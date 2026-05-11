import streamlit as st
import pandas as pd
import io
import re
import json
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from streamlit.components.v1 import html

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Automatizador de Preços PRO", layout="wide")

# --- INJEÇÃO DE JAVASCRIPT (F1 E F5) ---
js_shortcuts = """
<script>
    const doc = window.parent.document;
    doc.addEventListener('keydown', function(e) {
        if (e.key === 'F1') {
            e.preventDefault();
            const inputs = doc.querySelectorAll('input[type="text"]');
            const searchInput = Array.from(inputs).find(el => el.placeholder.includes("Digite e pressione"));
            if (searchInput) searchInput.focus();
        }
        if (e.key === 'F5') {
            e.preventDefault();
            const inputs = doc.querySelectorAll('input[type="text"]');
            const searchInput = Array.from(inputs).find(el => el.placeholder.includes("Digite e pressione"));
            if (searchInput) {
                searchInput.value = "";
                searchInput.dispatchEvent(new Event('input', { bubbles: true }));
                searchInput.dispatchEvent(new Event('change', { bubbles: true }));
                searchInput.focus();
            }
        }
    });
</script>
"""
html(js_shortcuts, height=0)

# --- CONSTANTES E STORAGE ---
DB_STORAGE = "master_database.csv"
USERS_STORAGE = "users_db.json"

# --- INICIALIZAÇÃO DE ESTADO ---
if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []
if 'replicar_data' not in st.session_state:
    st.session_state.replicar_data = None

def load_users():
    if not os.path.exists(USERS_STORAGE):
        return {"admin": {"password": "admin123", "expiry": "2099-12-31", "role": "admin"}}
    with open(USERS_STORAGE, "r") as f: return json.load(f)

def save_users(users):
    with open(USERS_STORAGE, "w") as f: json.dump(users, f)

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def login():
    st.sidebar.title("🔐 Acesso")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        users = load_users()
        if u in users and users[u]['password'] == p:
            st.session_state.autenticado = True
            st.session_state.user_role = users[u]['role']
            st.rerun()

if not st.session_state.autenticado:
    login()
    st.stop()

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if pd.isna(valor): return valor
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def format_barcode_robust(val):
    if pd.isna(val) or val == "" or val is None: return ""
    try:
        return "{:.0f}".format(float(val))
    except:
        return re.sub(r'\D', '', str(val))

def extract_all_barcodes(val):
    text = str(val)
    return re.findall(r'\d{8,14}', text)

def clean_barcode_prefix(barcode):
    if len(barcode) > 8 and (barcode.startswith('0') or barcode.startswith('1')):
        return barcode[1:]
    return barcode

def similarity(a, b):
    return SequenceMatcher(None, str(a).upper().strip(), str(b).upper().strip()).ratio()

def extrair_detalhes(texto):
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', str(texto).lower()))

@st.cache_data
def get_master_db():
    if os.path.exists(DB_STORAGE):
        df = pd.read_csv(DB_STORAGE)
        if 'codigo barras' in df.columns:
            df['codigo barras'] = df['codigo barras'].apply(format_barcode_robust)
        return df
    return pd.DataFrame(columns=['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'QUANT'])

# --- NAVEGAÇÃO ---
tabs = ["📊 Cotação", "💰 Vendas", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin":
    tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
    ignorar_01 = st.sidebar.checkbox("Ignorar 0 ou 1 à esquerda no EAN", value=True)
    estoque_minimo = st.sidebar.number_input("Estoque mínimo no banco para validar preço:", min_value=0, value=1)
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Planilha de Destino", type=["xlsx"])
    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início dos produtos:", 1, 1000, 11)
        
        t_df_view = pd.read_excel(target_file, header=header_pos-1)
        col1, col2, col3 = st.columns(3)
        desc_col = col1.selectbox("Coluna Descrição", t_df_view.columns)
        bar_col = col2.selectbox("Coluna Barras", t_df_view.columns)
        price_col = col3.selectbox("Coluna Preço", t_df_view.columns)

        if st.button("🚀 Processar Planilha"):
            product_map = {}
            for _, row in master_db.iterrows():
                product_map[str(row['codigo barras'])] = (row['preço'], row['estoque'])

            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            d_idx = col_indices.get(desc_col.strip())
            b_idx = col_indices.get(bar_col.strip())
            p_idx = col_indices.get(price_col.strip())

            if not all([d_idx, b_idx, p_idx]):
                st.error("Colunas não encontradas. Verifique a linha do cabeçalho.")
            else:
                preenchidos = 0
                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = ws.cell(row=r, column=b_idx).value
                    if not d_val: continue
                    
                    found_p = None
                    if "Barras" in modo or "Híbrido" in modo:
                        barcodes = extract_all_barcodes(b_val)
                        for b in barcodes:
                            b_f = format_barcode_robust(b)
                            if b_f in product_map:
                                p_val, s_val = product_map[b_f]
                                if s_val >= estoque_minimo: found_p = p_val; break
                            if found_p is None and ignorar_01:
                                b_l = clean_barcode_prefix(b_f)
                                if b_l in product_map:
                                    p_val, s_val = product_map[b_l]
                                    if s_val >= estoque_minimo: found_p = p_val; break
                    
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo):
                        best_sim = 0
                        d_det = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim = similarity(d_val, row_db['descrição'])
                            if sim >= 0.80 and d_det == extrair_detalhes(row_db['descrição']):
                                if sim > best_sim:
                                    best_sim = sim; found_p = row_db['preço']
                    
                    if found_p is not None:
                        final_p = float(found_p) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(final_p) if aplicar_arredondamento else final_p
                        preenchidos += 1

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Concluído! {preenchidos} preços carimbados.")
                st.download_button("📥 Baixar Resultado", out.getvalue(), target_file.name)

# --- ABA 2: VENDAS ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Vendas")
    master_db = get_master_db()
    
    query = st.text_input("🔍 Pesquisar Produto:", placeholder="Digite e pressione Enter (F1 para focar, F5 para limpar)")
    
    if query:
        terms = query.upper().split()
        mask = master_db['descrição'].apply(lambda x: all(t in str(x).upper() for t in terms))
        results = master_db[mask].head(15)
        
        for idx, row in results.iterrows():
            with st.container():
                c1, c2, c3, c4, c5 = st.columns([3, 1.5, 1, 1, 0.5])
                c1.write(f"**{row['descrição']}**")
                c2.write(f"EAN: `{row['codigo barras']}`")
                c3.write(f"R$ {row['preço']}")
                qtd = c4.number_input("Qtd", 1, 1000, 1, key=f"q_{row['codigo barras']}_{idx}")
                if c5.button("➕", key=f"b_{row['codigo barras']}_{idx}"):
                    st.session_state.carrinho.append({
                        "descrição": row['descrição'],
                        "ean": row['codigo barras'],
                        "qtd": qtd,
                        "preco": row['preço'],
                        "total": extra_round(row['preço'] * qtd)
                    })
                    st.rerun()

    if st.session_state.carrinho:
        st.divider()
        st.subheader("🛒 Carrinho")
        df_car = pd.DataFrame(st.session_state.carrinho)
        st.table(df_car)
        total_geral = df_car['total'].sum()
        st.write(f"### Total Geral: R$ {extra_round(total_geral)}")
        
        if st.button("🗑️ Limpar Carrinho"):
            st.session_state.carrinho = []
            st.rerun()

# --- ABA 3: GERENCIAR BANCO (6 COLUNAS E CÁLCULO) ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    f = st.file_uploader("Upload Banco (xlsx/csv)", type=["xlsx", "csv"])
    if f and st.button("💾 Salvar Banco"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        
        # Manutenção das 6 colunas
        df = df.iloc[:, [0, 1, 2, 3, 4, 5]]
        df.columns = ['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'QUANT']
        
        def extrair_inteiro(valor):
            numeros = re.findall(r'\d+', str(valor))
            return int(numeros[0]) if numeros else 1

        df['codigo barras'] = df['codigo barras'].apply(format_barcode_robust)
        df['caixa'] = pd.to_numeric(df['caixa'], errors='coerce').fillna(0)
        
        # Preço = Caixa / QUANT (número extraído)
        df['preço'] = df.apply(lambda r: r['caixa'] / extrair_inteiro(r['QUANT']), axis=1)
        
        df.to_csv(DB_STORAGE, index=False)
        st.cache_data.clear()
        st.success("Banco Atualizado com 06 Colunas e Preços Calculados!")
        st.dataframe(df.head())

# --- ABA 4: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    nu = st.text_input("Novo Usuário")
    np = st.text_input("Nova Senha")
    if st.button("Criar"):
        users[nu] = {"password": np, "expiry": "2099-12-31", "role": "user"}
        save_users(users)
        st.rerun()
    
    st.divider()
    for user, data in users.items():
        if user != "admin":
            st.write(f"Usuário: {user}")
            if st.button(f"Excluir {user}"):
                del users[user]
                save_users(users)
                st.rerun()

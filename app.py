import streamlit as st
import pandas as pd
import io
import re
import json
import os
import openpyxl
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from difflib import SequenceMatcher
from streamlit.components.v1 import html

# --- CONFIGURAÇÃO E LAYOUT ORIGINAL ---
st.set_page_config(page_title="Automatizador de Preços PRO", layout="wide")

# Atalhos F1 e F5
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

DB_STORAGE = "master_database.csv"
USERS_STORAGE = "users_db.json"

if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# --- GESTÃO DE USUÁRIOS ---
def load_users():
    if not os.path.exists(USERS_STORAGE):
        return {"admin": {"password": "admin123", "expiry": "2099-12-31", "role": "admin"}}
    with open(USERS_STORAGE, "r") as f: return json.load(f)

def save_users(users):
    with open(USERS_STORAGE, "w") as f: json.dump(users, f)

# --- LOGIN ---
if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acesso")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        users = load_users()
        if u in users and users[u]['password'] == p:
            st.session_state.autenticado = True
            st.session_state.user_role = users[u]['role']
            st.rerun()
    st.stop()

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if pd.isna(valor): return valor
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def format_barcode_robust(val):
    """Corrige notação científica do Excel (E+12) para texto puro"""
    if pd.isna(val) or val == "" or val is None: return ""
    try:
        return "{:.0f}".format(float(val))
    except:
        return re.sub(r'\D', '', str(val))

def similarity(a, b):
    return SequenceMatcher(None, str(a).upper().strip(), str(b).upper().strip()).ratio()

def extrair_detalhes(texto):
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', str(texto).lower()))

@st.cache_data
def get_master_db():
    if os.path.exists(DB_STORAGE):
        df = pd.read_csv(DB_STORAGE)
        df['codigo barras'] = df['codigo barras'].apply(format_barcode_robust)
        return df
    return pd.DataFrame(columns=['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'QUANT'])

# --- NAVEGAÇÃO ---
tabs = ["📊 Cotação", "💰 Vendas", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO (CARIMBAR PREÇOS) ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
    estoque_minimo = st.sidebar.number_input("Estoque mínimo no banco para carimbar:", min_value=0, value=0)
    discount = st.sidebar.number_input("Desconto (%)", 0.0)

    target_file = st.file_uploader("Upload Planilha de Destino", type=["xlsx"])
    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início dos produtos:", 1, 1000, 11)
        
        t_df_view = pd.read_excel(target_file, header=header_pos-1)
        col1, col2, col3 = st.columns(3)
        desc_col = col1.selectbox("Coluna Descrição na Planilha", t_df_view.columns)
        bar_col = col2.selectbox("Coluna Barras na Planilha", t_df_view.columns)
        price_col = col3.selectbox("Coluna Preço na Planilha (Destino)", t_df_view.columns)

        if st.button("🚀 Carimbar Preços na Planilha"):
            # Mapeia EAN -> Preço
            product_map = {}
            for _, row in master_db.iterrows():
                product_map[str(row['codigo barras'])] = (row['preço'], row['estoque'])

            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            # Localizar colunas
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            d_idx = col_indices.get(desc_col.strip())
            b_idx = col_indices.get(bar_col.strip())
            p_idx = col_indices.get(price_col.strip())

            if not all([d_idx, b_idx, p_idx]):
                st.error("Colunas não encontradas. Verifique a linha do cabeçalho.")
            else:
                count = 0
                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = format_barcode_robust(ws.cell(row=r, column=b_idx).value)
                    if not d_val: continue
                    
                    found_p = None
                    # Busca por EAN
                    if b_val in product_map:
                        p_val, s_val = product_map[b_val]
                        if s_val >= estoque_minimo:
                            found_p = p_val
                    
                    # Busca por Similaridade (se EAN falhar)
                    if found_p is None and "Barras" not in modo:
                        for _, row_db in master_db.iterrows():
                            if similarity(d_val, row_db['descrição']) > 0.85:
                                found_p = row_db['preço']
                                break
                    
                    if found_p is not None:
                        final_p = float(found_p) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(final_p)
                        count += 1

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Finalizado! {count} preços carimbados.")
                st.download_button("📥 Baixar Planilha Atualizada", out.getvalue(), f"atualizada_{target_file.name}")

# --- ABA 2: VENDAS ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Vendas")
    master_db = get_master_db()
    query = st.text_input("🔍 Buscar Produto (F1):", placeholder="Digite e pressione Enter")
    
    if query:
        results = master_db[master_db['descrição'].str.contains(query, case=False, na=False)].head(15)
        for _, row in results.iterrows():
            with st.container():
                c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
                c1.write(f"**{row['descrição']}**")
                c2.write(f"EAN: {row['codigo barras']}")
                c3.write(f"R$ {row['preço']}")
                if c4.button("Adicionar", key=row['codigo barras']):
                    st.session_state.carrinho.append(row.to_dict())
                    st.toast("Adicionado!")

    if st.session_state.carrinho:
        st.divider()
        st.subheader("🛒 Carrinho")
        st.table(pd.DataFrame(st.session_state.carrinho)[['descrição', 'preço']])

# --- ABA 3: GERENCIAR BANCO (AQUI ESTÁ A MUDANÇA DAS 06 COLUNAS) ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    st.info("O arquivo deve ter 6 colunas: descrição, codigo barras, preço, estoque, caixa, QUANT")
    
    f = st.file_uploader("Upload Banco Mestre", type=["xlsx", "csv"])
    if f and st.button("💾 Processar e Salvar"):
        df = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
        
        # Garante as 6 colunas
        df = df.iloc[:, [0, 1, 2, 3, 4, 5]]
        df.columns = ['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'QUANT']
        
        # Função para pegar só o número da coluna QUANT (ex: "12 un" vira 12)
        def extrair_numero(v):
            numeros = re.findall(r'\d+', str(v))
            return int(numeros[0]) if numeros else 1

        # Lógica pedida: Preço = Caixa / Quantidade
        df['codigo barras'] = df['codigo barras'].apply(format_barcode_robust)
        df['caixa'] = pd.to_numeric(df['caixa'], errors='coerce').fillna(0)
        df['preço'] = df.apply(lambda r: r['caixa'] / extrair_numero(r['QUANT']), axis=1)
        
        # Salva o CSV com as 6 colunas completas
        df.to_csv(DB_STORAGE, index=False)
        st.cache_data.clear()
        st.success("Banco de Dados salvo com sucesso (6 colunas)!")
        st.dataframe(df.head())

# --- ABA 4: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    with st.expander("Criar Novo Usuário"):
        new_u = st.text_input("Nome")
        new_p = st.text_input("Senha")
        if st.button("Adicionar"):
            users[new_u] = {"password": new_p, "expiry": "2026-12-31", "role": "user"}
            save_users(users)
            st.rerun()

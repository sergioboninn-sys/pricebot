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

# --- INJEÇÃO DE JAVASCRIPT (ATALHOS F1 E F5) ---
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
VENDAS_STORAGE = "vendas_history.json"

# --- INICIALIZAÇÃO DE ESTADO ---
if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []
if 'replicar_data' not in st.session_state:
    st.session_state.replicar_data = None
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

# --- FUNÇÕES DE USUÁRIO ---
def load_users():
    if not os.path.exists(USERS_STORAGE):
        return {"admin": {"password": "admin123", "expiry": "2099-12-31", "role": "admin"}}
    with open(USERS_STORAGE, "r") as f: return json.load(f)

def save_users(users):
    with open(USERS_STORAGE, "w") as f: json.dump(users, f)

# --- LOGIN ---
if not st.session_state.autenticado:
    st.sidebar.title("🔐 Acesso Restrito")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        users = load_users()
        if u in users and users[u]['password'] == p:
            exp_date = datetime.strptime(users[u]['expiry'], "%Y-%m-%d")
            if datetime.now() <= exp_date:
                st.session_state.autenticado = True
                st.session_state.user_role = users[u]['role']
                st.session_state.user_name = u
                st.rerun()
            else:
                st.sidebar.error("Acesso expirado.")
        else:
            st.sidebar.error("Credenciais inválidas.")
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
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido (Barras + Similaridade)", "Apenas Barras", "Apenas Similaridade"])
    ignorar_01 = st.sidebar.checkbox("Ignorar 0 ou 1 à esquerda no EAN", value=True)
    estoque_minimo = st.sidebar.number_input("Estoque mínimo no banco para validar preço:", min_value=0, value=1)
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

    target_file = st.file_uploader("Upload Planilha de Destino", type=["xlsx"])
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
                st.download_button("📥 Baixar Resultado", out.getvalue(), f"PRE_COTACAO_{target_file.name}")

# --- ABA 2: VENDAS (LOGICA COMPLETA RESTAURADA) ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
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
                    st.toast("Adicionado ao carrinho!")
                    st.rerun()

    if st.session_state.carrinho:
        st.divider()
        st.subheader("🛒 Itens do Pedido")
        df_car = pd.DataFrame(st.session_state.carrinho)
        st.dataframe(df_car, use_container_width=True)
        total_geral = df_car['total'].sum()
        st.write(f"### Total Geral: R$ {extra_round(total_geral)}")
        
        if st.button("🗑️ Limpar Carrinho"):
            st.session_state.carrinho = []
            st.rerun()

# --- ABA 3: GERENCIAR BANCO (AQUI ESTÃO AS NOVAS IMPLANTAÇÕES) ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    st.markdown("### Configurações de Importação")
    st.info("O arquivo deve conter as 6 colunas na ordem: descrição, codigo barras, preço, estoque, caixa, QUANT")
    
    file_db = st.file_uploader("Upload Banco Mestre (Excel ou CSV)", type=["xlsx", "csv"])
    
    if file_db:
        if st.button("💾 Processar e Salvar no Sistema"):
            try:
                df = pd.read_excel(file_db) if file_db.name.endswith('.xlsx') else pd.read_csv(file_db)
                
                # Garante que pegamos apenas as 6 primeiras colunas e renomeamos corretamente
                df = df.iloc[:, [0, 1, 2, 3, 4, 5]]
                df.columns = ['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'QUANT']
                
                # Função interna para extrair apenas o número da coluna QUANT (ex: "12 un" -> 12)
                def get_num_only(v):
                    m = re.search(r'\d+', str(v))
                    return int(m.group()) if m else 1

                # Limpeza de dados e cálculos
                df['codigo barras'] = df['codigo barras'].apply(format_barcode_robust)
                df['caixa'] = pd.to_numeric(df['caixa'], errors='coerce').fillna(0)
                df['estoque'] = pd.to_numeric(df['estoque'], errors='coerce').fillna(0)
                
                # O preço agora é calculado pela divisão solicitada
                df['preço'] = df.apply(lambda r: r['caixa'] / get_num_only(r['QUANT']), axis=1)
                
                # Salva o arquivo CSV final com as 6 colunas
                df.to_csv(DB_STORAGE, index=False)
                st.cache_data.clear()
                st.success("Banco de dados atualizado com 06 colunas e preços unitários calculados!")
                st.dataframe(df.head(10))
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {e}")

# --- ABA 4: USUÁRIOS (LOGICA COMPLETA RESTAURADA) ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários")
    users = load_users()
    
    with st.expander("➕ Adicionar Novo Usuário"):
        new_u = st.text_input("Nome do Usuário")
        new_p = st.text_input("Senha")
        new_v = st.number_input("Dias de Validade", 1, 365, 30)
        if st.button("Cadastrar"):
            if new_u and new_p:
                exp_date = (datetime.now() + timedelta(days=new_v)).strftime("%Y-%m-%d")
                users[new_u] = {"password": new_p, "expiry": exp_date, "role": "user"}
                save_users(users)
                st.success(f"Usuário {new_u} criado!")
                st.rerun()

    st.divider()
    st.subheader("Usuários Existentes")
    for user, data in users.items():
        if user != "admin":
            c1, c2, c3 = st.columns([2, 2, 1])
            c1.write(f"**{user}** ({data['role']})")
            c2.write(f"Expira em: {data['expiry']}")
            if c3.button("Remover", key=f"del_{user}"):
                del users[user]
                save_users(users)
                st.rerun()

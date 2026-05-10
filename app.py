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

# --- CONFIGURAÇÃO INICIAL ---
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

DB_STORAGE = "master_database.csv"
USERS_STORAGE = "users_db.json"

# --- PERSISTÊNCIA DE ESTADO ---
if 'carrinho' not in st.session_state:
    st.session_state.carrinho = []
if 'replicar_data' not in st.session_state:
    st.session_state.replicar_data = None
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
    st.sidebar.title("🔐 Acesso Restrito")
    u = st.sidebar.text_input("Usuário")
    p = st.sidebar.text_input("Senha", type="password")
    if st.sidebar.button("Entrar"):
        users = load_users()
        if u in users and users[u]['password'] == p:
            exp = datetime.strptime(users[u]['expiry'], "%Y-%m-%d")
            if datetime.now() <= exp:
                st.session_state.autenticado = True
                st.session_state.user_role = users[u]['role']
                st.rerun()
        st.sidebar.error("Usuário ou Senha inválidos")
    st.stop()

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if pd.isna(valor): return valor
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def format_barcode_robust(val):
    """Converte notação científica e floats para string numérica pura"""
    if pd.isna(val) or val == "" or val is None: return ""
    try:
        # Resolve o problema do 7.89E+12
        return "{:.0f}".format(float(val))
    except:
        return re.sub(r'\D', '', str(val))

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
        df['Barcode'] = df['Barcode'].apply(format_barcode_robust)
        if 'Stock' not in df.columns: df['Stock'] = 0
        return df
    return pd.DataFrame(columns=['Description', 'Barcode', 'Price', 'Stock'])

# --- NAVEGAÇÃO ---
tabs = ["📊 Cotação", "💰 Vendas", "⚙️ Gerenciar Banco"]
if st.session_state.user_role == "admin": tabs.append("👤 Usuários")
aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    
    st.sidebar.header("Configurações")
    modo = st.sidebar.selectbox("Regra de Busca:", ["Híbrido", "Apenas Barras", "Apenas Similaridade"])
    ignorar_01 = st.sidebar.checkbox("Ignorar 0 ou 1 à esquerda no EAN", value=True)
    estoque_minimo = st.sidebar.number_input("Estoque mínimo no banco para validar:", min_value=0, value=0)
    discount = st.sidebar.number_input("Desconto (%)", 0.0)
    
    target_file = st.file_uploader("Upload Planilha de Destino", type=["xlsx"])
    if target_file:
        c1, c2 = st.columns(2)
        header_pos = c1.number_input("Linha do Cabeçalho:", 1, 100, 10)
        start_row = c2.number_input("Linha de início dos produtos:", 1, 1000, 11)
        
        # Carrega apenas para visualização de colunas
        t_df_view = pd.read_excel(target_file, header=header_pos-1)
        col1, col2, col3 = st.columns(3)
        desc_col = col1.selectbox("Coluna Descrição", t_df_view.columns)
        bar_col = col2.selectbox("Coluna Barras", t_df_view.columns)
        price_col = col3.selectbox("Coluna Preço", t_df_view.columns)

        if st.button("🚀 Processar e Carimbar Preços"):
            # Mapeia o banco para busca rápida
            product_map = {}
            for _, row in master_db.iterrows():
                product_map[str(row['Barcode'])] = (row['Price'], row['Stock'])

            target_file.seek(0)
            wb = openpyxl.load_workbook(target_file)
            ws = wb.active
            
            # Localiza índices das colunas (base 1 para openpyxl)
            col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
            
            d_idx = col_indices.get(desc_col.strip())
            b_idx = col_indices.get(bar_col.strip())
            p_idx = col_indices.get(price_col.strip())

            if not all([d_idx, b_idx, p_idx]):
                st.error("Colunas não identificadas. Verifique a linha do cabeçalho.")
            else:
                preenchidos = 0
                for r in range(int(start_row), ws.max_row + 1):
                    d_val = ws.cell(row=r, column=d_idx).value
                    b_val = ws.cell(row=r, column=b_idx).value
                    if not d_val: continue
                    
                    found_p = None
                    # 1. Busca por Barras
                    if "Barras" in modo or "Híbrido" in modo:
                        ean_limpo = format_barcode_robust(b_val)
                        if ean_limpo in product_map:
                            p_val, s_val = product_map[ean_limpo]
                            if s_val >= estoque_minimo: found_p = p_val
                        
                        if found_p is None and ignorar_01:
                            ean_alt = clean_barcode_prefix(ean_limpo)
                            if ean_alt in product_map:
                                p_val, s_val = product_map[ean_alt]
                                if s_val >= estoque_minimo: found_p = p_val
                    
                    # 2. Busca por Similaridade
                    if found_p is None and ("Similaridade" in modo or "Híbrido" in modo):
                        best_sim = 0
                        d_det = extrair_detalhes(d_val)
                        for _, row_db in master_db.iterrows():
                            sim = similarity(d_val, row_db['Description'])
                            if sim >= 0.80 and d_det == extrair_detalhes(row_db['Description']):
                                if sim > best_sim:
                                    best_sim = sim
                                    found_p = row_db['Price']
                    
                    if found_p is not None:
                        final_p = float(found_p) * (1 - (discount/100))
                        ws.cell(row=r, column=p_idx).value = extra_round(final_p)
                        preenchidos += 1

                out = io.BytesIO()
                wb.save(out)
                st.success(f"Sucesso! {preenchidos} itens foram atualizados.")
                st.download_button("📥 Baixar Planilha Carimbada", out.getvalue(), f"resultado_{target_file.name}")

# --- ABA 2: VENDAS ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    
    query = st.text_input("🔍 Pesquisar Produto:", placeholder="Digite e pressione Enter (F1 para focar, F5 para limpar)")
    
    if query:
        terms = query.upper().split()
        mask = master_db['Description'].apply(lambda x: all(t in str(x).upper() for t in terms))
        results = master_db[mask].head(15)
        
        for _, row in results.iterrows():
            with st.container():
                c_desc, c_ean, c_pr, c_qtd, c_btn = st.columns([3, 1.5, 1, 1, 0.5])
                c_desc.write(f"**{row['Description']}**")
                c_ean.write(f"`{row['Barcode']}`")
                p_sug = extra_round(row['Price'])
                c_pr.write(f"R$ {p_sug}")
                qtd_v = c_qtd.number_input("Qtd", 1, 1000, 1, key=f"v_{row['Barcode']}")
                if c_btn.button("➕", key=f"b_{row['Barcode']}"):
                    st.session_state.carrinho.append({
                        "Descrição": row['Description'],
                        "EAN": row['Barcode'],
                        "Qtd": qtd_v,
                        "Preço Unit": p_sug,
                        "Total": extra_round(p_sug * qtd_v)
                    })
                    st.rerun()

    st.divider()
    if st.session_state.carrinho:
        st.subheader("🛒 Itens no Pedido")
        df_c = pd.DataFrame(st.session_state.carrinho)
        st.table(df_c)
        st.write(f"### Total: R$ {extra_round(df_c['Total'].sum())}")
        if st.button("🗑️ Esvaziar Carrinho"):
            st.session_state.carrinho = []
            st.rerun()

# --- ABA 3: GERENCIAR BANCO ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Configuração do Banco de Dados")
    st.markdown("""
    **Instruções para o Arquivo:**
    O arquivo deve conter 5 colunas na ordem: 
    1. Descrição | 2. Barras | 3. Valor da Caixa | 4. Quantidade (ex: 12 ou '12 un') | 5. Estoque
    """)
    
    f_banco = st.file_uploader("Upload Banco Mestre", type=["xlsx", "csv"])
    if f_banco and st.button("💾 Processar e Salvar"):
        df = pd.read_excel(f_banco) if f_banco.name.endswith('.xlsx') else pd.read_csv(f_banco)
        
        # Seleciona as 5 primeiras colunas e renomeia
        df = df.iloc[:, [0, 1, 2, 3, 4]]
        df.columns = ['Description', 'Barcode', 'Caixa', 'Quant', 'Stock']
        
        def extrair_qtd(v):
            num = re.search(r'\d+', str(v))
            return int(num.group()) if num else 1

        # Limpeza e Cálculo
        df['Barcode'] = df['Barcode'].apply(format_barcode_robust)
        df['Caixa'] = pd.to_numeric(df['Caixa'], errors='coerce').fillna(0)
        df['Stock'] = pd.to_numeric(df['Stock'], errors='coerce').fillna(0)
        
        # O PREÇO UNITÁRIO É CALCULADO AQUI
        df['Price'] = df.apply(lambda r: r['Caixa'] / extrair_qtd(r['Quant']), axis=1)
        
        # Salva apenas o necessário para o sistema
        df_save = df[['Description', 'Barcode', 'Price', 'Stock']]
        df_save.to_csv(DB_STORAGE, index=False)
        st.cache_data.clear()
        st.success("Banco de dados atualizado com preços unitários calculados!")

# --- ABA 4: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Acessos")
    users = load_users()
    
    nu = st.text_input("Novo Usuário")
    np = st.text_input("Nova Senha", type="password")
    nd = st.number_input("Dias de Validade", 1, 365, 30)
    
    if st.button("Criar Usuário"):
        if nu:
            exp_date = (datetime.now() + timedelta(days=nd)).strftime("%Y-%m-%d")
            users[nu] = {"password": np, "expiry": exp_date, "role": "user"}
            save_users(users)
            st.success(f"Usuário {nu} criado até {exp_date}")
            st.rerun()

    st.divider()
    st.subheader("Usuários Ativos")
    for user, data in users.items():
        if user != "admin":
            st.write(f"**{user}** - Expira em: {data['expiry']}")
            if st.button(f"Remover {user}"):
                del users[user]
                save_users(users)
                st.rerun()

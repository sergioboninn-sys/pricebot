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

# --- CONFIGURAÇÃO E PERSISTÊNCIA ---
st.set_page_config(page_title="Automatizador de Preços PRO", layout="wide")

# --- INJEÇÃO DE JAVASCRIPT ---
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

if 'replicar_data' not in st.session_state:
    st.session_state.replicar_data = None

def load_users():
    if not os.path.exists(USERS_STORAGE):
        return {
            "admin": {
                "password": "admin123", 
                "expiry": "2099-12-31", 
                "role": "admin",
                "permissions": {
                    "ver_estoque": True,
                    "ver_preco": True,
                    "ver_caixa": True,
                    "ver_quant": True,
                    "funcao_cotacao": True,
                    "funcao_vendas": True,
                    "busca_hibrido": True,
                    "busca_barras": True,
                    "busca_similaridade": True,
                    "ver_gerenciar_banco": True
                }
            }
        }
    with open(USERS_STORAGE, "r") as f: 
        return json.load(f)

def save_users(users):
    with open(USERS_STORAGE, "w") as f: 
        json.dump(users, f, indent=4)

# --- LOGIN ---
if 'autenticado' not in st.session_state:
    st.session_state.autenticado = False

def login():
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
                st.session_state.username = u
                st.rerun()
        st.sidebar.error("Acesso negado")

if not st.session_state.autenticado:
    login()
    st.stop()

# --- FUNÇÕES DE APOIO ---
def extra_round(valor):
    if pd.isna(valor): return valor
    return float(Decimal(str(valor)).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP))

def extract_all_barcodes(val):
    if pd.isna(val) or val is None: return []
    text = str(val).split('.')[0]
    return re.findall(r'\d{8,14}', text)

def clean_barcode_prefix(barcode):
    if len(barcode) > 8 and (barcode.startswith('0') or barcode.startswith('1')):
        return barcode[1:]
    return barcode

def similarity(a, b):
    return SequenceMatcher(None, str(a).upper().strip(), str(b).upper().strip()).ratio()

def extrair_detalhes(texto):
    return set(re.findall(r'(\d+\s?(?:g|gr|kg|l|lt|ml)\b)', str(texto).lower()))

# --- CARREGAMENTO DO BANCO ---
@st.cache_data
def get_master_db():
    if os.path.exists(DB_STORAGE):
        df = pd.read_csv(DB_STORAGE)
        df['Barcode'] = df['Barcode'].astype(str).str.replace(r'\.0$', '', regex=True)
        for col in ['Stock', 'caixa', 'QUANT']:
            if col not in df.columns: df[col] = 0
        return df
    return pd.DataFrame(columns=['Description', 'Barcode', 'Price', 'Stock', 'caixa', 'QUANT'])

# --- INTERFACE & CONTROLE DE NAVEGAÇÃO ---
users = load_users()
current_user_data = users.get(st.session_state.username, {})
current_perms = current_user_data.get("permissions", {})

tabs = ["📊 Cotação", "💰 Vendas"]
if current_perms.get("ver_gerenciar_banco", True) or st.session_state.user_role == "admin":
    tabs.append("⚙️ Gerenciar Banco")
if st.session_state.user_role == "admin": 
    tabs.append("👤 Usuários")

aba = st.sidebar.radio("Navegação", tabs)

# --- ABA 1: COTAÇÃO ---
if aba == "📊 Cotação":
    st.title("📊 Automatizador de Cotações")
    master_db = get_master_db()
    if master_db.empty: st.warning("Banco vazio.")
    
    modos_permitidos = []
    if current_perms.get("busca_hibrido", True) or st.session_state.user_role == "admin":
        modos_permitidos.append("Híbrido (Barras + Similaridade)")
    if current_perms.get("busca_barras", True) or st.session_state.user_role == "admin":
        modos_permitidos.append("Apenas Barras")
    if current_perms.get("busca_similaridade", True) or st.session_state.user_role == "admin":
        modos_permitidos.append("Apenas Similaridade")
        
    st.sidebar.header("Configurações")
    
    if modos_permitidos:
        modo = st.sidebar.selectbox("Regra de Busca:", modos_permitidos)
        
        ignorar_01 = st.sidebar.checkbox("Ignorar 0 ou 1 à esquerda no EAN", value=False)
        
        st.sidebar.subheader("Filtros de Estoque Mínimo")
        estoque_geral = st.sidebar.number_input("Estoque mínimo Geral:", min_value=0, value=1)
        estoque_unitario = st.sidebar.number_input("Estoque mínimo para Unitários (QUANT = 1):", min_value=0, value=1)
        estoque_caixas = st.sidebar.number_input("Estoque mínimo para Caixas/Dz/Dp/Cj (QUANT > 1):", min_value=0, value=1)
        
        discount = st.sidebar.number_input("Desconto (%)", 0.0)
        aplicar_arredondamento = st.sidebar.checkbox("Arredondar preços", value=True)

        st.sidebar.subheader("Opções de Download")
        opcao_salvamento = st.sidebar.radio(
            "Como deseja baixar o resultado?",
            ["Novo Arquivo (cotacao_corrigida.xlsx)", "Mesmo nome do arquivo original"]
        )

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

            if st.button("🚀 Processar e Preservar Formatação"):
                master_db['Stock'] = pd.to_numeric(master_db['Stock'], errors='coerce').fillna(0)
                
                # Criamos um mapeamento completo contendo: Preço, Estoque e QUANT
                product_map = {}
                for _, row_db in master_db.iterrows():
                    b_str = str(row_db['Barcode'])
                    try:
                        q_val = int(float(row_db['QUANT']))
                    except:
                        q_val = 1
                    product_map[b_str] = {
                        'Price': row_db['Price'], 
                        'Stock': row_db['Stock'], 
                        'QUANT': q_val
                    }
                
                if opcao_salvamento == "Mesmo nome do arquivo original":
                    output_name = target_file.name
                else:
                    output_name = "cotacao_corrigida.xlsx"

                target_file.seek(0)
                wb = openpyxl.load_workbook(target_file)
                ws = wb.active
                col_indices = {str(ws.cell(row=header_pos, column=i).value).strip(): i for i in range(1, ws.max_column + 1)}
                
                try:
                    d_idx = col_indices[desc_col.strip()]
                    b_idx = col_indices[bar_col.strip()]
                    p_idx = col_indices[price_col.strip()]
                    
                    for r in range(int(start_row), ws.max_row + 1):
                        d_val = ws.cell(row=r, column=d_idx).value
                        if not d_val or str(d_val).strip() == "": continue
                        
                        found_p = None
                        b_val = ws.cell(row=r, column=b_idx).value
                        
                        if "Barras" in modo or "Híbrido" in modo:
                            barcodes = extract_all_barcodes(b_val)
                            for b in barcodes:
                                target_barcode = None
                                if b in product_map:
                                    target_barcode = b
                                elif ignorar_01:
                                    b_limpo = clean_barcode_prefix(b)
                                    if b_limpo in product_map:
                                        target_barcode = b_limpo
                                
                                if target_barcode:
                                    p_info = product_map[target_barcode]
                                    s_val = p_info['Stock']
                                    q_val = p_info['QUANT']
                                    
                                    # Aplicação dinâmica das novas regras estruturadas de estoque mínimo
                                    if q_val == 1:
                                        if s_val >= estoque_unitario:
                                            found_p = p_info['Price']
                                            break
                                    elif q_val > 1:
                                        if s_val >= estoque_caixas:
                                            found_p = p_info['Price']
                                            break
                                    else:
                                        if s_val >= estoque_geral:
                                            found_p = p_info['Price']
                                            break
                        
                        if found_p is None and ("Similaridade" in modo or "Híbrido" in modo):
                            best_sim = 0
                            d_det = extrair_detalhes(d_val)
                            for _, row_db in master_db.iterrows():
                                sim = similarity(d_val, row_db['Description'])
                                if sim >= 0.75 and d_det == extrair_detalhes(row_db['Description']):
                                    if sim > best_sim:
                                        best_sim = sim
                                        found_p = row_db['Price']
                        
                        if found_p is not None:
                            f_p = float(found_p) * (1 - (discount/100))
                            ws.cell(row=r, column=p_idx).value = extra_round(f_p) if aplicar_arredondamento else f_p

                    contador_final = 0
                    for r in range(int(start_row), ws.max_row + 1):
                        celula_preco = ws.cell(row=r, column=p_idx).value
                        try:
                            if celula_preco is not None and float(celula_preco) > 0:
                                contador_final += 1
                        except: continue

                    out = io.BytesIO()
                    wb.save(out)
                    st.success(f"Sucesso! Foram preenchidos **{contador_final}** itens com preços.")
                    st.download_button(f"📥 Baixar Planilha ({output_name})", out.getvalue(), output_name)
                    
                except Exception as e:
                    st.error(f"Erro ao processar: {e}")
    else:
        st.error("Seu usuário não possui permissão para utilizar nenhuma regra de busca. Contate o administrador.")

# --- ABA 2: VENDAS ---
elif aba == "💰 Vendas":
    st.title("💰 Consulta e Pré-Pedido")
    master_db = get_master_db()
    col_vendas_1, col_vendas_2 = st.columns([3, 1])
    
    with col_vendas_1:
        st.subheader("🔍 Busca e Adição Rápida")
        query = st.text_input("Pesquisar produto:", placeholder="Digite e pressione Enter", key="search_query")
        desc_geral = st.number_input("Desconto Padrão na Tabela (%)", 0.0, 100.0, 0.0)
        
        if query:
            search_terms = query.replace('%', ' ').split()
            mask = master_db['Description'].apply(lambda x: all(term.upper() in str(x).upper() for term in search_terms))
            results = master_db[mask].head(20).copy() 
            
            if not results.empty:
                if st.session_state.replicar_data:
                    with st.container():
                        st.warning("🔄 **Replicação de Família Detectada**")
                        rep = st.session_state.replicar_data
                        st.write(f"Deseja replicar Qtd: {rep['qtd']} e Preço: R$ {rep['preco']} para família '{rep['familia']}'?")
                        familia_results = results[results['Description'].str.contains(rep['familia'], case=False) & (results['Barcode'] != rep['ean'])]
                        escolha = st.radio("Como deseja replicar?", ["Replicar em todos", "Escolher específicos"], horizontal=True)
                        
                        if escolha == "Escolher específicos":
                            selecionados = st.multiselect("Itens:", familia_results['Description'].tolist(), default=familia_results['Description'].tolist())
                            if st.button("Confirmar Seleção"):
                                for _, rf in familia_results.iterrows():
                                    if rf['Description'] in selecionados:
                                        st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                                st.session_state.replicar_data = None
                                st.rerun()
                        else:
                            if st.button("Confirmar em Todos"):
                                for _, rf in familia_results.iterrows():
                                    st.session_state.carrinho.append({"Descrição": rf['Description'], "EAN": rf['Barcode'], "Qtd": rep['qtd'], "Preço Unit": rep['preco'], "Total": extra_round(rep['preco'] * rep['qtd'])})
                                st.session_state.replicar_data = None
                                st.rerun()
                
                for idx, row in results.iterrows():
                    p_sugestao = extra_round(float(row['Price']) * (1 - (desc_geral/100)))
                    with st.container():
                        c_desc, c_ean, c_preco, c_qtd, c_add = st.columns([3, 1.5, 1.2, 1, 0.5])
                        c_desc.write(f"**{row['Description']}**")
                        c_ean.write(f"`{row['Barcode']}`")
                        c_preco.write(f"R$ {p_sugestao}")
                        input_qtd = c_qtd.number_input("Qtd", 1, 1000, 1, key=f"qtd_{idx}")
                        if c_add.button("➕", key=f"btn_{idx}"):
                            st.session_state.carrinho.append({"Descrição": row['Description'], "EAN": row['Barcode'], "Qtd": input_qtd, "Preço Unit": p_sugestao, "Total": extra_round(p_sugestao * input_qtd)})
                            palavras = row['Description'].split()
                            if len(palavras) >= 2:
                                st.session_state.replicar_data = {"familia": f"{palavras[0]} {palavras[1]}", "qtd": input_qtd, "preco": p_sugestao, "ean": row['Barcode']}
                            st.rerun()

    with col_vendas_2:
        st.subheader("🛒 Seu Pedido")
        if st.session_state.carrinho:
            df_cart = pd.DataFrame(st.session_state.carrinho)
            for item in st.session_state.carrinho:
                st.write(f"{item['Qtd']}x {item['Descrição']} - **R$ {item['Total']}**")
            total_pedido = df_cart['Total'].sum()
            st.metric("Total Geral", f"R$ {extra_round(total_pedido)}")
            if st.button("🗑️ Limpar Tudo"):
                st.session_state.carrinho = []
                st.rerun()
            output_xlsx = io.BytesIO()
            wb_ped = openpyxl.Workbook()
            ws_ped = wb_ped.active
            ws_ped.append(["Descrição", "Código EAN", "Qtd", "Preço Unit.", "Total"])
            for item in st.session_state.carrinho:
                ws_ped.append([item['Descrição'], item['EAN'], item['Qtd'], item['Preço Unit'], item['Total']])
            wb_ped.save(output_xlsx)
            st.download_button("📥 Baixar Pedido", output_xlsx.getvalue(), "pedido.xlsx")
        else:
            st.info("Carrinho vazio.")

# --- ABA 3: GERENCIAR BANCO ---
elif aba == "⚙️ Gerenciar Banco":
    st.title("⚙️ Gerenciar Banco de Dados")
    master_db = get_master_db()

    # =====================================================================
    # >>> NOVO BLOCO 5 (parte inicial): PERMISSÕES POR FUNÇÃO DO GERENCIADOR DE BANCO
    # =====================================================================
    perms_banco_usuario = users.get(st.session_state.username, {}).get("permissions", {})

    def _tem_perm_banco(chave):
        return st.session_state.user_role == "admin" or perms_banco_usuario.get(chave, True)

    perm_b_atualizar = _tem_perm_banco("banco_atualizar")
    perm_b_exportar = _tem_perm_banco("banco_exportar")
    perm_b_visualizar = _tem_perm_banco("banco_visualizar")
    perm_b_exp_desconto = _tem_perm_banco("banco_exportar_desconto")
    perm_b_exp_importacao = _tem_perm_banco("banco_exportar_importacao")
    perm_b_editar = _tem_perm_banco("banco_editar")
    perm_b_recalcular = _tem_perm_banco("banco_editar") and _tem_perm_banco("banco_recalcular")
    # =====================================================================
    
    col_upload, col_export = st.columns(2)
    
    if perm_b_atualizar:
        with col_upload:
            st.subheader("📥 Atualizar Base de Dados")
            f = st.file_uploader("Upload Banco (xlsx/csv)", type=["xlsx", "csv"])
            if f and st.button("💾 Salvar Banco"):
                df_file = pd.read_excel(f) if f.name.endswith('.xlsx') else pd.read_csv(f)
            
                pure_desc = [str(x).strip() if pd.notna(x) else "" for x in list(df_file.iloc[:, 0])] if len(df_file.columns) > 0 else [""] * len(df_file)
                pure_bar = [str(x).strip() if pd.notna(x) else "" for x in list(df_file.iloc[:, 1])] if len(df_file.columns) > 1 else [""] * len(df_file)
                pure_price = [float(x) if pd.notna(x) else 0.0 for x in list(df_file.iloc[:, 2])] if len(df_file.columns) > 2 else [0.0] * len(df_file)
            
                pure_stock = []
                for x in list(df_file.iloc[:, 3]):
                    try: pure_stock.append(int(float(x)) if pd.notna(x) else 0)
                    except: pure_stock.append(0)
                
                pure_caixa = []
                for x in list(df_file.iloc[:, 4]):
                    try: pure_caixa.append(float(x) if pd.notna(x) else 0.0)
                    except: pure_caixa.append(0.0)
                
                pure_quant = []
                for x in list(df_file.iloc[:, 5]):
                    try: pure_quant.append(int(float(x)) if pd.notna(x) else 1)
                    except: pure_quant.append(1)

                df_novo = pd.DataFrame({
                    'Description': pure_desc,
                    'Barcode': pure_bar,
                    'Price': pure_price,
                    'Stock': pure_stock,
                    'caixa': pure_caixa,
                    'QUANT': pure_quant
                })

                df_antigo = get_master_db()
            
                def get_num_only(v):
                    m = re.search(r'\d+', str(v))
                    return int(m.group()) if m else 1

                lista_final = []
                antigo_dict = {}
                if not df_antigo.empty:
                    antigo_dict = dict(zip(df_antigo['Barcode'].astype(str), df_antigo['QUANT']))
                
                for idx, row in df_novo.iterrows():
                    raw_desc = row['Description']
                    raw_barcode = row['Barcode']
                    raw_stock = row['Stock']
                    raw_caixa = row['caixa']
                    raw_quant = row['QUANT']
                
                    barcode_str = re.sub(r'\D', '', str(raw_barcode).split('.')[0])
                    if not barcode_str: continue
                
                    if barcode_str in antigo_dict:
                        quant_final = antigo_dict[barcode_str]
                    else:
                        quant_final = raw_quant
                    
                    divisor = get_num_only(quant_final)
                    nova_caixa = pd.to_numeric(raw_caixa, errors='coerce')
                    nova_caixa = float(nova_caixa) if not pd.isna(nova_caixa) else 0.0
                    preco_calculado = nova_caixa / divisor if divisor > 0 else 0.0
                
                    lista_final.append({
                        'Description': raw_desc,
                        'Barcode': barcode_str,
                        'Price': preco_calculado,
                        'Stock': int(raw_stock),
                        'caixa': nova_caixa,
                        'QUANT': quant_final
                    })
                
                df_resultado = pd.DataFrame(lista_final)
                df_resultado.to_csv(DB_STORAGE, index=False)
                st.cache_data.clear()
                st.success("Banco Updated!")
                st.rerun()

    if perm_b_exportar:
        with col_export:
            st.subheader("📤 Exportar Base de Dados")
            if not master_db.empty:
                st.write(f"A base possui atualmente **{len(master_db)}** produtos cadastrados.")
            
                buffer_db = io.BytesIO()
                with pd.ExcelWriter(buffer_db, engine='openpyxl') as writer:
                    df_export = master_db.copy()
                    df_export.columns = ['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'quant']
                    df_export.to_excel(writer, index=False, sheet_name='TabelaPreco')
                
                st.download_button(
                    label="📥 Exportar Banco em Excel (.xlsx)",
                    data=buffer_db.getvalue(),
                    file_name=f"banco_dados_{datetime.now().strftime('%d%m%y_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.info("Não há dados cadastrados na base de dados para exportação.")

    if perm_b_visualizar:
        st.divider()
        st.subheader("🔍 Visualização dos Dados Atuais")
        if not master_db.empty:
            st.dataframe(master_db, use_container_width=True)

    if perm_b_exp_desconto or perm_b_exp_importacao:
        # =====================================================================
        # >>> INÍCIO DO NOVO BLOCO 1: EXPORTAR BANCO COM DESCONTO E FILTRO DE ESTOQUE
        # =====================================================================
        st.divider()
        st.subheader("🏷️ Exportar Banco com Desconto e Filtro de Estoque")

        if not master_db.empty:
            def _quant_para_int(v):
                try:
                    return int(float(v))
                except:
                    return 1

            ex_c1, ex_c2 = st.columns(2)
            exp_desconto = ex_c1.number_input(
                "Desconto (%)", min_value=0.0, max_value=100.0, value=0.0, step=0.5, key="exp_desconto_banco"
            )
            exp_arredondar = ex_c2.checkbox("Arredondar preços", value=True, key="exp_arredondar_banco")

            st.write("**Filtros de Estoque Mínimo**")
            ex_c3, ex_c4, ex_c5 = st.columns(3)
            exp_est_geral = ex_c3.number_input(
                "Estoque mínimo Geral:", min_value=0, value=1, key="exp_est_geral_banco"
            )
            exp_est_unit = ex_c4.number_input(
                "Estoque mínimo para Unitários (QUANT = 1):", min_value=0, value=1, key="exp_est_unit_banco"
            )
            exp_est_cx = ex_c5.number_input(
                "Estoque mínimo para Caixas/Dz/Dp/Cj (QUANT > 1):", min_value=0, value=1, key="exp_est_cx_banco"
            )

            df_exp = master_db.copy()
            df_exp['Stock'] = pd.to_numeric(df_exp['Stock'], errors='coerce').fillna(0)
            df_exp['Price'] = pd.to_numeric(df_exp['Price'], errors='coerce').fillna(0.0)
            df_exp['caixa'] = pd.to_numeric(df_exp['caixa'], errors='coerce').fillna(0.0)

            quant_serie = df_exp['QUANT'].apply(_quant_para_int)

            # Mesma regra de estoque mínimo usada na aba de Cotação
            mask_unit = (quant_serie == 1) & (df_exp['Stock'] >= exp_est_unit)
            mask_cx = (quant_serie > 1) & (df_exp['Stock'] >= exp_est_cx)
            mask_geral = (quant_serie < 1) & (df_exp['Stock'] >= exp_est_geral)
            df_exp = df_exp[mask_unit | mask_cx | mask_geral].copy()

            fator_desc = 1 - (exp_desconto / 100)
            df_exp['Price'] = df_exp['Price'] * fator_desc
            df_exp['caixa'] = df_exp['caixa'] * fator_desc
            if exp_arredondar:
                df_exp['Price'] = df_exp['Price'].apply(extra_round)
                df_exp['caixa'] = df_exp['caixa'].apply(extra_round)

            st.write(f"**{len(df_exp)}** de **{len(master_db)}** produtos atendem aos filtros de estoque.")

            if df_exp.empty:
                st.warning("Nenhum produto atende aos filtros de estoque informados.")
            elif perm_b_exp_desconto:
                df_exp_saida = df_exp[['Description', 'Barcode', 'Price', 'Stock', 'caixa', 'QUANT']].copy()
                df_exp_saida.columns = ['descrição', 'codigo barras', 'preço', 'estoque', 'caixa', 'quant']

                with st.expander("👁️ Pré-visualizar exportação (primeiras 50 linhas)"):
                    st.dataframe(df_exp_saida.head(50), use_container_width=True)

                buffer_db_desc = io.BytesIO()
                with pd.ExcelWriter(buffer_db_desc, engine='openpyxl') as writer_desc:
                    df_exp_saida.to_excel(writer_desc, index=False, sheet_name='TabelaPreco')

                st.download_button(
                    label=f"📥 Exportar Banco com {exp_desconto:g}% de desconto (.xlsx)",
                    data=buffer_db_desc.getvalue(),
                    file_name=f"banco_desc_{exp_desconto:g}pct_{datetime.now().strftime('%d%m%y_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    key="btn_download_banco_desconto"
                )
        else:
            st.info("Não há dados cadastrados na base de dados para exportação.")
        # =====================================================================
        # >>> FIM DO NOVO BLOCO 1
        # =====================================================================

    if perm_b_exp_importacao:
        # =====================================================================
        # >>> INÍCIO DO NOVO BLOCO 4: SAÍDA ESPECÍFICA "exemplo-importacao" (CSV SEPARADO POR VÍRGULAS)
        # =====================================================================
        st.divider()
        st.subheader("📄 Exportar Planilha de Importação (exemplo-importacao.csv)")

        if not master_db.empty:
            st.caption(
                "Esta saída usa os mesmos critérios definidos no bloco acima (Desconto, Arredondamento e Filtros de Estoque Mínimo). "
                "Colunas geradas, nesta ordem: **EAN**, **DUN**, **Qtd. Embalagem** e **Vlr. Custo**. "
                "Formato igual ao arquivo de exemplo: colunas separadas por ponto e vírgula (;), preço com vírgula decimal e sem aspas."
            )

            if df_exp.empty:
                st.warning("Nenhum produto atende aos filtros de estoque informados.")
            else:
                def _formatar_preco_virgula(v):
                    try:
                        num_preco = float(v)
                    except (TypeError, ValueError):
                        return ""
                    texto_preco = f"{num_preco:.2f}" if exp_arredondar else format(Decimal(str(num_preco)), 'f')
                    if '.' in texto_preco:
                        texto_preco = texto_preco.rstrip('0').rstrip('.')
                    return texto_preco.replace('.', ',')

                df_importacao = pd.DataFrame({
                    'EAN': df_exp['Barcode'].astype(str),
                    'DUN': df_exp['Barcode'].astype(str),
                    'Qtd. Embalagem': 1,
                    'Vlr. Custo': df_exp['Price'].apply(_formatar_preco_virgula)
                })

                st.write(f"**{len(df_importacao)}** produtos serão exportados na planilha de importação.")

                with st.expander("👁️ Pré-visualizar planilha de importação (primeiras 50 linhas)"):
                    st.dataframe(df_importacao.head(50), use_container_width=True)

                # Formato idêntico ao arquivo de exemplo: separador ";" (padrão do CSV no Excel em português),
                # preço com vírgula decimal sem aspas e quebra de linha no padrão Windows (CRLF).
                import csv as csv_padrao
                buffer_csv_imp = io.StringIO()
                escritor_csv_imp = csv_padrao.writer(buffer_csv_imp, delimiter=";", lineterminator="\r\n")
                escritor_csv_imp.writerow(list(df_importacao.columns))
                for linha_imp in df_importacao.itertuples(index=False, name=None):
                    escritor_csv_imp.writerow(linha_imp)
                csv_importacao = buffer_csv_imp.getvalue().encode('utf-8')

                st.download_button(
                    label=f"📥 Exportar exemplo-importacao.csv com {exp_desconto:g}% de desconto",
                    data=csv_importacao,
                    file_name="exemplo-importacao.csv",
                    mime="text/csv",
                    type="primary",
                    key="btn_download_exemplo_importacao"
                )
        else:
            st.info("Não há dados cadastrados na base de dados para exportação.")
        # =====================================================================
        # >>> FIM DO NOVO BLOCO 4
        # =====================================================================

    if perm_b_editar:
        # =====================================================================
        # >>> INÍCIO DO NOVO BLOCO 2: EDITAR O BANCO DE DADOS DIRETAMENTE NA TELA
        # =====================================================================
        st.divider()
        st.subheader("✏️ Editar Banco de Dados")
        st.caption(
            "Clique em qualquer célula para corrigir. Para adicionar uma linha, use a última linha em branco da tabela. "
            "Para excluir, selecione a linha (caixa à esquerda) e pressione Delete. "
            "As alterações só valem depois de clicar em **Salvar Alterações no Banco**."
        )

        if 'editor_banco_versao' not in st.session_state:
            st.session_state.editor_banco_versao = 0

        df_para_editar = master_db.copy()
        df_para_editar['Description'] = df_para_editar['Description'].fillna("").astype(str)
        df_para_editar['Barcode'] = df_para_editar['Barcode'].astype(str)
        df_para_editar['Price'] = pd.to_numeric(df_para_editar['Price'], errors='coerce').fillna(0.0)
        df_para_editar['caixa'] = pd.to_numeric(df_para_editar['caixa'], errors='coerce').fillna(0.0)
        df_para_editar['Stock'] = pd.to_numeric(df_para_editar['Stock'], errors='coerce').fillna(0).astype(int)
        df_para_editar['QUANT'] = pd.to_numeric(df_para_editar['QUANT'], errors='coerce').fillna(1).astype(int)
        df_para_editar = df_para_editar[['Description', 'Barcode', 'Price', 'Stock', 'caixa', 'QUANT']].reset_index(drop=True)

        df_editado = st.data_editor(
            df_para_editar,
            num_rows="dynamic",
            use_container_width=True,
            key=f"editor_banco_{st.session_state.editor_banco_versao}",
            column_config={
                "Description": st.column_config.TextColumn("Descrição"),
                "Barcode": st.column_config.TextColumn("Código de Barras"),
                "Price": st.column_config.NumberColumn("Preço", format="%.2f", min_value=0.0),
                "Stock": st.column_config.NumberColumn("Estoque", step=1),
                "caixa": st.column_config.NumberColumn("Caixa", format="%.2f", min_value=0.0),
                "QUANT": st.column_config.NumberColumn("Quant", step=1, min_value=1),
            }
        )

        if st.button("💾 Salvar Alterações no Banco", type="primary", key="btn_salvar_edicao_banco"):
            df_salvar = df_editado.copy()
            df_salvar['Description'] = df_salvar['Description'].fillna("").astype(str).str.strip()
            df_salvar['Barcode'] = df_salvar['Barcode'].apply(
                lambda x: re.sub(r'\D', '', str(x).split('.')[0]) if pd.notna(x) else ""
            )
            df_salvar['Price'] = pd.to_numeric(df_salvar['Price'], errors='coerce').fillna(0.0)
            df_salvar['caixa'] = pd.to_numeric(df_salvar['caixa'], errors='coerce').fillna(0.0)
            df_salvar['Stock'] = pd.to_numeric(df_salvar['Stock'], errors='coerce').fillna(0).astype(int)
            df_salvar['QUANT'] = pd.to_numeric(df_salvar['QUANT'], errors='coerce').fillna(1).astype(int)

            qtd_sem_barras = int((df_salvar['Barcode'] == "").sum())

            if qtd_sem_barras > 0:
                st.error(
                    f"Existem {qtd_sem_barras} linha(s) sem código de barras válido. "
                    "Preencha ou exclua essas linhas antes de salvar."
                )
            else:
                qtd_duplicados = int(df_salvar['Barcode'].duplicated(keep=False).sum())

                # Cópia de segurança do banco antes de sobrescrever
                if os.path.exists(DB_STORAGE):
                    master_db.to_csv("master_database_backup.csv", index=False)

                df_salvar = df_salvar[['Description', 'Barcode', 'Price', 'Stock', 'caixa', 'QUANT']]
                df_salvar.to_csv(DB_STORAGE, index=False)
                st.cache_data.clear()
                st.session_state.editor_banco_versao += 1

                if qtd_duplicados > 0:
                    st.toast(f"Banco salvo com {len(df_salvar)} produtos. Atenção: {qtd_duplicados} linhas com código de barras repetido.", icon="⚠️")
                else:
                    st.toast(f"Banco salvo com sucesso! {len(df_salvar)} produtos.", icon="✅")
                st.rerun()
        # =====================================================================
        # >>> FIM DO NOVO BLOCO 2
        # =====================================================================

    if perm_b_recalcular:
        # =====================================================================
        # >>> INÍCIO DO NOVO BLOCO 3: RECALCULAR PREÇO (CAIXA ÷ QUANT) NAS LINHAS ALTERADAS
        # =====================================================================
        st.markdown("---")
        st.subheader("🔄 Recalcular Preço (Caixa ÷ Quant)")
        st.caption(
            "Recalcula o Preço (Preço = Caixa ÷ Quant, a mesma regra da importação) somente nas linhas em que você "
            "alterou **Caixa** ou **Quant** na tabela acima, e também nas linhas novas sem preço informado. "
            "Linhas em que você editou o **Preço** manualmente não são recalculadas. "
            "Ao clicar no botão, o preço é recalculado e o banco é salvo."
        )

        def _indices_para_recalcular(df_original, df_atual):
            comuns = df_atual.index.intersection(df_original.index)
            novos = df_atual.index.difference(df_original.index)

            alvo_alterados = []
            if len(comuns) > 0:
                o = df_original.loc[comuns]
                a = df_atual.loc[comuns]
                caixa_a = pd.to_numeric(a['caixa'], errors='coerce').fillna(0.0)
                quant_a = pd.to_numeric(a['QUANT'], errors='coerce').fillna(1)
                preco_a = pd.to_numeric(a['Price'], errors='coerce').fillna(0.0)
                caixa_mudou = (caixa_a - o['caixa']).abs() > 1e-9
                quant_mudou = (quant_a - o['QUANT']).abs() > 1e-9
                preco_mudou = (preco_a - o['Price']).abs() > 1e-9
                alvo_alterados = list(comuns[((caixa_mudou | quant_mudou) & ~preco_mudou).to_numpy()])

            alvo_novos = []
            if len(novos) > 0:
                preco_novos = pd.to_numeric(df_atual.loc[novos, 'Price'], errors='coerce').fillna(0.0)
                alvo_novos = list(novos[(preco_novos == 0).to_numpy()])

            return alvo_alterados + alvo_novos

        def _preco_recalculado(caixa, quant):
            c_val = pd.to_numeric(caixa, errors='coerce')
            c_val = float(c_val) if not pd.isna(c_val) else 0.0
            q_val = pd.to_numeric(quant, errors='coerce')
            q_val = int(q_val) if not pd.isna(q_val) else 1
            return c_val / q_val if q_val > 0 else 0.0

        rc_indices = _indices_para_recalcular(df_para_editar, df_editado)

        if rc_indices:
            rc_previa = []
            for i_rc in rc_indices:
                lin_rc = df_editado.loc[i_rc]
                rc_previa.append({
                    "Descrição": lin_rc['Description'],
                    "Código de Barras": lin_rc['Barcode'],
                    "Caixa": lin_rc['caixa'],
                    "Quant": lin_rc['QUANT'],
                    "Preço atual": lin_rc['Price'],
                    "Preço recalculado": _preco_recalculado(lin_rc['caixa'], lin_rc['QUANT'])
                })
            st.write(f"**{len(rc_indices)}** linha(s) serão recalculadas ao clicar no botão (exibindo até 50):")
            st.dataframe(pd.DataFrame(rc_previa).head(50), use_container_width=True)
        else:
            st.info("Nenhuma linha com Caixa ou Quant alterados no momento.")

        if st.button(
            "🔄 Recalcular Preço nas Linhas Alteradas e Salvar",
            type="primary",
            key="btn_recalcular_preco_banco",
            disabled=(len(rc_indices) == 0)
        ):
            df_rc = df_editado.copy()
            df_rc['Description'] = df_rc['Description'].fillna("").astype(str).str.strip()
            df_rc['Barcode'] = df_rc['Barcode'].apply(
                lambda x: re.sub(r'\D', '', str(x).split('.')[0]) if pd.notna(x) else ""
            )
            df_rc['Price'] = pd.to_numeric(df_rc['Price'], errors='coerce').fillna(0.0)
            df_rc['caixa'] = pd.to_numeric(df_rc['caixa'], errors='coerce').fillna(0.0)
            df_rc['Stock'] = pd.to_numeric(df_rc['Stock'], errors='coerce').fillna(0).astype(int)
            df_rc['QUANT'] = pd.to_numeric(df_rc['QUANT'], errors='coerce').fillna(1).astype(int)

            for i_rc in rc_indices:
                df_rc.loc[i_rc, 'Price'] = _preco_recalculado(df_rc.loc[i_rc, 'caixa'], df_rc.loc[i_rc, 'QUANT'])

            rc_sem_barras = int((df_rc['Barcode'] == "").sum())

            if rc_sem_barras > 0:
                st.error(
                    f"Existem {rc_sem_barras} linha(s) sem código de barras válido. "
                    "Preencha ou exclua essas linhas antes de salvar."
                )
            else:
                rc_duplicados = int(df_rc['Barcode'].duplicated(keep=False).sum())

                # Cópia de segurança do banco antes de sobrescrever
                if os.path.exists(DB_STORAGE):
                    master_db.to_csv("master_database_backup.csv", index=False)

                df_rc = df_rc[['Description', 'Barcode', 'Price', 'Stock', 'caixa', 'QUANT']]
                df_rc.to_csv(DB_STORAGE, index=False)
                st.cache_data.clear()
                st.session_state.editor_banco_versao += 1

                if rc_duplicados > 0:
                    st.toast(f"{len(rc_indices)} preço(s) recalculado(s) e banco salvo. Atenção: {rc_duplicados} linhas com código de barras repetido.", icon="⚠️")
                else:
                    st.toast(f"{len(rc_indices)} preço(s) recalculado(s) e banco salvo com sucesso!", icon="✅")
                st.rerun()
        # =====================================================================
        # >>> FIM DO NOVO BLOCO 3
        # =====================================================================

    # =====================================================================
    # >>> NOVO BLOCO 5 (parte final): AVISO QUANDO NENHUMA FUNÇÃO ESTÁ LIBERADA
    # =====================================================================
    if not any([perm_b_atualizar, perm_b_exportar, perm_b_visualizar, perm_b_exp_desconto, perm_b_exp_importacao, perm_b_editar]):
        st.info("Seu usuário não possui permissão para nenhuma função do Gerenciador de Banco. Contate o administrador.")
    # =====================================================================

# --- ABA 4: USUÁRIOS ---
elif aba == "👤 Usuários":
    st.title("👤 Gestão de Usuários e Permissões")
    
    c_b1, c_b2 = st.columns(2)
    with c_b1:
        df_u = pd.DataFrame.from_dict(users, orient='index').reset_index()
        st.download_button("📥 Backup Usuários", df_u.to_csv(index=False).encode('utf-8'), "backup_users.csv")
    with c_b2:
        up_b = st.file_uploader("Restore Usuários", type=["csv"])
        if up_b and st.button("🔥 Restaurar"):
            df_r = pd.read_csv(up_b)
            new_u = {}
            for _, r in df_r.iterrows():
                try: perm = json.loads(r['permissions'].replace("'", '"'))
                except: perm = {"ver_estoque": True, "ver_preco": True, "ver_caixa": True, "ver_quant": True, "funcao_cotacao": True, "funcao_vendas": True, "busca_hibrido": True, "busca_barras": True, "busca_similaridade": True, "ver_gerenciar_banco": True}
                new_u[str(r['index'])] = {"password": str(r['password']), "expiry": str(r['expiry']), "role": str(r['role']), "permissions": perm}
            save_users(new_u)
            st.rerun()
            
    st.divider()
    col_n, col_e, col_l = st.columns(3)
    
    with col_n:
        st.subheader("➕ Criar Novo Usuário")
        with st.form("Novo_Form"):
            nu = st.text_input("Nome do Usuário")
            np = st.text_input("Senha de Acesso")
            nd = st.number_input("Validade da Conta (dias)", 1, 365, 30)
            
            st.markdown("---")
            st.write("📋 **Permissões de Campos e Funções:**")
            p_estoque = st.checkbox("Acesso à Coluna: Estoque", value=True)
            p_preco = st.checkbox("Acesso à Coluna: Preço", value=True)
            p_caixa = st.checkbox("Acesso à Coluna: Caixa", value=True)
            p_quant = st.checkbox("Acesso à Coluna: Quantidade (Família)", value=True)
            p_cotacao = st.checkbox("Permitir Função: Processar Cotações", value=True)
            p_vendas = st.checkbox("Permitir Função: Consulta de Vendas / Pedido", value=True)
            p_gerenciar = st.checkbox("Permitir Aba: Gerenciar Banco de Dados", value=True)
            
            st.write("🔍 **Permissões de Regras de Busca:**")
            p_hibrido = st.checkbox("Regra: Híbrido (Barras + Similaridade)", value=True)
            p_barras = st.checkbox("Regra: Apenas Barras", value=True)
            p_similaridade = st.checkbox("Regra: Apenas Similaridade", value=True)

            st.write("🗄️ **Permissões das Funções do Gerenciador de Banco:**")
            st.caption("Valem somente se a opção 'Permitir Aba: Gerenciar Banco de Dados' estiver marcada.")
            p_bk_atualizar = st.checkbox("Banco: Atualizar Base de Dados (upload)", value=True, key="n_bk_atualizar")
            p_bk_exportar = st.checkbox("Banco: Exportar Base de Dados", value=True, key="n_bk_exportar")
            p_bk_visualizar = st.checkbox("Banco: Visualizar Dados Atuais", value=True, key="n_bk_visualizar")
            p_bk_exp_desconto = st.checkbox("Banco: Exportar com Desconto e Filtro de Estoque", value=True, key="n_bk_exp_desconto")
            p_bk_exp_importacao = st.checkbox("Banco: Exportar Planilha de Importação (CSV)", value=True, key="n_bk_exp_importacao")
            p_bk_editar = st.checkbox("Banco: Editar Banco de Dados", value=True, key="n_bk_editar")
            p_bk_recalcular = st.checkbox("Banco: Recalcular Preço (exige também 'Editar Banco de Dados')", value=True, key="n_bk_recalcular")
            
            if st.form_submit_button("Criar Usuário com Permissões"):
                if nu and np:
                    users[nu] = {
                        "password": np,
                        "expiry": (datetime.now() + timedelta(days=nd)).strftime("%Y-%m-%d"),
                        "role": "user",
                        "permissions": {
                            "ver_estoque": p_estoque,
                            "ver_preco": p_preco,
                            "ver_caixa": p_caixa,
                            "ver_quant": p_quant,
                            "funcao_cotacao": p_cotacao,
                            "funcao_vendas": p_vendas,
                            "busca_hibrido": p_hibrido,
                            "busca_barras": p_barras,
                            "busca_similaridade": p_similaridade,
                            "ver_gerenciar_banco": p_gerenciar
                        }
                    }
                    users[nu]["permissions"].update({
                        "banco_atualizar": p_bk_atualizar,
                        "banco_exportar": p_bk_exportar,
                        "banco_visualizar": p_bk_visualizar,
                        "banco_exportar_desconto": p_bk_exp_desconto,
                        "banco_exportar_importacao": p_bk_exp_importacao,
                        "banco_editar": p_bk_editar,
                        "banco_recalcular": p_bk_recalcular
                    })
                    save_users(users)
                    st.success(f"Usuário {nu} criado com sucesso!")
                    st.rerun()
                else:
                    st.error("Preencha o Usuário e Senha!")

    with col_e:
        st.subheader("📝 Editar Usuário Existente")
        u_sel = st.selectbox("Selecione o Usuário para Modificar", list(users.keys()))
        if u_sel:
            current_perms = users[u_sel].get("permissions", {
                "ver_estoque": True, "ver_preco": True, "ver_caixa": True, "ver_quant": True, "funcao_cotacao": True, "funcao_vendas": True,
                "busca_hibrido": True, "busca_barras": True, "busca_similaridade": True, "ver_gerenciar_banco": True
            })
            
            ep = st.text_input("Senha Atual/Nova", value=users[u_sel]["password"])
            ex = st.text_input("Data de Expiração (AAAA-MM-DD)", value=users[u_sel]["expiry"])
            
            st.markdown("---")
            st.write(f"📋 **Modificar Permissões de '{u_sel}':**")
            
            up_estoque = st.checkbox("Acesso à Coluna: Estoque", value=current_perms.get("ver_estoque", True), key="e_est")
            up_preco = st.checkbox("Acesso à Coluna: Preço", value=current_perms.get("ver_preco", True), key="e_prc")
            up_caixa = st.checkbox("Acesso à Coluna: Caixa", value=current_perms.get("ver_caixa", True), key="e_cx")
            up_quant = st.checkbox("Acesso à Coluna: Quantidade (Família)", value=current_perms.get("ver_quant", True), key="e_qnt")
            up_cotacao = st.checkbox("Permitir Função: Processar Cotações", value=current_perms.get("funcao_cotacao", True), key="e_cot")
            up_vendas = st.checkbox("Permitir Função: Consulta de Vendas / Pedido", value=current_perms.get("funcao_vendas", True), key="e_vnd")
            up_gerenciar = st.checkbox("Permitir Aba: Gerenciar Banco de Dados", value=current_perms.get("ver_gerenciar_banco", True), key="e_ger")
            
            st.write("🔍 **Modificar Regras de Busca:**")
            up_hibrido = st.checkbox("Regra: Híbrido (Barras + Similaridade)", value=current_perms.get("busca_hibrido", True), key="e_hib")
            up_barras = st.checkbox("Regra: Apenas Barras", value=current_perms.get("busca_barras", True), key="e_bar")
            up_similaridade = st.checkbox("Regra: Apenas Similaridade", value=current_perms.get("busca_similaridade", True), key="e_sim")

            st.write("🗄️ **Permissões das Funções do Gerenciador de Banco:**")
            st.caption("Valem somente se a opção 'Permitir Aba: Gerenciar Banco de Dados' estiver marcada.")
            up_bk_atualizar = st.checkbox("Banco: Atualizar Base de Dados (upload)", value=current_perms.get("banco_atualizar", True), key=f"e_bk_atualizar_{u_sel}")
            up_bk_exportar = st.checkbox("Banco: Exportar Base de Dados", value=current_perms.get("banco_exportar", True), key=f"e_bk_exportar_{u_sel}")
            up_bk_visualizar = st.checkbox("Banco: Visualizar Dados Atuais", value=current_perms.get("banco_visualizar", True), key=f"e_bk_visualizar_{u_sel}")
            up_bk_exp_desconto = st.checkbox("Banco: Exportar com Desconto e Filtro de Estoque", value=current_perms.get("banco_exportar_desconto", True), key=f"e_bk_exp_desconto_{u_sel}")
            up_bk_exp_importacao = st.checkbox("Banco: Exportar Planilha de Importação (CSV)", value=current_perms.get("banco_exportar_importacao", True), key=f"e_bk_exp_importacao_{u_sel}")
            up_bk_editar = st.checkbox("Banco: Editar Banco de Dados", value=current_perms.get("banco_editar", True), key=f"e_bk_editar_{u_sel}")
            up_bk_recalcular = st.checkbox("Banco: Recalcular Preço (exige também 'Editar Banco de Dados')", value=current_perms.get("banco_recalcular", True), key=f"e_bk_recalcular_{u_sel}")
            
            if st.button("Salvar Alterações do Usuário"):
                users[u_sel].update({
                    "password": ep,
                    "expiry": ex,
                    "permissions": {
                        "ver_estoque": up_estoque,
                        "ver_preco": up_preco,
                        "ver_caixa": up_caixa,
                        "ver_quant": up_quant,
                        "funcao_cotacao": up_cotacao,
                        "funcao_vendas": up_vendas,
                        "busca_hibrido": up_hibrido,
                        "busca_barras": up_barras,
                        "busca_similaridade": up_similaridade,
                        "ver_gerenciar_banco": up_gerenciar
                    }
                })
                users[u_sel]["permissions"].update({
                    "banco_atualizar": up_bk_atualizar,
                    "banco_exportar": up_bk_exportar,
                    "banco_visualizar": up_bk_visualizar,
                    "banco_exportar_desconto": up_bk_exp_desconto,
                    "banco_exportar_importacao": up_bk_exp_importacao,
                    "banco_editar": up_bk_editar,
                    "banco_recalcular": up_bk_recalcular
                })
                save_users(users)
                st.success(f"Permissões e dados de '{u_sel}' atualizados!")
                st.rerun()
                
            if u_sel != "admin" and st.button("❌ Excluir Usuário permanentemente"):
                del users[u_sel]
                save_users(users)
                st.rerun()

    with col_l:
        st.subheader("📦 Edição em Lote")
        lista_usuarios_lote = [u for u in list(users.keys()) if u != "admin"]
        
        usuarios_selecionados = st.multiselect(
            "Selecione os Usuários para aplicar em Lote", 
            options=lista_usuarios_lote,
            placeholder="Escolha um ou mais usuários"
        )
        
        if usuarios_selecionados:
            st.markdown("---")
            st.write("📋 **Permissões que serão replicadas nos selecionados:**")
            l_estoque = st.checkbox("Acesso à Coluna: Estoque", value=True, key="l_est")
            l_preco = st.checkbox("Acesso à Coluna: Preço", value=True, key="l_prc")
            l_caixa = st.checkbox("Acesso à Coluna: Caixa", value=True, key="l_cx")
            l_quant = st.checkbox("Acesso à Coluna: Quantidade (Família)", value=True, key="l_qnt")
            l_cotacao = st.checkbox("Permitir Função: Processar Cotações", value=True, key="l_cot")
            l_vendas = st.checkbox("Permitir Função: Consulta de Vendas / Pedido", value=True, key="l_vnd")
            l_gerenciar = st.checkbox("Permitir Aba: Gerenciar Banco de Dados", value=True, key="l_ger")
            
            st.write("🔍 **Regras de Busca que serão replicadas:**")
            l_hibrido = st.checkbox("Regra: Híbrido (Barras + Similaridade)", value=True, key="l_hib")
            l_barras = st.checkbox("Regra: Apenas Barras", value=True, key="l_bar")
            l_similaridade = st.checkbox("Regra: Apenas Similaridade", value=True, key="l_sim")

            st.write("🗄️ **Permissões das Funções do Gerenciador de Banco:**")
            st.caption("Valem somente se a opção 'Permitir Aba: Gerenciar Banco de Dados' estiver marcada.")
            l_bk_atualizar = st.checkbox("Banco: Atualizar Base de Dados (upload)", value=True, key="l_bk_atualizar")
            l_bk_exportar = st.checkbox("Banco: Exportar Base de Dados", value=True, key="l_bk_exportar")
            l_bk_visualizar = st.checkbox("Banco: Visualizar Dados Atuais", value=True, key="l_bk_visualizar")
            l_bk_exp_desconto = st.checkbox("Banco: Exportar com Desconto e Filtro de Estoque", value=True, key="l_bk_exp_desconto")
            l_bk_exp_importacao = st.checkbox("Banco: Exportar Planilha de Importação (CSV)", value=True, key="l_bk_exp_importacao")
            l_bk_editar = st.checkbox("Banco: Editar Banco de Dados", value=True, key="l_bk_editar")
            l_bk_recalcular = st.checkbox("Banco: Recalcular Preço (exige também 'Editar Banco de Dados')", value=True, key="l_bk_recalcular")
            
            if st.button("⚡ Aplicar Permissões em Lote", type="primary"):
                for u in usuarios_selecionados:
                    if u in users:
                        users[u]["permissions"] = {
                            "ver_estoque": l_estoque,
                            "ver_preco": l_preco,
                            "ver_caixa": l_caixa,
                            "ver_quant": l_quant,
                            "funcao_cotacao": l_cotacao,
                            "funcao_vendas": l_vendas,
                            "busca_hibrido": l_hibrido,
                            "busca_barras": l_barras,
                            "busca_similaridade": l_similaridade,
                            "ver_gerenciar_banco": l_gerenciar
                        }
                        users[u]["permissions"].update({
                            "banco_atualizar": l_bk_atualizar,
                            "banco_exportar": l_bk_exportar,
                            "banco_visualizar": l_bk_visualizar,
                            "banco_exportar_desconto": l_bk_exp_desconto,
                            "banco_exportar_importacao": l_bk_exp_importacao,
                            "banco_editar": l_bk_editar,
                            "banco_recalcular": l_bk_recalcular
                        })
                save_users(users)
                st.success(f"Permissões aplicadas com sucesso para {len(usuarios_selecionados)} usuários!")
                st.rerun()
        else:
            st.info("Selecione pelo menos um usuário para abrir a lista de edição em lote.")

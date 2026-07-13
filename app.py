import streamlit as st
import sqlite3
import re
import pandas as pd
import os
from datetime import datetime
import pypdf

# BIBLIOTECAS PARA RELATÓRIOS
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# CONFIGURAÇÃO DE CAMINHO AUTOMÁTICO
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
CAMINHO_BANCO = os.path.join(DIRETORIO_ATUAL, "diarias_sistema.db")
CAMINHO_BRASAO = os.path.join(DIRETORIO_ATUAL, "brasao_pr.png")

# Pasta de empenhos
PASTA_EMPENHOS = os.path.join(DIRETORIO_ATUAL, "arquivos_empenhos")
if not os.path.exists(PASTA_EMPENHOS):
    os.makedirs(PASTA_EMPENHOS)

# 1. CONFIGURAÇÃO DO BANCO DE DADOS
def conectar_banco():
    conn = sqlite3.connect(CAMINHO_BANCO)
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS servidores (
            id INTEGER PRIMARY KEY,
            nome_completo TEXT NOT NULL,
            cpf TEXT UNIQUE NOT NULL,
            unidade_lotacao TEXT NOT NULL,
            banco_codigo TEXT NOT NULL,
            agencia TEXT NOT NULL,
            conta_corrente TEXT NOT NULL,
            status TEXT DEFAULT 'Ativo'
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS unidades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_unidade TEXT UNIQUE NOT NULL
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS bancos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            nome_banco TEXT NOT NULL
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS empenhos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            servidor_nome TEXT,
            numero_empenho TEXT UNIQUE NOT NULL,
            valor_total REAL NOT NULL,
            valor_disponivel REAL NOT NULL,
            data_empenho TEXT,
            caminho_pdf TEXT
        )
    """)
    
    cursor.execute("PRAGMA table_info(empenhos)")
    colunas = [coluna[1] for coluna in cursor.fetchall()]
    
    if "servidor_nome" not in colunas:
        cursor.execute("ALTER TABLE empenhos ADD COLUMN servidor_nome TEXT")
    if "data_empenho" not in colunas:
        cursor.execute("ALTER TABLE empenhos ADD COLUMN data_empenho TEXT")
    if "caminho_pdf" not in colunas:
        cursor.execute("ALTER TABLE empenhos ADD COLUMN caminho_pdf TEXT")
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS diarias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            servidor_id INTEGER NOT NULL,
            empenho_id INTEGER NOT NULL,
            data_diaria TEXT NOT NULL,
            jornada TEXT NOT NULL,
            valor_diaria REAL NOT NULL,
            FOREIGN KEY(servidor_id) REFERENCES servidores(id),
            FOREIGN KEY(empenho_id) REFERENCES empenhos(id)
        )
    """)
    
    cursor.execute("SELECT COUNT(*) FROM bancos")
    if cursor.fetchone()[0] == 0:
        bancos_iniciais = [
            ("001", "BANCO DO BRASIL S.A."),
            ("104", "CAIXA ECONOMICA CEF"),
            ("341", "ITAU UNIBANCO S.A."),
            ("237", "BANCO BRADESCO S.A."),
            ("033", "BANCO SANTANDER (BRASIL) S.A.")
        ]
        cursor.executemany("INSERT INTO bancos (codigo, nome_banco) VALUES (?, ?)", bancos_iniciais)
        
    cursor.execute("SELECT COUNT(*) FROM unidades")
    if cursor.fetchone()[0] == 0:
        unidades_iniciais = [("Sede DEPEN"), ("CCCE"), ("CPA"), ("Complexo Piraquara")]
        cursor.executemany("INSERT INTO unidades (nome_unidade) VALUES (?)", [(u,) for u in unidades_iniciais])
    
    conn.commit()
    return conn, cursor

# Extração de PDF
def extrair_dados_pdf_pr(arquivo_pdf):
    try:
        leitor = pypdf.PdfReader(arquivo_pdf)
        texto = ""
        for pagina in leitor.pages:
            texto += pagina.extract_text() + "\n"
            
        dados = {"empenho": "", "data": datetime.now().date(), "servidor": "", "valor": 0.0}
        
        match_emp = re.search(r'(\d{4}NE\d{6})', texto)
        if match_emp:
            dados["empenho"] = match_emp.group(1).strip().upper()
            
        match_dt = re.search(r'\d{4}NE\d{6}\s+(\d{2}/\d{2}/\d{2,4})', texto)
        if match_dt:
            dt_str = match_dt.group(1)
            try:
                if len(dt_str.split('/')[-1]) == 2:
                    dados["data"] = datetime.strptime(dt_str, "%d/%m/%y").date()
                else:
                    dados["data"] = datetime.strptime(dt_str, "%d/%m/%Y").date()
            except:
                pass
                
        match_cred = re.search(r'Credor\s+\d+\s+-\s+([A-Z\sÀ-Ú]+)', texto)
        if match_cred:
            dados["servidor"] = match_cred.group(1).strip().upper()
            
        match_vl = re.search(r'Valor\s+([\d\.,]+)', texto)
        if match_vl:
            vl_str = match_vl.group(1).replace('.', '').replace(',', '.')
            dados["valor"] = float(vl_str)
            
        return dados
    except Exception as e:
        st.error(f"Erro ao processar a leitura do PDF: {e}")
        return None

# Funções auxiliares de listagem
def listar_bancos_cadastrados():
    conn, cursor = conectar_banco()
    cursor.execute("SELECT codigo, nome_banco FROM bancos ORDER BY codigo ASC")
    bancos = [f"{linha[0]} - {linha[1]}" for linha in cursor.fetchall()]
    conn.close()
    return bancos

def buscar_unidades_dict():
    conn, cursor = conectar_banco()
    cursor.execute("SELECT id, nome_unidade FROM unidades ORDER BY nome_unidade ASC")
    unidades = {linha[1]: linha[0] for linha in cursor.fetchall()}
    conn.close()
    return unidades

def listar_servidores_selecao():
    conn, cursor = conectar_banco()
    cursor.execute("SELECT id, nome_completo FROM servidores WHERE status = 'Ativo' ORDER BY nome_completo ASC")
    servidores = [{ "id": id_linha, "label": nome_linha } for id_linha, nome_linha in cursor.fetchall()]
    conn.close()
    return servidores

def listar_todos_empenhos():
    conn, cursor = conectar_banco()
    cursor.execute("SELECT id, numero_empenho, servidor_nome FROM empenhos ORDER BY numero_empenho ASC")
    linhas = cursor.fetchall()
    conn.close()
    return linhas

def listar_empenhos_com_saldo():
    conn, cursor = conectar_banco()
    cursor.execute("SELECT id, numero_empenho, valor_disponivel, servidor_nome FROM empenhos WHERE valor_disponivel > 0 ORDER BY numero_empenho ASC")
    linhas = cursor.fetchall()
    conn.close()
    
    empenhos = []
    for Proxy_linha in linhas:
        id_emp = Proxy_linha[0]
        num_emp = Proxy_linha[1]
        saldo = Proxy_linha[2]
        servidor = Proxy_linha[3] if Proxy_linha[3] else "Sem vínculo"
        saldo_formatado = f"R$ {saldo:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        empenhos.append({
            "id": id_emp,
            "label": f"NE: {num_emp} | {servidor} | Saldo: {saldo_formatado}"
        })
    return empenhos

# Formatação numérica no padrão brasileiro (sem "R$")
def formatar_valor_br(valor):
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# Formatação com "R$" (usada na planilha DEAEV REALIZADA)
def formatar_valor_br_moeda(valor):
    return f"R$ {formatar_valor_br(valor)}"

# Nomes dos meses por extenso
MESES_EXTENSO = {
    "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL",
    "05": "MAIO", "06": "JUNHO", "07": "JULHO", "08": "AGOSTO",
    "09": "SETEMBRO", "10": "OUTUBRO", "11": "NOVEMBRO", "12": "DEZEMBRO"
}

# CONTROLE DE MENU ATIVO
if "menu_ativo" not in st.session_state:
    st.session_state.menu_ativo = "inicio"

# Estados da leitura de PDF
if "pdf_num_empenho" not in st.session_state:
    st.session_state.pdf_num_empenho = ""
if "pdf_data_empenho" not in st.session_state:
    st.session_state.pdf_data_empenho = datetime.now().date()
if "pdf_nome_servidor" not in st.session_state:
    st.session_state.pdf_nome_servidor = ""
if "pdf_valor_empenho" not in st.session_state:
    st.session_state.pdf_valor_empenho = 0.0

st.set_page_config(page_title="Sistema de Diárias - Polícia Penal", layout="wide")

# CSS customizado para os botões do menu
st.markdown("""
    <style>
    div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stButton"] button {
        width: 210px !important;
        min-width: 210px !important;
        max-width: 210px !important;
        height: 50px !important;
        text-align: left !important;
        padding-left: 18px !important;
        white-space: nowrap !important;
        display: block !important;
    }
    div[data-testid="stHorizontalBlock"] div[data-testid="stButton"] button {
        width: 160px !important;
        min-width: 160px !important;
        max-width: 160px !important;
    }
    </style>
""", unsafe_allow_html=True)

conectar_banco()

# Topo do Sistema
col_logo, col_titulo = st.columns([1, 11])
with col_logo:
    if os.path.exists(CAMINHO_BRASAO):
        st.image(CAMINHO_BRASAO, width=80)
    else:
        st.subheader("🏛️")
with col_titulo:
    st.markdown("<h1 style='margin-top: -5px; padding-bottom: 0px;'>Sistema de Diárias Extraordinárias</h1>", unsafe_allow_html=True)
    st.markdown("<h4 style='color: #4F8BF9; margin-top: -15px; font-weight: bold; letter-spacing: 1px;'>POLÍCIA PENAL — ESTADO DO PARANÁ</h4>", unsafe_allow_html=True)
    
st.markdown("---")

col_menu_botoes, col_conteudo_dinamico = st.columns([2.5, 9.5])

# MENU LATERAL ESQUERDO
with col_menu_botoes:
    st.write("### Módulos")
    
    tipo_btn_servidores = "primary" if st.session_state.menu_ativo == "servidores" else "secondary"
    if st.button("👥 SERVIDORES", type=tipo_btn_servidores, key="btn_menu_servidores"):
        st.session_state.menu_ativo = "servidores"
        st.rerun()
        
    tipo_btn_empenhos = "primary" if st.session_state.menu_ativo == "empenhos" else "secondary"
    if st.button("📝 EMPENHOS", type=tipo_btn_empenhos, key="btn_menu_empenhos"):
        st.session_state.menu_ativo = "empenhos"
        st.rerun()
        
    tipo_btn_deaev = "primary" if st.session_state.menu_ativo == "deaev" else "secondary"
    if st.button("📋 DEAEV", type=tipo_btn_deaev, key="btn_menu_deaev"):
        st.session_state.menu_ativo = "deaev"
        st.rerun()

    tipo_btn_relatorios = "primary" if st.session_state.menu_ativo == "relatorios" else "secondary"
    if st.button("📊 RELATÓRIOS", type=tipo_btn_relatorios, key="btn_menu_relatorios"):
        st.session_state.menu_ativo = "relatorios"
        st.rerun()
        
    st.markdown("---")
    tipo_btn_sistema = "primary" if st.session_state.menu_ativo == "sistema" else "secondary"
    if st.button("⚙️ SISTEMA", type=tipo_btn_sistema, key="btn_menu_sistema"):
        st.session_state.menu_ativo = "sistema"
        st.rerun()

# CONTEÚDO DINÂMICO CENTRAL
with col_conteudo_dinamico:
    dict_unidades = buscar_unidades_dict()
    lista_unidades = list(dict_unidades.keys())
    lista_bancos = listar_bancos_cadastrados()
    lista_servidores = listar_servidores_selecao()
    
    if st.session_state.menu_ativo == "inicio":
        st.info("💡 Selecione um módulo ao lado esquerdo para exibir as opções e ferramentas na tela.")
        
    # =====================================================================
    # MÓDULO SERVIDORES
    # =====================================================================
    elif st.session_state.menu_ativo == "servidores":
        st.markdown("<h2>👥 Gerenciamento de Servidores</h2>", unsafe_allow_html=True)
        
        conn, cursor = conectar_banco()
        cursor.execute("SELECT id, cpf, nome_completo, unidade_lotacao, banco_codigo, agencia, conta_corrente, status FROM servidores ORDER BY nome_completo ASC")
        servidores_gravados = cursor.fetchall()
        conn.close()
        
        col_cadastro, col_consulta = st.columns([1.1, 1.1])
        servidor_selecionado_df = None
        
        with col_consulta:
            st.markdown("#### 🔍 Consulta de Servidores")
            st.caption("Clique no círculo ao lado esquerdo do CPF para editar ou excluir os dados do servidor.")
            
            if servidores_gravados:
                df_completo = pd.DataFrame(servidores_gravados, columns=["ID", "CPF", "Nome Completo", "Lotação", "Banco", "Agência", "Conta Corrente", "Status"])
                df_exibicao = df_completo[["CPF", "Nome Completo"]]
                
                tabela_interativa = st.dataframe(
                    df_exibicao, 
                    use_container_width=True, 
                    hide_index=True, 
                    height=400,
                    on_select="rerun",
                    selection_mode="single-row"
                )
                
                indices_selecionados = tabela_interativa.get("selection", {}).get("rows", [])
                if indices_selecionados:
                    index_linha = indices_selecionados[0]
                    servidor_selecionado_df = df_completo.iloc[index_linha]
            else:
                st.info("Nenhum servidor cadastrado até o momento.")
                
        with col_cadastro:
            if servidor_selecionado_df is not None:
                st.markdown("#### ✏️ Alterar / Excluir Servidor")
                id_atual = int(servidor_selecionado_df["ID"])
                nome_padrao = str(servidor_selecionado_df["Nome Completo"])
                cpf_padrao = str(servidor_selecionado_df["CPF"])
                
                lot_ori = str(servidor_selecionado_df["Lotação"])
                idx_lotacao = lista_unidades.index(lot_ori) if lot_ori in lista_unidades else 0
                
                bnc_ori = str(servidor_selecionado_df["Banco"])
                idx_banco = 0
                for idx, bnc in enumerate(lista_bancos):
                    if bnc.startswith(bnc_ori):
                        idx_banco = idx
                        break
                        
                agencia_padrao = str(servidor_selecionado_df["Agência"])
                conta_padrao = str(servidor_selecionado_df["Conta Corrente"])
                status_padrao = str(servidor_selecionado_df["Status"])
            else:
                st.markdown("#### ➕ Cadastrar Novo Servidor")
                id_atual = None
                nome_padrao = ""
                cpf_padrao = ""
                idx_lotacao = 0
                idx_banco = 0
                agencia_padrao = ""
                conta_padrao = ""
                status_padrao = "Ativo"

            nome_serv = st.text_input("Nome Completo", value=nome_padrao)
            cpf_serv = st.text_input("CPF", value=cpf_padrao)
            unidade_serv = st.selectbox("Unidade de Lotação", options=lista_unidades if lista_unidades else ["Padrão"], index=idx_lotacao)
            banco_serv = st.selectbox("Banco para depósito", options=lista_bancos, index=idx_banco)
            agencia_serv = st.text_input("Agência", value=agencia_padrao)
            conta_serv = st.text_input("Conta Corrente (Até 13 dígitos)", value=conta_padrao, max_chars=13)
            status_serv = st.selectbox("Status", options=["Ativo", "Inativo"], index=0 if status_padrao == "Ativo" else 1)
            
            if id_atual is None:
                if st.button("➕ Salvar Novo Servidor", type="primary", key="btn_salvar_novo_servidor"):
                    if not nome_serv or not cpf_serv or not conta_serv:
                        st.error("Nome, CPF e Conta Corrente são obrigatórios.")
                    else:
                        conta_formatada = conta_serv.strip().zfill(13)
                        conn, cursor = conectar_banco()
                        try:
                            cursor.execute("""
                                INSERT INTO servidores (nome_completo, cpf, unidade_lotacao, banco_codigo, agencia, conta_corrente, status)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (nome_serv.upper(), cpf_serv.strip(), unidade_serv, banco_serv.split(" - ")[0], agencia_serv.strip(), conta_formatada, status_serv))
                            conn.commit()
                            st.success(f"Servidor cadastrado com sucesso! Conta salva: {conta_formatada}")
                        except sqlite3.IntegrityError:
                            st.error("CPF já existente no banco de dados.")
                        finally:
                            conn.close()
                            st.rerun()
            else:
                c_atualizar, c_deletar = st.columns(2)
                with c_atualizar:
                    if st.button("💾 Atualizar Dados", type="primary", key="btn_atualizar_servidor"):
                        if not nome_serv or not cpf_serv or not conta_serv:
                            st.error("Nome, CPF e Conta Corrente são obrigatórios.")
                        else:
                            conta_formatada = conta_serv.strip().zfill(13)
                            conn, cursor = conectar_banco()
                            try:
                                cursor.execute("""
                                    UPDATE servidores 
                                    SET nome_completo = ?, cpf = ?, unidade_lotacao = ?, banco_codigo = ?, agencia = ?, conta_corrente = ?, status = ?
                                    WHERE id = ?
                                """, (nome_serv.upper(), cpf_serv.strip(), unidade_serv, banco_serv.split(" - ")[0], agencia_serv.strip(), conta_formatada, status_serv, id_atual))
                                conn.commit()
                                st.success("Servidor atualizado com sucesso!")
                            except sqlite3.IntegrityError:
                                st.error("Ocorreu um conflito (CPF duplicado).")
                            finally:
                                conn.close()
                                st.rerun()
                                
                with c_deletar:
                    if st.button("❌ Excluir Servidor", type="secondary", key="btn_excluir_servidor"):
                        conn, cursor = conectar_banco()
                        try:
                            cursor.execute("SELECT COUNT(*) FROM diarias WHERE servidor_id = ?", (id_atual,))
                            possui_diarias = cursor.fetchone()[0]
                            
                            if possui_diarias > 0:
                                st.error("Não é possível excluir este servidor pois ele já possui diárias lançadas no módulo DEAEV.")
                            else:
                                cursor.execute("DELETE FROM servidores WHERE id = ?", (id_atual,))
                                conn.commit()
                                st.success("Servidor removido com sucesso!")
                        except Exception as e:
                            st.error(f"Erro ao remover: {e}")
                        finally:
                            conn.close()
                            st.rerun()

    # =====================================================================
    # MÓDULO EMPENHOS (TRAVA REMOVIDA PARA FASE DE DESENVOLVIMENTO)
    # =====================================================================
    elif st.session_state.menu_ativo == "empenhos":
        st.markdown("<h2>📝 Gerenciamento de Notas de Empenho</h2>", unsafe_allow_html=True)
        st.caption("⚠️ Modo Desenvolvimento: A trava de segurança foi suspensa. Excluir um empenho removerá automaticamente suas diárias.")
        
        if not lista_servidores:
            st.warning("⚠️ Atenção: Não há servidores cadastrados no sistema.")
        else:
            conn, cursor = conectar_banco()
            cursor.execute("SELECT id, numero_empenho, servidor_nome, valor_total, valor_disponivel, data_empenho, caminho_pdf FROM empenhos ORDER BY id DESC")
            empenhos_gravados = cursor.fetchall()
            conn.close()
            
            col_cad_e, col_cons_e = st.columns([1.1, 1.2])
            empenho_selecionado_df = None
            
            with col_cons_e:
                st.markdown("#### 🔍 Consulta de Empenhos")
                st.caption("Clique no círculo ao lado esquerdo do número do Empenho para editar ou excluí-lo.")
                
                if empenhos_gravados:
                    df_empenhos_raw = pd.DataFrame(empenhos_gravados, columns=["ID", "Número Empenho", "Servidor Vinculado", "Valor Inicial Raw", "Saldo Disponível Raw", "Data Raw", "Caminho PDF"])
                    
                    df_exibicao_e = df_empenhos_raw[["Número Empenho", "Servidor Vinculado"]].copy()
                    df_exibicao_e["Saldo Remanescente"] = df_empenhos_raw["Saldo Disponível Raw"].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                    
                    tabela_interativa_e = st.dataframe(
                        df_exibicao_e, 
                        use_container_width=True, 
                        hide_index=True, 
                        height=400,
                        on_select="rerun",
                        selection_mode="single-row",
                        key="tabela_empenhos_interativa"
                    )
                    
                    indices_selecionados_e = tabela_interativa_e.get("selection", {}).get("rows", [])
                    if indices_selecionados_e:
                        index_linha_e = indices_selecionados_e[0]
                        empenho_selecionado_df = df_empenhos_raw.iloc[index_linha_e]
                else:
                    st.info("Nenhum empenho cadastrado.")
                    
            with col_cad_e:
                if empenho_selecionado_df is not None:
                    st.markdown("#### ✏️ Alterar / Excluir Empenho")
                    id_emp_atual = int(empenho_selecionado_df["ID"])
                    num_emp_padrao = str(empenho_selecionado_df["Número Empenho"])
                    val_emp_padrao = float(empenho_selecionado_df["Valor Inicial Raw"])
                    serv_vinc_padrao = str(empenho_selecionado_df["Servidor Vinculado"])
                    
                    dt_str_ori = str(empenho_selecionado_df["Data Raw"])
                    try:
                        dt_padrao = datetime.strptime(dt_str_ori, "%Y-%m-%d").date()
                    except:
                        dt_padrao = datetime.now().date()
                else:
                    st.markdown("#### 🤖 Importação Automática por PDF")
                    arquivo_subido = st.file_uploader("Arraste ou selecione a Nota de Empenho em PDF do Paraná", type=["pdf"])
                    
                    if arquivo_subido is not None:
                        dados_extraidos = extrair_dados_pdf_pr(arquivo_subido)
                        if dados_extraidos:
                            st.session_state.pdf_num_empenho = dados_extraidos["empenho"]
                            st.session_state.pdf_data_empenho = dados_extraidos["data"]
                            st.session_state.pdf_valor_empenho = dados_extraidos["valor"]
                            
                            servidor_encontrado_label = ""
                            for s in lista_servidores:
                                if s["label"] in dados_extraidos["servidor"]:
                                    servidor_encontrado_label = s["label"]
                                    break
                            st.session_state.pdf_nome_servidor = servidor_encontrado_label
                            st.success("⚡ Dados do PDF carregados no formulário!")
                    
                    st.markdown("---")
                    st.markdown("#### ➕ Cadastrar Novo Empenho")
                    id_emp_atual = None
                    num_emp_padrao = st.session_state.pdf_num_empenho
                    val_emp_padrao = st.session_state.pdf_valor_empenho
                    serv_vinc_padrao = st.session_state.pdf_nome_servidor
                    dt_padrao = st.session_state.pdf_data_empenho

                # Campos do Formulário de Empenho
                data_selecionada_e = st.date_input("Data de Emissão", value=dt_padrao, format="DD/MM/YYYY")
                
                opcoes_serv_emp = {s["label"]: s["label"] for s in lista_servidores}
                lista_chaves_servidores = [""] + list(opcoes_serv_emp.keys())
                
                try:
                    idx_selecionado_e = lista_chaves_servidores.index(serv_vinc_padrao)
                except ValueError:
                    idx_selecionado_e = 0
                    
                servidor_responsavel_e = st.selectbox("Nome do Servidor Vinculado", options=lista_chaves_servidores, index=idx_selecionado_e)
                numero_empenho_input = st.text_input("Número do Empenho", value=num_emp_padrao)
                valor_empenho_input = st.number_input("Valor do Empenho (R$)", min_value=0.0, value=val_emp_padrao, format="%.2f")
                
                # Regras de botões para Cadastrar ou Alterar/Excluir
                if id_emp_atual is None:
                    if st.button("➕ Salvar Empenho", type="primary", use_container_width=True):
                        if not servidor_responsavel_e or not numero_empenho_input.strip() or valor_empenho_input <= 0:
                            st.error("❌ Por favor, preencha todos os campos obrigatórios.")
                        else:
                            num_emp_limpo = numero_empenho_input.strip().upper()
                            caminho_salvar_pdf = ""
                            
                            if 'arquivo_subido' in locals() and arquivo_subido is not None:
                                nome_arquivo_fisico = f"{num_emp_limpo}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
                                caminho_salvar_pdf = os.path.join(PASTA_EMPENHOS, nome_arquivo_fisico)
                                with open(caminho_salvar_pdf, "wb") as f:
                                    f.write(arquivo_subido.getbuffer())
                            
                            conn, cursor = conectar_banco()
                            try:
                                cursor.execute("""
                                    INSERT INTO empenhos (servidor_nome, numero_empenho, valor_total, valor_disponivel, data_empenho, caminho_pdf)
                                    VALUES (?, ?, ?, ?, ?, ?)
                                """, (servidor_responsavel_e, num_emp_limpo, valor_empenho_input, valor_empenho_input, str(data_selecionada_e), caminho_salvar_pdf))
                                conn.commit()
                                st.success(f"✔️ Empenho '{num_emp_limpo}' cadastrado!")
                                st.session_state.pdf_num_empenho = ""
                                st.session_state.pdf_data_empenho = datetime.now().date()
                                st.session_state.pdf_nome_servidor = ""
                                st.session_state.pdf_valor_empenho = 0.0
                            except sqlite3.IntegrityError:
                                st.error("❌ Este número de empenho já existe.")
                            finally:
                                conn.close()
                                st.rerun()
                else:
                    col_act_e, col_del_e = st.columns(2)
                    with col_act_e:
                        if st.button("💾 Atualizar Empenho", type="primary", use_container_width=True, key="btn_atualizar_emp_ok"):
                            if not servidor_responsavel_e or not numero_empenho_input.strip() or valor_empenho_input <= 0:
                                st.error("❌ Preencha todos os campos obrigatórios.")
                            else:
                                num_emp_limpo = numero_empenho_input.strip().upper()
                                conn, cursor = conectar_banco()
                                try:
                                    cursor.execute("SELECT valor_total, valor_disponivel FROM empenhos WHERE id = ?", (id_emp_atual,))
                                    v_tot_antigo, v_disp_antigo = cursor.fetchone()
                                    diferenca = valor_empenho_input - v_tot_antigo
                                    novo_disponivel = v_disp_antigo + diferenca
                                    
                                    if novo_disponivel < 0:
                                        st.error("❌ Não é possível reduzir o valor porque o saldo ficaria negativo pelas diárias já pagas.")
                                    else:
                                        cursor.execute("""
                                            UPDATE empenhos 
                                            SET servidor_nome = ?, numero_empenho = ?, valor_total = ?, valor_disponivel = ?, data_empenho = ?
                                            WHERE id = ?
                                        """, (servidor_responsavel_e, num_emp_limpo, valor_empenho_input, novo_disponivel, str(data_selecionada_e), id_emp_atual))
                                        conn.commit()
                                        st.success("✔️ Empenho atualizado com sucesso!")
                                        st.rerun()
                                except sqlite3.IntegrityError:
                                    st.error("❌ Conflito: Esse número de empenho já existe em outro registro.")
                                finally:
                                    conn.close()
                                    
                    with col_del_e:
                        if st.button("❌ Excluir Empenho", type="secondary", use_container_width=True, key="btn_excluir_emp_ok"):
                            conn, cursor = conectar_banco()
                            try:
                                # REMOÇÃO DA TRAVA: Limpa primeiro os registros amarrados na tabela 'diarias'
                                cursor.execute("DELETE FROM diarias WHERE empenho_id = ?", (id_emp_atual,))
                                # Exclui a nota de empenho principal da tabela 'empenhos'
                                cursor.execute("DELETE FROM empenhos WHERE id = ?", (id_emp_atual,))
                                conn.commit()
                                st.success("✔️ Empenho e todas as suas diárias associadas foram removidos!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Erro ao remover: {e}")
                            finally:
                                conn.close()

    # =====================================================================
    # MÓDULO DEAEV
    # =====================================================================
    elif st.session_state.menu_ativo == "deaev":
        st.markdown("<h2>📋 Registro de Diárias - DEAEV</h2>", unsafe_allow_html=True)
        box_empenhos = listar_empenhos_com_saldo()
        
        if not lista_servidores:
            st.warning("⚠️ Atenção: Não há servidores cadastrados no sistema.")
        elif not box_empenhos:
            st.error("❌ Erro: Não existem empenhos com SALDO DISPONÍVEL.")
        else:
            st.markdown("### 🆕 Nova Diária")
            options_serv = {s["label"]: s["id"] for s in lista_servidores}
            servidor_selecionado = st.selectbox("Nome do Servidor", options=[""] + list(options_serv.keys()))
            
            options_emp = {e["label"]: e["id"] for e in box_empenhos}
            empenho_selecionado = st.selectbox("Empenho Destinado", options=[""] + list(options_emp.keys()))
            
            c_data, c_jornada, c_valor = st.columns([4, 4, 4])
            with c_data:
                data_diaria = st.date_input("Data do Plantão", format="DD/MM/YYYY")
            with c_jornada:
                jornada_selecionada = st.selectbox("Extra Jornada", options=["6 Horas", "12 Horas"])
            with c_valor:
                valor_calculado = 180.00 if jornada_selecionada == "6 Horas" else 360.00
                st.text_input("Valor da Diária", value=f"R$ {valor_calculado:,.2f}".replace(".", ","), disabled=True)
            
            if st.button("💾 Salvar Registro de Diária", type="primary"):
                if not servidor_selecionado or not empenho_selecionado:
                    st.error("❌ Por favor, preencha todos os campos.")
                else:
                    id_serv = options_serv[servidor_selecionado]
                    id_emp = options_emp[empenho_selecionado]
                    data_formatada_db = data_diaria.strftime("%Y-%m-%d")
                    
                    conn, cursor = conectar_banco()
                    try:
                        cursor.execute("SELECT valor_disponivel FROM empenhos WHERE id = ?", (id_emp,))
                        saldo_atual = cursor.fetchone()[0]
                        
                        if saldo_atual >= valor_calculado:
                            cursor.execute("INSERT INTO diarias (servidor_id, empenho_id, data_diaria, jornada, valor_diaria) VALUES (?, ?, ?, ?, ?)", 
                                           (id_serv, id_emp, data_formatada_db, jornada_selecionada, valor_calculado))
                            cursor.execute("UPDATE empenhos SET valor_disponivel = valor_disponivel - ? WHERE id = ?", (valor_calculado, id_emp))
                            conn.commit()
                            st.success("✔️ Diária registrada com sucesso!")
                        else:
                            st.error("❌ Saldo insuficiente neste empenho!")
                    except Exception as e:
                        conn.rollback()
                        st.error(f"Erro: {e}")
                    finally:
                        conn.close()
                        st.rerun()

    # =====================================================================
   # =====================================================================
    # MÓDULO RELATÓRIOS 📊
    # =====================================================================
    elif st.session_state.menu_ativo == "relatorios":
        st.markdown("<h2>📊 Relatórios e Prestação de Contas DEAEV</h2>", unsafe_allow_html=True)
        
        col_filtro_m, col_filtro_a = st.columns(2)
        with col_filtro_m:
            mes_relatorio = st.selectbox("Mês de Competência", options=[f"{i:02d}" for i in range(1, 13)], index=datetime.now().month - 1)
        with col_filtro_a:
            ano_relatorio = st.selectbox("Ano de Competência", options=[str(a) for a in range(2024, 2031)], index=2)
            
        competencia_busca = f"{ano_relatorio}-{mes_relatorio}"
        
        conn, cursor = conectar_banco()
        cursor.execute("""
            SELECT s.unidade_lotacao, s.nome_completo, s.cpf, e.numero_empenho, d.data_diaria, d.jornada, d.valor_diaria
            FROM diarias d
            JOIN servidores s ON d.servidor_id = s.id
            JOIN empenhos e ON d.empenho_id = e.id
            WHERE strftime('%Y-%m', d.data_diaria) = ?
            ORDER BY s.nome_completo ASC, e.numero_empenho ASC, d.data_diaria ASC
        """, (competencia_busca,))
        dados_diarias_mes = cursor.fetchall()
        conn.close()
        
        if not dados_diarias_mes:
            st.info(f"Nenhum registro de diária (DEAEV) encontrado para {mes_relatorio}/{ano_relatorio}.")
        else:
            st.success(f"Foram encontradas **{len(dados_diarias_mes)}** diárias lançadas neste mês.")
            st.markdown("---")
            
            if "relatorio_tipo_selecionado" not in st.session_state:
                st.session_state.relatorio_tipo_selecionado = None
                
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                tipo_btn_deaev = "primary" if st.session_state.relatorio_tipo_selecionado == "deaev_realizada" else "secondary"
                if st.button("📋 DEAEV REALIZADA", type=tipo_btn_deaev, use_container_width=True):
                    st.session_state.relatorio_tipo_selecionado = "deaev_realizada"
                    st.rerun()
            with col_btn2:
                tipo_btn_template = "primary" if st.session_state.relatorio_tipo_selecionado == "template_final" else "secondary"
                if st.button("🏢 TEMPLATE FINAL", type=tipo_btn_template, use_container_width=True):
                    st.session_state.relatorio_tipo_selecionado = "template_final"
                    st.rerun()
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            if st.session_state.relatorio_tipo_selecionado == "deaev_realizada":
                st.markdown("### 🔍 Opções para: **DEAEV REALIZADA**")
                formato_deaev = st.radio("Escolha o formato de saída:", options=["PDF", "PLANILHA"], index=0, horizontal=True)
                
                st.markdown("---")
                num_processo_op = st.text_input("Informe o Número do Processo / Protocolo:", placeholder="Ex: 26.216.461-6", key="proc_deaev")
                
                if num_processo_op.strip() == "":
                    st.warning("⚠️ Digite o número do processo acima para liberar a geração do relatório.")
                else:
                    titulo_tabela = f"TABELA 01 - PAGAMENTO {MESES_EXTENSO[mes_relatorio]} / {ano_relatorio}"
                    titulo_lote = f"SESP/DEPPEN – PAGAMENTO DE EXTRAJORNADA – REF. {mes_relatorio}/{ano_relatorio} – REGIONAL DE UMUARAMA – PROTOCOLO {num_processo_op.strip()}"
                    
                    if formato_deaev == "PLANILHA":
                        wb_op = Workbook()
                        ws_op = wb_op.active
                        ws_op.title = "Worksheet"
                        
                        FONTE_XLSX = "Calibri"
                        COR_CABECALHO = "FF729FCD"
                        COR_BRANCO = "FFFFFFFF"
                        borda_fina = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
                        alinhamento_centro = Alignment(horizontal="center", vertical="center")
                        
                        titulo_tabela_xlsx = f"TABELA 01 – PAGAMENTO {MESES_EXTENSO[mes_relatorio]} {ano_relatorio}"
                        
                        # Linha 1: Título
                        ws_op.append([titulo_tabela_xlsx])
                        ws_op.merge_cells("A1:I1")
                        ws_op["A1"].font = Font(name=FONTE_XLSX, size=11, bold=True)
                        ws_op["A1"].alignment = alinhamento_centro
                        ws_op.row_dimensions[1].height = 30
                        for col in range(1, 10): ws_op.cell(row=1, column=col).border = borda_fina
                        
                        # Linha 2: Subtítulo
                        ws_op.append([titulo_lote])
                        ws_op.merge_cells("A2:I2")
                        ws_op["A2"].font = Font(name=FONTE_XLSX, size=11, bold=True)
                        ws_op["A2"].alignment = alinhamento_centro
                        ws_op.row_dimensions[2].height = 30
                        for col in range(1, 10): ws_op.cell(row=2, column=col).border = borda_fina
                        
                        # Linha 3: Cabeçalho
                        headers_op = ["Unidade", "Policial", "CPF", "Empenho", "Data", "Horas", "Valor", "Total", "Liquidação"]
                        ws_op.append(headers_op)
                        ws_op.row_dimensions[3].height = 30
                        for col_num in range(1, len(headers_op) + 1):
                            cell = ws_op.cell(row=3, column=col_num)
                            cell.font = Font(name=FONTE_XLSX, size=11, bold=True)
                            cell.fill = PatternFill(start_color=COR_CABECALHO, end_color=COR_CABECALHO, fill_type="solid")
                            cell.alignment = alinhamento_centro
                            cell.border = borda_fina
                        
                        # Larguras das colunas
                        larguras_colunas = {"A": 12.0, "B": 50.559, "C": 16.282, "D": 17.567, "E": 15.139, "F": 9.283, "G": 13.997, "H": 16.282, "I": 15.139}
                        for col_letra, largura in larguras_colunas.items():
                            ws_op.column_dimensions[col_letra].width = largura
                        
                        # Monta os grupos tratando a string da unidade para extrair apenas a sigla
                        grupos_xlsx = []
                        grupo_atual_xlsx = None
                        for row_data in dados_diarias_mes:
                            unidade, policial, cpf, empenho, data_db, jornada, valor = row_data
                            
                            # Extrai apenas a sigla se houver o padrão "SIGLA - NOME DA UNIDADE"
                            sigla_unidade = unidade.split(" - ")[0].strip() if unidade else ""
                            
                            dt_br = datetime.strptime(data_db, "%Y-%m-%d").strftime("%d/%m/%Y")
                            num_horas = 6 if "6" in jornada else 12
                            
                            if grupo_atual_xlsx is not None and grupo_atual_xlsx["policial"] == policial and grupo_atual_xlsx["empenho"] == empenho:
                                grupo_atual_xlsx["linhas"].append((dt_br, num_horas, valor))
                            else:
                                grupo_atual_xlsx = {"unidade": sigla_unidade, "policial": policial, "cpf": cpf, "empenho": empenho, "linhas": [(dt_br, num_horas, valor)]}
                                grupos_xlsx.append(grupo_atual_xlsx)
                        
                        # Escreve os dados e trata bordas de células mescladas
                        linha_atual_xlsx = 4
                        for grupo in grupos_xlsx:
                            qtd_linhas_grupo = len(grupo["linhas"])
                            total_grupo = sum(item[2] for item in grupo["linhas"])
                            linha_inicio_grupo = linha_atual_xlsx
                            linha_fim_grupo = linha_atual_xlsx + qtd_linhas_grupo - 1
                            
                            for idx, (dt_br, num_horas, valor) in enumerate(grupo["linhas"]):
                                linha = linha_atual_xlsx + idx
                                ws_op.cell(row=linha, column=1, value=grupo["unidade"] if idx == 0 else None)
                                ws_op.cell(row=linha, column=2, value=grupo["policial"] if idx == 0 else None)
                                ws_op.cell(row=linha, column=3, value=str(grupo["cpf"]) if idx == 0 else None)
                                ws_op.cell(row=linha, column=4, value=grupo["empenho"] if idx == 0 else None)
                                ws_op.cell(row=linha, column=5, value=dt_br)
                                ws_op.cell(row=linha, column=6, value=num_horas)
                                ws_op.cell(row=linha, column=7, value=formatar_valor_br_moeda(valor))
                                ws_op.cell(row=linha, column=8, value=formatar_valor_br_moeda(total_grupo) if idx == 0 else None)
                                
                                for col_num in range(1, len(headers_op) + 1):
                                    cell = ws_op.cell(row=linha, column=col_num)
                                    cell.font = Font(name=FONTE_XLSX, size=11, bold=False)
                                    cell.fill = PatternFill(start_color=COR_BRANCO, end_color=COR_BRANCO, fill_type="solid")
                                    cell.alignment = alinhamento_centro
                                    cell.border = borda_fina
                            
                            if qtd_linhas_grupo > 1:
                                for col_num_mesclar in [1, 2, 3, 4, 8, 9]:
                                    letra_col = get_column_letter(col_num_mesclar)
                                    ws_op.merge_cells(f"{letra_col}{linha_inicio_grupo}:{letra_col}{linha_fim_grupo}")
                            
                            linha_atual_xlsx = linha_fim_grupo + 1
                        
                        # Linhas de totalizadores
                        total_policiais = len(set(item[2] for item in dados_diarias_mes))
                        total_deaev = len(dados_diarias_mes)
                        total_geral = sum(item[6] for item in dados_diarias_mes)
                        
                        textos_totais = [
                            f"Total de Policiais: {total_policiais}",
                            f"Total de DEAEV: {total_deaev}",
                            f"TOTAL de R$ {formatar_valor_br(total_geral)}"
                        ]
                        for texto_total in textos_totais:
                            ws_op.cell(row=linha_atual_xlsx, column=1, value=texto_total)
                            ws_op.merge_cells(f"A{linha_atual_xlsx}:I{linha_atual_xlsx}")
                            
                            # Garante que todas as células da mesclagem dos totais tenham bordas aplicadas
                            for col in range(1, 10):
                                cell_borda = ws_op.cell(row=linha_atual_xlsx, column=col)
                                cell_borda.border = borda_fina
                                cell_borda.fill = PatternFill(start_color=COR_CABECALHO, end_color=COR_CABECALHO, fill_type="solid")
                            
                            cell = ws_op.cell(row=linha_atual_xlsx, column=1)
                            cell.font = Font(name=FONTE_XLSX, size=11, bold=True)
                            cell.alignment = alinhamento_centro
                            ws_op.row_dimensions[linha_atual_xlsx].height = 20
                            linha_atual_xlsx += 1
                                    
                        caminho_temp_xlsx_op = os.path.join(DIRETORIO_ATUAL, "temp_op.xlsx")
                        wb_op.save(caminho_temp_xlsx_op)
                        with open(caminho_temp_xlsx_op, "rb") as f_op:
                            st.download_button("🟢 Baixar DEAEV REALIZADA em Excel (.xlsx)", data=f_op.read(), file_name=f"relatorio_mensal_{mes_relatorio}_{ano_relatorio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                            
                    elif formato_deaev == "PDF":
                        caminho_temp_pdf_op = os.path.join(DIRETORIO_ATUAL, "temp_op.pdf")
                        doc_op = SimpleDocTemplate(caminho_temp_pdf_op, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
                        story_op = []
                        
                        estilos = getSampleStyleSheet()
                        estilo_titulo = ParagraphStyle('TitOp', parent=estilos['Heading3'], alignment=1, spaceAfter=4)
                        estilo_subtitulo = ParagraphStyle('SubTitOp', parent=estilos['Normal'], alignment=1, fontSize=9, spaceAfter=15)
                        story_op.append(Paragraph(titulo_tabela, estilo_titulo))
                        story_op.append(Paragraph(titulo_lote, estilo_subtitulo))
                        
                        # Removida qualquer menção à Liquidação. Foco nas 7 colunas essenciais.
                        headers_op = ["Policial Penal", "CPF", "Empenhos", "Data", "Horas", "Valor", "Total"]
                        tabela_dados_pdf = [headers_op]
                        
                        # Estilo base da Tabela
                        estilos_tabela = [
                            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0,0), (-1,-1), 9),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ]
                        
                        grupos_pdf = []
                        grupo_atual_pdf = None
                        
                        # Agrupa os dados para saber exatamente quem deve ser mesclado
                        for row_data in dados_diarias_mes:
                            unidade, policial, cpf, empenho, data_db, jornada, valor = row_data
                            dt_br = datetime.strptime(data_db, "%Y-%m-%d").strftime("%d/%m/%Y")
                            num_horas = 6 if "6" in jornada else 12
                            
                            if grupo_atual_pdf is not None and grupo_atual_pdf["policial"] == policial and grupo_atual_pdf["empenho"] == empenho:
                                grupo_atual_pdf["linhas"].append((dt_br, num_horas, valor))
                            else:
                                grupo_atual_pdf = {"policial": policial, "cpf": cpf, "empenho": empenho, "linhas": [(dt_br, num_horas, valor)]}
                                grupos_pdf.append(grupo_atual_pdf)
                        
                        # Monta as linhas da tabela e gera os comandos de mesclagem (SPAN)
                        linha_atual_pdf = 1 # Linha 0 é o cabeçalho
                        for g in grupos_pdf:
                            qtd_linhas = len(g["linhas"])
                            total_grupo = sum(item[2] for item in g["linhas"])
                            
                            # Se houver mais de uma linha para o mesmo policial/empenho, aplica o SPAN nas colunas fixas
                            if qtd_linhas > 1:
                                fim_mesclagem = linha_atual_pdf + qtd_linhas - 1
                                # Mescla Policial (Col 0), CPF (Col 1), Empenho (Col 2) e Total (Col 6)
                                estilos_tabela.append(('SPAN', (0, linha_atual_pdf), (0, fim_mesclagem)))
                                estilos_tabela.append(('SPAN', (1, linha_atual_pdf), (1, fim_mesclagem)))
                                estilos_tabela.append(('SPAN', (2, linha_atual_pdf), (2, fim_mesclagem)))
                                estilos_tabela.append(('SPAN', (6, linha_atual_pdf), (6, fim_mesclagem)))
                            
                            for idx, (dt_br, num_horas, valor) in enumerate(g["linhas"]):
                                # Para células mescladas, mandamos os dados completos em todas as linhas. 
                                # O ReportLab se encarrega de exibir o valor centralizado usando a primeira célula do bloco.
                                tabela_dados_pdf.append([
                                    g["policial"][:30],
                                    g["cpf"],
                                    g["empenho"],
                                    dt_br,
                                    str(num_horas),
                                    formatar_valor_br(valor),
                                    formatar_valor_br(total_grupo)
                                ])
                            
                            linha_atual_pdf += qtd_linhas
                                    
                        # Definição perfeita de larguras para preencher a página paisagem de forma limpa
                        t_op = Table(tabela_dados_pdf, colWidths=[220, 95, 95, 75, 50, 70, 70], repeatRows=1)
                        t_op.setStyle(TableStyle(estilos_tabela))
                        story_op.append(t_op)
                        
                        # Rodapé com totalizadores
                        total_policiais = len(set(item[2] for item in dados_diarias_mes))
                        total_deaev = len(dados_diarias_mes)
                        total_geral = sum(item[6] for item in dados_diarias_mes)
                        
                        story_op.append(Spacer(1, 15))
                        tabela_totais_pdf = [
                            ["Total de Policiais", "Total de DEAEV", "Total em R$"],
                            [str(total_policiais), str(total_deaev), formatar_valor_br(total_geral)]
                        ]
                        t_op_totais = Table(tabela_totais_pdf, colWidths=[220, 220, 220])
                        t_op_totais.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                            ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
                            ('FONTSIZE', (0,0), (-1,-1), 10),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ]))
                        story_op.append(t_op_totais)
                        
                        doc_op.build(story_op)
                        
                        with open(caminho_temp_pdf_op, "rb") as f_pdf_op:
                            st.download_button("🔴 Baixar DEAEV REALIZADA em PDF (.pdf)", data=f_pdf_op.read(), file_name=f"relatorio_mensal_{mes_relatorio}_{ano_relatorio}.pdf", mime="application/pdf", use_container_width=True)
            
            elif st.session_state.relatorio_tipo_selecionado == "template_final":
                st.markdown("### 🔍 Opções para: **TEMPLATE FINAL**")
                formato_template = st.radio("Escolha o formato de saída:", options=["PDF", "PLANILHA"], index=0, horizontal=True)
                
                st.markdown("---")
                num_processo_tp = st.text_input("Informe o Número do Processo / Protocolo:", placeholder="Ex: 26.216.461-6", key="proc_template")
                
                if num_processo_tp.strip() == "":
                    st.warning("⚠️ Digite o número do processo acima para liberar a geração do template.")
                else:
                    df_bruto = pd.DataFrame(dados_diarias_mes, columns=["Unidade", "Policial", "CPF", "Empenho", "Data", "Jornada", "Valor"])
                    df_consolidado = df_bruto.groupby(["Empenho", "CPF", "Policial"]).agg({"Valor": "sum"}).reset_index()
                    
                    if mes_relatorio in ["01", "03", "05", "07", "08", "10", "12"]: dia_fim = "31"
                    elif mes_relatorio in ["04", "06", "09", "11"]: dia_fim = "30"
                    else: dia_fim = "28"
                    data_emissao_nl = f"{dia_fim}/{mes_relatorio}/{ano_relatorio}"
                    obs_padrao = f"SESP/DEPPEN - PAGAMENTO DE EXTRAJORNADA – REF. {mes_relatorio}/{ano_relatorio} - REGIONAL DE UMUARAMA"
                    
                    if formato_template == "PLANILHA":
                        wb_tp = Workbook()
                        ws_tp = wb_tp.active
                        ws_tp.title = "Worksheet"
                        
                        headers_tp = [
                            'Data de Emissão NL', 'UG Emitente', 'Nota de Empenho', 'Tipo Patrimonial', 'Item Patrimonial', 
                            'Operação Patrimonial', 'Valor do Item', 'Tipo de Retenção', 'Credor da Retenção', 
                            'Valor da Base de Cálculo da Retenção', 'Valor da Retenção', 'Observação', 'Mês de Competência', 
                            'Data do Processo', 'Código do Processo', 'Credor Secundário', 'DEA', 'Tipo de Documento Comprobatório', 
                            'Número Documento Comprobatório', 'Processo Documento Comprobatório', 'Data Documento Comprobatório', 
                            'Competência Documento Comprobatório', 'Tipo de Série Documento Comprobatório', 
                            'Descrição Tipo de Série Documento Comprobatório', 'Chave de Acesso Documento Comprobatório', 
                            'Código de Verificação Documento Comprobatório', 'Valor Documento Comprobatório', 'Categoria do PADV'
                        ]
                        ws_tp.append(headers_tp)
                        
                        for idx, row_c in df_consolidado.iterrows():
                            ws_tp.append([
                                data_emissao_nl, 390000, row_c["Empenho"], 15, 1902, 58, row_c["Valor"], "", "", "", "",
                                obs_padrao, int(mes_relatorio), "", num_processo_tp.strip(), "", 0, 23, f"{mes_relatorio}/{ano_relatorio}",
                                num_processo_tp.strip(), data_emissao_nl, f"{mes_relatorio}/{ano_relatorio}", 6, "", "", "", row_c["Valor"], ""
                            ])
                            
                        caminho_temp_xlsx_tp = os.path.join(DIRETORIO_ATUAL, "temp_tp.xlsx")
                        wb_tp.save(caminho_temp_xlsx_tp)
                        with open(caminho_temp_xlsx_tp, "rb") as f_tp:
                            st.download_button("🟢 Baixar TEMPLATE FINAL em Excel (.xlsx)", data=f_tp.read(), file_name=f"TEMPLATE_{mes_relatorio}_{ano_relatorio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                            
                    elif formato_template == "PDF":
                        caminho_temp_pdf_tp = os.path.join(DIRETORIO_ATUAL, "temp_tp.pdf")
                        doc_tp = SimpleDocTemplate(caminho_temp_pdf_tp, pagesize=landscape(letter), rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20)
                        story_tp = []
                        
                        estilos = getSampleStyleSheet()
                        estilo_titulo = ParagraphStyle('TitTp', parent=estilos['Heading3'], alignment=1, spaceAfter=15)
                        story_tp.append(Paragraph(f"<b>TEMPLATE DE PRESTAÇÃO DE CONTAS CONSOLIDADA - REF {mes_relatorio}/{ano_relatorio}</b>", estilo_titulo))
                        story_tp.append(Paragraph(f"Protocolo Geral do Lote: {num_processo_tp.strip()}<br/><br/>", estilos['Normal']))
                        
                        headers_pdf_tp = ["Empenho", "Policial / Credor", "CPF", "Valor Consolidado", "Competência"]
                        tabela_tp_pdf = [headers_pdf_tp]
                        
                        for idx, row_c in df_consolidado.iterrows():
                            tabela_tp_pdf.append([row_c["Empenho"], row_c["Policial"][:30], row_c["CPF"], f"R$ {row_c['Valor']:.2f}", f"{mes_relatorio}/{ano_relatorio}"])
                            
                        t_tp = Table(tabela_tp_pdf, colWidths=[120, 220, 100, 110, 90])
                        t_tp.setStyle(TableStyle([
                            ('BACKGROUND', (0,0), (-1,0), colors.navy),
                            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('FONTSIZE', (0,0), (-1,-1), 9),
                            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                        ]))
                        story_tp.append(t_tp)
                        doc_tp.build(story_tp)
                        
                        with open(caminho_temp_pdf_tp, "rb") as f_pdf_tp:
                            st.download_button("🔴 Baixar TEMPLATE FINAL em PDF (.pdf)", data=f_pdf_tp.read(), file_name=f"TEMPLATE_{mes_relatorio}_{ano_relatorio}.pdf", mime="application/pdf", use_container_width=True)
    # =====================================================================
    # MÓDULO SISTEMA ⚙️
    # =====================================================================
    elif st.session_state.menu_ativo == "sistema":
        st.markdown("<h2>⚙️ Parâmetros do Sistema</h2>", unsafe_allow_html=True)
        tab_unidades, tab_bancos = st.tabs(["🏛️ Unidades de Lotação", "🏦 Bancos Cadastrados"])
        
        with tab_unidades:
            st.markdown("#### Gerenciar Unidades Operacionais / Administrativas")
            with st.form("form_nova_unidade", clear_on_submit=True):
                nova_unidade_txt = st.text_input("Nome da Nova Unidade")
                if st.form_submit_button("➕ Adicionar Unidade"):
                    if not nova_unidade_txt.strip(): st.error("Digite um nome válido.")
                    else:
                        conn, cursor = conectar_banco()
                        try:
                            cursor.execute("INSERT INTO unidades (nome_unidade) VALUES (?)", (nova_unidade_txt.strip().upper(),))
                            conn.commit()
                            st.success("Unidade adicionada com sucesso!")
                        except sqlite3.IntegrityError: st.error("Esta unidade já está cadastrada.")
                        finally: conn.close(); st.rerun()
            
            conn, cursor = conectar_banco()
            cursor.execute("SELECT id, nome_unidade FROM unidades ORDER BY nome_unidade ASC")
            unidades_banco = cursor.fetchall()
            conn.close()
            if unidades_banco:
                st.dataframe(pd.DataFrame(unidades_banco, columns=["ID", "Nome da Unidade"]), use_container_width=True, hide_index=True)
                
        with tab_bancos:
            st.markdown("#### Gerenciar Instituições Bancárias")
            with st.form("form_novo_banco", clear_on_submit=True):
                col_c, col_n = st.columns([3, 9])
                with col_c: cod_banco_txt = st.text_input("Código Compe (3 dígitos)", max_chars=3)
                with col_n: nome_banco_txt = st.text_input("Nome Comercial do Banco")
                if st.form_submit_button("➕ Adicionar Banco"):
                    if not cod_banco_txt.strip() or not nome_banco_txt.strip(): st.error("Ambos os campos são obrigatórios.")
                    else:
                        conn, cursor = conectar_banco()
                        try:
                            cursor.execute("INSERT INTO bancos (codigo, nome_banco) VALUES (?, ?)", (cod_banco_txt.strip(), nome_banco_txt.strip().upper()))
                            conn.commit()
                            st.success("Banco cadastrado!")
                        except sqlite3.IntegrityError: st.error("Este código de banco já existe.")
                        finally: conn.close(); st.rerun()
            
            conn, cursor = conectar_banco()
            cursor.execute("SELECT codigo, nome_banco FROM bancos ORDER BY codigo ASC")
            bancos_banco = cursor.fetchall()
            conn.close()
            if bancos_banco:
                st.dataframe(pd.DataFrame(bancos_banco, columns=["Código Compe", "Nome do Banco"]), use_container_width=True, hide_index=True)
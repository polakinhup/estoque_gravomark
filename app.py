import streamlit as st
import gspread
import pandas as pd
import os
import io
import re
import time
from datetime import datetime
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

import streamlit as st

# 1. Usando a imagem desejada como ícone da página
NOME_DO_ICONE = "LOGO PNG COM FUNDO.png"  # Coloque o nome exato do arquivo aqui

import streamlit as st

# 1. Configuração da página
st.set_page_config(
    page_title="Estoque GravoMark 02.08", 
    page_icon="LOGO PNG COM FUNDO.png", 
    layout="wide"
)

# 2. CSS para remover o ícone do GitHub, menu do topo, rodapé e botões flutuantes
st.markdown(
    """
    <style>
    /* 1. Remove completamente o cabeçalho e o menu superior */
    header { display: none !important; }
    [data-testid="stHeader"] { display: none !important; }
    
    /* 2. Remove a barra branca inferior do modo Embed (Built with Streamlit / Fullscreen) */
    footer { display: none !important; }
    [data-testid="stFooter"] { display: none !important; }
    [data-testid="stBottom"] { display: none !important; }
    
    /* 3. Prevenção extra: oculta qualquer link de fullscreen caso o Streamlit mude a classe */
    a[title="View fullscreen"] { display: none !important; }
    svg[title="Fullscreen"] { display: none !important; }
    
    /* 4. Ajusta o espaçamento do aplicativo para aproveitar a tela toda */
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    [data-testid="stSidebar"] { padding-top: 1rem !important; }
    [data-testid="stSidebarHeader"] { display: none !important; }
    </style>
    """,
    unsafe_allow_html=True
)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def conectar_google():
    creds = None
    if "google_token" in st.secrets:
        token_info = dict(st.secrets["google_token"])
        creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    elif os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    if not creds:
        st.error("Erro nas credenciais de acesso ao Google. Verifique o registro em Secrets.")
        st.stop()

    client_sheets = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    return client_sheets, drive_service

client_sheets, drive_service = conectar_google()

# --- COLE OS SEUS IDs REAIS AQUI ---
SPREADSHEET_ID = "1mnR2hraUpJm5KIQRLk4JOCU35zTKgPaelr02CjxJ248"
FOLDER_ENTRADA_ID = "14b0Dp4LEEftPMIUkVxDd_0JFKGJhRbWL"
FOLDER_SAIDA_ID = "1iFmbto3DIRKW83SdON-QaMXTDrfrRmcx"

@st.cache_data(ttl=30)
def ler_dados_planilha(nome_aba, usar_formula=False):
    aba = client_sheets.open_by_key(SPREADSHEET_ID).worksheet(nome_aba)
    if usar_formula:
        return aba.get_all_records(value_render_option='FORMULA')
    return aba.get_all_records()

def formatar_codigo_peca(codigo_bruto):
    apenas_numeros = "".join(filter(str.isdigit, str(codigo_bruto)))
    if len(apenas_numeros) == 14:
        return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:9]}.{apenas_numeros[9:]}"
    return str(codigo_bruto).strip()

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

try:
    col_vazia1, col_logo, col_vazia2 = st.sidebar.columns([1, 2, 1])
    with col_logo:
        st.image("LOGO PNG.png", use_container_width=True)
    st.sidebar.markdown("<br>", unsafe_allow_html=True)
except Exception:
    pass

st.sidebar.title("Navegação")
aba = st.sidebar.radio("Ir para:", ["📦 Lançar Movimentação", "📊 Consultar Estoque", "📋 Histórico", "🛠️ Área Técnica"])

if aba == "📦 Lançar Movimentação":
    
    col_titulo, col_senha = st.columns([10, 1])
    with col_titulo:
        st.header("Lançamento Manual em Lote (Estilo Planilha)")
    with col_senha:
        if not st.session_state["autenticado"]:
            with st.popover("🔑"):
                with st.form("form_login"):
                    senha = st.text_input("Senha:", type="password")
                    submit = st.form_submit_button("Desbloquear", use_container_width=True)
                    if submit:
                        if senha == "1234":
                            st.session_state["autenticado"] = True
                            st.rerun()
                        else:
                            st.error("Senha incorreta!")
        else:
            if st.button("🔒 Sair"):
                st.session_state["autenticado"] = False
                st.rerun()

    if not st.session_state["autenticado"]:
        st.info("🔒 O sistema está bloqueado. Clique no ícone de chave (🔑) no canto superior direito para acessar.")
        st.stop() 

    dados_cadastro = ler_dados_planilha("Cadastro_Pecas")
    dados_estoque_atual = ler_dados_planilha("Estoque_Atual")

    lista_estoque_geral = []
    dict_cadastrados = {}
    if dados_cadastro:
        for row in dados_cadastro:
            cod = str(row.get('Codigo_Peca', '')).strip().replace("'", "")
            desc = str(row.get('Descricao', '')).strip()
            if cod:
                item_label = f"{cod} | {desc}"
                lista_estoque_geral.append(item_label)
                dict_cadastrados[cod] = desc
    lista_estoque_geral = sorted(list(set(lista_estoque_geral)))

    st.subheader("1. Cabeçalho da Nota Fiscal / Documento")
    col_tipo, col_nf_head, col_pdf_head = st.columns([2, 3, 4])
    
    with col_tipo:
        tipo_mov = st.radio("Operação:", ["Entrada", "Saida"], horizontal=True)
    with col_nf_head:
        num_nf_input = st.text_input("Número da NF ou Documento*", placeholder="Obrigatório")
        sem_nf_check = st.checkbox("🔓 Autorizar lançamento sem NF")
        senha_autorizacao = ""
        if sem_nf_check:
            senha_autorizacao = st.text_input("Senha de Autorização para Sem NF:", type="password")

    with col_pdf_head:
        arquivo_nf = st.file_uploader("Anexar PDF (Será vinculado a TODOS os itens)", type=["pdf"])

    st.divider()

    st.subheader(f"2. Itens para {tipo_mov} em Massa")

    if tipo_mov == "Entrada":
        st.caption("💡 Adicione quantos itens quiser na tabela abaixo antes de salvar.")
        
        df_template = pd.DataFrame([
            {"Peça Cadastrada": "", "Código Peça Nova (Se não existir)": "", "Descrição Peça Nova": "", "Quantidade": 1}
        ])

        editor_itens = st.data_editor(
            df_template,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Peça Cadastrada": st.column_config.SelectboxColumn(
                    "Selecione Peça Existente",
                    options=[""] + lista_estoque_geral,
                    width="large"
                ),
                "Código Peça Nova (Se não existir)": st.column_config.TextColumn("Ou digite Código Novo", width="medium"),
                "Descrição Peça Nova": st.column_config.TextColumn("Descrição da Peça Nova", width="large"),
                "Quantidade": st.column_config.NumberColumn("Quantidade", min_value=1, step=1, default=1, width="small")
            }
        )

        if st.button("💾 Registrar TODOS os Itens no Estoque", use_container_width=True, type="primary"):
            if not num_nf_input.strip() and not sem_nf_check:
                st.error("🛑 **ERRO DE BLOQUEIO:** O número da Nota Fiscal é OBRIGATÓRIO! Para lançar sem NF, marque a opção de autorização e digite a senha.")
                st.stop()
            
            if sem_nf_check and senha_autorizacao != "1234":
                st.error("🛑 **SENHA INCORRETA:** Senha de autorização inválida para lançamento sem Nota Fiscal.")
                st.stop()

            itens_validos = []
            novas_pecas_para_cadastrar = []

            for idx, row in editor_itens.iterrows():
                peca_existente = str(row.get("Peça Cadastrada", "")).strip()
                cod_novo = str(row.get("Código Peça Nova (Se não existir)", "")).strip()
                desc_nova = str(row.get("Descrição Peça Nova", "")).strip()
                qtd_item = int(row.get("Quantidade", 1))

                if peca_existente:
                    partes = peca_existente.split(" | ", 1)
                    c_fin = partes[0]
                    d_fin = partes[1] if len(partes) > 1 else ""
                    itens_validos.append({"codigo": c_fin, "descricao": d_fin, "qtd": qtd_item})
                elif cod_novo and desc_nova:
                    c_fin = formatar_codigo_peca(cod_novo)
                    itens_validos.append({"codigo": c_fin, "descricao": desc_nova, "qtd": qtd_item})
                    if c_fin not in dict_cadastrados:
                        novas_pecas_para_cadastrar.append([f"'{c_fin}", desc_nova])

            if not itens_validos:
                st.error("Preencha pelo menos um item válido na tabela acima.")
            else:
                with st.spinner("Processando lote de entrada..."):
                    data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    data_hora_segura = f"'{data_hora_bruta}"
                    
                    nf_formatada = "".join(filter(str.isdigit, num_nf_input)).lstrip('0') if num_nf_input else ""
                    nf_formatada = nf_formatada if nf_formatada else "S/N"

                    sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                    sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                    sheet_cad = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Cadastro_Pecas")

                    if novas_pecas_para_cadastrar:
                        sheet_cad.insert_rows(novas_pecas_para_cadastrar, row=2, value_input_option='USER_ENTERED')

                    link_arquivo = ""
                    if arquivo_nf is not None:
                        pasta_id = FOLDER_ENTRADA_ID
                        nome_padronizado = f"NF {nf_formatada}.pdf" if nf_formatada != "S/N" else arquivo_nf.name

                        query = f"'{pasta_id}' in parents and name='{nome_padronizado}' and trashed=false"
                        resultados = drive_service.files().list(q=query, fields="files(id)").execute().get('files', [])

                        if resultados:
                            arquivo_id = resultados[0]['id']
                        else:
                            file_metadata = {'name': nome_padronizado, 'parents': [pasta_id]}
                            media = MediaIoBaseUpload(io.BytesIO(arquivo_nf.getvalue()), mimetype=arquivo_nf.type)
                            arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                            arquivo_id = arquivo_salvo.get('id')

                        url_bruta = f"https://drive.google.com/file/d/{arquivo_id}/view"
                        link_arquivo = f'=HYPERLINK("{url_bruta}"; "📄 Abrir NF")'

                    novas_linhas_mov = []
                    novas_linhas_est = []
                    codigos_processados = set()

                    dados_est_cru = sheet_est.get_all_records()
                    linhas_a_deletar = []

                    for item in itens_validos:
                        c_seg = f"'{item['codigo']}"
                        novas_linhas_mov.append([data_hora_segura, "Entrada", nf_formatada, c_seg, item["descricao"], item["qtd"], link_arquivo, ""])

                        for idx_e, row_e in enumerate(dados_est_cru):
                            if str(row_e.get('Codigo_Peca', '')).replace("'", "").strip() == item['codigo'].strip() and str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0') == nf_formatada.strip():
                                linhas_a_deletar.append(idx_e + 2)

                        if item['codigo'] not in codigos_processados:
                            formula_saldo = f'=SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item["codigo"]}"; Movimentacoes!C:C; "{nf_formatada}"; Movimentacoes!B:B; "Entrada") + SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item["codigo"]}"; Movimentacoes!C:C; "{nf_formatada}"; Movimentacoes!B:B; "Retorno Técnica") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item["codigo"]}"; Movimentacoes!C:C; "{nf_formatada}"; Movimentacoes!B:B; "Saida") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item["codigo"]}"; Movimentacoes!C:C; "{nf_formatada}"; Movimentacoes!B:B; "Área Técnica")'
                            novas_linhas_est.append([c_seg, item["descricao"], formula_saldo, nf_formatada, 0, data_hora_segura])
                            codigos_processados.add(item['codigo'])

                    for r_del in sorted(list(set(linhas_a_deletar)), reverse=True):
                        sheet_est.delete_rows(r_del)

                    if novas_linhas_mov:
                        sheet_mov.insert_rows(novas_linhas_mov, row=2, value_input_option='USER_ENTERED')
                    if novas_linhas_est:
                        sheet_est.insert_rows(novas_linhas_est, row=2, value_input_option='USER_ENTERED')

                    st.cache_data.clear()
                    st.success(f"✅ Lote registrado com sucesso! {len(itens_validos)} itens adicionados ao estoque.")

    else:
        opcoes_saida_disponiveis = []
        for r in dados_estoque_atual:
            c_item = str(r.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
            q_item = pd.to_numeric(r.get('Quantidade_Atual', 0), errors='coerce') or 0
            
            raw_nf = str(r.get('Nota_Fiscal', '')).strip()
            nf_label = f"NF {raw_nf}" if raw_nf and raw_nf.upper() != "S/N" else "Sem NF (S/N)"
            
            desc_item = str(r.get('Descricao', '')).strip()
            if c_item and q_item > 0:
                opcoes_saida_disponiveis.append(f"{c_item} | {nf_label} | Saldo: {int(q_item)} un | {desc_item}")

        if not opcoes_saida_disponiveis:
            st.warning("⚠️ Não há peças com saldo disponível para dar saída no momento.")
        else:
            df_saida_template = pd.DataFrame([
                {"Selecione a Peça e NF de Origem": "", "Quantidade de Saída": 1}
            ])

            editor_saida = st.data_editor(
                df_saida_template,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Selecione a Peça e NF de Origem": st.column_config.SelectboxColumn(
                        "Selecione Peça e Lote de Saída",
                        options=[""] + sorted(opcoes_saida_disponiveis),
                        width="large"
                    ),
                    "Quantidade de Saída": st.column_config.NumberColumn("Quantidade", min_value=1, step=1, default=1, width="small")
                }
            )

            if st.button("📤 Confirmar Lote de Saída do Estoque", use_container_width=True, type="primary"):
                if not num_nf_input.strip() and not sem_nf_check:
                    st.error("🛑 **ERRO DE BLOQUEIO:** O número do Documento/NF de Saída é OBRIGATÓRIO! Para lançar sem NF, marque a opção de autorização e digite a senha.")
                    st.stop()

                if sem_nf_check and senha_autorizacao != "1234":
                    st.error("🛑 **SENHA INCORRETA:** Senha de autorização inválida para lançamento sem Nota Fiscal.")
                    st.stop()

                itens_saida_validos = []
                for idx, row in editor_saida.iterrows():
                    opcao_sel = str(row.get("Selecione a Peça e NF de Origem", "")).strip()
                    qtd_s = int(row.get("Quantidade de Saída", 1))

                    if opcao_sel:
                        partes_s = opcao_sel.split(" | ")
                        cod_s = partes_s[0].strip()
                        
                        str_nf_parte = partes_s[1].strip()
                        if "NF " in str_nf_parte:
                            nf_s = str_nf_parte.replace("NF ", "").strip()
                        else:
                            nf_s = "S/N"

                        desc_s = partes_s[3].strip() if len(partes_s) > 3 else ""
                        itens_saida_validos.append({"codigo": cod_s, "nf_origem": nf_s, "descricao": desc_s, "qtd": qtd_s})

                if not itens_saida_validos:
                    st.error("Selecione pelo menos uma peça para dar saída.")
                else:
                    with st.spinner("Processando lote de saída..."):
                        data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        data_hora_segura = f"'{data_hora_bruta}"

                        sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                        sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")

                        link_saida = ""
                        if arquivo_nf is not None:
                            nome_padronizado = f"SAIDA_{num_nf_input if num_nf_input else 'DOC'}.pdf"
                            file_metadata = {'name': nome_padronizado, 'parents': [FOLDER_SAIDA_ID]}
                            media = MediaIoBaseUpload(io.BytesIO(arquivo_nf.getvalue()), mimetype=arquivo_nf.type)
                            arquivo_salvo = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                            url_bruta = f"https://drive.google.com/file/d/{arquivo_salvo.get('id')}/view"
                            link_saida = f'=HYPERLINK("{url_bruta}"; "📄 Ver Comprovante")'

                        novas_linhas_mov_s = []
                        linhas_est_deletar = []
                        novas_linhas_est_s = []

                        dados_est_cru = sheet_est.get_all_records()

                        for item_s in itens_saida_validos:
                            c_seg = f"'{item_s['codigo']}"
                            novas_linhas_mov_s.append([data_hora_segura, "Saida", item_s["nf_origem"], c_seg, item_s["descricao"], item_s["qtd"], "", link_saida])

                            for idx_e, row_e in enumerate(dados_est_cru):
                                cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                if cod_e == item_s['codigo'] and nf_e == item_s['nf_origem'].lstrip('0'):
                                    linhas_est_deletar.append(idx_e + 2)

                            formula_saldo = f'=SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_s["codigo"]}"; Movimentacoes!C:C; "{item_s["nf_origem"]}"; Movimentacoes!B:B; "Entrada") + SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_s["codigo"]}"; Movimentacoes!C:C; "{item_s["nf_origem"]}"; Movimentacoes!B:B; "Retorno Técnica") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_s["codigo"]}"; Movimentacoes!C:C; "{item_s["nf_origem"]}"; Movimentacoes!B:B; "Saida") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_s["codigo"]}"; Movimentacoes!C:C; "{item_s["nf_origem"]}"; Movimentacoes!B:B; "Área Técnica")'
                            novas_linhas_est_s.append([c_seg, item_s["descricao"], formula_saldo, item_s["nf_origem"], 0, data_hora_segura])

                        for r_del in sorted(list(set(linhas_est_deletar)), reverse=True):
                            sheet_est.delete_rows(r_del)

                        if novas_linhas_mov_s:
                            sheet_mov.insert_rows(novas_linhas_mov_s, row=2, value_input_option='USER_ENTERED')
                        if novas_linhas_est_s:
                            sheet_est.insert_rows(novas_linhas_est_s, row=2, value_input_option='USER_ENTERED')

                        st.cache_data.clear()
                        st.success(f"✅ Saída em lote concluída para {len(itens_saida_validos)} itens!")

elif aba == "📊 Consultar Estoque":
    st.header("Estoque Atual")

    col_pesq, col_btn = st.columns([5, 1])
    with col_btn:
        if st.button("🔄 Atualizar Dados", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

    dados_estoque = ler_dados_planilha("Estoque_Atual")
    dados_tec = ler_dados_planilha("Area_Tecnica")

    dict_na_tecnica = {}
    if dados_tec:
        for rt in dados_tec:
            ct = str(rt.get('Codigo_Peca', '')).replace("'", "").strip()
            nft = str(rt.get('Nota_Fiscal', '')).strip().lstrip('0')
            qt = pd.to_numeric(rt.get('Quantidade', 0), errors='coerce') or 0
            chave = f"{ct}_{nft}"
            dict_na_tecnica[chave] = dict_na_tecnica.get(chave, 0) + qt

    if dados_estoque:
        df = pd.DataFrame(dados_estoque)
        df['Quantidade_Atual'] = pd.to_numeric(df['Quantidade_Atual'], errors='coerce').fillna(0)
        
        if 'Estoque_Minimo' in df.columns:
            df['Estoque_Minimo_Val'] = pd.to_numeric(df['Estoque_Minimo'], errors='coerce')
        else:
            df['Estoque_Minimo_Val'] = None
        
        df = df[df['Quantidade_Atual'] > 0].copy()

        if not df.empty:
            contagem_codigos = df['Codigo_Peca'].value_counts()
            codigos_com_multiplas_nfs = contagem_codigos[contagem_codigos > 1].index.tolist()

            def formatar_codigo_e_icone(row):
                cod_str = str(row['Codigo_Peca']).replace("'", "").strip()
                nf_str = str(row['Nota_Fiscal']).strip().lstrip('0')
                chave = f"{cod_str}_{nf_str}"
                
                tem_na_tec = dict_na_tecnica.get(chave, 0) > 0
                tem_mult_nfs = cod_str in codigos_com_multiplas_nfs

                prefixo = ""
                if tem_na_tec:
                    prefixo += "🛠️ "
                if tem_mult_nfs:
                    prefixo += "* "

                return f"{prefixo}{cod_str}"

            df['Codigo_Formatado'] = df.apply(formatar_codigo_e_icone, axis=1)

            def calc_qtd_tecnica(row):
                cod_str = str(row['Codigo_Peca']).replace("'", "").strip()
                nf_str = str(row['Nota_Fiscal']).strip().lstrip('0')
                return int(dict_na_tecnica.get(f"{cod_str}_{nf_str}", 0))

            df['Qtd_Na_Tecnica'] = df.apply(calc_qtd_tecnica, axis=1)

            with col_pesq:
                busca = st.text_input("🔍 Pesquisar por peça, código ou NF:")
            if busca:
                df = df[df.astype(str).apply(lambda row: row.str.contains(busca, case=False).any(), axis=1)]

            df['Estoque Mínimo'] = df['Estoque_Minimo_Val'].apply(
                lambda x: "" if pd.isna(x) or x == 0 else f"{x:.1f}".rstrip('0').rstrip('.')
            )

            colunas_exibicao = ['Codigo_Formatado', 'Descricao', 'Quantidade_Atual', 'Qtd_Na_Tecnica', 'Estoque Mínimo', 'Nota_Fiscal', 'Ultima_Atualizacao']
            colunas_exibicao = [c for c in colunas_exibicao if c in df.columns]

            def destacar_estoque_minimo(row):
                estilo = [''] * len(row)
                val_min = row.get('Estoque_Minimo_Val')
                if pd.notna(val_min) and val_min > 0 and row['Quantidade_Atual'] <= val_min:
                    if 'Quantidade_Atual' in row.index:
                        idx = row.index.get_loc('Quantidade_Atual')
                        estilo[idx] = 'background-color: rgba(255, 75, 75, 0.15); color: #ff8c8c;'
                return estilo

            df_estilizado = df[colunas_exibicao].style.apply(destacar_estoque_minimo, axis=1)

            st.dataframe(
                df_estilizado,
                height=700,
                use_container_width=True,
                column_config={
                    "Codigo_Formatado": st.column_config.TextColumn("Código da Peça (🛠️=Na Técnica | *=Várias NFs)"),
                    "Descricao": st.column_config.TextColumn("Descrição", width="large"),
                    "Quantidade_Atual": st.column_config.NumberColumn("Qtd\nAtual", width="small"),
                    "Qtd_Na_Tecnica": st.column_config.NumberColumn("Na\nTécnica", width="small"),
                    "Estoque Mínimo": st.column_config.TextColumn("Estoque Mínimo", width="small"),
                    "Nota_Fiscal": st.column_config.TextColumn("Nota Fiscal")
                },
                hide_index=True
            )
        else:
            st.info("Nenhum item com saldo positivo no estoque no momento.")
    else:
        st.info("Estoque vazio.")

elif aba == "📋 Histórico":
    st.header("Histórico de Movimentações")
    
    if st.button("🔄 Atualizar Histórico", use_container_width=False):
        st.cache_data.clear()
        st.rerun()

    dados_mov = ler_dados_planilha("Movimentacoes", usar_formula=True)

    if dados_mov:
        df_mov = pd.DataFrame(dados_mov)

        def extrair_url(valor):
            if isinstance(valor, str):
                if 'HYPERLINK' in valor.upper():
                    match = re.search(r'"(http.*?)"', valor)
                    return match.group(1) if match else ""
                elif valor.startswith('http'):
                    return valor
            return ""

        if 'NF_Entrada_File' in df_mov.columns:
            df_mov['NF_Entrada_File'] = df_mov['NF_Entrada_File'].apply(extrair_url)
        if 'NF_Saida_File' in df_mov.columns:
            df_mov['NF_Saida_File'] = df_mov['NF_Saida_File'].apply(extrair_url)

        st.dataframe(
            df_mov.head(200),
            height=700,
            use_container_width=True,
            column_config={
                "Data_Hora": st.column_config.TextColumn("Data e Hora"),
                "Tipo_Movimentacao": st.column_config.TextColumn("Operação / Status"),
                "Codigo_Peca": st.column_config.TextColumn("Código da Peça"),
                "Descricao": st.column_config.TextColumn("Descrição", width="large"),
                "NF_Entrada_File": st.column_config.LinkColumn("NF Entrada", display_text="🔗 Abrir NF", width="small"),
                "NF_Saida_File": st.column_config.LinkColumn("NF Saída", display_text="🔗 Abrir Doc", width="small")
            },
            hide_index=True
        )
    else:
        st.info("Nenhuma movimentação registrada.")

elif aba == "🛠️ Área Técnica":
    st.header("🛠️ Controle da Área Técnica (Bancada & Campo)")

    dados_estoque_atual = ler_dados_planilha("Estoque_Atual")
    dados_tec = ler_dados_planilha("Area_Tecnica")

    aba_tec_sub = st.tabs(["📤 Enviar Lote para Área Técnica", "📋 Peças na Área Técnica & Devolução por Checkbox"])

    with aba_tec_sub[0]:
        st.subheader("1. Identificação do Atendimento Técnico")
        col_at1, col_at2, col_at3 = st.columns([2, 2, 2])
        with col_at1:
            pedido_os = st.text_input("Pedido / OS / Atendimento*", placeholder="Ex: OS-1052")
        with col_at2:
            nome_tecnico = st.text_input("Técnico Responsável*", placeholder="Ex: Robinho")
        with col_at3:
            nome_cliente = st.text_input("Nome do Cliente (Opcional)")

        obs_tec = st.text_input("Observações Gerais")

        st.divider()

        st.subheader("2. Seleção das Peças em Lote (Diagnóstico / Manutenção)")
        
        opcoes_tec_disp = []
        for r in dados_estoque_atual:
            c_item = str(r.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
            q_item = pd.to_numeric(r.get('Quantidade_Atual', 0), errors='coerce') or 0
            raw_nf = str(r.get('Nota_Fiscal', '')).strip()
            desc_item = str(r.get('Descricao', '')).strip()
            if c_item and q_item > 0:
                opcoes_tec_disp.append(f"{c_item} | NF {raw_nf} | Saldo: {int(q_item)} un | {desc_item}")

        if not opcoes_tec_disp:
            st.warning("⚠️ Não há peças no estoque para transferir à Área Técnica.")
        else:
            df_tec_template = pd.DataFrame([
                {"Selecione Peça e NF no Estoque": "", "Quantidade": 1}
            ])

            editor_tec = st.data_editor(
                df_tec_template,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "Selecione Peça e NF no Estoque": st.column_config.SelectboxColumn(
                        "Selecione Peça e NF de Origem",
                        options=[""] + sorted(opcoes_tec_disp),
                        width="large"
                    ),
                    "Quantidade": st.column_config.NumberColumn("Qtd Enviar", min_value=1, step=1, default=1, width="small")
                }
            )

            if st.button("🚀 Transferir Lote para Área Técnica", use_container_width=True, type="primary"):
                if not pedido_os.strip() or not nome_tecnico.strip():
                    st.error("Por favor, preencha o número da OS/Pedido e o Técnico Responsável.")
                else:
                    itens_transferir = []
                    for idx, row in editor_tec.iterrows():
                        opcao_t = str(row.get("Selecione Peça e NF no Estoque", "")).strip()
                        qtd_t = int(row.get("Quantidade", 1))

                        if opcao_t:
                            partes_t = opcao_t.split(" | ")
                            cod_t = partes_t[0].strip()
                            raw_nf_t = partes_t[1].replace("NF ", "").strip()
                            desc_t = partes_t[3].strip() if len(partes_t) > 3 else ""
                            itens_transferir.append({"codigo": cod_t, "nf_origem": raw_nf_t, "descricao": desc_t, "qtd": qtd_t})

                    if not itens_transferir:
                        st.error("Selecione pelo menos uma peça na tabela acima.")
                    else:
                        with st.spinner("Transferindo para Área Técnica..."):
                            data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                            data_hora_segura = f"'{data_hora_bruta}"

                            sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                            sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                            sheet_tec = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Area_Tecnica")

                            novas_linhas_mov_t = []
                            novas_linhas_tec_t = []
                            linhas_est_a_deletar = []
                            novas_linhas_est_t = []

                            dados_est_cru = sheet_est.get_all_records()

                            for item_t in itens_transferir:
                                c_seg = f"'{item_t['codigo']}"
                                novas_linhas_mov_t.append([data_hora_segura, "Área Técnica", item_t["nf_origem"], c_seg, item_t["descricao"], item_t["qtd"], "", f"OS: {pedido_os} | Técnico: {nome_tecnico}"])

                                for idx_e, row_e in enumerate(dados_est_cru):
                                    cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                    nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                    if cod_e == item_t['codigo'] and nf_e == item_t['nf_origem'].lstrip('0'):
                                        linhas_est_a_deletar.append(idx_e + 2)

                                formula_saldo = f'=SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_t["codigo"]}"; Movimentacoes!C:C; "{item_t["nf_origem"]}"; Movimentacoes!B:B; "Entrada") + SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_t["codigo"]}"; Movimentacoes!C:C; "{item_t["nf_origem"]}"; Movimentacoes!B:B; "Retorno Técnica") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_t["codigo"]}"; Movimentacoes!C:C; "{item_t["nf_origem"]}"; Movimentacoes!B:B; "Saida") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{item_t["codigo"]}"; Movimentacoes!C:C; "{item_t["nf_origem"]}"; Movimentacoes!B:B; "Área Técnica")'
                                novas_linhas_est_t.append([c_seg, item_t["descricao"], formula_saldo, item_t["nf_origem"], 0, data_hora_segura])

                                novas_linhas_tec_t.append([c_seg, item_t["descricao"], item_t["qtd"], item_t["nf_origem"], pedido_os, nome_cliente, nome_tecnico, "Diagnóstico / Manutenção", data_hora_segura, obs_tec])

                            for r_del in sorted(list(set(linhas_est_a_deletar)), reverse=True):
                                sheet_est.delete_rows(r_del)

                            if novas_linhas_mov_t:
                                sheet_mov.insert_rows(novas_linhas_mov_t, row=2, value_input_option='USER_ENTERED')
                            if novas_linhas_est_t:
                                sheet_est.insert_rows(novas_linhas_est_t, row=2, value_input_option='USER_ENTERED')
                            if novas_linhas_tec_t:
                                sheet_tec.insert_rows(novas_linhas_tec_t, row=2, value_input_option='USER_ENTERED')

                            st.cache_data.clear()
                            st.success(f"🎉 **TRANSFERÊNCIA CONCLUÍDA COM SUCESSO!** {len(itens_transferir)} item(ns) enviado(s) para a Área Técnica.")
                            time.sleep(2)
                            st.rerun()

    with aba_tec_sub[1]:
        st.subheader("Painel de Controle e Devolução Rápida ao Estoque")
        st.caption("☑️ Marque as caixinhas na tabela dos itens que deseja devolver e clique no botão abaixo.")

        if dados_tec:
            df_tec = pd.DataFrame(dados_tec)
            if not df_tec.empty:
                cols_texto = ["Codigo_Peca", "Descricao", "Nota_Fiscal", "Pedido_OS", "Cliente", "Tecnico", "Status", "Observacao"]
                for col in cols_texto:
                    if col in df_tec.columns:
                        df_tec[col] = df_tec[col].astype(str)

                df_tec.insert(0, "Devolver?", False)
                df_tec["_Row_Idx"] = df_tec.index + 2

                df_tec_edit = st.data_editor(
                    df_tec,
                    height=600,
                    hide_index=True,
                    use_container_width=True,
                    column_config={
                        "Devolver?": st.column_config.CheckboxColumn("Devolver ao Estoque?", default=False, width="small"),
                        "Codigo_Peca": st.column_config.TextColumn("Código Peça"),
                        "Descricao": st.column_config.TextColumn("Descrição", width="large"),
                        "Quantidade": st.column_config.NumberColumn("Qtd", width="small"),
                        "Nota_Fiscal": st.column_config.TextColumn("NF Origem"),
                        "Pedido_OS": st.column_config.TextColumn("OS / Pedido"),
                        "Tecnico": st.column_config.TextColumn("Técnico"),
                        "_Row_Idx": None
                    }
                )

                if st.button("↩️ Devolver Peças Selecionadas ao Estoque 02.08", type="primary", use_container_width=True):
                    itens_para_devolver = df_tec_edit[df_tec_edit["Devolver?"] == True]

                    if itens_para_devolver.empty:
                        st.warning("Selecione pelo menos uma peça marcando a caixinha 'Devolver?'.")
                    else:
                        with st.spinner("Processando retorno das peças selecionadas..."):
                            data_hora_bruta = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                            data_hora_segura = f"'{data_hora_bruta}"

                            sheet_mov = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Movimentacoes")
                            sheet_est = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Estoque_Atual")
                            sheet_tec = client_sheets.open_by_key(SPREADSHEET_ID).worksheet("Area_Tecnica")

                            linhas_sheet_deletar = sorted(itens_para_devolver["_Row_Idx"].tolist(), reverse=True)

                            for r_idx in linhas_sheet_deletar:
                                sheet_tec.delete_rows(r_idx)

                            novas_mov_dev = []
                            novas_est_dev = []
                            linhas_est_substituir = []

                            dados_est_cru = sheet_est.get_all_records()

                            for _, row_d in itens_para_devolver.iterrows():
                                c_d = str(row_d.get('Codigo_Peca', '')).replace("'", "").strip()
                                nf_d = str(row_d.get('Nota_Fiscal', '')).strip()
                                q_d = int(row_d.get('Quantidade', 1))
                                desc_d = str(row_d.get('Descricao', '')).strip()

                                c_seg = f"'{c_d}"
                                novas_mov_dev.append([data_hora_segura, "Retorno Técnica", nf_d, c_seg, desc_d, q_d, "", "Retornou da bancada/campo"])

                                for idx_e, row_e in enumerate(dados_est_cru):
                                    cod_e = str(row_e.get('Codigo_Peca', '')).replace("'", "").replace("*", "").replace("🛠️", "").strip()
                                    nf_e = str(row_e.get('Nota_Fiscal', '')).strip().lstrip('0')
                                    if cod_e == c_d and nf_e == nf_d.lstrip('0'):
                                        linhas_est_substituir.append(idx_e + 2)

                                formula_saldo = f'=SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{c_d}"; Movimentacoes!C:C; "{nf_d}"; Movimentacoes!B:B; "Entrada") + SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{c_d}"; Movimentacoes!C:C; "{nf_d}"; Movimentacoes!B:B; "Retorno Técnica") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{c_d}"; Movimentacoes!C:C; "{nf_d}"; Movimentacoes!B:B; "Saida") - SUMIFS(Movimentacoes!F:F; Movimentacoes!D:D; "{c_d}"; Movimentacoes!C:C; "{nf_d}"; Movimentacoes!B:B; "Área Técnica")'
                                novas_est_dev.append([c_seg, desc_d, formula_saldo, nf_d, 0, data_hora_segura])

                            for r_del in sorted(list(set(linhas_est_substituir)), reverse=True):
                                sheet_est.delete_rows(r_del)

                            if novas_mov_dev:
                                sheet_mov.insert_rows(novas_mov_dev, row=2, value_input_option='USER_ENTERED')
                            if novas_est_dev:
                                sheet_est.insert_rows(novas_est_dev, row=2, value_input_option='USER_ENTERED')

                            st.cache_data.clear()
                            st.success(f"🎉 **DEVOLUÇÃO CONCLUÍDA COM SUCESSO!** {len(itens_para_devolver)} item(ns) retornado(s) ao Estoque 02.08.")
                            time.sleep(2)
                            st.rerun()

            else:
                st.info("Nenhuma peça atualmente na área técnica.")
        else:
            st.info("Nenhum registro de área técnica.")

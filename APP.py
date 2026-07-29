import streamlit as st
import pandas as pd
import io
import re
import requests
from datetime import datetime

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
ESPECIE_UNICA       = "39"
ACUMULADOR_EXTERIOR = "2551"
CFOP_SP             = "1933"
CFOP_FORA           = "2933"
COD_ISS_RETIDO      = "18"
COD_ISS_NORMAL      = "3"

# ─────────────────────────────────────────────
# URLs DOS ARQUIVOS NO GITHUB
# ─────────────────────────────────────────────
GITHUB_RAW_BASE  = "https://raw.githubusercontent.com/roneferrera/NFTS-LEIAUTEPADRAODOMINIOSISTEMAS/main"
URL_ACUMULADORES = f"{GITHUB_RAW_BASE}/Acumuladores.xlsx"
URL_PAISES       = f"{GITHUB_RAW_BASE}/Pa%C3%ADses.xlsx"

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def safe(row, col, default=""):
    val = row.get(col, default)
    if pd.isna(val):
        return default
    return str(val).strip()

def fmt_decimal(value, decimals=2):
    try:
        s = str(value).strip().replace(".", "").replace(",", ".")
        v = float(s)
        return f"{v:.{decimals}f}"
    except Exception:
        return f"0.{'0'*decimals}"

def fmt_date(value):
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return ""
    s = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue
    return s

def limpa_cnpj(value):
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return ""
    return re.sub(r"\D", "", str(value))

def limpa_numero(value):
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return ""
    try:
        return str(int(float(str(value).replace(",", "."))))
    except Exception:
        return str(value).strip()

# ─────────────────────────────────────────────
# CARREGA ACUMULADORES DO EXCEL
# ─────────────────────────────────────────────
def carrega_acumuladores(file_bytes):
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Acumuladores",
        dtype=str
    )
    df.columns = [c.strip() for c in df.columns]
    lookup = {}
    for _, row in df.iterrows():
        paulistana = str(row.get("PAULISTANA", "")).strip()
        paulistana = re.sub(r"\.0$", "", paulistana).strip()
        acumulador = str(row.get("Codigo ACUMULADOR", "")).strip()
        acumulador = re.sub(r"\.0$", "", acumulador).strip()
        if paulistana and paulistana not in ("", "nan"):
            lookup[paulistana] = acumulador
    return lookup

# ─────────────────────────────────────────────
# CARREGA PAÍSES DO EXCEL
# ─────────────────────────────────────────────
def carrega_paises(file_bytes):
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="RELAÇÃO DE PAÍSES",
        dtype=str
    )
    df.columns = [c.strip() for c in df.columns]
    lookup = {}
    for _, row in df.iterrows():
        nome   = str(row.get("Nome", "")).strip().upper()
        codigo = str(row.get("Código", "")).strip()
        codigo = re.sub(r"\.0$", "", codigo).strip()
        if nome and codigo and nome != "NAN":
            lookup[nome] = codigo
    return lookup

# ─────────────────────────────────────────────
# CARREGA DO GITHUB COM CACHE
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carrega_acumuladores_github(url: str) -> dict:
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    return carrega_acumuladores(resp.content)

@st.cache_data(show_spinner=False)
def carrega_paises_github(url: str) -> dict:
    resp = requests.get(url, timeout=15)
    resp.raise_for_status()
    return carrega_paises(resp.content)

# ─────────────────────────────────────────────
# REGRAS DE NEGÓCIO
# ─────────────────────────────────────────────
def determina_especie(_row) -> str:
    return ESPECIE_UNICA

def determina_cfop(row) -> str:
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    ind_num = re.sub(r"\.0$", "", ind).strip()
    if ind_num == "3":
        return CFOP_FORA
    uf = safe(row, "UF do Prestador").upper().strip()
    return CFOP_SP if uf == "SP" else CFOP_FORA

def determina_acumulador(row, lookup_acum: dict) -> str:
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    ind_num = re.sub(r"\.0$", "", ind).strip()
    if ind_num == "3":
        return ACUMULADOR_EXTERIOR
    paulistana_raw = safe(row, "Código do Serviço Prestado na NFTS")
    paulistana = limpa_numero(paulistana_raw)
    acum = lookup_acum.get(paulistana, "")
    if not acum:
        return f"AVISO: PAULISTANA {paulistana} NAO MAPEADA"
    return acum

def determina_cod_iss(row) -> str:
    retido = safe(row, "ISS Retido").upper().strip()
    return COD_ISS_RETIDO if retido == "S" else COD_ISS_NORMAL

def determina_fornecedor(row) -> str:
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    ind_num = re.sub(r"\.0$", "", ind).strip()
    if ind_num == "3":
        return ""
    return limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))

def determina_uf(row) -> str:
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    ind_num = re.sub(r"\.0$", "", ind).strip()
    if ind_num == "3":
        return "EX"
    return safe(row, "UF do Prestador").strip()

def determina_cod_pais(row, lookup_pais: dict) -> str:
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    ind_num = re.sub(r"\.0$", "", ind).strip()
    if ind_num != "3":
        return ""
    bairro   = safe(row, "Bairro do Prestador").upper().strip()
    cidade   = safe(row, "Cidade do Prestador").upper().strip()
    endereco = safe(row, "Endereço do Prestador").upper().strip()
    if any(x in bairro + cidade + endereco for x in ["SACRAMENTO", "NORTH STREET", "USA", "EUA"]):
        return lookup_pais.get("ESTADOS UNIDOS", "76")
    return lookup_pais.get("ESTADOS UNIDOS", "76")

# ─────────────────────────────────────────────
# MONTA REGISTROS
# ─────────────────────────────────────────────
def reg_0000(cnpj_empresa: str) -> str:
    return "|".join(["0000", cnpj_empresa])

def reg_0020(row, lookup_pais: dict) -> str:
    ind_num     = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    fornecedor  = determina_fornecedor(row)
    razao       = safe(row, "Razão Social do Prestador")[:150]
    endereco    = safe(row, "Endereço do Prestador")
    numero_end  = limpa_numero(safe(row, "Número do Endereço do Prestador"))
    complemento = safe(row, "Complemento do Endereço do Prestador")
    bairro      = safe(row, "Bairro do Prestador")
    uf          = determina_uf(row)
    cep         = re.sub(r"\D", "", safe(row, "CEP do Prestador"))
    cod_pais    = determina_cod_pais(row, lookup_pais) if ind_num == "3" else ""
    email       = safe(row, "Email do Prestador")

    campos = [
        "0020", fornecedor, razao, "", endereco, numero_end,
        complemento, bairro, "", uf, cod_pais, cep,
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", email,
        "", "", "", "",
    ]
    return "|".join(campos)

def reg_1000(row, lookup_acum: dict, lookup_pais: dict) -> str:
    especie    = determina_especie(row)
    fornecedor = determina_fornecedor(row)
    acumulador = determina_acumulador(row, lookup_acum)
    cfop       = determina_cfop(row)
    num_doc    = limpa_numero(safe(row, "Número do Documento"))
    serie      = safe(row, "Série do Documento")
    if serie in ("-", "nan", ""):
        serie = ""
    dt_entrada = fmt_date(safe(row, "Data da Prestação de Serviços"))
    dt_emissao = fmt_date(safe(row, "Data Hora Emissão NFTS"))
    valor      = fmt_decimal(safe(row, "Valor dos Serviços"))
    cod_iss    = determina_cod_iss(row)

    campos = [
        "1000", especie, fornecedor, "", acumulador, cfop, "",
        num_doc, serie, "", dt_entrada, dt_emissao, valor, "", "", "", "", "", "",
        cod_iss, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", "", "", "", "",
    ]
    return "|".join(campos)

def reg_1020(row) -> str:
    iss_retido    = safe(row, "ISS Retido").upper().strip()
    valor_iss_raw = safe(row, "Valor ISS")
    aliquota_raw  = safe(row, "Alíquota")
    valor_serv    = fmt_decimal(safe(row, "Valor dos Serviços"))
    valor_iss     = fmt_decimal(valor_iss_raw)
    aliquota      = fmt_decimal(aliquota_raw)
    cod_iss       = determina_cod_iss(row)

    try:
        v_iss = float(str(valor_iss_raw).strip().replace(".", "").replace(",", "."))
    except Exception:
        v_iss = 0.0

    if v_iss == 0.0 and iss_retido != "S":
        return ""

    campos = [
        "1020", cod_iss, "", valor_serv, aliquota, valor_iss,
        "", "", "", "", valor_serv, "", "", "", "", "", "", "", "",
    ]
    return "|".join(campos)

def reg_1150(_row) -> str:
    return "|".join(["1150", "", "", "", ""])

def reg_1151(_row) -> str:
    return "|".join(["1151", "", "", "", ""])

# ─────────────────────────────────────────────
# PROCESSAMENTO PRINCIPAL
# ─────────────────────────────────────────────
def converte_nfts(
    df_notas, lookup_acum, lookup_pais, cnpj_empresa,
    incluir_0000, incluir_0020, incluir_1020, incluir_1150, incluir_1151,
) -> tuple[str, pd.DataFrame]:

    linhas  = []
    preview = []
    fornecedores_vistos = set()

    if incluir_0000:
        linhas.append(reg_0000(cnpj_empresa))

    for _, row in df_notas.iterrows():
        especie    = determina_especie(row)
        cfop       = determina_cfop(row)
        acumulador = determina_acumulador(row, lookup_acum)
        fornecedor = determina_fornecedor(row)
        uf         = determina_uf(row)
        cod_iss    = determina_cod_iss(row)
        num_nfts   = safe(row, "Nº NFTS")
        razao      = safe(row, "Razão Social do Prestador")
        valor      = fmt_decimal(safe(row, "Valor dos Serviços"))
        valor_iss  = fmt_decimal(safe(row, "Valor ISS"))
        iss_retido = safe(row, "ISS Retido").upper().strip()
        dt_entrada = fmt_date(safe(row, "Data da Prestação de Serviços"))
        dt_emissao = fmt_date(safe(row, "Data Hora Emissão NFTS"))

        if incluir_0020:
            chave_forn = fornecedor if fornecedor else razao
            if chave_forn not in fornecedores_vistos:
                fornecedores_vistos.add(chave_forn)
                linhas.append(reg_0020(row, lookup_pais))

        linhas.append(reg_1000(row, lookup_acum, lookup_pais))

        if incluir_1020:
            linha_1020 = reg_1020(row)
            if linha_1020:
                linhas.append(linha_1020)

        if incluir_1150:
            linhas.append(reg_1150(row))

        if incluir_1151:
            linhas.append(reg_1151(row))

        preview.append({
            "NFTS":       num_nfts,
            "Prestador":  razao,
            "Especie":    especie,
            "CFOP":       cfop,
            "Acumulador": acumulador,
            "UF":         uf,
            "Dt Entrada": dt_entrada,
            "Dt Emissão": dt_emissao,
            "Valor":      valor,
            "ISS":        valor_iss,
            "ISS Retido": iss_retido,
            "Cód ISS":    cod_iss,
            "Fornecedor": fornecedor,
        })

    return "\n".join(linhas), pd.DataFrame(preview)

# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="Conversor NFTS", layout="wide")
st.title("📄 Conversor NFTS → Arquivo de Importação")

# ── Sidebar ──────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configurações")

    cnpj_empresa = st.text_input(
        "CNPJ da Empresa (só números)",
        value="20586841000130",
        max_chars=14,
    )

    st.markdown("---")
    st.subheader("Registros a gerar")
    incluir_0000 = st.checkbox("Reg. 0000 (Empresa)",    value=True)
    incluir_0020 = st.checkbox("Reg. 0020 (Fornecedor)", value=True)
    incluir_1020 = st.checkbox("Reg. 1020 (ISS)",        value=True)
    incluir_1150 = st.checkbox("Reg. 1150 (IVA/IBS)",    value=False)
    incluir_1151 = st.checkbox("Reg. 1151 (IVA/CBS)",    value=False)

    st.markdown("---")
    st.info(
        "**Regras aplicadas:**\n\n"
        "**Especie** — 39 para todas as notas\n\n"
        "**CFOP** — 1933 (SP) / 2933 (outros/EXT)\n\n"
        "**Acumulador** — 2551 (exterior) / lookup PAULISTANA\n\n"
        "**ISS** — Cód.18 (retido) / Cód.3 (normal)\n\n"
        "**Fornecedor** — CNPJ (nacional) / vazio (exterior)\n\n"
        "**UF** — EX (exterior) / UF real (nacional)"
    )

# ── Carrega Acumuladores e Países do GitHub ──
with st.spinner("Carregando tabelas de referência do GitHub..."):
    try:
        lookup_acum = carrega_acumuladores_github(URL_ACUMULADORES)
        lookup_pais = carrega_paises_github(URL_PAISES)
        st.sidebar.success(
            f"✅ {len(lookup_acum)} acumuladores | {len(lookup_pais)} países carregados"
        )
    except Exception as e:
        st.error(
            f"❌ Erro ao carregar arquivos do GitHub: {e}\n\n"
            "Verifique se a URL está correta e o repositório é público."
        )
        st.stop()

# ── Upload apenas do CSV NFTS ─────────────────
st.subheader("📂 Upload do Arquivo NFTS")

file_nfts = st.file_uploader(
    "CSV NFTS",
    type=["csv"],
    key="nfts",
    help="Arquivo exportado do portal NFTS da Prefeitura de SP"
)

# ── Processamento ─────────────────────────────
if file_nfts:

    try:
        df_nfts = pd.read_csv(file_nfts, sep=",", dtype=str, encoding="utf-8")
    except Exception:
        file_nfts.seek(0)
        try:
            df_nfts = pd.read_csv(file_nfts, sep=";", dtype=str, encoding="utf-8")
        except Exception:
            file_nfts.seek(0)
            df_nfts = pd.read_csv(file_nfts, sep=";", dtype=str, encoding="latin-1")

    df_nfts.columns = [c.strip() for c in df_nfts.columns]

    col_tipo = "Tipo de Registro"
    if col_tipo not in df_nfts.columns:
        st.error(f"Coluna '{col_tipo}' não encontrada. Colunas: {list(df_nfts.columns)}")
        st.stop()

    df_notas = df_nfts[df_nfts[col_tipo].str.strip().str.upper() == "4"].copy()

    if df_notas.empty:
        st.warning("Nenhuma nota com Tipo de Registro = 4 encontrada no CSV.")
        st.stop()

    st.success(
        f"✅ {len(df_notas)} nota(s) carregada(s) | "
        f"{len(lookup_acum)} acumuladores | {len(lookup_pais)} países"
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Preview", "📄 Arquivo Gerado", "🔍 Debug", "❓ Ajuda"
    ])

    conteudo, df_prev = converte_nfts(
        df_notas, lookup_acum, lookup_pais, cnpj_empresa,
        incluir_0000, incluir_0020, incluir_1020, incluir_1150, incluir_1151,
    )

    # ── TAB 1: PREVIEW ────────────────────────
    with tab1:
        st.subheader("Preview das Notas")

        if not df_prev.empty:
            n_ret   = len(df_prev[df_prev["ISS Retido"] == "S"])
            n_nor   = len(df_prev[df_prev["ISS Retido"] != "S"])
            n_1933  = len(df_prev[df_prev["CFOP"] == CFOP_SP])
            n_2933  = len(df_prev[df_prev["CFOP"] == CFOP_FORA])
            n_aviso = len(df_prev[df_prev["Acumulador"].str.startswith("AVISO", na=False)])

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Total notas",          len(df_prev))
            m2.metric("ISS Retido (cód.18)",  n_ret)
            m3.metric("ISS Normal (cód.3)",   n_nor)
            m4.metric("CFOP 1933 (SP)",       n_1933)
            m5.metric("CFOP 2933 (fora/EXT)", n_2933)
            m6.metric("Acum. não mapeado",    n_aviso, delta_color="inverse")

            def highlight_row(row):
                if str(row.get("Acumulador", "")).startswith("AVISO"):
                    cor = "#f8d7da"
                elif row.get("ISS Retido", "") == "S":
                    cor = "#d4edda"
                else:
                    cor = "#fff3cd"
                return [f"background-color: {cor}"] * len(row)

            st.dataframe(
                df_prev.style.apply(highlight_row, axis=1),
                use_container_width=True,
            )

            col_l1, col_l2, col_l3 = st.columns(3)
            col_l1.markdown(
                "<span style='background:#d4edda;padding:2px 8px;border-radius:4px'>"
                "🟢 Verde = ISS Retido (Cód.18)</span>", unsafe_allow_html=True)
            col_l2.markdown(
                "<span style='background:#fff3cd;padding:2px 8px;border-radius:4px'>"
                "🟡 Amarelo = ISS Normal (Cód.3)</span>", unsafe_allow_html=True)
            col_l3.markdown(
                "<span style='background:#f8d7da;padding:2px 8px;border-radius:4px'>"
                "🔴 Vermelho = Acumulador não mapeado</span>", unsafe_allow_html=True)

    # ── TAB 2: ARQUIVO GERADO ─────────────────
    with tab2:
        st.subheader("Arquivo de Importação Gerado")

        if conteudo:
            st.code(conteudo, language="text")

            nome_arquivo = f"importacao_nfts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            st.download_button(
                label="⬇️ Baixar arquivo .txt",
                data=conteudo.encode("utf-8"),
                file_name=nome_arquivo,
                mime="text/plain",
            )

            avisos = [l for l in conteudo.splitlines() if "AVISO" in l]
            if avisos:
                st.warning(f"⚠️ {len(avisos)} linha(s) com acumulador não mapeado:")
                for a in avisos:
                    st.code(a)

    # ── TAB 3: DEBUG ──────────────────────────
    with tab3:
        st.subheader("🔍 Debug — Dados Carregados")

        with st.expander("CSV NFTS — Dados brutos"):
            st.dataframe(df_notas, use_container_width=True)

        with st.expander(f"Acumuladores carregados ({len(lookup_acum)} itens)"):
            st.dataframe(
                pd.DataFrame(list(lookup_acum.items()), columns=["PAULISTANA", "Acumulador"]),
                use_container_width=True,
            )

        with st.expander(f"Países carregados ({len(lookup_pais)} itens)"):
            st.dataframe(
                pd.DataFrame(list(lookup_pais.items()), columns=["Nome", "Código"]),
                use_container_width=True,
            )

        with st.expander("Mapeamento campo a campo por nota"):
            for _, row in df_notas.iterrows():
                nfts = safe(row, "Nº NFTS")
                st.markdown(f"**NFTS {nfts}**")
                debug_data = {
                    "Campo": [
                        "Indicador Prestador", "CPF/CNPJ Prestador", "Razão Social",
                        "UF Prestador", "PAULISTANA", "ISS Retido",
                        "Valor Serviços", "Valor ISS", "Alíquota",
                        "→ Especie", "→ CFOP", "→ Acumulador",
                        "→ Fornecedor", "→ UF", "→ Cód ISS", "→ Cód País",
                    ],
                    "Valor": [
                        safe(row, "Indicador de CPF/CNPJ do Prestador"),
                        safe(row, "CPF/CNPJ do Prestador"),
                        safe(row, "Razão Social do Prestador"),
                        safe(row, "UF do Prestador"),
                        limpa_numero(safe(row, "Código do Serviço Prestado na NFTS")),
                        safe(row, "ISS Retido"),
                        safe(row, "Valor dos Serviços"),
                        safe(row, "Valor ISS"),
                        safe(row, "Alíquota"),
                        determina_especie(row),
                        determina_cfop(row),
                        determina_acumulador(row, lookup_acum),
                        determina_fornecedor(row),
                        determina_uf(row),
                        determina_cod_iss(row),
                        determina_cod_pais(row, lookup_pais),
                    ]
                }
                st.dataframe(
                    pd.DataFrame(debug_data),
                    use_container_width=True,
                    hide_index=True,
                )
                st.markdown("---")

    # ── TAB 4: AJUDA ──────────────────────────
    with tab4:
        st.subheader("❓ Ajuda — Regras de Importação")

        st.markdown("### Regras automáticas aplicadas")
        st.dataframe(pd.DataFrame([
            ["Indicador=3 (Exterior)", "39", "2933", "2551 - SERVIÇOS TOMADOS IMPORTAÇÃO", "EX",      "76 (EUA)"],
            ["Indicador=1/2, UF=SP",   "39", "1933", "Lookup PAULISTANA",                  "SP",      ""],
            ["Indicador=1/2, UF≠SP",   "39", "2933", "Lookup PAULISTANA",                  "UF real", ""],
            ["Indicador=1/2, UF vazia","39", "2933", "Lookup PAULISTANA",                  "-",       ""],
        ], columns=["Situação", "Espécie", "CFOP", "Acumulador", "UF", "Cód País"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Campos principais do Registro 1000")
        st.dataframe(pd.DataFrame([
            ["Campo 2",  "Espécie",    "Todas as notas",                     "39"],
            ["Campo 3",  "Fornecedor", "Nacional = CNPJ / Exterior = vazio", "—"],
            ["Campo 5",  "Acumulador", "Exterior = 2551 / Nacional = lookup","—"],
            ["Campo 6",  "CFOP",       "SP = 1933 / Outros/EXT = 2933",      "—"],
            ["Campo 11", "Dt Entrada", "Data da Prestação de Serviços",      "dd/mm/aaaa"],
            ["Campo 12", "Dt Emissão", "Data Hora Emissão NFTS",             "dd/mm/aaaa"],
            ["Campo 13", "Valor",      "Valor dos Serviços",                 "decimal"],
            ["Campo 20", "Cód ISS",    "ISS Retido S = 18 / N = 3",         "—"],
        ], columns=["Campo", "Nome", "Regra", "Valor/Formato"]),
        use_container_width=True, hide_index=True)

else:
    st.info(
        "👆 Faça upload do CSV NFTS para iniciar.\n\n"
        "Os arquivos **Acumuladores.xlsx** e **Países.xlsx** são carregados "
        "automaticamente do GitHub."
    )

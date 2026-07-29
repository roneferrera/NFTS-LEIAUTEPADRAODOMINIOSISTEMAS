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
# URLs RAW DO GITHUB
# ─────────────────────────────────────────────
_BASE = "https://raw.githubusercontent.com/roneferrera/NFTS-LEIAUTEPADRAODOMINIOSISTEMAS/main"
URL_ACUMULADORES = f"{_BASE}/Acumuladores.xlsx"
URL_PAISES       = f"{_BASE}/Pa%C3%ADses.xlsx"

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def safe(row, col, default=""):
    val = row.get(col, default)
    if pd.isna(val):
        return default
    return str(val).strip()

def fmt_decimal(value, casas=2):
    try:
        s = str(value).strip()
        if not s or s in ("nan", ""):
            return "0"
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        v = float(s)
        return f"{v:.{casas}f}".replace(".", ",")
    except Exception:
        return "0"

def fmt_date(value):
    if not value or str(value).strip() in ("", "nan"):
        return ""
    s = str(value).strip()
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except Exception:
            continue
    return s

def limpa_cnpj(value):
    if not value or str(value).strip() in ("", "nan"):
        return ""
    return re.sub(r"\D", "", str(value))

def limpa_numero(value):
    if not value or str(value).strip() in ("", "nan"):
        return ""
    try:
        return str(int(float(str(value).replace(",", "."))))
    except Exception:
        return str(value).strip()

def monta_linha(campos: list) -> str:
    return "|".join([str(c) for c in campos]) + "|"

# ─────────────────────────────────────────────
# CARREGA ACUMULADORES DO EXCEL
# ─────────────────────────────────────────────
def _parse_acumuladores(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Acumuladores", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    lookup = {}
    for _, row in df.iterrows():
        paulistana = re.sub(r"\.0$", "", str(row.get("PAULISTANA", "")).strip())
        acumulador = re.sub(r"\.0$", "", str(row.get("Codigo ACUMULADOR", "")).strip())
        if paulistana and paulistana not in ("", "nan"):
            lookup[paulistana] = acumulador
    return lookup

# ─────────────────────────────────────────────
# CARREGA PAÍSES DO EXCEL
# ─────────────────────────────────────────────
def _parse_paises(file_bytes: bytes) -> dict:
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="RELAÇÃO DE PAÍSES", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    lookup = {}
    for _, row in df.iterrows():
        nome   = str(row.get("Nome", "")).strip().upper()
        codigo = re.sub(r"\.0$", "", str(row.get("Código", "")).strip())
        if nome and codigo and nome != "NAN":
            lookup[nome] = codigo
    return lookup

# ─────────────────────────────────────────────
# CARREGA DO GITHUB COM CACHE
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def carrega_acumuladores_github(url: str) -> dict:
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    return _parse_acumuladores(resp.content)

@st.cache_data(show_spinner=False)
def carrega_paises_github(url: str) -> dict:
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()
    return _parse_paises(resp.content)

# ─────────────────────────────────────────────
# REGRAS DE NEGÓCIO
# ─────────────────────────────────────────────
def determina_especie(_row) -> str:
    return ESPECIE_UNICA

def determina_cfop(row) -> str:
    ind_num = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    if ind_num == "3":
        return CFOP_FORA
    uf = safe(row, "UF do Prestador").upper().strip()
    return CFOP_SP if uf == "SP" else CFOP_FORA

def determina_acumulador(row, lookup_acum: dict) -> str:
    ind_num = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    if ind_num == "3":
        return ACUMULADOR_EXTERIOR
    paulistana = limpa_numero(safe(row, "Código do Serviço Prestado na NFTS"))
    acum = lookup_acum.get(paulistana, "")
    if not acum:
        return f"AVISO: PAULISTANA {paulistana} NAO MAPEADA"
    return acum

def determina_cod_iss(row) -> str:
    retido = safe(row, "ISS Retido").upper().strip()
    return COD_ISS_RETIDO if retido == "S" else COD_ISS_NORMAL

def determina_fornecedor(row) -> str:
    ind_num = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    if ind_num == "3":
        return ""
    return limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))

def determina_uf(row) -> str:
    ind_num = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    if ind_num == "3":
        return "EX"
    return safe(row, "UF do Prestador").strip()

def determina_cod_pais(row, lookup_pais: dict) -> str:
    ind_num = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    if ind_num != "3":
        return ""
    return lookup_pais.get("ESTADOS UNIDOS", "76")

# ─────────────────────────────────────────────
# ALERTA: NOTAS DUPLICADAS
# ─────────────────────────────────────────────
def detecta_duplicatas(df_notas) -> pd.DataFrame:
    df = df_notas.copy()
    df["_razao"]   = df.apply(lambda r: safe(r, "Razão Social do Prestador"), axis=1)
    df["_num_doc"] = df.apply(lambda r: limpa_numero(safe(r, "Número do Documento")), axis=1)
    df["_chave"]   = df["_razao"] + "||" + df["_num_doc"]
    duplicados = df[df.duplicated(subset=["_chave"], keep=False)].copy()
    if duplicados.empty:
        return pd.DataFrame()
    return duplicados[[
        "Nº NFTS", "Razão Social do Prestador",
        "Número do Documento", "Data Hora Emissão NFTS",
        "Valor dos Serviços", "ISS Retido",
    ]].rename(columns={
        "Nº NFTS":                   "NFTS",
        "Razão Social do Prestador": "Prestador",
        "Número do Documento":       "Nº Documento",
        "Data Hora Emissão NFTS":    "Data Emissão",
        "Valor dos Serviços":        "Valor",
        "ISS Retido":                "ISS Retido",
    })

# ─────────────────────────────────────────────
# MONTA REGISTROS
# ─────────────────────────────────────────────

def reg_0000(cnpj_empresa: str) -> str:
    return monta_linha([
        "0000",
        cnpj_empresa,
    ])

def reg_0020(row, lookup_pais: dict) -> str:
    ind_num     = re.sub(r"\.0$", "", safe(row, "Indicador de CPF/CNPJ do Prestador")).strip()
    fornecedor  = determina_fornecedor(row)
    razao       = safe(row, "Razão Social do Prestador")[:150]
    apelido     = razao[:40]
    endereco    = safe(row, "Endereço do Prestador")
    numero_end  = limpa_numero(safe(row, "Número do Endereço do Prestador"))
    complemento = safe(row, "Complemento do Endereço do Prestador")
    bairro      = safe(row, "Bairro do Prestador")
    uf          = determina_uf(row)
    cep         = re.sub(r"\D", "", safe(row, "CEP do Prestador"))
    cod_pais    = determina_cod_pais(row, lookup_pais) if ind_num == "3" else ""
    email       = safe(row, "Email do Prestador")

    campos = [
        "0020",      # 1
        fornecedor,  # 2
        razao,       # 3
        apelido,     # 4
        endereco,    # 5
        numero_end,  # 6
        complemento, # 7
        bairro,      # 8
        "",          # 9
        uf,          # 10
        cod_pais,    # 11
        cep,         # 12
        "",          # 13
        "",          # 14
        "",          # 15
        "",          # 16
        "",          # 17
        "",          # 18
        "",          # 19
        "",          # 20
        "",          # 21
        "N",         # 22 - Agropecuário
        "",          # 23
        "N",         # 24 - Regime de apuração
        "N",         # 25 - Contribuinte ICMS
        "",          # 26
        "",          # 27
        "",          # 28
        email,       # 29
        "",          # 30
        "N",         # 31 - Contribuinte CPRB
        "",          # 32
        "",          # 33
    ]
    assert len(campos) == 33, f"reg_0020: esperado 33 campos, encontrado {len(campos)}"
    return monta_linha(campos)

def reg_1000(row, lookup_acum: dict, lookup_pais: dict) -> str:
    especie    = determina_especie(row)
    fornecedor = determina_fornecedor(row)
    acumulador = determina_acumulador(row, lookup_acum)
    cfop       = determina_cfop(row)
    num_doc    = limpa_numero(safe(row, "Número do Documento"))
    serie      = safe(row, "Série do Documento")
    if serie in ("-", "nan", ""):
        serie = ""

    # ✅ Ambos os campos de data usam "Data da Prestação de Serviços"
    dt_prestacao = fmt_date(safe(row, "Data da Prestação de Serviços"))
    dt_campo11   = dt_prestacao  # Campo 11 - Data da entrada
    dt_campo12   = dt_prestacao  # Campo 12 - Data emissão

    valor   = fmt_decimal(safe(row, "Valor dos Serviços"), casas=2)
    cod_iss = determina_cod_iss(row)

    campos = [
        "1000",      # 1
        especie,     # 2
        fornecedor,  # 3
        "",          # 4
        acumulador,  # 5
        cfop,        # 6
        "",          # 7
        num_doc,     # 8
        serie,       # 9
        "",          # 10
        dt_campo11,  # 11 ✅ Data da Prestação de Serviços
        dt_campo12,  # 12 ✅ Data da Prestação de Serviços
        valor,       # 13
        "",          # 14
        "",          # 15
        "",          # 16
        "",          # 17
        "",          # 18
        "",          # 19
        cod_iss,     # 20
        "",          # 21
        "",          # 22
        "",          # 23
        "",          # 24
        "",          # 25
        "",          # 26
        "",          # 27
        "",          # 28
        "",          # 29
        "",          # 30
        "",          # 31
        "",          # 32
        "",          # 33
        "",          # 34
        "",          # 35
        "",          # 36
        "",          # 37
        "",          # 38
        valor,       # 39 - Valor produtos = Valor contábil
        "",          # 40
        "",          # 41
        "",          # 42
        "",          # 43
        "",          # 44
        "",          # 45
        "",          # 46
        "",          # 47
        "",          # 48
        "",          # 49
        "",          # 50
        "",          # 51
        "",          # 52
        "",          # 53
        "",          # 54
        "",          # 55
        "",          # 56
        "",          # 57
        "",          # 58
        "",          # 59
        "",          # 60
        "",          # 61
        "",          # 62
        "",          # 63
        "",          # 64
        "",          # 65
        "",          # 66
        "",          # 67
        "",          # 68
        "",          # 69
        "",          # 70
        "",          # 71
        "",          # 72
        "",          # 73
        "",          # 74
        "",          # 75
        "",          # 76
        "",          # 77
        "",          # 78
        "",          # 79
        "",          # 80
        "",          # 81
        "",          # 82
        "",          # 83
        "",          # 84
        "",          # 85
        "",          # 86
        "",          # 87
        "",          # 88
        "",          # 89
        "",          # 90
        "",          # 91
        "",          # 92
        "",          # 93
        "",          # 94
        "",          # 95
        "",          # 96
        "",          # 97
        "",          # 98
    ]
    assert len(campos) == 98, f"reg_1000: esperado 98 campos, encontrado {len(campos)}"
    return monta_linha(campos)

def reg_1020(row) -> str:
    iss_retido    = safe(row, "ISS Retido").upper().strip()
    valor_iss_raw = safe(row, "Valor ISS")
    aliquota_raw  = safe(row, "Alíquota")
    cod_iss       = determina_cod_iss(row)
    valor_serv    = fmt_decimal(safe(row, "Valor dos Serviços"), casas=2)

    if iss_retido == "S":
        # ✅ ISS Retido → alíquota do CSV | Campo 6 = valor ISS | Campo 8 vazio
        aliquota  = fmt_decimal(aliquota_raw, casas=2)
        valor_iss = fmt_decimal(valor_iss_raw, casas=2)
        campo6    = valor_iss
        campo8    = ""
    else:
        # ✅ ISS Normal → alíquota ZERADA | Campo 6 vazio | Campo 8 = valor serviços
        aliquota = "0,00"
        campo6   = ""
        campo8   = valor_serv

    campos = [
        "1020",     # 1
        cod_iss,    # 2  - 18=Retido / 3=Normal
        "",         # 3
        valor_serv, # 4  - Base de cálculo
        aliquota,   # 5  ✅ Alíquota: CSV se Retido | 0,00 se Normal
        campo6,     # 6  ✅ Valor Imposto (só ISS Retido)
        "",         # 7
        campo8,     # 8  ✅ Valor Outras (ISS Normal = valor dos serviços)
        "",         # 9
        "",         # 10
        valor_serv, # 11 - Valor Contábil
        "",         # 12
        "",         # 13
        "",         # 14
        "",         # 15
        "",         # 16
        "",         # 17
        "",         # 18
        "",         # 19
    ]
    assert len(campos) == 19, f"reg_1020: esperado 19 campos, encontrado {len(campos)}"
    return monta_linha(campos)

def reg_1150(_row) -> str:
    return monta_linha(["1150", "", "", "", ""])

def reg_1151(_row) -> str:
    return monta_linha(["1151", "", "", "", ""])

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
        valor_raw  = safe(row, "Valor dos Serviços")
        valor      = fmt_decimal(valor_raw, casas=2)
        valor_iss  = fmt_decimal(safe(row, "Valor ISS"), casas=2)
        iss_retido = safe(row, "ISS Retido").upper().strip()

        # ✅ Ambas as datas do preview também usam Data da Prestação de Serviços
        dt_prestacao = fmt_date(safe(row, "Data da Prestação de Serviços"))
        dt_campo11   = dt_prestacao
        dt_campo12   = dt_prestacao

        if incluir_0020:
            chave_forn = fornecedor if fornecedor else razao
            if chave_forn not in fornecedores_vistos:
                fornecedores_vistos.add(chave_forn)
                linhas.append(reg_0020(row, lookup_pais))

        linhas.append(reg_1000(row, lookup_acum, lookup_pais))

        if incluir_1020:
            linhas.append(reg_1020(row))

        if incluir_1150:
            linhas.append(reg_1150(row))

        if incluir_1151:
            linhas.append(reg_1151(row))

        preview.append({
            "NFTS":           num_nfts,
            "Prestador":      razao,
            "Especie":        especie,
            "CFOP":           cfop,
            "Acumulador":     acumulador,
            "UF":             uf,
            "C11-Dt Entrada": dt_campo11,
            "C12-Dt Emissão": dt_campo12,
            "Valor (raw)":    valor_raw,
            "Valor (dom)":    valor,
            "ISS (dom)":      valor_iss,
            "ISS Retido":     iss_retido,
            "Cód ISS":        cod_iss,
            "Fornecedor":     fornecedor,
        })

    return "\n".join(linhas), pd.DataFrame(preview)

# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────
st.set_page_config(page_title="Conversor NFTS", layout="wide")
st.title("📄 Conversor NFTS → Arquivo de Importação")

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
        "**ISS Retido** — Cód.18 → Campo 6 | Alíquota do CSV\n\n"
        "**ISS Normal** — Cód.3  → Campo 8 | Alíquota = 0,00\n\n"
        "**Fornecedor** — CNPJ (nacional) / vazio (exterior)\n\n"
        "**UF** — EX (exterior) / UF real (nacional)\n\n"
        "**Datas C11/C12** — Data da Prestação de Serviços ✅\n\n"
        "**Decimais** — Domínio: 1.747,85 → 1747,85 ✅"
    )

with st.spinner("🔄 Carregando tabelas de referência do GitHub..."):
    try:
        lookup_acum = carrega_acumuladores_github(URL_ACUMULADORES)
        lookup_pais = carrega_paises_github(URL_PAISES)
        st.sidebar.success(
            f"✅ {len(lookup_acum)} acumuladores\n\n"
            f"✅ {len(lookup_pais)} países\n\n"
            f"_(carregados do GitHub)_"
        )
    except Exception as e:
        st.error(f"❌ Erro ao carregar arquivos do GitHub:\n\n`{e}`")
        st.stop()

st.subheader("📂 Upload do Arquivo NFTS")
file_nfts = st.file_uploader(
    "CSV NFTS — exportado do portal NFTS da Prefeitura de SP",
    type=["csv"],
    key="nfts",
)

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
        st.error(f"Coluna '{col_tipo}' não encontrada.\n\nColunas: `{list(df_nfts.columns)}`")
        st.stop()

    df_notas = df_nfts[df_nfts[col_tipo].str.strip().str.upper() == "4"].copy()

    if df_notas.empty:
        st.warning("Nenhuma nota com Tipo de Registro = 4 encontrada no CSV.")
        st.stop()

    # ✅ Alerta de duplicatas
    df_dup = detecta_duplicatas(df_notas)
    if not df_dup.empty:
        st.error(
            f"⚠️ **{len(df_dup)} nota(s) com número de documento duplicado "
            f"para o mesmo fornecedor!** Verifique antes de importar."
        )
        st.dataframe(
            df_dup.style.apply(
                lambda x: ["background-color: #f8d7da"] * len(x), axis=1
            ),
            use_container_width=True,
        )
        st.markdown("---")

    st.success(
        f"✅ {len(df_notas)} nota(s) | "
        f"{len(lookup_acum)} acumuladores | "
        f"{len(lookup_pais)} países"
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Preview", "📄 Arquivo Gerado", "🔍 Debug", "❓ Ajuda"
    ])

    conteudo, df_prev = converte_nfts(
        df_notas, lookup_acum, lookup_pais, cnpj_empresa,
        incluir_0000, incluir_0020, incluir_1020, incluir_1150, incluir_1151,
    )

    with tab1:
        st.subheader("Preview das Notas")
        if not df_prev.empty:
            n_ret   = len(df_prev[df_prev["ISS Retido"] == "S"])
            n_nor   = len(df_prev[df_prev["ISS Retido"] != "S"])
            n_1933  = len(df_prev[df_prev["CFOP"] == CFOP_SP])
            n_2933  = len(df_prev[df_prev["CFOP"] == CFOP_FORA])
            n_aviso = len(df_prev[df_prev["Acumulador"].str.startswith("AVISO", na=False)])
            n_dup   = len(df_dup) if not df_dup.empty else 0

            m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
            m1.metric("Total notas",          len(df_prev))
            m2.metric("ISS Retido (cód.18)",  n_ret)
            m3.metric("ISS Normal (cód.3)",   n_nor)
            m4.metric("CFOP 1933 (SP)",       n_1933)
            m5.metric("CFOP 2933 (fora/EXT)", n_2933)
            m6.metric("Acum. não mapeado",    n_aviso, delta_color="inverse")
            m7.metric("🔴 Duplicatas",         n_dup,   delta_color="inverse")

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
                "🟢 Verde = ISS Retido (Cód.18) → Campo 6</span>", unsafe_allow_html=True)
            col_l2.markdown(
                "<span style='background:#fff3cd;padding:2px 8px;border-radius:4px'>"
                "🟡 Amarelo = ISS Normal (Cód.3) → Campo 8 | Alíquota 0,00</span>", unsafe_allow_html=True)
            col_l3.markdown(
                "<span style='background:#f8d7da;padding:2px 8px;border-radius:4px'>"
                "🔴 Vermelho = Acumulador não mapeado</span>", unsafe_allow_html=True)

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

    with tab3:
        st.subheader("🔍 Debug — Dados Carregados")

        with st.expander("CSV NFTS — Dados brutos"):
            st.dataframe(df_notas, use_container_width=True)

        with st.expander(f"Acumuladores ({len(lookup_acum)} itens)"):
            st.dataframe(
                pd.DataFrame(list(lookup_acum.items()), columns=["PAULISTANA", "Acumulador"]),
                use_container_width=True,
            )

        with st.expander(f"Países ({len(lookup_pais)} itens)"):
            st.dataframe(
                pd.DataFrame(list(lookup_pais.items()), columns=["Nome", "Código"]),
                use_container_width=True,
            )

        with st.expander("Mapeamento campo a campo por nota"):
            for _, row in df_notas.iterrows():
                nfts      = safe(row, "Nº NFTS")
                val_bruto = safe(row, "Valor dos Serviços")
                iss_bruto = safe(row, "Valor ISS")
                ali_bruto = safe(row, "Alíquota")
                iss_ret   = safe(row, "ISS Retido").upper().strip()
                st.markdown(f"**NFTS {nfts}**")
                debug_data = {
                    "Campo": [
                        "Indicador Prestador",
                        "CPF/CNPJ Prestador",
                        "Razão Social",
                        "UF Prestador",
                        "PAULISTANA",
                        "ISS Retido",
                        "Valor Serviços (bruto CSV)",
                        "Valor Serviços (Domínio)",
                        "Valor ISS (bruto CSV)",
                        "Valor ISS (Domínio)",
                        "Alíquota (bruto CSV)",
                        "Alíquota (1020 campo 5)",
                        "→ Especie",
                        "→ CFOP",
                        "→ Acumulador",
                        "→ Fornecedor",
                        "→ UF",
                        "→ Cód ISS",
                        "→ Cód País",
                        "→ C11 Data entrada",
                        "→ C12 Data emissão",
                    ],
                    "Valor": [
                        safe(row, "Indicador de CPF/CNPJ do Prestador"),
                        safe(row, "CPF/CNPJ do Prestador"),
                        safe(row, "Razão Social do Prestador"),
                        safe(row, "UF do Prestador"),
                        limpa_numero(safe(row, "Código do Serviço Prestado na NFTS")),
                        iss_ret,
                        val_bruto,
                        fmt_decimal(val_bruto, casas=2),
                        iss_bruto,
                        fmt_decimal(iss_bruto, casas=2),
                        ali_bruto,
                        fmt_decimal(ali_bruto, casas=2) if iss_ret == "S" else "0,00 (forçado)",
                        determina_especie(row),
                        determina_cfop(row),
                        determina_acumulador(row, lookup_acum),
                        determina_fornecedor(row),
                        determina_uf(row),
                        determina_cod_iss(row),
                        determina_cod_pais(row, lookup_pais),
                        fmt_date(safe(row, "Data da Prestação de Serviços")),  # ✅
                        fmt_date(safe(row, "Data da Prestação de Serviços")),  # ✅
                    ],
                }
                st.dataframe(
                    pd.DataFrame(debug_data),
                    use_container_width=True,
                    hide_index=True,
                )
                st.markdown("---")

    with tab4:
        st.subheader("❓ Ajuda — Regras de Importação")

        st.markdown("### Regras automáticas aplicadas")
        st.dataframe(pd.DataFrame([
            ["Indicador=3 (Exterior)", "39", "2933", "2551",               "EX",      "76 (EUA)"],
            ["Indicador=1/2, UF=SP",   "39", "1933", "Lookup PAULISTANA", "SP",       ""],
            ["Indicador=1/2, UF≠SP",   "39", "2933", "Lookup PAULISTANA", "UF real",  ""],
        ], columns=["Situação", "Espécie", "CFOP", "Acumulador", "UF", "Cód País"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Formato de decimais — Padrão Domínio ✅")
        st.dataframe(pd.DataFrame([
            ["1.747,85", "ponto=milhar, vírgula=decimal", "1747,85 ✅"],
            ["963,02",   "vírgula=decimal",               "963,02  ✅"],
            ["2,9",      "alíquota ISS Retido",           "2,90    ✅"],
            ["0",        "alíquota ISS Normal",           "0,00    ✅"],
        ], columns=["Valor no CSV", "Interpretação", "Resultado no arquivo"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Datas — Registro 1000 ✅")
        st.dataframe(pd.DataFrame([
            ["Campo 11", "Data da entrada", "Data da Prestação de Serviços", "✅"],
            ["Campo 12", "Data emissão",    "Data da Prestação de Serviços", "✅"],
        ], columns=["Campo", "Nome Domínio", "Fonte CSV", "Status"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Regra ISS — Registro 1020 ✅")
        st.dataframe(pd.DataFrame([
            ["ISS Retido (S)", "Cód.18", "do CSV",  "valor_iss",  '""',       "Campo 6 = Valor Imposto"],
            ["ISS Normal (N)", "Cód. 3", "0,00 ✅", '""',          "valor_serv","Campo 8 = Valor Outras"],
        ], columns=["Situação", "Cód ISS", "Alíquota C5", "Campo 6", "Campo 8", "Observação"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Todas as correções aplicadas")
        st.dataframe(pd.DataFrame([
            ["fmt_decimal", "—",  "Formato decimal",    "×100 inteiro",  "vírgula decimal ✅"],
            ["0020",        "4",  "Nome reduzido",       '""',            "razao[:40] ✅"],
            ["0020",        "22", "Agropecuário",        '""',            '"N" ✅'],
            ["0020",        "24", "Regime de apuração",  '""',            '"N" ✅'],
            ["0020",        "25", "Contribuinte ICMS",   '""',            '"N" ✅'],
            ["0020",        "31", "Contribuinte CPRB",   '""',            '"N" ✅'],
            ["1000",        "11", "Data entrada",        "Emissão NFTS",  "Prestação Serviços ✅"],
            ["1000",        "12", "Data emissão",        "Prestação Serv","Prestação Serviços ✅"],
            ["1000",        "39", "Valor produtos",      '""',            "= campo 13 ✅"],
            ["1020",        "5",  "Alíquota ISS Normal", "do CSV",        "0,00 ✅"],
            ["1020",        "6",  "Valor Imposto",       "sempre",        "só ISS Retido ✅"],
            ["1020",        "8",  "Valor Outras",        '""',            "ISS Normal = valor_serv ✅"],
            ["1020",        "—",  "Geração",             "condicional",   "SEMPRE gerado ✅"],
            ["—",           "—",  "Duplicatas",          "—",             "Alerta 🔴 ✅"],
        ], columns=["Onde", "Campo", "Nome", "Antes", "Depois"]),
        use_container_width=True, hide_index=True)

else:
    st.info(
        "👆 Faça upload do **CSV NFTS** para iniciar.\n\n"
        "Os arquivos **Acumuladores.xlsx** e **Países.xlsx** são carregados "
        "automaticamente do GitHub."
    )

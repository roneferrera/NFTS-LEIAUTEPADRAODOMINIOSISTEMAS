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
    """
    Converte valor para formato Domínio Sistemas (Decimal com N casas).
    Regra oficial: remover separadores e manter dígitos × 10^casas.
    Exemplos:
      1.747,85  → 174785   (ponto=milhar, vírgula=decimal)
      1747.85   → 174785   (ponto=decimal)
      1747,85   → 174785   (vírgula=decimal)
      2,90      → 290      (alíquota)
      50,68     → 5068     (ISS)
    """
    try:
        s = str(value).strip()
        if not s or s in ("nan", ""):
            return "0"
        # ponto E vírgula → ponto=milhar, vírgula=decimal
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        # só vírgula → vírgula=decimal
        elif "," in s:
            s = s.replace(",", ".")
        # só ponto → ponto=decimal (não mexe)
        v = float(s)
        return str(int(round(v * (10 ** casas))))
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
    """
    Une campos com pipe E adiciona pipe final.
    Regra: todo pipe que abre tem que fechar.
    N campos = N-1 pipes internos + 1 pipe final = N pipes total.
    """
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
# MONTA REGISTROS — conforme leiautes oficiais
#
# 0000 → 2 campos  → 2 pipes  (1 interno + 1 final)
# 0020 → 33 campos → 33 pipes (32 internos + 1 final)
# 1000 → 98 campos → 98 pipes (97 internos + 1 final)
# 1020 → 19 campos → 19 pipes (18 internos + 1 final)
#
# Decimais: formato Domínio = inteiro sem separador
#   1747,85 → 174785  |  2,90 → 290  |  50,68 → 5068
# ─────────────────────────────────────────────

def reg_0000(cnpj_empresa: str) -> str:
    # Leiaute: 2 campos
    # Campo 1: Fixo "0000"
    # Campo 2: CNPJ/CPF/CEI/CAEPF da empresa (só números)
    return monta_linha([
        "0000",        # 1 - Identificação do registro
        cnpj_empresa,  # 2 - Inscrição da empresa
    ])
    # resultado: 0000|20586841000130|

def reg_0020(row, lookup_pais: dict) -> str:
    # Leiaute: 33 campos (verificado campo a campo)
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
        "0020",      # 1  - Identificação do registro (fixo)
        fornecedor,  # 2  - Inscrição CNPJ/CPF/CEI/CAEPF (só números)
        razao,       # 3  - Razão Social (max 150 chars)
        "",          # 4  - Apelido (max 40 chars)
        endereco,    # 5  - Endereço
        numero_end,  # 6  - Número do endereço (numérico)
        complemento, # 7  - Complemento
        bairro,      # 8  - Bairro
        "",          # 9  - Código do município (numérico: estadual/federal/IBGE/RAIS)
        uf,          # 10 - UF (EX para exterior)
        cod_pais,    # 11 - Código do País (numérico, só exterior)
        cep,         # 12 - CEP
        "",          # 13 - Inscrição Estadual
        "",          # 14 - Inscrição Municipal
        "",          # 15 - Inscrição Suframa
        "",          # 16 - DDD
        "",          # 17 - Telefone
        "",          # 18 - FAX
        "",          # 19 - Data do cadastro (dd/mm/aaaa)
        "",          # 20 - Conta contábil (numérico)
        "",          # 21 - Conta contábil cliente (numérico)
        "",          # 22 - Agropecuário (S/N)
        "",          # 23 - Natureza jurídica (1-8)
        "",          # 24 - Regime de apuração (N/M/E/O/U/I)
        "",          # 25 - Contribuinte ICMS (S/N)
        "",          # 26 - Alíquota ICMS (decimal)
        "",          # 27 - Categoria do estabelecimento (ARM/CNF/CPQ/DIS/...)
        "",          # 28 - Inscrição Estadual ST
        email,       # 29 - Email
        "",          # 30 - Interdependência com a empresa (S/N)
        "",          # 31 - Contribuinte da CPRB (S/N)
        "",          # 32 - Processo administrativo/judicial (max 21 chars)
        "",          # 33 - Tipo Inscrição (1=CAEPF)
    ]
    # Verifica contagem: deve ser exatamente 33
    assert len(campos) == 33, f"reg_0020: esperado 33 campos, encontrado {len(campos)}"
    return monta_linha(campos)

def reg_1000(row, lookup_acum: dict, lookup_pais: dict) -> str:
    # Leiaute: 98 campos (verificado campo a campo)
    especie    = determina_especie(row)
    fornecedor = determina_fornecedor(row)
    acumulador = determina_acumulador(row, lookup_acum)
    cfop       = determina_cfop(row)
    num_doc    = limpa_numero(safe(row, "Número do Documento"))
    serie      = safe(row, "Série do Documento")
    if serie in ("-", "nan", ""):
        serie = ""

    # DATAS:
    # Campo 11 = Data da entrada → recebe Data Hora Emissão NFTS (mais antiga)
    # Campo 12 = Data emissão    → recebe Data da Prestação de Serviços (mais recente)
    # Regra Domínio: campo 11 (entrada) deve ser <= campo 12 (emissão)
    dt_campo11 = fmt_date(safe(row, "Data Hora Emissão NFTS"))
    dt_campo12 = fmt_date(safe(row, "Data da Prestação de Serviços"))

    # VALOR: Decimal(2) → inteiro sem separador
    # Ex.: 1747,85 → 174785
    valor   = fmt_decimal(safe(row, "Valor dos Serviços"), casas=2)
    cod_iss = determina_cod_iss(row)

    campos = [
        "1000",      # 1  - Identificação do registro (fixo)
        especie,     # 2  - Código da espécie (numérico)
        fornecedor,  # 3  - Inscrição fornecedor (CNPJ/CPF/CEI/CAEPF)
        "",          # 4  - Código de Exclusão da DIEF (numérico)
        acumulador,  # 5  - Código do acumulador (numérico)
        cfop,        # 6  - CFOP (numérico)
        "",          # 7  - Segmento (numérico)
        num_doc,     # 8  - Número do documento (numérico)
        serie,       # 9  - Série (caractere)
        "",          # 10 - Numero do documento final (numérico)
        dt_campo11,  # 11 - Data da entrada (dd/mm/aaaa) ← Emissão NFTS
        dt_campo12,  # 12 - Data emissão (dd/mm/aaaa)   ← Prestação Serviço
        valor,       # 13 - Valor contábil (decimal 2) → ex: 174785
        "",          # 14 - Valor da exclusão da DIEF (decimal 2)
        "",          # 15 - Observação (caractere)
        "",          # 16 - Modalidade do frete (C/F/S/T/R/D)
        "",          # 17 - Emitente da nota fiscal (P/T)
        "",          # 18 - CFOP estendido/detalhamento (numérico, só SE)
        "",          # 19 - Código da transferência de crédito (numérico, só RS)
        cod_iss,     # 20 - Código do Recolhimento do ISS Retido (caractere)
        "",          # 21 - Código do Recolhimento do IRRF (caractere)
        "",          # 22 - Código da observação (numérico)
        "",          # 23 - Data do visto (dd/mm/aaaa, só MG)
        "",          # 24 - Fato gerador da CRF (E/P)
        "",          # 25 - Fato gerador do IRRF (E/P)
        "",          # 26 - Valor do frete (decimal 2)
        "",          # 27 - Valor do seguro (decimal 2)
        "",          # 28 - Valor das despesas (decimal 2)
        "",          # 29 - Valor do PIS (decimal 2)
        "",          # 30 - Código Antecipação Tributária (numérico)
        "",          # 31 - Valor do COFINS (decimal 2)
        "",          # 32 - Valor DARE (decimal 2, só SE)
        "",          # 33 - Alíquota DARE (decimal 2, só SE)
        "",          # 34 - Valor base ICMS ST (numérico: 0/1/2)
        "",          # 35 - Entradas cuja saída é isenta (decimal 2, só MG)
        "",          # 36 - Outras entradas isentas (decimal 2, só MG)
        "",          # 37 - Valor transporte incluído na base (decimal 2, só MG)
        "",          # 38 - Código de ressarcimento (numérico)
        "",          # 39 - Valor produtos (decimal 2)
        "",          # 40 - Município Origem (numérico)
        "",          # 41 - Situação da Nota (0-10)
        "",          # 42 - Código da situação tributária (numérico)
        "",          # 43 - Sub serie (numérico)
        "",          # 44 - Inscrição estadual do fornecedor (caractere)
        "",          # 45 - Inscrição municipal do fornecedor (caractere)
        "",          # 46 - Código da operação e prestação (caractere)
        "",          # 47 - Valor a ser deduzido da receita tributável (decimal 2)
        "",          # 48 - Competência (dd/mm/aaaa)
        "",          # 49 - Operação (numérico, só PA)
        "",          # 50 - Número do parecer fiscal (caractere)
        "",          # 51 - Data do parecer fiscal (dd/mm/aaaa)
        "",          # 52 - Número da declaração de Importação (caractere)
        "",          # 53 - Possui benefício fiscal (S/N)
        "",          # 54 - Chave da nota fiscal eletrônica (caractere)
        "",          # 55 - Código de recolhimento do FETHAB (caractere)
        "",          # 56 - Responsável pelo recolhimento do FETHAB (E/C)
        "",          # 57 - CFOP documento fiscal (numérico)
        "",          # 58 - Tipo de CT-e (0/1/2)
        "",          # 59 - CT-e referência (caractere)
        "",          # 60 - Modalidade da importação (1-5)
        "",          # 61 - Código da informação complementar (numérico)
        "",          # 62 - Informação complementar (caractere)
        "",          # 63 - Classe de consumo (numérico)
        "",          # 64 - Tipo de ligação (numérico)
        "",          # 65 - Grupo de tensão (numérico)
        "",          # 66 - Tipo de assinante (numérico)
        "",          # 67 - KWH consumido (numérico)
        "",          # 68 - Valor fornecido/consumido (decimal 2)
        "",          # 69 - Valor cobrado de terceiros (decimal 2)
        "",          # 70 - Tipo do documento de importação (10/1)
        "",          # 71 - Número do Ato Concessório Drawback (caractere)
        "",          # 72 - Natureza do frete PIS/COFINS (numérico)
        "",          # 73 - CST PIS/COFINS (numérico)
        "",          # 74 - Base do crédito PIS/COFINS (numérico)
        "",          # 75 - Valor serviços/itens PIS/COFINS (decimal 2)
        "",          # 76 - Base de cálculo PIS/COFINS (decimal 2)
        "",          # 77 - Alíquota de PIS (decimal 2)
        "",          # 78 - Alíquota de COFINS (decimal 2)
        "",          # 79 - Chave de NFSe (caractere)
        "",          # 80 - Número do processo ou ato concessório (caractere)
        "",          # 81 - Origem do processo (1/3/9)
        "",          # 82 - Data da escrituração (dd/mm/aaaa)
        "",          # 83 - CFPS (numérico, só DF)
        "",          # 84 - Natureza da receita PIS/COFINS (numérico)
        "",          # 85 - CST IPI (caractere)
        "",          # 86 - Lançamentos de SCP (numérico)
        "",          # 87 - Tipo de serviço (1/2)
        "",          # 88 - Município destino (numérico)
        "",          # 89 - Pedágio (decimal 2)
        "",          # 90 - IPI (decimal 2)
        "",          # 91 - ICMS ST (decimal 2)
        "",          # 92 - Classificação EFD-Reinf tipo (numérico)
        "",          # 93 - Classificação EFD-Reinf indicativo (0/1/2)
        "",          # 94 - Número do documento de arrecadação (max 255, só RS)
        "",          # 95 - Tipo do título (0-3/99)
        "",          # 96 - Identificação (max 255 chars)
        "",          # 97 - ICMS Desonerado (decimal 2)
        "",          # 98 - IPI Devolução (decimal 2)
    ]
    assert len(campos) == 98, f"reg_1000: esperado 98 campos, encontrado {len(campos)}"
    return monta_linha(campos)

def reg_1020(row) -> str:
    # Leiaute: 19 campos (verificado campo a campo)
    # Decimais: formato Domínio → inteiro sem separador
    # Alíquota campo 5: Decimal(2) para ISS (não é IRRFP que usa 3 casas)
    iss_retido    = safe(row, "ISS Retido").upper().strip()
    valor_iss_raw = safe(row, "Valor ISS")
    aliquota_raw  = safe(row, "Alíquota")
    cod_iss       = determina_cod_iss(row)

    valor_serv = fmt_decimal(safe(row, "Valor dos Serviços"), casas=2)
    valor_iss  = fmt_decimal(valor_iss_raw, casas=2)
    aliquota   = fmt_decimal(aliquota_raw, casas=2)

    # Verifica se há ISS a lançar
    try:
        s = str(valor_iss_raw).strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        v_iss = float(s)
    except Exception:
        v_iss = 0.0

    if v_iss == 0.0 and iss_retido != "S":
        return ""

    campos = [
        "1020",     # 1  - Identificação do registro (fixo)
        cod_iss,    # 2  - Código do imposto (numérico)
        "",         # 3  - Percentual de redução da base de cálculo (decimal 2)
        valor_serv, # 4  - Base de cálculo (decimal 2) → ex: 174785
        aliquota,   # 5  - Alíquota (decimal 2 para ISS) → ex: 290 = 2,90%
        valor_iss,  # 6  - Valor do Imposto (decimal 2) → ex: 5068
        "",         # 7  - Valor de Isentas (decimal 2)
        "",         # 8  - Valor de Outras (decimal 2)
        "",         # 9  - Valor do IPI (decimal 2)
        "",         # 10 - Valor da substituição Tributária (decimal 2)
        valor_serv, # 11 - Valor Contábil (decimal 2)
        "",         # 12 - Código do recolhimento do imposto (caractere)
        "",         # 13 - Valor não tributadas (decimal 2, só GO)
        "",         # 14 - Valor parcela reduzida (decimal 2, só GO)
        "",         # 15 - Alíq. Interest. (numérico/decimal, só RS/SP)
        "",         # 16 - Nat. rend. (numérico max 5, só impostos específicos)
        "",         # 17 - Tipo de Dedução (numérico max 1, só 63-IRRF-APF)
        "",         # 18 - Tipo de Isenção (numérico max 2, só 63-IRRF-APF)
        "",         # 19 - Descrição (caractere max 100)
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
        dt_campo11 = fmt_date(safe(row, "Data Hora Emissão NFTS"))
        dt_campo12 = fmt_date(safe(row, "Data da Prestação de Serviços"))

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
        "**ISS** — Cód.18 (retido) / Cód.3 (normal)\n\n"
        "**Fornecedor** — CNPJ (nacional) / vazio (exterior)\n\n"
        "**UF** — EX (exterior) / UF real (nacional)\n\n"
        "**Decimais** — Domínio: 1747,85 → 174785"
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
                        "Alíquota (Domínio)",
                        "→ Especie",
                        "→ CFOP",
                        "→ Acumulador",
                        "→ Fornecedor",
                        "→ UF",
                        "→ Cód ISS",
                        "→ Cód País",
                        "→ C11 Data entrada (Emissão NFTS)",
                        "→ C12 Data emissão (Prestação Serviço)",
                    ],
                    "Valor": [
                        safe(row, "Indicador de CPF/CNPJ do Prestador"),
                        safe(row, "CPF/CNPJ do Prestador"),
                        safe(row, "Razão Social do Prestador"),
                        safe(row, "UF do Prestador"),
                        limpa_numero(safe(row, "Código do Serviço Prestado na NFTS")),
                        safe(row, "ISS Retido"),
                        val_bruto,
                        fmt_decimal(val_bruto, casas=2),
                        iss_bruto,
                        fmt_decimal(iss_bruto, casas=2),
                        ali_bruto,
                        fmt_decimal(ali_bruto, casas=2),
                        determina_especie(row),
                        determina_cfop(row),
                        determina_acumulador(row, lookup_acum),
                        determina_fornecedor(row),
                        determina_uf(row),
                        determina_cod_iss(row),
                        determina_cod_pais(row, lookup_pais),
                        fmt_date(safe(row, "Data Hora Emissão NFTS")),
                        fmt_date(safe(row, "Data da Prestação de Serviços")),
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
            ["Indicador=3 (Exterior)", "39", "2933", "2551", "EX",      "76 (EUA)"],
            ["Indicador=1/2, UF=SP",   "39", "1933", "Lookup PAULISTANA", "SP",   ""],
            ["Indicador=1/2, UF≠SP",   "39", "2933", "Lookup PAULISTANA", "UF real", ""],
        ], columns=["Situação", "Espécie", "CFOP", "Acumulador", "UF", "Cód País"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Formato de decimais — Regra oficial Domínio")
        st.dataframe(pd.DataFrame([
            ["1.747,85", "ponto=milhar, vírgula=decimal", "174785"],
            ["1747.85",  "ponto=decimal",                 "174785"],
            ["1747,85",  "vírgula=decimal",               "174785"],
            ["2,90",     "alíquota ISS",                  "290"],
            ["50,68",    "valor ISS",                     "5068"],
        ], columns=["Valor no CSV", "Interpretação", "Resultado no arquivo"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Lógica das datas no Registro 1000")
        st.dataframe(pd.DataFrame([
            ["Campo 11", "Data da entrada", "Data Hora Emissão NFTS",        "Mais antiga ≤ campo 12"],
            ["Campo 12", "Data emissão",    "Data da Prestação de Serviços", "Mais recente ≥ campo 11"],
        ], columns=["Campo", "Nome Domínio", "Fonte no CSV", "Regra"]),
        use_container_width=True, hide_index=True)

        st.markdown("### Campos principais do Registro 1000")
        st.dataframe(pd.DataFrame([
            ["Campo 2",  "Espécie",     "Todas as notas",                     "39"],
            ["Campo 3",  "Fornecedor",  "Nacional = CNPJ / Exterior = vazio", "—"],
            ["Campo 5",  "Acumulador",  "Exterior = 2551 / Nacional = lookup","—"],
            ["Campo 6",  "CFOP",        "SP = 1933 / Outros/EXT = 2933",      "—"],
            ["Campo 11", "Dt entrada",  "Data Hora Emissão NFTS",             "dd/mm/aaaa"],
            ["Campo 12", "Dt emissão",  "Data da Prestação de Serviços",      "dd/mm/aaaa"],
            ["Campo 13", "Valor",       "Valor dos Serviços",                 "174785 = 1747,85"],
            ["Campo 20", "Cód ISS",     "ISS Retido S=18 / N=3",             "—"],
        ], columns=["Campo", "Nome", "Regra", "Exemplo"]),
        use_container_width=True, hide_index=True)

else:
    st.info(
        "👆 Faça upload do **CSV NFTS** para iniciar.\n\n"
        "Os arquivos **Acumuladores.xlsx** e **Países.xlsx** são carregados "
        "automaticamente do GitHub."
    )

"""
app.py — Conversor NFTS Paulistana → Domínio Sistemas
Versão final — tudo em um único arquivo
"""

import csv
import io
import re
import time
import streamlit as st
import pandas as pd
from datetime import datetime
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

st.set_page_config(
    page_title="NFTS → Domínio Sistemas",
    page_icon="🧾",
    layout="wide"
)

# ════════════════════════════════════════════════════════════════════════════
# TABELA OFICIAL DE PAÍSES – DOMÍNIO SISTEMAS
# Fonte: Países.xls (arquivo oficial Domínio)
# Estrutura: ISO Alpha-2 → (Nome no Domínio, Código interno Domínio)
# ════════════════════════════════════════════════════════════════════════════
TABELA_PAISES: dict[str, tuple[str, str]] = {
    "AF": ("AFEGANISTAO",                                        "1"),
    "ZA": ("AFRICA DO SUL",                                      "2"),
    "AL": ("ALBANIA, REPUBLICA DA",                              "3"),
    "DE": ("ALEMANHA",                                           "4"),
    "AD": ("ANDORRA",                                            "5"),
    "AO": ("ANGOLA",                                             "6"),
    "AI": ("ANGUILLA",                                           "7"),
    "AG": ("ANTIGUA E BARBUDA",                                  "8"),
    "AN": ("ANTILHAS HOLANDESAS",                                "9"),
    "SA": ("ARABIA SAUDITA",                                     "10"),
    "DZ": ("ARGELIA",                                            "11"),
    "AR": ("ARGENTINA",                                          "12"),
    "AM": ("ARMENIA, REPUBLICA DA",                              "13"),
    "AW": ("ARUBA",                                              "14"),
    "AU": ("AUSTRALIA",                                          "15"),
    "AT": ("AUSTRIA",                                            "16"),
    "AZ": ("AZERBAIJAO, REPUBLICA DO",                           "17"),
    "BS": ("BAHAMAS, ILHAS",                                     "18"),
    "BH": ("BAHREIN, ILHAS",                                     "19"),
    "BD": ("BANGLADESH",                                         "20"),
    "BB": ("BARBADOS",                                           "21"),
    "BY": ("Belarus, República da",                              "22"),
    "BE": ("BELGICA",                                            "23"),
    "BZ": ("BELIZE",                                             "24"),
    "BJ": ("BENIN",                                              "25"),
    "BM": ("BERMUDAS",                                           "26"),
    "BO": ("Bolívia, Estado Plurinacional da",                   "27"),
    "BA": ("Bósnia-Herzegovina, República da",                   "28"),
    "BW": ("BOTSUANA",                                           "29"),
    "BR": ("BRASIL",                                             "30"),
    "BN": ("BRUNEI",                                             "31"),
    "BG": ("BULGARIA, REPUBLICA DA",                             "32"),
    "BF": ("BURKINA FASO",                                       "33"),
    "BI": ("BURUNDI",                                            "34"),
    "BT": ("BUTAO",                                              "35"),
    "CV": ("CABO VERDE, REPUBLICA DE",                           "36"),
    "CM": ("CAMAROES",                                           "37"),
    "KH": ("CAMBOJA",                                            "38"),
    "CA": ("CANADA",                                             "39"),
    "GG": ("GUERNSEY, ILHA DO CANAL (INCLUI ALDERNEY E SARK)",   "40"),
    "IC": ("CANARIAS, ILHAS",                                    "41"),
    "QA": ("CATAR",                                              "42"),
    "KY": ("CAYMAN, ILHA",                                       "43"),
    "KZ": ("CAZAQUISTAO, REPUBLICA DO",                          "44"),
    "TD": ("CHADE",                                              "45"),
    "CL": ("CHILE",                                              "46"),
    "CN": ("CHINA, REPUBLICA POPULAR DA",                        "47"),
    "CY": ("CHIPRE",                                             "48"),
    "CX": ("CHRISTMAS, ILHA (NAVIDAD)",                          "49"),
    "SG": ("Singapura",                                          "50"),
    "CC": ("COCOS (KEELING), ILHAS",                             "51"),
    "CO": ("COLOMBIA",                                           "52"),
    "KM": ("COMORES, ILHAS",                                     "53"),
    "CD": ("CONGO, REPUBLICA DEMOCRATICA DO",                    "54"),
    "CG": ("CONGO, REPUBLICA DO",                                "55"),
    "CK": ("COOK, ILHA",                                         "56"),
    "KP": ("Coreia (do Norte), Rep. Pop. Democrática da",        "57"),
    "KR": ("Coreia (do Sul), República da",                      "58"),
    "CI": ("COSTA DO MARFIM",                                    "59"),
    "CR": ("COSTA RICA",                                         "60"),
    "KW": ("KUWAIT",                                             "61"),
    "HR": ("CROACIA, REPUBLICA DA",                              "62"),
    "CU": ("CUBA",                                               "63"),
    "DK": ("DINAMARCA",                                          "64"),
    "DJ": ("DJIBUTI",                                            "65"),
    "DM": ("DOMINICA, ILHA",                                     "66"),
    "EG": ("EGITO",                                              "67"),
    "SV": ("EL SALVADOR",                                        "68"),
    "AE": ("EMIRADOS ARABES UNIDOS",                             "69"),
    "EC": ("EQUADOR",                                            "70"),
    "ER": ("ERITREIA",                                           "71"),
    "GB-SCT": ("ESCOCIA",                                        "72"),
    "SK": ("ESLOVACA, REPUBLICA",                                "73"),
    "SI": ("ESLOVENIA, REPUBLICA DA",                            "74"),
    "ES": ("ESPANHA",                                            "75"),
    "US": ("ESTADOS UNIDOS",                                     "76"),
    "EE": ("ESTONIA, REPUBLICA DA",                              "77"),
    "ET": ("ETIOPIA",                                            "78"),
    "FK": ("FALKLAND (ILHAS MALVINAS)",                          "79"),
    "FO": ("FEROE, ILHAS",                                       "80"),
    "FJ": ("FIJI",                                               "81"),
    "PH": ("FILIPINAS",                                          "82"),
    "FI": ("FINLANDIA",                                          "83"),
    "TW": ("FORMOSA (TAIWAN)",                                   "84"),
    "FR": ("FRANCA",                                             "85"),
    "GA": ("GABAO",                                              "86"),
    "GB-WLS": ("GALES, PAIS DE",                                 "87"),
    "GM": ("GAMBIA",                                             "88"),
    "GH": ("GANA",                                               "89"),
    "GE": ("GEORGIA, REPUBLICA DA",                              "90"),
    "GI": ("GIBRALTAR",                                          "91"),
    "GB": ("GRA-BRETANHA",                                       "92"),
    "GD": ("GRANADA",                                            "93"),
    "GR": ("GRECIA",                                             "94"),
    "GL": ("GROENLANDIA",                                        "95"),
    "GP": ("GUADALUPE",                                          "96"),
    "GU": ("GUAM",                                               "97"),
    "GT": ("GUATEMALA",                                          "98"),
    "GY": ("GUIANA",                                             "99"),
    "GF": ("GUIANA FRANCESA",                                    "100"),
    "GN": ("GUINE",                                              "101"),
    "GW": ("GUINE-BISSAU",                                       "102"),
    "GQ": ("GUINE-EQUATORIAL",                                   "103"),
    "HT": ("HAITI",                                              "104"),
    "NL": ("Países Baixos (Holanda)",                            "105"),
    "HN": ("HONDURAS",                                           "106"),
    "HK": ("HONG KONG, REGIAO ADM. ESPECIAL",                    "107"),
    "HU": ("HUNGRIA, REPUBLICA DA",                              "108"),
    "YE": ("IEMEN",                                              "109"),
    "IN": ("INDIA",                                              "110"),
    "ID": ("INDONESIA",                                          "111"),
    "GB-ENG": ("INGLATERRA",                                     "112"),
    "IR": ("IRA, REPUBLICA ISLAMICA DO",                         "113"),
    "IQ": ("IRAQUE",                                             "114"),
    "IE": ("IRLANDA",                                            "115"),
    "GB-NIR": ("IRLANDA DO NORTE",                               "116"),
    "IS": ("ISLANDIA",                                           "117"),
    "IL": ("ISRAEL",                                             "118"),
    "IT": ("ITALIA",                                             "119"),
    "RS": ("SERVIA",                                             "120"),
    "JM": ("JAMAICA",                                            "121"),
    "JP": ("JAPAO",                                              "122"),
    "UM-67": ("JOHNSTON, ILHAS",                                 "123"),
    "JO": ("JORDANIA",                                           "124"),
    "KI": ("KIRIBATI",                                           "125"),
    "LA": ("LAOS, REP. POP. DEMOCRATICA DO",                     "126"),
    "BN-LB": ("LEBUAN",                                          "127"),
    "LS": ("LESOTO",                                             "128"),
    "LV": ("LETONIA, REPUBLICA DA",                              "129"),
    "LB": ("LIBANO",                                             "130"),
    "LR": ("LIBERIA",                                            "131"),
    "LY": ("LIBIA",                                              "132"),
    "LI": ("LIECHTENSTEIN",                                      "133"),
    "LT": ("LITUANIA, REPUBLICA DA",                             "134"),
    "LU": ("LUXEMBURGO",                                         "135"),
    "MO": ("MACAU",                                              "136"),
    "MK": ("MACEDONIA DO NORTE",                                 "137"),
    "MG": ("MADAGASCAR",                                         "138"),
    "PT-30": ("MADEIRA, ILHA DA",                                "139"),
    "MY": ("MALASIA",                                            "140"),
    "MW": ("MALAVI",                                             "141"),
    "MV": ("MALDIVAS",                                           "142"),
    "ML": ("MALI",                                               "143"),
    "MT": ("MALTA",                                              "144"),
    "IM": ("MAN, ILHAS",                                         "145"),
    "MP": ("MARIANAS DO NORTE",                                  "146"),
    "MA": ("MARROCOS",                                           "147"),
    "MH": ("MARSHALL, ILHAS",                                    "148"),
    "MQ": ("MARTINICA",                                          "149"),
    "MU": ("MAURICIO",                                           "150"),
    "MR": ("MAURITANIA",                                         "151"),
    "MX": ("MEXICO",                                             "152"),
    "MM": ("MIANMAR (BIRMANIA)",                                 "153"),
    "FM": ("MICRONESIA",                                         "154"),
    "UM-71": ("MIDWAY, ILHAS",                                   "155"),
    "MZ": ("MOCAMBIQUE",                                         "156"),
    "MD": ("MOLDAVIA, REPUBLICA DA",                             "157"),
    "MC": ("MONACO",                                             "158"),
    "MN": ("MONGOLIA",                                           "159"),
    "MS": ("MONTSERRAT, ILHA",                                   "160"),
    "NA": ("NAMIBIA",                                            "161"),
    "NR": ("NAURU",                                              "162"),
    "NP": ("NEPAL",                                              "163"),
    "NI": ("NICARAGUA",                                          "164"),
    "NE": ("NIGER",                                              "165"),
    "NG": ("NIGERIA",                                            "166"),
    "NU": ("NIUE, ILHA",                                         "167"),
    "NF": ("NORFOLK, ILHA",                                      "168"),
    "NO": ("NORUEGA",                                            "169"),
    "NC": ("NOVA CALEDONIA",                                     "170"),
    "NZ": ("NOVA ZELANDIA",                                      "171"),
    "OM": ("OMA",                                                "172"),
    "PW": ("PALAU",                                              "173"),
    "PA": ("PANAMA",                                             "174"),
    "PG": ("PAPUA NOVA GUINE",                                   "175"),
    "PK": ("PAQUISTAO",                                          "176"),
    "PY": ("PARAGUAI",                                           "177"),
    "PE": ("PERU",                                               "178"),
    "PN": ("PITCAIRN, ILHA",                                     "179"),
    "PF": ("POLINESIA FRANCESA",                                 "180"),
    "PL": ("POLONIA, REPUBLICA DA",                              "181"),
    "PR": ("PORTO RICO",                                         "182"),
    "PT": ("PORTUGAL",                                           "183"),
    "KE": ("QUENIA",                                             "184"),
    "KG": ("QUIRGUIZ, REPUBLICA",                                "185"),
    "UK": ("REINO UNIDO",                                        "186"),
    "CF": ("REPUBLICA CENTRO-AFRICANA",                          "187"),
    "DO": ("REPUBLICA DOMINICANA",                               "188"),
    "RE": ("REUNIAO, ILHA",                                      "189"),
    "RO": ("ROMENIA",                                            "190"),
    "RW": ("RUANDA",                                             "191"),
    "RU": ("Rússia, Federação da",                               "192"),
    "EH": ("SAARA OCIDENTAL",                                    "193"),
    "SB": ("SALOMAO, ILHAS",                                     "194"),
    "WS": ("SAMOA",                                              "195"),
    "AS": ("SAMOA AMERICANA",                                    "196"),
    "SM": ("San Marino",                                         "197"),
    "SH": ("SANTA HELENA",                                       "198"),
    "LC": ("SANTA LUCIA",                                        "199"),
    "KN": ("SAO CRISTOVAO E NEVES",                              "200"),
    "PM": ("SAO PEDRO E MIQUELON",                               "201"),
    "ST": ("SAO TOME E PRINCIPE, ILHAS",                         "202"),
    "VC": ("SAO VICENTE E GRANADINA",                            "203"),
    "SN": ("SENEGAL",                                            "204"),
    "SL": ("SERRA LEOA",                                         "205"),
    "SC": ("SEYCHELLE",                                          "206"),
    "SY": ("SIRIA, REPUBLICA ARABE DA",                          "207"),
    "SO": ("SOMALIA",                                            "208"),
    "LK": ("SRI LANKA",                                          "209"),
    "SZ": ("eSwatini (Essuatíni, Suazilândia)",                  "210"),
    "SD": ("SUDAO",                                              "211"),
    "SE": ("SUECIA",                                             "212"),
    "CH": ("SUICA",                                              "213"),
    "SR": ("SURINAME",                                           "214"),
    "TJ": ("TADJIQUISTAO",                                       "215"),
    "TH": ("TAILANDIA",                                          "216"),
    "TZ": ("TANZANIA, REPUBLICA UNIDA DA",                       "217"),
    "CZ": ("TCHECA, REPUBLICA",                                  "218"),
    "IO": ("TERRITORIO BRITANICO OC. INDICO",                    "219"),
    "TL": ("TIMOR LESTE",                                        "220"),
    "TG": ("TOGO",                                               "221"),
    "TO": ("TONGA",                                              "222"),
    "TK": ("TOQUELAU, ILHAS",                                    "223"),
    "TT": ("TRINIDAD E TOBAGO",                                  "224"),
    "TN": ("TUNISIA",                                            "225"),
    "TC": ("TURCAS E CAICOS, ILHAS",                             "226"),
    "TM": ("TURCOMENISTAO, REPUBLICA DO",                        "227"),
    "TR": ("TURQUIA",                                            "228"),
    "TV": ("TUVALU",                                             "229"),
    "UA": ("UCRANIA",                                            "230"),
    "UG": ("UGANDA",                                             "231"),
    "UY": ("URUGUAI",                                            "232"),
    "UZ": ("UZBEQUISTAO, REPUBLICA DO",                          "233"),
    "VU": ("VANUATU",                                            "234"),
    "VA": ("VATICANO, ESTADO DA CIDADE DO",                      "235"),
    "VE": ("VENEZUELA",                                          "236"),
    "VN": ("VIETNA",                                             "237"),
    "VG": ("VIRGENS, ILHAS (BRITANICAS)",                        "238"),
    "VI": ("VIRGENS, ILHAS (E.U.A.)",                            "239"),
    "UM-79": ("WAKE, ILHA",                                      "240"),
    "WF": ("WALLIS E FUTUNA, ILHAS",                             "241"),
    "ZM": ("ZAMBIA",                                             "242"),
    "ZW": ("ZIMBABUE",                                           "243"),
    "PA-CZ": ("ZONA DO CANAL DO PANAMA",                         "244"),
    "ME": ("MONTENEGRO",                                         "245"),
    "XX": ("EXTERIOR",                                           "246"),
    "UM": ("Pacífico, Ilhas do (Possessão dos EUA)",             "248"),
    "QA2": ("QATAR",                                             "249"),
    "KN2": ("SAINT KITTS E NEVIS",                               "250"),
    "CS": ("SERVIA E MONTENEGRO",                                "251"),
    "AX": ("ALAND, ILHAS",                                       "252"),
    "AQ": ("ANTARTICA",                                          "253"),
    "BQ": ("Bonaire, Saint Eustatius e Saba",                    "254"),
    "BV": ("BOUVET, ILHA",                                       "255"),
    "CW": ("CURACAO",                                            "256"),
    "HM": ("Heard e Ilhas McDonald, Ilha",                       "257"),
    "MF": ("São Martinho, Ilha de (Parte Francesa)",             "258"),
    "GS": ("Geórgia do Sul e Sandwich do Sul, Ilhas",            "259"),
    "JE": ("Jersey. Ilha do Canal",                              "260"),
    "YT": ("Mayotte",                                            "261"),
    "BL": ("São Bartolomeu",                                     "262"),
    "SJ": ("Svalbard e Jan Mayen",                               "263"),
    "TF": ("Terras Austrais Francesas",                          "264"),
    "SX": ("SAO MARTINHO, ILHA DE (PARTE HOLANDESA)",            "265"),
    "PS": ("Palestina",                                          "266"),
    "SS": ("Sudão do Sul",                                       "267"),
    "GG2": ("Guernsey, Ilha do Canal",                           "268"),
    "XB": ("Bancos Centrais",                                    "269"),
    "XO": ("Organizações Internacionais",                        "270"),
    "XF": ("FEZZAN",                                             "271"),
    "XD": ("DUBAI",                                              "272"),
    "XP": ("DELEGAÇÃO ESPECIAL DA PALESTINA",                    "273"),
}

# Mapa reverso: nome em inglês (Nominatim) → ISO Alpha-2
_NOME_EN_PARA_ISO: dict[str, str] = {
    "Afghanistan": "AF", "South Africa": "ZA", "Albania": "AL",
    "Germany": "DE", "Andorra": "AD", "Angola": "AO",
    "Anguilla": "AI", "Antigua and Barbuda": "AG",
    "Saudi Arabia": "SA", "Algeria": "DZ", "Argentina": "AR",
    "Armenia": "AM", "Aruba": "AW", "Australia": "AU",
    "Austria": "AT", "Azerbaijan": "AZ", "Bahamas": "BS",
    "Bahrain": "BH", "Bangladesh": "BD", "Barbados": "BB",
    "Belarus": "BY", "Belgium": "BE", "Belize": "BZ",
    "Benin": "BJ", "Bermuda": "BM", "Bolivia": "BO",
    "Bosnia and Herzegovina": "BA", "Botswana": "BW",
    "Brazil": "BR", "Brunei": "BN", "Bulgaria": "BG",
    "Burkina Faso": "BF", "Burundi": "BI", "Bhutan": "BT",
    "Cape Verde": "CV", "Cabo Verde": "CV", "Cameroon": "CM",
    "Cambodia": "KH", "Canada": "CA", "Guernsey": "GG",
    "Qatar": "QA", "Cayman Islands": "KY", "Kazakhstan": "KZ",
    "Chad": "TD", "Chile": "CL", "China": "CN", "Cyprus": "CY",
    "Christmas Island": "CX", "Singapore": "SG",
    "Cocos (Keeling) Islands": "CC", "Colombia": "CO",
    "Comoros": "KM", "Democratic Republic of the Congo": "CD",
    "Republic of the Congo": "CG", "Congo-Brazzaville": "CG",
    "Congo-Kinshasa": "CD", "Cook Islands": "CK",
    "North Korea": "KP", "South Korea": "KR",
    "Ivory Coast": "CI", "Côte d'Ivoire": "CI",
    "Costa Rica": "CR", "Kuwait": "KW", "Croatia": "HR",
    "Cuba": "CU", "Denmark": "DK", "Djibouti": "DJ",
    "Dominica": "DM", "Egypt": "EG", "El Salvador": "SV",
    "United Arab Emirates": "AE", "Ecuador": "EC",
    "Eritrea": "ER", "Scotland": "GB-SCT", "Slovakia": "SK",
    "Slovenia": "SI", "Spain": "ES",
    "United States": "US", "United States of America": "US",
    "Estonia": "EE", "Ethiopia": "ET",
    "Falkland Islands": "FK", "Faroe Islands": "FO",
    "Fiji": "FJ", "Philippines": "PH", "Finland": "FI",
    "Taiwan": "TW", "France": "FR", "Gabon": "GA",
    "Wales": "GB-WLS", "Gambia": "GM", "Ghana": "GH",
    "Georgia": "GE", "Gibraltar": "GI",
    "United Kingdom": "GB", "Great Britain": "GB",
    "Grenada": "GD", "Greece": "GR", "Greenland": "GL",
    "Guadeloupe": "GP", "Guam": "GU", "Guatemala": "GT",
    "Guyana": "GY", "French Guiana": "GF", "Guinea": "GN",
    "Guinea-Bissau": "GW", "Equatorial Guinea": "GQ",
    "Haiti": "HT", "Netherlands": "NL", "Holland": "NL",
    "Honduras": "HN", "Hong Kong": "HK", "Hungary": "HU",
    "Yemen": "YE", "India": "IN", "Indonesia": "ID",
    "England": "GB-ENG", "Iran": "IR", "Iraq": "IQ",
    "Ireland": "IE", "Northern Ireland": "GB-NIR",
    "Iceland": "IS", "Israel": "IL", "Italy": "IT",
    "Serbia": "RS", "Jamaica": "JM", "Japan": "JP",
    "Jordan": "JO", "Kiribati": "KI", "Laos": "LA",
    "Lesotho": "LS", "Latvia": "LV", "Lebanon": "LB",
    "Liberia": "LR", "Libya": "LY", "Liechtenstein": "LI",
    "Lithuania": "LT", "Luxembourg": "LU", "Macau": "MO",
    "Macao": "MO", "North Macedonia": "MK",
    "Madagascar": "MG", "Malaysia": "MY", "Malawi": "MW",
    "Maldives": "MV", "Mali": "ML", "Malta": "MT",
    "Isle of Man": "IM", "Northern Mariana Islands": "MP",
    "Morocco": "MA", "Marshall Islands": "MH",
    "Martinique": "MQ", "Mauritius": "MU",
    "Mauritania": "MR", "Mexico": "MX", "Myanmar": "MM",
    "Micronesia": "FM", "Mozambique": "MZ", "Moldova": "MD",
    "Monaco": "MC", "Mongolia": "MN", "Montserrat": "MS",
    "Namibia": "NA", "Nauru": "NR", "Nepal": "NP",
    "Nicaragua": "NI", "Niger": "NE", "Nigeria": "NG",
    "Niue": "NU", "Norfolk Island": "NF", "Norway": "NO",
    "New Caledonia": "NC", "New Zealand": "NZ", "Oman": "OM",
    "Palau": "PW", "Panama": "PA", "Papua New Guinea": "PG",
    "Pakistan": "PK", "Paraguay": "PY", "Peru": "PE",
    "Pitcairn Islands": "PN", "French Polynesia": "PF",
    "Poland": "PL", "Puerto Rico": "PR", "Portugal": "PT",
    "Kenya": "KE", "Kyrgyzstan": "KG",
    "Central African Republic": "CF",
    "Dominican Republic": "DO", "Réunion": "RE",
    "Romania": "RO", "Rwanda": "RW", "Russia": "RU",
    "Western Sahara": "EH", "Solomon Islands": "SB",
    "Samoa": "WS", "American Samoa": "AS",
    "San Marino": "SM", "Saint Helena": "SH",
    "Saint Lucia": "LC", "Saint Kitts and Nevis": "KN",
    "Saint Pierre and Miquelon": "PM",
    "Sao Tome and Principe": "ST",
    "Saint Vincent and the Grenadines": "VC",
    "Senegal": "SN", "Sierra Leone": "SL",
    "Seychelles": "SC", "Syria": "SY", "Somalia": "SO",
    "Sri Lanka": "LK", "Eswatini": "SZ", "Swaziland": "SZ",
    "Sudan": "SD", "Sweden": "SE", "Switzerland": "CH",
    "Suriname": "SR", "Tajikistan": "TJ", "Thailand": "TH",
    "Tanzania": "TZ", "Czech Republic": "CZ", "Czechia": "CZ",
    "British Indian Ocean Territory": "IO",
    "East Timor": "TL", "Timor-Leste": "TL", "Togo": "TG",
    "Tonga": "TO", "Tokelau": "TK",
    "Trinidad and Tobago": "TT", "Tunisia": "TN",
    "Turks and Caicos Islands": "TC", "Turkmenistan": "TM",
    "Turkey": "TR", "Türkiye": "TR", "Tuvalu": "TV",
    "Ukraine": "UA", "Uganda": "UG", "Uruguay": "UY",
    "Uzbekistan": "UZ", "Vanuatu": "VU",
    "Vatican City": "VA", "Venezuela": "VE",
    "Vietnam": "VN", "Viet Nam": "VN",
    "British Virgin Islands": "VG",
    "United States Virgin Islands": "VI",
    "Wallis and Futuna": "WF", "Zambia": "ZM",
    "Zimbabwe": "ZW", "Montenegro": "ME",
    "Curaçao": "CW", "Mayotte": "YT",
    "South Sudan": "SS", "Palestine": "PS",
    "State of Palestine": "PS", "Jersey": "JE",
    "Svalbard and Jan Mayen": "SJ",
}


# ════════════════════════════════════════════════════════════════════════════
# HELPERS GERAIS
# ════════════════════════════════════════════════════════════════════════════

def limpa_valor(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan", "-"):
        return ""
    v = str(v).strip()
    v = re.sub(r"[R$\s]", "", v)
    v = re.sub(r"\.(?=\d{3}[,.])", "", v)
    v = v.replace(",", ".")
    try:
        return f"{float(v):.2f}"
    except ValueError:
        return ""


def limpa_aliquota(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan", "-", "0"):
        return "0.00"
    v = str(v).strip().replace(",", ".")
    try:
        return f"{float(v):.2f}"
    except ValueError:
        return "0.00"


def formata_data(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan"):
        return ""
    v = str(v).strip().split(" ")[0]
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try:
            return datetime.strptime(v, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return v


def limpa_cnpj(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan"):
        return ""
    return re.sub(r"\D", "", str(v))


def limpa_im(v) -> str:
    if pd.isna(v) or str(v).strip() in ("", "nan"):
        return ""
    return re.sub(r"\D", "", str(v))


def safe(row, col: str, default: str = "") -> str:
    try:
        val = row[col]
        if pd.isna(val):
            return default
        return str(val).strip()
    except (KeyError, TypeError):
        return default


def gera_csv_dominio(registros: list) -> str:
    buf = io.StringIO()
    writer = csv.writer(
        buf, delimiter="|", lineterminator="\n",
        quoting=csv.QUOTE_NONE, escapechar="\\"
    )
    for reg in registros:
        writer.writerow(reg)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# GEOCODER — DETECÇÃO DE PAÍS
# ════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False, ttl=3600)
def detectar_pais_por_endereco(endereco: str) -> dict:
    """
    Consulta Nominatim/OSM e retorna o código interno do Domínio Sistemas.
    Retorna dict com: iso2, nome_en, nome_dominio, cod_dominio,
                      endereco_completo, confianca, erro.
    """
    resultado = {
        "iso2": "", "nome_en": "", "nome_dominio": "",
        "cod_dominio": "", "endereco_completo": "",
        "confianca": "baixa", "erro": None,
    }
    if not endereco or not endereco.strip():
        resultado["erro"] = "Endereço vazio."
        return resultado
    try:
        geolocator = Nominatim(
            user_agent="nfts_dominio_conversor/3.0",
            timeout=10,
        )
        time.sleep(1.1)
        location = geolocator.geocode(
            endereco, language="en",
            addressdetails=True, exactly_one=True,
        )
        if location is None:
            resultado["erro"] = f"Endereço não encontrado: '{endereco}'"
            return resultado

        raw     = location.raw.get("address", {})
        nome_en = raw.get("country", "")
        iso2    = raw.get("country_code", "").upper()

        resultado["nome_en"]           = nome_en
        resultado["iso2"]              = iso2
        resultado["endereco_completo"] = location.address

        # Resolve pelo ISO
        if iso2 in TABELA_PAISES:
            nome_dom, cod = TABELA_PAISES[iso2]
        else:
            # Fallback pelo nome em inglês
            iso2 = _NOME_EN_PARA_ISO.get(nome_en, "")
            resultado["iso2"] = iso2
            nome_dom, cod = TABELA_PAISES.get(iso2, ("", ""))

        resultado["nome_dominio"] = nome_dom
        resultado["cod_dominio"]  = cod

        if iso2 and nome_en and cod:
            resultado["confianca"] = "alta"
        elif cod:
            resultado["confianca"] = "media"
        else:
            resultado["erro"] = (
                f"País '{nome_en}' (ISO: {iso2}) não encontrado "
                "na tabela Domínio. Selecione manualmente."
            )
    except GeocoderTimedOut:
        resultado["erro"] = "Timeout ao consultar o geocoder. Tente novamente."
    except GeocoderServiceError as e:
        resultado["erro"] = f"Erro no serviço de geocoding: {e}"
    except Exception as e:
        resultado["erro"] = f"Erro inesperado: {e}"
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# MAPEADORES DE REGISTRO
# ════════════════════════════════════════════════════════════════════════════

def reg_0000(cnpj_empresa: str) -> list:
    """Registro 0000 – Identificação da empresa."""
    return ["0000", limpa_cnpj(cnpj_empresa)]


def reg_0020(row, cod_pais: str = "") -> list:
    """
    Registro 0020 – Cadastro de fornecedores.
    cod_pais: código interno Domínio (ex: "76" para EUA).
              Preenchido pelo geocoder para prestadores estrangeiros.
    """
    ind      = safe(row, "Indicador de CPF/CNPJ do Prestador")
    cnpj     = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))

    # Estrangeiro (Indicador = 3) ou sem CNPJ
    if ind == "3" or not cnpj:
        cnpj = "Outros"

    razao    = safe(row, "Razão Social do Prestador")[:150]
    endereco = safe(row, "Endereço do Prestador")
    numero   = re.sub(r"\D", "", safe(row, "Número do Endereço do Prestador"))
    compl    = safe(row, "Complemento do Endereço do Prestador")
    bairro   = safe(row, "Bairro do Prestador")
    uf       = safe(row, "UF do Prestador")
    cep      = re.sub(r"\D", "", safe(row, "CEP do Prestador"))
    email    = safe(row, "Email do Prestador")
    insc_mun = limpa_im(safe(row, "Inscrição Municipal do Prestador"))

    # Exterior: UF = EX, código do país preenchido
    pais_campo = ""
    if ind == "3":
        uf         = "EX"
        pais_campo = cod_pais   # campo 11 — código interno Domínio

    return [
        "0020",      # 1  - Identificação do registro
        cnpj,        # 2  - Inscrição (CNPJ/CPF/Outros)
        razao,       # 3  - Razão Social
        "",          # 4  - Apelido
        endereco,    # 5  - Endereço
        numero,      # 6  - Número do endereço
        compl,       # 7  - Complemento
        bairro,      # 8  - Bairro
        "",          # 9  - Código do município
        uf,          # 10 - UF (EX para exterior)
        pais_campo,  # 11 - Código do País (interno Domínio)
        cep,         # 12 - CEP
        "",          # 13 - Inscrição Estadual
        insc_mun,    # 14 - Inscrição Municipal
        "",          # 15 - Inscrição Suframa
        "",          # 16 - DDD
        "",          # 17 - Telefone
        "",          # 18 - FAX
        "",          # 19 - Data do cadastro
        "",          # 20 - Conta contábil
        "",          # 21 - Conta contábil cliente
        "N",         # 22 - Agropecuário
        "7",         # 23 - Natureza jurídica (7=Empresa Privada)
        "N",         # 24 - Regime de apuração (N=Normal)
        "N",         # 25 - Contribuinte ICMS
        "",          # 26 - Alíquota ICMS
        "",          # 27 - Categoria do estabelecimento
        "",          # 28 - Inscrição Estadual ST
        email,       # 29 - Email
        "N",         # 30 - Interdependência
        "N",         # 31 - Contribuinte CPRB
        "",          # 32 - Processo administrativo/judicial
        "",          # 33 - Tipo Inscrição
    ]


def reg_1000(row, cfop: str, cod_acumulador: str, cod_especie: str) -> list:
    """
    Registro 1000 – Nota Fiscal de Entrada.

    Mapeamento NFTS → Domínio:
      Número do Documento           → campo 8
      Série do Documento            → campo 9
      Data da Prestação de Serviços → campo 11 (Data entrada) e campo 48 (Competência)
      Data Hora Emissão NFTS        → campo 12 (Data emissão)
      Valor dos Serviços            → campo 13 (Valor contábil) e campo 39 (Valor produtos)
      ISS Retido = S                → campo 20 = "18"
      ISS Retido = N                → campo 20 = ""
    """
    ind  = safe(row, "Indicador de CPF/CNPJ do Prestador")
    cnpj = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))
    if ind == "3" or not cnpj:
        cnpj = "Outros"

    num_doc = re.sub(r"\.0$", "", safe(row, "Número do Documento"))
    serie   = safe(row, "Série do Documento")
    if serie in ("-", "nan", ""):
        serie = ""

    data_entrada = formata_data(safe(row, "Data da Prestação de Serviços"))
    data_emissao = formata_data(safe(row, "Data Hora Emissão NFTS"))
    valor_cont   = limpa_valor(safe(row, "Valor dos Serviços"))

    iss_retido     = safe(row, "ISS Retido").strip().upper()
    cod_recolh_iss = "18" if iss_retido == "S" else ""

    insc_mun = limpa_im(safe(row, "Inscrição Municipal do Prestador"))

    return [
        "1000",          # 1  - Identificação do registro
        cod_especie,     # 2  - Código da espécie (03=NFS)
        cnpj,            # 3  - Inscrição fornecedor
        "",              # 4  - Código de Exclusão da DIEF
        cod_acumulador,  # 5  - Código do acumulador
        cfop,            # 6  - CFOP
        "",              # 7  - Segmento
        num_doc,         # 8  - Número do documento
        serie,           # 9  - Série
        "",              # 10 - Número do documento final
        data_entrada,    # 11 - Data da entrada (dd/mm/aaaa)
        data_emissao,    # 12 - Data emissão (dd/mm/aaaa)
        valor_cont,      # 13 - Valor contábil
        "",              # 14 - Valor da exclusão da DIEF
        "",              # 15 - Observação
        "S",             # 16 - Modalidade do frete (S=Sem frete)
        "T",             # 17 - Emitente (T=Terceiros)
        "",              # 18 - CFOP estendido (só SE)
        "",              # 19 - Código transferência crédito (só RS)
        cod_recolh_iss,  # 20 - Código Recolhimento ISS Retido
        "",              # 21 - Código Recolhimento IRRF
        "",              # 22 - Código da observação
        "",              # 23 - Data do visto (só MG)
        "E",             # 24 - Fato gerador CRF (E=Emissão)
        "E",             # 25 - Fato gerador IRRF (E=Emissão)
        "",              # 26 - Valor do frete
        "",              # 27 - Valor do seguro
        "",              # 28 - Valor das despesas
        "",              # 29 - Valor do PIS
        "",              # 30 - Código Antecipação Tributária
        "",              # 31 - Valor do COFINS
        "",              # 32 - Valor DARE (só SE)
        "",              # 33 - Alíquota DARE (só SE)
        "",              # 34 - Valor base cálculo ICMS ST
        "",              # 35 - Entradas isentas (só MG)
        "",              # 36 - Outras entradas isentas (só MG)
        "",              # 37 - Valor transporte incluído na base (só MG)
        "",              # 38 - Código de ressarcimento
        valor_cont,      # 39 - Valor produtos
        "",              # 40 - Município Origem
        "0",             # 41 - Situação da Nota (0=Documento Regular)
        "",              # 42 - Código da situação tributária
        "",              # 43 - Sub série
        "",              # 44 - Inscrição estadual do fornecedor
        insc_mun,        # 45 - Inscrição municipal do fornecedor
        "",              # 46 - Código da operação e prestação
        "",              # 47 - Valor a ser deduzido da receita tributável
        data_entrada,    # 48 - Competência (dd/mm/aaaa)
        "",              # 49 - Operação (só PA)
        "",              # 50 - Número do parecer fiscal
        "",              # 51 - Data do parecer fiscal
        "",              # 52 - Número da declaração de Importação
        "N",             # 53 - Possui benefício fiscal
        "",              # 54 - Chave da nota fiscal eletrônica
        "",              # 55 - Código de recolhimento FETHAB
        "",              # 56 - Responsável pelo recolhimento FETHAB
        "",              # 57 - CFOP documento fiscal
        "",              # 58 - Tipo de CT-e
        "",              # 59 - CT-e referência
        "",              # 60 - Modalidade da importação
        "",              # 61 - Código da informação complementar
        "",              # 62 - Informação complementar
        "",              # 63 - Classe de consumo
        "",              # 64 - Tipo de ligação
        "",              # 65 - Grupo de tensão
        "",              # 66 - Tipo de assinante
        "",              # 67 - KWH consumido
        "",              # 68 - Valor fornecido/consumido gás ou energia
        "",              # 69 - Valor cobrado de terceiros
        "",              # 70 - Tipo do documento de importação
        "",              # 71 - Número do Ato Concessório Drawback
        "",              # 72 - Natureza do frete PIS/COFINS
        "",              # 73 - CST PIS/COFINS
        "",              # 74 - Base do crédito PIS/COFINS
        "",              # 75 - Valor serviços/itens PIS/COFINS
        "",              # 76 - Base de cálculo PIS/COFINS
        "",              # 77 - Alíquota de PIS
        "",              # 78 - Alíquota de COFINS
        "",              # 79 - Chave de NFSe
        "",              # 80 - Número do processo ou ato concessório
        "",              # 81 - Origem do processo
        "",              # 82 - Data da escrituração
        "",              # 83 - CFPS (só DF)
        "",              # 84 - Natureza da receita PIS/COFINS
        "",              # 85 - CST IPI
        "",              # 86 - Lançamentos de SCP
        "",              # 87 - Tipo de serviço
        "",              # 88 - Município destino
        "",              # 89 - Pedágio
        "",              # 90 - IPI
        "",              # 91 - ICMS ST
        "",              # 92 - EFD-Reinf Tipo de serviço
        "",              # 93 - EFD-Reinf Indicativo Prestação
        "",              # 94 - Número doc. arrecadação (só RS)
        "",              # 95 - Tipo do título
        "",              # 96 - Identificação
        "",              # 97 - ICMS Desonerado
        "",              # 98 - IPI Devolução
    ]


def reg_1020(row) -> list:
    """
    Registro 1020 – Impostos da Nota Fiscal de Entrada.

    ISS Retido = S → Código 18 (ISS Retido)
                     Base    = Valor dos Serviços
                     Alíq.   = Alíquota da nota
                     Valor   = Valor ISS
                     Outras  = vazio
                     V.Cont. = Valor dos Serviços

    ISS Retido = N → Código 3 (ISS)
                     Base    = 0.00  (zerado)
                     Alíq.   = 0.00  (zerado)
                     Valor   = 0.00  (zerado)
                     Outras  = Valor dos Serviços
                     V.Cont. = Valor dos Serviços
    """
    iss_retido     = safe(row, "ISS Retido").strip().upper()
    valor_servicos = limpa_valor(safe(row, "Valor dos Serviços"))
    aliquota       = limpa_aliquota(safe(row, "Alíquota"))
    valor_iss      = limpa_valor(safe(row, "Valor ISS"))

    if iss_retido == "S":
        cod_imposto = "18"          # ISS Retido
        base        = valor_servicos
        aliq        = aliquota
        valor       = valor_iss
        outras      = ""
    else:
        cod_imposto = "3"           # ISS Normal
        base        = "0.00"
        aliq        = "0.00"
        valor       = "0.00"
        outras      = valor_servicos

    return [
        "1020",         # 1  - Identificação do registro
        cod_imposto,    # 2  - Código do imposto (18=ISS Retido / 3=ISS)
        "",             # 3  - Percentual de redução da base de cálculo
        base,           # 4  - Base de cálculo
        aliq,           # 5  - Alíquota
        valor,          # 6  - Valor do Imposto
        "",             # 7  - Valor de Isentas
        outras,         # 8  - Valor de Outras (Valor Serviços quando ISS Normal)
        "",             # 9  - Valor do IPI
        "",             # 10 - Valor da substituição Tributária
        valor_servicos, # 11 - Valor Contábil (sempre = Valor dos Serviços)
        "",             # 12 - Código do recolhimento do imposto
        "",             # 13 - Valor não tributadas (só GO)
        "",             # 14 - Valor parcela reduzida (só GO)
        "",             # 15 - Alíq. Interestadual
        "",             # 16 - Nat. rend.
        "",             # 17 - Tipo de Dedução
        "",             # 18 - Tipo de Isenção
        "",             # 19 - Descrição
    ]


def reg_1150(row) -> list:
    """
    Registro 1150 – IBS (Imposto sobre Bens e Serviços).
    Filho do registro 1000. Campos: cClassTrib, Base, Alíquota, Valor.
    Gerado em branco — alíquotas IBS ainda não definidas para a maioria dos serviços.
    """
    return [
        "1150",  # 1 - Identificação do registro (fixo)
        "",      # 2 - IBS cClassTrib
        "",      # 3 - IBS Base de cálculo (2 decimais)
        "",      # 4 - IBS Alíquota (2 decimais)
        "",      # 5 - IBS Valor (2 decimais)
    ]


def reg_1151(row) -> list:
    """
    Registro 1151 – CBS (Contribuição sobre Bens e Serviços).
    Filho do registro 1000. Campos: cClassTrib, Base, Alíquota, Valor.
    Gerado em branco — alíquotas CBS ainda não definidas para a maioria dos serviços.
    """
    return [
        "1151",  # 1 - Identificação do registro (fixo)
        "",      # 2 - CBS cClassTrib
        "",      # 3 - CBS Base de cálculo (2 decimais)
        "",      # 4 - CBS Alíquota (2 decimais)
        "",      # 5 - CBS Valor (2 decimais)
    ]


# ════════════════════════════════════════════════════════════════════════════
# CONVERSOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

def converte_nfts(
    df: pd.DataFrame,
    cnpj_empresa: str,
    cfop: str,
    cod_acumulador: str,
    cod_especie: str,
    gerar_ibs_cbs: bool,
    apenas_iss_retido: bool,
    auto_geocode: bool,
    pais_manual_override: str,
) -> tuple:
    """
    Converte DataFrame da NFTS para registros Domínio.
    Retorna (csv_string, df_preview, erros_list).
    """
    registros = []
    preview   = []
    erros     = []

    # Registro 0000 — uma única vez
    registros.append(reg_0000(cnpj_empresa))

    # Filtra apenas Tipo de Registro = "4" (notas)
    df_notas = df[df["Tipo de Registro"].astype(str).str.strip() == "4"].copy()
    if df_notas.empty:
        erros.append("Nenhum registro do tipo '4' (nota) encontrado no arquivo.")
        return "", pd.DataFrame(), erros

    if apenas_iss_retido:
        df_notas = df_notas[
            df_notas["ISS Retido"].astype(str).str.strip().str.upper() == "S"
        ]
        if df_notas.empty:
            erros.append("Nenhuma nota com ISS Retido = 'S' encontrada.")
            return "", pd.DataFrame(), erros

    fornecedores_gerados: set = set()

    for idx, row in df_notas.iterrows():
        try:
            ind   = safe(row, "Indicador de CPF/CNPJ do Prestador")
            cnpj  = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))
            razao = safe(row, "Razão Social do Prestador")

            chave_forn = (
                f"EXT_{razao[:50]}" if (ind == "3" or not cnpj) else cnpj
            )

            # ── Detecta país para prestadores estrangeiros ──────────────
            cod_pais_ext  = ""
            nome_pais_ext = ""
            geo_info      = ""

            if ind == "3":
                # Prioridade 1: override manual do usuário
                if pais_manual_override:
                    cod_pais_ext  = pais_manual_override
                    nome_pais_ext = next(
                        (n for iso, (n, c) in TABELA_PAISES.items()
                         if c == pais_manual_override), ""
                    )
                    geo_info = f"Manual: {nome_pais_ext} (cód. {cod_pais_ext})"

                # Prioridade 2: geocoder automático
                elif auto_geocode:
                    end_parts = [
                        safe(row, "Endereço do Prestador"),
                        safe(row, "Número do Endereço do Prestador"),
                        safe(row, "Bairro do Prestador"),
                        safe(row, "Cidade do Prestador"),
                        safe(row, "UF do Prestador"),
                    ]
                    end_completo = " ".join(p for p in end_parts if p)
                    if end_completo.strip():
                        geo = detectar_pais_por_endereco(end_completo)
                        cod_pais_ext  = geo.get("cod_dominio", "")
                        nome_pais_ext = geo.get("nome_dominio", "")
                        if geo.get("erro") and not cod_pais_ext:
                            geo_info = f"⚠️ {geo['erro']}"
                            erros.append(
                                f"Linha {idx} ({razao}): {geo['erro']}"
                            )
                        else:
                            geo_info = (
                                f"OSM: {nome_pais_ext} "
                                f"(cód. {cod_pais_ext})"
                            )

            # ── Registro 0020 — apenas uma vez por fornecedor ───────────
            if chave_forn not in fornecedores_gerados:
                registros.append(reg_0020(row, cod_pais=cod_pais_ext))
                fornecedores_gerados.add(chave_forn)

            # ── Registro 1000 ───────────────────────────────────────────
            registros.append(
                reg_1000(row, cfop, cod_acumulador, cod_especie)
            )

            # ── Registro 1020 — sempre gerado ───────────────────────────
            registros.append(reg_1020(row))

            # ── Registros 1150 / 1151 — opcionais ──────────────────────
            if gerar_ibs_cbs:
                registros.append(reg_1150(row))
                registros.append(reg_1151(row))

            # ── Preview ─────────────────────────────────────────────────
            iss_ret = safe(row, "ISS Retido").upper()
            preview.append({
                "Nº NFTS"        : safe(row, "Nº NFTS"),
                "Prestador"      : razao,
                "CNPJ/CPF"       : safe(row, "CPF/CNPJ do Prestador"),
                "Tipo"           : "Estrangeiro" if ind == "3" else "Nacional",
                "País (Domínio)" : f"{nome_pais_ext} [{cod_pais_ext}]"
                                   if cod_pais_ext else (
                                       geo_info if geo_info else "—"
                                   ),
                "Emissão"        : safe(row, "Data Hora Emissão NFTS"),
                "Prestação"      : safe(row, "Data da Prestação de Serviços"),
                "Valor Serviços" : safe(row, "Valor dos Serviços"),
                "Alíquota ISS"   : safe(row, "Alíquota"),
                "Valor ISS"      : safe(row, "Valor ISS"),
                "ISS Retido"     : iss_ret,
                "Cód. Imposto"   : "18 – ISS Retido" if iss_ret == "S"
                                   else "3 – ISS",
                "Base 1020"      : limpa_valor(safe(row, "Valor dos Serviços"))
                                   if iss_ret == "S" else "0.00",
                "Outras 1020"    : "" if iss_ret == "S"
                                   else limpa_valor(
                                       safe(row, "Valor dos Serviços")
                                   ),
            })

        except Exception as e:
            erros.append(f"Linha {idx}: {e}")

    csv_out    = gera_csv_dominio(registros)
    df_preview = pd.DataFrame(preview)
    return csv_out, df_preview, erros


# ════════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ════════════════════════════════════════════════════════════════════════════

st.title("🧾 Conversor NFTS Paulistana → Domínio Sistemas")
st.caption(
    "Converte o CSV exportado da **NFS-e Tomadas (NFTS)** da Prefeitura de São Paulo "
    "para o layout de importação do **Domínio Sistemas** — "
    "Registros 0000 · 0020 · 1000 · 1020 · 1150 · 1151"
)

# ── Tabs principais ──────────────────────────────────────────────────────
tab_converter, tab_paises, tab_ajuda = st.tabs([
    "🔄 Converter", "🌍 Países", "📖 Ajuda"
])

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Configurações
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ Configurações")

    st.subheader("Empresa Tomadora")
    cnpj_empresa = st.text_input(
        "CNPJ (apenas números)",
        placeholder="20586841000130",
        help="CNPJ do tomador dos serviços (sua empresa).",
    )

    st.subheader("Parâmetros Domínio")
    cfop = st.text_input(
        "CFOP", value="2933",
        help="CFOP para serviços tomados. Ex.: 2933.",
    )
    cod_acumulador = st.text_input(
        "Código do Acumulador", value="",
        help="Código do acumulador configurado no Domínio.",
    )
    cod_especie = st.text_input(
        "Código da Espécie", value="03",
        help="03 = Nota Fiscal de Serviço.",
    )

    st.subheader("Opções de Geração")
    apenas_iss_retido = st.checkbox(
        "Apenas notas com ISS Retido", value=False,
    )
    gerar_ibs_cbs = st.checkbox(
        "Gerar registros 1150 (IBS) e 1151 (CBS)", value=False,
        help="Gera os registros em branco. Ative se o Domínio exigir.",
    )

    st.subheader("🌍 Prestadores Estrangeiros")
    auto_geocode = st.checkbox(
        "Detectar país automaticamente (OSM/Nominatim)",
        value=True,
        help="Consulta o OpenStreetMap para identificar o país pelo endereço.",
    )

    # Selectbox de override manual — ordenado pelo código Domínio
    opcoes_iso = sorted(
        TABELA_PAISES.keys(), key=lambda k: int(TABELA_PAISES[k][1])
    )
    opcoes_label = ["(Automático / sem override)"] + [
        f"{TABELA_PAISES[iso][1]:>3} – {TABELA_PAISES[iso][0]}"
        for iso in opcoes_iso
    ]
    sel_override = st.selectbox(
        "Forçar país para todos os estrangeiros",
        options=opcoes_label,
        index=0,
        help=(
            "Selecione para forçar um país específico para TODOS os "
            "prestadores estrangeiros, ignorando o geocoder."
        ),
    )
    pais_manual_override = ""
    if sel_override != "(Automático / sem override)":
        pais_manual_override = sel_override.split(" – ")[0].strip()

    st.divider()
    st.caption("v3.0 · NFTS → Domínio Sistemas")


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — CONVERTER
# ════════════════════════════════════════════════════════════════════════════
with tab_converter:
    st.subheader("📂 Upload do arquivo NFTS")
    arquivo = st.file_uploader(
        "Selecione o CSV exportado da NFTS (Prefeitura de SP)",
        type=["csv"],
        help="Ex.: NFTS_50431480_20260601_20260630.csv",
    )

    if arquivo:
        # Leitura com fallback de encoding
        try:
            df_raw = pd.read_csv(
                arquivo, sep=",", dtype=str,
                encoding="utf-8", on_bad_lines="skip",
            )
        except UnicodeDecodeError:
            arquivo.seek(0)
            df_raw = pd.read_csv(
                arquivo, sep=",", dtype=str,
                encoding="latin-1", on_bad_lines="skip",
            )

        total_linhas = len(df_raw)
        notas_tipo4  = (
            df_raw["Tipo de Registro"].astype(str).str.strip() == "4"
        ).sum()
        notas_ext = 0
        if "Indicador de CPF/CNPJ do Prestador" in df_raw.columns:
            df_t4 = df_raw[
                df_raw["Tipo de Registro"].astype(str).str.strip() == "4"
            ]
            notas_ext = (
                df_t4["Indicador de CPF/CNPJ do Prestador"]
                .astype(str).str.strip() == "3.0"
            ).sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Linhas no arquivo", total_linhas)
        c2.metric("Notas (Tipo 4)",    notas_tipo4)
        c3.metric("Prestadores EXT",   notas_ext)

        with st.expander("👁️ Dados brutos do CSV"):
            st.dataframe(df_raw, use_container_width=True)

        # Validações
        avisos = []
        if not cnpj_empresa:
            avisos.append("⚠️ Informe o **CNPJ da empresa** na barra lateral.")
        if not cfop:
            avisos.append("⚠️ Informe o **CFOP** na barra lateral.")
        if not cod_acumulador:
            avisos.append(
                "⚠️ Informe o **Código do Acumulador** na barra lateral."
            )
        for av in avisos:
            st.warning(av)

        if not avisos:
            if notas_ext > 0 and auto_geocode and not pais_manual_override:
                st.info(
                    f"🌍 **{notas_ext}** nota(s) com prestador estrangeiro "
                    "detectada(s). O país será buscado automaticamente pelo "
                    "OpenStreetMap (pode levar alguns segundos por nota)."
                )

            if st.button("🚀 Converter para layout Domínio", type="primary"):
                with st.spinner("Convertendo..."):
                    csv_saida, df_prev, erros = converte_nfts(
                        df=df_raw,
                        cnpj_empresa=cnpj_empresa,
                        cfop=cfop,
                        cod_acumulador=cod_acumulador,
                        cod_especie=cod_especie,
                        gerar_ibs_cbs=gerar_ibs_cbs,
                        apenas_iss_retido=apenas_iss_retido,
                        auto_geocode=auto_geocode,
                        pais_manual_override=pais_manual_override,
                    )

                if erros:
                    with st.expander(
                        f"⚠️ {len(erros)} aviso(s) durante a conversão"
                    ):
                        for e in erros:
                            st.warning(e)

                if not df_prev.empty:
                    st.success(
                        f"✅ Conversão concluída! "
                        f"**{len(df_prev)}** nota(s) processada(s)."
                    )

                    # Métricas
                    m1, m2, m3, m4 = st.columns(4)
                    n_ret = (df_prev["ISS Retido"] == "S").sum()
                    n_nor = (df_prev["ISS Retido"] == "N").sum()
                    total_val = df_prev["Valor Serviços"].apply(
                        lambda x: float(limpa_valor(x))
                        if limpa_valor(x) else 0.0
                    ).sum()
                    m1.metric("Total de notas",       len(df_prev))
                    m2.metric("ISS Retido (cód. 18)", n_ret)
                    m3.metric("ISS Normal (cód. 3)",  n_nor)
                    m4.metric("Total Serviços R$",    f"{total_val:,.2f}")

                    # Preview com destaque por ISS Retido
                    st.subheader("📋 Preview das notas convertidas")

                    def highlight_iss(row):
                        cor = (
                            "#d4edda" if row["ISS Retido"] == "S"
                            else "#fff3cd"
                        )
                        return [f"background-color: {cor}"] * len(row)

                    st.dataframe(
                        df_prev.style.apply(highlight_iss, axis=1),
                        use_container_width=True,
                    )

                    # Regra aplicada
                    with st.expander("📐 Regra aplicada no Registro 1020"):
                        st.markdown("""
| Situação | Cód. | Base | Alíquota | Valor ISS | Campo **Outras** | V. Contábil |
|---|---|---|---|---|---|---|
| **ISS Retido** (`S`) | `18` | Valor Serviços | Alíquota nota | Valor ISS | *(vazio)* | Valor Serviços |
| **ISS Normal** (`N`) | `3` | `0.00` | `0.00` | `0.00` | **Valor Serviços** | Valor Serviços |
                        """)

                    # Prévia do arquivo gerado
                    with st.expander(
                        "📄 Prévia do arquivo gerado (50 primeiras linhas)"
                    ):
                        linhas = csv_saida.split("\n")[:50]
                        st.code("\n".join(linhas), language="text")

                    # Download
                    nome_saida = arquivo.name.replace(".csv", "_DOMINIO.txt")
                    st.download_button(
                        label="⬇️ Baixar arquivo para importação no Domínio",
                        data=csv_saida.encode("utf-8"),
                        file_name=nome_saida,
                        mime="text/plain",
                        type="primary",
                    )
                    st.info(
                        "💡 **Como importar:** No Domínio Sistemas acesse "
                        "*Utilitários → Importação → Notas Fiscais de Entrada* "
                        "e selecione o arquivo gerado."
                    )
    else:
        st.info("⬆️ Faça o upload do arquivo CSV da NFTS para iniciar.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — PAÍSES
# ════════════════════════════════════════════════════════════════════════════
with tab_paises:
    st.subheader("🌍 Consulta de País por Endereço")
    st.markdown(
        "Identifique o **código interno do Domínio Sistemas** de qualquer país "
        "a partir de um endereço. Útil para conferir prestadores estrangeiros "
        "antes de converter."
    )

    col_end, col_btn = st.columns([4, 1])
    with col_end:
        end_consulta = st.text_input(
            "Endereço para busca",
            value="2108 North Street, Sacramento, CA",
            placeholder="Ex.: 2108 North Street, Sacramento, CA",
            key="geo_manual_end",
        )
    with col_btn:
        st.write("")
        st.write("")
        buscar_pais = st.button("🔍 Buscar", key="geo_manual_btn")

    iso_detectado = ""
    if buscar_pais and end_consulta:
        with st.spinner("Consultando OpenStreetMap..."):
            res = detectar_pais_por_endereco(end_consulta)

        if res["cod_dominio"]:
            icone = {"alta": "✅", "media": "🟡", "baixa": "🔴"}.get(
                res["confianca"], "❓"
            )
            st.success(
                f"{icone} **País detectado:** {res['nome_dominio']}  \n"
                f"🌐 **Nome (EN):** {res['nome_en']}  \n"
                f"🏷️ **ISO Alpha-2:** `{res['iso2']}`  \n"
                f"🔢 **Código Domínio:** `{res['cod_dominio']}`  \n"
                f"📍 **Endereço completo:** {res['endereco_completo']}"
            )
            iso_detectado = res["iso2"]
        else:
            st.error(f"❌ {res['erro']}")
            if res.get("nome_en"):
                st.info(
                    f"País detectado pelo OSM: **{res['nome_en']}** "
                    f"(ISO: `{res['iso2']}`) — não encontrado na tabela Domínio. "
                    "Selecione manualmente abaixo."
                )

    # Selectbox de confirmação manual
    st.markdown("---")
    st.markdown("##### Selecionar / confirmar país manualmente")
    opcoes_iso_t = sorted(
        TABELA_PAISES.keys(), key=lambda k: int(TABELA_PAISES[k][1])
    )
    opcoes_label_t = [
        f"{TABELA_PAISES[iso][1]:>3} – {TABELA_PAISES[iso][0]}"
        for iso in opcoes_iso_t
    ]
    idx_def = 0
    if iso_detectado and iso_detectado in opcoes_iso_t:
        idx_def = opcoes_iso_t.index(iso_detectado)

    sel_manual = st.selectbox(
        "País (Tabela Domínio — ordenado por código)",
        options=opcoes_label_t,
        index=idx_def,
        key="geo_manual_sel",
    )
    if sel_manual:
        cod_sel = sel_manual.split(" – ")[0].strip()
        nome_sel = sel_manual.split(" – ", 1)[1] if " – " in sel_manual else ""
        st.success(
            f"✅ Código Domínio: **`{cod_sel}`** — {nome_sel}  \n"
            "Use este valor no campo 11 do Registro 0020 "
            "ou selecione-o no override da barra lateral."
        )

    # Tabela completa
    st.markdown("---")
    st.markdown("##### 📋 Tabela completa de Países — Domínio Sistemas")

    busca_pais = st.text_input(
        "🔎 Filtrar tabela",
        placeholder="Digite nome ou código...",
        key="filtro_tabela_paises",
    )

    df_paises = pd.DataFrame([
        {
            "Código Domínio": int(cod),
            "ISO Alpha-2":    iso,
            "Nome (Domínio)": nome,
        }
        for iso, (nome, cod) in TABELA_PAISES.items()
    ]).sort_values("Código Domínio").reset_index(drop=True)

    if busca_pais:
        mask = (
            df_paises["Nome (Domínio)"].str.contains(
                busca_pais, case=False, na=False
            ) |
            df_paises["Código Domínio"].astype(str).str.contains(
                busca_pais, na=False
            ) |
            df_paises["ISO Alpha-2"].str.contains(
                busca_pais, case=False, na=False
            )
        )
        df_paises = df_paises[mask]

    st.dataframe(df_paises, use_container_width=True, height=400)
    st.caption(f"Total: {len(df_paises)} países")


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — AJUDA
# ════════════════════════════════════════════════════════════════════════════
with tab_ajuda:
    st.subheader("📖 Documentação e Mapeamento de Campos")

    st.markdown("""
### Registros gerados

| Registro | Descrição | Fonte |
|---|---|---|
| `0000` | Identificação da empresa (tomadora) | Configuração |
| `0020` | Cadastro do fornecedor (prestador) — 1 por CNPJ | NFTS |
| `1000` | Nota Fiscal de Entrada | NFTS |
| `1020` | Impostos (ISS Retido cód. 18 **ou** ISS Normal cód. 3) | NFTS |
| `1150` | IBS — Imposto sobre Bens e Serviços (opcional) | Em branco |
| `1151` | CBS — Contribuição sobre Bens e Serviços (opcional) | Em branco |

---

### Regra do Registro 1020

| ISS Retido | Cód. | Campo 4 (Base) | Campo 5 (Alíq.) | Campo 6 (Valor) | Campo 8 (Outras) | Campo 11 (V.Cont.) |
|---|---|---|---|---|---|---|
| **S** | `18` | Valor Serviços | Alíquota nota | Valor ISS | *(vazio)* | Valor Serviços |
| **N** | `3` | `0.00` | `0.00` | `0.00` | **Valor Serviços** | Valor Serviços |

---

### Mapeamento NFTS → Domínio (Reg. 1000)

| Campo NFTS | Campo Domínio | Nº |
|---|---|---|
| CPF/CNPJ do Prestador | Inscrição fornecedor | 3 |
| Código da espécie | Configurável (padrão `03`) | 2 |
| Código do acumulador | Configurável na sidebar | 5 |
| CFOP | Configurável na sidebar | 6 |
| Número do Documento | Número do documento | 8 |
| Série do Documento | Série | 9 |
| Data da Prestação de Serviços | Data da entrada + Competência | 11, 48 |
| Data Hora Emissão NFTS | Data emissão | 12 |
| Valor dos Serviços | Valor contábil + Valor produtos | 13, 39 |
| ISS Retido = S | Cód. Recolhimento ISS = `18` | 20 |
| Inscrição Municipal do Prestador | Inscrição municipal fornecedor | 45 |

---

### Mapeamento NFTS → Domínio (Reg. 0020)

| Campo NFTS | Campo Domínio | Nº | Observação |
|---|---|---|---|
| CPF/CNPJ do Prestador | Inscrição | 2 | `Outros` para estrangeiros |
| Razão Social do Prestador | Razão Social | 3 | Máx. 150 chars |
| Endereço do Prestador | Endereço | 5 | |
| Número do Endereço | Número | 6 | Apenas números |
| Complemento | Complemento | 7 | |
| Bairro do Prestador | Bairro | 8 | |
| UF do Prestador | UF | 10 | `EX` para exterior |
| País (geocoder/manual) | Código do País | 11 | Código interno Domínio |
| CEP do Prestador | CEP | 12 | |
| Inscrição Municipal do Prestador | Inscrição Municipal | 14 | |
| Email do Prestador | Email | 29 | |

---

### Detecção automática de país (prestadores estrangeiros)

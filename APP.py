"""
app.py - Conversor NFTS Paulistana para Dominio Sistemas
Versao 3.4 Final - arquivo unico
Regras automaticas:
  - CNPJ empresa tomadora: lido do campo CPF/CNPJ do Tomador do CSV
  - CFOP: 1933 (UF prestador = SP) ou 2933 (fora SP / exterior)
  - Especie: 03 (nacional) ou 39 (exterior - Indicador=3)
  - Acumulador: lookup PAULISTANA ou 2551 (importacao)
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
    page_title="NFTS para Dominio Sistemas",
    page_icon="🧾",
    layout="wide",
)

# ════════════════════════════════════════════════════════════════════════════
# TABELA OFICIAL DE PAISES - DOMINIO SISTEMAS
# ════════════════════════════════════════════════════════════════════════════
TABELA_PAISES: dict = {
    "AF": ("AFEGANISTAO", "1"),
    "ZA": ("AFRICA DO SUL", "2"),
    "AL": ("ALBANIA, REPUBLICA DA", "3"),
    "DE": ("ALEMANHA", "4"),
    "AD": ("ANDORRA", "5"),
    "AO": ("ANGOLA", "6"),
    "AI": ("ANGUILLA", "7"),
    "AG": ("ANTIGUA E BARBUDA", "8"),
    "AN": ("ANTILHAS HOLANDESAS", "9"),
    "SA": ("ARABIA SAUDITA", "10"),
    "DZ": ("ARGELIA", "11"),
    "AR": ("ARGENTINA", "12"),
    "AM": ("ARMENIA, REPUBLICA DA", "13"),
    "AW": ("ARUBA", "14"),
    "AU": ("AUSTRALIA", "15"),
    "AT": ("AUSTRIA", "16"),
    "AZ": ("AZERBAIJAO, REPUBLICA DO", "17"),
    "BS": ("BAHAMAS, ILHAS", "18"),
    "BH": ("BAHREIN, ILHAS", "19"),
    "BD": ("BANGLADESH", "20"),
    "BB": ("BARBADOS", "21"),
    "BY": ("Belarus, Republica da", "22"),
    "BE": ("BELGICA", "23"),
    "BZ": ("BELIZE", "24"),
    "BJ": ("BENIN", "25"),
    "BM": ("BERMUDAS", "26"),
    "BO": ("Bolivia, Estado Plurinacional da", "27"),
    "BA": ("Bosnia-Herzegovina, Republica da", "28"),
    "BW": ("BOTSUANA", "29"),
    "BR": ("BRASIL", "30"),
    "BN": ("BRUNEI", "31"),
    "BG": ("BULGARIA, REPUBLICA DA", "32"),
    "BF": ("BURKINA FASO", "33"),
    "BI": ("BURUNDI", "34"),
    "BT": ("BUTAO", "35"),
    "CV": ("CABO VERDE, REPUBLICA DE", "36"),
    "CM": ("CAMAROES", "37"),
    "KH": ("CAMBOJA", "38"),
    "CA": ("CANADA", "39"),
    "GG": ("GUERNSEY, ILHA DO CANAL (INCLUI ALDERNEY E SARK)", "40"),
    "IC": ("CANARIAS, ILHAS", "41"),
    "QA": ("CATAR", "42"),
    "KY": ("CAYMAN, ILHA", "43"),
    "KZ": ("CAZAQUISTAO, REPUBLICA DO", "44"),
    "TD": ("CHADE", "45"),
    "CL": ("CHILE", "46"),
    "CN": ("CHINA, REPUBLICA POPULAR DA", "47"),
    "CY": ("CHIPRE", "48"),
    "CX": ("CHRISTMAS, ILHA (NAVIDAD)", "49"),
    "SG": ("Singapura", "50"),
    "CC": ("COCOS (KEELING), ILHAS", "51"),
    "CO": ("COLOMBIA", "52"),
    "KM": ("COMORES, ILHAS", "53"),
    "CD": ("CONGO, REPUBLICA DEMOCRATICA DO", "54"),
    "CG": ("CONGO, REPUBLICA DO", "55"),
    "CK": ("COOK, ILHA", "56"),
    "KP": ("Coreia (do Norte), Rep. Pop. Democratica da", "57"),
    "KR": ("Coreia (do Sul), Republica da", "58"),
    "CI": ("COSTA DO MARFIM", "59"),
    "CR": ("COSTA RICA", "60"),
    "KW": ("KUWAIT", "61"),
    "HR": ("CROACIA, REPUBLICA DA", "62"),
    "CU": ("CUBA", "63"),
    "DK": ("DINAMARCA", "64"),
    "DJ": ("DJIBUTI", "65"),
    "DM": ("DOMINICA, ILHA", "66"),
    "EG": ("EGITO", "67"),
    "SV": ("EL SALVADOR", "68"),
    "AE": ("EMIRADOS ARABES UNIDOS", "69"),
    "EC": ("EQUADOR", "70"),
    "ER": ("ERITREIA", "71"),
    "GB-SCT": ("ESCOCIA", "72"),
    "SK": ("ESLOVACA, REPUBLICA", "73"),
    "SI": ("ESLOVENIA, REPUBLICA DA", "74"),
    "ES": ("ESPANHA", "75"),
    "US": ("ESTADOS UNIDOS", "76"),
    "EE": ("ESTONIA, REPUBLICA DA", "77"),
    "ET": ("ETIOPIA", "78"),
    "FK": ("FALKLAND (ILHAS MALVINAS)", "79"),
    "FO": ("FEROE, ILHAS", "80"),
    "FJ": ("FIJI", "81"),
    "PH": ("FILIPINAS", "82"),
    "FI": ("FINLANDIA", "83"),
    "TW": ("FORMOSA (TAIWAN)", "84"),
    "FR": ("FRANCA", "85"),
    "GA": ("GABAO", "86"),
    "GB-WLS": ("GALES, PAIS DE", "87"),
    "GM": ("GAMBIA", "88"),
    "GH": ("GANA", "89"),
    "GE": ("GEORGIA, REPUBLICA DA", "90"),
    "GI": ("GIBRALTAR", "91"),
    "GB": ("GRA-BRETANHA", "92"),
    "GD": ("GRANADA", "93"),
    "GR": ("GRECIA", "94"),
    "GL": ("GROENLANDIA", "95"),
    "GP": ("GUADALUPE", "96"),
    "GU": ("GUAM", "97"),
    "GT": ("GUATEMALA", "98"),
    "GY": ("GUIANA", "99"),
    "GF": ("GUIANA FRANCESA", "100"),
    "GN": ("GUINE", "101"),
    "GW": ("GUINE-BISSAU", "102"),
    "GQ": ("GUINE-EQUATORIAL", "103"),
    "HT": ("HAITI", "104"),
    "NL": ("Paises Baixos (Holanda)", "105"),
    "HN": ("HONDURAS", "106"),
    "HK": ("HONG KONG, REGIAO ADM. ESPECIAL", "107"),
    "HU": ("HUNGRIA, REPUBLICA DA", "108"),
    "YE": ("IEMEN", "109"),
    "IN": ("INDIA", "110"),
    "ID": ("INDONESIA", "111"),
    "GB-ENG": ("INGLATERRA", "112"),
    "IR": ("IRA, REPUBLICA ISLAMICA DO", "113"),
    "IQ": ("IRAQUE", "114"),
    "IE": ("IRLANDA", "115"),
    "GB-NIR": ("IRLANDA DO NORTE", "116"),
    "IS": ("ISLANDIA", "117"),
    "IL": ("ISRAEL", "118"),
    "IT": ("ITALIA", "119"),
    "RS": ("SERVIA", "120"),
    "JM": ("JAMAICA", "121"),
    "JP": ("JAPAO", "122"),
    "UM-67": ("JOHNSTON, ILHAS", "123"),
    "JO": ("JORDANIA", "124"),
    "KI": ("KIRIBATI", "125"),
    "LA": ("LAOS, REP. POP. DEMOCRATICA DO", "126"),
    "BN-LB": ("LEBUAN", "127"),
    "LS": ("LESOTO", "128"),
    "LV": ("LETONIA, REPUBLICA DA", "129"),
    "LB": ("LIBANO", "130"),
    "LR": ("LIBERIA", "131"),
    "LY": ("LIBIA", "132"),
    "LI": ("LIECHTENSTEIN", "133"),
    "LT": ("LITUANIA, REPUBLICA DA", "134"),
    "LU": ("LUXEMBURGO", "135"),
    "MO": ("MACAU", "136"),
    "MK": ("MACEDONIA DO NORTE", "137"),
    "MG": ("MADAGASCAR", "138"),
    "PT-30": ("MADEIRA, ILHA DA", "139"),
    "MY": ("MALASIA", "140"),
    "MW": ("MALAVI", "141"),
    "MV": ("MALDIVAS", "142"),
    "ML": ("MALI", "143"),
    "MT": ("MALTA", "144"),
    "IM": ("MAN, ILHAS", "145"),
    "MP": ("MARIANAS DO NORTE", "146"),
    "MA": ("MARROCOS", "147"),
    "MH": ("MARSHALL, ILHAS", "148"),
    "MQ": ("MARTINICA", "149"),
    "MU": ("MAURICIO", "150"),
    "MR": ("MAURITANIA", "151"),
    "MX": ("MEXICO", "152"),
    "MM": ("MIANMAR (BIRMANIA)", "153"),
    "FM": ("MICRONESIA", "154"),
    "UM-71": ("MIDWAY, ILHAS", "155"),
    "MZ": ("MOCAMBIQUE", "156"),
    "MD": ("MOLDAVIA, REPUBLICA DA", "157"),
    "MC": ("MONACO", "158"),
    "MN": ("MONGOLIA", "159"),
    "MS": ("MONTSERRAT, ILHA", "160"),
    "NA": ("NAMIBIA", "161"),
    "NR": ("NAURU", "162"),
    "NP": ("NEPAL", "163"),
    "NI": ("NICARAGUA", "164"),
    "NE": ("NIGER", "165"),
    "NG": ("NIGERIA", "166"),
    "NU": ("NIUE, ILHA", "167"),
    "NF": ("NORFOLK, ILHA", "168"),
    "NO": ("NORUEGA", "169"),
    "NC": ("NOVA CALEDONIA", "170"),
    "NZ": ("NOVA ZELANDIA", "171"),
    "OM": ("OMA", "172"),
    "PW": ("PALAU", "173"),
    "PA": ("PANAMA", "174"),
    "PG": ("PAPUA NOVA GUINE", "175"),
    "PK": ("PAQUISTAO", "176"),
    "PY": ("PARAGUAI", "177"),
    "PE": ("PERU", "178"),
    "PN": ("PITCAIRN, ILHA", "179"),
    "PF": ("POLINESIA FRANCESA", "180"),
    "PL": ("POLONIA, REPUBLICA DA", "181"),
    "PR": ("PORTO RICO", "182"),
    "PT": ("PORTUGAL", "183"),
    "KE": ("QUENIA", "184"),
    "KG": ("QUIRGUIZ, REPUBLICA", "185"),
    "UK": ("REINO UNIDO", "186"),
    "CF": ("REPUBLICA CENTRO-AFRICANA", "187"),
    "DO": ("REPUBLICA DOMINICANA", "188"),
    "RE": ("REUNIAO, ILHA", "189"),
    "RO": ("ROMENIA", "190"),
    "RW": ("RUANDA", "191"),
    "RU": ("Russia, Federacao da", "192"),
    "EH": ("SAARA OCIDENTAL", "193"),
    "SB": ("SALOMAO, ILHAS", "194"),
    "WS": ("SAMOA", "195"),
    "AS": ("SAMOA AMERICANA", "196"),
    "SM": ("San Marino", "197"),
    "SH": ("SANTA HELENA", "198"),
    "LC": ("SANTA LUCIA", "199"),
    "KN": ("SAO CRISTOVAO E NEVES", "200"),
    "PM": ("SAO PEDRO E MIQUELON", "201"),
    "ST": ("SAO TOME E PRINCIPE, ILHAS", "202"),
    "VC": ("SAO VICENTE E GRANADINA", "203"),
    "SN": ("SENEGAL", "204"),
    "SL": ("SERRA LEOA", "205"),
    "SC": ("SEYCHELLE", "206"),
    "SY": ("SIRIA, REPUBLICA ARABE DA", "207"),
    "SO": ("SOMALIA", "208"),
    "LK": ("SRI LANKA", "209"),
    "SZ": ("eSwatini (Essuatini, Suazilândia)", "210"),
    "SD": ("SUDAO", "211"),
    "SE": ("SUECIA", "212"),
    "CH": ("SUICA", "213"),
    "SR": ("SURINAME", "214"),
    "TJ": ("TADJIQUISTAO", "215"),
    "TH": ("TAILANDIA", "216"),
    "TZ": ("TANZANIA, REPUBLICA UNIDA DA", "217"),
    "CZ": ("TCHECA, REPUBLICA", "218"),
    "IO": ("TERRITORIO BRITANICO OC. INDICO", "219"),
    "TL": ("TIMOR LESTE", "220"),
    "TG": ("TOGO", "221"),
    "TO": ("TONGA", "222"),
    "TK": ("TOQUELAU, ILHAS", "223"),
    "TT": ("TRINIDAD E TOBAGO", "224"),
    "TN": ("TUNISIA", "225"),
    "TC": ("TURCAS E CAICOS, ILHAS", "226"),
    "TM": ("TURCOMENISTAO, REPUBLICA DO", "227"),
    "TR": ("TURQUIA", "228"),
    "TV": ("TUVALU", "229"),
    "UA": ("UCRANIA", "230"),
    "UG": ("UGANDA", "231"),
    "UY": ("URUGUAI", "232"),
    "UZ": ("UZBEQUISTAO, REPUBLICA DO", "233"),
    "VU": ("VANUATU", "234"),
    "VA": ("VATICANO, ESTADO DA CIDADE DO", "235"),
    "VE": ("VENEZUELA", "236"),
    "VN": ("VIETNA", "237"),
    "VG": ("VIRGENS, ILHAS (BRITANICAS)", "238"),
    "VI": ("VIRGENS, ILHAS (E.U.A.)", "239"),
    "UM-79": ("WAKE, ILHA", "240"),
    "WF": ("WALLIS E FUTUNA, ILHAS", "241"),
    "ZM": ("ZAMBIA", "242"),
    "ZW": ("ZIMBABUE", "243"),
    "PA-CZ": ("ZONA DO CANAL DO PANAMA", "244"),
    "ME": ("MONTENEGRO", "245"),
    "XX": ("EXTERIOR", "246"),
    "UM": ("Pacifico, Ilhas do (Possessao dos EUA)", "248"),
    "QA2": ("QATAR", "249"),
    "KN2": ("SAINT KITTS E NEVIS", "250"),
    "CS": ("SERVIA E MONTENEGRO", "251"),
    "AX": ("ALAND, ILHAS", "252"),
    "AQ": ("ANTARTICA", "253"),
    "BQ": ("Bonaire, Saint Eustatius e Saba", "254"),
    "BV": ("BOUVET, ILHA", "255"),
    "CW": ("CURACAO", "256"),
    "HM": ("Heard e Ilhas McDonald, Ilha", "257"),
    "MF": ("Sao Martinho, Ilha de (Parte Francesa)", "258"),
    "GS": ("Georgia do Sul e Sandwich do Sul, Ilhas", "259"),
    "JE": ("Jersey. Ilha do Canal", "260"),
    "YT": ("Mayotte", "261"),
    "BL": ("Sao Bartolomeu", "262"),
    "SJ": ("Svalbard e Jan Mayen", "263"),
    "TF": ("Terras Austrais Francesas", "264"),
    "SX": ("SAO MARTINHO, ILHA DE (PARTE HOLANDESA)", "265"),
    "PS": ("Palestina", "266"),
    "SS": ("Sudao do Sul", "267"),
    "GG2": ("Guernsey, Ilha do Canal", "268"),
    "XB": ("Bancos Centrais", "269"),
    "XO": ("Organizacoes Internacionais", "270"),
    "XF": ("FEZZAN", "271"),
    "XD": ("DUBAI", "272"),
    "XP": ("DELEGACAO ESPECIAL DA PALESTINA", "273"),
}

_NOME_EN_PARA_ISO: dict = {
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
    "Ivory Coast": "CI", "Cote d'Ivoire": "CI",
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
    "Dominican Republic": "DO", "Reunion": "RE",
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
    "Turkey": "TR", "Turkiye": "TR", "Tuvalu": "TV",
    "Ukraine": "UA", "Uganda": "UG", "Uruguay": "UY",
    "Uzbekistan": "UZ", "Vanuatu": "VU",
    "Vatican City": "VA", "Venezuela": "VE",
    "Vietnam": "VN", "Viet Nam": "VN",
    "British Virgin Islands": "VG",
    "United States Virgin Islands": "VI",
    "Wallis and Futuna": "WF", "Zambia": "ZM",
    "Zimbabwe": "ZW", "Montenegro": "ME",
    "Curacao": "CW", "Mayotte": "YT",
    "South Sudan": "SS", "Palestine": "PS",
    "State of Palestine": "PS", "Jersey": "JE",
    "Svalbard and Jan Mayen": "SJ",
}

# ════════════════════════════════════════════════════════════════════════════
# TABELA DE ACUMULADORES
# Fonte: Acumuladores.xlsx
# Chave: int(PAULISTANA) -> (codigo_acumulador_str, descricao)
# ════════════════════════════════════════════════════════════════════════════
ACUMULADORES_RAW = [
    (2104,2660,"ANALISE E DESENV DE SISTEMA"),(2105,2668,"SERVICOS DE PROGRAMACAO"),
    (2106,2684,"PROC, ARMAZ OU HOSP DE DADOS"),(2107,2692,"ELABORACAO DE PROGRAMAS DE C"),
    (2108,2800,"LICENCIAMENTO OU CESSAO DE D"),(2109,2881,"ASSESSORIA E CONSULT EM INFO"),
    (2110,2919,"SUPORTE TECNICO EM INFORMATI"),(2111,2935,"PLANEJ CONFEC MANUT E PAG EL"),
    (2112,2961,"DISPON, S"),(2113,2963,"DISPON, S"),(2114,2962,"DISPONIBILIZACAO, SEM CESSAO"),
    (2115,3085,"SERV PESQ E DESENVOL Q NATUR"),(2116,7765,"CESS DIR USO MARC SINAIS PRO"),
    (2117,7773,"EXPL SAL FEST CT CNV ESC VIR"),(2118,7774,"EXPLORACAO DE STANDS E CENTR"),
    (2119,7790,"LOC SUBLOC ARR DI PRASS PERM"),(2120,7803,"CESSAO AND PALC COBERT ESTR"),
    (2121,4030,"MEDICINA E BIOMEDICINA"),(2122,4111,"MEDICINA E BIOMED (R. SPE"),
    (2123,4154,"ANALISES CLIN PATOL (R. SPE"),(2124,5576,"PATOLOGIA E ELETRICIDADE MED"),
    (2125,4139,"ANALISES CLINICAS."),(2126,4140,"RADIOTERAPIA, QUIMIOTERAPIA,"),
    (2127,4170,"LABORATORIOS"),(2128,4189,"HOSPITAIS"),(2129,4197,"CLINICAS E CASAS DE SAUDE"),
    (2130,4219,"AMBULATORIOS E PRONTOS SOCOR"),(2131,4235,"SANATORIOS MANICOMIOS E CONG"),
    (2132,4251,"INSTRUMENTACAO CIRURGICA"),(2133,4260,"ACUPUNTURA"),
    (2134,4316,"ENFERMAGEM INCLUS SERV AUXIL"),(2135,4359,"ENFER INCLUS SERV AUX (R. SP"),
    (2136,4383,"SERV FARMACEUTICOS"),(2137,4391,"FISIOTERAPIA"),
    (2138,4430,"FISIOTERAPIA (REG SPEC"),(2139,4472,"FONOAUDIOLOGIA"),
    (2140,4502,"FONOAUDIOLOGIA (REG SPEC"),(2141,4510,"TERAPIA OCUPACIONAL"),
    (2142,4553,"TERAPIA OCUPACIONAL (R SPE"),(2143,4588,"TERAPIA Q ESP DESTIN TRAT FI"),
    (2144,4626,"NUTRICAO"),(2145,4634,"OBSTETRICIA"),(2146,4677,"OBSTETRICIA (REG SPEC"),
    (2147,4693,"ODONTOLOGIA"),(2148,4731,"ODONTOLOGIA (REG SPEC"),
    (2149,4774,"ORTOPTICA"),(2150,4901,"ORTOPTICA (REG SPEC"),
    (2151,5037,"PROTESES SOB ENCOMENDA"),(2152,5096,"PROTESES SOB ENC (R. SPE"),
    (2153,5100,"PSICANALISE"),(2154,5118,"PSICOLOGIA"),
    (2155,5142,"PSICOL CLINICA OU NAO (R. SP"),(2156,5150,"CASAS DE REPOUSO"),
    (2157,5177,"CRECHES"),(2158,5185,"ASILOS"),(2159,5584,"CASAS DE RECUPERACAO"),
    (2160,5193,"INSEMINACAO ARTIF FERTIL IN"),(2161,5223,"BANC D SANGUE LEITE PELE OLH"),
    (2162,5231,"COL SANGUE LEITE TEC SEMEN O"),(2163,5266,"UNID ATEND ASSIST TRAT MOVEL"),
    (2164,5274,"PLANOS MED GRUPO INDIV E CON"),(2165,5312,"OUT PLANOS SDE Q CUMP SERV T"),
    (2166,5380,"MEDICINA VETERINARIA E ZOOTE"),(2167,5410,"MED VET E ZOOTEC (R. SPE"),
    (2168,5428,"HOSP CLIN AMB PTO"),(2169,5436,"LAB DE ANALISE NA AREA VETER"),
    (2170,5460,"INSEM ART FERTIL IN VITRO AR"),(2171,5479,"BANCOS SANGUE ORGAOS AREA VE"),
    (2172,5495,"COL SAN LEIT TECI SEM ORG AR"),(2173,5517,"UNID ATEND ASSIST TRAT M ARE"),
    (2174,8648,"GUARD TRAT AMES RELAT ANIMAI"),(2175,5533,"PLAN ATEND E ASSIST MED"),
    (2176,8494,"BARBEARIA CABELEIREIROS MANI"),(2177,8516,"ESTETICISTAS TRAT DE PELE DE"),
    (2178,8532,"BANHOS DUCHAS SAUNA MASSAGEN"),(2179,5657,"GINAST DANCA ESP NAT ART MAR"),
    (2180,8567,"CENTROS DE EMAGRECIMENTO SPA"),(2181,8658,"APLICACAO DE TATUAGENS, PIE"),
    (2182,1210,"PAISAGISMO"),(2183,1520,"ENGEN AGRON ARQU URBANISMO"),
    (2184,1546,"ENGEN AGRO ARQU URBA (R. ESP"),(2185,1589,"AGRIMENSURA GEOLOGIA"),
    (2186,1627,"AGRIMENS GEOL (R. ESP"),(2187,1015,"EXEC P"),(2188,1023,"EXEC P"),
    (2189,1024,"EXECUCAO, POR ADMINISTRACAO,"),(2190,1113,"EXECUCAO, POR ADMINISTRACAO,"),
    (2191,1114,"EXECUCAO, POR ADMINISTRACAO,"),(2192,1694,"ELAB PLAN DIRET EST VIAB EST"),
    (2193,1031,"DEMOLICAO"),(2194,1032,"DEMOLICAO. EXCLUSIVO PARA O"),
    (2195,1115,"DEMOLICAO, NO CASO DE SERVIC"),(2196,1058,"REPAR CONSERV REF EDIF ESTRA"),
    (2197,1059,"REPARACAO, CONSERVACAO E REF"),(2198,1116,"REPARACAO, CONSERVACAO E REF"),
    (2199,1228,"COLOC INSTAL TAPETES CARPETE"),(2200,1236,"RECUP RASP POLIM LUSTR PISOS"),
    (2201,1244,"CALAFETACAO"),(2202,1325,"VARRICIO COLETA REM INCINE"),
    (2203,1384,"LIMP MANUT CONS VIAS E LOGRA"),(2204,1406,"LIMP MANUT CONS IMOVEIS"),
    (2205,1430,"DECORACAO"),(2206,1449,"JARD INCLUS CORTE E PODA"),
    (2207,1724,"CONTROL E TRAT EFLU Q NATUR"),(2208,1465,"DEDET DESINF DESINSET IMUNIZ"),
    (2209,1741,"FLORESTAMENTO, REFLORESTAMEN"),(2210,1090,"ESCOR CONTENCAO DE ENCOSTAS"),
    (2211,1117,"ESCORAMENTO, CONTENCAO DE EN"),(2212,1473,"LIMP DRAG RIOS PORTOS CANAIS"),
    (2213,1805,"ACOMP E FISCAL EXEC OBRAS"),(2214,1119,"ACOMPANHAMENTO E FISCALIZACA"),
    (2215,1821,"AEROFOTOGRAMET (INC. INTER)"),(2216,1121,"AEROFOTOGRAMETRIA (INCLUSIVE"),
    (2217,1864,"PESQ PERF CIMENT MERG PERFIL"),(2218,1872,"NUCLEACAO E BOMBARD DE NUVEN"),
    (2219,5673,"ENSI REG PRE FUNDL MEDIO INC"),(2220,5690,"ENS SUP CURS GRAD DEM CURS S"),
    (2221,5711,"ENS SUP CURS POS MESTR DOUT"),(2222,5738,"AUTO"),
    (2223,5762,"OUTR SERV INSTR TREIN PEDAG"),(2224,7005,"HOSP HOTEIS E HOTELARIA MARI"),
    (2225,7013,"HOSP PENS ALB POUS HOSP OCUP"),(2226,7056,"HOSP EM MOTEIS"),
    (2227,7099,"HOSP APRT SERV COND FLAT APR"),(2228,7123,"ORG PROM EXEC DE TURISMO PAS"),
    (2229,7109,"AGENCIAMENTO E INTERMEDIACAO"),(2230,7137,"GUIAS DE TURISMO"),
    (2231,6050,"AGENC CORRET INTERM PLN D P"),(2232,6076,"AGENC CORRET OU INTERM DE C"),
    (2233,6084,"AGENC OU INTERM DE SEGUROS"),(2234,6092,"AGENC CORRET INTERM CART CR"),
    (2235,6114,"AGENC CORRET INTERM PLN SAU"),(2236,6130,"CORRET DE SEGUROS"),
    (2237,6157,"AGENC CORRET INTERM TIT GRL"),(2238,6173,"AGENC CORRET INTERM DIR PRO"),
    (2239,6190,"AGENC CORRET INTERM CONT AR"),(2240,6238,"AGENC CORRET INTERM CONT FA"),
    (2241,6221,"AGENCIAMENTO, CORRETAGEM OU"),(2242,6270,"AGENC CORR INTER BENS BLS M"),
    (2243,6302,"INTERM, PLATAFORM DIGIT, EN"),(2244,6299,"INTERMEDIACAO, VIA PLATAFOR"),
    (2245,6301,"INTERMEDIACAO, VIA PLATAFOR"),(2246,6303,"INTERMEDIACAO, VIA PLATAFOR"),
    (2247,6335,"AGENC MARITIMO"),(2248,6351,"AGENC DE NOTICIAS"),
    (2249,6394,"AGENC PUBLICIDADE E PROPAGA"),(2250,6009,"REPRESENT Q NAT INCLUS COME"),
    (2251,6041,"DISTRIB DE BENS DE TERCEIRO"),(2252,7811,"GUARD ESTAC VEIC TERR AUT"),
    (2253,7838,"GUARD ESTAC VEIC TERR AUTO"),(2254,7846,"GUARD ESTAC VEIC AUTOM VAL"),
    (2255,7854,"GUARDA E ESTAC AERON EMBARC"),(2256,7870,"VIGIL SEGUR MONIT BENS E PE"),
    (2257,7897,"ESCOLTA INCLUS DE VEIC E CA"),(2258,7927,"ARMAZ DEPOS CARGA DESC ARRU"),
    (2259,8052,"ESPETACULOS TEATRAIS"),(2260,8274,"ESPET TEATRAIS E ESPETAC CI"),
    (2261,8079,"EXIBICOES CINEMATOGRAFICAS"),(2262,8087,"ESPETACULOS CIRCENSES"),
    (2263,8095,"PROGRAMAS DE AUDITORIO"),(2264,8117,"PARQUES DIVERSOES CENTROS L"),
    (2265,8257,"PQ DVS CT LAZER N ESTAB N M"),(2266,8273,"PREST SERV DVS PUBL N ESTAB"),
    (2267,8125,"BOATES TAXI DANCING NIGHT C"),(2268,8133,"SHOWS BAILES DESFILES FESTI"),
    (2269,8168,"OPERAS BALLET DANCAS CONC R"),(2270,8176,"FEIRAS EXPOSICOES CONGRESSO"),
    (2271,8311,"BILHAR POR TEMPO (SNOOKER)"),(2272,8320,"BOLICHE"),
    (2273,8338,"DIVERT ELETR INCLUS COMP GA"),(2274,8354,"DIVERT ELETR INCLUS VITRO A"),
    (2275,8362,"MAQ ELETR PROGR DISTRIB PRE"),(2276,8370,"MAQ ELETR PROGR DISTRIB PRE"),
    (2277,8397,"CARTEADO DOMINO VISPORA C C"),(2278,8192,"CORRIDAS E COMPETICOES DE A"),
    (2279,8206,"COMPET ESPORT DESTR FISICA"),(2280,8281,"COMP ESP GD PR BRL F1"),
    (2281,8290,"COMP ESP GD PR BRL F1 N EST"),(2282,8400,"EXEC DE MUSICA INDIVIDE OU"),
    (2283,6777,"PROD MEDI S ENC PREVIA D EV"),(2284,8419,"FORN MUSICA PARA AMBI FECHA"),
    (2285,8214,"DESFILES BLC CARNAVAL FOLC"),(2286,8230,"EXIB FILMES ENTREVISTAS MU"),
    (2287,7218,"RECREA ANIMA INCL FESTAS EV"),(2288,6794,"FONOGRAFIA OU GRAVACAO DE S"),
    (2289,6808,"FOTOGRAFIA E CINEMATOGRAFIA"),(2290,6817,"REPROGRAFIA, MICROFILMAGEM"),
    (2291,6912,"ART GRAF TIPO DIAGR PAG GRA"),(2292,6940,"COMP GRAF INCL IMPRES GRAF"),
    (2293,7331,"LUSTRACAO DE BENS MOVEIS"),(2294,7366,"LUB LAV LIMP N AUTO VEIC EX"),
    (2295,7390,"LUBRIF LAV LIMP VEIC INCL A"),(2296,7412,"LUBR LAV LIMP AUT VEIC EXCT"),
    (2297,7439,"LUBRIF LIMP REV MAQS AP EQU"),(2298,7447,"CARGA RECARGA APARELHOS EQU"),
    (2299,7455,"CONS REST MANU CONS PNT VC"),(2300,7471,"CONS REST MANUT CONS PINT D"),
    (2301,7498,"CONS REST MANUT E CONSERV D"),(2302,7510,"BLINDAGEM"),
    (2303,1880,"ASSISTENCIA TECNICA"),(2304,7552,"RETIF RECOND MOTOR (EXCT PC"),
    (2305,7560,"RECAUC REGENERACAO PNEUS BO"),(2306,7579,"REST RECOND ACOND PINTURA B"),
    (2307,7285,"INST MONT AP MAQS EQU PRES"),(2308,7315,"INSTAL E MONT IND PREST USU"),
    (2309,6831,"COLOCACAO DE MOLDURAS"),(2310,6858,"ENCAD GRAV DOURA LIVR VER"),
    (2311,7595,"ALFAIAT COST MAT FORN USU F"),(2312,7617,"TINTURARIA E LAVANDERIA"),
    (2313,7641,"TAPEC REF ESTOFAMENTOS GERA"),(2314,7676,"FUNILARIA LANTERNAGEM INCLU"),
    (2315,1104,"CARPINTARIA E SERRALHERIA"),(2316,1118,"CARPINTARIA E SERRALHERIA,"),
    (2317,7324,"GUINCHO INTRAMUNICIPAL, GUI"),(2318,5771,"ADM DE FUNDOS QUAISQUER"),
    (2319,5800,"ORG E ADM DE CONSORCIOS"),(2320,5820,"ADM DE CARTAO DE CREDITO DE"),
    (2321,5836,"ADM DE CARTEIRA DE CLIENTES"),(2322,5837,"ADM DE CHEQUES PRE"),
    (2323,5878,"ABERTURA DE CONTAS EM GERAL"),(2324,5870,"LOC MANUT COFRES PART TER E"),
    (2325,5871,"FORN EMISS ATEST EM GERAL"),(2326,5872,"CAD ELAB DE FICHA CAD RENOV"),
    (2327,5873,"EMISS REEM FORN AVS COMPRO"),(2328,5874,"ACESS MOV ATEND CONS CONTAS"),
    (2329,5875,"EMISS REEM ALT CESSAO SUBST"),(2330,5851,"ARRENDAMENTO MERCANTIL (LEA"),
    (2331,5876,"SERV REL COBR RECEB GERAL T"),(2332,5877,"SERV REL PAGTOS GERAL D TIT"),
    (2333,5895,"SERVICOS RELACIONADOS A PAG"),(2334,5879,"DEVOL TIT PROTESTO SUSTACA"),
    (2335,5888,"CUST GRL TIT VAL MOB BMFS B"),(2336,5889,"CUSTORIA EM GERAL INCLUS DE"),
    (2337,5881,"SERV REL OPER CAMBIO EM GER"),(2338,5887,"FORN EMISS REEM RENOV MANUT"),
    (2339,5890,"COMPS TIT BOL VAL MERC FUT"),(2340,5891,"COMPS DE CHEQUES E TIT"),
    (2341,5892,"EMI REEM LIQU CANC BX BLS B"),(2342,5893,"EMISS REEM LIQU ALT CANCEL"),
    (2343,5885,"EMISS FORN DEVOL SUST CANC"),(2344,5886,"SERV REL CRED IMOB AVAL VIS"),
    (2345,2330,"TRANSP P ONIBUS (CONCES E P"),(2346,2340,"SERVICOS DE TRANSPORTE COLE"),
    (2347,2431,"TRANSP PESSOAS QQR MEIO NO"),(2348,2366,"TRANSPORTE POR TAXI, EXPLOR"),
    (2349,2404,"TRANSPORTE DE ESCOLARES."),(2350,2447,"TRANSPORTE DE BENS OU VALOR"),
    (2351,3093,"ANALISE EXAME PESQ COLETA C"),(2352,3115,"ASSES CONS Q NAT N CONT OT"),
    (2353,3123,"TRADUCAO E INTERPRETACAO"),(2354,3158,"DATILOG DIGIT ESTENO SECRET"),
    (2355,3159,"RESPOSTA AUDIVEL (CENTRAIS"),(2356,1899,"PLANEJ COORD PROGR ORGAN TE"),
    (2357,6475,"RECR AGENC SELEC COLOC MAO"),(2358,6491,"FORN MAO D OBRA CARATER TEM"),
    (2359,2496,"PROPAGANDA E PUBLICIDADE"),(2360,6522,"FRANQUIA (FRANCHISING)."),
    (2361,1902,"PERICIAS LAUDOS EXAMES TECN"),(2362,1903,"INSPECAO AMBIENTAL VEICULAR"),
    (2363,7161,"PLANEJ ORG E ADM DE FEIRAS"),(2364,7196,"ORG DE FESTAS E RECEPCOES;"),
    (2365,3204,"ADM GERAL INCLUS BENS NEG T"),(2366,5894,"ADM DE DISTRIB DE CO"),
    (2367,3205,"FORNEC. ADMINS. DE V.A., V."),(2368,3210,"ADMINISTRACAO DE BENEFICIOS"),
    (2369,3213,"ADMINISTRACAO DE IMOVEIS RE"),(2370,6530,"LEILAO E CONGENERES"),
    (2371,3220,"ADVOCACIA"),(2372,3379,"ADVOCACIA (REG SPEC"),
    (2373,3387,"ARBITRAGEM Q ESPECIE INCLUS"),(2374,3395,"AUDITORIA"),
    (2375,3433,"AUDITORIA (REG SPEC"),(2376,2038,"ANALISE DE ORG E METODOS"),
    (2377,3450,"ATUARIA E CALCULOS TECNICOS"),(2378,3476,"CONTABIL INCLUS SERV TECN E"),
    (2379,3620,"CONTADOR TEC CONT (R.ESP"),(2380,3654,"CONSULTORIA E ASSESSORIA EC"),
    (2381,3700,"ECONOMISTAS (REG SPEC"),(2382,3719,"ESTATISTICA"),
    (2383,6564,"COBR REC CON D TERC PROT TI"),(2384,3743,"ASSES ANALI AVAL CONS (FACT"),
    (2385,3751,"APRES PALESTRAS CONF SEMINA"),(2386,2498,"INSERCAO DE TEXTOS, DESENHO"),
    (2387,5916,"SERV REGU SINIST VINC CONT"),(2388,8478,"DISTRIB E VENDA CARTELAS SO"),
    (2389,8486,"SERV DISTRIB E VENDA DE BIL"),(2390,7951,"SERV PORT FERR UTIL PORT MO"),
    (2391,7960,"SERV AEROP UTILIZ AEROP MOV"),(2392,7978,"SERV DE TERMI RODOV FERROV"),
    (2393,3878,"AUTENT DOC RECONHECIMENTO F"),(2394,1481,"SERV EXPLOR RODOV PEDAGIO"),
    (2395,2054,"DESENHO INDUSTRIAL"),(2396,2501,"SERVICOS DE PROGRAMACAO VIS"),
    (2397,6963,"SERV CHAV CONF CARIM PLC SI"),(2398,6572,"FUNE INCL FORN CAIXAO URNA"),
    (2399,6599,"CREMACAO DE CORPOS PART D C"),(2400,6581,"TRASLADO INTRAMUNICIPAL DE"),
    (2401,6602,"PLANOS OU CONVENIO FUNERARI"),(2402,6610,"MANUT E CONSERV DE JAZIGOS"),
    (2403,6613,"CESSAO DE USO DE ESPACOS EM"),(2404,2453,"SERV COLETA REM OU ENT CORR"),
    (2405,2461,"SERV COLETA REM OU ENT COUR"),(2406,2097,"SERV DE ASSISTENCIA SOCIAL"),
    (2407,2119,"SERV AVAL DE BENS SERV. Q N"),(2408,3956,"SERV DE BIBLIOTECONOMIA"),
    (2409,2143,"SERV DE BIOLOGIA BIOTECNOLO"),(2410,2151,"SERV TECNICOS EDIF ELETRON"),
    (2411,2186,"SERV DE DESENHOS TECNICOS"),(2412,6637,"SERV DE DESEMBARACO ADUANEI"),
    (2413,8672,"SERV DE INVESTIG PART DETET"),(2414,2534,"SERV D REPORT ASSESS D IMPR"),
    (2415,2216,"SERV DE METEOROLOGIA"),(2416,8842,"SERV ARTI ATL MODELOS E MAN"),
    (2417,2224,"SERV DE MUSEOLOGIA"),(2418,8885,"SERV DE OURIVESARIA E LAPID"),
    (2419,8893,"SERV RELAT OBRAS DE ARTE SO"),
]

ACUMULADOR_POR_PAULISTANA: dict = {
    paulistana: (str(cod_acum), descricao)
    for cod_acum, paulistana, descricao in ACUMULADORES_RAW
}

ACUMULADOR_IMPORTACAO      = "2551"
ACUMULADOR_IMPORTACAO_DESC = "SERVICOS TOMADOS IMPORTACAO"

# ════════════════════════════════════════════════════════════════════════════
# CONSTANTES DE REGRAS
# ════════════════════════════════════════════════════════════════════════════
CFOP_DENTRO_SP   = "1933"
CFOP_FORA_SP     = "2933"
UF_TOMADOR       = "SP"
ESPECIE_NACIONAL = "03"
ESPECIE_EXTERIOR = "39"


# ════════════════════════════════════════════════════════════════════════════
# HELPERS GERAIS
# ════════════════════════════════════════════════════════════════════════════

def safe(row, col: str, default: str = "") -> str:
    try:
        val = row[col]
        if pd.isna(val):
            return default
        return str(val).strip()
    except (KeyError, TypeError):
        return default


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


def normaliza_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Remove acentos e normaliza nomes de colunas."""
    _ACENTO = str.maketrans(
        "çãéêáâíóúõüôèàÇÃÉÊÁÂÍÓÚÕÜÔÈÀºª",
        "caeeaaioouuoeaCAEEAAIOUOUOEAoa  ",
    )
    df.columns = [c.translate(_ACENTO).strip() for c in df.columns]
    return df


def gera_csv_dominio(registros: list) -> str:
    buf = io.StringIO()
    writer = csv.writer(
        buf, delimiter="|", lineterminator="\n",
        quoting=csv.QUOTE_NONE, escapechar="\\",
    )
    for reg in registros:
        writer.writerow(reg)
    return buf.getvalue()


def extrai_cnpj_tomador(df: pd.DataFrame) -> str:
    """
    Extrai o CNPJ do tomador diretamente do CSV (coluna CPF/CNPJ do Tomador).
    Pega o primeiro valor nao nulo das linhas de Tipo de Registro = 4.
    """
    try:
        df_notas = df[df["Tipo de Registro"].astype(str).str.strip() == "4"]
        for _, row in df_notas.iterrows():
            val = safe(row, "CPF/CNPJ do Tomador")
            cnpj = limpa_cnpj(val)
            if cnpj:
                return cnpj
    except Exception:
        pass
    return ""


# ════════════════════════════════════════════════════════════════════════════
# REGRAS AUTOMATICAS
# ════════════════════════════════════════════════════════════════════════════

def determina_cfop(row) -> str:
    """1933 se prestador SP nacional, 2933 caso contrario."""
    ind      = safe(row, "Indicador de CPF/CNPJ do Prestador")
    uf_prest = safe(row, "UF do Prestador").strip().upper()
    if ind == "3":
        return CFOP_FORA_SP
    if not uf_prest:
        return CFOP_FORA_SP
    return CFOP_DENTRO_SP if uf_prest == UF_TOMADOR else CFOP_FORA_SP


def determina_especie(row) -> str:
    """39 para exterior (Indicador=3), 03 para nacional."""
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")
    return ESPECIE_EXTERIOR if ind == "3" else ESPECIE_NACIONAL


def determina_acumulador(row) -> tuple:
    """
    Retorna (codigo: str, descricao: str, status: str).
    status: 'importacao' | 'ok' | 'aviso' | 'erro'
    """
    ind = safe(row, "Indicador de CPF/CNPJ do Prestador")

    # Exterior -> acumulador fixo de importacao
    if ind == "3":
        return (ACUMULADOR_IMPORTACAO, ACUMULADOR_IMPORTACAO_DESC, "importacao")

    # Lookup pelo codigo paulistana
    cod_serv_raw = safe(row, "Codigo do Servico Prestado na NFTS")
    if not cod_serv_raw or cod_serv_raw in ("nan", ""):
        return ("", "SEM CODIGO DE SERVICO", "aviso")

    try:
        cod_serv_int = int(float(cod_serv_raw))
    except (ValueError, TypeError):
        return ("", f"CODIGO INVALIDO: {cod_serv_raw}", "erro")

    if cod_serv_int in ACUMULADOR_POR_PAULISTANA:
        cod_acum, descricao = ACUMULADOR_POR_PAULISTANA[cod_serv_int]
        return (cod_acum, descricao, "ok")

    return ("", f"PAULISTANA {cod_serv_int} NAO MAPEADA", "aviso")


# ════════════════════════════════════════════════════════════════════════════
# GEOCODER
# ════════════════════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False, ttl=3600)
def detectar_pais_por_endereco(endereco: str) -> dict:
    resultado = {
        "iso2": "", "nome_en": "", "nome_dominio": "",
        "cod_dominio": "", "endereco_completo": "",
        "confianca": "baixa", "erro": None,
    }
    if not endereco or not endereco.strip():
        resultado["erro"] = "Endereco vazio."
        return resultado
    try:
        geolocator = Nominatim(
            user_agent="nfts_dominio_conversor/3.4", timeout=10,
        )
        time.sleep(1.1)
        location = geolocator.geocode(
            endereco, language="en", addressdetails=True, exactly_one=True,
        )
        if location is None:
            resultado["erro"] = f"Endereco nao encontrado: '{endereco}'"
            return resultado

        raw     = location.raw.get("address", {})
        nome_en = raw.get("country", "")
        iso2    = raw.get("country_code", "").upper()

        resultado["nome_en"]           = nome_en
        resultado["iso2"]              = iso2
        resultado["endereco_completo"] = location.address

        nome_dom, cod = TABELA_PAISES.get(iso2, ("", ""))
        if not cod:
            iso2 = _NOME_EN_PARA_ISO.get(nome_en, "")
            resultado["iso2"] = iso2
            nome_dom, cod = TABELA_PAISES.get(iso2, ("", ""))

        resultado["nome_dominio"] = nome_dom
        resultado["cod_dominio"]  = cod
        resultado["confianca"]    = (
            "alta" if (iso2 and nome_en and cod)
            else "media" if cod else "baixa"
        )
        if not cod:
            resultado["erro"] = (
                f"Pais '{nome_en}' (ISO: {iso2}) nao encontrado "
                "na tabela Dominio. Selecione manualmente."
            )
    except GeocoderTimedOut:
        resultado["erro"] = "Timeout ao consultar o geocoder. Tente novamente."
    except GeocoderServiceError as e:
        resultado["erro"] = f"Erro no servico de geocoding: {e}"
    except Exception as e:
        resultado["erro"] = f"Erro inesperado: {e}"
    return resultado


# ════════════════════════════════════════════════════════════════════════════
# MAPEADORES DE REGISTRO
# ════════════════════════════════════════════════════════════════════════════

def reg_0000(cnpj_tomador: str) -> list:
    """
    Registro 0000 - Identificacao da empresa.
    CNPJ lido do campo CPF/CNPJ do Tomador do CSV.
    """
    return ["0000", limpa_cnpj(cnpj_tomador)]


def reg_0020(row, cod_pais: str = "") -> list:
    """Registro 0020 - Cadastro de fornecedores (prestador)."""
    ind  = safe(row, "Indicador de CPF/CNPJ do Prestador")
    cnpj = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))
    if ind == "3" or not cnpj:
        cnpj = "Outros"

    razao    = safe(row, "Razao Social do Prestador")[:150]
    endereco = safe(row, "Endereco do Prestador")
    numero   = re.sub(r"\D", "", safe(row, "Numero do Endereco do Prestador"))
    compl    = safe(row, "Complemento do Endereco do Prestador")
    bairro   = safe(row, "Bairro do Prestador")
    uf       = safe(row, "UF do Prestador")
    cep      = re.sub(r"\D", "", safe(row, "CEP do Prestador"))
    email    = safe(row, "Email do Prestador")
    insc_mun = limpa_im(safe(row, "Inscricao Municipal do Prestador"))

    pais_campo = ""
    if ind == "3":
        uf         = "EX"
        pais_campo = cod_pais

    return [
        "0020",     # 1  - Identificacao do registro (fixo)
        cnpj,       # 2  - Inscricao CNPJ/CPF/CEI/CAEPF ou Outros
        razao,      # 3  - Razao Social (max 150)
        "",         # 4  - Apelido (max 40)
        endereco,   # 5  - Endereco
        numero,     # 6  - Numero do endereco (numerico)
        compl,      # 7  - Complemento
        bairro,     # 8  - Bairro
        "",         # 9  - Codigo do municipio
        uf,         # 10 - UF (EX para exterior)
        pais_campo, # 11 - Codigo do Pais (interno Dominio, so exterior)
        cep,        # 12 - CEP
        "",         # 13 - Inscricao Estadual
        insc_mun,   # 14 - Inscricao Municipal
        "",         # 15 - Inscricao Suframa
        "",         # 16 - DDD
        "",         # 17 - Telefone
        "",         # 18 - FAX
        "",         # 19 - Data do cadastro (dd/mm/aaaa)
        "",         # 20 - Conta contabil
        "",         # 21 - Conta contabil cliente
        "N",        # 22 - Agropecuario (S/N)
        "7",        # 23 - Natureza juridica (7=Empresa Privada)
        "N",        # 24 - Regime de apuracao (N=Normal)
        "N",        # 25 - Contribuinte ICMS (S/N)
        "",         # 26 - Aliquota ICMS
        "",         # 27 - Categoria do estabelecimento
        "",         # 28 - Inscricao Estadual ST
        email,      # 29 - Email
        "N",        # 30 - Interdependencia (S/N)
        "N",        # 31 - Contribuinte CPRB (S/N)
        "",         # 32 - Processo administrativo/judicial (max 21)
        "",         # 33 - Tipo Inscricao (1=CAEPF)
    ]


def reg_1000(row) -> list:
    """
    Registro 1000 - Nota Fiscal de Entrada.
    Todos os campos automaticos:
      Campo 2  (Especie):    03=Nacional / 39=Exterior
      Campo 5  (Acumulador): lookup PAULISTANA ou 2551
      Campo 6  (CFOP):       1933=SP / 2933=outros/exterior
      Campo 20 (ISS Ret):    18=Retido / vazio=Normal
    """
    especie    = determina_especie(row)
    cfop       = determina_cfop(row)
    acum, _, _ = determina_acumulador(row)

    ind  = safe(row, "Indicador de CPF/CNPJ do Prestador")
    cnpj = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))
    if ind == "3" or not cnpj:
        cnpj = "Outros"

    num_doc = re.sub(r"\.0$", "", safe(row, "Numero do Documento"))
    serie   = safe(row, "Serie do Documento")
    if serie in ("-", "nan", ""):
        serie = ""

    data_entrada = formata_data(safe(row, "Data da Prestacao de Servicos"))
    data_emissao = formata_data(safe(row, "Data Hora Emissao NFTS"))
    valor_cont   = limpa_valor(safe(row, "Valor dos Servicos"))

    iss_retido     = safe(row, "ISS Retido").strip().upper()
    cod_recolh_iss = "18" if iss_retido == "S" else ""

    insc_mun = limpa_im(safe(row, "Inscricao Municipal do Prestador"))

    return [
        "1000",          # 1  - Identificacao do registro (fixo)
        especie,         # 2  - Codigo da especie (03=Nacional / 39=Exterior)
        cnpj,            # 3  - Inscricao fornecedor (CNPJ/CPF/CEI/Outros/CAEPF)
        "",              # 4  - Codigo de Exclusao da DIEF
        acum,            # 5  - Codigo do acumulador (automatico)
        cfop,            # 6  - CFOP (1933=dentro SP / 2933=fora SP ou exterior)
        "",              # 7  - Segmento
        num_doc,         # 8  - Numero do documento
        serie,           # 9  - Serie
        "",              # 10 - Numero do documento final
        data_entrada,    # 11 - Data da entrada (dd/mm/aaaa)
        data_emissao,    # 12 - Data emissao (dd/mm/aaaa)
        valor_cont,      # 13 - Valor contabil (2 decimais)
        "",              # 14 - Valor da exclusao da DIEF
        "",              # 15 - Observacao
        "S",             # 16 - Modalidade do frete (S=Sem frete)
        "T",             # 17 - Emitente (T=Terceiros)
        "",              # 18 - CFOP estendido (so SE)
        "",              # 19 - Codigo transferencia credito (so RS)
        cod_recolh_iss,  # 20 - Codigo Recolhimento ISS (18=Retido / vazio=Normal)
        "",              # 21 - Codigo Recolhimento IRRF
        "",              # 22 - Codigo da observacao
        "",              # 23 - Data do visto (so MG, dd/mm/aaaa)
        "E",             # 24 - Fato gerador CRF (E=Emissao / P=Pagamento)
        "E",             # 25 - Fato gerador IRRF (E=Emissao / P=Pagamento)
        "",              # 26 - Valor do frete (2 decimais)
        "",              # 27 - Valor do seguro (2 decimais)
        "",              # 28 - Valor das despesas (2 decimais)
        "",              # 29 - Valor do PIS (2 decimais)
        "",              # 30 - Codigo Antecipacao Tributaria
        "",              # 31 - Valor do COFINS (2 decimais)
        "",              # 32 - Valor DARE (so SE, 2 decimais)
        "",              # 33 - Aliquota DARE (so SE, 2 decimais)
        "",              # 34 - Valor base calculo ICMS ST
        "",              # 35 - Entradas isentas (so MG, 2 decimais)
        "",              # 36 - Outras entradas isentas (so MG, 2 decimais)
        "",              # 37 - Valor transporte incluido na base (so MG, 2 decimais)
        "",              # 38 - Codigo de ressarcimento
        valor_cont,      # 39 - Valor produtos (2 decimais)
        "",              # 40 - Municipio Origem (cMunIni)
        "0",             # 41 - Situacao da Nota (0=Documento Regular)
        "",              # 42 - Codigo da situacao tributaria
        "",              # 43 - Sub serie
        "",              # 44 - Inscricao estadual do fornecedor
        insc_mun,        # 45 - Inscricao municipal do fornecedor
        "",              # 46 - Codigo da operacao e prestacao
        "",              # 47 - Valor a ser deduzido da receita tributavel (2 decimais)
        data_entrada,    # 48 - Competencia (dd/mm/aaaa)
        "",              # 49 - Operacao (so PA)
        "",              # 50 - Numero do parecer fiscal
        "",              # 51 - Data do parecer fiscal (dd/mm/aaaa)
        "",              # 52 - Numero da declaracao de Importacao
        "N",             # 53 - Possui beneficio fiscal (S/N)
        "",              # 54 - Chave da nota fiscal eletronica
        "",              # 55 - Codigo de recolhimento FETHAB
        "",              # 56 - Responsavel pelo recolhimento FETHAB (E/C)
        "",              # 57 - CFOP documento fiscal
        "",              # 58 - Tipo de CT-e (0=Normal/1=Complemento/2=Anulacao)
        "",              # 59 - CT-e referencia
        "",              # 60 - Modalidade da importacao (1/2/3/4/5)
        "",              # 61 - Codigo da informacao complementar
        "",              # 62 - Informacao complementar
        "",              # 63 - Classe de consumo
        "",              # 64 - Tipo de ligacao
        "",              # 65 - Grupo de tensao
        "",              # 66 - Tipo de assinante
        "",              # 67 - KWH consumido
        "",              # 68 - Valor fornecido/consumido gas ou energia (2 decimais)
        "",              # 69 - Valor cobrado de terceiros (2 decimais)
        "",              # 70 - Tipo do documento de importacao (10=DI/1=DSI)
        "",              # 71 - Numero do Ato Concessorio Drawback
        "",              # 72 - Natureza do frete PIS/COFINS
        "",              # 73 - CST PIS/COFINS
        "",              # 74 - Base do credito PIS/COFINS
        "",              # 75 - Valor servicos/itens PIS/COFINS (2 decimais)
        "",              # 76 - Base de calculo PIS/COFINS (2 decimais)
        "",              # 77 - Aliquota de PIS (2 decimais)
        "",              # 78 - Aliquota de COFINS (2 decimais)
        "",              # 79 - Chave de NFSe
        "",              # 80 - Numero do processo ou ato concessorio
        "",              # 81 - Origem do processo (1=Jus.Fed./3=SRF/9=Outros)
        "",              # 82 - Data da escrituracao (dd/mm/aaaa)
        "",              # 83 - CFPS (so DF)
        "",              # 84 - Natureza da receita PIS/COFINS
        "",              # 85 - CST IPI
        "",              # 86 - Lancamentos de SCP
        "",              # 87 - Tipo de servico (1=Transp.Cargas/2=Transp.Pass.)
        "",              # 88 - Municipio destino
        "",              # 89 - Pedagio (2 decimais)
        "",              # 90 - IPI (2 decimais)
        "",              # 91 - ICMS ST (2 decimais)
        "",              # 92 - EFD-Reinf Tipo de servico
        "",              # 93 - EFD-Reinf Indicativo Prestacao (0/1/2)
        "",              # 94 - Numero doc. arrecadacao (so RS, max 255)
        "",              # 95 - Tipo do titulo (0=Dup/1=Cheq/2=Prom/3=Rec/99=Out)
        "",              # 96 - Identificacao (max 255)
        "",              # 97 - ICMS Desonerado (2 decimais)
        "",              # 98 - IPI Devolucao (2 decimais)
    ]


def reg_1020(row) -> list:
    """
    Registro 1020 - Impostos da Nota Fiscal de Entrada.

    ISS Retido = S -> Cod 18 | Base=ValorServ | Aliq=Aliquota | Valor=ValorISS | Outras=vazio
    ISS Retido = N -> Cod 3  | Base=0.00 | Aliq=0.00 | Valor=0.00 | Outras=ValorServ
    Campo 11 (Valor Contabil) = sempre ValorServicos
    """
    iss_retido     = safe(row, "ISS Retido").strip().upper()
    valor_servicos = limpa_valor(safe(row, "Valor dos Servicos"))
    aliquota       = limpa_aliquota(safe(row, "Aliquota"))
    valor_iss      = limpa_valor(safe(row, "Valor ISS"))

    if iss_retido == "S":
        cod_imposto = "18"
        base        = valor_servicos
        aliq        = aliquota
        valor       = valor_iss
        outras      = ""
    else:
        cod_imposto = "3"
        base        = "0.00"
        aliq        = "0.00"
        valor       = "0.00"
        outras      = valor_servicos

    return [
        "1020",         # 1  - Identificacao do registro (fixo)
        cod_imposto,    # 2  - Codigo do imposto (18=ISS Retido / 3=ISS)
        "",             # 3  - Percentual de reducao da base de calculo (2 decimais)
        base,           # 4  - Base de calculo (2 decimais)
        aliq,           # 5  - Aliquota (2 casas decimais; 3 para 39-IRRFP)
        valor,          # 6  - Valor do Imposto (2 decimais)
        "",             # 7  - Valor de Isentas (2 decimais)
        outras,         # 8  - Valor de Outras (2 decimais) - ValorServ se ISS Normal
        "",             # 9  - Valor do IPI (2 decimais)
        "",             # 10 - Valor da substituicao Tributaria (2 decimais)
        valor_servicos, # 11 - Valor Contabil (2 decimais) - sempre ValorServicos
        "",             # 12 - Codigo do recolhimento do imposto
        "",             # 13 - Valor nao tributadas (so GO, 2 decimais)
        "",             # 14 - Valor parcela reduzida (so GO, 2 decimais)
        "",             # 15 - Aliq. Interestadual
        "",             # 16 - Nat. rend. (max 5 chars)
        "",             # 17 - Tipo de Deducao (max 1 char, so 63-IRRF-APF)
        "",             # 18 - Tipo de Isencao (max 2 chars, so 63-IRRF-APF)
        "",             # 19 - Descricao (max 100 chars)
    ]


def reg_1150(row) -> list:
    """Registro 1150 - IBS. Filho do 1000. Gerado em branco."""
    return ["1150", "", "", "", ""]


def reg_1151(row) -> list:
    """Registro 1151 - CBS. Filho do 1000. Gerado em branco."""
    return ["1151", "", "", "", ""]


# ════════════════════════════════════════════════════════════════════════════
# CONVERSOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

def converte_nfts(
    df: pd.DataFrame,
    cnpj_tomador: str,
    gerar_ibs_cbs: bool,
    apenas_iss_retido: bool,
    auto_geocode: bool,
    pais_manual_override: str,
) -> tuple:
    """
    Converte DataFrame da NFTS para registros Dominio.
    Retorna (csv_string, df_preview, erros_list).
    """
    registros = []
    preview   = []
    erros     = []

    # Registro 0000 com CNPJ lido do CSV
    registros.append(reg_0000(cnpj_tomador))

    df_notas = df[df["Tipo de Registro"].astype(str).str.strip() == "4"].copy()
    if df_notas.empty:
        erros.append("Nenhum registro do tipo '4' (nota) encontrado no arquivo.")
        return "", pd.DataFrame(), erros

    if apenas_iss_retido:
        df_notas = df_notas[
            df_notas["ISS Retido"].astype(str).str.strip().str.upper() == "S"
        ]
        if df_notas.empty:
            erros.append("Nenhuma nota com ISS Retido = S encontrada.")
            return "", pd.DataFrame(), erros

    fornecedores_gerados: set = set()

    for idx, row in df_notas.iterrows():
        try:
            ind      = safe(row, "Indicador de CPF/CNPJ do Prestador")
            cnpj     = limpa_cnpj(safe(row, "CPF/CNPJ do Prestador"))
            razao    = safe(row, "Razao Social do Prestador")
            uf_prest = safe(row, "UF do Prestador").strip().upper()

            chave_forn = (
                f"EXT_{razao[:50]}" if (ind == "3" or not cnpj) else cnpj
            )

            # Determina automaticos
            cfop_nota    = determina_cfop(row)
            especie_nota = determina_especie(row)
            acum_cod, acum_desc, acum_status = determina_acumulador(row)

            if acum_status in ("erro", "aviso") and acum_cod == "":
                erros.append(
                    f"Linha {idx} ({razao}): Acumulador nao encontrado - {acum_desc}"
                )

            # Detecta pais para prestadores estrangeiros
            cod_pais_ext  = ""
            nome_pais_ext = ""
            geo_info      = ""

            if ind == "3":
                if pais_manual_override:
                    cod_pais_ext  = pais_manual_override
                    nome_pais_ext = next(
                        (n for iso, (n, c) in TABELA_PAISES.items()
                         if c == pais_manual_override), ""
                    )
                    geo_info = f"Manual: {nome_pais_ext} (cod. {cod_pais_ext})"
                elif auto_geocode:
                    end_parts = [
                        safe(row, "Endereco do Prestador"),
                        safe(row, "Numero do Endereco do Prestador"),
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
                            geo_info = f"Aviso: {geo['erro']}"
                            erros.append(f"Linha {idx} ({razao}): {geo['erro']}")
                        else:
                            geo_info = f"OSM: {nome_pais_ext} (cod. {cod_pais_ext})"

            # Registro 0020 - uma vez por fornecedor
            if chave_forn not in fornecedores_gerados:
                registros.append(reg_0020(row, cod_pais=cod_pais_ext))
                fornecedores_gerados.add(chave_forn)

            # Registros 1000 e 1020
            registros.append(reg_1000(row))
            registros.append(reg_1020(row))

            # Registros 1150/1151 opcionais
            if gerar_ibs_cbs:
                registros.append(reg_1150(row))
                registros.append(reg_1151(row))

            # Preview
            iss_ret        = safe(row, "ISS Retido").upper()
            valor_servicos = safe(row, "Valor dos Servicos")
            cod_serv       = safe(row, "Codigo do Servico Prestado na NFTS")

            cfop_desc = (
                f"{cfop_nota} - dentro SP"
                if cfop_nota == CFOP_DENTRO_SP
                else f"{cfop_nota} - fora SP/EXT"
            )
            especie_desc = (
                f"{especie_nota} - Exterior"
                if especie_nota == ESPECIE_EXTERIOR
                else f"{especie_nota} - Nacional"
            )
            acum_display = (
                f"{acum_cod} - {acum_desc}"
                if acum_cod
                else f"AVISO: {acum_desc}"
            )

            preview.append({
                "Nr NFTS"        : safe(row, "Nr NFTS"),
                "Prestador"      : razao,
                "CNPJ/CPF"       : safe(row, "CPF/CNPJ do Prestador"),
                "UF Prest."      : uf_prest if uf_prest else "EXT",
                "Cod. Servico"   : cod_serv,
                "Acumulador"     : acum_display,
                "Especie"        : especie_desc,
                "CFOP"           : cfop_desc,
                "Pais Dominio"   : (
                    f"{nome_pais_ext} [{cod_pais_ext}]"
                    if cod_pais_ext else (geo_info if geo_info else "-")
                ),
                "Emissao"        : safe(row, "Data Hora Emissao NFTS"),
                "Prestacao"      : safe(row, "Data da Prestacao de Servicos"),
                "Valor Servicos" : valor_servicos,
                "Aliquota ISS"   : safe(row, "Aliquota"),
                "Valor ISS"      : safe(row, "Valor ISS"),
                "ISS Retido"     : iss_ret,
                "Cod Imposto"    : "18 - ISS Retido" if iss_ret == "S" else "3 - ISS",
                "Base 1020"      : limpa_valor(valor_servicos) if iss_ret == "S" else "0.00",
                "Outras 1020"    : "" if iss_ret == "S" else limpa_valor(valor_servicos),
            })

        except Exception as e:
            erros.append(f"Linha {idx}: {e}")

    csv_out    = gera_csv_dominio(registros)
    df_preview = pd.DataFrame(preview)
    return csv_out, df_preview, erros


# ════════════════════════════════════════════════════════════════════════════
# HELPERS DE UI
# ════════════════════════════════════════════════════════════════════════════

def _opcoes_paises() -> tuple:
    opcoes_iso = sorted(TABELA_PAISES.keys(), key=lambda k: int(TABELA_PAISES[k][1]))
    opcoes_label = [
        f"{TABELA_PAISES[iso][1]:>3} - {TABELA_PAISES[iso][0]}"
        for iso in opcoes_iso
    ]
    return opcoes_iso, opcoes_label


# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — apenas opcoes que o usuario controla
# ════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.header("Configuracoes")

    # Resumo das regras automaticas (informativo)
    st.info(
        "Campos determinados automaticamente pelo CSV:\n\n"
        "**CNPJ Empresa** — CPF/CNPJ do Tomador\n\n"
        "**CFOP** — 1933 (SP) ou 2933 (fora/EXT)\n\n"
        "**Especie** — 03 (nacional) ou 39 (exterior)\n\n"
        "**Acumulador** — lookup PAULISTANA ou 2551 (importacao)"
    )

    st.subheader("Opcoes de Geracao")
    apenas_iss_retido = st.checkbox(
        "Apenas notas com ISS Retido", value=False,
        help="Filtra somente notas onde ISS Retido = S.",
    )
    gerar_ibs_cbs = st.checkbox(
        "Gerar registros 1150 (IBS) e 1151 (CBS)", value=False,
        help="Gera os registros em branco. Ative se o Dominio exigir.",
    )

    st.subheader("Prestadores Estrangeiros")
    auto_geocode = st.checkbox(
        "Detectar pais automaticamente (OSM/Nominatim)",
        value=True,
        help="Consulta o OpenStreetMap para identificar o pais pelo endereco.",
    )
    opcoes_iso_sb, opcoes_label_sb = _opcoes_paises()
    sel_override = st.selectbox(
        "Forcar pais para todos os estrangeiros",
        options=["(Automatico / sem override)"] + opcoes_label_sb,
        index=0,
        help="Selecione para forcar um pais especifico para TODOS os prestadores estrangeiros.",
    )
    pais_manual_override = ""
    if sel_override != "(Automatico / sem override)":
        pais_manual_override = sel_override.split(" - ")[0].strip()

    st.divider()
    st.caption("v3.4 - NFTS Paulistana para Dominio Sistemas")


# ════════════════════════════════════════════════════════════════════════════
# TITULO E TABS
# ════════════════════════════════════════════════════════════════════════════

st.title("Conversor NFTS Paulistana para Dominio Sistemas")
st.caption(
    "Converte o CSV da NFS-e Tomadas (NFTS) da Prefeitura de Sao Paulo "
    "para o layout de importacao do Dominio Sistemas — "
    "Registros 0000, 0020, 1000, 1020, 1150, 1151"
)

tab_converter, tab_acumuladores, tab_paises, tab_ajuda = st.tabs(
    ["Converter", "Acumuladores", "Paises", "Ajuda"]
)


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 - CONVERTER
# ════════════════════════════════════════════════════════════════════════════

with tab_converter:
    st.subheader("Upload do arquivo NFTS")
    arquivo = st.file_uploader(
        "Selecione o CSV exportado da NFTS (Prefeitura de SP)",
        type=["csv"],
        help="Ex.: NFTS_50431480_20260601_20260630.csv",
    )

    if arquivo:
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

        df_raw = normaliza_colunas(df_raw)

        # Extrai CNPJ do tomador automaticamente do CSV
        cnpj_tomador = extrai_cnpj_tomador(df_raw)
        razao_tomador = ""
        try:
            df_t4 = df_raw[df_raw["Tipo de Registro"].astype(str).str.strip() == "4"]
            if not df_t4.empty:
                razao_tomador = safe(df_t4.iloc[0], "Razao Social do Tomador")
        except Exception:
            pass

        total_linhas = len(df_raw)
        notas_tipo4  = (df_raw["Tipo de Registro"].astype(str).str.strip() == "4").sum()

        col_ind = "Indicador de CPF/CNPJ do Prestador"
        notas_ext  = 0
        notas_sp   = 0
        notas_fsp  = 0
        notas_s_ac = 0

        if col_ind in df_raw.columns:
            df_t4     = df_raw[df_raw["Tipo de Registro"].astype(str).str.strip() == "4"]
            notas_ext = df_t4[col_ind].astype(str).str.strip().isin(["3", "3.0"]).sum()
            mask_nac  = ~df_t4[col_ind].astype(str).str.strip().isin(["3", "3.0"])
            df_nac    = df_t4[mask_nac]
            if "UF do Prestador" in df_raw.columns:
                notas_sp  = (df_nac["UF do Prestador"].astype(str).str.strip().str.upper() == "SP").sum()
                notas_fsp = notas_tipo4 - notas_sp - notas_ext
            col_serv = "Codigo do Servico Prestado na NFTS"
            if col_serv in df_raw.columns:
                for _, row in df_nac.iterrows():
                    cod_raw = safe(row, col_serv)
                    if cod_raw and cod_raw not in ("nan", ""):
                        try:
                            cod_int = int(float(cod_raw))
                            if cod_int not in ACUMULADOR_POR_PAULISTANA:
                                notas_s_ac += 1
                        except (ValueError, TypeError):
                            notas_s_ac += 1

        # Exibe dados da empresa tomadora detectados
        if cnpj_tomador:
            st.success(
                f"Empresa tomadora detectada no CSV: "
                f"**{razao_tomador}** — CNPJ: **{cnpj_tomador}**"
            )
        else:
            st.error(
                "Nao foi possivel detectar o CNPJ do tomador no CSV. "
                "Verifique se o arquivo contem a coluna 'CPF/CNPJ do Tomador'."
            )

        # Metricas
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Linhas",             total_linhas)
        c2.metric("Notas (Tipo 4)",     notas_tipo4)
        c3.metric("Exterior (Esp.39)",  notas_ext)
        c4.metric("CFOP 1933 (SP)",     notas_sp)
        c5.metric("CFOP 2933 (fora)",   notas_fsp + notas_ext)
        c6.metric("Sem acumulador",     notas_s_ac,
                  delta=None if notas_s_ac == 0 else f"{notas_s_ac} avisos",
                  delta_color="inverse")

        with st.expander("Visualizar dados brutos do CSV"):
            st.dataframe(df_raw, use_container_width=True)

        # Validacoes
        avisos = []
        if not cnpj_tomador:
            avisos.append("CNPJ do tomador nao detectado no CSV.")
        for av in avisos:
            st.warning(av)

        if not avisos:
            if notas_ext > 0 and auto_geocode and not pais_manual_override:
                st.info(
                    f"{notas_ext} nota(s) com prestador estrangeiro. "
                    "O pais sera buscado automaticamente pelo OpenStreetMap."
                )
            if notas_s_ac > 0:
                st.warning(
                    f"{notas_s_ac} nota(s) com codigo de servico nao mapeado. "
                    "O acumulador sera gerado em branco. Veja a aba Acumuladores."
                )

            if st.button("Converter para layout Dominio", type="primary"):
                with st.spinner("Convertendo..."):
                    csv_saida, df_prev, erros = converte_nfts(
                        df=df_raw,
                        cnpj_tomador=cnpj_tomador,
                        gerar_ibs_cbs=gerar_ibs_cbs,
                        apenas_iss_retido=apenas_iss_retido,
                        auto_geocode=auto_geocode,
                        pais_manual_override=pais_manual_override,
                    )

                if erros:
                    with st.expander(f"{len(erros)} aviso(s) durante a conversao"):
                        for e in erros:
                            st.warning(e)

                if not df_prev.empty:
                    st.success(
                        f"Conversao concluida! {len(df_prev)} nota(s) processada(s)."
                    )

                    m1, m2, m3, m4, m5, m6 = st.columns(6)
                    n_ret   = (df_prev["ISS Retido"] == "S").sum()
                    n_nor   = (df_prev["ISS Retido"] == "N").sum()
                    n_1933  = df_prev["CFOP"].str.startswith("1933").sum()
                    n_2933  = df_prev["CFOP"].str.startswith("2933").sum()
                    n_ext   = df_prev["Especie"].str.startswith("39").sum()
                    n_aviso = df_prev["Acumulador"].str.startswith("AVISO").sum()
                    total_val = df_prev["Valor Servicos"].apply(
                        lambda x: float(limpa_valor(x)) if limpa_valor(x) else 0.0
                    ).sum()
                    m1.metric("Total notas",         len(df_prev))
                    m2.metric("ISS Retido (cod.18)", n_ret)
                    m3.metric("ISS Normal (cod.3)",  n_nor)
                    m4.metric("CFOP 1933",           n_1933)
                    m5.metric("CFOP 2933",           n_2933)
                    m6.metric("Acum. nao mapeado",   n_aviso, delta_color="inverse")

                    st.subheader("Preview das notas convertidas")

                    def highlight_row(row):
                        if row["Acumulador"].startswith("AVISO"):
                            cor = "#f8d7da"   # vermelho - acumulador nao encontrado
                        elif row["Especie"].startswith("39"):
                            cor = "#cce5ff"   # azul - exterior
                        elif row["ISS Retido"] == "S":
                            cor = "#d4edda"   # verde - ISS retido
                        else:
                            cor = "#fff3cd"   # amarelo - ISS normal
                        return [f"background-color: {cor}"] * len(row)

                    st.dataframe(
                        df_prev.style.apply(highlight_row, axis=1),
                        use_container_width=True,
                    )

                    # Legenda de cores
                    col_l1, col_l2, col_l3, col_l4 = st.columns(4)
                    col_l1.markdown(
                        "<span style='background:#f8d7da;padding:2px 8px;border-radius:4px'>"
                        "Vermelho = Acumulador nao mapeado</span>",
                        unsafe_allow_html=True,
                    )
                    col_l2.markdown(
                        "<span style='background:#cce5ff;padding:2px 8px;border-radius:4px'>"
                        "Azul = Exterior (Esp.39 / Acum.2551)</span>",
                        unsafe_allow_html=True,
                    )
                    col_l3.markdown(
                        "<span style='background:#d4edda;padding:2px 8px;border-radius:4px'>"
                        "Verde = ISS Retido (cod.18)</span>",
                        unsafe_allow_html=True,
                    )
                    col_l4.markdown(
                        "<span style='background:#fff3cd;padding:2px 8px;border-radius:4px'>"
                        "Amarelo = ISS Normal (cod.3)</span>",
                        unsafe_allow_html=True,
                    )

                    with st.expander("Regras automaticas aplicadas"):
                        st.dataframe(pd.DataFrame([
                            ["CNPJ Empresa",    "CPF/CNPJ do Tomador do CSV",                  cnpj_tomador],
                            ["Especie 03",      "Indicador Prestador = 1 ou 2 (nacional)",     "03 - Nota Fiscal de Servico"],
                            ["Especie 39",      "Indicador Prestador = 3 (exterior)",          "39 - Nota de Servico Exterior"],
                            ["CFOP 1933",       "Prestador nacional com UF = SP",              "Dentro do estado"],
                            ["CFOP 2933",       "Prestador fora de SP ou exterior",            "Fora do estado / exterior"],
                            ["Acumulador",      "Lookup Codigo Servico vs coluna PAULISTANA",  "Codigo automatico"],
                            ["Acumulador 2551", "Prestador estrangeiro (Indicador = 3)",       "SERVICOS TOMADOS IMPORTACAO"],
                            ["ISS Retido",      "ISS Retido = S -> Cod 18 | N -> Cod 3",       "Campo 2 do Reg. 1020"],
                        ], columns=["Campo", "Regra", "Resultado"]),
                        use_container_width=True, hide_index=True)

                    with st.expander("Regra Registro 1020"):
                        st.dataframe(pd.DataFrame([
                            ["ISS Retido (S)", "18", "Valor Servicos", "Aliquota nota", "Valor ISS", "(vazio)", "Valor Servicos"],
                            ["ISS Normal (N)", "3",  "0.00",           "0.00",          "0.00",      "Valor Servicos", "Valor Servicos"],
                        ], columns=["Situacao", "Cod", "Campo 4 Base", "Campo 5 Aliq", "Campo 6 Valor", "Campo 8 Outras", "Campo 11 VCont"]),
                        use_container_width=True, hide_index=True)

                    with st.expander("Previa do arquivo gerado (50 primeiras linhas)"):
                        st.code("\n".join(csv_saida.split("\n")[:50]), language="text")

                    nome_saida = arquivo.name.replace(".csv", "_DOMINIO.txt")
                    st.download_button(
                        label="Baixar arquivo para importacao no Dominio",
                        data=csv_saida.encode("utf-8"),
                        file_name=nome_saida,
                        mime="text/plain",
                        type="primary",
                    )
                    st.info(
                        "Como importar: No Dominio Sistemas acesse "
                        "Utilitarios > Importacao > Notas Fiscais de Entrada "
                        "e selecione o arquivo gerado."
                    )
    else:
        st.info("Faca o upload do arquivo CSV da NFTS para iniciar.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 2 - ACUMULADORES
# ════════════════════════════════════════════════════════════════════════════

with tab_acumuladores:
    st.subheader("Tabela de Acumuladores - Lookup por Codigo Paulistana")
    st.markdown(
        "O sistema confronta o campo **Codigo do Servico Prestado na NFTS** "
        "com a coluna **PAULISTANA** para determinar o acumulador automaticamente."
    )
    st.info(
        "Acumulador especial: **2551 - SERVICOS TOMADOS IMPORTACAO** "
        "aplicado automaticamente para prestadores estrangeiros (Indicador = 3)."
    )

    busca_acum = st.text_input(
        "Filtrar tabela de acumuladores",
        placeholder="Digite codigo, paulistana ou descricao...",
        key="filtro_acumuladores",
    )

    df_acum = pd.DataFrame([
        {"Codigo Acumulador": cod_acum, "Codigo Paulistana": paulistana, "Descricao": descricao}
        for cod_acum, paulistana, descricao in ACUMULADORES_RAW
    ]).sort_values("Codigo Paulistana").reset_index(drop=True)

    if busca_acum:
        mask = (
            df_acum["Codigo Acumulador"].astype(str).str.contains(busca_acum, na=False)
            | df_acum["Codigo Paulistana"].astype(str).str.contains(busca_acum, na=False)
            | df_acum["Descricao"].str.contains(busca_acum, case=False, na=False)
        )
        df_acum = df_acum[mask]

    st.markdown("##### Acumulador de Importacao (fixo para estrangeiros)")
    st.dataframe(pd.DataFrame([{
        "Codigo Acumulador": 2551,
        "Codigo Paulistana": "N/A",
        "Descricao": "SERVICOS TOMADOS IMPORTACAO (Prestador Estrangeiro - Indicador 3)",
    }]), use_container_width=True, hide_index=True)

    st.markdown("##### Acumuladores por Codigo Paulistana")
    st.dataframe(df_acum, use_container_width=True, height=450)
    st.caption(f"Total: {len(df_acum)} acumuladores mapeados")

    st.markdown("---")
    st.markdown("##### Simulador de Lookup")
    col_sim1, col_sim2 = st.columns([2, 3])
    with col_sim1:
        cod_sim = st.number_input(
            "Codigo do Servico Prestado na NFTS",
            min_value=0, max_value=99999, value=2800,
            help="Digite o codigo paulistana para ver qual acumulador sera usado.",
        )
    with col_sim2:
        st.write("")
        st.write("")
        if int(cod_sim) in ACUMULADOR_POR_PAULISTANA:
            acum_sim, desc_sim = ACUMULADOR_POR_PAULISTANA[int(cod_sim)]
            st.success(f"Paulistana **{int(cod_sim)}** -> Acumulador **{acum_sim}** - {desc_sim}")
        else:
            st.error(f"Paulistana **{int(cod_sim)}** nao encontrada na tabela.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 - PAISES
# ════════════════════════════════════════════════════════════════════════════

with tab_paises:
    st.subheader("Consulta de Pais por Endereco")

    col_end, col_btn = st.columns([4, 1])
    with col_end:
        end_consulta = st.text_input(
            "Endereco para busca",
            value="2108 North Street, Sacramento, CA",
            key="geo_manual_end",
        )
    with col_btn:
        st.write("")
        st.write("")
        buscar_pais = st.button("Buscar", key="geo_manual_btn")

    iso_detectado = ""
    if buscar_pais and end_consulta:
        with st.spinner("Consultando OpenStreetMap..."):
            res = detectar_pais_por_endereco(end_consulta)

        if res["cod_dominio"]:
            icone = {"alta": "OK - Alta confianca", "media": "Parcial", "baixa": "Baixa"}.get(
                res["confianca"], "?"
            )
            st.success(
                f"Confianca: {icone}\n\n"
                f"Pais detectado: {res['nome_dominio']}\n\n"
                f"Nome (EN): {res['nome_en']}\n\n"
                f"ISO Alpha-2: {res['iso2']}\n\n"
                f"Codigo Dominio: {res['cod_dominio']}\n\n"
                f"Endereco completo: {res['endereco_completo']}"
            )
            iso_detectado = res["iso2"]
        else:
            st.error(f"Erro: {res['erro']}")

    st.markdown("---")
    st.markdown("##### Selecionar pais manualmente")
    opcoes_iso_t, opcoes_label_t = _opcoes_paises()
    idx_def = 0
    if iso_detectado and iso_detectado in opcoes_iso_t:
        idx_def = opcoes_iso_t.index(iso_detectado)

    sel_manual = st.selectbox(
        "Pais (Tabela Dominio - ordenado por codigo)",
        options=opcoes_label_t, index=idx_def, key="geo_manual_sel",
    )
    if sel_manual:
        partes  = sel_manual.split(" - ", 1)
        cod_sel = partes[0].strip()
        nom_sel = partes[1] if len(partes) > 1 else ""
        st.success(f"Codigo Dominio: {cod_sel} - {nom_sel}")

    st.markdown("---")
    st.markdown("##### Tabela completa de Paises - Dominio Sistemas")
    busca_filtro = st.text_input("Filtrar tabela", placeholder="Nome, codigo ou ISO...", key="filtro_paises")
    df_paises = pd.DataFrame([
        {"Codigo Dominio": int(cod), "ISO Alpha-2": iso, "Nome (Dominio)": nome}
        for iso, (nome, cod) in TABELA_PAISES.items()
    ]).sort_values("Codigo Dominio").reset_index(drop=True)

    if busca_filtro:
        mask = (
            df_paises["Nome (Dominio)"].str.contains(busca_filtro, case=False, na=False)
            | df_paises["Codigo Dominio"].astype(str).str.contains(busca_filtro, na=False)
            | df_paises["ISO Alpha-2"].str.contains(busca_filtro, case=False, na=False)
        )
        df_paises = df_paises[mask]

    st.dataframe(df_paises, use_container_width=True, height=400)
    st.caption(f"Total: {len(df_paises)} paises")


# ════════════════════════════════════════════════════════════════════════════
# TAB 4 - AJUDA
# ════════════════════════════════════════════════════════════════════════════

with tab_ajuda:
    st.subheader("Documentacao e Mapeamento de Campos")

    st.markdown("### Registros gerados")
    st.dataframe(pd.DataFrame([
        ["0000", "Identificacao da empresa (tomadora)", "CPF/CNPJ do Tomador do CSV"],
        ["0020", "Cadastro do fornecedor (prestador) - 1 por CNPJ", "NFTS"],
        ["1000", "Nota Fiscal de Entrada", "NFTS"],
        ["1020", "Impostos (ISS Retido cod.18 ou ISS Normal cod.3)", "NFTS"],
        ["1150", "IBS - Imposto sobre Bens e Servicos (opcional)", "Em branco"],
        ["1151", "CBS - Contribuicao sobre Bens e Servicos (opcional)", "Em branco"],
    ], columns=["Registro", "Descricao", "Fonte"]), use_container_width=True, hide_index=True)

    st.markdown("### Todas as regras automaticas")
    st.dataframe(pd.DataFrame([
        ["CNPJ Empresa (Reg.0000)",  "CPF/CNPJ do Tomador",                    "Lido diretamente do CSV - nao requer configuracao"],
        ["Especie (campo 2)",        "Indicador Prestador = 3",                 "39 - Nota de Servico do Exterior"],
        ["Especie (campo 2)",        "Indicador Prestador = 1 ou 2",            "03 - Nota Fiscal de Servico"],
        ["CFOP (campo 6)",           "Prestador nacional, UF = SP",             "1933 - dentro do estado"],
        ["CFOP (campo 6)",           "Prestador nacional, UF != SP",            "2933 - fora do estado"],
        ["CFOP (campo 6)",           "Prestador estrangeiro (Ind.=3)",          "2933 - exterior"],
        ["Acumulador (campo 5)",     "Indicador = 3 (exterior)",                "2551 - SERVICOS TOMADOS IMPORTACAO"],
        ["Acumulador (campo 5)",     "Lookup Cod.Servico vs PAULISTANA",        "Codigo do acumulador correspondente"],
        ["ISS Retido (campo 20)",    "ISS Retido = S",                          "18 - codigo de recolhimento ISS retido"],
        ["ISS Retido (campo 20)",    "ISS Retido = N",                          "vazio - ISS normal nao retido"],
        ["UF Reg.0020 (campo 10)",   "Prestador estrangeiro",                   "EX"],
        ["Pais Reg.0020 (campo 11)", "Prestador estrangeiro",                   "Codigo interno Dominio (geocoder OSM)"],
    ], columns=["Campo", "Condicao", "Resultado"]), use_container_width=True, hide_index=True)

    st.markdown("### Regra do Registro 1020")
    st.dataframe(pd.DataFrame([
        ["ISS Retido (S)", "18", "Valor Servicos", "Aliquota nota", "Valor ISS", "(vazio)", "Valor Servicos"],
        ["ISS Normal (N)", "3",  "0.00",           "0.00",          "0.00",      "Valor Servicos", "Valor Servicos"],
    ], columns=["Situacao", "Cod", "Campo 4 Base", "Campo 5 Aliq", "Campo 6 Valor", "Campo 8 Outras", "Campo 11 VCont"]),
    use_container_width=True, hide_index=True)

    st.markdown("### Exemplo com dados reais do CSV (MOYMER)")
    st.dataframe(pd.DataFrame([
        ["0000", "20586841000130", "CNPJ lido de CPF/CNPJ do Tomador", "-"],
        ["196/195", "Tenjin INC",  "Ind.=3 / EXT", "39 / 2933 / 2551 / Pais=76(EUA)"],
        ["194",     "ALELO S.A.", "Ind.=2 / SP",  "03 / 1933 / 2237 / ISS Normal cod.3"],
    ], columns=["Nr/Reg", "Empresa/Prestador", "Indicador/UF", "Especie/CFOP/Acum/ISS"]),
    use_container_width=True, hide_index=True)

    st.markdown("### Como importar no Dominio Sistemas")
    st.dataframe(pd.DataFrame([
        ["1", "Faca o upload do CSV da NFTS"],
        ["2", "Confirme os dados da empresa tomadora detectados automaticamente"],
        ["3", "Clique em Converter para layout Dominio"],
        ["4", "Baixe o arquivo .txt gerado"],
        ["5", "No Dominio: Utilitarios > Importacao > Notas Fiscais de Entrada"],
        ["6", "Selecione o arquivo e confirme a importacao"],
    ], columns=["Passo", "Acao"]), use_container_width=True, hide_index=True)

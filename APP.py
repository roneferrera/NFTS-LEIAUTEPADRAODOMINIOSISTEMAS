def reg_1000(row, lookup_acum: dict, lookup_pais: dict) -> str:
    especie    = determina_especie(row)
    fornecedor = determina_fornecedor(row)
    acumulador = determina_acumulador(row, lookup_acum)
    cfop       = determina_cfop(row)
    num_doc    = limpa_numero(safe(row, "Número do Documento"))
    serie      = safe(row, "Série do Documento")
    if serie in ("-", "nan", ""):
        serie = ""

    dt_campo11 = fmt_date(safe(row, "Data Hora Emissão NFTS"))
    dt_campo12 = fmt_date(safe(row, "Data da Prestação de Serviços"))

    valor   = fmt_decimal(safe(row, "Valor dos Serviços"), casas=2)
    cod_iss = determina_cod_iss(row)

    campos = [
        "1000",      # 1  - Identificação do registro (fixo)
        especie,     # 2  - Código da espécie
        fornecedor,  # 3  - Inscrição fornecedor
        "",          # 4  - Código de Exclusão da DIEF
        acumulador,  # 5  - Código do acumulador
        cfop,        # 6  - CFOP
        "",          # 7  - Segmento
        num_doc,     # 8  - Número do documento
        serie,       # 9  - Série
        "",          # 10 - Numero do documento final
        dt_campo11,  # 11 - Data da entrada
        dt_campo12,  # 12 - Data emissão
        valor,       # 13 - Valor contábil ✅
        "",          # 14 - Valor da exclusão da DIEF
        "",          # 15 - Observação
        "",          # 16 - Modalidade do frete
        "",          # 17 - Emitente da nota fiscal
        "",          # 18 - CFOP estendido
        "",          # 19 - Código da transferência de crédito
        cod_iss,     # 20 - Código do Recolhimento do ISS Retido
        "",          # 21 - Código do Recolhimento do IRRF
        "",          # 22 - Código da observação
        "",          # 23 - Data do visto
        "",          # 24 - Fato gerador da CRF
        "",          # 25 - Fato gerador do IRRF
        "",          # 26 - Valor do frete
        "",          # 27 - Valor do seguro
        "",          # 28 - Valor das despesas
        "",          # 29 - Valor do PIS
        "",          # 30 - Código Antecipação Tributária
        "",          # 31 - Valor do COFINS
        "",          # 32 - Valor DARE
        "",          # 33 - Alíquota DARE
        "",          # 34 - Valor base ICMS ST
        "",          # 35 - Entradas cuja saída é isenta
        "",          # 36 - Outras entradas isentas
        "",          # 37 - Valor transporte incluído na base
        "",          # 38 - Código de ressarcimento
        valor,       # 39 - Valor produtos ✅ CORRIGIDO = mesmo que Valor contábil
        "",          # 40 - Município Origem
        "",          # 41 - Situação da Nota
        "",          # 42 - Código da situação tributária
        "",          # 43 - Sub serie
        "",          # 44 - Inscrição estadual do fornecedor
        "",          # 45 - Inscrição municipal do fornecedor
        "",          # 46 - Código da operação e prestação
        "",          # 47 - Valor a ser deduzido da receita tributável
        "",          # 48 - Competência
        "",          # 49 - Operação
        "",          # 50 - Número do parecer fiscal
        "",          # 51 - Data do parecer fiscal
        "",          # 52 - Número da declaração de Importação
        "",          # 53 - Possui benefício fiscal
        "",          # 54 - Chave da nota fiscal eletrônica
        "",          # 55 - Código de recolhimento do FETHAB
        "",          # 56 - Responsável pelo recolhimento do FETHAB
        "",          # 57 - CFOP documento fiscal
        "",          # 58 - Tipo de CT-e
        "",          # 59 - CT-e referência
        "",          # 60 - Modalidade da importação
        "",          # 61 - Código da informação complementar
        "",          # 62 - Informação complementar
        "",          # 63 - Classe de consumo
        "",          # 64 - Tipo de ligação
        "",          # 65 - Grupo de tensão
        "",          # 66 - Tipo de assinante
        "",          # 67 - KWH consumido
        "",          # 68 - Valor fornecido/consumido
        "",          # 69 - Valor cobrado de terceiros
        "",          # 70 - Tipo do documento de importação
        "",          # 71 - Número do Ato Concessório Drawback
        "",          # 72 - Natureza do frete PIS/COFINS
        "",          # 73 - CST PIS/COFINS
        "",          # 74 - Base do crédito PIS/COFINS
        "",          # 75 - Valor serviços/itens PIS/COFINS
        "",          # 76 - Base de cálculo PIS/COFINS
        "",          # 77 - Alíquota de PIS
        "",          # 78 - Alíquota de COFINS
        "",          # 79 - Chave de NFSe
        "",          # 80 - Número do processo ou ato concessório
        "",          # 81 - Origem do processo
        "",          # 82 - Data da escrituração
        "",          # 83 - CFPS
        "",          # 84 - Natureza da receita PIS/COFINS
        "",          # 85 - CST IPI
        "",          # 86 - Lançamentos de SCP
        "",          # 87 - Tipo de serviço
        "",          # 88 - Município destino
        "",          # 89 - Pedágio
        "",          # 90 - IPI
        "",          # 91 - ICMS ST
        "",          # 92 - Classificação EFD-Reinf tipo
        "",          # 93 - Classificação EFD-Reinf indicativo
        "",          # 94 - Número do documento de arrecadação
        "",          # 95 - Tipo do título
        "",          # 96 - Identificação
        "",          # 97 - ICMS Desonerado
        "",          # 98 - IPI Devolução
    ]
    assert len(campos) == 98, f"reg_1000: esperado 98 campos, encontrado {len(campos)}"
    return monta_linha(campos)

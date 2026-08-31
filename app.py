"""
Gerador de Relatório de NF-e/NFC-e com Precificação — v3.0
UFISCAL — Inteligência em Negócios

Novidades da v3.0:
  - Colunas uTrib, qTrib e vUnTrib (logo após vUnitConv)
  - Bloco de precificação após Custo_Total_da_Mercadoria (antecipação de ICMS,
    custo de aquisição, preço ideal de venda, margem bruta e comparação com o
    preço praticado)
  - Excel com fórmulas vivas: o usuário altera margem, % ICMS saída, % PIS/COFINS
    e o preço praticado direto na planilha e tudo recalcula
  - Linhas zebradas, cabeçalhos por bloco, filtros e painéis congelados na coluna NCM
"""

import time
import zipfile
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
import xlsxwriter
from lxml import etree as ET
from xlsxwriter.utility import xl_col_to_name

try:
    import rarfile
    RAR_DISPONIVEL = True
except ImportError:  # rarfile/unrar não instalados no servidor
    rarfile = None
    RAR_DISPONIVEL = False


# =============================================================================
# 1. LEITURA DO XML
# =============================================================================

def strip_namespace(root):
    """Remove o namespace de todos os elementos da árvore XML (lxml)."""
    for el in root.iter():
        if isinstance(el.tag, str) and '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]
    return root


def get_text(node, tag, default=''):
    """Obtém o texto de um nó XML de forma segura."""
    if node is None:
        return default
    child = node.find(tag)
    return child.text if child is not None and child.text is not None else default


def to_float(value):
    """Converte um valor para float de forma segura."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def format_ncm(ncm):
    return f"{ncm[0:4]}.{ncm[4:6]}.{ncm[6:8]}" if len(ncm) >= 8 else ncm


def format_cest(cest):
    if not cest:
        return ''
    cest = cest.zfill(7)
    return f"{cest[0:2]}.{cest[2:5]}.{cest[5:7]}"


def format_cfop(cfop):
    cfop = cfop.zfill(4)
    return f"{cfop[0]}.{cfop[1:4]}"


# Ordem das colunas extraídas do XML
COLS_BASE = [
    'CRT', 'Chave_NFe', 'Numero_NFe', 'Data_Emis', 'nItem', 'cProd', 'xProd', 'NCM', 'CEST', 'CFOP',
    'CST_ICMS', 'vProd', 'vBC', 'pICMS', 'vICMS', 'vFCP', 'vFrete', 'vSeg', 'vDesc', 'vOutro',
    'pMVAST', 'vBCST', 'pICMSST', 'vICMSST', 'vFCPST', 'vICMSDeson', 'Desonerado_Abate', 'vIPI',
    'vBC_IPI', 'pIPI', 'CST_PIS', 'vPIS', 'vCOFINS', 'vII', 'Total_s_IPI', 'Base_PIS_COFINS',
    'Base_COFINS', 'uCom', 'qCom', 'vUnitConv', 'uTrib', 'qTrib', 'vUnTrib',
    'cEANTrib', 'cEAN', 'Comparacao_EAN', 'Custo_Total_da_Mercadoria',
]

# Bloco de precificação (na ordem solicitada)
COLS_CALC = [
    'Antecipacao_ICMS', 'Custo_Aquisicao', 'Custo_Unitario_Aquisicao',
    'Perc_Margem', 'Perc_ICMS_Saida', 'Perc_PIS_COFINS', 'Perc_IRPJ_CSLL', 'Total_Perc',
    'Preco_Ideal_Venda', 'VL_ICMS_Saida', 'VL_PIS_COFINS_Saida', 'VL_IRPJ_CSLL',
    'Margem_Bruta_RS', 'Margem_Bruta_Perc',
    'Preco_Praticado', 'Diferenca_RS', 'Diferenca_Perc', 'Situacao_Preco',
]


def process_xml_streams(xml_streams):
    """Processa uma lista de tuplas (filename, file_stream) -> (DataFrame, log de erros)."""
    rows = []
    error_log = []
    total_files = len(xml_streams)
    progress_bar = st.progress(0, text="Processando XMLs...")

    for i, (filename, xml_stream) in enumerate(xml_streams):
        try:
            xml_stream.seek(0)
            parser = ET.XMLParser(remove_blank_text=True, recover=True)
            tree = ET.parse(xml_stream, parser)
            root = strip_namespace(tree.getroot())

            inf = root.find('.//infNFe')
            if inf is None:
                raise ValueError("Tag <infNFe> não encontrada. O arquivo é uma NF-e válida?")

            num = get_text(inf, 'ide/nNF')
            dh = get_text(inf, 'ide/dhEmi').split('T')[0]
            chave = inf.get('Id')[-44:] if inf.get('Id') else ''
            crt = get_text(inf, 'emit/CRT')

            for prod in root.findall('.//det'):
                imposto = prod.find('imposto')
                if imposto is None:
                    error_log.append({
                        "arquivo": filename,
                        "item": prod.get('nItem'),
                        "erro": "Grupo <imposto> não encontrado no item."
                    })
                    continue

                # ---------------- ICMS ----------------
                icms_node = imposto.find('ICMS')
                icms_group = None
                cst_icms = ''

                if icms_node is not None and len(icms_node) > 0:
                    icms_group = icms_node[0]
                    orig = get_text(icms_group, 'orig')
                    cst = get_text(icms_group, 'CST')
                    csosn = get_text(icms_group, 'CSOSN')
                    cst_icms = orig + (cst if cst else csosn)

                vBC = to_float(get_text(icms_group, 'vBC'))
                pICMS = to_float(get_text(icms_group, 'pICMS')) / 100
                vICMS = to_float(get_text(icms_group, 'vICMS'))
                vFCP = to_float(get_text(icms_group, 'vFCP'))
                pMV = to_float(get_text(icms_group, 'pMVAST')) / 100
                vBCST = to_float(get_text(icms_group, 'vBCST'))
                pICMSST = to_float(get_text(icms_group, 'pICMSST')) / 100
                vICMSST = to_float(get_text(icms_group, 'vICMSST'))
                vFCPST = to_float(get_text(icms_group, 'vFCPST'))
                vICMSDeson = to_float(get_text(icms_group, 'vICMSDeson'))

                # indDeduzDeson: 1 = deduz do total da nota | ausente/0 = não deduz
                ind_deduz = get_text(icms_group, 'indDeduzDeson')
                desonerado_abate = 'SIM' if ind_deduz == '1' else 'NAO'

                # ---------------- Produto ----------------
                vProd = to_float(get_text(prod, 'prod/vProd'))
                vFre = to_float(get_text(prod, 'prod/vFrete'))
                vSeg = to_float(get_text(prod, 'prod/vSeg'))
                vDesc = to_float(get_text(prod, 'prod/vDesc'))
                vOut = to_float(get_text(prod, 'prod/vOutro'))

                # ---------------- IPI ----------------
                ipi_node = imposto.find('IPI')
                vIPI = to_float(get_text(ipi_node, 'IPITrib/vIPI')) if ipi_node is not None else 0.0
                vBC_IPI = to_float(get_text(ipi_node, 'IPITrib/vBC')) if ipi_node is not None else 0.0
                pIPI = to_float(get_text(ipi_node, 'IPITrib/pIPI')) / 100 if ipi_node is not None else 0.0

                # ---------------- PIS / COFINS ----------------
                pis_node = imposto.find('PIS')
                cst_pis = ''
                if pis_node is not None and len(pis_node) > 0:
                    cst_pis = get_text(pis_node[0], 'CST')

                vPIS = to_float(get_text(pis_node, './/vPIS'))
                vCOF = to_float(get_text(imposto, 'COFINS/.//vCOFINS'))
                vII = to_float(get_text(imposto, 'II/vII'))

                # ---------------- Cálculos ----------------
                tot_no_ipi = vProd + vFre + vSeg - vDesc + vOut
                base_pis = tot_no_ipi + vIPI
                base_cof = tot_no_ipi + vICMSST + vFCPST + vIPI

                qCom = to_float(get_text(prod, 'prod/qCom'))
                unit_conv = base_cof / (qCom or 1)

                # Unidade tributável
                uTrib = get_text(prod, 'prod/uTrib')
                qTrib = to_float(get_text(prod, 'prod/qTrib'))
                vUnTrib = to_float(get_text(prod, 'prod/vUnTrib'))

                cEANTrib = get_text(prod, 'prod/cEANTrib')
                cEAN = get_text(prod, 'prod/cEAN')
                cmpEAN = 'Igual' if cEANTrib == cEAN else 'Diferente'

                custo_total = (vProd + vFre + vSeg) - vDesc + (vOut + vICMSST + vFCPST + vIPI)
                if desonerado_abate == 'SIM':
                    custo_total -= vICMSDeson

                rows.append({
                    'Numero_NFe': num, 'Data_Emis': dh, 'Chave_NFe': chave,
                    'nItem': prod.get('nItem'), 'cProd': get_text(prod, 'prod/cProd'),
                    'xProd': get_text(prod, 'prod/xProd'), 'NCM': format_ncm(get_text(prod, 'prod/NCM')),
                    'CEST': format_cest(get_text(prod, 'prod/CEST')), 'CFOP': format_cfop(get_text(prod, 'prod/CFOP')),
                    'CST_ICMS': cst_icms, 'vProd': vProd, 'vBC': vBC, 'pICMS': pICMS,
                    'vICMS': vICMS, 'vFCP': vFCP, 'vFrete': vFre, 'vSeg': vSeg,
                    'vDesc': vDesc, 'vOutro': vOut, 'pMVAST': pMV, 'vBCST': vBCST,
                    'pICMSST': pICMSST, 'vICMSST': vICMSST, 'vFCPST': vFCPST, 'vIPI': vIPI,
                    'Total_s_IPI': tot_no_ipi, 'Custo_Total_da_Mercadoria': custo_total,
                    'Base_PIS_COFINS': base_pis, 'Base_COFINS': base_cof, 'uCom': get_text(prod, 'prod/uCom'),
                    'qCom': qCom, 'vUnitConv': unit_conv,
                    'uTrib': uTrib, 'qTrib': qTrib, 'vUnTrib': vUnTrib,
                    'cEANTrib': cEANTrib, 'cEAN': cEAN, 'Comparacao_EAN': cmpEAN,
                    'vBC_IPI': vBC_IPI, 'pIPI': pIPI, 'CST_PIS': cst_pis, 'CRT': crt,
                    'vICMSDeson': vICMSDeson, 'Desonerado_Abate': desonerado_abate,
                    'vPIS': vPIS, 'vCOFINS': vCOF, 'vII': vII
                })

        except ET.ParseError as e:
            error_log.append({"arquivo": filename, "item": "-", "erro": f"XML mal formatado ou corrompido: {e}"})
        except Exception as e:
            error_log.append({"arquivo": filename, "item": "-", "erro": f"Erro inesperado: {e}"})

        progress_bar.progress((i + 1) / total_files, text=f"Processando XMLs... ({i + 1}/{total_files})")

    progress_bar.empty()
    df = pd.DataFrame(rows)

    if not df.empty:
        cols_existentes = [c for c in COLS_BASE if c in df.columns]
        df = df[cols_existentes]

    return df, error_log


# =============================================================================
# 2. PRECIFICAÇÃO
# =============================================================================

def aplicar_precificacao(df, perc_margem, perc_icms_saida, perc_pis_cofins, perc_irpj_csll):
    """
    Acrescenta o bloco de precificação ao DataFrame.

    Antecipação ICMS ....... (Custo_Total * %ICMS_Saida) - vICMS
    Custo Aquisição ........ Custo_Total - vICMS, se %ICMS_Saida > 0; senão Custo_Total
    Custo Unit. Aquisição .. Custo_Aquisicao / qTrib
    Total % ................ Margem + ICMS Saída + PIS/COFINS + IRPJ/CSLL
    Preço Ideal ............ Custo Unit. Aquisição / (1 - Total %)
    Margem Bruta R$ ........ Preço Ideal - Custo Unit. - ICMS - PIS/COFINS - IRPJ/CSLL
    """
    df = df.copy()

    custo_total = df['Custo_Total_da_Mercadoria']
    v_icms = df['vICMS']
    q_trib = df['qTrib'].replace(0, np.nan)

    df['Perc_Margem'] = perc_margem
    df['Perc_ICMS_Saida'] = perc_icms_saida
    df['Perc_PIS_COFINS'] = perc_pis_cofins
    df['Perc_IRPJ_CSLL'] = perc_irpj_csll

    # Antecipação só existe quando há ICMS de saída e quando o débito supera o
    # crédito de entrada; nos demais casos o campo zera
    if perc_icms_saida > 0:
        df['Antecipacao_ICMS'] = (custo_total * perc_icms_saida - v_icms).clip(lower=0)
    else:
        df['Antecipacao_ICMS'] = 0.0

    if perc_icms_saida > 0:
        df['Custo_Aquisicao'] = custo_total - v_icms
    else:
        df['Custo_Aquisicao'] = custo_total

    df['Custo_Unitario_Aquisicao'] = (df['Custo_Aquisicao'] / q_trib).fillna(0)

    df['Total_Perc'] = perc_margem + perc_icms_saida + perc_pis_cofins + perc_irpj_csll

    divisor = 1 - df['Total_Perc']
    df['Preco_Ideal_Venda'] = np.where(divisor > 0, df['Custo_Unitario_Aquisicao'] / divisor, 0.0)

    df['VL_ICMS_Saida'] = df['Preco_Ideal_Venda'] * perc_icms_saida
    df['VL_PIS_COFINS_Saida'] = df['Preco_Ideal_Venda'] * perc_pis_cofins
    df['VL_IRPJ_CSLL'] = df['Preco_Ideal_Venda'] * perc_irpj_csll

    df['Margem_Bruta_RS'] = (
        df['Preco_Ideal_Venda'] - df['Custo_Unitario_Aquisicao']
        - df['VL_ICMS_Saida'] - df['VL_PIS_COFINS_Saida'] - df['VL_IRPJ_CSLL']
    )
    df['Margem_Bruta_Perc'] = np.where(
        df['Preco_Ideal_Venda'] > 0, df['Margem_Bruta_RS'] / df['Preco_Ideal_Venda'], 0.0
    )

    # Preenchidos manualmente na planilha (fórmulas prontas no Excel)
    df['Preco_Praticado'] = np.nan
    df['Diferenca_RS'] = np.nan
    df['Diferenca_Perc'] = np.nan
    df['Situacao_Preco'] = ''

    return df[[c for c in COLS_BASE + COLS_CALC if c in df.columns]]


# =============================================================================
# 3. GERAÇÃO DO EXCEL
# =============================================================================

FMT_MOEDA = '#,##0.00'
FMT_MOEDA4 = '#,##0.0000'
FMT_PERC = '0.00%'
FMT_QTD = '#,##0.0000'

COLS_PERC = {
    'pICMS', 'pMVAST', 'pICMSST', 'pIPI', 'Perc_Margem', 'Perc_ICMS_Saida',
    'Perc_PIS_COFINS', 'Perc_IRPJ_CSLL', 'Total_Perc', 'Margem_Bruta_Perc', 'Diferenca_Perc',
}
COLS_MOEDA4 = {'vUnitConv', 'vUnTrib', 'Custo_Unitario_Aquisicao', 'Preco_Ideal_Venda', 'Preco_Praticado'}
COLS_QTD = {'qCom', 'qTrib'}
COLS_MOEDA = {
    'vProd', 'vBC', 'vICMS', 'vFCP', 'vFrete', 'vSeg', 'vDesc', 'vOutro', 'vBCST', 'vICMSST',
    'vFCPST', 'vICMSDeson', 'vIPI', 'vBC_IPI', 'vPIS', 'vCOFINS', 'vII', 'Total_s_IPI',
    'Base_PIS_COFINS', 'Base_COFINS', 'Custo_Total_da_Mercadoria', 'Antecipacao_ICMS',
    'Custo_Aquisicao', 'VL_ICMS_Saida', 'VL_PIS_COFINS_Saida', 'VL_IRPJ_CSLL',
    'Margem_Bruta_RS', 'Diferenca_RS',
}
# Colunas de digitação do usuário (destacadas em amarelo)
COLS_EDITAVEIS = {'Perc_Margem', 'Perc_ICMS_Saida', 'Perc_PIS_COFINS', 'Perc_IRPJ_CSLL', 'Preco_Praticado'}

# Cabeçalhos por bloco
GRUPO_IDENT = {'CRT', 'Chave_NFe', 'Numero_NFe', 'Data_Emis', 'nItem', 'cProd', 'xProd',
               'NCM', 'CEST', 'CFOP'}
GRUPO_RESULTADO = {'Preco_Ideal_Venda', 'Margem_Bruta_RS', 'Margem_Bruta_Perc',
                   'Diferenca_RS', 'Diferenca_Perc', 'Situacao_Preco'}

COR_HEADER_IDENT = '#1F3864'
COR_HEADER_FISCAL = '#2E75B6'
COR_HEADER_CALC = '#375623'
COR_HEADER_RESULT = '#7F6000'
COR_HEADER_EDIT = '#BF8F00'
COR_BANDA = ('#FFFFFF', '#EAF1F8')
COR_EDIT = ('#FFF8E1', '#FDF0C9')
COR_CALC_BANDA = ('#F4F9F0', '#E7F2DF')


def _tipo_num(coluna):
    if coluna in COLS_PERC:
        return FMT_PERC
    if coluna in COLS_MOEDA4:
        return FMT_MOEDA4
    if coluna in COLS_QTD:
        return FMT_QTD
    if coluna in COLS_MOEDA:
        return FMT_MOEDA
    return None


def gerar_excel(df):
    """Gera o Excel formatado, com fórmulas vivas no bloco de precificação."""
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})
    ws = wb.add_worksheet('Relatorio')

    colunas = list(df.columns)
    idx = {c: i for i, c in enumerate(colunas)}

    def L(c):
        return xl_col_to_name(idx[c])

    # ---------- Formatos ----------
    cache = {}

    def fmt(coluna, banda):
        key = (coluna, banda)
        if key in cache:
            return cache[key]
        props = {'border': 1, 'border_color': '#BFBFBF', 'valign': 'vcenter'}
        if coluna in COLS_EDITAVEIS:
            props['bg_color'] = COR_EDIT[banda]
            props['bold'] = True
        elif coluna in COLS_CALC:
            props['bg_color'] = COR_CALC_BANDA[banda]
        else:
            props['bg_color'] = COR_BANDA[banda]
        num = _tipo_num(coluna)
        if num:
            props['num_format'] = num
        if coluna in GRUPO_RESULTADO:
            props['bold'] = True
        cache[key] = wb.add_format(props)
        return cache[key]

    def header_fmt(coluna):
        if coluna in COLS_EDITAVEIS:
            cor = COR_HEADER_EDIT
        elif coluna in GRUPO_RESULTADO:
            cor = COR_HEADER_RESULT
        elif coluna in COLS_CALC:
            cor = COR_HEADER_CALC
        elif coluna in GRUPO_IDENT:
            cor = COR_HEADER_IDENT
        else:
            cor = COR_HEADER_FISCAL
        key = ('header', cor)
        if key not in cache:
            cache[key] = wb.add_format({
                'bold': True, 'font_color': '#FFFFFF', 'bg_color': cor,
                'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
                'border': 1, 'border_color': '#FFFFFF',
            })
        return cache[key]

    fmt_neg = wb.add_format({'font_color': '#C00000', 'bold': True})
    fmt_pos = wb.add_format({'font_color': '#1F7A1F', 'bold': True})

    # ---------- Cabeçalho ----------
    ws.set_row(0, 34)
    for c, nome in enumerate(colunas):
        ws.write(0, c, nome, header_fmt(nome))

    # ---------- Fórmulas do bloco de precificação ----------
    ct, vi, qt = L('Custo_Total_da_Mercadoria'), L('vICMS'), L('qTrib')
    pm, pic, ppc, pir = L('Perc_Margem'), L('Perc_ICMS_Saida'), L('Perc_PIS_COFINS'), L('Perc_IRPJ_CSLL')
    tp, ca, cu, pv = L('Total_Perc'), L('Custo_Aquisicao'), L('Custo_Unitario_Aquisicao'), L('Preco_Ideal_Venda')
    vic, vpc, vir = L('VL_ICMS_Saida'), L('VL_PIS_COFINS_Saida'), L('VL_IRPJ_CSLL')
    mb, pp = L('Margem_Bruta_RS'), L('Preco_Praticado')

    formulas = {
        'Antecipacao_ICMS': lambda r: f"=IF({pic}{r}=0,0,MAX(0,{ct}{r}*{pic}{r}-{vi}{r}))",
        'Custo_Aquisicao': lambda r: f"=IF({pic}{r}>0,{ct}{r}-{vi}{r},{ct}{r})",
        'Custo_Unitario_Aquisicao': lambda r: f"=IFERROR({ca}{r}/{qt}{r},0)",
        'Total_Perc': lambda r: f"={pm}{r}+{pic}{r}+{ppc}{r}+{pir}{r}",
        'Preco_Ideal_Venda': lambda r: f"=IF({tp}{r}>=1,0,IFERROR({cu}{r}/(1-{tp}{r}),0))",
        'VL_ICMS_Saida': lambda r: f"={pv}{r}*{pic}{r}",
        'VL_PIS_COFINS_Saida': lambda r: f"={pv}{r}*{ppc}{r}",
        'VL_IRPJ_CSLL': lambda r: f"={pv}{r}*{pir}{r}",
        'Margem_Bruta_RS': lambda r: f"={pv}{r}-{cu}{r}-{vic}{r}-{vpc}{r}-{vir}{r}",
        'Margem_Bruta_Perc': lambda r: f"=IFERROR({mb}{r}/{pv}{r},0)",
        'Diferenca_RS': lambda r: f'=IF({pp}{r}="","",{pp}{r}-{pv}{r})',
        'Diferenca_Perc': lambda r: f'=IF(OR({pp}{r}="",{pv}{r}=0),"",{pp}{r}/{pv}{r}-1)',
        'Situacao_Preco': lambda r: (
            f'=IF({pp}{r}="","-",'
            f'IF({pp}{r}<{pv}{r},"ABAIXO DO IDEAL",'
            f'IF({pp}{r}>{pv}{r},"ACIMA DO IDEAL","NO IDEAL")))'
        ),
    }

    # ---------- Dados ----------
    for r in range(len(df)):
        banda = r % 2
        linha_excel = r + 2  # 1-based, pulando o cabeçalho
        registro = df.iloc[r]
        for c, nome in enumerate(colunas):
            formato = fmt(nome, banda)
            valor = registro[nome]

            if nome == 'Preco_Praticado':
                ws.write_blank(r + 1, c, None, formato)
            elif nome in formulas:
                cached = valor
                if nome in ('Diferenca_RS', 'Diferenca_Perc'):
                    cached = ''
                elif nome == 'Situacao_Preco':
                    cached = '-'
                elif pd.isna(cached):
                    cached = 0
                ws.write_formula(r + 1, c, formulas[nome](linha_excel), formato, cached)
            elif isinstance(valor, float) and pd.isna(valor):
                ws.write_blank(r + 1, c, None, formato)
            elif isinstance(valor, (int, float, np.integer, np.floating)):
                ws.write_number(r + 1, c, float(valor), formato)
            else:
                ws.write(r + 1, c, str(valor), formato)

    # ---------- Larguras ----------
    for c, nome in enumerate(colunas):
        if df.empty:
            largura = len(nome) + 4
        else:
            amostra = df[nome].head(400).map(lambda v: len(str(v)) if pd.notna(v) else 0)
            largura = max(len(nome) + 4, int(amostra.max()) + 2)
        ws.set_column(c, c, min(max(largura, 11), 42))

    # ---------- Congelamento, filtro e destaques ----------
    col_congela = idx.get('NCM', 0) + 1  # congela tudo até a coluna NCM (inclusive)
    ws.freeze_panes(1, col_congela)

    if len(df) > 0:
        ws.autofilter(0, 0, len(df), len(colunas) - 1)
        faixa_dif = f"{L('Diferenca_Perc')}2:{L('Diferenca_Perc')}{len(df) + 1}"
        ws.conditional_format(faixa_dif, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_neg})
        ws.conditional_format(faixa_dif, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': fmt_pos})
        faixa_mb = f"{L('Margem_Bruta_RS')}2:{L('Margem_Bruta_RS')}{len(df) + 1}"
        ws.conditional_format(faixa_mb, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_neg})

    # ---------- Aba de instruções ----------
    ws2 = wb.add_worksheet('Parametros')
    tit = wb.add_format({'bold': True, 'font_size': 13, 'font_color': '#1F3864'})
    neg = wb.add_format({'bold': True, 'bg_color': '#EAF1F8', 'border': 1})
    txt = wb.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
    ws2.set_column(0, 0, 34)
    ws2.set_column(1, 1, 78)
    ws2.write(0, 0, 'Memória de cálculo da precificação', tit)

    memoria = [
        ('Antecipacao_ICMS', '(Custo_Total_da_Mercadoria x % ICMS Saída) - vICMS. '
                             'Zera quando o % ICMS Saída for 0% ou quando o resultado '
                             'for negativo (crédito de entrada maior que o débito de saída)'),
        ('Custo_Aquisicao', 'Custo_Total_da_Mercadoria - vICMS quando % ICMS Saída > 0; '
                            'caso contrário, igual ao Custo_Total_da_Mercadoria'),
        ('Custo_Unitario_Aquisicao', 'Custo_Aquisicao / qTrib'),
        ('Total_Perc', '% Margem + % ICMS Saída + % PIS/COFINS + % IRPJ/CSLL'),
        ('Preco_Ideal_Venda', 'Custo_Unitario_Aquisicao / (1 - Total %)'),
        ('VL_ICMS_Saida', 'Preço Ideal x % ICMS Saída'),
        ('VL_PIS_COFINS_Saida', 'Preço Ideal x % PIS/COFINS'),
        ('VL_IRPJ_CSLL', 'Preço Ideal x % IRPJ/CSLL'),
        ('Margem_Bruta_RS', 'Preço Ideal - Custo Unitário - VL ICMS - VL PIS/COFINS - VL IRPJ/CSLL'),
        ('Margem_Bruta_Perc', 'Margem Bruta R$ / Preço Ideal'),
        ('Preco_Praticado', 'DIGITAÇÃO MANUAL (célula amarela)'),
        ('Diferenca_RS', 'Preço Praticado - Preço Ideal (negativo = praticando abaixo do ideal)'),
        ('Diferenca_Perc', 'Preço Praticado / Preço Ideal - 1'),
        ('Situacao_Preco', 'ABAIXO DO IDEAL / NO IDEAL / ACIMA DO IDEAL'),
    ]
    for i, (campo, desc) in enumerate(memoria, start=2):
        ws2.write(i, 0, campo, neg)
        ws2.write(i, 1, desc, txt)

    linha = len(memoria) + 4
    ws2.write(linha, 0, 'Colunas amarelas', neg)
    ws2.write(linha, 1, 'Perc_Margem, Perc_ICMS_Saida, Perc_PIS_COFINS, Perc_IRPJ_CSLL e '
                        'Preco_Praticado podem ser alterados linha a linha — todo o bloco de '
                        'precificação recalcula automaticamente na planilha.', txt)

    wb.close()
    return output.getvalue()


# =============================================================================
# 4. INTERFACE STREAMLIT
# =============================================================================

st.set_page_config(page_title="Relatório NF-e + Precificação", layout="wide")
st.title("📄 Relatório de NF-e/NFC-e com Precificação (v3.0)")
st.caption("Extrai dados de `.xml`, `.zip` ou `.rar`, calcula custo de aquisição e preço ideal de venda.")

# ---------- Parâmetros ----------
OPC_ICMS = {"0,00%": 0.0, "4,00%": 0.04, "5,60%": 0.056, "20,50%": 0.205}
OPC_PIS = {"3,65%": 0.0365, "0,365%": 0.00365}

with st.sidebar:
    st.header("⚙️ Parâmetros de precificação")
    margem_pct = st.number_input("Margem (%)", min_value=0.0, max_value=99.0,
                                 value=20.0, step=0.5, format="%.2f")
    icms_label = st.selectbox("% ICMS Saída", list(OPC_ICMS.keys()), index=3)
    pis_label = st.selectbox("% PIS e COFINS", list(OPC_PIS.keys()), index=0)
    irpj_pct = st.number_input("% IRPJ e CSLL", min_value=0.0, max_value=99.0,
                               value=3.90, step=0.05, format="%.2f")

    perc_margem = margem_pct / 100
    perc_icms = OPC_ICMS[icms_label]
    perc_pis = OPC_PIS[pis_label]
    perc_irpj = irpj_pct / 100
    total_perc = perc_margem + perc_icms + perc_pis + perc_irpj

    st.metric("Total %", f"{total_perc * 100:.2f}%")
    if total_perc >= 1:
        st.error("Total % ≥ 100%. O preço ideal não pode ser calculado — reduza a margem.")

    if not RAR_DISPONIVEL:
        st.warning("Biblioteca `rarfile`/`unrar` indisponível: arquivos .rar não serão lidos.")

uploaded_files = st.file_uploader(
    "Selecione os arquivos XML, ZIP ou RAR",
    type=["xml", "zip", "rar"],
    accept_multiple_files=True,
    help="Você pode arrastar múltiplos arquivos ou arquivos compactados."
)

if uploaded_files and st.button("🚀 Processar XMLs", type="primary"):
    start_time = time.time()
    xml_streams = []
    unpack_errors = []

    with st.spinner('Preparando e descompactando arquivos...'):
        for file in uploaded_files:
            filename = file.name
            ext = filename.lower().split('.')[-1]
            try:
                if ext == 'xml':
                    xml_streams.append((filename, file))

                elif ext == 'zip':
                    with zipfile.ZipFile(file, 'r') as zf:
                        for nome in zf.namelist():
                            if nome.lower().endswith('.xml'):
                                xml_streams.append((f"{filename}/{nome}", BytesIO(zf.read(nome))))

                elif ext == 'rar':
                    if not RAR_DISPONIVEL:
                        unpack_errors.append({"arquivo": filename,
                                              "erro": "Processamento de RAR não configurado no servidor."})
                        continue
                    try:
                        with rarfile.RarFile(file, 'r') as rf:
                            for nome in rf.namelist():
                                if nome.lower().endswith('.xml'):
                                    xml_streams.append((f"{filename}/{nome}", BytesIO(rf.read(nome))))
                    except rarfile.NeedFirstVolume:
                        unpack_errors.append({"arquivo": filename, "erro": "RAR multi-volume não suportado."})
                    except Exception as e:
                        unpack_errors.append({"arquivo": filename, "erro": f"Erro ao ler RAR: {e}"})

            except Exception as e:
                unpack_errors.append({"arquivo": filename, "erro": f"Erro ao abrir arquivo: {e}"})

    if not xml_streams:
        st.warning("Nenhum arquivo XML válido foi encontrado nos uploads.")
        if unpack_errors:
            st.dataframe(pd.DataFrame(unpack_errors), use_container_width=True)
    else:
        st.info(f"Arquivos preparados. Processando {len(xml_streams)} XMLs...")
        df_base, process_errors = process_xml_streams(xml_streams)
        st.session_state['df_base'] = df_base
        st.session_state['erros'] = unpack_errors + process_errors
        st.session_state['tempo'] = time.time() - start_time

# ---------- Resultado ----------
if 'df_base' in st.session_state:
    df_base = st.session_state['df_base']
    erros = st.session_state.get('erros', [])

    if df_base.empty:
        st.warning("Nenhum dado foi extraído. Verifique os XMLs.")
    else:
        df_final = aplicar_precificacao(df_base, perc_margem, perc_icms, perc_pis, perc_irpj)

        st.success(
            f"{len(df_final)} itens processados em {st.session_state.get('tempo', 0):.2f}s. "
            f"Parâmetros aplicados: margem {margem_pct:.2f}% | ICMS {icms_label} | "
            f"PIS/COFINS {pis_label} | IRPJ/CSLL {irpj_pct:.2f}%."
        )

        c1, c2, c3 = st.columns(3)
        c1.metric("Custo total das mercadorias", f"R$ {df_final['Custo_Total_da_Mercadoria'].sum():,.2f}")
        c2.metric("Antecipação de ICMS estimada", f"R$ {df_final['Antecipacao_ICMS'].sum():,.2f}")
        c3.metric("Margem bruta média", f"{df_final['Margem_Bruta_Perc'].mean() * 100:.2f}%")

        st.dataframe(df_final, use_container_width=True, height=460)

        st.download_button(
            label="📥 Baixar Relatório em Excel",
            data=gerar_excel(df_final),
            file_name="relatorio_nfe_precificacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    if erros:
        st.error(f"Ocorreram {len(erros)} erros durante o processo:")
        with st.expander("Ver detalhes dos erros"):
            st.dataframe(pd.DataFrame(erros), use_container_width=True)

elif not uploaded_files:
    st.info("Aguardando o upload de arquivos `.xml`, `.zip` ou `.rar`.")

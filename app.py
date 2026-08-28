"""
====================================================================================
 Gerador de Relatorio de NF-e / NFC-e  -  v3.0 (Reforma Tributaria do Consumo)
====================================================================================
 Extrai dados de XML de NF-e/NFC-e (mod. 55/65) e consolida em relatorio Excel.

 NOVIDADE v3.0: leitura completa do Grupo UB da NT 2025.002-RTC (leiaute ate v1.40)
   - IBSCBS  (CST + cClassTrib + indDoacao)
   - gIBSCBS (vBC, gIBSUF, gIBSMun, vIBS, gCBS) com gDif / gDevTrib / gRed
   - gTribRegular  (beneficio fiscal condicional - ZFM/ALC)
   - gTribCompraGov (compras governamentais)
   - gALCZFMCBS    (aliquota zero CBS em ALC/ZFM - novo na v1.40)
   - gIBSCBSMono   (gMonoPadrao / gMonoReten / gMonoRet / gMonoDif) - CST 620
   - gTransfCred   - CST 800
   - gAjusteCompet - CST 811
   - gEstornoCred  (UB116)
   - gCredPresOper (UB120) com gIBSCredPres / gCBSCredPres
   - gCredPresIBSZFM - CST 810
   - IS - Imposto Seletivo (CSTIS, cClassTribIS, vBCIS, pIS, adRemIS, uTrib, qTrib, vIS)
   - Cabecalho: cMunFGIBS, tpNFDebito, tpNFCredito, cIndOp, gCompraGov
   - Item: vItem, indBemMovelUsado, tpCredPresIBSZFM
   - Totais: IBSCBSTot (vBCIBSCBS, gIBS, gCBS, gMono, gEstornoCred), vNFTot

 Estrategia anti-erro: nenhuma tag e obrigatoria. Toda leitura passa por funcoes
 que retornam vazio/zero quando o no nao existe, e cada grupo e lido no seu
 escopo proprio (evita colisao de nomes como vIBS, vCredPres, vDif e pDif, que
 aparecem em mais de um grupo com significados diferentes).
====================================================================================
"""

import streamlit as st
from lxml import etree as ET
import pandas as pd
from io import BytesIO
import zipfile
import time

# rarfile e opcional: se nao estiver instalado, o app continua funcionando
try:
    import rarfile
    RARFILE_OK = True
except Exception:
    rarfile = None
    RARFILE_OK = False


# ====================================================================================
# CONFIGURACAO
# ====================================================================================

# Percentuais do ICMS/IPI ja eram divididos por 100 no codigo original (fracao).
# Para IBS/CBS/IS os percentuais sao mantidos COMO ESTAO NO XML (pIBSUF 0.1000 = 0,1%).
# Mude para True se preferir que tambem virem fracao (0.1000 -> 0.001).
PERC_RTC_COMO_FRACAO = False

# Tolerancia usada nas conferencias de calculo do IBS/CBS
TOLERANCIA = 0.02


# ====================================================================================
# TABELAS OFICIAIS (NT 2025.002 / Portal DF-e SVRS)
# ====================================================================================

CST_IBSCBS = {
    '000': 'Tributacao integral',
    '010': 'Tributacao com aliquotas uniformes',
    '011': 'Tributacao com aliquotas uniformes reduzidas',
    '200': 'Aliquota reduzida',
    '220': 'Aliquota fixa',
    '221': 'Aliquota fixa proporcional',
    '222': 'Reducao de base de calculo',
    '400': 'Isencao',
    '410': 'Imunidade e nao incidencia',
    '510': 'Diferimento',
    '515': 'Diferimento com reducao de aliquota',
    '550': 'Suspensao',
    '620': 'Tributacao monofasica',
    '800': 'Transferencia de credito',
    '810': 'Ajuste de IBS na ZFM',
    '811': 'Ajustes',
    '820': 'Tributacao em declaracao de regime especifico',
    '830': 'Exclusao da base de calculo',
}

# Grupo de tributo exigido por CST (base para o diagnostico)
CST_EXIGE_GIBSCBS = {'000', '010', '011', '200', '220', '221', '222', '510', '515', '550'}
CST_EXIGE_MONO = {'620'}
CST_EXIGE_TRANSFCRED = {'800'}
CST_EXIGE_AJUSTECOMPET = {'811'}
CST_EXIGE_CREDPRESZFM = {'810'}
CST_SEM_GRUPO_TRIBUTO = {'400', '410', '810', '820', '830'}
CST_EXIGE_GRED = {'200', '515'}
CST_EXIGE_GDIF = {'510', '515'}

# cClassTrib - tabela oficial (descricao | % reducao IBS | % reducao CBS)
CCLASSTRIB = {
    '000001': ('Situacoes tributadas integralmente pelo IBS e CBS', 0, 0),
    '000002': ('Exploracao de via', 0, 0),
    '000003': ('Regime automotivo - projetos incentivados (art. 311)', 0, 0),
    '000004': ('Regime automotivo - projetos incentivados (art. 312)', 0, 0),
    '000005': ('Operacao com EAC destinado a mistura com gasolina A, com destinacao diversa', 0, 0),
    '010001': ('Operacoes do FGTS nao realizadas pela Caixa Economica Federal', 0, 0),
    '010002': ('Operacoes do servico financeiro', 0, 0),
    '011001': ('Planos de assistencia funeraria', 60, 60),
    '011002': ('Planos de assistencia a saude', 60, 60),
    '011003': ('Intermediacao de planos de assistencia a saude', 60, 60),
    '011004': ('Concursos e prognosticos', 0, 0),
    '011005': ('Planos de assistencia a saude de animais domesticos', 30, 30),
    '200001': ('Aquisicoes entre empresas em zonas de processamento de exportacao', 100, 100),
    '200002': ('Fornecimento ou importacao para produtor rural nao contribuinte ou TAC', 100, 100),
    '200003': ('Vendas de produtos destinados a alimentacao humana (Anexo I)', 100, 100),
    '200004': ('Venda de dispositivos medicos (Anexo XII)', 100, 100),
    '200005': ('Venda de dispositivos medicos adquiridos por orgaos publicos (Anexo IV)', 100, 100),
    '200006': ('Situacao de emergencia de saude publica (Anexo XII)', 100, 100),
    '200007': ('Dispositivos de acessibilidade para pessoas com deficiencia (Anexo XIII)', 100, 100),
    '200008': ('Dispositivos de acessibilidade adquiridos por orgaos publicos (Anexo V)', 100, 100),
    '200009': ('Fornecimento de medicamentos (Anexo XIV)', 100, 100),
    '200010': ('Medicamentos registrados na Anvisa adquiridos por orgaos publicos', 100, 100),
    '200011': ('Nutricao enteral e parenteral adquirida por orgaos publicos (Anexo VI)', 100, 100),
    '200012': ('Situacao de emergencia de saude publica (Anexo XIV)', 100, 100),
    '200013': ('Tampoes e absorventes higienicos', 100, 100),
    '200014': ('Produtos horticolas, frutas e ovos (Anexo XV)', 100, 100),
    '200015': ('Automoveis nacionais para motoristas profissionais ou PcD', 100, 100),
    '200016': ('Pesquisa e desenvolvimento por ICT', 100, 100),
    '200017': ('Operacoes relacionadas ao FGTS', 100, 100),
    '200018': ('Operacoes de resseguro e retrocessao', 100, 100),
    '200019': ('Importador dos servicos financeiros contribuinte', 100, 100),
    '200020': ('Sociedades cooperativas optantes por regime especifico', 100, 100),
    '200021': ('Transporte publico coletivo ferroviario e hidroviario', 100, 100),
    '200022': ('Operacao de fora da ZFM que destine bem industrializado a contribuinte na ZFM', 100, 100),
    '200023': ('Industria incentivada que destine bem intermediario a outra industria na ZFM', 100, 100),
    '200024': ('Operacao de fora das ALC destinada a contribuinte nas ALC', 100, 100),
    '200025': ('Servicos de educacao - Prouni', 60, 100),
    '200026': ('Locacao de imoveis em zonas reabilitadas', 80, 80),
    '200027': ('Locacao, cessao onerosa e arrendamento de bens imoveis', 70, 70),
    '200028': ('Servicos de educacao (Anexo II)', 60, 60),
    '200029': ('Servicos de saude humana (Anexo III)', 60, 60),
    '200030': ('Venda de dispositivos medicos (Anexo IV)', 60, 60),
    '200031': ('Dispositivos de acessibilidade para pessoas com deficiencia (Anexo V)', 60, 60),
    '200032': ('Medicamentos registrados na Anvisa ou de farmacia de manipulacao', 60, 60),
    '200033': ('Nutricao enteral e parenteral (Anexo VI)', 60, 60),
    '200034': ('Alimentos destinados ao consumo humano (Anexo VII)', 60, 60),
    '200035': ('Produtos de higiene pessoal e limpeza (Anexo VIII)', 60, 60),
    '200036': ('Produtos agropecuarios, aquicolas, pesqueiros, florestais e extrativistas in natura', 60, 60),
    '200037': ('Servicos ambientais de conservacao ou recuperacao da vegetacao nativa', 60, 60),
    '200038': ('Fornecimento de insumos agropecuarios e aquicolas (Anexo IX)', 60, 60),
    '200039': ('Servicos e cessao de direitos para producoes nacionais artisticas (Anexo X)', 60, 60),
    '200040': ('Servicos de comunicacao institucional a administracao publica', 60, 60),
    '200041': ('Servico de educacao desportiva (art. 141, I)', 60, 60),
    '200042': ('Servico de educacao desportiva (art. 141, II)', 60, 60),
    '200043': ('Fornecimento a administracao publica de bens e servicos de soberania (Anexo XI)', 60, 60),
    '200044': ('Seguranca da informacao e ciberseguranca - socio brasileiro (Anexo XI)', 60, 60),
    '200045': ('Projetos de reabilitacao urbana de zonas historicas e areas criticas', 60, 60),
    '200046': ('Operacoes com bens imoveis', 50, 50),
    '200047': ('Bares e restaurantes', 40, 40),
    '200048': ('Hotelaria, parques de diversao e parques tematicos', 40, 40),
    '200049': ('Transporte coletivo rodoviario, ferroviario e hidroviario', 40, 40),
    '200050': ('Transporte aereo regional coletivo de passageiros ou de carga', 40, 40),
    '200051': ('Agencias de turismo', 40, 40),
    '200052': ('Prestacao de servicos de profissoes intelectuais', 30, 30),
    '220001': ('Incorporacao imobiliaria submetida ao regime especial de tributacao', 0, 0),
    '220002': ('Incorporacao imobiliaria submetida ao regime especial de tributacao', 0, 0),
    '220003': ('Alienacao de imovel decorrente de parcelamento do solo', 0, 0),
    '221001': ('Locacao, cessao onerosa ou arrendamento de imovel sobre a receita bruta', 0, 0),
    '222001': ('Transporte internacional de passageiros com ida e volta em conjunto', 0, 0),
    '400001': ('Transporte publico coletivo de passageiros rodoviario e metroviario', 0, 0),
    '410001': ('Bonificacoes constantes no documento fiscal que nao dependam de evento posterior', 0, 0),
    '410002': ('Transferencias entre estabelecimentos do mesmo contribuinte', 0, 0),
    '410003': ('Doacoes sem contraprestacao em beneficio do doador', 0, 0),
    '410004': ('Exportacoes de bens e servicos', 0, 0),
    '410005': ('Fornecimentos realizados pela Uniao, Estados, DF e Municipios', 0, 0),
    '410006': ('Fornecimentos por entidades religiosas e templos de qualquer culto', 0, 0),
    '410007': ('Fornecimentos realizados por partidos politicos', 0, 0),
    '410008': ('Livros, jornais, periodicos e papel destinado a sua impressao', 0, 0),
    '410009': ('Fonogramas e videofonogramas musicais produzidos no Brasil', 0, 0),
    '410010': ('Radiodifusao sonora e de sons e imagens de recepcao livre e gratuita', 0, 0),
    '410011': ('Ouro definido em lei como ativo financeiro ou instrumento cambial', 0, 0),
    '410012': ('Condominio edilicio nao optante pelo regime regular', 0, 0),
    '410013': ('Exportacoes de combustiveis', 0, 0),
    '410014': ('Fornecimento de produtor rural nao contribuinte', 0, 0),
    '410015': ('Fornecimento por transportador autonomo nao contribuinte', 0, 0),
    '410016': ('Fornecimento ou aquisicao de residuos solidos', 0, 0),
    '410017': ('Aquisicao de bem movel com credito presumido sob condicao de revenda', 0, 0),
    '410018': ('Fundos garantidores e executores de politicas publicas', 0, 0),
    '410019': ('Exclusao da gorjeta na base de calculo no fornecimento de alimentacao', 0, 0),
    '410020': ('Exclusao do valor de intermediacao na BC no fornecimento de alimentacao', 0, 0),
    '410021': ('Contribuicao de que trata o art. 149-A da Constituicao Federal', 0, 0),
    '410022': ('Consolidacao da propriedade do bem pelo credor', 0, 0),
    '410023': ('Alienacao de bem objeto de garantia - prestador nao contribuinte', 0, 0),
    '410024': ('Consolidacao da propriedade do bem pelo grupo de consorcio', 0, 0),
    '410025': ('Alienacao de bem objeto de garantia - prestador nao contribuinte', 0, 0),
    '410026': ('Doacao com anulacao de credito', 0, 0),
    '410027': ('Exportacao de servico ou de bem imaterial', 0, 0),
    '410028': ('Operacoes com bens imoveis por pessoas fisicas nao contribuintes', 0, 0),
    '410029': ('Operacoes acobertadas somente pelo ICMS', 0, 0),
    '410030': ('Estorno de credito por perecimento, deterioracao, roubo, furto ou extravio', 0, 0),
    '410031': ('Fornecimento em periodo anterior ao inicio de vigencia de CBS e IBS', 0, 0),
    '410999': ('Operacoes nao onerosas sem previsao de tributacao, nao especificadas', 0, 0),
    '510001': ('Diferimento com energia eletrica (geracao, comercializacao, distribuicao, transmissao)', 0, 0),
    '515001': ('Diferimento com insumos agropecuarios e aquicolas (Anexo IX)', 60, 60),
    '550001': ('Exportacoes de bens materiais', 0, 0),
    '550002': ('Regime de Transito', 0, 0),
    '550003': ('Regimes de Deposito (art. 85)', 0, 0),
    '550004': ('Regimes de Deposito (art. 87)', 0, 0),
    '550005': ('Regimes de Deposito (art. 87, paragrafo unico)', 0, 0),
    '550006': ('Regimes de Permanencia Temporaria', 0, 0),
    '550007': ('Regimes de Aperfeicoamento', 0, 0),
    '550008': ('Importacao de bens para o Regime de Repetro-Temporario', 0, 0),
    '550009': ('GNL-Temporario', 0, 0),
    '550010': ('Repetro-Permanente', 0, 0),
    '550011': ('Repetro-Industrializacao', 0, 0),
    '550012': ('Repetro-Nacional', 0, 0),
    '550013': ('Repetro-Entreposto', 0, 0),
    '550014': ('Zona de Processamento de Exportacao', 0, 0),
    '550015': ('Incentivo a Modernizacao e Ampliacao da Estrutura Portuaria', 0, 0),
    '550016': ('Regime Especial de Incentivos para Desenvolvimento da Infraestrutura', 0, 0),
    '550017': ('Incentivo a Atividade Economica Naval', 0, 0),
    '550018': ('Desoneracao da aquisicao de bens de capital', 0, 0),
    '550019': ('Importacao por industria incentivada para utilizacao na ZFM', 0, 0),
    '550020': ('Areas de livre comercio', 0, 0),
    '550021': ('Industrializacao destinada a exportacoes', 0, 0),
    '620001': ('Tributacao monofasica sobre combustiveis', 0, 0),
    '620002': ('Monofasica com responsabilidade pela retencao sobre combustiveis', 0, 0),
    '620003': ('Monofasica com tributos retidos por responsabilidade sobre combustiveis', 0, 0),
    '620004': ('Monofasica sobre mistura de EAC com gasolina A acima do percentual obrigatorio', 0, 0),
    '620005': ('Monofasica sobre mistura de EAC com gasolina A abaixo do percentual obrigatorio', 0, 0),
    '620006': ('Monofasica sobre combustiveis cobrada anteriormente', 0, 0),
    '800001': ('Fusao, cisao ou incorporacao', 0, 0),
    '800002': ('Transferencia de credito do associado, inclusive cooperativas singulares', 0, 0),
    '810001': ('Credito presumido sobre o valor apurado nos fornecimentos a partir da ZFM', 0, 0),
    '811001': ('Anulacao de credito por saidas imunes ou isentas', 0, 0),
    '811002': ('Debitos de notas fiscais nao processadas na apuracao', 0, 0),
    '811003': ('Desenquadramento do Simples Nacional', 0, 0),
    '820001': ('Documento com informacoes de servicos de planos de assistencia a saude', 0, 0),
    '820002': ('Documento com informacoes de servicos de planos de assistencia funeraria', 0, 0),
    '820003': ('Documento com informacoes de planos de saude de animais domesticos', 0, 0),
    '820004': ('Documento com informacoes de concursos de prognosticos', 0, 0),
    '820005': ('Documento com informacoes de alienacao de bens imoveis', 0, 0),
    '820006': ('Documento com informacoes de servicos de exploracao de via', 0, 0),
    '820007': ('Documento com informacoes de fornecimento de servicos financeiros', 0, 0),
    '820008': ('Documento com informacoes de fornecimento tributado em fatura anterior', 0, 0),
    '830001': ('Exclusao da BC de energia eletrica fornecida pela distribuidora', 0, 0),
}

# cCredPres - Anexo IV (codigos de credito presumido)
CCREDPRES = {
    '01': 'Aquisicao de produtor rural nao contribuinte',
    '02': 'Tomador de servico de transporte de TAC PF nao contribuinte',
    '03': 'Aquisicao de PF com destino a reciclagem',
    '04': 'Aquisicao de bens moveis de PF nao contribuinte para revenda',
    '05': 'Regime opcional para cooperativa',
}

# Aliquotas de referencia do periodo de transicao (LC 214/25, arts. 343/344/346)
ALIQ_TRANSICAO = {
    2025: {'IBSUF': 0.10, 'CBS': 0.90},
    2026: {'IBSUF': 0.10, 'CBS': 0.90},
    2027: {'IBSUF': 0.05, 'CBS': None},
    2028: {'IBSUF': 0.05, 'CBS': None},
}


# ====================================================================================
# HELPERS DE LEITURA DE XML  (tolerantes a ausencia de qualquer no)
# ====================================================================================

def strip_namespace(root):
    """Remove o namespace de todos os elementos da arvore (lxml)."""
    for el in root.iter():
        if isinstance(el.tag, str) and '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]
    ET.cleanup_namespaces(root)
    return root


def get_text(node, tag, default=''):
    """Texto de um no filho, com seguranca contra node/child None."""
    if node is None:
        return default
    child = node.find(tag)
    if child is None or child.text is None:
        return default
    return child.text.strip()


def to_float(value):
    """Converte para float com seguranca (aceita virgula decimal)."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    v = str(value).strip()
    if not v:
        return 0.0
    try:
        return float(v)
    except ValueError:
        try:
            return float(v.replace('.', '').replace(',', '.'))
        except ValueError:
            return 0.0


def sub(node, *names):
    """Retorna o primeiro subgrupo existente entre varios nomes possiveis.

    Usado onde a NT/emissores divergem na grafia (gDif/gDIF, gDevTrib/gDev).
    Se nenhum existir, devolve o proprio node para que a leitura das tags
    filhas caia no comportamento padrao (vazio/zero) sem quebrar.
    """
    if node is None:
        return None
    for n in names:
        found = node.find(n)
        if found is not None:
            return found
    return None


def val(node, tag):
    """Valor monetario/quantitativo: float, 0.0 se ausente."""
    return to_float(get_text(node, tag))


def perc(node, tag):
    """Percentual do RTC: mantido como no XML, salvo config em contrario."""
    p = to_float(get_text(node, tag))
    return p / 100 if PERC_RTC_COMO_FRACAO else p


def desc_cst(cst):
    return CST_IBSCBS.get(cst, '')


def desc_cclass(c):
    item = CCLASSTRIB.get(c)
    return item[0] if item else ''


def red_cclass(c):
    item = CCLASSTRIB.get(c)
    return (item[1], item[2]) if item else ('', '')


# ====================================================================================
# FORMATACAO
# ====================================================================================

def format_ncm(ncm):
    return f"{ncm[0:4]}.{ncm[4:6]}.{ncm[6:8]}" if len(ncm) >= 8 else ncm


def format_cest(cest):
    if not cest:
        return ''
    cest = cest.zfill(7)
    return f"{cest[0:2]}.{cest[2:5]}.{cest[5:7]}"


def format_cfop(cfop):
    if not cfop:
        return ''
    cfop = cfop.zfill(4)
    return f"{cfop[0]}.{cfop[1:4]}"


# ====================================================================================
# LEITURA DOS GRUPOS DA REFORMA TRIBUTARIA (Grupo UB)
# ====================================================================================

def _ler_bloco_aliquota(grupo, prefixo, tag_valor):
    """Le um bloco gIBSUF / gIBSMun / gCBS.

    Estrutura: p<Tributo>, gDif(pDif,vDif), gDevTrib(pDevTrib,vDevTrib),
               gRed(pRedAliq,pAliqEfet), v<Tributo>.
    Le cada subgrupo no seu escopo para nao confundir pDif do IBS com o do CBS.
    """
    d = {
        f'p{prefixo}': 0.0,
        f'pRedAliq_{prefixo}': 0.0,
        f'pAliqEfet_{prefixo}': 0.0,
        f'pDif_{prefixo}': 0.0,
        f'vDif_{prefixo}': 0.0,
        f'pDevTrib_{prefixo}': 0.0,
        f'vDevTrib_{prefixo}': 0.0,
        f'v{prefixo}': 0.0,
    }
    if grupo is None:
        return d

    d[f'p{prefixo}'] = perc(grupo, f'p{prefixo}')
    d[f'v{prefixo}'] = val(grupo, tag_valor)

    g_red = sub(grupo, 'gRed')
    if g_red is not None:
        d[f'pRedAliq_{prefixo}'] = perc(g_red, 'pRedAliq')
        d[f'pAliqEfet_{prefixo}'] = perc(g_red, 'pAliqEfet')

    g_dif = sub(grupo, 'gDif', 'gDIF')
    if g_dif is not None:
        d[f'pDif_{prefixo}'] = perc(g_dif, 'pDif')
        d[f'vDif_{prefixo}'] = val(g_dif, 'vDif')

    # gDevTrib (nome do grupo variou entre versoes da NT: gDevTrib / gDev)
    g_dev = sub(grupo, 'gDevTrib', 'gDev')
    if g_dev is not None:
        d[f'pDevTrib_{prefixo}'] = perc(g_dev, 'pDevTrib')
        d[f'vDevTrib_{prefixo}'] = val(g_dev, 'vDevTrib')
    else:
        # alguns emissores colocam as tags soltas dentro do grupo
        d[f'pDevTrib_{prefixo}'] = perc(grupo, 'pDevTrib')
        d[f'vDevTrib_{prefixo}'] = val(grupo, 'vDevTrib')

    return d


def colunas_rtc_vazias():
    """Dicionario com todas as colunas do Grupo UB zeradas/vazias.

    Garante que o DataFrame tenha sempre o mesmo conjunto de colunas,
    mesmo que nenhum XML do lote traga IBS/CBS.
    """
    d = {
        # --- identificacao IBSCBS ---
        'CST_IBSCBS': '', 'CST_IBSCBS_Desc': '',
        'cClassTrib': '', 'cClassTrib_Desc': '',
        'pRed_Tabela_IBS': '', 'pRed_Tabela_CBS': '',
        'indDoacao': '', 'Grupo_Tributo_RTC': '',
        # --- gIBSCBS ---
        'vBC_IBSCBS': 0.0,
        'pIBSUF': 0.0, 'pRedAliq_IBSUF': 0.0, 'pAliqEfet_IBSUF': 0.0,
        'pDif_IBSUF': 0.0, 'vDif_IBSUF': 0.0,
        'pDevTrib_IBSUF': 0.0, 'vDevTrib_IBSUF': 0.0, 'vIBSUF': 0.0,
        'pIBSMun': 0.0, 'pRedAliq_IBSMun': 0.0, 'pAliqEfet_IBSMun': 0.0,
        'pDif_IBSMun': 0.0, 'vDif_IBSMun': 0.0,
        'pDevTrib_IBSMun': 0.0, 'vDevTrib_IBSMun': 0.0, 'vIBSMun': 0.0,
        'vIBS': 0.0,
        'pCBS': 0.0, 'pRedAliq_CBS': 0.0, 'pAliqEfet_CBS': 0.0,
        'pDif_CBS': 0.0, 'vDif_CBS': 0.0,
        'pDevTrib_CBS': 0.0, 'vDevTrib_CBS': 0.0, 'vCBS': 0.0,
        # --- gALCZFMCBS (v1.40) ---
        'tpALCZFMCBS': '', 'pAliqEfeRegCBS_ALCZFM': 0.0, 'vTribRegCBS_ALCZFM': 0.0,
        # --- gTribRegular ---
        'CSTReg': '', 'cClassTribReg': '',
        'pAliqEfetRegIBSUF': 0.0, 'vTribRegIBSUF': 0.0,
        'pAliqEfetRegIBSMun': 0.0, 'vTribRegIBSMun': 0.0,
        'pAliqEfetRegCBS': 0.0, 'vTribRegCBS': 0.0,
        # --- gTribCompraGov ---
        'pAliqIBSUF_Gov': 0.0, 'vTribIBSUF_Gov': 0.0,
        'pAliqIBSMun_Gov': 0.0, 'vTribIBSMun_Gov': 0.0,
        'pAliqCBS_Gov': 0.0, 'vTribCBS_Gov': 0.0,
        # --- gIBSCBSMono (CST 620) ---
        'qBCMono': 0.0, 'adRemIBS': 0.0, 'adRemCBS': 0.0,
        'vIBSMono': 0.0, 'vCBSMono': 0.0,
        'qBCMonoReten': 0.0, 'adRemIBSReten': 0.0, 'vIBSMonoReten': 0.0,
        'adRemCBSReten': 0.0, 'vCBSMonoReten': 0.0,
        'qBCMonoRet': 0.0, 'adRemIBSRet': 0.0, 'vIBSMonoRet': 0.0,
        'adRemCBSRet': 0.0, 'vCBSMonoRet': 0.0,
        'pDifIBS_Mono': 0.0, 'vIBSMonoDif': 0.0,
        'pDifCBS_Mono': 0.0, 'vCBSMonoDif': 0.0,
        'vTotIBSMonoItem': 0.0, 'vTotCBSMonoItem': 0.0,
        # --- gTransfCred (CST 800) ---
        'vIBS_TransfCred': 0.0, 'vCBS_TransfCred': 0.0,
        # --- gAjusteCompet (CST 811) ---
        'competApur_Ajuste': '', 'vIBS_AjusteCompet': 0.0, 'vCBS_AjusteCompet': 0.0,
        # --- gEstornoCred ---
        'vIBSEstCred': 0.0, 'vCBSEstCred': 0.0,
        # --- gCredPresOper ---
        'vBCCredPres': 0.0, 'cCredPres': '', 'cCredPres_Desc': '',
        'pCredPres_IBS': 0.0, 'vCredPres_IBS': 0.0, 'vCredPresCondSus_IBS': 0.0,
        'pCredPres_CBS': 0.0, 'vCredPres_CBS': 0.0, 'vCredPresCondSus_CBS': 0.0,
        # --- gCredPresIBSZFM (CST 810) ---
        'competApur_ZFM': '', 'tpCredPresIBSZFM': '', 'vCredPres_ZFM': 0.0,
        # --- IS - Imposto Seletivo ---
        'CST_IS': '', 'cClassTrib_IS': '', 'vBC_IS': 0.0,
        'pIS': 0.0, 'adRemIS': 0.0, 'uTrib_IS': '', 'qTrib_IS': 0.0, 'vIS': 0.0,
        # --- item ---
        'vItem': 0.0, 'indBemMovelUsado': '', 'tpCredPresIBSZFM_Prod': '',
        # --- consolidado do item ---
        'Total_IBS_Item': 0.0, 'Total_CBS_Item': 0.0, 'Total_IBSCBS_Item': 0.0,
        # --- diagnostico ---
        'Diag_RTC': '',
    }
    return d


def extrair_ibscbs(imposto, det):
    """Le todo o Grupo UB (IBS/CBS/IS) de um item. Nunca levanta excecao."""
    d = colunas_rtc_vazias()
    if imposto is None:
        return d

    # ---------- IS - Imposto Seletivo ----------
    g_is = imposto.find('IS')
    if g_is is not None:
        d['CST_IS'] = get_text(g_is, 'CSTIS')
        d['cClassTrib_IS'] = get_text(g_is, 'cClassTribIS')
        d['vBC_IS'] = val(g_is, 'vBCIS')
        d['pIS'] = perc(g_is, 'pIS')
        # v1.40 renomeou pISEspec para adRemIS: aceita os dois
        d['adRemIS'] = perc(g_is, 'adRemIS') or perc(g_is, 'pISEspec')
        d['uTrib_IS'] = get_text(g_is, 'uTrib')
        d['qTrib_IS'] = val(g_is, 'qTrib')
        d['vIS'] = val(g_is, 'vIS')

    # ---------- IBSCBS ----------
    ibscbs = imposto.find('IBSCBS')
    if ibscbs is None:
        return d

    d['CST_IBSCBS'] = get_text(ibscbs, 'CST')
    d['cClassTrib'] = get_text(ibscbs, 'cClassTrib')
    d['CST_IBSCBS_Desc'] = desc_cst(d['CST_IBSCBS'])
    d['cClassTrib_Desc'] = desc_cclass(d['cClassTrib'])
    r_ibs, r_cbs = red_cclass(d['cClassTrib'])
    d['pRed_Tabela_IBS'] = r_ibs
    d['pRed_Tabela_CBS'] = r_cbs
    d['indDoacao'] = get_text(ibscbs, 'indDoacao')

    # ---------- gIBSCBS (tributacao padrao) ----------
    g = ibscbs.find('gIBSCBS')
    if g is not None:
        d['Grupo_Tributo_RTC'] = 'gIBSCBS'
        d['vBC_IBSCBS'] = val(g, 'vBC')
        d['vIBS'] = val(g, 'vIBS')          # filho direto: nao confundir com gTransfCred/vIBS

        d.update(_ler_bloco_aliquota(g.find('gIBSUF'), 'IBSUF', 'vIBSUF'))
        d.update(_ler_bloco_aliquota(g.find('gIBSMun'), 'IBSMun', 'vIBSMun'))

        g_cbs = g.find('gCBS')
        d.update(_ler_bloco_aliquota(g_cbs, 'CBS', 'vCBS'))

        # gALCZFMCBS (novo na v1.40)
        if g_cbs is not None:
            alc = g_cbs.find('gALCZFMCBS')
            if alc is not None:
                d['tpALCZFMCBS'] = get_text(alc, 'tpALCZFMCBS')
                d['pAliqEfeRegCBS_ALCZFM'] = perc(alc, 'pAliqEfeRegCBS')
                d['vTribRegCBS_ALCZFM'] = val(alc, 'vTribRegCBS')

        # gTribRegular
        reg = g.find('gTribRegular')
        if reg is not None:
            d['CSTReg'] = get_text(reg, 'CSTReg')
            d['cClassTribReg'] = get_text(reg, 'cClassTribReg')
            d['pAliqEfetRegIBSUF'] = perc(reg, 'pAliqEfetRegIBSUF')
            d['vTribRegIBSUF'] = val(reg, 'vTribRegIBSUF')
            d['pAliqEfetRegIBSMun'] = perc(reg, 'pAliqEfetRegIBSMun')
            d['vTribRegIBSMun'] = val(reg, 'vTribRegIBSMun')
            d['pAliqEfetRegCBS'] = perc(reg, 'pAliqEfetRegCBS')
            d['vTribRegCBS'] = val(reg, 'vTribRegCBS')

    # gTribCompraGov: a NT posiciona dentro de gIBSCBS, mas ha exemplos
    # com o grupo irmao de gIBSCBS. Busca nos dois lugares.
    gov = None
    if g is not None:
        gov = g.find('gTribCompraGov')
    if gov is None:
        gov = ibscbs.find('gTribCompraGov')
    if gov is not None:
        d['pAliqIBSUF_Gov'] = perc(gov, 'pAliqIBSUF')
        d['vTribIBSUF_Gov'] = val(gov, 'vTribIBSUF')
        d['pAliqIBSMun_Gov'] = perc(gov, 'pAliqIBSMun')
        d['vTribIBSMun_Gov'] = val(gov, 'vTribIBSMun')
        d['pAliqCBS_Gov'] = perc(gov, 'pAliqCBS')
        d['vTribCBS_Gov'] = val(gov, 'vTribCBS')

    # ---------- gIBSCBSMono (CST 620) ----------
    mono = ibscbs.find('gIBSCBSMono')
    if mono is not None:
        d['Grupo_Tributo_RTC'] = 'gIBSCBSMono'
        # v1.30 introduziu o subgrupo gMonoPadrao; versoes anteriores traziam
        # as tags soltas dentro de gIBSCBSMono. Aceita as duas formas.
        padrao = mono.find('gMonoPadrao')
        base = padrao if padrao is not None else mono
        d['qBCMono'] = val(base, 'qBCMono')
        d['adRemIBS'] = perc(base, 'adRemIBS')
        d['adRemCBS'] = perc(base, 'adRemCBS')
        d['vIBSMono'] = val(base, 'vIBSMono')
        d['vCBSMono'] = val(base, 'vCBSMono')

        reten = sub(mono, 'gMonoReten')
        if reten is not None:
            d['qBCMonoReten'] = val(reten, 'qBCMonoReten')
            d['adRemIBSReten'] = perc(reten, 'adRemIBSReten')
            d['vIBSMonoReten'] = val(reten, 'vIBSMonoReten')
            d['adRemCBSReten'] = perc(reten, 'adRemCBSReten')
            d['vCBSMonoReten'] = val(reten, 'vCBSMonoReten')

        ret = sub(mono, 'gMonoRet')
        if ret is not None:
            d['qBCMonoRet'] = val(ret, 'qBCMonoRet')
            d['adRemIBSRet'] = perc(ret, 'adRemIBSRet')
            d['vIBSMonoRet'] = val(ret, 'vIBSMonoRet')
            d['adRemCBSRet'] = perc(ret, 'adRemCBSRet')
            d['vCBSMonoRet'] = val(ret, 'vCBSMonoRet')

        dif = sub(mono, 'gMonoDif')
        if dif is not None:
            d['pDifIBS_Mono'] = perc(dif, 'pDifIBS')
            d['vIBSMonoDif'] = val(dif, 'vIBSMonoDif')
            d['pDifCBS_Mono'] = perc(dif, 'pDifCBS')
            d['vCBSMonoDif'] = val(dif, 'vCBSMonoDif')

        d['vTotIBSMonoItem'] = val(mono, 'vTotIBSMonoItem')
        d['vTotCBSMonoItem'] = val(mono, 'vTotCBSMonoItem')

    # ---------- gTransfCred (CST 800) ----------
    tc = ibscbs.find('gTransfCred')
    if tc is not None:
        d['Grupo_Tributo_RTC'] = 'gTransfCred'
        d['vIBS_TransfCred'] = val(tc, 'vIBS')
        d['vCBS_TransfCred'] = val(tc, 'vCBS')

    # ---------- gAjusteCompet (CST 811) ----------
    ac = ibscbs.find('gAjusteCompet')
    if ac is not None:
        d['Grupo_Tributo_RTC'] = 'gAjusteCompet'
        d['competApur_Ajuste'] = get_text(ac, 'competApur')
        d['vIBS_AjusteCompet'] = val(ac, 'vIBS')
        d['vCBS_AjusteCompet'] = val(ac, 'vCBS')

    # ---------- gEstornoCred (UB116) ----------
    ec = ibscbs.find('gEstornoCred')
    if ec is not None:
        d['vIBSEstCred'] = val(ec, 'vIBSEstCred')
        d['vCBSEstCred'] = val(ec, 'vCBSEstCred')

    # ---------- gCredPresOper (UB120) ----------
    cp = ibscbs.find('gCredPresOper')
    if cp is not None:
        d['vBCCredPres'] = val(cp, 'vBCCredPres')
        d['cCredPres'] = get_text(cp, 'cCredPres')
        d['cCredPres_Desc'] = CCREDPRES.get(d['cCredPres'].zfill(2), '')
        cp_ibs = cp.find('gIBSCredPres')
        if cp_ibs is not None:
            d['pCredPres_IBS'] = perc(cp_ibs, 'pCredPres')
            d['vCredPres_IBS'] = val(cp_ibs, 'vCredPres')
            d['vCredPresCondSus_IBS'] = val(cp_ibs, 'vCredPresCondSus')
        cp_cbs = cp.find('gCBSCredPres')
        if cp_cbs is not None:
            d['pCredPres_CBS'] = perc(cp_cbs, 'pCredPres')
            d['vCredPres_CBS'] = val(cp_cbs, 'vCredPres')
            d['vCredPresCondSus_CBS'] = val(cp_cbs, 'vCredPresCondSus')

    # ---------- gCredPresIBSZFM (CST 810) ----------
    zfm = ibscbs.find('gCredPresIBSZFM')
    if zfm is not None:
        d['competApur_ZFM'] = get_text(zfm, 'competApur')
        d['tpCredPresIBSZFM'] = get_text(zfm, 'tpCredPresIBSZFM')
        d['vCredPres_ZFM'] = val(zfm, 'vCredPres')

    if not d['Grupo_Tributo_RTC']:
        d['Grupo_Tributo_RTC'] = 'sem grupo de tributo'

    # ---------- consolidado do item ----------
    d['Total_IBS_Item'] = round(d['vIBS'] + d['vTotIBSMonoItem'], 2)
    d['Total_CBS_Item'] = round(d['vCBS'] + d['vTotCBSMonoItem'], 2)
    d['Total_IBSCBS_Item'] = round(d['Total_IBS_Item'] + d['Total_CBS_Item'], 2)

    return d


# ====================================================================================
# DIAGNOSTICO (conferencias nao bloqueantes)
# ====================================================================================

def diagnosticar_rtc(d, ano, crt):
    """Verifica coerencia CST x grupos x calculo. Retorna lista de avisos."""
    avisos = []
    cst = d.get('CST_IBSCBS', '')
    grupo = d.get('Grupo_Tributo_RTC', '')

    if not cst:
        if crt == '3' and ano and ano >= 2026:
            avisos.append('Item sem grupo IBSCBS (obrigatorio p/ CRT=3 a partir de 2026)')
        return avisos

    if cst not in CST_IBSCBS:
        avisos.append(f'CST {cst} fora da tabela oficial')

    cc = d.get('cClassTrib', '')
    if cc and cc not in CCLASSTRIB:
        avisos.append(f'cClassTrib {cc} fora da tabela oficial')
    elif cc and not cc.startswith(cst):
        avisos.append(f'cClassTrib {cc} incompativel com CST {cst}')

    # grupo de tributo exigido pelo CST
    if cst in CST_EXIGE_GIBSCBS and grupo != 'gIBSCBS':
        avisos.append(f'CST {cst} exige o grupo gIBSCBS')
    if cst in CST_EXIGE_MONO and grupo != 'gIBSCBSMono':
        avisos.append(f'CST {cst} exige o grupo gIBSCBSMono')
    if cst in CST_EXIGE_TRANSFCRED and grupo != 'gTransfCred':
        avisos.append(f'CST {cst} exige o grupo gTransfCred')
    if cst in CST_EXIGE_AJUSTECOMPET and grupo != 'gAjusteCompet':
        avisos.append(f'CST {cst} exige o grupo gAjusteCompet')
    if cst in CST_EXIGE_CREDPRESZFM and not d.get('tpCredPresIBSZFM'):
        avisos.append(f'CST {cst} exige o grupo gCredPresIBSZFM')
    if cst in CST_SEM_GRUPO_TRIBUTO and grupo in ('gIBSCBS', 'gIBSCBSMono'):
        avisos.append(f'CST {cst} nao admite grupo de tributacao')

    # gRed / gDif
    if cst in CST_EXIGE_GRED and grupo == 'gIBSCBS':
        if d['pRedAliq_IBSUF'] == 0 and d['pAliqEfet_IBSUF'] == 0 \
                and d['pRedAliq_CBS'] == 0 and d['pAliqEfet_CBS'] == 0:
            avisos.append(f'CST {cst} exige o grupo gRed (reducao de aliquota)')
    if cst in CST_EXIGE_GDIF and grupo == 'gIBSCBS':
        if d['pDif_IBSUF'] == 0 and d['pDif_CBS'] == 0:
            avisos.append(f'CST {cst} exige o grupo gDif (diferimento)')

    if grupo != 'gIBSCBS':
        return avisos

    fator = 100 if not PERC_RTC_COMO_FRACAO else 1

    # vIBS = vIBSUF + vIBSMun
    soma = d['vIBSUF'] + d['vIBSMun']
    if abs(soma - d['vIBS']) > 0.01:
        avisos.append(f"vIBS ({d['vIBS']:.2f}) != vIBSUF+vIBSMun ({soma:.2f})")

    # vIBSUF = vBC x aliquota efetiva - vDif - vDevTrib
    for pref, tag_v in (('IBSUF', 'vIBSUF'), ('IBSMun', 'vIBSMun'), ('CBS', 'vCBS')):
        aliq = d[f'pAliqEfet_{pref}'] if d[f'pRedAliq_{pref}'] or d[f'pAliqEfet_{pref}'] else d[f'p{pref}']
        esperado = d['vBC_IBSCBS'] * aliq / fator - d[f'vDif_{pref}'] - d[f'vDevTrib_{pref}']
        if abs(esperado - d[tag_v]) > TOLERANCIA:
            avisos.append(f'{tag_v} informado {d[tag_v]:.2f}, calculado {esperado:.2f}')

    # aliquotas de referencia do periodo de transicao
    ref = ALIQ_TRANSICAO.get(ano)
    if ref and cst not in CST_SEM_GRUPO_TRIBUTO and not d.get('CSTReg'):
        alvo_ibs = ref['IBSUF'] if not PERC_RTC_COMO_FRACAO else ref['IBSUF'] / 100
        if d['pIBSUF'] and abs(d['pIBSUF'] - alvo_ibs) > 0.0001:
            avisos.append(f"pIBSUF {d['pIBSUF']} diverge da aliquota de {ano} ({alvo_ibs})")
        if ref['CBS'] is not None:
            alvo_cbs = ref['CBS'] if not PERC_RTC_COMO_FRACAO else ref['CBS'] / 100
            if d['pCBS'] and abs(d['pCBS'] - alvo_cbs) > 0.0001:
                avisos.append(f"pCBS {d['pCBS']} diverge da aliquota de {ano} ({alvo_cbs})")

    return avisos


# ====================================================================================
# PROCESSAMENTO PRINCIPAL
# ====================================================================================

def processar_nfe(inf, filename, rows, error_log):
    """Processa uma <infNFe> (uma NF-e) e acrescenta uma linha por item."""

    ide = inf.find('ide')
    emit = inf.find('emit')
    dest = inf.find('dest')
    total = inf.find('total')

    num = get_text(ide, 'nNF')
    dh_full = get_text(ide, 'dhEmi') or get_text(ide, 'dEmi')
    dh = dh_full.split('T')[0]
    try:
        ano = int(dh[:4])
    except (ValueError, TypeError):
        ano = None
    chave = inf.get('Id')[-44:] if inf.get('Id') else ''
    crt = get_text(emit, 'CRT')
    mod = get_text(ide, 'mod')
    serie = get_text(ide, 'serie')
    tpNF = get_text(ide, 'tpNF')
    finNFe = get_text(ide, 'finNFe')

    # --- campos de cabecalho da reforma ---
    cMunFGIBS = get_text(ide, 'cMunFGIBS')
    tpNFDebito = get_text(ide, 'tpNFDebito')
    tpNFCredito = get_text(ide, 'tpNFCredito')
    cIndOp = get_text(ide, 'cIndOp')
    dPrevEntrega = get_text(ide, 'dPrevEntrega')

    gov = ide.find('gCompraGov') if ide is not None else None
    tpEnteGov = get_text(gov, 'tpEnteGov')
    pRedutorGov = perc(gov, 'pRedutor') if gov is not None else 0.0

    emit_cnpj = get_text(emit, 'CNPJ') or get_text(emit, 'CPF')
    emit_nome = get_text(emit, 'xNome')
    emit_uf = get_text(emit, 'enderEmit/UF')
    dest_cnpj = get_text(dest, 'CNPJ') or get_text(dest, 'CPF')
    dest_nome = get_text(dest, 'xNome')
    dest_uf = get_text(dest, 'enderDest/UF')

    # --- totais da nota (Grupo W03) ---
    tot_rtc = total.find('IBSCBSTot') if total is not None else None
    vBCIBSCBS_Tot = val(tot_rtc, 'vBCIBSCBS')
    tot_gibs = tot_rtc.find('gIBS') if tot_rtc is not None else None
    vIBS_Tot = val(tot_gibs, 'vIBS')
    tot_gcbs = tot_rtc.find('gCBS') if tot_rtc is not None else None
    vCBS_Tot = val(tot_gcbs, 'vCBS')
    tot_gmono = tot_rtc.find('gMono') if tot_rtc is not None else None
    vIBSMono_Tot = val(tot_gmono, 'vIBSMono')
    vCBSMono_Tot = val(tot_gmono, 'vCBSMono')
    vNFTot = val(total, 'vNFTot')
    vIS_Tot = val(total.find('ISTot'), 'vIS') if total is not None else 0.0

    dets = inf.findall('.//det')
    if not dets:
        error_log.append({
            'arquivo': filename, 'chave': chave, 'item': '-',
            'erro': 'NF-e sem itens <det> (XML truncado ou incompleto).'
        })
        return

    for prod in dets:
        try:
            imposto = prod.find('imposto')
            n_item = prod.get('nItem')

            if imposto is None:
                error_log.append({
                    'arquivo': filename, 'chave': chave, 'item': n_item,
                    'erro': 'Grupo <imposto> nao encontrado no item.'
                })
                continue

            # ------------------ ICMS ------------------
            icms_node = imposto.find('ICMS')
            icms_group = None
            cst_icms = ''
            if icms_node is not None and len(icms_node) > 0:
                icms_group = icms_node[0]
                orig = get_text(icms_group, 'orig')
                cst = get_text(icms_group, 'CST')
                csosn = get_text(icms_group, 'CSOSN')
                cst_icms = orig + (cst if cst else csosn)

            vBC = val(icms_group, 'vBC')
            pICMS = val(icms_group, 'pICMS') / 100
            vICMS = val(icms_group, 'vICMS')
            vFCP = val(icms_group, 'vFCP')
            pMV = val(icms_group, 'pMVAST') / 100
            vBCST = val(icms_group, 'vBCST')
            pICMSST = val(icms_group, 'pICMSST') / 100
            vICMSST = val(icms_group, 'vICMSST')
            vFCPST = val(icms_group, 'vFCPST')
            vICMSDeson = val(icms_group, 'vICMSDeson')

            desonerado_abate = 'SIM' if get_text(icms_group, 'indDeduzDeson') == '1' else 'NAO'

            # ------------------ Produto ------------------
            vProd = val(prod, 'prod/vProd')
            vFre = val(prod, 'prod/vFrete')
            vSeg = val(prod, 'prod/vSeg')
            vDesc = val(prod, 'prod/vDesc')
            vOut = val(prod, 'prod/vOutro')

            # ------------------ IPI ------------------
            ipi_node = imposto.find('IPI')
            vIPI = val(ipi_node, 'IPITrib/vIPI')
            vBC_IPI = val(ipi_node, 'IPITrib/vBC')
            pIPI = val(ipi_node, 'IPITrib/pIPI') / 100

            # ------------------ PIS / COFINS ------------------
            pis_node = imposto.find('PIS')
            cst_pis = ''
            vPIS = 0.0
            if pis_node is not None and len(pis_node) > 0:
                pis_group = pis_node[0]
                cst_pis = get_text(pis_group, 'CST')
                el = pis_node.find('.//vPIS')
                vPIS = to_float(el.text) if el is not None else 0.0

            cof_node = imposto.find('COFINS')
            cst_cofins = ''
            vCOF = 0.0
            if cof_node is not None and len(cof_node) > 0:
                cof_group = cof_node[0]
                cst_cofins = get_text(cof_group, 'CST')
                el = cof_node.find('.//vCOFINS')
                vCOF = to_float(el.text) if el is not None else 0.0

            vII = val(imposto, 'II/vII')

            # ------------------ Calculos ------------------
            tot_no_ipi = vProd + vFre + vSeg - vDesc + vOut
            base_pis = tot_no_ipi + vIPI
            base_cof = tot_no_ipi + vICMSST + vFCPST + vIPI

            qCom = val(prod, 'prod/qCom')
            unit_conv = base_cof / (qCom or 1)

            cEANTrib = get_text(prod, 'prod/cEANTrib')
            cEAN = get_text(prod, 'prod/cEAN')
            cmpEAN = 'Igual' if cEANTrib == cEAN else 'Diferente'

            custo_total = (vProd + vFre + vSeg) - vDesc + (vOut + vICMSST + vFCPST + vIPI)
            if desonerado_abate == 'SIM':
                custo_total -= vICMSDeson

            # ------------------ Reforma Tributaria (Grupo UB) ------------------
            rtc = extrair_ibscbs(imposto, prod)
            rtc['vItem'] = val(prod, 'vItem')
            rtc['indBemMovelUsado'] = get_text(prod, 'prod/indBemMovelUsado')
            rtc['tpCredPresIBSZFM_Prod'] = get_text(prod, 'prod/tpCredPresIBSZFM')

            avisos = diagnosticar_rtc(rtc, ano, crt)
            rtc['Diag_RTC'] = ' | '.join(avisos)

            # Custo simulado com IBS/CBS (na transicao o IBS/CBS NAO integra o
            # total do item nem o total da nota - fica so para simulacao)
            custo_com_rtc = custo_total + rtc['Total_IBSCBS_Item'] + rtc['vIS']

            linha = {
                'CRT': crt, 'mod': mod, 'serie': serie, 'tpNF': tpNF, 'finNFe': finNFe,
                'Chave_NFe': chave, 'Numero_NFe': num, 'Data_Emis': dh,
                'Emit_CNPJ': emit_cnpj, 'Emit_Nome': emit_nome, 'Emit_UF': emit_uf,
                'Dest_CNPJ': dest_cnpj, 'Dest_Nome': dest_nome, 'Dest_UF': dest_uf,
                'nItem': n_item, 'cProd': get_text(prod, 'prod/cProd'),
                'xProd': get_text(prod, 'prod/xProd'),
                'NCM': format_ncm(get_text(prod, 'prod/NCM')),
                'CEST': format_cest(get_text(prod, 'prod/CEST')),
                'CFOP': format_cfop(get_text(prod, 'prod/CFOP')),
                'CST_ICMS': cst_icms, 'vProd': vProd, 'vBC': vBC, 'pICMS': pICMS,
                'vICMS': vICMS, 'vFCP': vFCP, 'vFrete': vFre, 'vSeg': vSeg,
                'vDesc': vDesc, 'vOutro': vOut, 'pMVAST': pMV, 'vBCST': vBCST,
                'pICMSST': pICMSST, 'vICMSST': vICMSST, 'vFCPST': vFCPST,
                'vICMSDeson': vICMSDeson, 'Desonerado_Abate': desonerado_abate,
                'vIPI': vIPI, 'vBC_IPI': vBC_IPI, 'pIPI': pIPI,
                'CST_PIS': cst_pis, 'vPIS': vPIS,
                'CST_COFINS': cst_cofins, 'vCOFINS': vCOF, 'vII': vII,
                'Total_s_IPI': tot_no_ipi, 'Base_PIS_COFINS': base_pis,
                'Base_COFINS': base_cof, 'uCom': get_text(prod, 'prod/uCom'),
                'qCom': qCom, 'vUnitConv': unit_conv,
                'cEANTrib': cEANTrib, 'cEAN': cEAN, 'Comparacao_EAN': cmpEAN,
                'Custo_Total_da_Mercadoria': custo_total,
                'Custo_Total_com_IBSCBS_IS': custo_com_rtc,
                # cabecalho reforma
                'cMunFGIBS': cMunFGIBS, 'tpNFDebito': tpNFDebito,
                'tpNFCredito': tpNFCredito, 'cIndOp': cIndOp,
                'dPrevEntrega': dPrevEntrega,
                'tpEnteGov': tpEnteGov, 'pRedutorGov': pRedutorGov,
                # totais da nota
                'vBCIBSCBS_Tot': vBCIBSCBS_Tot, 'vIBS_Tot': vIBS_Tot,
                'vCBS_Tot': vCBS_Tot, 'vIBSMono_Tot': vIBSMono_Tot,
                'vCBSMono_Tot': vCBSMono_Tot, 'vIS_Tot': vIS_Tot,
                'vNFTot': vNFTot,
                'arquivo_origem': filename,
            }
            linha.update(rtc)
            rows.append(linha)

        except Exception as e:
            error_log.append({
                'arquivo': filename, 'chave': chave,
                'item': prod.get('nItem'),
                'erro': f'Erro ao processar item: {type(e).__name__}: {e}'
            })


def process_xml_streams(xml_streams):
    """Processa uma lista de tuplas (filename, file_stream)."""
    rows = []
    error_log = []
    total_files = len(xml_streams)
    if total_files == 0:
        return pd.DataFrame(), error_log

    progress_bar = st.progress(0, text='Processando XMLs...')
    parser = ET.XMLParser(remove_blank_text=True, recover=True, huge_tree=True,
                          resolve_entities=False)

    for i, (filename, xml_stream) in enumerate(xml_streams):
        try:
            xml_stream.seek(0)
            tree = ET.parse(xml_stream, parser)
            root = tree.getroot()
            if root is None:
                raise ValueError('XML vazio ou ilegivel.')
            root = strip_namespace(root)

            # Um arquivo pode conter mais de uma NF-e (lotes / resNFe).
            infs = root.findall('.//infNFe')
            if not infs:
                tipo = root.tag
                error_log.append({
                    'arquivo': filename, 'chave': '', 'item': '-',
                    'erro': f'Sem tag <infNFe> (raiz <{tipo}>). Provavel evento, '
                            f'cancelamento, inutilizacao ou XML de outro modelo.'
                })
                continue

            for inf in infs:
                processar_nfe(inf, filename, rows, error_log)

        except ET.XMLSyntaxError as e:
            error_log.append({'arquivo': filename, 'chave': '', 'item': '-',
                              'erro': f'XML mal formado: {e}'})
        except Exception as e:
            error_log.append({'arquivo': filename, 'chave': '', 'item': '-',
                              'erro': f'Erro inesperado: {type(e).__name__}: {e}'})

        progress_bar.progress((i + 1) / total_files,
                              text=f'Processando XMLs... ({i + 1}/{total_files})')

    progress_bar.empty()
    df = pd.DataFrame(rows)

    if not df.empty:
        cols = [
            # identificacao
            'CRT', 'mod', 'serie', 'tpNF', 'finNFe', 'Chave_NFe', 'Numero_NFe', 'Data_Emis',
            'Emit_CNPJ', 'Emit_Nome', 'Emit_UF', 'Dest_CNPJ', 'Dest_Nome', 'Dest_UF',
            'nItem', 'cProd', 'xProd', 'NCM', 'CEST', 'CFOP',
            # ICMS / IPI / PIS / COFINS (tributos atuais)
            'CST_ICMS', 'vProd', 'vBC', 'pICMS', 'vICMS', 'vFCP', 'vFrete', 'vSeg',
            'vDesc', 'vOutro', 'pMVAST', 'vBCST', 'pICMSST', 'vICMSST', 'vFCPST',
            'vICMSDeson', 'Desonerado_Abate', 'vIPI', 'vBC_IPI', 'pIPI',
            'CST_PIS', 'vPIS', 'CST_COFINS', 'vCOFINS', 'vII',
            'Total_s_IPI', 'Base_PIS_COFINS', 'Base_COFINS',
            'uCom', 'qCom', 'vUnitConv', 'cEANTrib', 'cEAN', 'Comparacao_EAN',
            'Custo_Total_da_Mercadoria',
            # ---------------- REFORMA TRIBUTARIA ----------------
            'CST_IBSCBS', 'CST_IBSCBS_Desc', 'cClassTrib', 'cClassTrib_Desc',
            'pRed_Tabela_IBS', 'pRed_Tabela_CBS', 'indDoacao', 'Grupo_Tributo_RTC',
            'vBC_IBSCBS',
            'pIBSUF', 'pRedAliq_IBSUF', 'pAliqEfet_IBSUF', 'pDif_IBSUF', 'vDif_IBSUF',
            'pDevTrib_IBSUF', 'vDevTrib_IBSUF', 'vIBSUF',
            'pIBSMun', 'pRedAliq_IBSMun', 'pAliqEfet_IBSMun', 'pDif_IBSMun', 'vDif_IBSMun',
            'pDevTrib_IBSMun', 'vDevTrib_IBSMun', 'vIBSMun',
            'vIBS',
            'pCBS', 'pRedAliq_CBS', 'pAliqEfet_CBS', 'pDif_CBS', 'vDif_CBS',
            'pDevTrib_CBS', 'vDevTrib_CBS', 'vCBS',
            'tpALCZFMCBS', 'pAliqEfeRegCBS_ALCZFM', 'vTribRegCBS_ALCZFM',
            'CSTReg', 'cClassTribReg', 'pAliqEfetRegIBSUF', 'vTribRegIBSUF',
            'pAliqEfetRegIBSMun', 'vTribRegIBSMun', 'pAliqEfetRegCBS', 'vTribRegCBS',
            'pAliqIBSUF_Gov', 'vTribIBSUF_Gov', 'pAliqIBSMun_Gov', 'vTribIBSMun_Gov',
            'pAliqCBS_Gov', 'vTribCBS_Gov',
            'qBCMono', 'adRemIBS', 'adRemCBS', 'vIBSMono', 'vCBSMono',
            'qBCMonoReten', 'adRemIBSReten', 'vIBSMonoReten', 'adRemCBSReten', 'vCBSMonoReten',
            'qBCMonoRet', 'adRemIBSRet', 'vIBSMonoRet', 'adRemCBSRet', 'vCBSMonoRet',
            'pDifIBS_Mono', 'vIBSMonoDif', 'pDifCBS_Mono', 'vCBSMonoDif',
            'vTotIBSMonoItem', 'vTotCBSMonoItem',
            'vIBS_TransfCred', 'vCBS_TransfCred',
            'competApur_Ajuste', 'vIBS_AjusteCompet', 'vCBS_AjusteCompet',
            'vIBSEstCred', 'vCBSEstCred',
            'vBCCredPres', 'cCredPres', 'cCredPres_Desc',
            'pCredPres_IBS', 'vCredPres_IBS', 'vCredPresCondSus_IBS',
            'pCredPres_CBS', 'vCredPres_CBS', 'vCredPresCondSus_CBS',
            'competApur_ZFM', 'tpCredPresIBSZFM', 'vCredPres_ZFM',
            'CST_IS', 'cClassTrib_IS', 'vBC_IS', 'pIS', 'adRemIS',
            'uTrib_IS', 'qTrib_IS', 'vIS',
            'vItem', 'indBemMovelUsado', 'tpCredPresIBSZFM_Prod',
            'Total_IBS_Item', 'Total_CBS_Item', 'Total_IBSCBS_Item',
            'Custo_Total_com_IBSCBS_IS',
            # cabecalho / totais da nota
            'cMunFGIBS', 'tpNFDebito', 'tpNFCredito', 'cIndOp', 'dPrevEntrega',
            'tpEnteGov', 'pRedutorGov',
            'vBCIBSCBS_Tot', 'vIBS_Tot', 'vCBS_Tot', 'vIBSMono_Tot', 'vCBSMono_Tot',
            'vIS_Tot', 'vNFTot',
            'Diag_RTC', 'arquivo_origem',
        ]
        cols_existentes = [c for c in cols if c in df.columns]
        restantes = [c for c in df.columns if c not in cols_existentes]
        df = df[cols_existentes + restantes]

    return df, error_log


# ====================================================================================
# EXPORTACAO
# ====================================================================================

COLS_RTC = [c for c in colunas_rtc_vazias().keys()]


def remover_colunas_rtc_vazias(df):
    """Remove colunas do bloco RTC que estao 100% vazias/zeradas no lote."""
    if df.empty:
        return df
    dropar = []
    for c in COLS_RTC:
        if c not in df.columns or c == 'Diag_RTC':
            continue
        serie = df[c]
        if serie.dtype == object:
            if serie.fillna('').astype(str).str.strip().eq('').all():
                dropar.append(c)
        else:
            if (serie.fillna(0) == 0).all():
                dropar.append(c)
    return df.drop(columns=dropar)


def resumo_rtc(df):
    """Resumo por CST/cClassTrib do IBS e da CBS."""
    if df.empty or 'CST_IBSCBS' not in df.columns:
        return pd.DataFrame()
    base = df[df['CST_IBSCBS'].astype(str).str.strip() != '']
    if base.empty:
        return pd.DataFrame()
    g = base.groupby(['CST_IBSCBS', 'CST_IBSCBS_Desc', 'cClassTrib', 'cClassTrib_Desc'],
                     dropna=False).agg(
        Itens=('vProd', 'size'),
        vProd=('vProd', 'sum'),
        vBC_IBSCBS=('vBC_IBSCBS', 'sum'),
        vIBSUF=('vIBSUF', 'sum'),
        vIBSMun=('vIBSMun', 'sum'),
        vIBS=('vIBS', 'sum'),
        vCBS=('vCBS', 'sum'),
        vIBSMono=('vTotIBSMonoItem', 'sum'),
        vCBSMono=('vTotCBSMonoItem', 'sum'),
        vIS=('vIS', 'sum'),
    ).reset_index()
    return g.sort_values(['CST_IBSCBS', 'cClassTrib'])


def to_excel(df, df_resumo, df_erros, df_diag):
    """Gera o arquivo Excel em memoria com todas as abas."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt_moeda = wb.add_format({'num_format': '#,##0.00'})
        fmt_qtd = wb.add_format({'num_format': '#,##0.0000'})
        fmt_head = wb.add_format({'bold': True, 'bg_color': '#1F3864',
                                  'font_color': 'white', 'border': 1,
                                  'align': 'center', 'valign': 'vcenter',
                                  'text_wrap': True})

        def escrever(dframe, aba):
            if dframe is None or dframe.empty:
                return
            dframe.to_excel(writer, index=False, sheet_name=aba, startrow=1, header=False)
            ws = writer.sheets[aba]
            for j, col in enumerate(dframe.columns):
                ws.write(0, j, col, fmt_head)
                try:
                    largura = int(dframe[col].astype(str).map(len).max())
                except (ValueError, TypeError):
                    largura = 12
                largura = max(min(max(largura, len(str(col))) + 2, 45), 10)
                nome = str(col)
                if nome.startswith('v') or nome.startswith('Total') or nome.startswith('Base') \
                        or nome.startswith('Custo'):
                    ws.set_column(j, j, largura, fmt_moeda)
                elif nome.startswith('q') or nome.startswith('adRem'):
                    ws.set_column(j, j, largura, fmt_qtd)
                else:
                    ws.set_column(j, j, largura)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, len(dframe), len(dframe.columns) - 1)

        escrever(df, 'Relatorio')
        escrever(df_resumo, 'Resumo IBS-CBS')
        escrever(df_diag, 'Divergencias RTC')
        escrever(df_erros, 'Erros')

    return output.getvalue()


# ====================================================================================
# INTERFACE STREAMLIT
# ====================================================================================

st.set_page_config(page_title='Relatorio NF-e - Reforma Tributaria', layout='wide')
st.title('📄 Gerador de Relatorio de NF-e/NFC-e (v3.0 - IBS/CBS/IS)')
st.markdown(
    'Extrai dados de `.xml`, `.zip` ou `.rar` e consolida em Excel, '
    'incluindo todos os grupos da **NT 2025.002-RTC** (leiaute ate a v1.40).'
)

with st.sidebar:
    st.header('Opcoes')
    ocultar_vazias = st.checkbox(
        'Ocultar colunas IBS/CBS/IS vazias', value=True,
        help='Remove do relatorio as colunas da reforma que ficaram 100% zeradas no lote.'
    )
    deduplicar = st.checkbox(
        'Remover itens duplicados (mesma chave + item)', value=True,
        help='Util quando o mesmo XML aparece em mais de um ZIP.'
    )
    st.caption(f'Percentuais RTC como fracao: **{PERC_RTC_COMO_FRACAO}**')
    if not RARFILE_OK:
        st.warning('Biblioteca `rarfile` nao instalada: arquivos .rar serao ignorados.')

tipos = ['xml', 'zip'] + (['rar'] if RARFILE_OK else [])
uploaded_files = st.file_uploader(
    'Selecione os arquivos XML, ZIP' + (' ou RAR' if RARFILE_OK else ''),
    type=tipos, accept_multiple_files=True,
    help='Voce pode arrastar multiplos arquivos ou compactados.'
)

if uploaded_files:
    if st.button('🚀 Gerar Relatorio', type='primary'):
        start_time = time.time()
        xml_streams_to_process = []
        unpack_error_log = []

        with st.spinner('Preparando e descompactando arquivos...'):
            for file in uploaded_files:
                filename = file.name
                ext = filename.lower().rsplit('.', 1)[-1]
                try:
                    if ext == 'xml':
                        xml_streams_to_process.append((filename, file))

                    elif ext == 'zip':
                        with zipfile.ZipFile(file, 'r') as zf:
                            for nome in zf.namelist():
                                if nome.lower().endswith('.xml'):
                                    xml_streams_to_process.append(
                                        (f'{filename}/{nome}', BytesIO(zf.read(nome)))
                                    )

                    elif ext == 'rar' and RARFILE_OK:
                        try:
                            with rarfile.RarFile(file, 'r') as rf:
                                for nome in rf.namelist():
                                    if nome.lower().endswith('.xml'):
                                        xml_streams_to_process.append(
                                            (f'{filename}/{nome}', BytesIO(rf.read(nome)))
                                        )
                        except rarfile.NeedFirstVolume:
                            unpack_error_log.append(
                                {'arquivo': filename, 'chave': '', 'item': '-',
                                 'erro': 'RAR multi-volume nao suportado.'})
                        except rarfile.RarCannotExec:
                            unpack_error_log.append(
                                {'arquivo': filename, 'chave': '', 'item': '-',
                                 'erro': 'Utilitario unrar nao instalado no servidor.'})
                        except Exception as e:
                            unpack_error_log.append(
                                {'arquivo': filename, 'chave': '', 'item': '-',
                                 'erro': f'Erro ao ler RAR: {e}'})

                except Exception as e:
                    unpack_error_log.append(
                        {'arquivo': filename, 'chave': '', 'item': '-',
                         'erro': f'Erro ao abrir arquivo: {e}'})

        if not xml_streams_to_process:
            st.warning('Nenhum arquivo XML valido foi encontrado nos uploads.')
            if unpack_error_log:
                st.error('Erros ao descompactar:')
                st.dataframe(pd.DataFrame(unpack_error_log), use_container_width=True)
        else:
            st.info(f'Arquivos preparados. Processando {len(xml_streams_to_process)} XMLs...')
            df_report, process_errors = process_xml_streams(xml_streams_to_process)
            all_errors = unpack_error_log + process_errors
            total_time = time.time() - start_time

            if df_report.empty:
                st.warning('Nenhum dado foi extraido. Verifique os XMLs.')
            else:
                if deduplicar:
                    antes = len(df_report)
                    df_report = df_report.drop_duplicates(
                        subset=['Chave_NFe', 'nItem'], keep='first')
                    if antes != len(df_report):
                        st.caption(f'{antes - len(df_report)} itens duplicados removidos.')

                df_resumo = resumo_rtc(df_report)
                df_diag = df_report[df_report['Diag_RTC'].astype(str) != ''][
                    ['Chave_NFe', 'Numero_NFe', 'Data_Emis', 'nItem', 'xProd',
                     'CST_IBSCBS', 'cClassTrib', 'Diag_RTC']
                ] if 'Diag_RTC' in df_report.columns else pd.DataFrame()

                df_exibir = remover_colunas_rtc_vazias(df_report) if ocultar_vazias else df_report

                st.success(
                    f'Relatorio gerado. {len(df_report)} itens em {total_time:.2f}s '
                    f'({len(df_exibir.columns)} colunas).'
                )

                c1, c2, c3, c4 = st.columns(4)
                c1.metric('Itens', len(df_report))
                c2.metric('Notas', df_report['Chave_NFe'].nunique())
                c3.metric('IBS total', f"R$ {df_report['Total_IBS_Item'].sum():,.2f}")
                c4.metric('CBS total', f"R$ {df_report['Total_CBS_Item'].sum():,.2f}")

                aba1, aba2, aba3 = st.tabs(['Relatorio', 'Resumo IBS/CBS', 'Divergencias'])
                with aba1:
                    st.dataframe(df_exibir, use_container_width=True)
                with aba2:
                    if df_resumo.empty:
                        st.info('Nenhum item com grupo IBS/CBS informado neste lote.')
                    else:
                        st.dataframe(df_resumo, use_container_width=True)
                with aba3:
                    if df_diag.empty:
                        st.success('Nenhuma divergencia encontrada nos grupos de IBS/CBS.')
                    else:
                        st.warning(f'{len(df_diag)} itens com possiveis inconsistencias.')
                        st.dataframe(df_diag, use_container_width=True)

                excel_data = to_excel(
                    df_exibir, df_resumo,
                    pd.DataFrame(all_errors) if all_errors else pd.DataFrame(),
                    df_diag
                )
                st.download_button(
                    label='📥 Baixar Relatorio em Excel',
                    data=excel_data,
                    file_name='relatorio_nfe_rtc.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            if all_errors:
                st.error(f'{len(all_errors)} ocorrencia(s) durante o processo:')
                with st.expander('Ver detalhes'):
                    st.dataframe(pd.DataFrame(all_errors), use_container_width=True)

else:
    st.info('Aguardando o upload de arquivos `.xml`, `.zip`' +
            (' ou `.rar`.' if RARFILE_OK else '.'))

import io
import os
import traceback
import pandas as pd
from io import StringIO
from datetime import datetime
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================
# VERSÃO
# ==============================
VERSAO          = "V2.6"
NOME_CFOP_XLSX  = "160314_Tabela_CFOP.xlsx"
NOME_IMP_XLSX   = "Impostos.xlsx"

# ==============================
# MAPEAMENTO FIXO: nome/fragmento → código Domínio
# Usado como fallback se a planilha não carregar
# ==============================
IMPOSTOS_FALLBACK = {
    'ICMS':                     1,
    'IPI':                      2,
    'ISS':                      3,
    'PIS':                      4,
    'COFINS':                   5,
    'CONTRIBUICAO SOCIAL':      6,
    'DIFAL':                    8,
    'SUBST. TRIBUTARIA':        9,
    'SIMPLES':                 10,
    'ICMS RETIDO':             11,
    'ICMS SUBSTITUTO':         12,
    'PIS NÃO CUMULATIVO':      17,
    'COFINS NÃO CUMULATIVO':   19,
    'INSS RETIDO':             26,
    'ICMS ANTECIPADO':         27,
    'SIMPLES NACIONAL':        44,
    'ICMS IMPORTACAO':         45,
    'PIS IMPORTACAO':         133,
    'COFINS IMPORTACAO':      134,
    'ICMS DIFERIDO':          116,
    'ICMS COMPLEMENTAR':      125,
    'DIFAL NÃO CONTRIBUINTE': 145,
    'DIFAL FCP':              146,
    'IBS':                    183,
    'CBS':                    184,
}

# ==============================
# ALÍQUOTAS PIS/COFINS → CÓDIGO DOMÍNIO
# Regra: alíquota determina regime (cumulativo x não cumulativo)
# ==============================
# Tolerância para comparação de float
_ALIQ_TOL = 0.001

# Mapa: (alíquota_percentual, tolerância) → código Domínio
# Formato: (valor_central, tolerância) → codigo
ALIQ_PIS_COFINS = {
    # PIS
    'PIS': [
        (0.65,  _ALIQ_TOL,  4),   # PIS Cumulativo
        (1.65,  _ALIQ_TOL, 17),   # PIS Não Cumulativo
    ],
    # COFINS
    'COFINS': [
        (3.00,  _ALIQ_TOL,  5),   # COFINS Cumulativo
        (7.60,  _ALIQ_TOL, 19),   # COFINS Não Cumulativo
    ],
}


def _aliq_float(valor: str) -> float:
    """Converte string de alíquota para float, tratando vírgula e ponto."""
    v = valor.strip().replace(',', '.')
    if not v:
        return 0.0
    try:
        f = float(v)
        # Se vier como decimal (ex: 0.0065) converte para percentual (0.65)
        # Heurística: se o valor for < 0.20 e > 0, assume que é fração → multiplica por 100
        if 0 < f < 0.20:
            f = round(f * 100, 6)
        return f
    except ValueError:
        return 0.0


def get_codigo_pis(aliq_str: str, por_nome: dict) -> int:
    """
    Retorna o código Domínio do PIS com base na alíquota:
      0,65% → PIS Cumulativo      (código 4)
      1,65% → PIS Não Cumulativo  (código 17)
    Se não reconhecer a alíquota, usa o código genérico de PIS.
    """
    aliq = _aliq_float(aliq_str)
    for (central, tol, cod) in ALIQ_PIS_COFINS['PIS']:
        if abs(aliq - central) <= tol:
            return cod
    # Fallback: código genérico PIS
    return get_codigo_imposto('PIS', por_nome, 4)


def get_codigo_cofins(aliq_str: str, por_nome: dict) -> int:
    """
    Retorna o código Domínio do COFINS com base na alíquota:
      3,00% → COFINS Cumulativo      (código 5)
      7,60% → COFINS Não Cumulativo  (código 19)
    Se não reconhecer a alíquota, usa o código genérico de COFINS.
    """
    aliq = _aliq_float(aliq_str)
    for (central, tol, cod) in ALIQ_PIS_COFINS['COFINS']:
        if abs(aliq - central) <= tol:
            return cod
    # Fallback: código genérico COFINS
    return get_codigo_imposto('COFINS', por_nome, 5)


# ==============================
# TEMA TR
# ==============================
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif; color: #444444;
        }
        h1, h2, h3 { color: #FF8000; font-weight: 700; }
        section[data-testid="stSidebar"] { background-color: #444444; color: #FFFFFF; }
        section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
        .stButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stButton > button:hover { background-color: #D64001; color: #FFFFFF; }
        .stDownloadButton > button {
            background-color: #FF8000; color: #FFFFFF;
            border: none; border-radius: 4px; font-weight: bold;
        }
        .stDownloadButton > button:hover { background-color: #D64001; color: #FFFFFF; }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 16px 20px; margin: 12px 0;
            color: #444444; font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 { color: #FF8000; margin-top: 14px; margin-bottom: 6px; }
        .instrucoes-box h4:first-child { margin-top: 0; }
        </style>
    """, unsafe_allow_html=True)


# ==============================
# DECODE / ENCODE
# ==============================
def decode_arquivo(raw: bytes) -> str:
    for enc in ('utf-8', 'latin-1', 'cp1252'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def encode_ansi_seguro(conteudo: str, log: list) -> bytes:
    resultado = []
    substituicoes = 0
    for char in conteudo:
        try:
            resultado.append(char.encode('latin-1'))
        except UnicodeEncodeError:
            resultado.append(b'?')
            substituicoes += 1
    if substituicoes:
        log.append(f"AVISO: {substituicoes} caractere(s) fora do ANSI substituídos por '?'.")
    return b''.join(resultado)


# ==============================
# CAMINHOS CANDIDATOS
# ==============================
def _candidatos(nome: str) -> list:
    try:
        base = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        base = os.getcwd()
    return [
        os.path.join(base, nome),
        os.path.join(os.getcwd(), nome),
        nome,
    ]


# ==============================
# TABELA IMPOSTOS — Domínio Sistemas
# ==============================
@st.cache_data(show_spinner=False)
def carregar_tabela_impostos() -> tuple:
    caminho = next((c for c in _candidatos(NOME_IMP_XLSX) if os.path.isfile(c)), None)
    por_codigo: dict = {}
    por_nome:   dict = {}

    if caminho is None:
        for nome, cod in IMPOSTOS_FALLBACK.items():
            por_codigo[cod]        = nome
            por_nome[nome.upper()] = cod
        return por_codigo, por_nome

    try:
        df = pd.read_excel(caminho, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]

        col_cod  = next((c for c in df.columns if 'CÓD' in c or 'COD' in c or c == 'CÓDIGO'), None)
        col_nome = next((c for c in df.columns if 'NOME' in c), None)

        if col_cod is None or col_nome is None:
            if len(df.columns) >= 2:
                col_cod  = df.columns[0]
                col_nome = df.columns[1]
            else:
                for nome, cod in IMPOSTOS_FALLBACK.items():
                    por_codigo[cod] = nome
                    por_nome[nome.upper()] = cod
                return por_codigo, por_nome

        for _, row in df.iterrows():
            raw_cod = str(row[col_cod]).strip().split('.')[0]
            nome    = str(row[col_nome]).strip()
            if not raw_cod or raw_cod.upper() == 'NAN' or not nome or nome.upper() == 'NAN':
                continue
            try:
                cod = int(raw_cod)
            except ValueError:
                continue
            por_codigo[cod]        = nome
            por_nome[nome.upper()] = cod

    except Exception:
        for nome, cod in IMPOSTOS_FALLBACK.items():
            por_codigo[cod] = nome
            por_nome[nome.upper()] = cod

    return por_codigo, por_nome


def get_codigo_imposto(nome_fragmento: str, por_nome: dict, default: int = 0) -> int:
    chave = nome_fragmento.strip().upper()
    if chave in por_nome:
        return por_nome[chave]
    for k, v in por_nome.items():
        if chave in k or k in chave:
            return v
    return default


# ==============================
# TABELA CFOP — Receita Federal
# ==============================
@st.cache_data(show_spinner=False)
def carregar_tabela_cfop_oficial() -> tuple:
    caminho = next((c for c in _candidatos(NOME_CFOP_XLSX) if os.path.isfile(c)), None)
    tabela_descr: dict = {}
    tabela_flags: dict = {}

    if caminho is None:
        return tabela_descr, tabela_flags

    try:
        df = pd.read_excel(caminho, sheet_name="CFOP", dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]

        col_cfop  = next((c for c in df.columns if c == 'CFOP'), None)
        col_descr = next((c for c in df.columns if 'DESCRI' in c or 'RESUMIDA' in c), None)
        col_nfe   = next((c for c in df.columns if 'INDNFE'      in c.replace(' ', '').upper()), None)
        col_com   = next((c for c in df.columns if 'INDCOMUNICA' in c.replace(' ', '').upper()), None)
        col_trp   = next((c for c in df.columns if 'INDTRANSP'   in c.replace(' ', '').upper()), None)
        col_dev   = next((c for c in df.columns if 'INDDEVOL'    in c.replace(' ', '').upper()), None)

        if col_cfop is None or col_descr is None:
            return tabela_descr, tabela_flags

        def _flag(row, col):
            if col is None: return 0
            v = str(row[col]).strip().split('.')[0]
            try: return int(v)
            except ValueError: return 0

        for _, row in df.iterrows():
            raw = str(row[col_cfop]).strip()
            if '.' in raw: raw = raw.split('.')[0]
            raw = ''.join(filter(str.isdigit, raw))
            if not raw: continue
            cfop  = raw.zfill(4)
            descr = str(row[col_descr]).strip()
            if cfop == '0000' or not descr or descr.lower() == 'nan': continue
            tabela_descr[cfop] = descr
            tabela_flags[cfop] = {
                'indNFe':      _flag(row, col_nfe),
                'indComunica': _flag(row, col_com),
                'indTransp':   _flag(row, col_trp),
                'indDevol':    _flag(row, col_dev),
            }
    except Exception:
        return {}, {}

    return tabela_descr, tabela_flags


def get_descricao_cfop(cfop: str, tabela_descr: dict) -> str:
    return tabela_descr.get(str(cfop).strip().zfill(4), '— descrição não encontrada —')

def get_flags_cfop(cfop: str, tabela_flags: dict) -> dict:
    return tabela_flags.get(
        str(cfop).strip().zfill(4),
        {'indNFe': 0, 'indComunica': 0, 'indTransp': 0, 'indDevol': 0}
    )

def get_tipo_operacao(cfop: str) -> str:
    p = str(cfop).strip()[:1]
    if p in ('1', '2', '3'): return 'Entrada'
    if p in ('5', '6', '7'): return 'Saída'
    return 'Desconhecido'

def is_devolucao(cfop: str, tabela_flags: dict) -> bool:
    return get_flags_cfop(cfop, tabela_flags).get('indDevol', 0) == 1


# ==============================
# PARSER SPED FISCAL
# ==============================
def parse_sped(content: str) -> dict:
    linhas_ordenadas = []
    por_tipo = {}
    for num_linha, linha in enumerate(content.splitlines(), start=1):
        linha = linha.strip()
        if not linha: continue
        campos = linha.split('|')
        if campos and campos[0] == '':  campos = campos[1:]
        if campos and campos[-1] == '': campos = campos[:-1]
        if not campos: continue
        tipo = campos[0].strip()
        if not tipo: continue
        linhas_ordenadas.append((tipo, campos, num_linha))
        por_tipo.setdefault(tipo, []).append((campos, num_linha))
    return {'linhas_ordenadas': linhas_ordenadas, 'por_tipo': por_tipo}

def _c(campos: list, idx: int, default: str = '') -> str:
    return campos[idx].strip() if len(campos) > idx else default


# ==============================
# ÍNDICES SPED FISCAL
# ==============================
SPED_0000_CNPJ       = 6
SPED_0150_COD        = 1
SPED_0150_CNPJ       = 4
SPED_C100_IND_OPER   = 1
SPED_C100_COD_PART   = 3
SPED_C100_COD_MOD    = 4
SPED_C100_COD_SIT    = 5
SPED_C100_SER        = 6
SPED_C100_NUM_DOC    = 7
SPED_C100_CHV_NFE    = 8
SPED_C100_DT_DOC     = 9
SPED_C100_DT_ES      = 10
SPED_C100_VL_DOC     = 11
SPED_C100_VL_ICMS    = 21
SPED_C100_VL_IPI     = 24
SPED_C100_VL_PIS     = 25
SPED_C100_VL_COFINS  = 26
SPED_C170_NUM_ITEM   = 1
SPED_C170_COD_ITEM   = 2
SPED_C170_DESCR      = 3
SPED_C170_QTD        = 4
SPED_C170_UNID       = 5
SPED_C170_VL_ITEM    = 6
SPED_C170_VL_DESC    = 7
SPED_C170_IND_MOV    = 8
SPED_C170_CST_ICMS   = 9
SPED_C170_CFOP       = 10
SPED_C170_COD_NAT    = 11
SPED_C170_VL_BC      = 12
SPED_C170_ALIQ_ICMS  = 13
SPED_C170_VL_ICMS    = 14
SPED_C170_VL_BC_ST   = 15
SPED_C170_ALIQ_ST    = 16
SPED_C170_VL_ICMS_ST = 17
SPED_C170_IND_APUR   = 18
SPED_C170_CST_IPI    = 19
SPED_C170_COD_ENQ    = 20
SPED_C170_VL_BC_IPI  = 21
SPED_C170_ALIQ_IPI   = 22
SPED_C170_VL_IPI     = 23
SPED_C170_CST_PIS    = 24
SPED_C170_VL_BC_PIS  = 25
SPED_C170_ALIQ_PIS   = 26
SPED_C170_VL_PIS     = 28
SPED_C170_CST_COFINS = 29
SPED_C170_VL_BC_COF  = 30
SPED_C170_ALIQ_COF   = 31
SPED_C170_VL_COFINS  = 33
SPED_C190_CFOP       = 2
SPED_C190_ALIQ       = 3
SPED_C190_VL_OPR     = 4
SPED_C190_VL_BC      = 5
SPED_C190_VL_ICMS    = 6
SPED_D100_IND_OPER   = 1
SPED_D100_COD_PART   = 3
SPED_D100_COD_MOD    = 4
SPED_D100_COD_SIT    = 5
SPED_D100_SER        = 6
SPED_D100_NUM_DOC    = 8
SPED_D100_DT_DOC     = 10
SPED_D100_VL_DOC     = 14
SPED_D100_ALIQ       = 19
SPED_D100_VL_ICMS    = 20
SPED_H010_COD_ITEM   = 1
SPED_H010_UNID       = 2
SPED_H010_QTD        = 3
SPED_H010_VL_UNIT    = 4
SPED_H010_VL_ITEM    = 5


# ==============================
# EXTRAÇÃO DE CFOPs DO SPED
# ==============================
def extrair_cfops_do_sped(parsed: dict, tabela_flags: dict, log: list) -> dict:
    cfops = {}

    def registrar(cfop_raw: str, tipo_reg: str):
        cfop = str(cfop_raw).strip()
        if not cfop or not cfop.isdigit(): return
        cfop = cfop.zfill(4)
        if cfop == '0000': return
        if cfop not in cfops:
            flags = get_flags_cfop(cfop, tabela_flags)
            cfops[cfop] = {
                'registros':     set(),
                'ocorrencias':   0,
                'tipo_operacao': get_tipo_operacao(cfop),
                'indNFe':        flags['indNFe'],
                'indComunica':   flags['indComunica'],
                'indTransp':     flags['indTransp'],
                'indDevol':      flags['indDevol'],
            }
        cfops[cfop]['registros'].add(tipo_reg)
        cfops[cfop]['ocorrencias'] += 1

    contadores = {'C100': 0, 'C170': 0, 'C190': 0, 'D100': 0}
    for tipo, campos, _ in parsed['linhas_ordenadas']:
        if tipo in contadores: contadores[tipo] += 1
        if tipo == 'C170':   registrar(_c(campos, SPED_C170_CFOP), 'C170')
        elif tipo == 'C190': registrar(_c(campos, SPED_C190_CFOP), 'C190')

    log.append(
        f"Registros lidos: C100={contadores['C100']} | C170={contadores['C170']} | "
        f"C190={contadores['C190']} | D100={contadores['D100']}"
    )
    log.append(f"CFOPs únicos encontrados: {len(cfops)} — {sorted(cfops.keys())}")
    return cfops


# ==============================
# GERADOR XLSX — TEMA TR
# ==============================
def gerar_xlsx_acumuladores_tr(cfops_dict: dict, tabela_descr: dict, tabela_flags: dict) -> bytes:
    wb = Workbook()
    COR_LARANJA   = "FF8000"
    COR_CINZA_ESC = "444444"
    COR_CINZA_CLR = "E9E9E9"
    COR_BRANCO    = "FFFFFF"
    COR_LARANJA_C = "FFF3E0"
    COR_VERDE_CLR = "E8F5E9"
    COR_VERM_CLR  = "FFEBEE"

    borda_fina = Border(
        left=Side(style='thin', color="CCCCCC"), right=Side(style='thin', color="CCCCCC"),
        top=Side(style='thin',  color="CCCCCC"), bottom=Side(style='thin', color="CCCCCC"),
    )
    def fill(h):   return PatternFill("solid", fgColor=h)
    def center():  return Alignment(horizontal='center', vertical='center')
    def left_al(): return Alignment(horizontal='left',   vertical='center')
    def wrap_al(): return Alignment(horizontal='left',   vertical='center', wrap_text=True)

    cfops_ord = sorted(
        cfops_dict.items(),
        key=lambda x: (0 if x[1]['tipo_operacao'] == 'Entrada' else 1, x[0])
    )

    # ── Aba 1: Acumuladores ───────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Acumuladores"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:I1"); ws1.row_dimensions[1].height = 36
    c = ws1["A1"]
    c.value = "Thomson Reuters  |  Domínio Sistemas  —  Tabela de Acumuladores CFOP"
    c.fill = fill(COR_CINZA_ESC); c.font = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA); c.alignment = left_al()

    ws1.merge_cells("A2:I2"); ws1.row_dimensions[2].height = 20
    c2 = ws1["A2"]
    c2.value = (f"Extraído do SPED Fiscal  |  {len(cfops_dict)} CFOP(s)  |  "
                f"Descrições: Receita Federal  |  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    c2.fill = fill(COR_LARANJA); c2.font = Font(name='Segoe UI', size=9, color=COR_BRANCO); c2.alignment = left_al()

    ws1.merge_cells("A3:I3"); ws1.row_dimensions[3].height = 18
    c3 = ws1["A3"]
    c3.value = "⚠  Preencha a coluna ACUMULADOR para cada CFOP antes de fazer o upload no conversor."
    c3.fill = fill(COR_CINZA_CLR); c3.font = Font(name='Segoe UI', bold=True, size=9, color=COR_CINZA_ESC); c3.alignment = left_al()

    ws1.row_dimensions[4].height = 6
    ws1.row_dimensions[5].height = 22

    cabecalhos = ['CFOP', 'DESCRIÇÃO OFICIAL (Receita Federal)', 'TIPO OPERAÇÃO',
                  'OCORRÊNCIAS', 'DEVOLUÇÃO', 'NFe', 'COMUNICAÇÃO', 'TRANSPORTE', 'ACUMULADOR']
    col_widths  = [10, 60, 16, 14, 12, 8, 14, 14, 16]

    for ci, (cab, w) in enumerate(zip(cabecalhos, col_widths), start=1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
        cell = ws1.cell(row=5, column=ci, value=cab)
        cell.fill = fill(COR_LARANJA); cell.font = Font(name='Segoe UI', bold=True, size=11, color=COR_BRANCO)
        cell.alignment = center(); cell.border = borda_fina

    linha = 6
    for idx, (cfop, info) in enumerate(cfops_ord):
        ws1.row_dimensions[linha].height = 30
        bg = (COR_CINZA_CLR if idx % 2 == 0 else COR_BRANCO) if info['tipo_operacao'] == 'Entrada' \
             else (COR_LARANJA_C if idx % 2 == 0 else COR_BRANCO)
        descricao = get_descricao_cfop(cfop, tabela_descr)
        eh_devol = '✔' if info.get('indDevol')    == 1 else ''
        eh_nfe   = '✔' if info.get('indNFe')       == 1 else ''
        eh_com   = '✔' if info.get('indComunica')  == 1 else ''
        eh_trp   = '✔' if info.get('indTransp')    == 1 else ''
        valores  = [cfop, descricao, info['tipo_operacao'], info['ocorrencias'],
                    eh_devol, eh_nfe, eh_com, eh_trp, '']

        for ci, valor in enumerate(valores, start=1):
            cell = ws1.cell(row=linha, column=ci, value=valor)
            cell.border = borda_fina
            if ci == 1:
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC); cell.alignment = center()
            elif ci == 2:
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', size=9, color=COR_CINZA_ESC); cell.alignment = wrap_al()
            elif ci == 3:
                cell.fill = fill(bg)
                cor = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor); cell.alignment = center()
            elif ci == 4:
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC); cell.alignment = center()
            elif ci in (5, 6, 7, 8):
                cor_flag = "B71C1C" if (ci == 5 and valor == '✔') else ("1B5E20" if valor == '✔' else COR_CINZA_ESC)
                cell.fill = fill(COR_VERM_CLR if ci == 5 and valor == '✔' else bg)
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor_flag); cell.alignment = center()
            elif ci == 9:
                cell.fill = fill("FFF8F0"); cell.font = Font(name='Segoe UI', bold=True, size=10, color=COR_LARANJA)
                cell.alignment = center()
                cell.border = Border(left=Side(style='medium', color=COR_LARANJA), right=Side(style='medium', color=COR_LARANJA),
                                     top=Side(style='thin', color="CCCCCC"), bottom=Side(style='thin', color="CCCCCC"))
        linha += 1

    ws1.merge_cells(f"A{linha}:I{linha}"); ws1.row_dimensions[linha].height = 18
    cr = ws1.cell(row=linha, column=1, value="Thomson Reuters  |  Domínio Sistemas  |  Descrições: Receita Federal")
    cr.fill = fill(COR_CINZA_ESC); cr.font = Font(name='Segoe UI', size=8, color="888888")
    cr.alignment = Alignment(horizontal='right', vertical='center')
    ws1.freeze_panes = "A6"; ws1.auto_filter.ref = f"A5:I{linha - 1}"

    # ── Aba 2: CFOPs Encontrados ──────────────────────────────────────────
    ws2 = wb.create_sheet(title="CFOPs Encontrados")
    ws2.sheet_view.showGridLines = False
    n_ent = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Entrada')
    n_sai = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Saída')
    n_dev = sum(1 for v in cfops_dict.values() if v.get('indDevol') == 1)

    ws2.merge_cells("A1:H1"); ws2.row_dimensions[1].height = 36
    c = ws2["A1"]
    c.value = "Thomson Reuters  |  Domínio Sistemas  —  CFOPs Identificados no SPED Fiscal"
    c.fill = fill(COR_CINZA_ESC); c.font = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA); c.alignment = left_al()

    ws2.merge_cells("A2:H2"); ws2.row_dimensions[2].height = 20
    c2 = ws2["A2"]
    c2.value = (f"Analisado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
                f"Total: {len(cfops_dict)}  |  Entradas: {n_ent}  |  Saídas: {n_sai}  |  "
                f"Devoluções: {n_dev}  |  Fonte: Receita Federal")
    c2.fill = fill(COR_LARANJA); c2.font = Font(name='Segoe UI', size=9, color=COR_BRANCO); c2.alignment = left_al()

    ws2.row_dimensions[3].height = 6; ws2.row_dimensions[4].height = 22
    cab2  = ['CFOP', 'DESCRIÇÃO OFICIAL (Receita Federal)', 'TIPO OPERAÇÃO',
             'DEVOLUÇÃO', 'REGISTROS SPED', 'OCORRÊNCIAS', 'STATUS RF', 'STATUS']
    wids2 = [10, 60, 16, 12, 18, 14, 12, 22]

    for ci, (cab, w) in enumerate(zip(cab2, wids2), start=1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
        cell = ws2.cell(row=4, column=ci, value=cab)
        cell.fill = fill(COR_LARANJA); cell.font = Font(name='Segoe UI', bold=True, size=11, color=COR_BRANCO)
        cell.alignment = center(); cell.border = borda_fina

    linha2 = 5
    for idx, (cfop, info) in enumerate(cfops_ord):
        ws2.row_dimensions[linha2].height = 30
        bg      = COR_CINZA_CLR if idx % 2 == 0 else COR_BRANCO
        descr   = get_descricao_cfop(cfop, tabela_descr)
        mapeado = descr != '— descrição não encontrada —'
        status  = '✔ Catalogado' if mapeado else '✘ Não catalogado'
        bg_st   = COR_VERDE_CLR if mapeado else COR_VERM_CLR
        regs    = ', '.join(sorted(info['registros']))
        eh_dev  = '✔ Devolução' if info.get('indDevol') == 1 else '—'
        status_rf = 'NFe' if info.get('indNFe') == 1 else (
                    'CT-e' if info.get('indTransp') == 1 else (
                    'Comunic.' if info.get('indComunica') == 1 else '—'))
        vals = [cfop, descr, info['tipo_operacao'], eh_dev, regs, info['ocorrencias'], status_rf, status]

        for ci, valor in enumerate(vals, start=1):
            cell = ws2.cell(row=linha2, column=ci, value=valor)
            cell.border = borda_fina
            if ci == 1:
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC); cell.alignment = center()
            elif ci == 2:
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', size=9, color=COR_CINZA_ESC); cell.alignment = wrap_al()
            elif ci == 3:
                cell.fill = fill(bg)
                cor = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor); cell.alignment = center()
            elif ci == 4:
                cor_d = "B71C1C" if info.get('indDevol') == 1 else COR_CINZA_ESC
                bg_d  = COR_VERM_CLR if info.get('indDevol') == 1 else bg
                cell.fill = fill(bg_d); cell.font = Font(name='Segoe UI', bold=True, size=9, color=cor_d); cell.alignment = center()
            elif ci in (5, 6, 7):
                cell.fill = fill(bg); cell.font = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC); cell.alignment = center()
            elif ci == 8:
                cell.fill = fill(bg_st)
                cor = "1B5E20" if mapeado else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor); cell.alignment = center()
        linha2 += 1

    ws2.merge_cells(f"A{linha2}:H{linha2}"); ws2.row_dimensions[linha2].height = 18
    cr2 = ws2.cell(row=linha2, column=1, value="Thomson Reuters  |  Domínio Sistemas  |  Fonte: Receita Federal")
    cr2.fill = fill(COR_CINZA_ESC); cr2.font = Font(name='Segoe UI', size=8, color="888888")
    cr2.alignment = Alignment(horizontal='right', vertical='center')
    ws2.freeze_panes = "A5"; ws2.auto_filter.ref = f"A4:H{linha2 - 1}"

    # ── Aba 3: Tabela Completa CFOP ───────────────────────────────────────
    ws3 = wb.create_sheet(title="Tabela CFOP Receita Federal")
    ws3.sheet_view.showGridLines = False
    ws3.merge_cells("A1:G1"); ws3.row_dimensions[1].height = 36
    c = ws3["A1"]
    c.value = "Thomson Reuters  |  Tabela Completa de CFOPs — Receita Federal"
    c.fill = fill(COR_CINZA_ESC); c.font = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA); c.alignment = left_al()
    ws3.merge_cells("A2:G2"); ws3.row_dimensions[2].height = 20
    c2 = ws3["A2"]
    c2.value = f"Total: {len(tabela_descr)} CFOPs  |  Fonte: {NOME_CFOP_XLSX} — Receita Federal"
    c2.fill = fill(COR_LARANJA); c2.font = Font(name='Segoe UI', size=9, color=COR_BRANCO); c2.alignment = left_al()
    ws3.row_dimensions[3].height = 6; ws3.row_dimensions[4].height = 22
    cab3  = ['CFOP', 'TIPO OPERAÇÃO', 'DESCRIÇÃO OFICIAL', 'DEVOLUÇÃO', 'NFe', 'COMUNICAÇÃO', 'TRANSPORTE']
    wids3 = [10, 16, 70, 12, 8, 14, 14]
    for ci, (cab, w) in enumerate(zip(cab3, wids3), start=1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
        cell = ws3.cell(row=4, column=ci, value=cab)
        cell.fill = fill(COR_LARANJA); cell.font = Font(name='Segoe UI', bold=True, size=11, color=COR_BRANCO)
        cell.alignment = center(); cell.border = borda_fina
    linha3 = 5
    for idx, (cfop, descr) in enumerate(sorted(tabela_descr.items())):
        ws3.row_dimensions[linha3].height = 28
        bg = COR_CINZA_CLR if idx % 2 == 0 else COR_BRANCO
        tipo  = get_tipo_operacao(cfop)
        flags = tabela_flags.get(cfop, {})
        f_dev = '✔' if flags.get('indDevol')    == 1 else ''
        f_nfe = '✔' if flags.get('indNFe')       == 1 else ''
        f_com = '✔' if flags.get('indComunica')  == 1 else ''
        f_trp = '✔' if flags.get('indTransp')    == 1 else ''
        for ci, valor in enumerate([cfop, tipo, descr, f_dev, f_nfe, f_com, f_trp], start=1):
            cell = ws3.cell(row=linha3, column=ci, value=valor)
            cell.fill = fill(bg); cell.border = borda_fina
            if ci == 1:
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC); cell.alignment = center()
            elif ci == 2:
                cor = "1B5E20" if tipo == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor); cell.alignment = center()
            elif ci == 3:
                cell.font = Font(name='Segoe UI', size=9, color=COR_CINZA_ESC); cell.alignment = wrap_al()
            else:
                cor_f = "B71C1C" if (ci == 4 and valor == '✔') else ("1B5E20" if valor == '✔' else COR_CINZA_ESC)
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor_f); cell.alignment = center()
        linha3 += 1
    ws3.freeze_panes = "A5"; ws3.auto_filter.ref = f"A4:G{linha3 - 1}"

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.read()


# ==============================
# CARREGAMENTO DA TABELA DE ACUMULADORES
# ==============================
def carregar_acumuladores(arquivo_bytes: bytes, nome_arquivo: str, log: list) -> dict:
    try:
        ext = os.path.splitext(nome_arquivo)[1].lower()
        if ext in ('.xlsx', '.xls'):
            df = None
            for header_row in range(6):
                try:
                    df_tent = pd.read_excel(
                        io.BytesIO(arquivo_bytes), sheet_name=0,
                        header=header_row, dtype=str
                    )
                    cols = [str(c).strip().upper() for c in df_tent.columns]
                    if 'CFOP' in cols and 'ACUMULADOR' in cols:
                        df_tent.columns = cols
                        df = df_tent
                        log.append(f"Planilha lida com sucesso (header linha Excel {header_row + 1}). "
                                   f"Colunas: {list(df.columns)}")
                        break
                except Exception:
                    continue
            if df is None:
                log.append("ERRO: Não foi possível localizar as colunas 'CFOP' e 'ACUMULADOR' na planilha.")
                return None
        else:
            raw_str = arquivo_bytes.decode('latin-1', errors='replace')
            sep     = ';' if raw_str.count(';') >= raw_str.count(',') else ','
            df      = pd.read_csv(io.StringIO(raw_str), sep=sep, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            if 'CFOP' not in df.columns or 'ACUMULADOR' not in df.columns:
                log.append("ERRO: O arquivo CSV deve conter as colunas 'CFOP' e 'ACUMULADOR'.")
                return None

        tabela = {}
        for _, row in df.iterrows():
            raw_cfop = str(row['CFOP']).strip()
            if '.' in raw_cfop: raw_cfop = raw_cfop.split('.')[0]
            raw_cfop = ''.join(filter(str.isdigit, raw_cfop))
            if not raw_cfop: continue
            cfop = raw_cfop.zfill(4)

            raw_acum = str(row['ACUMULADOR']).strip()
            if raw_acum.endswith('.0'): raw_acum = raw_acum[:-2]
            acum = raw_acum.strip()

            if cfop == '0000': continue
            if not acum or acum.upper() in ('NAN', '', '0'): continue
            tabela[cfop] = acum

        if not tabela:
            log.append("ERRO: Nenhum par CFOP → Acumulador preenchido.")
            return None

        log.append(f"Tabela de acumuladores carregada: {len(tabela)} CFOPs mapeados.")
        log.append(f"Amostra: { {k: tabela[k] for k in list(tabela)[:5]} }")
        return tabela

    except Exception as e:
        log.append(f"ERRO ao carregar tabela de acumuladores: {e}")
        log.append(traceback.format_exc())
        return None


def get_acumulador(cfop: str, tabela: dict, nao_mapeados: set) -> str:
    cfop_norm = str(cfop).strip().zfill(4)
    acum = tabela.get(cfop_norm)
    if acum is None:
        nao_mapeados.add(cfop_norm)
        return '9999'
    return acum


# ==============================
# CONVERSÃO SPED FISCAL → DOMÍNIO SISTEMAS
# V2.6 — PIS/COFINS com código por alíquota (cumulativo x não cumulativo)
# ==============================
def converter_sped_para_dominio(
    parsed: dict,
    tabela_acum: dict,
    tabela_flags: dict,
    por_nome_imp: dict,
    log: list,
) -> tuple:
    saida        = StringIO()
    nao_mapeados = set()
    stats = {
        'nf_entrada': 0, 'nf_saida': 0, 'itens': 0,
        'analiticos': 0, 'transporte': 0, 'inventario': 0,
        'devolucoes': 0, 'erros': 0,
    }

    # ── Códigos de impostos fixos (ICMS, IPI, ST) ─────────────────────────
    COD_ICMS    = get_codigo_imposto('ICMS',             por_nome_imp, 1)
    COD_IPI     = get_codigo_imposto('IPI',              por_nome_imp, 2)
    COD_ISS     = get_codigo_imposto('ISS',              por_nome_imp, 3)
    COD_ST      = get_codigo_imposto('SUBST. TRIBUTARIA',por_nome_imp, 9)
    COD_ICMS_ST = get_codigo_imposto('ICMS RETIDO',      por_nome_imp, 11)

    # ── PIS/COFINS: código resolvido por alíquota em tempo de execução ─────
    # get_codigo_pis(aliq_str, por_nome_imp)   → 4 (cumulativo) ou 17 (não cumulativo)
    # get_codigo_cofins(aliq_str, por_nome_imp) → 5 (cumulativo) ou 19 (não cumulativo)

    log.append(
        f"Códigos fixos: ICMS={COD_ICMS} | IPI={COD_IPI} | ISS={COD_ISS} | "
        f"ST={COD_ST} | ICMS_ST={COD_ICMS_ST}"
    )
    log.append(
        "PIS/COFINS: código resolvido por alíquota — "
        "0,65%→PIS Cum.(4) | 1,65%→PIS N.Cum.(17) | "
        "3,00%→COFINS Cum.(5) | 7,60%→COFINS N.Cum.(19)"
    )

    # ── 0000 ──────────────────────────────────────────────────────────────
    if '0000' in parsed['por_tipo']:
        campos, _ = parsed['por_tipo']['0000'][0]
        cnpj = _c(campos, SPED_0000_CNPJ)
        saida.write(f"|0000|{cnpj}|\n")
        log.append(f"0000: CNPJ={cnpj}")
    else:
        log.append("AVISO: Registro 0000 não encontrado.")

    # ── Participantes ─────────────────────────────────────────────────────
    participantes = {}
    if '0150' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['0150']:
            participantes[_c(campos, SPED_0150_COD)] = campos
        log.append(f"Participantes carregados: {len(participantes)}")

    # ── Hierarquia C100 → C170 → C190 ────────────────────────────────────
    blocos_c    = []
    bloco_atual = None
    for tipo, campos, num_linha in parsed['linhas_ordenadas']:
        if tipo == 'C100':
            if bloco_atual is not None: blocos_c.append(bloco_atual)
            bloco_atual = {'c100': campos, 'c170': [], 'c190': []}
        elif tipo == 'C170':
            if bloco_atual is not None: bloco_atual['c170'].append(campos)
        elif tipo == 'C190':
            if bloco_atual is not None: bloco_atual['c190'].append(campos)
    if bloco_atual is not None: blocos_c.append(bloco_atual)
    log.append(f"Blocos C100 montados: {len(blocos_c)}")

    # ── C100 → 1000 + 1020 + 1030(s) ─────────────────────────────────────
    for bloco in blocos_c:
        campos_c100 = bloco['c100']
        try:
            ind_oper = _c(campos_c100, SPED_C100_IND_OPER)
            cod_part = _c(campos_c100, SPED_C100_COD_PART)
            cod_mod  = _c(campos_c100, SPED_C100_COD_MOD)
            cod_sit  = _c(campos_c100, SPED_C100_COD_SIT)
            serie    = _c(campos_c100, SPED_C100_SER)
            num_doc  = _c(campos_c100, SPED_C100_NUM_DOC)
            chv_nfe  = _c(campos_c100, SPED_C100_CHV_NFE)
            dt_doc   = _c(campos_c100, SPED_C100_DT_DOC)
            dt_es    = _c(campos_c100, SPED_C100_DT_ES)
            vl_doc   = _c(campos_c100, SPED_C100_VL_DOC)
            vl_icms  = _c(campos_c100, SPED_C100_VL_ICMS)

            part_campos = participantes.get(cod_part, [])
            cnpj_part   = _c(part_campos, SPED_0150_CNPJ) if part_campos else ''
            tipo_es     = 'E' if ind_oper == '0' else 'S'

            cfop_principal = ''
            if bloco['c170']:   cfop_principal = _c(bloco['c170'][0], SPED_C170_CFOP)
            elif bloco['c190']: cfop_principal = _c(bloco['c190'][0], SPED_C190_CFOP)

            acum_principal = get_acumulador(cfop_principal, tabela_acum, nao_mapeados)
            eh_devol       = is_devolucao(cfop_principal, tabela_flags)
            if eh_devol: stats['devolucoes'] += 1

            aliq_icms_nf = '0,00'
            if bloco['c190']: aliq_icms_nf = _c(bloco['c190'][0], SPED_C190_ALIQ) or '0,00'

            obs = 'DEVOLUCAO' if eh_devol else 'OBSERVACAO'
            saida.write(
                f"|1000|{num_doc}|{cnpj_part}||{ind_oper}|{cfop_principal}|"
                f"{serie}|{cod_mod}|{cod_sit}|{chv_nfe}|||"
                f"{dt_doc}|{dt_es}|{vl_doc}||{obs}|C||||||||{tipo_es}|"
                f"0,00|0,00|0,00|0,00||0,00||||0,00|0,00|0,00||{vl_doc}|"
                f"0|0||||{acum_principal}||0,00||||||N|S||{tipo_es}||0|||||"
                f"||||||||||||0|{cod_sit}|0||0,00|0,00|0,00|||||||||||\n"
            )
            if ind_oper == '0': stats['nf_entrada'] += 1
            else:               stats['nf_saida']   += 1

            saida.write(
                f"|1020|{num_doc}||{vl_doc}|{aliq_icms_nf}|{vl_icms}|"
                f"0,00|0,00|0,00|0,00|{vl_doc}||||\n"
            )
            stats['analiticos'] += 1

            # ── Registros 1030 — um por item C170 ────────────────────────
            for campos_c170 in bloco['c170']:
                num_item  = _c(campos_c170, SPED_C170_NUM_ITEM)
                qtd       = _c(campos_c170, SPED_C170_QTD)       or '0'
                unid      = _c(campos_c170, SPED_C170_UNID)
                vl_item   = _c(campos_c170, SPED_C170_VL_ITEM)   or '0,00'
                vl_desc_i = _c(campos_c170, SPED_C170_VL_DESC)   or '0,00'
                cfop_item = _c(campos_c170, SPED_C170_CFOP)

                # ── ICMS ──────────────────────────────────────────────────
                vl_bc_icms   = _c(campos_c170, SPED_C170_VL_BC)       or '0,00'
                aliq_icms    = _c(campos_c170, SPED_C170_ALIQ_ICMS)   or '0,00'
                vl_icms_i    = _c(campos_c170, SPED_C170_VL_ICMS)     or '0,00'

                # ── ICMS-ST ───────────────────────────────────────────────
                vl_bc_st     = _c(campos_c170, SPED_C170_VL_BC_ST)    or '0,00'
                aliq_st      = _c(campos_c170, SPED_C170_ALIQ_ST)     or '0,00'
                vl_icms_st   = _c(campos_c170, SPED_C170_VL_ICMS_ST)  or '0,00'

                # ── IPI ───────────────────────────────────────────────────
                vl_bc_ipi    = _c(campos_c170, SPED_C170_VL_BC_IPI)   or '0,00'
                aliq_ipi     = _c(campos_c170, SPED_C170_ALIQ_IPI)    or '0,00'
                vl_ipi       = _c(campos_c170, SPED_C170_VL_IPI)      or '0,00'

                # ── PIS — alíquota determina código ───────────────────────
                vl_bc_pis    = _c(campos_c170, SPED_C170_VL_BC_PIS)   or '0,00'
                aliq_pis     = _c(campos_c170, SPED_C170_ALIQ_PIS)    or '0,00'
                vl_pis       = _c(campos_c170, SPED_C170_VL_PIS)      or '0,00'
                # V2.6: resolve código PIS pela alíquota
                cod_pis      = get_codigo_pis(aliq_pis, por_nome_imp)

                # ── COFINS — alíquota determina código ────────────────────
                vl_bc_cof    = _c(campos_c170, SPED_C170_VL_BC_COF)   or '0,00'
                aliq_cof     = _c(campos_c170, SPED_C170_ALIQ_COF)    or '0,00'
                vl_cofins    = _c(campos_c170, SPED_C170_VL_COFINS)   or '0,00'
                # V2.6: resolve código COFINS pela alíquota
                cod_cofins   = get_codigo_cofins(aliq_cof, por_nome_imp)

                # ── Valor unitário ────────────────────────────────────────
                try:
                    vl_unit = f"{float(vl_item.replace(',', '.')) / float(qtd.replace(',', '.')):.3f}".replace('.', ',')
                except Exception:
                    vl_unit = vl_item

                # ── Imposto principal do item (ICMS-ST > ICMS > IPI) ──────
                def _val(s): 
                    try: return float(s.replace(',', '.'))
                    except: return 0.0

                if _val(vl_icms_st) > 0:
                    cod_imp_item = COD_ICMS_ST
                    vl_bc_princ  = vl_bc_st
                    aliq_princ   = aliq_st
                    vl_imp_princ = vl_icms_st
                elif _val(vl_icms_i) > 0:
                    cod_imp_item = COD_ICMS
                    vl_bc_princ  = vl_bc_icms
                    aliq_princ   = aliq_icms
                    vl_imp_princ = vl_icms_i
                elif _val(vl_ipi) > 0:
                    cod_imp_item = COD_IPI
                    vl_bc_princ  = vl_bc_ipi
                    aliq_princ   = aliq_ipi
                    vl_imp_princ = vl_ipi
                else:
                    cod_imp_item = COD_ICMS
                    vl_bc_princ  = vl_bc_icms
                    aliq_princ   = aliq_icms
                    vl_imp_princ = vl_icms_i

                # ── Registro 1030 ─────────────────────────────────────────
                saida.write(
                    f"|1030|{num_item}|{qtd}|{vl_unit}|0|0|1|{dt_doc}||"
                    f"{cod_sit}|{vl_item}|{vl_desc_i}|{vl_item}|0,00|"
                    f"{aliq_princ}|||"
                    f"{cod_pis}|{vl_bc_pis}|{aliq_pis}|{vl_pis}|0,000|{vl_imp_princ}|"
                    f"{cod_cofins}|{vl_bc_cof}|{aliq_cof}|{vl_cofins}||{vl_item}|0,00|"
                    f"{cod_imp_item}|{vl_bc_princ}|{vl_ipi}|{vl_ipi}|0,00|"
                    f"{cfop_item}||0,0000|0,00|0,00|0,00|{vl_bc_icms}|"
                    f"{COD_ICMS_ST}|{vl_bc_st}|{COD_ICMS_ST}|{vl_icms_st}|||||"
                    f"{dt_doc}|{dt_doc}||||||S|{unid}|||"
                    f"{vl_item}|||||||1|||||01|01||||||||\n"
                )
                stats['itens'] += 1

        except Exception as e:
            log.append(f"ERRO ao converter C100 NF={_c(campos_c100, SPED_C100_NUM_DOC)}: {e}")
            stats['erros'] += 1

    # ── D100 → 1000 + 1020 ────────────────────────────────────────────────
    if 'D100' in parsed['por_tipo']:
        for campos, num_linha in parsed['por_tipo']['D100']:
            try:
                ind_oper = _c(campos, SPED_D100_IND_OPER)
                cod_part = _c(campos, SPED_D100_COD_PART)
                cod_mod  = _c(campos, SPED_D100_COD_MOD)
                cod_sit  = _c(campos, SPED_D100_COD_SIT)
                serie    = _c(campos, SPED_D100_SER)
                num_doc  = _c(campos, SPED_D100_NUM_DOC)
                dt_doc   = _c(campos, SPED_D100_DT_DOC)
                vl_doc   = _c(campos, SPED_D100_VL_DOC)
                aliq     = _c(campos, SPED_D100_ALIQ)
                vl_icms  = _c(campos, SPED_D100_VL_ICMS)
                tipo_es  = 'E' if ind_oper == '0' else 'S'
                part_campos = participantes.get(cod_part, [])
                cnpj_part   = _c(part_campos, SPED_0150_CNPJ) if part_campos else ''

                saida.write(
                    f"|1000|{num_doc}|{cnpj_part}||{ind_oper}||"
                    f"{serie}|{cod_mod}|{cod_sit}||||"
                    f"{dt_doc}|{dt_doc}|{vl_doc}||FRETE|C||||||||{tipo_es}|"
                    f"0,00|0,00|0,00|0,00||0,00||||0,00|0,00|0,00||{vl_doc}|"
                    f"0|0|||||0,00||||||N|S||{tipo_es}||0|||||"
                    f"||||||||||||0|{cod_sit}|0||0,00|0,00|0,00|||||||||||\n"
                )
                saida.write(
                    f"|1020|{num_doc}||{vl_doc}|{aliq}|{vl_icms}|"
                    f"0,00|0,00|0,00|0,00|{vl_doc}||||\n"
                )
                stats['transporte'] += 1
            except Exception as e:
                log.append(f"ERRO ao converter D100 linha {num_linha}: {e}")
                stats['erros'] += 1

    # ── H010 ──────────────────────────────────────────────────────────────
    if 'H010' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['H010']:
            saida.write(
                f"|H010|{_c(campos, SPED_H010_COD_ITEM)}|"
                f"{_c(campos, SPED_H010_UNID)}|"
                f"{_c(campos, SPED_H010_QTD)}|"
                f"{_c(campos, SPED_H010_VL_UNIT)}|"
                f"{_c(campos, SPED_H010_VL_ITEM)}|\n"
            )
            stats['inventario'] += 1

    saida.write("|9999|\n")

    if nao_mapeados:
        log.append(
            f"AVISO: {len(nao_mapeados)} CFOP(s) sem acumulador (receberam 9999): "
            f"{', '.join(sorted(nao_mapeados))}"
        )
    log.append(
        f"Conversão concluída — "
        f"NFs entrada={stats['nf_entrada']} | NFs saída={stats['nf_saida']} | "
        f"Itens={stats['itens']} | Analíticos={stats['analiticos']} | "
        f"Devoluções={stats['devolucoes']} | "
        f"Transporte={stats['transporte']} | Inventário={stats['inventario']} | "
        f"Erros={stats['erros']}"
    )
    return saida.getvalue(), stats


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="SPED Fiscal → Domínio | Thomson Reuters",
        page_icon="🟠", layout="wide", initial_sidebar_state="expanded",
    )
    apply_tr_theme()

    tabela_cfop, tabela_flags   = carregar_tabela_cfop_oficial()
    por_codigo_imp, por_nome_imp = carregar_tabela_impostos()

    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                    border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                📄 Conversor SPED Fiscal → Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Entrada: <strong>SPED Fiscal EFD ICMS/IPI</strong> &nbsp;→&nbsp;
                Saída: <strong>Leiaute padrão Domínio Sistemas</strong>
                &nbsp;|&nbsp; CFOPs: <strong>{len(tabela_cfop)}</strong>
                &nbsp;|&nbsp; Impostos: <strong>{len(por_codigo_imp)}</strong>
            </p>
        </div>
        """, unsafe_allow_html=True,
    )

    # ── Sidebar ───────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters  |  Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 📥 Entrada (SPED Fiscal)")
        st.markdown("- **0000** Abertura\n- **0150** Participantes\n- **C100** Notas Fiscais\n"
                    "- **C170** Itens de NF\n- **C190** Analítico ICMS\n"
                    "- **D100** Conhecimento Transporte\n- **H010** Inventário\n")
        st.markdown("### 📤 Saída (Domínio Sistemas)")
        st.markdown("- **0000** Cabeçalho\n- **1000** Nota Fiscal\n"
                    "- **1020** Totais da NF\n- **1030** Itens da NF\n- **9999** Encerramento\n")
        st.markdown("---")

        if tabela_cfop:
            n_dev_total = sum(1 for f in tabela_flags.values() if f.get('indDevol') == 1)
            st.success(f"✅ {len(tabela_cfop)} CFOPs carregados\n↩ {n_dev_total} de devolução")
        else:
            st.error("❌ Tabela CFOP não carregada!")
            for c in _candidatos(NOME_CFOP_XLSX):
                st.caption(f"{'✔' if os.path.isfile(c) else '✘'} `{c}`")

        if por_codigo_imp:
            st.success(f"✅ {len(por_codigo_imp)} impostos carregados")
        else:
            st.warning("⚠ Tabela de impostos não carregada (usando fallback).")

        st.markdown("---")
        st.markdown("### 📑 Fluxo")
        st.markdown("1. Upload do SPED Fiscal `.txt`\n2. **Extrair CFOPs → baixar XLSX**\n"
                    "3. Preencher coluna `ACUMULADOR`\n4. Upload do XLSX preenchido\n"
                    "5. **Converter** e baixar saída\n")
        st.markdown("---")
        st.markdown("### 🧾 PIS/COFINS por Alíquota")
        st.markdown(
            "| Alíquota | Imposto | Cód. |\n"
            "|---|---|---|\n"
            "| 0,65% | PIS Cumulativo | 4 |\n"
            "| 1,65% | PIS Não Cum. | 17 |\n"
            "| 3,00% | COFINS Cumulativo | 5 |\n"
            "| 7,60% | COFINS Não Cum. | 19 |"
        )

    # ── Instruções ────────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown("""
            <div class="instrucoes-box">
            <h4>🔹 Etapa 1 — Upload do SPED e extração de CFOPs</h4>
            <p>Faça o upload do arquivo <code>.txt</code> do SPED Fiscal e clique em
            <b>🔍 Extrair CFOPs e Gerar Planilha</b>.</p>
            <h4>🔹 Etapa 2 — Preencher acumuladores</h4>
            <p>Abra o XLSX, preencha a coluna <b>ACUMULADOR</b> na aba
            <i>Acumuladores</i> e salve. <b>Não altere a estrutura do arquivo.</b></p>
            <h4>🔹 Etapa 3 — Converter</h4>
            <p>Faça o upload do XLSX preenchido e clique em
            <b>▶ Converter SPED → Domínio</b>.</p>
            <hr>
            <h4>⚠ Observações</h4>
            <ul>
                <li>CFOPs sem acumulador preenchido receberão <b>9999</b>.</li>
                <li>PIS/COFINS: código determinado automaticamente pela alíquota
                    (0,65%→4 | 1,65%→17 | 3,00%→5 | 7,60%→19).</li>
                <li>ICMS, IPI, ST: código lido de <code>Impostos.xlsx</code>.</li>
                <li>Saída em <b>ANSI (Latin-1)</b>.</li>
            </ul>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # ── Session state ─────────────────────────────────────────────────────
    defaults = {
        "log":             [f"Aplicação pronta. Versão: {VERSAO} | CFOPs: {len(tabela_cfop)} | Impostos: {len(por_codigo_imp)}"],
        "resultado":       None,
        "nome_saida":      "saida_dominio.txt",
        "stats":           None,
        "xlsx_bytes":      None,
        "xlsx_nome":       "acumuladores.xlsx",
        "cfops_extraidos": None,
        "tabela_acum_ok":  False,
        "arquivo_raw":     None,
        "arquivo_nome":    None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ════════════════════════════════════════════════════════════════════
    # ETAPA 1
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### 🔍 Etapa 1 — Upload do SPED Fiscal e extração de CFOPs")

    uploaded_file = st.file_uploader(
        "📂 Arquivo SPED Fiscal (.txt)", type=["txt"],
        help="Arquivo EFD ICMS/IPI exportado pelo sistema ERP", key="upload_sped",
    )

    if uploaded_file is not None:
        raw_atual = uploaded_file.read()
        if raw_atual != st.session_state.arquivo_raw:
            st.session_state.arquivo_raw     = raw_atual
            st.session_state.arquivo_nome    = uploaded_file.name
            st.session_state.cfops_extraidos = None
            st.session_state.xlsx_bytes      = None
            st.session_state.resultado       = None
            st.session_state.stats           = None
            st.session_state.tabela_acum_ok  = False
            st.session_state.log             = [
                f"Arquivo carregado: {uploaded_file.name} ({len(raw_atual)/1024:.1f} KB)"
            ]

    if st.session_state.arquivo_raw is not None:
        st.info(f"📄 Arquivo em memória: **{st.session_state.arquivo_nome}** "
                f"({len(st.session_state.arquivo_raw)/1024:.1f} KB)")

    col_e1, col_e2 = st.columns([1, 1])
    with col_e1:
        extrair = st.button("🔍 Extrair CFOPs e Gerar Planilha",
                            disabled=(st.session_state.arquivo_raw is None),
                            use_container_width=True, type="primary")
    with col_e2:
        if st.session_state.xlsx_bytes is not None:
            st.download_button(
                label="⬇ Baixar Planilha de Acumuladores (.xlsx)",
                data=st.session_state.xlsx_bytes,
                file_name=st.session_state.xlsx_nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )

    if extrair:
        st.session_state.log             = ["Extraindo CFOPs do SPED Fiscal..."]
        st.session_state.xlsx_bytes      = None
        st.session_state.cfops_extraidos = None
        try:
            content    = decode_arquivo(st.session_state.arquivo_raw)
            parsed     = parse_sped(content)
            st.session_state.log.append(f"Registros encontrados: {', '.join(parsed['por_tipo'].keys())}")
            cfops_dict = extrair_cfops_do_sped(parsed, tabela_flags, st.session_state.log)
            if not cfops_dict:
                st.session_state.log.append("AVISO: Nenhum CFOP encontrado.")
            else:
                xlsx_bytes = gerar_xlsx_acumuladores_tr(cfops_dict, tabela_cfop, tabela_flags)
                nome_base  = st.session_state.arquivo_nome.replace('.txt', '')
                nome_xlsx  = f"{nome_base}_acumuladores.xlsx"
                st.session_state.xlsx_bytes      = xlsx_bytes
                st.session_state.xlsx_nome       = nome_xlsx
                st.session_state.cfops_extraidos = cfops_dict
                n_dev = sum(1 for v in cfops_dict.values() if v.get('indDevol') == 1)
                st.session_state.log.append(
                    f"✔ {len(cfops_dict)} CFOP(s) extraído(s) | {n_dev} devolução(ões) | Planilha: {nome_xlsx}"
                )
        except Exception:
            st.session_state.log.append("ERRO FATAL na extração de CFOPs.")
            st.session_state.log.append(traceback.format_exc())
        st.rerun()

    if st.session_state.cfops_extraidos:
        cfops_dict = st.session_state.cfops_extraidos
        n_ent = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Entrada')
        n_sai = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Saída')
        n_dev = sum(1 for v in cfops_dict.values() if v.get('indDevol') == 1)
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("CFOPs únicos", len(cfops_dict))
        col_m2.metric("Entradas",     n_ent)
        col_m3.metric("Saídas",       n_sai)
        col_m4.metric("Devoluções",   n_dev)
        with st.expander("📋 CFOPs identificados no SPED", expanded=False):
            rows = [{'CFOP': cfop, 'Descrição (RF)': get_descricao_cfop(cfop, tabela_cfop),
                     'Tipo': info['tipo_operacao'],
                     'Devolução': '✔' if info.get('indDevol') == 1 else '',
                     'Ocorrências': info['ocorrencias'],
                     'Registros': ', '.join(sorted(info['registros']))}
                    for cfop, info in sorted(cfops_dict.items())]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════
    # ETAPA 2
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### ▶ Etapa 2 — Converter com a tabela de acumuladores preenchida")

    arquivo_acum = st.file_uploader(
        "📂 Tabela de Acumuladores preenchida (.xlsx ou .csv)",
        type=["xlsx", "xls", "csv"],
        help="Planilha gerada na Etapa 1 com a coluna ACUMULADOR preenchida",
        key="upload_acum",
    )

    if arquivo_acum is not None:
        log_temp = []
        raw_acum = arquivo_acum.read()
        tab_prev = carregar_acumuladores(raw_acum, arquivo_acum.name, log_temp)
        arquivo_acum.seek(0)
        if tab_prev is not None:
            st.success(f"✅ Tabela válida — **{len(tab_prev)} CFOPs** com acumulador preenchido.")
            st.session_state.tabela_acum_ok = True
        else:
            for msg in log_temp: st.error(msg)
            st.session_state.tabela_acum_ok = False
    else:
        if not st.session_state.tabela_acum_ok:
            st.info("⬆ Faça o upload da tabela de acumuladores preenchida para converter.")

    pode_converter = st.session_state.tabela_acum_ok and st.session_state.arquivo_raw is not None

    col1, col2 = st.columns([1, 1])
    with col1:
        converter = st.button("▶ Converter SPED → Domínio",
                              disabled=not pode_converter,
                              use_container_width=True, type="primary")
    with col2:
        limpar = st.button("🗑 Limpar Tudo", use_container_width=True)

    if limpar:
        for k in list(st.session_state.keys()): del st.session_state[k]
        st.rerun()

    if converter and pode_converter:
        st.session_state.log       = ["Iniciando conversão SPED → Domínio Sistemas..."]
        st.session_state.resultado = None
        st.session_state.stats     = None
        try:
            arquivo_acum.seek(0)
            tabela_acum = carregar_acumuladores(
                arquivo_acum.read(), arquivo_acum.name, st.session_state.log,
            )
            if tabela_acum is None:
                st.session_state.log.append("ERRO: Tabela inválida. Abortando.")
                st.rerun()

            content = decode_arquivo(st.session_state.arquivo_raw)
            parsed  = parse_sped(content)

            resultado_txt, stats = converter_sped_para_dominio(
                parsed, tabela_acum, tabela_flags, por_nome_imp, st.session_state.log
            )
            resultado_bytes = encode_ansi_seguro(resultado_txt, st.session_state.log)
            st.session_state.resultado  = resultado_bytes
            st.session_state.stats      = stats
            st.session_state.nome_saida = st.session_state.arquivo_nome.replace('.txt', '_dominio.txt')
        except Exception:
            st.session_state.log.append("ERRO FATAL durante a conversão.")
            st.session_state.log.append(traceback.format_exc())
        st.rerun()

    if st.session_state.resultado is not None:
        st.success("✅ Arquivo convertido com sucesso!")
        stats = st.session_state.stats or {}
        st.markdown("#### 📊 Estatísticas da Conversão")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("NFs Entrada",  stats.get('nf_entrada',  0))
        col2.metric("NFs Saída",    stats.get('nf_saida',    0))
        col3.metric("Devoluções",   stats.get('devolucoes',  0))
        col4.metric("Itens",        stats.get('itens',       0))
        col5, col6, col7 = st.columns(3)
        col5.metric("Analíticos",   stats.get('analiticos',  0))
        col6.metric("Transporte",   stats.get('transporte',  0))
        col7.metric("Erros",        stats.get('erros',       0))
        st.markdown("---")
        with st.expander("👁️ Prévia do arquivo gerado (primeiras 60 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado.decode('latin-1', errors='replace').splitlines()[:60]
            )
            st.code(preview, language='text')
        st.download_button(
            label="⬇ Baixar Arquivo Domínio Sistemas",
            data=st.session_state.resultado,
            file_name=st.session_state.nome_saida,
            mime="text/plain", use_container_width=True, type="primary",
        )

    # ── Log ────────────────────────────────────────────────────────────────
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"
    st.markdown(
        f"""<div style="background:#FCFCFC; border:1px solid {cor_borda}; border-radius:6px;
                    padding:14px; font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:340px; overflow-y:auto; color:#1F1F1F;">
{log_texto}</div>""", unsafe_allow_html=True,
    )
    st.markdown("---")
    st.caption("Conversor SPED Fiscal → Domínio Sistemas | Thomson Reuters | Python + Streamlit")


if __name__ == "__main__":
    main()

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
VERSAO          = "V2.8"
NOME_CFOP_XLSX  = "160314_Tabela_CFOP.xlsx"
NOME_IMP_XLSX   = "Impostos.xlsx"

# ==============================
# MAPEAMENTO FIXO: nome/fragmento → código Domínio
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
# ==============================
_ALIQ_TOL = 0.001
ALIQ_PIS_COFINS = {
    'PIS':    [(0.65, _ALIQ_TOL,  4), (1.65, _ALIQ_TOL, 17)],
    'COFINS': [(3.00, _ALIQ_TOL,  5), (7.60, _ALIQ_TOL, 19)],
}

# ==============================
# CSTs QUE EXIGEM NATUREZA DE PIS/COFINS
# ==============================
# CSTs que, quando presentes, devem ter o campo Natureza preenchido
CSTS_COM_NATUREZA = {
    '04': 'Operação Tributável Monofásica – Revenda a Alíquota Zero',
    '05': 'Operação Tributável por Substituição Tributária',
    '06': 'Operação Tributável a Alíquota Zero',
    '07': 'Operação Isenta da Contribuição',
    '08': 'Operação sem Incidência da Contribuição',
    '09': 'Operação com Suspensão da Contribuição',
    # CSTs de 2 dígitos com zeros à esquerda
    '004': 'Operação Tributável Monofásica – Revenda a Alíquota Zero',
    '005': 'Operação Tributável por Substituição Tributária',
    '006': 'Operação Tributável a Alíquota Zero',
    '007': 'Operação Isenta da Contribuição',
    '008': 'Operação sem Incidência da Contribuição',
    '009': 'Operação com Suspensão da Contribuição',
    # CSTs de 3 dígitos (SPED às vezes usa 3 dígitos)
    '50':  'Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Tributada',
    '51':  'Operação com Direito a Crédito – Vinculada Exclusivamente a Receita Não Tributada',
    '52':  'Operação com Direito a Crédito – Vinculada Exclusivamente a Receita de Exportação',
    '53':  'Operação com Direito a Crédito – Vinculada a Receitas Tributadas e Não-Tributadas',
    '54':  'Operação com Direito a Crédito – Vinculada a Receitas Tributadas e de Exportação',
    '55':  'Operação com Direito a Crédito – Vinculada a Receitas Não-Tributadas e de Exportação',
    '56':  'Operação com Direito a Crédito – Vinculada a Receitas Tributadas, Não-Tributadas e de Exportação',
    '60':  'Crédito Presumido – Oper. de Aquisição Vinculada Exclusivamente a Receita Tributada',
    '61':  'Crédito Presumido – Oper. de Aquisição Vinculada Exclusivamente a Receita Não-Tributada',
    '62':  'Crédito Presumido – Oper. de Aquisição Vinculada Exclusivamente a Receita de Exportação',
    '63':  'Crédito Presumido – Oper. de Aquisição Vinculada a Receitas Tributadas e Não-Tributadas',
    '64':  'Crédito Presumido – Oper. de Aquisição Vinculada a Receitas Tributadas e de Exportação',
    '65':  'Crédito Presumido – Oper. de Aquisição Vinculada a Receitas Não-Tributadas e de Exportação',
    '66':  'Crédito Presumido – Oper. de Aquisição Vinculada a Receitas Tributadas, Não-Tributadas e de Exportação',
    '67':  'Crédito Presumido – Outras Operações',
    '70':  'Operação de Aquisição sem Direito a Crédito',
    '71':  'Operação de Aquisição com Isenção',
    '72':  'Operação de Aquisição com Suspensão',
    '73':  'Operação de Aquisição a Alíquota Zero',
    '74':  'Operação de Aquisição sem Incidência da Contribuição',
    '75':  'Operação de Aquisição por Substituição Tributária',
    '98':  'Outras Operações de Entrada',
    '99':  'Outras Operações',
    '2':   'Operação Tributável com Alíquota Diferenciada',
    '02':  'Operação Tributável com Alíquota Diferenciada',
}

# CSTs que REALMENTE precisam do campo Natureza (conforme leiaute Domínio)
CSTS_NATUREZA_OBRIGATORIA = {'04', '05', '06', '07', '08', '09',
                              '4',  '5',  '6',  '7',  '8',  '9'}

# Opções de Natureza de Receita (campo 71/72 do 2030 e campo 67 do 1030)
# Conforme tabela Domínio Sistemas
NATUREZA_RECEITA_OPCOES = {
    '': '-- Não informar --',
    '01': '01 – Receita de Venda de Bens e Serviços',
    '02': '02 – Receita de Prestação de Serviços',
    '03': '03 – Receita de Locação de Bens Móveis',
    '04': '04 – Receita de Locação de Bens Imóveis',
    '05': '05 – Receita de Juros',
    '06': '06 – Receita de Dividendos',
    '07': '07 – Receita Financeira',
    '08': '08 – Receita de Exportação',
    '09': '09 – Receita de Atividade Imobiliária',
    '10': '10 – Receita de Serviços de Telecomunicações',
    '11': '11 – Receita de Serviços de Transporte',
    '12': '12 – Receita de Atividade de Seguros',
    '13': '13 – Receita de Atividade de Previdência Privada',
    '14': '14 – Receita de Atividade de Saúde',
    '15': '15 – Receita de Serviços de Educação',
    '16': '16 – Receita de Serviços Hospitalares',
    '17': '17 – Receita de Serviços de Limpeza e Conservação',
    '18': '18 – Receita de Serviços de Vigilância',
    '19': '19 – Receita de Serviços de Construção Civil',
    '20': '20 – Receita de Serviços de Informática',
    '21': '21 – Receita Cooperativa',
    '99': '99 – Outras Receitas',
}

# Opções de Vínculo de Crédito (campo 72/73 do 1030 e 77/78 do 2030)
VINCULO_CREDITO_OPCOES = {
    '': '-- Não informar --',
    '01': '01 – Crédito vinculado à alíquota básica',
    '02': '02 – Crédito vinculado à alíquota diferenciada',
    '03': '03 – Crédito vinculado à alíquota por unidade de produto',
    '05': '05 – Crédito vinculado a aquisição de embalagem',
    '06': '06 – Crédito presumido agroindústria e aquisição de combustível',
    '08': '08 – Crédito de importação',
    '99': '99 – Outros créditos',
}

# ==============================
# CONFIGURAÇÃO PADRÃO DE NATUREZA POR CST
# ==============================
NATUREZA_PADRAO_POR_CST = {
    '04': '', '05': '', '06': '', '07': '', '08': '', '09': '',
}


def _aliq_float(valor: str) -> float:
    v = str(valor).strip().replace(',', '.')
    if not v or v.upper() == 'NAN':
        return 0.0
    try:
        f = float(v)
        if 0 < f < 0.20:
            f = round(f * 100, 6)
        return f
    except ValueError:
        return 0.0


def get_codigo_pis(aliq_str: str, por_nome: dict) -> int:
    aliq = _aliq_float(aliq_str)
    for (central, tol, cod) in ALIQ_PIS_COFINS['PIS']:
        if abs(aliq - central) <= tol:
            return cod
    return get_codigo_imposto('PIS', por_nome, 4)


def get_codigo_cofins(aliq_str: str, por_nome: dict) -> int:
    aliq = _aliq_float(aliq_str)
    for (central, tol, cod) in ALIQ_PIS_COFINS['COFINS']:
        if abs(aliq - central) <= tol:
            return cod
    return get_codigo_imposto('COFINS', por_nome, 5)


def converter_data(data_sped: str) -> str:
    d = str(data_sped).strip().replace('/', '').replace('-', '')
    if len(d) == 8 and d.isdigit():
        return f"{d[0:2]}/{d[2:4]}/{d[4:8]}"
    return data_sped


def converter_data_planilha(valor) -> str:
    if pd.isna(valor) if not isinstance(valor, str) else False:
        return ''
    s = str(valor).strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%d%m%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%d/%m/%Y')
        except ValueError:
            continue
    return s


def _normalizar_cst(cst: str) -> str:
    """Normaliza CST removendo zeros à esquerda para comparação."""
    s = str(cst).strip().split('.')[0]
    try:
        return str(int(s))
    except ValueError:
        return s


def get_natureza_por_cst(
    cst_pis: str,
    cst_cofins: str,
    config_natureza: dict,
) -> tuple:
    """
    Retorna (natureza_pis, natureza_cofins) com base na configuração do usuário.
    config_natureza: dict { 'cst_XX_pis': 'NN', 'cst_XX_cofins': 'NN' }
    """
    cst_p = _normalizar_cst(cst_pis)
    cst_c = _normalizar_cst(cst_cofins)

    nat_pis    = ''
    nat_cofins = ''

    # Busca natureza configurada para o CST do PIS
    chave_pis = f"cst_{cst_p.zfill(2)}_pis"
    if chave_pis in config_natureza:
        nat_pis = config_natureza[chave_pis]

    # Busca natureza configurada para o CST do COFINS
    chave_cof = f"cst_{cst_c.zfill(2)}_cofins"
    if chave_cof in config_natureza:
        nat_cofins = config_natureza[chave_cof]

    return nat_pis, nat_cofins


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
        .natureza-box {
            background-color: #FFF8F0; border-left: 4px solid #FF8000;
            border-radius: 4px; padding: 14px 18px; margin: 8px 0;
        }
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
                col_cod, col_nome = df.columns[0], df.columns[1]
            else:
                for nome, cod in IMPOSTOS_FALLBACK.items():
                    por_codigo[cod] = nome; por_nome[nome.upper()] = cod
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
            por_codigo[cod] = nome
            por_nome[nome.upper()] = cod
    except Exception:
        for nome, cod in IMPOSTOS_FALLBACK.items():
            por_codigo[cod] = nome; por_nome[nome.upper()] = cod

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
# CARREGAMENTO DA PLANILHA CLIENTE
# ==============================
def carregar_planilha_cliente(arquivo_bytes: bytes, nome_arquivo: str, log: list) -> pd.DataFrame | None:
    try:
        ext = os.path.splitext(nome_arquivo)[1].lower()
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(io.BytesIO(arquivo_bytes), dtype=str)
        else:
            raw_str = arquivo_bytes.decode('latin-1', errors='replace')
            sep     = ';' if raw_str.count(';') >= raw_str.count(',') else ','
            df      = pd.read_csv(io.StringIO(raw_str), sep=sep, dtype=str)

        df.columns = [str(c).strip().upper() for c in df.columns]

        colunas_obrigatorias = [
            'NF', 'COD.ITEM',
            'PIS CST', 'PIS PC ALIQ', 'PIS VL BASE', 'PIS VALOR',
            'COFINS CST', 'COFINS PC ALIQ', 'COFINS VL BASE', 'COFINS VALOR',
        ]
        ausentes = [c for c in colunas_obrigatorias if c not in df.columns]
        if ausentes:
            log.append(f"ERRO Planilha Cliente: colunas ausentes → {ausentes}")
            log.append(f"Colunas encontradas: {list(df.columns)}")
            return None

        df['NF']       = df['NF'].apply(lambda x: str(x).strip().split('.')[0])
        df['COD.ITEM'] = df['COD.ITEM'].apply(lambda x: str(x).strip().split('.')[0])

        log.append(f"Planilha Cliente carregada: {len(df)} linhas | "
                   f"NFs únicas: {df['NF'].nunique()} | "
                   f"Itens únicos: {df['COD.ITEM'].nunique()}")
        return df

    except Exception as e:
        log.append(f"ERRO ao carregar Planilha Cliente: {e}")
        log.append(traceback.format_exc())
        return None


def _safe_str(valor, casas: int = 2) -> str:
    s = str(valor).strip()
    if not s or s.upper() == 'NAN':
        return f"0,{'0' * casas}"
    s = s.replace(',', '.')
    try:
        f = float(s)
        return f"{f:.{casas}f}".replace('.', ',')
    except ValueError:
        return f"0,{'0' * casas}"


def _safe_int(valor) -> str:
    s = str(valor).strip()
    if not s or s.upper() == 'NAN':
        return ''
    s = s.split('.')[0]
    return s


def buscar_pis_cofins_planilha(
    df_cliente: pd.DataFrame,
    num_nf: str,
    cod_item: str,
    log: list,
) -> dict | None:
    nf_norm   = str(num_nf).strip().split('.')[0]
    item_norm = str(cod_item).strip().split('.')[0]

    mascara = (df_cliente['NF'] == nf_norm) & (df_cliente['COD.ITEM'] == item_norm)
    linhas  = df_cliente[mascara]

    if linhas.empty:
        log.append(f"  AVISO: NF={nf_norm} / Item={item_norm} não encontrado na Planilha Cliente.")
        return None

    row = linhas.iloc[0]

    def _col(nome: str, default: str = '') -> str:
        return str(row[nome]).strip() if nome in df_cliente.columns else default

    resultado = {
        'pis_cst':      _safe_int(_col('PIS CST')),
        'pis_base':     _safe_str(_col('PIS VL BASE'), 2),
        'pis_aliq':     _safe_str(_col('PIS PC ALIQ'), 4),
        'pis_valor':    _safe_str(_col('PIS VALOR'),   2),
        'cofins_cst':   _safe_int(_col('COFINS CST')),
        'cofins_base':  _safe_str(_col('COFINS VL BASE'), 2),
        'cofins_aliq':  _safe_str(_col('COFINS PC ALIQ'), 4),
        'cofins_valor': _safe_str(_col('COFINS VALOR'),   2),
        'cbs_class':    _safe_int(_col('CLASS TRIB CBS')),
        'cbs_base':     _safe_str(_col('CBS BASE'),  2),
        'cbs_aliq':     _safe_str(_col('CBS ALIQ'),  2),
        'cbs_valor':    _safe_str(_col('CBS VALOR'), 2),
        'ibs_class':    _safe_int(_col('CLASS TRIB IBS')),
        'ibs_base':     _safe_str(_col('IBS BASE'),  2),
        'ibs_aliq':     _safe_str(_col('IBS ALIQ'),  2),
        'ibs_valor':    _safe_str(_col('IBS VALOR'), 2),
    }
    return resultado


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
SPED_0000_DT_INI     = 3
SPED_0000_DT_FIN     = 4
SPED_0150_COD        = 1
SPED_0150_CNPJ       = 4
SPED_0200_COD_ITEM   = 1
SPED_0200_DESCR      = 2
SPED_0200_COD_BARRA  = 3
SPED_0200_COD_ANT    = 4
SPED_0200_UNID_INV   = 5
SPED_0200_TIPO_ITEM  = 6
SPED_0200_COD_NCM    = 7
SPED_0200_EX_IPI     = 8
SPED_0200_COD_GEN    = 9
SPED_0200_COD_LST    = 10
SPED_0200_ALIQ_ICMS  = 11
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
# EXTRAÇÃO DE PRODUTOS DO SPED
# ==============================
def extrair_produtos_do_sped(parsed: dict, log: list) -> dict:
    produtos = {}

    if '0200' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['0200']:
            cod  = _c(campos, SPED_0200_COD_ITEM).strip()
            if not cod: continue
            produtos[cod] = {
                'descr':      _c(campos, SPED_0200_DESCR),
                'unid':       _c(campos, SPED_0200_UNID_INV),
                'ncm':        _c(campos, SPED_0200_COD_NCM),
                'tipo_item':  _c(campos, SPED_0200_TIPO_ITEM),
                'aliq_icms':  _c(campos, SPED_0200_ALIQ_ICMS) or '0,00',
                'aliq_ipi':   '0,00',
                'cst_icms':   '',
                'cst_ipi':    '',
                'cst_pis':    '',
                'cst_cofins': '',
                'aliq_pis':   '0,00',
                'aliq_cofins':'0,00',
            }
        log.append(f"Produtos carregados do 0200: {len(produtos)}")

    itens_c170 = 0
    for tipo, campos, _ in parsed['linhas_ordenadas']:
        if tipo != 'C170': continue
        cod  = _c(campos, SPED_C170_COD_ITEM).strip()
        if not cod: continue
        itens_c170 += 1

        aliq_icms_c170  = _c(campos, SPED_C170_ALIQ_ICMS)  or '0,00'
        aliq_ipi_c170   = _c(campos, SPED_C170_ALIQ_IPI)   or '0,00'
        aliq_pis_c170   = _c(campos, SPED_C170_ALIQ_PIS)   or '0,00'
        aliq_cof_c170   = _c(campos, SPED_C170_ALIQ_COF)   or '0,00'
        cst_icms_c170   = _c(campos, SPED_C170_CST_ICMS)
        cst_ipi_c170    = _c(campos, SPED_C170_CST_IPI)
        cst_pis_c170    = _c(campos, SPED_C170_CST_PIS)
        cst_cofins_c170 = _c(campos, SPED_C170_CST_COFINS)

        if cod not in produtos:
            produtos[cod] = {
                'descr':      _c(campos, SPED_C170_DESCR),
                'unid':       _c(campos, SPED_C170_UNID),
                'ncm':        '',
                'tipo_item':  '00',
                'aliq_icms':  aliq_icms_c170,
                'aliq_ipi':   aliq_ipi_c170,
                'cst_icms':   cst_icms_c170,
                'cst_ipi':    cst_ipi_c170,
                'cst_pis':    cst_pis_c170,
                'cst_cofins': cst_cofins_c170,
                'aliq_pis':   aliq_pis_c170,
                'aliq_cofins':aliq_cof_c170,
            }
        else:
            p = produtos[cod]
            if not p.get('cst_icms'):   p['cst_icms']   = cst_icms_c170
            if not p.get('cst_ipi'):    p['cst_ipi']    = cst_ipi_c170
            if not p.get('cst_pis'):    p['cst_pis']    = cst_pis_c170
            if not p.get('cst_cofins'): p['cst_cofins'] = cst_cofins_c170
            if p.get('aliq_ipi',   '0,00') in ('', '0,00', '0'): p['aliq_ipi']    = aliq_ipi_c170
            if p.get('aliq_pis',   '0,00') in ('', '0,00', '0'): p['aliq_pis']    = aliq_pis_c170
            if p.get('aliq_cofins','0,00') in ('', '0,00', '0'): p['aliq_cofins'] = aliq_cof_c170
            if p.get('aliq_icms',  '0,00') in ('', '0,00', '0'): p['aliq_icms']   = aliq_icms_c170
            if not p.get('unid'):  p['unid']  = _c(campos, SPED_C170_UNID)
            if not p.get('descr'): p['descr'] = _c(campos, SPED_C170_DESCR)

    log.append(f"Itens C170 processados: {itens_c170} | Total produtos únicos: {len(produtos)}")
    return produtos


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
# GERAÇÃO DOS REGISTROS 0100 + 0110
# ==============================
def gerar_registros_produtos(produtos: dict, dt_ini: str, por_nome_imp: dict, log: list) -> str:
    saida = StringIO()
    n_prod = 0
    n_vig  = 0

    for cod, p in sorted(produtos.items()):
        descr      = (p.get('descr') or '').replace('|', ' ')[:60]
        unid       = p.get('unid', 'UN') or 'UN'
        ncm        = p.get('ncm', '') or ''
        aliq_icms  = p.get('aliq_icms', '0,00') or '0,00'
        aliq_ipi   = p.get('aliq_ipi',  '0,00') or '0,00'
        cst_icms   = p.get('cst_icms',  '') or ''
        cst_ipi    = p.get('cst_ipi',   '') or ''
        cst_pis    = p.get('cst_pis',   '') or ''
        cst_cofins = p.get('cst_cofins','') or ''
        aliq_pis   = p.get('aliq_pis',  '0,00') or '0,00'
        aliq_cof   = p.get('aliq_cofins','0,00') or '0,00'

        campos_opcionais = '|' * 64

        saida.write(
            f"|0100|{cod}|{descr}|||{ncm}||||{unid}|N|O|||"
            f"|N||0,000|0,00000|0,000|{cst_icms}|{aliq_icms}|{aliq_ipi}|M||N|"
            f"{campos_opcionais}\n"
        )
        n_prod += 1

        campos_adicionais_0110 = '|' * 30

        saida.write(
            f"|0110|Vigência||01|N|N|"
            f"{aliq_pis}|{aliq_cof}|N|N|||"
            f"0,0000|0,0000|"
            f"{cst_cofins}|N||||||"
            f"N|{aliq_pis}|{aliq_cof}|N|N|||"
            f"0,0000|0,0000|||"
            f"N|N|"
            f"{cst_icms}|{cst_icms}|{aliq_icms}|"
            f"{cst_ipi}|{cst_ipi}|M|{aliq_ipi}|"
            f"{campos_adicionais_0110}\n"
        )
        n_vig += 1

    log.append(f"Produtos gerados: {n_prod} registros 0100 + {n_vig} registros 0110")
    return saida.getvalue()


# ==============================
# WIDGET DE CONFIGURAÇÃO DE NATUREZA PIS/COFINS
# ==============================
def render_configuracao_natureza() -> dict:
    """
    Renderiza o painel de configuração de Natureza de PIS/COFINS por CST.
    Retorna dict com as configurações do usuário.
    Chaves: 'cst_XX_pis' e 'cst_XX_cofins' → valor da natureza selecionada.
    """
    st.markdown("### 🏷️ Configuração de Natureza de PIS/COFINS por CST")

    st.markdown("""
    <div class="natureza-box">
    <strong>ℹ️ O que é Natureza de PIS/COFINS?</strong><br>
    Para CSTs que representam operações não tributadas (04, 05, 06, 07, 08, 09),
    o Domínio Sistemas exige que seja informado o código de <strong>Natureza da Receita</strong>
    (campo 71/72 do registro 2030 para saídas e campo 67 do registro 1030 para entradas).
    Configure abaixo qual natureza usar para cada CST encontrado na operação.
    </div>
    """, unsafe_allow_html=True)

    # CSTs que precisam de natureza
    csts_config = [
        ('04', 'CST 04 – Monofásica / Revenda Alíquota Zero'),
        ('05', 'CST 05 – Substituição Tributária'),
        ('06', 'CST 06 – Alíquota Zero'),
        ('07', 'CST 07 – Isenta da Contribuição'),
        ('08', 'CST 08 – Sem Incidência'),
        ('09', 'CST 09 – Suspensão da Contribuição'),
    ]

    config = {}

    # Opções para o selectbox — lista de labels
    opcoes_labels = list(NATUREZA_RECEITA_OPCOES.values())
    opcoes_codigos = list(NATUREZA_RECEITA_OPCOES.keys())

    # Recupera configuração anterior da session_state
    cfg_anterior = st.session_state.get('config_natureza', {})

    with st.expander("⚙️ Configurar Natureza por CST — clique para expandir", expanded=True):
        col_info, col_pis, col_cofins = st.columns([2, 1.5, 1.5])
        with col_info:
            st.markdown("**CST / Descrição**")
        with col_pis:
            st.markdown("**Natureza do PIS**")
        with col_cofins:
            st.markdown("**Natureza do COFINS**")

        st.markdown("---")

        for cst_cod, cst_descr in csts_config:
            chave_pis = f"cst_{cst_cod}_pis"
            chave_cof = f"cst_{cst_cod}_cofins"

            # Recupera valor anterior ou vazio
            val_pis_ant = cfg_anterior.get(chave_pis, '')
            val_cof_ant = cfg_anterior.get(chave_cof, '')

            # Índice atual
            idx_pis = opcoes_codigos.index(val_pis_ant) if val_pis_ant in opcoes_codigos else 0
            idx_cof = opcoes_codigos.index(val_cof_ant) if val_cof_ant in opcoes_codigos else 0

            col_i, col_p, col_c = st.columns([2, 1.5, 1.5])
            with col_i:
                st.markdown(
                    f"<div style='padding:6px 0; font-size:13px;'>"
                    f"<strong style='color:#FF8000;'>CST {cst_cod}</strong><br>"
                    f"<span style='color:#666; font-size:11px;'>{cst_descr.split('–')[1].strip() if '–' in cst_descr else cst_descr}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            with col_p:
                sel_pis = st.selectbox(
                    label=f"PIS CST {cst_cod}",
                    options=opcoes_labels,
                    index=idx_pis,
                    key=f"nat_pis_{cst_cod}",
                    label_visibility="collapsed",
                )
                # Converte label → código
                cod_pis = opcoes_codigos[opcoes_labels.index(sel_pis)]
                config[chave_pis] = cod_pis

            with col_c:
                sel_cof = st.selectbox(
                    label=f"COFINS CST {cst_cod}",
                    options=opcoes_labels,
                    index=idx_cof,
                    key=f"nat_cof_{cst_cod}",
                    label_visibility="collapsed",
                )
                cod_cof = opcoes_codigos[opcoes_labels.index(sel_cof)]
                config[chave_cof] = cod_cof

        st.markdown("---")

        # Resumo visual das configurações ativas
        configs_ativas = {k: v for k, v in config.items() if v}
        if configs_ativas:
            st.markdown("**✅ Configurações ativas:**")
            resumo_cols = st.columns(3)
            col_idx = 0
            for chave, valor in configs_ativas.items():
                partes = chave.split('_')  # ['cst', 'XX', 'pis'/'cofins']
                cst_n  = partes[1]
                imp    = partes[2].upper()
                label  = NATUREZA_RECEITA_OPCOES.get(valor, valor)
                resumo_cols[col_idx % 3].markdown(
                    f"<div style='background:#E8F5E9; border-left:3px solid #388E3C; "
                    f"padding:6px 10px; border-radius:4px; margin:3px 0; font-size:12px;'>"
                    f"<strong>CST {cst_n} / {imp}</strong><br>{valor} – {label.split('–')[1].strip() if '–' in label else label}"
                    f"</div>",
                    unsafe_allow_html=True
                )
                col_idx += 1
        else:
            st.info("Nenhuma natureza configurada. Os campos serão deixados em branco no arquivo gerado.")

    # Salva na session_state para persistir entre reruns
    st.session_state['config_natureza'] = config
    return config


# ==============================
# CONVERSÃO SPED FISCAL → DOMÍNIO SISTEMAS
# V2.8 — com Natureza de PIS/COFINS por CST
# ==============================
def converter_sped_para_dominio(
    parsed: dict,
    tabela_acum: dict,
    tabela_flags: dict,
    por_nome_imp: dict,
    log: list,
    df_cliente: pd.DataFrame | None = None,
    config_natureza: dict | None = None,
) -> tuple:
    saida        = StringIO()
    nao_mapeados = set()
    stats = {
        'nf_entrada': 0, 'nf_saida': 0, 'itens': 0,
        'analiticos': 0, 'transporte': 0, 'inventario': 0,
        'devolucoes': 0, 'produtos': 0, 'erros': 0,
        'itens_planilha': 0, 'itens_sped': 0,
        'natureza_aplicada': 0,
    }

    cfg_nat = config_natureza or {}

    COD_ICMS    = get_codigo_imposto('ICMS',              por_nome_imp, 1)
    COD_IPI     = get_codigo_imposto('IPI',               por_nome_imp, 2)
    COD_ISS     = get_codigo_imposto('ISS',               por_nome_imp, 3)
    COD_ST      = get_codigo_imposto('SUBST. TRIBUTARIA', por_nome_imp, 9)
    COD_ICMS_ST = get_codigo_imposto('ICMS RETIDO',       por_nome_imp, 11)

    usa_planilha = df_cliente is not None and not df_cliente.empty
    log.append(
        f"Fonte PIS/COFINS/IBS/CBS: {'Planilha Cliente ✔' if usa_planilha else 'SPED Fiscal (fallback)'}"
    )
    log.append(
        f"Configuração de Natureza: {len([v for v in cfg_nat.values() if v])} CSTs com natureza definida"
    )

    dt_ini_sped = ''
    if '0000' in parsed['por_tipo']:
        campos_0000, _ = parsed['por_tipo']['0000'][0]
        dt_ini_sped = _c(campos_0000, SPED_0000_DT_INI)

    if '0000' in parsed['por_tipo']:
        campos, _ = parsed['por_tipo']['0000'][0]
        cnpj = _c(campos, SPED_0000_CNPJ)
        saida.write(f"|0000|{cnpj}|\n")
        log.append(f"0000: CNPJ={cnpj}")
    else:
        log.append("AVISO: Registro 0000 não encontrado.")

    participantes = {}
    if '0150' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['0150']:
            participantes[_c(campos, SPED_0150_COD)] = campos
        log.append(f"Participantes carregados: {len(participantes)}")

    produtos = extrair_produtos_do_sped(parsed, log)
    if produtos:
        bloco_produtos = gerar_registros_produtos(produtos, dt_ini_sped, por_nome_imp, log)
        saida.write(bloco_produtos)
        stats['produtos'] = len(produtos)
    else:
        log.append("AVISO: Nenhum produto encontrado no SPED (0200/C170).")

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
            dt_doc   = converter_data(_c(campos_c100, SPED_C100_DT_DOC))
            dt_es    = converter_data(_c(campos_c100, SPED_C100_DT_ES))
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

            # ── Itens C170 ────────────────────────────────────────────────
            for campos_c170 in bloco['c170']:
                cod_item  = _c(campos_c170, SPED_C170_COD_ITEM)
                num_item  = _c(campos_c170, SPED_C170_NUM_ITEM)
                qtd       = _c(campos_c170, SPED_C170_QTD)       or '0'
                unid      = _c(campos_c170, SPED_C170_UNID)
                vl_item   = _c(campos_c170, SPED_C170_VL_ITEM)   or '0,00'
                vl_desc_i = _c(campos_c170, SPED_C170_VL_DESC)   or '0,00'
                cfop_item = _c(campos_c170, SPED_C170_CFOP)

                # ICMS/IPI — sempre do SPED
                vl_bc_icms   = _c(campos_c170, SPED_C170_VL_BC)       or '0,00'
                aliq_icms    = _c(campos_c170, SPED_C170_ALIQ_ICMS)   or '0,00'
                vl_icms_i    = _c(campos_c170, SPED_C170_VL_ICMS)     or '0,00'
                vl_bc_st     = _c(campos_c170, SPED_C170_VL_BC_ST)    or '0,00'
                aliq_st      = _c(campos_c170, SPED_C170_ALIQ_ST)     or '0,00'
                vl_icms_st   = _c(campos_c170, SPED_C170_VL_ICMS_ST)  or '0,00'
                vl_bc_ipi    = _c(campos_c170, SPED_C170_VL_BC_IPI)   or '0,00'
                aliq_ipi     = _c(campos_c170, SPED_C170_ALIQ_IPI)    or '0,00'
                vl_ipi       = _c(campos_c170, SPED_C170_VL_IPI)      or '0,00'

                # PIS/COFINS/IBS/CBS — Planilha Cliente (prioridade)
                dados_planilha = None
                if usa_planilha:
                    dados_planilha = buscar_pis_cofins_planilha(
                        df_cliente, num_doc, cod_item, log
                    )

                if dados_planilha is not None:
                    pis_cst      = dados_planilha['pis_cst']
                    pis_base     = dados_planilha['pis_base']
                    pis_aliq     = dados_planilha['pis_aliq']
                    pis_valor    = dados_planilha['pis_valor']
                    cofins_cst   = dados_planilha['cofins_cst']
                    cofins_base  = dados_planilha['cofins_base']
                    cofins_aliq  = dados_planilha['cofins_aliq']
                    cofins_valor = dados_planilha['cofins_valor']
                    cbs_class    = dados_planilha['cbs_class']
                    cbs_base     = dados_planilha['cbs_base']
                    cbs_aliq     = dados_planilha['cbs_aliq']
                    cbs_valor    = dados_planilha['cbs_valor']
                    ibs_class    = dados_planilha['ibs_class']
                    ibs_base     = dados_planilha['ibs_base']
                    ibs_aliq     = dados_planilha['ibs_aliq']
                    ibs_valor    = dados_planilha['ibs_valor']
                    cod_pis    = get_codigo_pis(pis_aliq, por_nome_imp)
                    cod_cofins = get_codigo_cofins(cofins_aliq, por_nome_imp)
                    stats['itens_planilha'] += 1
                else:
                    pis_cst      = _c(campos_c170, SPED_C170_CST_PIS)
                    pis_base     = _c(campos_c170, SPED_C170_VL_BC_PIS)  or '0,00'
                    pis_aliq     = _c(campos_c170, SPED_C170_ALIQ_PIS)   or '0,00'
                    pis_valor    = _c(campos_c170, SPED_C170_VL_PIS)     or '0,00'
                    cofins_cst   = _c(campos_c170, SPED_C170_CST_COFINS)
                    cofins_base  = _c(campos_c170, SPED_C170_VL_BC_COF)  or '0,00'
                    cofins_aliq  = _c(campos_c170, SPED_C170_ALIQ_COF)   or '0,00'
                    cofins_valor = _c(campos_c170, SPED_C170_VL_COFINS)  or '0,00'
                    cbs_class    = ''
                    cbs_base     = '0,00'
                    cbs_aliq     = '0,00'
                    cbs_valor    = '0,00'
                    ibs_class    = ''
                    ibs_base     = '0,00'
                    ibs_aliq     = '0,00'
                    ibs_valor    = '0,00'
                    cod_pis    = get_codigo_pis(pis_aliq, por_nome_imp)
                    cod_cofins = get_codigo_cofins(cofins_aliq, por_nome_imp)
                    stats['itens_sped'] += 1

                # ── NATUREZA DE PIS/COFINS por CST ────────────────────────
                # Normaliza CST para 2 dígitos (ex: "50" → "50", "4" → "04")
                cst_p_norm = str(pis_cst).strip().split('.')[0]
                cst_c_norm = str(cofins_cst).strip().split('.')[0]
                try:
                    cst_p_norm = str(int(cst_p_norm)).zfill(2)
                except ValueError:
                    cst_p_norm = cst_p_norm.zfill(2)
                try:
                    cst_c_norm = str(int(cst_c_norm)).zfill(2)
                except ValueError:
                    cst_c_norm = cst_c_norm.zfill(2)

                nat_pis, nat_cofins = get_natureza_por_cst(
                    cst_p_norm, cst_c_norm, cfg_nat
                )

                if nat_pis or nat_cofins:
                    stats['natureza_aplicada'] += 1

                try:
                    vl_unit = f"{float(vl_item.replace(',', '.')) / float(qtd.replace(',', '.')):.3f}".replace('.', ',')
                except Exception:
                    vl_unit = vl_item

                def _val(s):
                    try: return float(str(s).replace(',', '.'))
                    except: return 0.0

                if _val(vl_icms_st) > 0:
                    cod_imp_item = COD_ICMS_ST; vl_bc_princ = vl_bc_st
                    aliq_princ = aliq_st; vl_imp_princ = vl_icms_st
                elif _val(vl_icms_i) > 0:
                    cod_imp_item = COD_ICMS; vl_bc_princ = vl_bc_icms
                    aliq_princ = aliq_icms; vl_imp_princ = vl_icms_i
                elif _val(vl_ipi) > 0:
                    cod_imp_item = COD_IPI; vl_bc_princ = vl_bc_ipi
                    aliq_princ = aliq_ipi; vl_imp_princ = vl_ipi
                else:
                    cod_imp_item = COD_ICMS; vl_bc_princ = vl_bc_icms
                    aliq_princ = aliq_icms; vl_imp_princ = vl_icms_i

                # ── Determina registro: 1030 (entrada) ou 2030 (saída) ────
                reg = '1030' if ind_oper == '0' else '2030'

                # ── IBS e CBS ─────────────────────────────────────────────
                campos_ibs_cbs = (
                    f"{ibs_class}|{ibs_base}|{ibs_aliq}|{ibs_valor}|"
                    f"{cbs_class}|{cbs_base}|{cbs_aliq}|{cbs_valor}"
                )

                if ind_oper == '0':
                    # ── REGISTRO 1030 (ENTRADA) ───────────────────────────
                    # Campo 67 = Base do crédito (natureza de entrada PIS)
                    # Campo 72 = Vínculo de Crédito PIS
                    # Campo 73 = Vínculo de Crédito COFINS
                    # Campos 41 = CST PIS | 43 = CST COFINS
                    # Campos 42 = Base PIS | 44 = Base COFINS
                    # Campos 36 = Alíq PIS | 37 = Vlr PIS
                    # Campos 38 = Alíq COFINS | 39 = Vlr COFINS
                    # Campos 104-111 = IBS/CBS
                    saida.write(
                        f"|1030|{cod_item}|{qtd}|{vl_unit}|0|0|1|{dt_doc}||"
                        f"{cod_sit}|{vl_item}|{vl_desc_i}|{vl_item}|0,00|"
                        f"{aliq_princ}|||"
                        # campos 18-20 (frete, seguro, despesas) = 0,00
                        # campos 21-25 (qtd gasolina, ICMS, SUBTRI, isentas, outras IPI) = 0,00
                        # campo 27 = valor unitário
                        f"{cod_pis}|{pis_base}|{pis_aliq}|{pis_valor}|0,000|{vl_imp_princ}|"
                        f"{cod_cofins}|{cofins_base}|{cofins_aliq}|{cofins_valor}||{vl_item}|0,00|"
                        f"{cod_imp_item}|{vl_bc_princ}|{vl_ipi}|{vl_ipi}|0,00|"
                        f"{cfop_item}||0,0000|0,00|0,00|0,00|{vl_bc_icms}|"
                        f"{COD_ICMS_ST}|{vl_bc_st}|{COD_ICMS_ST}|{vl_icms_st}|"
                        # campo 41 = CST PIS | campo 43 = CST COFINS
                        f"{pis_cst}|{cofins_cst}|||"
                        f"{dt_doc}|{dt_doc}||||||S|{unid}|||"
                        f"{vl_item}|||||||1|||||01|01|||"
                        # campo 67 = Base do crédito (natureza PIS entrada)
                        f"{nat_pis}|"
                        # campo 68 = Nº nota devolvida
                        # campo 69 = Descrição complementar
                        # campo 70 = CST PIS nota devolvida
                        # campo 71 = CST COFINS nota devolvida
                        # campo 72 = Vínculo crédito PIS
                        # campo 73 = Vínculo crédito COFINS
                        f"||||{nat_pis}|{nat_cofins}|"
                        # campos 74-75 = Exclusão PIS/COFINS
                        # campos 76-78 = ICMS Carga Média
                        f"||0,00|0,00|0,00|"
                        # campos 79-84 = ECF, Redução, Cód.Rec.PIS/COFINS devolvidos
                        f"||0,00|||"
                        # campos 83-84 = Cód.Rec.PIS/COFINS
                        f"{nat_pis}|{nat_cofins}|"
                        # campos 85-89 = Crédito Presumido, ICMS ST Antecipação
                        f"0,00|0,00|0,00|0,00|0,00|"
                        # campos 90-103 = IPI, CEST, ICMS ST Retido, Identificador, etc.
                        f"||0,00|0,00|||||||||||"
                        # campos 104-111 = IBS e CBS
                        f"{campos_ibs_cbs}||||\n"
                    )
                else:
                    # ── REGISTRO 2030 (SAÍDA) ─────────────────────────────
                    # Campo 48 = CST PIS | campo 49 = Base PIS
                    # Campo 50 = Alíq PIS | campo 51 = Vlr PIS
                    # Campo 52 = CST COFINS | campo 53 = Base COFINS
                    # Campo 54 = Alíq COFINS | campo 55 = Vlr COFINS
                    # Campo 71 = Natureza da receita PIS
                    # Campo 72 = Natureza da receita COFINS
                    # Campos 111-118 = IBS e CBS
                    saida.write(
                        f"|2030|{cod_item}|{qtd}|{vl_unit}|0|0|1|{dt_doc}|"
                        # campo 9 = CST ICMS | campo 10 = Vlr Bruto | campo 11 = Desconto
                        f"{cod_sit}|{vl_item}|{vl_desc_i}|"
                        # campo 12 = Base ICMS | campo 13 = Base ST | campo 14 = Alíq ICMS
                        f"{vl_bc_icms}|{vl_bc_st}|{aliq_princ}|"
                        # campos 15-18 = Chassi, Incentivado, Apuração, Sit.ECF
                        f"||||"
                        # campos 19-21 = Frete, Seguro, Despesas
                        f"0,00|0,00|0,00|"
                        # campos 22-24 = Qtd Gasolina, Vlr ICMS, Vlr SUBTRI
                        f"0,00|{vl_icms_i}|{vl_icms_st}|"
                        # campos 25-26 = Isentas IPI, Outras IPI
                        f"0,00|0,00|"
                        # campo 27 = Valor Unitário
                        f"{vl_unit}|"
                        # campos 28-30 = Alíq ST, Cód Trib IPI, Alíq IPI
                        f"{aliq_st}|{cod_imp_item}|{aliq_ipi}|"
                        # campos 31-33 = Base ISSQN, Alíq ISSQN, Vlr ISSQN
                        f"0,00|0,00|0,00|"
                        # campos 34-47 = Lote Med, Validade, Ref.BC, Vl.Tab, Arma, Cano, TipoOp,
                        #                QtdCancel, VlrCancel, Isentas, NaoInc, AcumST, FabMed, ECF
                        f"|||||||0,00|0,00|0,00|0,00|0,00|||"
                        # campo 48 = CST PIS
                        f"{pis_cst}|"
                        # campo 49 = Base PIS
                        f"{pis_base}|"
                        # campo 50 = Alíq PIS
                        f"{pis_aliq}|"
                        # campo 51 = Vlr PIS
                        f"{pis_valor}|"
                        # campo 52 = CST COFINS
                        f"{cofins_cst}|"
                        # campo 53 = Base COFINS
                        f"{cofins_base}|"
                        # campo 54 = Alíq COFINS
                        f"{cofins_aliq}|"
                        # campo 55 = Vlr COFINS
                        f"{cofins_valor}|"
                        # campos 56-62 = QtdLote, Enquad.IPI, Qtd(16,5), MovFísica, Unid, CodBico, VlrContábil
                        f"||{aliq_ipi}|{qtd}|S|{unid}||{vl_item}|"
                        # campos 63-68 = Qtd/Vlr/Vlr PIS e COFINS por unidade medida
                        f"0,000|0,0000|0,00|0,000|0,0000|0,00|"
                        # campo 69 = Nota devolvida | campo 70 = Descrição complementar
                        f"||"
                        # campo 71 = Natureza da receita PIS ← AQUI
                        f"{nat_pis}|"
                        # campo 72 = Natureza da receita COFINS ← AQUI
                        f"{nat_cofins}|"
                        # campo 73 = Exclusão cooperativa
                        f"0,00|"
                        # campos 74-75 = CST PIS/COFINS nota devolvida
                        f"||"
                        # campo 76 = Data | campos 77-78 = Vínculo Crédito PIS/COFINS
                        f"{dt_doc}|{nat_pis}|{nat_cofins}|"
                        # campos 79-90 = Bases/Impostos Frete/Seguro/Desp PIS e COFINS
                        f"0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|0,00|"
                        # campos 91-110 = Tanque, BC/Alíq/Vlr Carga Média, Redução,
                        #                  DIFAL, Cód.Rec.PIS/COFINS, IPI, CEST,
                        #                  PR-Benefício, Identificador, Desonerado, Código
                        f"||0,00|0,00|0,00|0,00||||||||||||"
                        # campos 111-118 = IBS e CBS
                        f"{campos_ibs_cbs}||||\n"
                    )

                stats['itens'] += 1

        except Exception as e:
            log.append(f"ERRO ao converter C100 NF={_c(campos_c100, SPED_C100_NUM_DOC)}: {e}")
            log.append(traceback.format_exc())
            stats['erros'] += 1

    # ── D100 ──────────────────────────────────────────────────────────────
    if 'D100' in parsed['por_tipo']:
        for campos, num_linha in parsed['por_tipo']['D100']:
            try:
                ind_oper = _c(campos, SPED_D100_IND_OPER)
                cod_part = _c(campos, SPED_D100_COD_PART)
                cod_mod  = _c(campos, SPED_D100_COD_MOD)
                cod_sit  = _c(campos, SPED_D100_COD_SIT)
                serie    = _c(campos, SPED_D100_SER)
                num_doc  = _c(campos, SPED_D100_NUM_DOC)
                dt_doc   = converter_data(_c(campos, SPED_D100_DT_DOC))
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
        f"Produtos={stats['produtos']} | "
        f"NFs entrada={stats['nf_entrada']} | NFs saída={stats['nf_saida']} | "
        f"Itens={stats['itens']} | "
        f"  ↳ Da Planilha Cliente={stats['itens_planilha']} | "
        f"  ↳ Do SPED (fallback)={stats['itens_sped']} | "
        f"Natureza aplicada={stats['natureza_aplicada']} itens | "
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

    tabela_cfop, tabela_flags    = carregar_tabela_cfop_oficial()
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

    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters  |  Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 📥 Entrada (SPED Fiscal)")
        st.markdown("- **0000** Abertura\n- **0150** Participantes\n- **0200** Produtos\n"
                    "- **C100** Notas Fiscais\n- **C170** Itens de NF\n- **C190** Analítico ICMS\n"
                    "- **D100** Conhecimento Transporte\n- **H010** Inventário\n")
        st.markdown("### 📤 Saída (Domínio Sistemas)")
        st.markdown("- **0000** Cabeçalho\n- **0100** Cadastro de Produtos\n"
                    "- **0110** Vigência Fiscal\n- **1000** Nota Fiscal\n"
                    "- **1020** Totais da NF\n- **1030** Itens de Entrada\n"
                    "- **2030** Itens de Saída\n- **9999** Encerramento\n")
        st.markdown("---")
        st.markdown("### 🏷️ Natureza PIS/COFINS")
        st.markdown(
            "CSTs que exigem Natureza:\n\n"
            "| CST | Descrição |\n|---|---|\n"
            "| 04 | Monofásica/Alíq.Zero |\n"
            "| 05 | Substituição Tributária |\n"
            "| 06 | Alíquota Zero |\n"
            "| 07 | Isenta |\n"
            "| 08 | Sem Incidência |\n"
            "| 09 | Suspensão |\n\n"
            "Configure na **Etapa 2** abaixo."
        )
        st.markdown("---")
        st.markdown("### 📊 Planilha Cliente")
        st.markdown(
            "PIS, COFINS, IBS e CBS lidos da Planilha Cliente.\n\n"
            "Colunas obrigatórias:\n"
            "`NF`, `COD.ITEM`, `PIS CST`, `PIS PC ALIQ`,\n"
            "`PIS VL BASE`, `PIS VALOR`, `COFINS CST`,\n"
            "`COFINS PC ALIQ`, `COFINS VL BASE`, `COFINS VALOR`"
        )
        st.markdown("---")
        st.markdown("### 📑 Fluxo")
        st.markdown(
            "1. Upload do SPED Fiscal `.txt`\n"
            "2. **Extrair CFOPs → baixar XLSX**\n"
            "3. Preencher coluna `ACUMULADOR`\n"
            "4. **Configurar Natureza PIS/COFINS** por CST\n"
            "5. Upload da Planilha Cliente (opcional)\n"
            "6. Upload do XLSX de acumuladores\n"
            "7. **Converter** e baixar saída\n"
        )

    st.markdown("---")

    # ── Defaults da session_state ─────────────────────────────────────────
    defaults = {
        "log":              [f"Aplicação pronta. Versão: {VERSAO} | CFOPs: {len(tabela_cfop)} | Impostos: {len(por_codigo_imp)}"],
        "resultado":        None,
        "nome_saida":       "saida_dominio.txt",
        "stats":            None,
        "xlsx_bytes":       None,
        "xlsx_nome":        "acumuladores.xlsx",
        "cfops_extraidos":  None,
        "tabela_acum_ok":   False,
        "arquivo_raw":      None,
        "arquivo_nome":     None,
        "df_cliente":       None,
        "cliente_ok":       False,
        "config_natureza":  {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Etapa 1: SPED ─────────────────────────────────────────────────────
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

    # ── Etapa 2: Configuração de Natureza PIS/COFINS ──────────────────────
    config_natureza = render_configuracao_natureza()

    st.markdown("---")

    # ── Etapa 3: Planilha Cliente ──────────────────────────────────────────
    st.markdown("### 📊 Etapa 3 — Upload da Planilha Cliente (PIS / COFINS / IBS / CBS)")

    arquivo_cliente = st.file_uploader(
        "📂 Planilha Cliente (.xlsx ou .csv)",
        type=["xlsx", "xls", "csv"],
        help="Planilha com colunas NF, COD.ITEM, PIS CST, PIS PC ALIQ, PIS VL BASE, PIS VALOR, "
             "COFINS CST, COFINS PC ALIQ, COFINS VL BASE, COFINS VALOR, CBS BASE, CBS ALIQ, "
             "CBS VALOR, CLASS TRIB CBS, IBS BASE, IBS ALIQ, IBS VALOR, CLASS TRIB IBS",
        key="upload_cliente",
    )

    if arquivo_cliente is not None:
        log_temp   = []
        raw_cli    = arquivo_cliente.read()
        df_cli_prev = carregar_planilha_cliente(raw_cli, arquivo_cliente.name, log_temp)
        if df_cli_prev is not None:
            st.session_state.df_cliente = df_cli_prev
            st.session_state.cliente_ok = True
            st.success(
                f"✅ Planilha Cliente válida — **{len(df_cli_prev)} linhas** | "
                f"NFs: {df_cli_prev['NF'].nunique()} | Itens únicos: {df_cli_prev['COD.ITEM'].nunique()}"
            )
            with st.expander("👁️ Prévia da Planilha Cliente (primeiras 10 linhas)"):
                st.dataframe(df_cli_prev.head(10), use_container_width=True, hide_index=True)
        else:
            st.session_state.df_cliente = None
            st.session_state.cliente_ok = False
            for msg in log_temp:
                st.error(msg)
    else:
        if not st.session_state.cliente_ok:
            st.info(
                "⬆ Faça o upload da Planilha Cliente para usar PIS/COFINS/IBS/CBS dela. "
                "Se não fizer upload, os valores do SPED serão usados como fallback."
            )

    st.markdown("---")

    # ── Etapa 4: Acumuladores + Conversão ─────────────────────────────────
    st.markdown("### ▶ Etapa 4 — Converter com a tabela de acumuladores preenchida")

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

            df_cliente_conv = st.session_state.get('df_cliente', None)
            cfg_nat_conv    = st.session_state.get('config_natureza', {})

            resultado_txt, stats = converter_sped_para_dominio(
                parsed, tabela_acum, tabela_flags, por_nome_imp,
                st.session_state.log,
                df_cliente=df_cliente_conv,
                config_natureza=cfg_nat_conv,
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
        col1.metric("Produtos",            stats.get('produtos',          0))
        col2.metric("NFs Entrada",         stats.get('nf_entrada',        0))
        col3.metric("NFs Saída",           stats.get('nf_saida',          0))
        col4.metric("Devoluções",          stats.get('devolucoes',         0))
        col5, col6, col7, col8 = st.columns(4)
        col5.metric("Itens (total)",       stats.get('itens',              0))
        col6.metric("Da Planilha Cliente", stats.get('itens_planilha',     0))
        col7.metric("Do SPED (fallback)",  stats.get('itens_sped',         0))
        col8.metric("Natureza Aplicada",   stats.get('natureza_aplicada',  0))
        col9, col10 = st.columns([1, 3])
        col9.metric("Erros", stats.get('erros', 0))
        st.markdown("---")
        with st.expander("👁️ Prévia do arquivo gerado (primeiras 80 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado.decode('latin-1', errors='replace').splitlines()[:80]
            )
            st.code(preview, language='text')
        st.download_button(
            label="⬇ Baixar Arquivo Domínio Sistemas",
            data=st.session_state.resultado,
            file_name=st.session_state.nome_saida,
            mime="text/plain", use_container_width=True, type="primary",
        )

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

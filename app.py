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
VERSAO = "V2.0"

# ==============================
# TEMA TR
# ==============================
def apply_tr_theme():
    st.markdown("""
        <style>
        html, body, [class*="css"] {
            font-family: 'Segoe UI', 'Arial', sans-serif;
            color: #444444;
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
    resultado     = []
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
# PARSER SPED FISCAL (ENTRADA)
# Layout EFD ICMS/IPI — separador pipe
# ==============================
def parse_sped(content: str) -> dict:
    """
    Lê o SPED Fiscal e agrupa linhas por tipo de registro,
    preservando a ordem original.
    Retorna:
        {
          'linhas_ordenadas': [(tipo, campos[], num_linha), ...],
          'por_tipo':         {tipo: [(campos[], num_linha), ...]}
        }
    """
    linhas_ordenadas = []
    por_tipo         = {}

    for num_linha, linha in enumerate(content.splitlines(), start=1):
        linha = linha.strip()
        if not linha:
            continue
        campos = linha.split('|')
        if campos and campos[0] == '':
            campos = campos[1:]
        if campos and campos[-1] == '':
            campos = campos[:-1]
        if not campos:
            continue
        tipo = campos[0].strip()
        if not tipo:
            continue
        linhas_ordenadas.append((tipo, campos, num_linha))
        por_tipo.setdefault(tipo, []).append((campos, num_linha))

    return {'linhas_ordenadas': linhas_ordenadas, 'por_tipo': por_tipo}


# ==============================
# HELPER
# ==============================
def _c(campos: list, idx: int, default: str = '') -> str:
    """Retorna campo pelo índice com fallback seguro."""
    return campos[idx].strip() if len(campos) > idx else default


# ==============================
# ÍNDICES SPED FISCAL (ENTRADA)
# Layout oficial EFD ICMS/IPI
# ==============================

# Registro 0000
# [0]=0000 [1]=COD_VER [2]=COD_FIN [3]=DT_INI [4]=DT_FIN
# [5]=NOME [6]=CNPJ [7]=CPF [8]=UF [9]=IE
SPED_0000_DT_INI = 3
SPED_0000_DT_FIN = 4
SPED_0000_NOME   = 5
SPED_0000_CNPJ   = 6
SPED_0000_UF     = 8

# Registro 0150 — Participantes
# [0]=0150 [1]=COD_PART [2]=NOME [3]=COD_PAIS [4]=CNPJ [5]=CPF
# [6]=IE [7]=COD_MUN [8]=SUFRAMA [9]=END [10]=NUM [11]=COMPL
# [12]=BAIRRO
SPED_0150_COD   = 1
SPED_0150_NOME  = 2
SPED_0150_CNPJ  = 4
SPED_0150_CPF   = 5
SPED_0150_IE    = 6
SPED_0150_END   = 9
SPED_0150_NUM   = 10
SPED_0150_COMPL = 11
SPED_0150_BAIRRO= 12

# Registro C100 — Nota Fiscal
# [0]=C100 [1]=IND_OPER [2]=IND_EMIT [3]=COD_PART [4]=COD_MOD
# [5]=COD_SIT [6]=SER [7]=NUM_DOC [8]=CHV_NFE [9]=DT_DOC
# [10]=DT_E_S [11]=VL_DOC [12]=IND_PGTO [13]=VL_DESC [14]=VL_ABAT_NT
# [15]=VL_MERC [16]=IND_FRT [17]=VL_FRT [18]=VL_SEG [19]=VL_OUT_DA
# [20]=VL_BC_ICMS [21]=VL_ICMS [22]=VL_BC_ICMS_ST [23]=VL_ICMS_ST
# [24]=VL_IPI [25]=VL_PIS [26]=VL_COFINS [27]=VL_PIS_ST [28]=VL_COFINS_ST
SPED_C100_IND_OPER  = 1
SPED_C100_COD_PART  = 3
SPED_C100_COD_MOD   = 4
SPED_C100_COD_SIT   = 5
SPED_C100_SER       = 6
SPED_C100_NUM_DOC   = 7
SPED_C100_CHV_NFE   = 8
SPED_C100_DT_DOC    = 9
SPED_C100_DT_ES     = 10
SPED_C100_VL_DOC    = 11
SPED_C100_VL_BC_ICMS= 20
SPED_C100_VL_ICMS   = 21
SPED_C100_VL_IPI    = 24
SPED_C100_VL_PIS    = 25
SPED_C100_VL_COFINS = 26

# Registro C170 — Itens da NF
# [0]=C170 [1]=NUM_ITEM [2]=COD_ITEM [3]=DESCR_COMPL [4]=QTD
# [5]=UNID [6]=VL_ITEM [7]=VL_DESC [8]=IND_MOV [9]=CST_ICMS
# [10]=CFOP [11]=COD_NAT [12]=VL_BC_ICMS [13]=ALIQ_ICMS [14]=VL_ICMS
# [15]=VL_BC_ICMS_ST [16]=ALIQ_ST [17]=VL_ICMS_ST [18]=IND_APUR
# [19]=CST_IPI [20]=COD_ENQ [21]=VL_BC_IPI [22]=ALIQ_IPI [23]=VL_IPI
# [24]=CST_PIS [25]=VL_BC_PIS [26]=ALIQ_PIS [27]=QUANT_BC_PIS [28]=VL_PIS
# [29]=CST_COFINS [30]=VL_BC_COFINS [31]=ALIQ_COFINS [32]=QUANT_BC_COFINS
# [33]=VL_COFINS [34]=COD_CTA [35]=VL_ABAT_NT
SPED_C170_NUM_ITEM  = 1
SPED_C170_COD_ITEM  = 2
SPED_C170_DESCR     = 3
SPED_C170_QTD       = 4
SPED_C170_UNID      = 5
SPED_C170_VL_ITEM   = 6
SPED_C170_VL_DESC   = 7
SPED_C170_CFOP      = 10
SPED_C170_VL_BC     = 12
SPED_C170_ALIQ_ICMS = 13
SPED_C170_VL_ICMS   = 14

# Registro C190 — Analítico ICMS
# [0]=C190 [1]=CST_ICMS [2]=CFOP [3]=ALIQ_ICMS [4]=VL_OPR
# [5]=VL_BC_ICMS [6]=VL_ICMS [7]=VL_BC_ICMS_ST [8]=VL_ICMS_ST
# [9]=VL_RED_BC [10]=COD_OBS
SPED_C190_CST_ICMS  = 1
SPED_C190_CFOP      = 2
SPED_C190_ALIQ      = 3
SPED_C190_VL_OPR    = 4
SPED_C190_VL_BC     = 5
SPED_C190_VL_ICMS   = 6
SPED_C190_VL_RED    = 9

# Registro D100 — Conhecimento de Transporte
# [0]=D100 [1]=IND_OPER [2]=IND_EMIT [3]=COD_PART [4]=COD_MOD
# [5]=COD_SIT [6]=SER [7]=SUB [8]=NUM_DOC [9]=CHV_CTE
# [10]=DT_DOC [11]=DT_A_P [12]=TP_CT [13]=CHV_DOC_ANT
# [14]=VL_DOC [15]=VL_DESC [16]=IND_FRT [17]=VL_SERV
# [18]=VL_BC_ICMS [19]=ALIQ_ICMS [20]=VL_ICMS [21]=VL_NT
SPED_D100_IND_OPER = 1
SPED_D100_COD_PART = 3
SPED_D100_COD_MOD  = 4
SPED_D100_COD_SIT  = 5
SPED_D100_SER      = 6
SPED_D100_NUM_DOC  = 8
SPED_D100_DT_DOC   = 10
SPED_D100_VL_DOC   = 14
SPED_D100_VL_BC    = 18
SPED_D100_ALIQ     = 19
SPED_D100_VL_ICMS  = 20

# Registro H010 — Inventário
# [0]=H010 [1]=COD_ITEM [2]=UNID [3]=QTD [4]=VL_UNIT [5]=VL_ITEM [6]=IND_PROP
SPED_H010_COD_ITEM = 1
SPED_H010_UNID     = 2
SPED_H010_QTD      = 3
SPED_H010_VL_UNIT  = 4
SPED_H010_VL_ITEM  = 5


# ==============================
# EXTRAÇÃO DE CFOPs DO SPED
# ==============================
def extrair_cfops_do_sped(parsed: dict, log: list) -> dict:
    """
    Extrai todos os CFOPs únicos presentes nos registros
    C170 (itens NF) e C190 (analítico) do SPED Fiscal.
    """
    cfops = {}

    def registrar(cfop_raw: str, tipo_reg: str):
        cfop = str(cfop_raw).strip()
        if not cfop or not cfop.isdigit():
            return
        cfop = cfop.zfill(4)
        if cfop == '0000':
            return
        if cfop not in cfops:
            primeiro = cfop[0]
            cfops[cfop] = {
                'registros':     set(),
                'ocorrencias':   0,
                'tipo_operacao': 'Entrada' if primeiro in ('1','2','3')
                                 else ('Saída' if primeiro in ('5','6','7')
                                       else 'Desconhecido'),
            }
        cfops[cfop]['registros'].add(tipo_reg)
        cfops[cfop]['ocorrencias'] += 1

    total = {'C100': 0, 'C170': 0, 'C190': 0, 'D100': 0}

    for tipo, campos, _ in parsed['linhas_ordenadas']:
        if tipo == 'C100':
            total['C100'] += 1
        elif tipo == 'C170':
            total['C170'] += 1
            registrar(_c(campos, SPED_C170_CFOP), 'C170')
        elif tipo == 'C190':
            total['C190'] += 1
            registrar(_c(campos, SPED_C190_CFOP), 'C190')
        elif tipo == 'D100':
            total['D100'] += 1

    log.append(
        f"Registros lidos: C100={total['C100']} | "
        f"C170={total['C170']} | C190={total['C190']} | D100={total['D100']}"
    )
    log.append(
        f"CFOPs únicos encontrados: {len(cfops)} — {sorted(cfops.keys())}"
    )
    return cfops


# ==============================
# DESCRIÇÕES DE CFOP
# ==============================
DESCRICOES_CFOP = {
    '1101': 'Compra p/ industrialização - dentro do estado',
    '1102': 'Compra p/ comercialização - dentro do estado',
    '1111': 'Compra p/ industrialização em zona franca',
    '1113': 'Compra p/ comercialização em zona franca',
    '1116': 'Compra p/ industrialização - devol. de remessa',
    '1120': 'Compra p/ industrialização sem transitar pelo estab.',
    '1201': 'Devolução de venda - dentro do estado',
    '1202': 'Devolução de venda - dentro do estado - tributada',
    '1203': 'Devolução de venda - dentro do estado - não tributada',
    '1251': 'Compra p/ ativo imobilizado - dentro do estado',
    '1252': 'Compra p/ ativo imobilizado - fora do estado',
    '1301': 'Aquisição de serviço de transporte - dentro do estado',
    '1351': 'Aquisição de serviço de comunicação - dentro do estado',
    '1352': 'Aquisição de serviço de comunicação - fora do estado',
    '1401': 'Compra p/ uso e consumo - dentro do estado',
    '1403': 'Compra p/ uso e consumo - fora do estado',
    '1501': 'Entrada de mercadoria c/ fim específico de exportação',
    '1601': 'Recebimento, por transferência, de crédito de ICMS',
    '1701': 'Entrada de mercadoria recebida em consignação mercantil',
    '1801': 'Entradas p/ industrialização por encomenda',
    '1901': 'Outras entradas de mercadorias não especificadas',
    '2101': 'Compra p/ industrialização - fora do estado',
    '2102': 'Compra p/ comercialização - fora do estado',
    '2111': 'Compra p/ industrialização em zona franca - fora do estado',
    '2113': 'Compra p/ comercialização em zona franca - fora do estado',
    '2201': 'Devolução de venda - fora do estado',
    '2251': 'Compra p/ ativo imobilizado - fora do estado',
    '2301': 'Aquisição de serviço de transporte - fora do estado',
    '2401': 'Compra p/ uso e consumo - fora do estado',
    '2901': 'Outras entradas não especificadas - fora do estado',
    '3101': 'Compra p/ industrialização - exterior',
    '3102': 'Compra p/ comercialização - exterior',
    '5101': 'Venda de produção do estabelecimento - dentro do estado',
    '5102': 'Venda de mercadoria adquirida ou recebida - dentro do estado',
    '5111': 'Venda de produção - zona franca - dentro do estado',
    '5113': 'Venda de mercadoria - zona franca - dentro do estado',
    '5151': 'Transferência de produção - dentro do estado',
    '5152': 'Transferência de mercadoria - dentro do estado',
    '5201': 'Devolução de compra p/ industrialização - dentro do estado',
    '5202': 'Devolução de compra p/ comercialização - dentro do estado',
    '5251': 'Venda de ativo imobilizado - dentro do estado',
    '5301': 'Prestação de serviço de transporte - dentro do estado',
    '5351': 'Prestação de serviço de comunicação - dentro do estado',
    '5401': 'Venda de produção - subst. tributária - dentro do estado',
    '5403': 'Venda de mercadoria - subst. tributária - dentro do estado',
    '5501': 'Remessa de mercadoria p/ formação de lote de exportação',
    '5601': 'Transferência de crédito de ICMS - dentro do estado',
    '5701': 'Venda de mercadoria em consignação - dentro do estado',
    '5801': 'Remessa p/ industrialização por encomenda - dentro do estado',
    '5901': 'Outras saídas de mercadorias não especificadas',
    '6101': 'Venda de produção do estabelecimento - fora do estado',
    '6102': 'Venda de mercadoria adquirida ou recebida - fora do estado',
    '6151': 'Transferência de produção - fora do estado',
    '6152': 'Transferência de mercadoria - fora do estado',
    '6201': 'Devolução de compra p/ industrialização - fora do estado',
    '6202': 'Devolução de compra p/ comercialização - fora do estado',
    '6251': 'Venda de ativo imobilizado - fora do estado',
    '6301': 'Prestação de serviço de transporte - fora do estado',
    '6401': 'Venda de produção - subst. tributária - fora do estado',
    '6403': 'Venda de mercadoria - subst. tributária - fora do estado',
    '6501': 'Remessa p/ formação de lote de exportação - fora do estado',
    '6601': 'Transferência de crédito de ICMS - fora do estado',
    '6701': 'Venda de mercadoria em consignação - fora do estado',
    '6801': 'Remessa p/ industrialização por encomenda - fora do estado',
    '6901': 'Outras saídas não especificadas - fora do estado',
}


# ==============================
# GERADOR XLSX DE ACUMULADORES — TEMA TR
# ==============================
def gerar_xlsx_acumuladores_tr(cfops_dict: dict) -> bytes:
    wb = Workbook()
    COR_LARANJA   = "FF8000"
    COR_CINZA_ESC = "444444"
    COR_CINZA_CLR = "E9E9E9"
    COR_BRANCO    = "FFFFFF"
    COR_LARANJA_C = "FFF3E0"
    COR_VERDE_CLR = "E8F5E9"
    COR_VERM_CLR  = "FFEBEE"

    borda_fina = Border(
        left=Side(style='thin',   color="CCCCCC"),
        right=Side(style='thin',  color="CCCCCC"),
        top=Side(style='thin',    color="CCCCCC"),
        bottom=Side(style='thin', color="CCCCCC"),
    )

    def fill(h):
        return PatternFill("solid", fgColor=h)

    def center():
        return Alignment(horizontal='center', vertical='center')

    def left_al():
        return Alignment(horizontal='left', vertical='center')

    cfops_ord = sorted(
        cfops_dict.items(),
        key=lambda x: (0 if x[1]['tipo_operacao'] == 'Entrada' else 1, x[0])
    )

    # ── Aba 1: Acumuladores ──────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Acumuladores"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:E1")
    ws1.row_dimensions[1].height = 36
    c = ws1["A1"]
    c.value     = "Thomson Reuters  |  Domínio Sistemas  —  Tabela de Acumuladores CFOP"
    c.fill      = fill(COR_CINZA_ESC)
    c.font      = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA)
    c.alignment = left_al()

    ws1.merge_cells("A2:E2")
    ws1.row_dimensions[2].height = 20
    c2 = ws1["A2"]
    c2.value = (
        f"Extraído do SPED Fiscal  |  {len(cfops_dict)} CFOP(s) identificado(s)  |  "
        f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    c2.fill      = fill(COR_LARANJA)
    c2.font      = Font(name='Segoe UI', size=9, color=COR_BRANCO)
    c2.alignment = left_al()

    ws1.merge_cells("A3:E3")
    ws1.row_dimensions[3].height = 18
    c3 = ws1["A3"]
    c3.value     = (
        "⚠  Preencha a coluna ACUMULADOR para cada CFOP "
        "antes de fazer o upload no conversor."
    )
    c3.fill      = fill(COR_CINZA_CLR)
    c3.font      = Font(name='Segoe UI', bold=True, size=9, color=COR_CINZA_ESC)
    c3.alignment = left_al()

    ws1.row_dimensions[4].height = 6
    ws1.row_dimensions[5].height = 22

    cabecalhos = ['CFOP', 'DESCRIÇÃO', 'TIPO OPERAÇÃO', 'OCORRÊNCIAS', 'ACUMULADOR']
    col_widths  = [10,     48,          16,               14,            16]
    for ci, (cab, w) in enumerate(zip(cabecalhos, col_widths), start=1):
        ws1.column_dimensions[get_column_letter(ci)].width = w
        cell           = ws1.cell(row=5, column=ci, value=cab)
        cell.fill      = fill(COR_LARANJA)
        cell.font      = Font(name='Segoe UI', bold=True, size=11, color=COR_BRANCO)
        cell.alignment = center()
        cell.border    = borda_fina

    linha = 6
    for idx, (cfop, info) in enumerate(cfops_ord):
        ws1.row_dimensions[linha].height = 18
        bg = (COR_CINZA_CLR if idx % 2 == 0 else COR_BRANCO) \
             if info['tipo_operacao'] == 'Entrada' \
             else (COR_LARANJA_C if idx % 2 == 0 else COR_BRANCO)
        descricao = DESCRICOES_CFOP.get(cfop, '— preencha a descrição —')
        valores   = [cfop, descricao, info['tipo_operacao'], info['ocorrencias'], '']
        for ci, valor in enumerate(valores, start=1):
            cell        = ws1.cell(row=linha, column=ci, value=valor)
            cell.border = borda_fina
            if ci == 1:
                cell.fill = fill(bg)
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 3:
                cell.fill = fill(bg)
                cor = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor)
                cell.alignment = center()
            elif ci == 4:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 5:
                cell.fill      = fill("FFF8F0")
                cell.font      = Font(name='Segoe UI', bold=True, size=10, color=COR_LARANJA)
                cell.alignment = center()
                cell.border    = Border(
                    left=Side(style='medium', color=COR_LARANJA),
                    right=Side(style='medium', color=COR_LARANJA),
                    top=Side(style='thin',    color="CCCCCC"),
                    bottom=Side(style='thin', color="CCCCCC"),
                )
            else:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC)
                cell.alignment = left_al()
        linha += 1

    ws1.merge_cells(f"A{linha}:E{linha}")
    ws1.row_dimensions[linha].height = 18
    cr = ws1.cell(row=linha, column=1,
                  value="Thomson Reuters  |  Domínio Sistemas  |  Gerado automaticamente")
    cr.fill      = fill(COR_CINZA_ESC)
    cr.font      = Font(name='Segoe UI', size=8, color="888888")
    cr.alignment = Alignment(horizontal='right', vertical='center')
    ws1.freeze_panes    = "A6"
    ws1.auto_filter.ref = f"A5:E{linha - 1}"

    # ── Aba 2: CFOPs Encontrados ─────────────────────────────────────────
    ws2 = wb.create_sheet(title="CFOPs Encontrados")
    ws2.sheet_view.showGridLines = False
    n_ent = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Entrada')
    n_sai = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Saída')

    ws2.merge_cells("A1:F1")
    ws2.row_dimensions[1].height = 36
    c = ws2["A1"]
    c.value     = "Thomson Reuters  |  Domínio Sistemas  —  CFOPs Identificados no SPED Fiscal"
    c.fill      = fill(COR_CINZA_ESC)
    c.font      = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA)
    c.alignment = left_al()

    ws2.merge_cells("A2:F2")
    ws2.row_dimensions[2].height = 20
    c2 = ws2["A2"]
    c2.value = (
        f"Analisado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
        f"Total: {len(cfops_dict)}  |  Entradas: {n_ent}  |  Saídas: {n_sai}"
    )
    c2.fill      = fill(COR_LARANJA)
    c2.font      = Font(name='Segoe UI', size=9, color=COR_BRANCO)
    c2.alignment = left_al()

    ws2.row_dimensions[3].height = 6
    ws2.row_dimensions[4].height = 22
    cab2  = ['CFOP', 'DESCRIÇÃO', 'TIPO OPERAÇÃO', 'REGISTROS SPED', 'OCORRÊNCIAS', 'STATUS']
    wids2 = [10,     48,          16,               18,                14,            22]
    for ci, (cab, w) in enumerate(zip(cab2, wids2), start=1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
        cell           = ws2.cell(row=4, column=ci, value=cab)
        cell.fill      = fill(COR_LARANJA)
        cell.font      = Font(name='Segoe UI', bold=True, size=11, color=COR_BRANCO)
        cell.alignment = center()
        cell.border    = borda_fina

    linha2 = 5
    for idx, (cfop, info) in enumerate(cfops_ord):
        ws2.row_dimensions[linha2].height = 18
        bg      = COR_CINZA_CLR if idx % 2 == 0 else COR_BRANCO
        mapeado = cfop in DESCRICOES_CFOP
        status  = '✔ Catalogado' if mapeado else '✘ Não catalogado'
        bg_st   = COR_VERDE_CLR if mapeado else COR_VERM_CLR
        descr   = DESCRICOES_CFOP.get(cfop, '— CFOP não catalogado —')
        regs    = ', '.join(sorted(info['registros']))
        vals    = [cfop, descr, info['tipo_operacao'], regs, info['ocorrencias'], status]
        for ci, valor in enumerate(vals, start=1):
            cell        = ws2.cell(row=linha2, column=ci, value=valor)
            cell.border = borda_fina
            if ci == 6:
                cell.fill = fill(bg_st)
                cor = "1B5E20" if mapeado else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor)
                cell.alignment = center()
            elif ci == 1:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 3:
                cell.fill = fill(bg)
                cor = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor)
                cell.alignment = center()
            elif ci in (4, 5):
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            else:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC)
                cell.alignment = left_al()
        linha2 += 1

    ws2.merge_cells(f"A{linha2}:F{linha2}")
    ws2.row_dimensions[linha2].height = 18
    cr2 = ws2.cell(row=linha2, column=1,
                   value="Thomson Reuters  |  Domínio Sistemas  |  Gerado automaticamente")
    cr2.fill      = fill(COR_CINZA_ESC)
    cr2.font      = Font(name='Segoe UI', size=8, color="888888")
    cr2.alignment = Alignment(horizontal='right', vertical='center')
    ws2.freeze_panes    = "A5"
    ws2.auto_filter.ref = f"A4:F{linha2 - 1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ==============================
# CARREGAMENTO DA TABELA DE ACUMULADORES
# ==============================
def carregar_acumuladores(arquivo_bytes: bytes, nome_arquivo: str, log: list) -> dict | None:
    try:
        ext = os.path.splitext(nome_arquivo)[1].lower()
        if ext in ('.xlsx', '.xls'):
            df = pd.read_excel(io.BytesIO(arquivo_bytes), dtype=str)
        else:
            raw_str = arquivo_bytes.decode('latin-1', errors='replace')
            sep     = ';' if raw_str.count(';') >= raw_str.count(',') else ','
            df      = pd.read_csv(io.StringIO(raw_str), sep=sep, dtype=str)

        df.columns = [str(c).strip().upper() for c in df.columns]
        if 'CFOP' not in df.columns or 'ACUMULADOR' not in df.columns:
            log.append("ERRO: O arquivo deve conter as colunas 'CFOP' e 'ACUMULADOR'.")
            return None

        tabela = {}
        for _, row in df.iterrows():
            cfop = str(row['CFOP']).strip().zfill(4)
            acum = str(row['ACUMULADOR']).strip()
            if not cfop or not acum:
                continue
            if cfop.upper() == 'NAN' or acum.upper() == 'NAN':
                continue
            if acum in ('', '0', 'nan', 'NAN', 'None'):
                continue
            tabela[cfop] = acum

        if not tabela:
            log.append(
                "ERRO: Nenhum par CFOP → Acumulador preenchido. "
                "Preencha a coluna ACUMULADOR e tente novamente."
            )
            return None

        log.append(f"Tabela de acumuladores carregada: {len(tabela)} CFOPs mapeados.")
        return tabela

    except Exception as e:
        log.append(f"ERRO ao carregar tabela de acumuladores: {e}")
        return None


def get_acumulador(cfop: str, tabela: dict, nao_mapeados: set) -> str:
    cfop_norm = str(cfop).strip().zfill(4)
    acum      = tabela.get(cfop_norm)
    if acum is None:
        nao_mapeados.add(cfop_norm)
        return '9999'
    return acum


# ==============================
# CONVERSÃO SPED → DOMÍNIO
#
# SPED (entrada)          DOMÍNIO (saída)
# ─────────────────────── ──────────────────────────────
# 0000                 →  0000  (cabeçalho)
# C100 + C170 + C190   →  1000 (NF) + 1020 (totais)
#                          + 1030 (itens) por C100
# D100                 →  1000 (CT-e) + 1020
# H010                 →  registro de inventário (passthrough)
# —                    →  9999  (encerramento)
# ==============================

def converter_sped_para_dominio(parsed: dict, tabela_acum: dict, log: list) -> tuple:
    """
    Converte o SPED Fiscal para o leiaute Domínio Sistemas.
    Retorna (conteudo_str, stats_dict)
    """
    saida        = StringIO()
    nao_mapeados = set()
    stats = {
        'nf_entrada':  0,
        'nf_saida':    0,
        'itens':       0,
        'analiticos':  0,
        'transporte':  0,
        'inventario':  0,
        'erros':       0,
    }

    # ── 0000 → 0000 ──────────────────────────────────────────────────────
    if '0000' in parsed['por_tipo']:
        campos, _ = parsed['por_tipo']['0000'][0]
        cnpj  = _c(campos, SPED_0000_CNPJ)
        saida.write(f"|0000|{cnpj}|\n")
        log.append(f"0000: CNPJ={cnpj}")
    else:
        log.append("AVISO: Registro 0000 não encontrado no SPED.")

    # ── Monta dicionário de participantes (0150) para lookup ──────────────
    participantes = {}
    if '0150' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['0150']:
            cod = _c(campos, SPED_0150_COD)
            participantes[cod] = campos

    # ── Monta hierarquia C100 → C170/C190 ────────────────────────────────
    # Percorre as linhas em ordem para associar filhos ao pai C100
    blocos_c = []       # lista de {'c100': campos, 'c170': [campos,...], 'c190': [campos,...]}
    bloco_atual = None

    for tipo, campos, num_linha in parsed['linhas_ordenadas']:
        if tipo == 'C100':
            if bloco_atual is not None:
                blocos_c.append(bloco_atual)
            bloco_atual = {'c100': campos, 'c170': [], 'c190': []}
        elif tipo == 'C170':
            if bloco_atual is not None:
                bloco_atual['c170'].append(campos)
            else:
                log.append(f"AVISO: C170 na linha {num_linha} sem C100 pai. Ignorado.")
        elif tipo == 'C190':
            if bloco_atual is not None:
                bloco_atual['c190'].append(campos)
            else:
                log.append(f"AVISO: C190 na linha {num_linha} sem C100 pai. Ignorado.")

    if bloco_atual is not None:
        blocos_c.append(bloco_atual)

    log.append(f"Blocos C100 montados: {len(blocos_c)}")

    # ── Converte cada bloco C100 → 1000 + 1020 + 1030(s) ────────────────
    for bloco in blocos_c:
        campos_c100 = bloco['c100']
        try:
            ind_oper   = _c(campos_c100, SPED_C100_IND_OPER)
            cod_part   = _c(campos_c100, SPED_C100_COD_PART)
            cod_mod    = _c(campos_c100, SPED_C100_COD_MOD)
            cod_sit    = _c(campos_c100, SPED_C100_COD_SIT)
            serie      = _c(campos_c100, SPED_C100_SER)
            num_doc    = _c(campos_c100, SPED_C100_NUM_DOC)
            chv_nfe    = _c(campos_c100, SPED_C100_CHV_NFE)
            dt_doc     = _c(campos_c100, SPED_C100_DT_DOC)
            dt_es      = _c(campos_c100, SPED_C100_DT_ES)
            vl_doc     = _c(campos_c100, SPED_C100_VL_DOC)
            vl_bc_icms = _c(campos_c100, SPED_C100_VL_BC_ICMS)
            vl_icms    = _c(campos_c100, SPED_C100_VL_ICMS)
            vl_ipi     = _c(campos_c100, SPED_C100_VL_IPI)
            vl_pis     = _c(campos_c100, SPED_C100_VL_PIS)
            vl_cofins  = _c(campos_c100, SPED_C100_VL_COFINS)

            # Busca CNPJ do participante
            part_campos = participantes.get(cod_part, [])
            cnpj_part   = _c(part_campos, SPED_0150_CNPJ) if part_campos else ''

            # Determina tipo E/S pelo IND_OPER
            # 0=Entrada, 1=Saída
            tipo_es = 'E' if ind_oper == '0' else 'S'

            # CFOP principal — pega do primeiro C170 ou C190
            cfop_principal = ''
            if bloco['c170']:
                cfop_principal = _c(bloco['c170'][0], SPED_C170_CFOP)
            elif bloco['c190']:
                cfop_principal = _c(bloco['c190'][0], SPED_C190_CFOP)

            acum_principal = get_acumulador(cfop_principal, tabela_acum, nao_mapeados)

            # ── Registro 1000 (cabeçalho da NF no Domínio) ───────────────
            # Layout Domínio 1000 (baseado no arquivo de exemplo):
            # |1000|NUM_NF|CNPJ_EMIT|CNPJ_DEST|IND_OPER|CFOP|SERIE|
            # COD_MOD|COD_SIT|CHV_NFE|...|DT_DOC|DT_ES|VL_DOC|OBS|
            # ...|VL_BC_ICMS|VL_ICMS|VL_IPI|VL_PIS|VL_COFINS|...
            saida.write(
                f"|1000|{num_doc}|{cnpj_part}||{ind_oper}|{cfop_principal}|"
                f"{serie}|{cod_mod}|{cod_sit}|{chv_nfe}|||"
                f"{dt_doc}|{dt_es}|{vl_doc}||||||||||{tipo_es}|"
                f"0,00|0,00|0,00|0,00||0,00||||0,00|0,00|0,00||{vl_doc}|"
                f"0|0||||{acum_principal}||0,00||||||N|S||{tipo_es}||0|||||"
                f"||||||||||||0|{cod_sit}|0||0,00|0,00|0,00|||||||||||\n"
            )

            if ind_oper == '0':
                stats['nf_entrada'] += 1
            else:
                stats['nf_saida'] += 1

            # ── Registro 1020 (totais da NF no Domínio) ──────────────────
            saida.write(
                f"|1020|{num_doc}||{vl_doc}|0,00|{vl_icms}|"
                f"0,00|0,00|0,00|0,00|{vl_doc}||||\n"
            )

            # ── Registros 1030 (itens da NF no Domínio) ──────────────────
            for campos_c170 in bloco['c170']:
                num_item  = _c(campos_c170, SPED_C170_NUM_ITEM)
                qtd       = _c(campos_c170, SPED_C170_QTD)
                vl_unit   = ''   # SPED não tem VL_UNIT direto no C170
                dt_item   = dt_doc
                cod_sit_i = cod_sit
                vl_item   = _c(campos_c170, SPED_C170_VL_ITEM)
                vl_desc_i = _c(campos_c170, SPED_C170_VL_DESC)
                cfop_item = _c(campos_c170, SPED_C170_CFOP)
                vl_bc_i   = _c(campos_c170, SPED_C170_VL_BC)
                aliq_i    = _c(campos_c170, SPED_C170_ALIQ_ICMS)
                vl_icms_i = _c(campos_c170, SPED_C170_VL_ICMS)
                acum_item = get_acumulador(cfop_item, tabela_acum, nao_mapeados)

                # Calcula VL_UNIT a partir de VL_ITEM / QTD
                try:
                    vl_unit = f"{float(vl_item.replace(',', '.')) / float(qtd.replace(',', '.')):.3f}".replace('.', ',')
                except Exception:
                    vl_unit = vl_item

                # Layout Domínio 1030 (baseado no arquivo de exemplo):
                # |1030|NUM_ITEM|QTD|VL_UNIT|0|0|1|DT_DOC||COD_SIT|
                # VL_ITEM|VL_DESC|VL_ITEM|0,00|ALIQ_ICMS|||0,00|0,00|0,00|
                # 0,000|VL_ICMS|0,00|0,00|0,00||VL_ITEM|0,00|02|0,00|0,00|
                # 0,00|0,00|CFOP||0,0000|0,00|0,00|0,00|VL_BC|70|0,00|70|
                # 0,00|||||DT_DOC|DT_DOC||||||S|UN|||VL_ITEM|||||||1|||||
                # 01|01||||||||
                saida.write(
                    f"|1030|{num_item}|{qtd}|{vl_unit}|0|0|1|{dt_item}||"
                    f"{cod_sit_i}|{vl_item}|{vl_desc_i}|{vl_item}|0,00|"
                    f"{aliq_i}|||0,00|0,00|0,00|0,000|{vl_icms_i}|0,00|"
                    f"0,00|0,00||{vl_item}|0,00|02|0,00|0,00|0,00|0,00|"
                    f"{cfop_item}||0,0000|0,00|0,00|0,00|{vl_bc_i}|70|"
                    f"0,00|70|0,00|||||{dt_item}|{dt_item}||||||S|UN|||"
                    f"{vl_item}|||||||1|||||01|01||||||||\n"
                )
                stats['itens'] += 1

        except Exception as e:
            log.append(f"ERRO ao converter C100 num_doc={_c(campos_c100, SPED_C100_NUM_DOC)}: {e}")
            stats['erros'] += 1

    # ── D100 → 1000 + 1020 (Conhecimento de Transporte) ─────────────────
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
                vl_bc    = _c(campos, SPED_D100_VL_BC)
                aliq     = _c(campos, SPED_D100_ALIQ)
                vl_icms  = _c(campos, SPED_D100_VL_ICMS)
                tipo_es  = 'E' if ind_oper == '0' else 'S'

                part_campos = participantes.get(cod_part, [])
                cnpj_part   = _c(part_campos, SPED_0150_CNPJ) if part_campos else ''

                saida.write(
                    f"|1000|{num_doc}|{cnpj_part}||{ind_oper}||"
                    f"{serie}|{cod_mod}|{cod_sit}||||"
                    f"{dt_doc}|{dt_doc}|{vl_doc}||||||||||{tipo_es}|"
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

    # ── H010 → passthrough como comentário (inventário não tem reg. Domínio direto) ──
    if 'H010' in parsed['por_tipo']:
        for campos, _ in parsed['por_tipo']['H010']:
            cod_item = _c(campos, SPED_H010_COD_ITEM)
            unid     = _c(campos, SPED_H010_UNID)
            qtd      = _c(campos, SPED_H010_QTD)
            vl_unit  = _c(campos, SPED_H010_VL_UNIT)
            vl_item  = _c(campos, SPED_H010_VL_ITEM)
            saida.write(
                f"|H010|{cod_item}|{unid}|{qtd}|{vl_unit}|{vl_item}|\n"
            )
            stats['inventario'] += 1

    # ── 9999 ──────────────────────────────────────────────────────────────
    saida.write("|9999|\n")

    # Avisos de CFOPs não mapeados
    if nao_mapeados:
        log.append(
            f"AVISO: {len(nao_mapeados)} CFOP(s) sem acumulador (9999): "
            f"{', '.join(sorted(nao_mapeados))}"
        )

    log.append(
        f"Conversão concluída — "
        f"NFs entrada={stats['nf_entrada']} | NFs saída={stats['nf_saida']} | "
        f"Itens={stats['itens']} | Transporte={stats['transporte']} | "
        f"Inventário={stats['inventario']} | Erros={stats['erros']}"
    )

    return saida.getvalue(), stats


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="SPED Fiscal → Domínio | Thomson Reuters",
        page_icon="🟠",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_tr_theme()

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
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ──────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters**")
        st.markdown("**Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 📥 Entrada (SPED Fiscal)")
        st.markdown(
            "- **0000** Abertura\n"
            "- **0150** Participantes\n"
            "- **C100** Notas Fiscais\n"
            "- **C170** Itens de NF\n"
            "- **C190** Analítico ICMS\n"
            "- **D100** Conhecimento de Transporte\n"
            "- **H010** Inventário\n"
        )
        st.markdown("### 📤 Saída (Domínio Sistemas)")
        st.markdown(
            "- **0000** Cabeçalho\n"
            "- **1000** Nota Fiscal\n"
            "- **1020** Totais da NF\n"
            "- **1030** Itens da NF\n"
            "- **9999** Encerramento\n"
        )
        st.markdown("---")
        st.markdown("### 📑 Fluxo")
        st.markdown(
            "1. Upload do SPED Fiscal `.txt`\n"
            "2. **Extrair CFOPs → baixar XLSX**\n"
            "3. Preencher coluna `ACUMULADOR`\n"
            "4. Upload do XLSX preenchido\n"
            "5. **Converter** e baixar saída\n"
        )
        st.markdown("---")
        st.markdown("### ⚙ Encoding de saída")
        st.markdown("**ANSI (Latin-1)**")

    # ── Instruções ────────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">
            <h4>🔹 Etapa 1 — Upload do SPED Fiscal e extração de CFOPs</h4>
            <p>Faça o upload do arquivo <code>.txt</code> do SPED Fiscal (EFD ICMS/IPI)
            e clique em <b>🔍 Extrair CFOPs e Gerar Planilha</b>. O sistema identificará
            todos os CFOPs presentes e gerará um XLSX no tema Thomson Reuters.</p>

            <h4>🔹 Etapa 2 — Preencher acumuladores</h4>
            <p>Abra o XLSX, preencha a coluna <b>ACUMULADOR</b> para cada CFOP
            conforme a parametrização do Domínio Sistemas e salve.</p>

            <h4>🔹 Etapa 3 — Converter</h4>
            <p>Faça o upload do XLSX preenchido e clique em
            <b>▶ Converter SPED → Domínio</b>. O arquivo gerado estará no
            leiaute padrão do Domínio Sistemas pronto para importação.</p>

            <hr>
            <h4>⚠ Observações</h4>
            <ul>
                <li>Entrada: SPED Fiscal EFD ICMS/IPI com registros C100/C170/C190/D100.</li>
                <li>Saída: leiaute Domínio Sistemas com registros 1000/1020/1030.</li>
                <li>CFOPs sem acumulador preenchido receberão <b>9999</b>.</li>
                <li>Saída em <b>ANSI (Latin-1)</b>, padrão do Domínio Sistemas.</li>
            </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Session state ─────────────────────────────────────────────────────
    defaults = {
        "log":             [f"Aplicação pronta. Versão: {VERSAO}"],
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
    # ETAPA 1 — UPLOAD DO SPED + EXTRAÇÃO DE CFOPs
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### 🔍 Etapa 1 — Upload do SPED Fiscal e extração de CFOPs")

    uploaded_file = st.file_uploader(
        "📂 Arquivo SPED Fiscal (.txt)",
        type=["txt"],
        help="Arquivo EFD ICMS/IPI exportado pelo sistema ERP",
        key="upload_sped",
    )

    # Cacheia os bytes imediatamente para evitar perda após st.rerun()
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
                f"Arquivo carregado: {uploaded_file.name} "
                f"({len(raw_atual)/1024:.1f} KB)"
            ]

    if st.session_state.arquivo_raw is not None:
        st.info(
            f"📄 Arquivo em memória: **{st.session_state.arquivo_nome}** "
            f"({len(st.session_state.arquivo_raw)/1024:.1f} KB)"
        )

    col_e1, col_e2 = st.columns([1, 1])
    with col_e1:
        extrair = st.button(
            "🔍 Extrair CFOPs e Gerar Planilha",
            disabled=(st.session_state.arquivo_raw is None),
            use_container_width=True,
            type="primary",
        )
    with col_e2:
        if st.session_state.xlsx_bytes is not None:
            st.download_button(
                label="⬇ Baixar Planilha de Acumuladores (.xlsx)",
                data=st.session_state.xlsx_bytes,
                file_name=st.session_state.xlsx_nome,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )

    if extrair:
        st.session_state.log             = ["Extraindo CFOPs do SPED Fiscal..."]
        st.session_state.xlsx_bytes      = None
        st.session_state.cfops_extraidos = None

        try:
            content    = decode_arquivo(st.session_state.arquivo_raw)
            parsed     = parse_sped(content)

            tipos_enc = list(parsed['por_tipo'].keys())
            st.session_state.log.append(
                f"Registros encontrados: {', '.join(tipos_enc)}"
            )

            cfops_dict = extrair_cfops_do_sped(parsed, st.session_state.log)

            if not cfops_dict:
                st.session_state.log.append(
                    "AVISO: Nenhum CFOP encontrado nos registros C170/C190. "
                    "Verifique se o arquivo é um SPED Fiscal EFD ICMS/IPI válido."
                )
            else:
                xlsx_bytes = gerar_xlsx_acumuladores_tr(cfops_dict)
                nome_base  = st.session_state.arquivo_nome.replace('.txt', '')
                nome_xlsx  = f"{nome_base}_acumuladores.xlsx"
                st.session_state.xlsx_bytes      = xlsx_bytes
                st.session_state.xlsx_nome       = nome_xlsx
                st.session_state.cfops_extraidos = cfops_dict
                st.session_state.log.append(
                    f"✔ {len(cfops_dict)} CFOP(s) extraído(s) — "
                    f"planilha gerada: {nome_xlsx}"
                )

        except Exception:
            st.session_state.log.append("ERRO FATAL na extração de CFOPs.")
            st.session_state.log.append(traceback.format_exc())

        st.rerun()

    # Métricas dos CFOPs extraídos
    if st.session_state.cfops_extraidos:
        cfops_dict = st.session_state.cfops_extraidos
        n_ent = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Entrada')
        n_sai = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Saída')

        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("CFOPs únicos", len(cfops_dict))
        col_m2.metric("Entradas",     n_ent)
        col_m3.metric("Saídas",       n_sai)

        with st.expander("📋 CFOPs identificados no SPED", expanded=False):
            rows = [
                {
                    'CFOP':        cfop,
                    'Descrição':   DESCRICOES_CFOP.get(cfop, '—'),
                    'Tipo':        info['tipo_operacao'],
                    'Ocorrências': info['ocorrencias'],
                    'Registros':   ', '.join(sorted(info['registros'])),
                }
                for cfop, info in sorted(cfops_dict.items())
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════
    # ETAPA 2 — TABELA DE ACUMULADORES + CONVERSÃO
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### ▶ Etapa 2 — Converter com a tabela de acumuladores preenchida")

    arquivo_acum = st.file_uploader(
        "📂 Tabela de Acumuladores preenchida (.xlsx ou .csv)",
        type=["xlsx", "xls", "csv"],
        help="Planilha com colunas CFOP e ACUMULADOR preenchidos",
        key="upload_acum",
    )

    if arquivo_acum is not None:
        log_temp = []
        raw_acum = arquivo_acum.read()
        tab_prev = carregar_acumuladores(raw_acum, arquivo_acum.name, log_temp)
        arquivo_acum.seek(0)
        if tab_prev is not None:
            st.success(
                f"✅ Tabela válida — "
                f"**{len(tab_prev)} CFOPs** com acumulador preenchido."
            )
            st.session_state.tabela_acum_ok = True
        else:
            for msg in log_temp:
                st.error(msg)
            st.session_state.tabela_acum_ok = False
    else:
        if not st.session_state.tabela_acum_ok:
            st.info("⬆ Faça o upload da tabela de acumuladores preenchida para converter.")

    pode_converter = (
        st.session_state.tabela_acum_ok and
        st.session_state.arquivo_raw is not None
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        converter = st.button(
            "▶ Converter SPED → Domínio",
            disabled=not pode_converter,
            use_container_width=True,
            type="primary",
        )
    with col2:
        limpar = st.button("🗑 Limpar Tudo", use_container_width=True)

    if limpar:
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    # ── Conversão ─────────────────────────────────────────────────────────
    if converter and pode_converter:
        st.session_state.log       = ["Iniciando conversão SPED → Domínio Sistemas..."]
        st.session_state.resultado = None
        st.session_state.stats     = None

        try:
            arquivo_acum.seek(0)
            tabela_acum = carregar_acumuladores(
                arquivo_acum.read(),
                arquivo_acum.name,
                st.session_state.log,
            )
            if tabela_acum is None:
                st.session_state.log.append("ERRO: Tabela inválida. Abortando.")
                st.rerun()

            st.session_state.log.append(
                f"Usando arquivo SPED em memória: {st.session_state.arquivo_nome}"
            )
            content = decode_arquivo(st.session_state.arquivo_raw)
            parsed  = parse_sped(content)

            resultado_txt, stats = converter_sped_para_dominio(
                parsed, tabela_acum, st.session_state.log
            )

            resultado_bytes = encode_ansi_seguro(
                resultado_txt, st.session_state.log
            )

            st.session_state.resultado  = resultado_bytes
            st.session_state.stats      = stats
            st.session_state.nome_saida = (
                st.session_state.arquivo_nome.replace('.txt', '_dominio.txt')
            )

        except Exception:
            st.session_state.log.append("ERRO FATAL durante a conversão.")
            st.session_state.log.append(traceback.format_exc())

        st.rerun()

    # ── Resultado ──────────────────────────────────────────────────────────
    if st.session_state.resultado is not None:
        st.success("✅ Arquivo convertido com sucesso!")

        stats = st.session_state.stats or {}
        st.markdown("#### 📊 Estatísticas da Conversão")
        col1, col2, col3 = st.columns(3)
        col1.metric("NFs Entrada",  stats.get('nf_entrada',  0))
        col2.metric("NFs Saída",    stats.get('nf_saida',    0))
        col3.metric("Itens",        stats.get('itens',       0))

        col4, col5, col6 = st.columns(3)
        col4.metric("Transporte",   stats.get('transporte',  0))
        col5.metric("Inventário",   stats.get('inventario',  0))
        col6.metric("Erros",        stats.get('erros',       0))

        st.markdown("---")

        with st.expander("👁️ Prévia do arquivo gerado (primeiras 60 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado
                .decode('latin-1', errors='replace')
                .splitlines()[:60]
            )
            st.code(preview, language='text')

        st.download_button(
            label="⬇ Baixar Arquivo Domínio Sistemas",
            data=st.session_state.resultado,
            file_name=st.session_state.nome_saida,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log ────────────────────────────────────────────────────────────────
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log)
    cor_borda = "#D32F2F" if tem_erro else "#388E3C"
    st.markdown(
        f"""
        <div style="background:#FCFCFC; border:1px solid {cor_borda};
                    border-radius:6px; padding:14px;
                    font-family:Consolas,monospace; font-size:13px;
                    white-space:pre-wrap; max-height:340px;
                    overflow-y:auto; color:#1F1F1F;">
{log_texto}
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.caption(
        "Conversor SPED Fiscal → Domínio Sistemas | "
        "Thomson Reuters | Desenvolvido com Python + Streamlit"
    )


if __name__ == "__main__":
    main()

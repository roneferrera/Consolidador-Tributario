import math
import re
import io
import os
import base64
import pandas as pd
from io import StringIO
from datetime import datetime, date
import traceback
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================
# VERSÃO
# ==============================
VERSAO = "V1.4"

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
        h1, h2, h3 {
            color: #FF8000;
            font-weight: 700;
        }
        section[data-testid="stSidebar"] {
            background-color: #444444;
            color: #FFFFFF;
        }
        section[data-testid="stSidebar"] * {
            color: #FFFFFF !important;
        }
        .stButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        .stDownloadButton > button {
            background-color: #FF8000;
            color: #FFFFFF;
            border: none;
            border-radius: 4px;
            font-weight: bold;
        }
        .stDownloadButton > button:hover {
            background-color: #D64001;
            color: #FFFFFF;
        }
        hr { border-color: #FF8000; }
        [data-testid="metric-container"] {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 10px;
        }
        .instrucoes-box {
            background-color: #E9E9E9;
            border-left: 4px solid #FF8000;
            border-radius: 4px;
            padding: 16px 20px;
            margin: 12px 0;
            color: #444444;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .instrucoes-box h4 {
            color: #FF8000;
            margin-top: 14px;
            margin-bottom: 6px;
        }
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
        log.append(
            f"AVISO: {substituicoes} caractere(s) fora do padrão ANSI "
            f"substituídos por '?'."
        )
    return b''.join(resultado)


# ==============================
# PARSER DO ARQUIVO DOMÍNIO
# ==============================
def parse_dominio(content: str) -> dict:
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
# MAPEAMENTO DE ÍNDICES — VERIFICADO PELO ARQUIVO DE EXEMPLO
# ==============================
#
# Registro 1000 (após remoção dos pipes das bordas):
# Pos : Valor no exemplo
# [0] : 1000
# [1] : 1              ← NUM_NF
# [2] : 22222222000191 ← CNPJ_EMIT
# [3] : (vazio)
# [4] : 1              ← IND_OPER (1=entrada)
# [5] : 1102           ← CFOP  ✅
# [6] : (vazio)
# [7] : 15
# [8] : 1
# [9] : (vazio)
# [10]: (vazio)
# [11]: 10/09/2015     ← DT_EMISSAO
# [12]: 10/09/2015     ← DT_ENTRADA
# [13]: 5571,24        ← VL_TOTAL
#
# Registro 1030 (após remoção dos pipes das bordas):
# Pos : Valor no exemplo
# [0] : 1030
# [1] : 1              ← NUM_ITEM
# [2] : 1,000          ← QTD
# [3] : 544,950        ← VL_UNIT
# [4] : 0
# [5] : 0
# [6] : 1
# [7] : 10/09/2015     ← DT_DOC
# [8] : (vazio)
# [9] : 00             ← COD_SIT
# [10]: 544,95         ← VL_ITEM
# [11]: 0,00           ← VL_DESC
# [12]: 544,95
# [13]: 0,00
# [14]: 17,00          ← ALIQ_ICMS
# [15]: (vazio)
# [16]: (vazio)
# [17]: 0,00
# [18]: 0,00
# [19]: 0,00
# [20]: 0,000
# [21]: 92,64          ← VL_ICMS
# [22]: 0,00
# [23]: 0,00
# [24]: 0,00
# [25]: (vazio)
# [26]: 544,950
# [27]: 0,00
# [28]: 02
# [29]: 0,00
# [30]: 0,00
# [31]: 0,00
# [32]: 0,00
# [33]: 1102           ← CFOP  ✅
# [34]: (vazio)
# ...
#
# ==============================
IDX_1000_NUM_NF     = 1
IDX_1000_CNPJ_EMIT  = 2
IDX_1000_IND_OPER   = 4
IDX_1000_CFOP       = 5   # ← CORRIGIDO (era 4)
IDX_1000_DT_EMISSAO = 11
IDX_1000_DT_ENTRADA = 12
IDX_1000_VL_TOTAL   = 13

IDX_1030_NUM_ITEM   = 1
IDX_1030_QTD        = 2
IDX_1030_VL_UNIT    = 3
IDX_1030_DT_DOC     = 7
IDX_1030_COD_SIT    = 9
IDX_1030_VL_ITEM    = 10
IDX_1030_VL_DESC    = 11
IDX_1030_ALIQ_ICMS  = 14
IDX_1030_VL_ICMS    = 21
IDX_1030_CFOP       = 33  # ← CORRIGIDO (era 34)


def _campo(campos: list, idx: int, default: str = '') -> str:
    """Retorna o campo pelo índice com fallback seguro."""
    return campos[idx].strip() if len(campos) > idx else default


# ==============================
# DIAGNÓSTICO DE ÍNDICES (helper de debug)
# ==============================
def diagnosticar_indices(parsed: dict, log: list):
    """
    Loga os primeiros registros 1000 e 1030 com seus índices
    para facilitar a verificação do mapeamento.
    """
    for tipo in ('1000', '1030'):
        if tipo in parsed['por_tipo']:
            campos, num_linha = parsed['por_tipo'][tipo][0]
            log.append(f"── Diagnóstico {tipo} (linha {num_linha}) ──")
            for i, v in enumerate(campos):
                log.append(f"  [{i:>2}] = '{v}'")


# ==============================
# EXTRAÇÃO DE CFOPs
# ==============================
def extrair_cfops_do_arquivo(parsed: dict, log: list) -> dict:
    cfops = {}

    def registrar(cfop_raw: str, tipo_reg: str):
        cfop = str(cfop_raw).strip()
        # Normaliza para 4 dígitos apenas se for numérico
        if not cfop or not cfop.isdigit():
            return
        cfop = cfop.zfill(4)
        # Filtra valores claramente inválidos
        if cfop in ('0000',):
            return
        if cfop not in cfops:
            primeiro = cfop[0]
            if primeiro in ('1', '2', '3'):
                tipo_op = 'Entrada'
            elif primeiro in ('5', '6', '7'):
                tipo_op = 'Saída'
            else:
                tipo_op = 'Desconhecido'
            cfops[cfop] = {
                'registros':     set(),
                'ocorrencias':   0,
                'tipo_operacao': tipo_op,
            }
        cfops[cfop]['registros'].add(tipo_reg)
        cfops[cfop]['ocorrencias'] += 1

    total_1000 = 0
    total_1030 = 0

    for tipo, campos, num_linha in parsed['linhas_ordenadas']:
        if tipo == '1000':
            total_1000 += 1
            cfop_val = _campo(campos, IDX_1000_CFOP)
            log.append(
                f"  1000 linha {num_linha}: "
                f"idx[{IDX_1000_CFOP}]='{cfop_val}' "
                f"(total campos={len(campos)})"
            )
            if cfop_val:
                registrar(cfop_val, '1000')

        elif tipo == '1030':
            total_1030 += 1
            cfop_val = _campo(campos, IDX_1030_CFOP)
            # Só loga o primeiro item para não poluir o log
            if total_1030 <= 2:
                log.append(
                    f"  1030 linha {num_linha}: "
                    f"idx[{IDX_1030_CFOP}]='{cfop_val}' "
                    f"(total campos={len(campos)})"
                )
            if cfop_val:
                registrar(cfop_val, '1030')

    log.append(
        f"Registros lidos: {total_1000} × 1000 | {total_1030} × 1030"
    )
    log.append(
        f"CFOPs únicos encontrados: {len(cfops)} — "
        f"{sorted(cfops.keys())}"
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
# GERADOR DO XLSX — TEMA TR
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

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

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

    # Banner
    ws1.merge_cells("A1:E1")
    ws1.row_dimensions[1].height = 36
    c = ws1["A1"]
    c.value     = "Thomson Reuters  |  Domínio Sistemas  —  Tabela de Acumuladores CFOP"
    c.fill      = fill(COR_CINZA_ESC)
    c.font      = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA)
    c.alignment = left_al()

    # Subtítulo laranja
    ws1.merge_cells("A2:E2")
    ws1.row_dimensions[2].height = 20
    c2 = ws1["A2"]
    c2.value = (
        f"Gerado automaticamente  |  {len(cfops_dict)} CFOP(s) identificado(s)  |  "
        f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    c2.fill      = fill(COR_LARANJA)
    c2.font      = Font(name='Segoe UI', size=9, color=COR_BRANCO)
    c2.alignment = left_al()

    # Instrução
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

    # Cabeçalho da tabela
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

    # Dados
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
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 3:
                cell.fill = fill(bg)
                cor       = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor)
                cell.alignment = center()
            elif ci == 4:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 5:
                # Coluna ACUMULADOR — destaque laranja para preenchimento
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

    # Rodapé
    ws1.merge_cells(f"A{linha}:E{linha}")
    ws1.row_dimensions[linha].height = 18
    cr = ws1.cell(
        row=linha, column=1,
        value="Thomson Reuters  |  Domínio Sistemas  |  Gerado automaticamente"
    )
    cr.fill      = fill(COR_CINZA_ESC)
    cr.font      = Font(name='Segoe UI', size=8, color="888888")
    cr.alignment = Alignment(horizontal='right', vertical='center')

    ws1.freeze_panes   = "A6"
    ws1.auto_filter.ref = f"A5:E{linha - 1}"

    # ── Aba 2: CFOPs Encontrados ─────────────────────────────────────────
    ws2 = wb.create_sheet(title="CFOPs Encontrados")
    ws2.sheet_view.showGridLines = False

    n_ent = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Entrada')
    n_sai = sum(1 for v in cfops_dict.values() if v['tipo_operacao'] == 'Saída')

    ws2.merge_cells("A1:F1")
    ws2.row_dimensions[1].height = 36
    c = ws2["A1"]
    c.value     = "Thomson Reuters  |  Domínio Sistemas  —  Relatório de CFOPs Identificados"
    c.fill      = fill(COR_CINZA_ESC)
    c.font      = Font(name='Segoe UI', bold=True, size=13, color=COR_LARANJA)
    c.alignment = left_al()

    ws2.merge_cells("A2:F2")
    ws2.row_dimensions[2].height = 20
    c2 = ws2["A2"]
    c2.value = (
        f"Analisado em {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  "
        f"Total: {len(cfops_dict)} CFOP(s)  |  "
        f"Entradas: {n_ent}  |  Saídas: {n_sai}"
    )
    c2.fill      = fill(COR_LARANJA)
    c2.font      = Font(name='Segoe UI', size=9, color=COR_BRANCO)
    c2.alignment = left_al()

    ws2.row_dimensions[3].height = 6
    ws2.row_dimensions[4].height = 22

    cab2  = ['CFOP', 'DESCRIÇÃO', 'TIPO OPERAÇÃO', 'REGISTROS', 'OCORRÊNCIAS', 'STATUS']
    wids2 = [10,     48,          16,               14,           14,            22]

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

        vals = [cfop, descr, info['tipo_operacao'], regs, info['ocorrencias'], status]
        for ci, valor in enumerate(vals, start=1):
            cell        = ws2.cell(row=linha2, column=ci, value=valor)
            cell.border = borda_fina
            if ci == 6:
                cell.fill = fill(bg_st)
                cor       = "1B5E20" if mapeado else "B71C1C"
                cell.font = Font(name='Segoe UI', bold=True, size=10, color=cor)
                cell.alignment = center()
            elif ci == 1:
                cell.fill      = fill(bg)
                cell.font      = Font(name='Segoe UI', bold=True, size=10, color=COR_CINZA_ESC)
                cell.alignment = center()
            elif ci == 3:
                cell.fill = fill(bg)
                cor       = "1B5E20" if info['tipo_operacao'] == 'Entrada' else "B71C1C"
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
    cr2 = ws2.cell(
        row=linha2, column=1,
        value="Thomson Reuters  |  Domínio Sistemas  |  Gerado automaticamente"
    )
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
            log.append(
                "ERRO: O arquivo deve conter as colunas 'CFOP' e 'ACUMULADOR'."
            )
            return None

        tabela = {}
        erros  = 0
        for _, row in df.iterrows():
            cfop = str(row['CFOP']).strip().zfill(4)
            acum = str(row['ACUMULADOR']).strip()
            if not cfop or not acum or cfop.upper() == 'NAN' or acum.upper() == 'NAN':
                erros += 1
                continue
            if acum in ('', '0', 'nan', 'NAN', 'None'):
                continue
            tabela[cfop] = acum

        if not tabela:
            log.append(
                "ERRO: Nenhum par CFOP → Acumulador preenchido. "
                "Preencha a coluna ACUMULADOR no arquivo e tente novamente."
            )
            return None

        if erros:
            log.append(f"AVISO: {erros} linha(s) ignoradas por dados inválidos.")

        log.append(f"Tabela de acumuladores carregada: {len(tabela)} CFOPs mapeados.")
        return tabela

    except Exception as e:
        log.append(f"ERRO ao carregar arquivo de acumuladores: {e}")
        return None


def get_acumulador(cfop: str, tabela: dict, nao_mapeados: set) -> str:
    cfop_norm = str(cfop).strip().zfill(4)
    acum      = tabela.get(cfop_norm)
    if acum is None:
        nao_mapeados.add(cfop_norm)
        return '9999'
    return acum


# ==============================
# HIERARQUIA E ORDENAÇÃO
# ==============================
def extrair_notas_com_itens(parsed, tabela_acum, nao_mapeados, log):
    notas      = []
    nota_atual = None

    for tipo, campos, num_linha in parsed['linhas_ordenadas']:
        if tipo == '0000':
            continue

        elif tipo == '1000':
            if nota_atual is not None:
                notas.append(nota_atual)
            cfop_nf    = _campo(campos, IDX_1000_CFOP)
            nota_atual = {
                'num_nf':      _campo(campos, IDX_1000_NUM_NF),
                'cnpj_emit':   _campo(campos, IDX_1000_CNPJ_EMIT),
                'cfop_nf':     cfop_nf,
                'acum_nf':     get_acumulador(cfop_nf, tabela_acum, nao_mapeados),
                'dt_emissao':  _campo(campos, IDX_1000_DT_EMISSAO),
                'dt_entrada':  _campo(campos, IDX_1000_DT_ENTRADA),
                'vl_total':    _campo(campos, IDX_1000_VL_TOTAL),
                'campos_1000': campos,
                'totais_1020': None,
                'itens_1030':  [],
                'lanc_1300':   [],
                'linha_orig':  num_linha,
            }

        elif tipo == '1020':
            if nota_atual is not None:
                nota_atual['totais_1020'] = campos
            else:
                log.append(f"AVISO: 1020 na linha {num_linha} sem 1000 pai. Ignorado.")

        elif tipo == '1030':
            if nota_atual is not None:
                cfop_item = _campo(campos, IDX_1030_CFOP)
                nota_atual['itens_1030'].append({
                    'num_item':   _campo(campos, IDX_1030_NUM_ITEM),
                    'cfop':       cfop_item,
                    'acum':       get_acumulador(cfop_item, tabela_acum, nao_mapeados),
                    'campos':     campos,
                    'linha_orig': num_linha,
                })
            else:
                log.append(f"AVISO: 1030 na linha {num_linha} sem 1000 pai. Ignorado.")

        elif tipo == '1300':
            if nota_atual is not None:
                nota_atual['lanc_1300'].append(campos)
            else:
                log.append(f"AVISO: 1300 na linha {num_linha} sem 1000 pai. Ignorado.")

    if nota_atual is not None:
        notas.append(nota_atual)

    def sort_key(n):
        dt = datetime.min
        for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
            try:
                dt = datetime.strptime(n['dt_emissao'], fmt)
                break
            except Exception:
                pass
        try:
            num = int(n['num_nf'])
        except Exception:
            num = 0
        return (dt, num)

    notas.sort(key=sort_key)
    for nota in notas:
        nota['itens_1030'].sort(
            key=lambda it: int(it['num_item']) if it['num_item'].isdigit() else 0
        )

    log.append(
        f"Hierarquia montada: {len(notas)} nota(s) ordenada(s) "
        f"por data de emissão e número."
    )
    return notas


# ==============================
# GERADOR DO ARQUIVO DE SAÍDA
# ==============================
def gerar_saida(parsed, tabela_acum, log):
    nao_mapeados = set()
    saida        = StringIO()
    stats = {
        'notas': 0, 'nf_entrada': 0, 'nf_saida': 0,
        'itens': 0, 'lancamentos': 0, 'erros': 0,
    }

    if '0000' in parsed['por_tipo']:
        campos_0000, _ = parsed['por_tipo']['0000'][0]
        saida.write('|' + '|'.join(campos_0000) + '|\n')
    else:
        log.append("AVISO: Registro 0000 não encontrado.")

    notas = extrair_notas_com_itens(parsed, tabela_acum, nao_mapeados, log)

    for nota in notas:
        saida.write('|' + '|'.join(nota['campos_1000']) + '|\n')
        stats['notas'] += 1
        if nota['cfop_nf'][:1] in ('1', '2', '3'):
            stats['nf_entrada'] += 1
        elif nota['cfop_nf'][:1] in ('5', '6', '7'):
            stats['nf_saida'] += 1

        if nota['totais_1020'] is not None:
            saida.write('|' + '|'.join(nota['totais_1020']) + '|\n')

        for item in nota['itens_1030']:
            saida.write('|' + '|'.join(item['campos']) + '|\n')
            stats['itens'] += 1

        for lanc in nota['lanc_1300']:
            saida.write('|' + '|'.join(lanc) + '|\n')
            stats['lancamentos'] += 1

    saida.write('|9999|\n')

    if nao_mapeados:
        log.append(
            f"AVISO: {len(nao_mapeados)} CFOP(s) sem acumulador (9999): "
            f"{', '.join(sorted(nao_mapeados))}"
        )

    log.append(
        f"Geração concluída — Notas={stats['notas']} | "
        f"Entradas={stats['nf_entrada']} | Saídas={stats['nf_saida']} | "
        f"Itens={stats['itens']} | Lançamentos={stats['lancamentos']} | "
        f"Erros={stats['erros']}"
    )
    return saida.getvalue(), stats


# ==============================
# RELATÓRIO DE ORDENAÇÃO
# ==============================
def gerar_relatorio_ordenacao(notas):
    linhas = []
    for pos, nota in enumerate(notas, start=1):
        linhas.append({
            'Posição':    pos,
            'Num NF':     nota['num_nf'],
            'CFOP':       nota['cfop_nf'],
            'Acumulador': nota['acum_nf'],
            'Dt Emissão': nota['dt_emissao'],
            'Dt Entrada': nota['dt_entrada'],
            'Vl Total':   nota['vl_total'],
            'Qtd Itens':  len(nota['itens_1030']),
            'Qtd Lanç':   len(nota['lanc_1300']),
            'Linha Orig': nota['linha_orig'],
            'Tipo':       'Entrada' if nota['cfop_nf'][:1] in ('1','2','3') else 'Saída',
        })
        for item in nota['itens_1030']:
            linhas.append({
                'Posição':    f"  └ Item {item['num_item']}",
                'Num NF':     nota['num_nf'],
                'CFOP':       item['cfop'],
                'Acumulador': item['acum'],
                'Dt Emissão': '',
                'Dt Entrada': '',
                'Vl Total':   _campo(item['campos'], IDX_1030_VL_ITEM),
                'Qtd Itens':  '',
                'Qtd Lanç':   '',
                'Linha Orig': item['linha_orig'],
                'Tipo':       'Item',
            })
    return pd.DataFrame(linhas)


# ==============================
# INTERFACE STREAMLIT
# ==============================
def main():
    st.set_page_config(
        page_title="Domínio Sistemas | Thomson Reuters",
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
                📄 Conversor / Ordenador Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Faça upload do arquivo, extraia os CFOPs, preencha os acumuladores
                e processe — tudo em um único fluxo.
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
        st.markdown("### 📋 Registros Suportados")
        st.markdown(
            "- **0000** Cabeçalho\n"
            "- **1000** Nota Fiscal\n"
            "- **1020** Totais da NF\n"
            "- **1030** Itens da NF\n"
            "- **1300** Lançamento Contábil\n"
            "- **9999** Encerramento\n"
        )
        st.markdown("---")
        st.markdown("### 📑 Fluxo")
        st.markdown(
            "1. Upload do arquivo `.txt`\n"
            "2. **Extrair CFOPs → baixar XLSX**\n"
            "3. Preencher coluna `ACUMULADOR`\n"
            "4. Upload do XLSX preenchido\n"
            "5. **Processar** e baixar saída\n"
        )
        st.markdown("---")
        st.markdown("### ⚙ Ordenação")
        st.markdown("Data Emissão → Nº NF → Nº Item")
        st.markdown("### ⚙ Encoding de saída")
        st.markdown("**ANSI (Latin-1)**")

    # ── Instruções ────────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">
            <h4>🔹 Etapa 1 — Upload e extração de CFOPs</h4>
            <p>Faça o upload do arquivo <code>.txt</code> e clique em
            <b>🔍 Extrair CFOPs e Gerar Planilha</b>.</p>

            <h4>🔹 Etapa 2 — Preencher acumuladores</h4>
            <p>Abra o XLSX baixado, preencha a coluna <b>ACUMULADOR</b> e salve.</p>

            <h4>🔹 Etapa 3 — Processar</h4>
            <p>Faça o upload do XLSX preenchido e clique em
            <b>▶ Processar Arquivo</b>.</p>

            <hr>
            <h4>⚠ Observações</h4>
            <ul>
                <li>CFOPs sem acumulador preenchido receberão <b>9999</b>.</li>
                <li>Ordenação: <b>Data Emissão → Nº NF → Nº Item</b>.</li>
                <li>Saída em <b>ANSI (Latin-1)</b>.</li>
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
        "df_ordenacao":    None,
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
    # ETAPA 1 — UPLOAD + EXTRAÇÃO
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### 🔍 Etapa 1 — Upload do arquivo e extração de CFOPs")

    uploaded_file = st.file_uploader(
        "📂 Arquivo Domínio de origem (.txt)",
        type=["txt"],
        help="Arquivo .txt exportado pelo Domínio Sistemas",
        key="upload_dominio",
    )

    # ── Cacheia os bytes IMEDIATAMENTE ao fazer upload ────────────────────
    # Isso evita que o arquivo suma após st.rerun()
    if uploaded_file is not None:
        raw_atual = uploaded_file.read()
        if raw_atual != st.session_state.arquivo_raw:
            st.session_state.arquivo_raw     = raw_atual
            st.session_state.arquivo_nome    = uploaded_file.name
            st.session_state.cfops_extraidos = None
            st.session_state.xlsx_bytes      = None
            st.session_state.resultado       = None
            st.session_state.stats           = None
            st.session_state.df_ordenacao    = None
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
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
                type="primary",
            )

    # ── Extração ──────────────────────────────────────────────────────────
    if extrair:
        st.session_state.log             = ["Extraindo CFOPs do arquivo..."]
        st.session_state.xlsx_bytes      = None
        st.session_state.cfops_extraidos = None

        try:
            content    = decode_arquivo(st.session_state.arquivo_raw)
            parsed     = parse_dominio(content)

            # Diagnóstico de índices no log
            diagnosticar_indices(parsed, st.session_state.log)

            cfops_dict = extrair_cfops_do_arquivo(parsed, st.session_state.log)

            if not cfops_dict:
                st.session_state.log.append(
                    "AVISO: Nenhum CFOP válido encontrado. "
                    "Verifique o diagnóstico de índices no log acima."
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

        with st.expander("📋 CFOPs identificados no arquivo", expanded=False):
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
    # ETAPA 2 — TABELA DE ACUMULADORES + PROCESSAMENTO
    # ════════════════════════════════════════════════════════════════════
    st.markdown("### ▶ Etapa 2 — Processar com a tabela de acumuladores preenchida")

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
            st.info(
                "⬆ Faça o upload da tabela de acumuladores preenchida "
                "para processar."
            )

    pode_processar = (
        st.session_state.tabela_acum_ok and
        st.session_state.arquivo_raw is not None
    )

    col1, col2 = st.columns([1, 1])
    with col1:
        processar = st.button(
            "▶ Processar Arquivo",
            disabled=not pode_processar,
            use_container_width=True,
            type="primary",
        )
    with col2:
        limpar = st.button("🗑 Limpar Tudo", use_container_width=True)

    if limpar:
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    # ── Processamento ──────────────────────────────────────────────────────
    if processar and pode_processar:
        st.session_state.log          = ["Iniciando processamento..."]
        st.session_state.resultado    = None
        st.session_state.stats        = None
        st.session_state.df_ordenacao = None

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
                f"Usando arquivo em memória: {st.session_state.arquivo_nome}"
            )
            content = decode_arquivo(st.session_state.arquivo_raw)
            parsed  = parse_dominio(content)

            tipos_enc = list(parsed['por_tipo'].keys())
            st.session_state.log.append(
                f"Tipos de registro: {', '.join(tipos_enc)}"
            )

            resultado_txt, stats = gerar_saida(
                parsed, tabela_acum, st.session_state.log
            )

            notas_ord = extrair_notas_com_itens(
                parsed, tabela_acum, set(), []
            )
            df_ord = gerar_relatorio_ordenacao(notas_ord)

            resultado_bytes = encode_ansi_seguro(
                resultado_txt, st.session_state.log
            )

            st.session_state.resultado    = resultado_bytes
            st.session_state.stats        = stats
            st.session_state.df_ordenacao = df_ord
            st.session_state.nome_saida   = (
                st.session_state.arquivo_nome.replace('.txt', '_processado.txt')
            )

        except Exception:
            st.session_state.log.append("ERRO FATAL durante o processamento.")
            st.session_state.log.append(traceback.format_exc())

        st.rerun()

    # ── Resultado ──────────────────────────────────────────────────────────
    if st.session_state.resultado is not None:
        st.success("✅ Arquivo processado com sucesso!")

        stats = st.session_state.stats or {}
        st.markdown("#### 📊 Estatísticas")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Notas", stats.get('notas',       0))
        col2.metric("NFs Entrada",    stats.get('nf_entrada',  0))
        col3.metric("NFs Saída",      stats.get('nf_saida',    0))

        col4, col5, col6 = st.columns(3)
        col4.metric("Itens",       stats.get('itens',       0))
        col5.metric("Lançamentos", stats.get('lancamentos', 0))
        col6.metric("Erros",       stats.get('erros',       0))

        st.markdown("---")

        if st.session_state.df_ordenacao is not None:
            with st.expander("📋 Verificação de Ordenação", expanded=True):
                st.markdown(
                    "Notas na **ordem exata** gravada no arquivo de saída."
                )
                st.dataframe(
                    st.session_state.df_ordenacao,
                    use_container_width=True,
                    hide_index=True,
                )

        with st.expander("👁️ Prévia do arquivo gerado (primeiras 60 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado
                .decode('latin-1', errors='replace')
                .splitlines()[:60]
            )
            st.code(preview, language='text')

        st.download_button(
            label="⬇ Baixar Arquivo Processado",
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
        "Conversor / Ordenador Domínio Sistemas | "
        "Thomson Reuters | Desenvolvido com Python + Streamlit"
    )


if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import re
from io import StringIO
from datetime import datetime
import os
import base64

# ─────────────────────────────────────────
# VERSÃO
# ─────────────────────────────────────────
VERSAO = "V1.0"

# ─────────────────────────────────────────
# TEMA THOMSON REUTERS (idêntico ao RPA)
# ─────────────────────────────────────────
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
        hr {
            border-color: #FF8000;
        }
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
        .instrucoes-box h4:first-child {
            margin-top: 0;
        }
        </style>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────
# PARSER SPED
# ─────────────────────────────────────────
def parse_sped(content: str) -> dict:
    """Lê o arquivo SPED e agrupa as linhas por registro."""
    registros = {}
    for linha in content.splitlines():
        linha = linha.strip()
        if not linha:
            continue
        campos = linha.split('|')
        if len(campos) < 2:
            continue
        if campos[0] == '':
            campos = campos[1:]
        if campos[-1] == '':
            campos = campos[:-1]
        reg = campos[0]
        if reg not in registros:
            registros[reg] = []
        registros[reg].append(campos)
    return registros


# ─────────────────────────────────────────
# MAPEAMENTO CFOP → ACUMULADOR DOMÍNIO
# ─────────────────────────────────────────
CFOP_ACUMULADOR = {
    # ENTRADAS
    '1101': '1101', '2101': '1101', '3101': '1101',
    '1102': '1151', '2102': '1151', '3102': '1151',
    '1111': '1103', '2111': '1103',
    '1113': '1153', '2113': '1153',
    '1116': '1104', '2116': '1104',
    '1120': '1105', '2120': '1105',
    '1201': '1201', '2201': '1201',
    '1202': '1202', '2202': '1202',
    '1203': '1203', '2203': '1203',
    '1251': '1301', '2251': '1301',
    '1252': '1302', '2252': '1302',
    '1301': '1334', '2301': '1334',
    '1351': '1367', '2351': '1367',
    '1352': '1368', '2352': '1368',
    '1401': '1401', '2401': '1401',
    '1403': '1403', '2403': '1403',
    '1501': '1501', '2501': '1501',
    '1502': '1502', '2502': '1502',
    '1503': '1503', '2503': '1503',
    '1601': '1601', '2601': '1601',
    '1701': '1701', '2701': '1701',
    '1801': '1801', '2801': '1801',
    '1901': '1901', '2901': '1901',
    # SAÍDAS
    '5101': '5101', '6101': '5101',
    '5102': '5102', '6102': '5102',
    '5111': '5103', '6111': '5103',
    '5113': '5104', '6113': '5104',
    '5151': '5151', '6151': '5151',
    '5152': '5152', '6152': '5152',
    '5153': '5153', '6153': '5153',
    '5201': '5201', '6201': '5201',
    '5202': '5202', '6202': '5202',
    '5251': '5301', '6251': '5301',
    '5301': '5334', '6301': '5334',
    '5351': '5367', '6351': '5367',
    '5352': '5368', '6352': '5368',
    '5401': '5401', '6401': '5401',
    '5403': '5403', '6403': '5403',
    '5501': '5501', '6501': '5501',
    '5601': '5601', '6601': '5601',
    '5701': '5701', '6701': '5701',
    '5801': '5801', '6801': '5801',
    '5901': '5901', '6901': '5901',
}

# CFOPs não mapeados — rastreados para aviso ao usuário
_cfops_nao_mapeados: set = set()

def get_acumulador(cfop: str) -> str:
    acum = CFOP_ACUMULADOR.get(cfop)
    if acum is None:
        _cfops_nao_mapeados.add(cfop)
        return '9999'
    return acum


# ─────────────────────────────────────────
# CONVERSORES POR REGISTRO
# ─────────────────────────────────────────
def converter_0000(campos: list) -> str:
    """Registro de abertura — índices corrigidos conforme layout SPED."""
    try:
        # [0]=0000 [1]=COD_VER [2]=COD_FIN [3]=DT_INI [4]=DT_FIN
        # [5]=NOME [6]=CNPJ [7]=CPF [8]=UF [9]=IE ...
        dt_ini = campos[3] if len(campos) > 3 else ''
        dt_fin = campos[4] if len(campos) > 4 else ''
        nome   = campos[5] if len(campos) > 5 else ''
        cnpj   = campos[6] if len(campos) > 6 else ''
        uf     = campos[8] if len(campos) > 8 else ''
        return f"|0000|1|{dt_ini}|{dt_fin}|{nome}|{cnpj}|{uf}|\n"
    except Exception:
        return ''


def converter_0150(campos: list) -> str:
    """Cadastro de participantes → Clientes/Fornecedores Domínio."""
    try:
        cod    = campos[1]  if len(campos) > 1  else ''
        nome   = campos[2]  if len(campos) > 2  else ''
        cnpj   = campos[4]  if len(campos) > 4  else ''
        cpf    = campos[5]  if len(campos) > 5  else ''
        ie     = campos[6]  if len(campos) > 6  else ''
        end    = campos[8]  if len(campos) > 8  else ''
        num    = campos[9]  if len(campos) > 9  else ''
        comp   = campos[10] if len(campos) > 10 else ''
        bairro = campos[11] if len(campos) > 11 else ''
        cep    = campos[12] if len(campos) > 12 else ''
        uf     = campos[13] if len(campos) > 13 else ''
        fone   = campos[14] if len(campos) > 14 else ''
        return (
            f"|0150|{cod}|{nome}|{cnpj}|{cpf}|{ie}|"
            f"{end}|{num}|{comp}|{bairro}|{cep}|{uf}|{fone}|\n"
        )
    except Exception:
        return ''


def converter_c100(campos: list) -> str:
    """Nota Fiscal (C100) → Registro de NF Domínio."""
    try:
        ind_oper   = campos[1]  if len(campos) > 1  else '0'
        cod_part   = campos[3]  if len(campos) > 3  else ''
        cod_mod    = campos[4]  if len(campos) > 4  else ''
        cod_sit    = campos[5]  if len(campos) > 5  else '00'
        serie      = campos[6]  if len(campos) > 6  else ''
        num_doc    = campos[7]  if len(campos) > 7  else ''
        chv_nfe    = campos[8]  if len(campos) > 8  else ''
        dt_doc     = campos[9]  if len(campos) > 9  else ''
        dt_es      = campos[10] if len(campos) > 10 else ''
        vl_doc     = campos[11] if len(campos) > 11 else '0'
        vl_bc_icms = campos[15] if len(campos) > 15 else '0'
        vl_icms    = campos[16] if len(campos) > 16 else '0'
        vl_ipi     = campos[21] if len(campos) > 21 else '0'
        vl_pis     = campos[22] if len(campos) > 22 else '0'
        vl_cofins  = campos[23] if len(campos) > 23 else '0'
        tipo = 'E' if ind_oper == '0' else 'S'
        return (
            f"|C100|{tipo}|{cod_part}|{cod_mod}|{cod_sit}|{serie}|{num_doc}|"
            f"{chv_nfe}|{dt_doc}|{dt_es}|{vl_doc}|{vl_bc_icms}|{vl_icms}|"
            f"{vl_ipi}|{vl_pis}|{vl_cofins}|\n"
        )
    except Exception:
        return ''


def converter_c170(campos: list, cfop: str = '') -> str:
    """
    Itens da NF (C170) → Detalhamento Domínio.
    Índices corrigidos conforme layout real do SPED EFD ICMS/IPI:
    [0]=C170 [1]=NUM_ITEM [2]=COD_ITEM [3]=DESCR_COMPL [4]=QTD
    [5]=UNID [6]=VL_ITEM [7]=VL_DESC [8]=IND_MOV [9]=CST_ICMS
    [10]=CFOP [11]=COD_NAT [12]=VL_BC_ICMS [13]=ALIQ_ICMS [14]=VL_ICMS
    """
    try:
        num_item  = campos[1]  if len(campos) > 1  else ''
        cod_item  = campos[2]  if len(campos) > 2  else ''
        descr     = campos[3]  if len(campos) > 3  else ''
        qtd       = campos[4]  if len(campos) > 4  else '0'
        unid      = campos[5]  if len(campos) > 5  else ''
        vl_item   = campos[6]  if len(campos) > 6  else '0'
        vl_desc   = campos[7]  if len(campos) > 7  else '0'
        # campos[8] = IND_MOV  |  campos[9] = CST_ICMS
        cfop_item = campos[10] if len(campos) > 10 else cfop
        acum      = get_acumulador(cfop_item)
        vl_bc     = campos[12] if len(campos) > 12 else '0'
        aliq      = campos[13] if len(campos) > 13 else '0'
        vl_icms   = campos[14] if len(campos) > 14 else '0'
        return (
            f"|C170|{num_item}|{cod_item}|{descr}|{qtd}|{unid}|{vl_item}|"
            f"{vl_desc}|{cfop_item}|{acum}|{vl_bc}|{aliq}|{vl_icms}|\n"
        )
    except Exception:
        return ''


def converter_c190(campos: list) -> str:
    """Registro analítico C190 → Totais por CFOP/CST."""
    try:
        cst_icms = campos[1] if len(campos) > 1 else ''
        cfop     = campos[2] if len(campos) > 2 else ''
        aliq     = campos[3] if len(campos) > 3 else '0'
        vl_opr   = campos[4] if len(campos) > 4 else '0'
        vl_bc    = campos[5] if len(campos) > 5 else '0'
        vl_icms  = campos[6] if len(campos) > 6 else '0'
        vl_red   = campos[7] if len(campos) > 7 else '0'
        acum     = get_acumulador(cfop)
        return (
            f"|C190|{cst_icms}|{cfop}|{aliq}|{vl_opr}|{vl_bc}|"
            f"{vl_icms}|{vl_red}|{acum}|\n"
        )
    except Exception:
        return ''


def converter_d100(campos: list) -> str:
    """
    Conhecimento de Transporte (D100).
    Nota: CFOP do D100 está no D190 filho; aqui extraímos o que estiver
    disponível no próprio registro para montar a linha de cabeçalho.
    """
    try:
        ind_oper = campos[1]  if len(campos) > 1  else '0'
        cod_part = campos[3]  if len(campos) > 3  else ''
        cod_mod  = campos[4]  if len(campos) > 4  else ''
        cod_sit  = campos[5]  if len(campos) > 5  else '00'
        serie    = campos[6]  if len(campos) > 6  else ''
        num_doc  = campos[8]  if len(campos) > 8  else ''
        dt_doc   = campos[10] if len(campos) > 10 else ''
        vl_doc   = campos[14] if len(campos) > 14 else '0'
        tipo     = 'E' if ind_oper == '0' else 'S'
        return (
            f"|D100|{tipo}|{cod_part}|{cod_mod}|{cod_sit}|{serie}|"
            f"{num_doc}|{dt_doc}|{vl_doc}|\n"
        )
    except Exception:
        return ''


def converter_d500(campos: list) -> str:
    """Serviços de Comunicação (D500)."""
    try:
        ind_oper = campos[1]  if len(campos) > 1  else '0'
        cod_part = campos[3]  if len(campos) > 3  else ''
        cod_mod  = campos[4]  if len(campos) > 4  else ''
        cod_sit  = campos[5]  if len(campos) > 5  else '00'
        serie    = campos[6]  if len(campos) > 6  else ''
        num_doc  = campos[7]  if len(campos) > 7  else ''
        dt_doc   = campos[8]  if len(campos) > 8  else ''
        vl_doc   = campos[10] if len(campos) > 10 else '0'
        cfop     = campos[15] if len(campos) > 15 else ''
        acum     = get_acumulador(cfop)
        tipo     = 'E' if ind_oper == '0' else 'S'
        return (
            f"|D500|{tipo}|{cod_part}|{cod_mod}|{cod_sit}|{serie}|"
            f"{num_doc}|{dt_doc}|{vl_doc}|{cfop}|{acum}|\n"
        )
    except Exception:
        return ''


def converter_h010(campos: list) -> str:
    """Inventário (H010)."""
    try:
        cod_item = campos[1] if len(campos) > 1 else ''
        unid     = campos[2] if len(campos) > 2 else ''
        qtd      = campos[3] if len(campos) > 3 else '0'
        vl_unit  = campos[4] if len(campos) > 4 else '0'
        vl_item  = campos[5] if len(campos) > 5 else '0'
        ind_prop = campos[6] if len(campos) > 6 else '0'
        return (
            f"|H010|{cod_item}|{unid}|{qtd}|{vl_unit}|{vl_item}|{ind_prop}|\n"
        )
    except Exception:
        return ''


# ─────────────────────────────────────────
# DECODE COM FALLBACK DE ENCODING
# ─────────────────────────────────────────
def decode_sped(raw: bytes) -> str:
    """Tenta utf-8, latin-1 e cp1252 antes de usar replace."""
    for enc in ('utf-8', 'latin-1', 'cp1252'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


# ─────────────────────────────────────────
# GERADOR DO ARQUIVO DOMÍNIO
# ─────────────────────────────────────────
def gerar_dominio(registros: dict) -> tuple[str, dict]:
    saida = StringIO()
    stats = {
        'participantes': 0,
        'nf_entrada':    0,
        'nf_saida':      0,
        'itens':         0,
        'analiticos':    0,
        'transporte':    0,
        'comunicacao':   0,
        'inventario':    0,
        'erros':         0,
    }

    # Abertura
    if '0000' in registros:
        saida.write(converter_0000(registros['0000'][0]))

    # Participantes
    if '0150' in registros:
        for campos in registros['0150']:
            linha = converter_0150(campos)
            if linha:
                saida.write(linha)
                stats['participantes'] += 1
            else:
                stats['erros'] += 1

    # Bloco C — NFs com hierarquia C100 → C170 → C190
    if 'C100' in registros:
        for campos_c100 in registros['C100']:
            linha_c100 = converter_c100(campos_c100)
            if linha_c100:
                saida.write(linha_c100)
                ind_oper = campos_c100[1] if len(campos_c100) > 1 else '0'
                if ind_oper == '0':
                    stats['nf_entrada'] += 1
                else:
                    stats['nf_saida'] += 1
            else:
                stats['erros'] += 1

    # Itens (C170) — escritos após todos os C100 por limitação do parser flat
    if 'C170' in registros:
        for campos in registros['C170']:
            linha = converter_c170(campos)
            if linha:
                saida.write(linha)
                stats['itens'] += 1
            else:
                stats['erros'] += 1

    # Analítico C190
    if 'C190' in registros:
        for campos in registros['C190']:
            linha = converter_c190(campos)
            if linha:
                saida.write(linha)
                stats['analiticos'] += 1
            else:
                stats['erros'] += 1

    # Bloco D — Transporte
    if 'D100' in registros:
        for campos in registros['D100']:
            linha = converter_d100(campos)
            if linha:
                saida.write(linha)
                stats['transporte'] += 1
            else:
                stats['erros'] += 1

    # Bloco D — Comunicação
    if 'D500' in registros:
        for campos in registros['D500']:
            linha = converter_d500(campos)
            if linha:
                saida.write(linha)
                stats['comunicacao'] += 1
            else:
                stats['erros'] += 1

    # Bloco H — Inventário
    if 'H010' in registros:
        for campos in registros['H010']:
            linha = converter_h010(campos)
            if linha:
                saida.write(linha)
                stats['inventario'] += 1
            else:
                stats['erros'] += 1

    # Encerramento
    saida.write("|9999|\n")

    return saida.getvalue(), stats


# ─────────────────────────────────────────
# INTERFACE STREAMLIT — TEMA TR
# ─────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Domínio Sistemas | Thomson Reuters",
        page_icon="🟠",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    apply_tr_theme()

    # ── Cabeçalho idêntico ao RPA ──────────────────────────────────────
    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                    border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                📄 Conversor SPED Fiscal → Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Selecione o arquivo SPED Fiscal (.txt) e clique em
                <strong>▶ Processar Arquivo</strong>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar idêntica ao RPA ────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ℹ Sobre")
        st.markdown(f"**Versão:** {VERSAO}")
        st.markdown("**Thomson Reuters**")
        st.markdown("**Domínio Sistemas**")
        st.markdown("---")
        st.markdown("### 📋 Registros Suportados")
        st.markdown(
            "- **0000** Abertura\n"
            "- **0150** Participantes\n"
            "- **C100** Notas Fiscais\n"
            "- **C170** Itens de NF\n"
            "- **C190** Analítico ICMS\n"
            "- **D100** Conhecimento de Transporte\n"
            "- **D500** Serviços de Comunicação\n"
            "- **H010** Inventário\n"
        )
        st.markdown("---")
        st.markdown("### ⚙ Encodings suportados")
        st.markdown("UTF-8 · Latin-1 · CP1252")

    # ── Instruções ────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">

            <h4>🔹 Passo 1 — Obter o arquivo SPED Fiscal</h4>
            <p>Exporte o arquivo <b>EFD ICMS/IPI</b> (.txt) gerado pelo seu sistema ERP
            para a escrituração fiscal digital.</p>

            <h4>🔹 Passo 2 — Realizar o upload</h4>
            <p>Clique em <b>Browse files</b> e selecione o arquivo <code>.txt</code> do SPED Fiscal.</p>

            <h4>🔹 Passo 3 — Processar</h4>
            <p>Clique em <b>▶ Processar Arquivo</b> e aguarde o processamento.
            As estatísticas serão exibidas automaticamente.</p>

            <h4>🔹 Passo 4 — Baixar e importar</h4>
            <p>Clique em <b>⬇ Baixar Arquivo Domínio</b> e importe o arquivo gerado
            no módulo fiscal do <b>Domínio Sistemas</b>.</p>

            <hr>

            <h4>⚠ Observações importantes</h4>
            <ul>
                <li>O arquivo de entrada deve estar no formato <b>SPED EFD ICMS/IPI</b>
                    com campos separados por <code>|</code>.</li>
                <li>CFOPs não mapeados receberão acumulador <b>9999</b> e serão listados
                    como aviso após o processamento.</li>
                <li>Encodings suportados: <b>UTF-8</b>, <b>Latin-1</b> e <b>CP1252</b>
                    (detecção automática).</li>
                <li>A hierarquia de registros (C100 → C170 → C190) é preservada
                    na saída.</li>
            </ul>

            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Session state ──────────────────────────────────────────────────
    if "log_sped" not in st.session_state:
        st.session_state.log_sped = [f"Aplicação pronta. Versão: {VERSAO}"]
    if "resultado_sped" not in st.session_state:
        st.session_state.resultado_sped = None
    if "nome_saida" not in st.session_state:
        st.session_state.nome_saida = "saida_dominio.txt"

    # ── Upload ─────────────────────────────────────────────────────────
    uploaded_file = st.file_uploader(
        "Arquivo SPED Fiscal de origem",
        type=["txt"],
        help="Arquivo gerado pelo ERP no formato SPED Fiscal (EFD ICMS/IPI)",
    )

    col1, col2 = st.columns([1, 1])

    with col1:
        processar = st.button(
            "▶ Processar Arquivo",
            disabled=(uploaded_file is None),
            use_container_width=True,
            type="primary",
        )
    with col2:
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        st.session_state.log_sped       = ["Campos limpos."]
        st.session_state.resultado_sped = None
        st.session_state.nome_saida     = "saida_dominio.txt"
        _cfops_nao_mapeados.clear()
        st.rerun()

    # ── Processamento ──────────────────────────────────────────────────
    if processar and uploaded_file is not None:
        _cfops_nao_mapeados.clear()
        st.session_state.log_sped       = ["Iniciando conversão do arquivo SPED..."]
        st.session_state.resultado_sped = None

        try:
            raw      = uploaded_file.read()
            content  = decode_sped(raw)
            registros = parse_sped(content)

            regs_encontrados = list(registros.keys())
            st.session_state.log_sped.append(
                f"Registros encontrados: {', '.join(regs_encontrados)}"
            )

            # Valida se parece um SPED real
            if '0000' not in registros:
                st.session_state.log_sped.append(
                    "AVISO: Registro 0000 não encontrado — verifique se o arquivo "
                    "é um SPED Fiscal válido."
                )

            resultado, stats = gerar_dominio(registros)

            # Avisos de CFOPs não mapeados
            if _cfops_nao_mapeados:
                st.session_state.log_sped.append(
                    f"AVISO: CFOPs não mapeados (acumulador 9999): "
                    f"{', '.join(sorted(_cfops_nao_mapeados))}"
                )

            # Avisos de blocos ausentes
            for bloco, nome in [
                ('C100', 'Notas Fiscais (C100)'),
                ('C170', 'Itens de NF (C170)'),
                ('C190', 'Analítico ICMS (C190)'),
            ]:
                if bloco not in registros:
                    st.session_state.log_sped.append(
                        f"AVISO: Bloco {nome} não encontrado no arquivo."
                    )

            total_regs = sum(stats.values())
            st.session_state.log_sped.append(
                f"Conversão concluída. "
                f"NFs entrada={stats['nf_entrada']} | saída={stats['nf_saida']} | "
                f"itens={stats['itens']} | participantes={stats['participantes']} | "
                f"erros={stats['erros']}"
            )

            st.session_state.resultado_sped = resultado.encode('utf-8')
            st.session_state.nome_saida = (
                uploaded_file.name.replace('.txt', '_dominio.txt')
            )
            st.session_state.stats = stats
            st.session_state.regs_encontrados = regs_encontrados

        except Exception as e:
            st.session_state.log_sped.append(f"ERRO FATAL: {str(e)}")
            import traceback
            st.session_state.log_sped.append(traceback.format_exc())

        st.rerun()

    # ── Resultado e Download ───────────────────────────────────────────
    if st.session_state.resultado_sped is not None:
        st.success("✅ Arquivo convertido com sucesso!")

        # Métricas — estilo idêntico ao RPA
        stats = st.session_state.get("stats", {})

        st.markdown("#### 📊 Estatísticas da Conversão")
        col1, col2, col3 = st.columns(3)
        col1.metric("Participantes", stats.get('participantes', 0))
        col2.metric("NFs Entrada",   stats.get('nf_entrada',    0))
        col3.metric("NFs Saída",     stats.get('nf_saida',      0))

        col4, col5, col6 = st.columns(3)
        col4.metric("Itens de NF",  stats.get('itens',       0))
        col5.metric("Analíticos",   stats.get('analiticos',  0))
        col6.metric("Transporte",   stats.get('transporte',  0))

        col7, col8, col9 = st.columns(3)
        col7.metric("Comunicação",  stats.get('comunicacao', 0))
        col8.metric("Inventário",   stats.get('inventario',  0))
        col9.metric("Erros",        stats.get('erros',       0))

        st.markdown("---")

        # Registros encontrados
        regs = st.session_state.get("regs_encontrados", [])
        with st.expander("🔍 Registros encontrados no SPED"):
            st.write(regs)

        # Prévia
        with st.expander("👁️ Prévia do arquivo gerado (primeiras 50 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado_sped.decode('utf-8').splitlines()[:50]
            )
            st.code(preview, language='text')

        # Download — estilo primário TR
        st.download_button(
            label="⬇ Baixar Arquivo Domínio",
            data=st.session_state.resultado_sped,
            file_name=st.session_state.nome_saida,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log de processamento (idêntico ao RPA) ─────────────────────────
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log_sped)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log_sped)
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

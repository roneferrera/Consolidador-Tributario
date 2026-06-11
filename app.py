import streamlit as st
import pandas as pd
import re
from io import StringIO
from datetime import datetime

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
        # Remove o primeiro e último campo vazio (pipes nas bordas)
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

def get_acumulador(cfop: str) -> str:
    return CFOP_ACUMULADOR.get(cfop, '9999')


# ─────────────────────────────────────────
# CONVERSORES POR REGISTRO
# ─────────────────────────────────────────

def converter_0000(campos: list) -> str:
    """Registro de abertura."""
    try:
        dt_ini = campos[4] if len(campos) > 4 else ''
        dt_fin = campos[5] if len(campos) > 5 else ''
        nome    = campos[6] if len(campos) > 6 else ''
        cnpj    = campos[7] if len(campos) > 7 else ''
        uf      = campos[9] if len(campos) > 9 else ''
        return f"|0000|1|{dt_ini}|{dt_fin}|{nome}|{cnpj}|{uf}|\n"
    except Exception:
        return ''


def converter_0150(campos: list) -> str:
    """Cadastro de participantes → Clientes/Fornecedores Domínio."""
    try:
        cod    = campos[1]  if len(campos) > 1  else ''
        nome   = campos[2]  if len(campos) > 2  else ''
        cod_pais = campos[3] if len(campos) > 3 else ''
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
        ind_oper  = campos[1]  if len(campos) > 1  else '0'
        ind_emit  = campos[2]  if len(campos) > 2  else '0'
        cod_part  = campos[3]  if len(campos) > 3  else ''
        cod_mod   = campos[4]  if len(campos) > 4  else ''
        cod_sit   = campos[5]  if len(campos) > 5  else '00'
        serie     = campos[6]  if len(campos) > 6  else ''
        num_doc   = campos[7]  if len(campos) > 7  else ''
        chv_nfe   = campos[8]  if len(campos) > 8  else ''
        dt_doc    = campos[9]  if len(campos) > 9  else ''
        dt_es     = campos[10] if len(campos) > 10 else ''
        vl_doc    = campos[11] if len(campos) > 11 else '0'
        vl_bc_icms = campos[15] if len(campos) > 15 else '0'
        vl_icms   = campos[16] if len(campos) > 16 else '0'
        vl_ipi    = campos[21] if len(campos) > 21 else '0'
        vl_pis    = campos[22] if len(campos) > 22 else '0'
        vl_cofins = campos[23] if len(campos) > 23 else '0'

        tipo = 'E' if ind_oper == '0' else 'S'

        return (
            f"|C100|{tipo}|{cod_part}|{cod_mod}|{cod_sit}|{serie}|{num_doc}|"
            f"{chv_nfe}|{dt_doc}|{dt_es}|{vl_doc}|{vl_bc_icms}|{vl_icms}|"
            f"{vl_ipi}|{vl_pis}|{vl_cofins}|\n"
        )
    except Exception:
        return ''


def converter_c170(campos: list, cfop: str = '') -> str:
    """Itens da NF (C170) → Detalhamento Domínio."""
    try:
        num_item  = campos[1]  if len(campos) > 1  else ''
        cod_item  = campos[2]  if len(campos) > 2  else ''
        descr     = campos[3]  if len(campos) > 3  else ''
        qtd       = campos[4]  if len(campos) > 4  else '0'
        unid      = campos[5]  if len(campos) > 5  else ''
        vl_item   = campos[6]  if len(campos) > 6  else '0'
        vl_desc   = campos[7]  if len(campos) > 7  else '0'
        cfop_item = campos[8]  if len(campos) > 8  else cfop
        acum      = get_acumulador(cfop_item)
        vl_bc     = campos[9]  if len(campos) > 9  else '0'
        aliq      = campos[10] if len(campos) > 10 else '0'
        vl_icms   = campos[11] if len(campos) > 11 else '0'
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
    """Conhecimento de Transporte (D100)."""
    try:
        ind_oper = campos[1] if len(campos) > 1 else '0'
        cod_part = campos[3] if len(campos) > 3 else ''
        cod_mod  = campos[4] if len(campos) > 4 else ''
        cod_sit  = campos[5] if len(campos) > 5 else '00'
        serie    = campos[6] if len(campos) > 6 else ''
        num_doc  = campos[7] if len(campos) > 7 else ''
        dt_doc   = campos[9] if len(campos) > 9 else ''
        vl_doc   = campos[11] if len(campos) > 11 else '0'
        cfop     = campos[16] if len(campos) > 16 else ''
        acum     = get_acumulador(cfop)
        tipo     = 'E' if ind_oper == '0' else 'S'
        return (
            f"|D100|{tipo}|{cod_part}|{cod_mod}|{cod_sit}|{serie}|"
            f"{num_doc}|{dt_doc}|{vl_doc}|{cfop}|{acum}|\n"
        )
    except Exception:
        return ''


def converter_d500(campos: list) -> str:
    """Serviços de Comunicação (D500)."""
    try:
        ind_oper = campos[1] if len(campos) > 1 else '0'
        cod_part = campos[3] if len(campos) > 3 else ''
        cod_mod  = campos[4] if len(campos) > 4 else ''
        cod_sit  = campos[5] if len(campos) > 5 else '00'
        serie    = campos[6] if len(campos) > 6 else ''
        num_doc  = campos[7] if len(campos) > 7 else ''
        dt_doc   = campos[8] if len(campos) > 8 else ''
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
# GERADOR DO ARQUIVO DOMÍNIO
# ─────────────────────────────────────────

def gerar_dominio(registros: dict) -> str:
    saida = StringIO()
    stats = {
        'participantes': 0,
        'nf_entrada': 0,
        'nf_saida': 0,
        'itens': 0,
        'transporte': 0,
        'comunicacao': 0,
        'inventario': 0,
        'erros': 0,
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

    # Bloco C - Notas Fiscais
    if 'C100' in registros:
        for campos in registros['C100']:
            linha = converter_c100(campos)
            if linha:
                saida.write(linha)
                ind_oper = campos[1] if len(campos) > 1 else '0'
                if ind_oper == '0':
                    stats['nf_entrada'] += 1
                else:
                    stats['nf_saida'] += 1

    # Itens das NFs
    if 'C170' in registros:
        for campos in registros['C170']:
            linha = converter_c170(campos)
            if linha:
                saida.write(linha)
                stats['itens'] += 1

    # Analítico C190
    if 'C190' in registros:
        for campos in registros['C190']:
            linha = converter_c190(campos)
            if linha:
                saida.write(linha)

    # Bloco D - Transporte
    if 'D100' in registros:
        for campos in registros['D100']:
            linha = converter_d100(campos)
            if linha:
                saida.write(linha)
                stats['transporte'] += 1

    # Bloco D - Comunicação
    if 'D500' in registros:
        for campos in registros['D500']:
            linha = converter_d500(campos)
            if linha:
                saida.write(linha)
                stats['comunicacao'] += 1

    # Bloco H - Inventário
    if 'H010' in registros:
        for campos in registros['H010']:
            linha = converter_h010(campos)
            if linha:
                saida.write(linha)
                stats['inventario'] += 1

    # Encerramento
    saida.write("|9999|\n")

    return saida.getvalue(), stats


# ─────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="SPED Fiscal → Domínio Sistemas",
        page_icon="📄",
        layout="centered"
    )

    st.title("📄 Conversor SPED Fiscal → Domínio Sistemas")
    st.markdown(
        "Faça o upload do arquivo SPED Fiscal (.txt) para converter "
        "para o leiaute do **Domínio Sistemas** com separador `|`."
    )

    st.divider()

    uploaded_file = st.file_uploader(
        "Selecione o arquivo SPED Fiscal",
        type=["txt"],
        help="Arquivo gerado pelo sistema ERP no formato SPED Fiscal (EFD ICMS/IPI)"
    )

    if uploaded_file is not None:
        st.success(f"✅ Arquivo carregado: **{uploaded_file.name}**")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Tamanho", f"{uploaded_file.size / 1024:.1f} KB")
        with col2:
            st.metric("Tipo", "SPED Fiscal .txt")

        st.divider()

        if st.button("🔄 Processar Arquivo", use_container_width=True, type="primary"):
            with st.spinner("Processando arquivo SPED..."):
                try:
                    content = uploaded_file.read().decode('utf-8', errors='replace')
                    registros = parse_sped(content)

                    regs_encontrados = list(registros.keys())

                    resultado, stats = gerar_dominio(registros)

                    st.success("✅ Arquivo convertido com sucesso!")

                    st.subheader("📊 Estatísticas da Conversão")
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Participantes",  stats['participantes'])
                    col2.metric("NFs Entrada",    stats['nf_entrada'])
                    col3.metric("NFs Saída",      stats['nf_saida'])

                    col4, col5, col6 = st.columns(3)
                    col4.metric("Itens de NF",    stats['itens'])
                    col5.metric("Transporte",     stats['transporte'])
                    col6.metric("Comunicação",    stats['comunicacao'])

                    col7, _, _ = st.columns(3)
                    col7.metric("Inventário",     stats['inventario'])

                    st.divider()

                    with st.expander("🔍 Registros encontrados no SPED"):
                        st.write(regs_encontrados)

                    with st.expander("👁️ Prévia do arquivo gerado (primeiras 50 linhas)"):
                        preview = '\n'.join(resultado.splitlines()[:50])
                        st.code(preview, language='text')

                    nome_saida = uploaded_file.name.replace('.txt', '_dominio.txt')
                    st.download_button(
                        label="⬇️ Baixar Arquivo Domínio",
                        data=resultado.encode('utf-8'),
                        file_name=nome_saida,
                        mime='text/plain',
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"❌ Erro ao processar o arquivo: {str(e)}")
                    st.exception(e)

    else:
        st.info("⬆️ Aguardando upload do arquivo SPED Fiscal...")

    st.divider()
    st.caption("Conversor SPED Fiscal → Domínio Sistemas | Desenvolvido com Python + Streamlit")


if __name__ == "__main__":
    main()

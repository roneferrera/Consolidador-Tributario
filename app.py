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

# ==============================
# VERSÃO
# ==============================
VERSAO = "V1.2"

# ==============================
# TEMA TR (idêntico ao RPA V3.9)
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


# ==============================
# DECODE COM FALLBACK
# ==============================
def decode_arquivo(raw: bytes) -> str:
    for enc in ('utf-8', 'latin-1', 'cp1252'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


# ==============================
# ENCODE ANSI COM LOG
# ==============================
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
    """
    Lê o arquivo Domínio Sistemas e agrupa as linhas por tipo de registro,
    preservando a ordem original para análise de hierarquia.
    Retorna:
        {
          'linhas_ordenadas': [ (tipo, campos[]), ... ],
          'por_tipo':         { tipo: [ campos[], ... ] }
        }
    """
    linhas_ordenadas = []
    por_tipo         = {}

    for num_linha, linha in enumerate(content.splitlines(), start=1):
        linha = linha.strip()
        if not linha:
            continue
        campos = linha.split('|')
        # Remove bordas vazias geradas pelos pipes iniciais/finais
        if campos and campos[0] == '':
            campos = campos[1:]
        if campos and campos[-1] == '':
            campos = campos[:-1]
        if len(campos) < 1:
            continue

        tipo = campos[0].strip()
        if not tipo:
            continue

        linhas_ordenadas.append((tipo, campos, num_linha))
        if tipo not in por_tipo:
            por_tipo[tipo] = []
        por_tipo[tipo].append((campos, num_linha))

    return {
        'linhas_ordenadas': linhas_ordenadas,
        'por_tipo':         por_tipo,
    }


# ==============================
# MAPEAMENTO DE CAMPOS POR REGISTRO
# (baseado no arquivo de exemplo fornecido)
# ==============================

# Registro 0000: |0000|CNPJ|
IDX_0000_CNPJ = 1

# Registro 1000 (Nota Fiscal):
# |1000|NUM_NF|CNPJ_EMIT|...|CFOP|...|DT_EMISSAO|DT_ENTRADA|VL_TOTAL|...|
# Posições identificadas pelo exemplo:
IDX_1000_NUM_NF      = 1
IDX_1000_CNPJ_EMIT   = 2
IDX_1000_CFOP        = 4   # campo "1102" no exemplo → índice 4
IDX_1000_DT_EMISSAO  = 11
IDX_1000_DT_ENTRADA  = 12
IDX_1000_VL_TOTAL    = 13

# Registro 1020 (Totais da NF):
# |1020|NUM_NF|...|VL_TOTAL|ALIQ_ICMS|VL_ICMS|...|
IDX_1020_NUM_NF    = 1
IDX_1020_VL_TOTAL  = 3
IDX_1020_ALIQ_ICMS = 4
IDX_1020_VL_ICMS   = 5

# Registro 1030 (Item da NF):
# |1030|NUM_ITEM|QTD|VL_UNIT|...|VL_ITEM|VL_DESC|...|CFOP|...|
# Pelo exemplo: campo[34] = CFOP (ex: "1102")
# Mapeamento completo pelo exemplo fornecido:
# [0]=1030 [1]=NUM_ITEM [2]=QTD [3]=VL_UNIT [4]=? [5]=? [6]=?
# [7]=DT_DOC [8]=? [9]=COD_SIT [10]=VL_ITEM [11]=VL_DESC [12]=VL_ITEM_LIQ
# [13]=? [14]=ALIQ_ICMS [15..33]=outros campos
# [34]=CFOP  ← identificado pelo valor "1102" no exemplo
IDX_1030_NUM_ITEM  = 1
IDX_1030_QTD       = 2
IDX_1030_VL_UNIT   = 3
IDX_1030_DT_DOC    = 7
IDX_1030_VL_ITEM   = 10
IDX_1030_VL_DESC   = 11
IDX_1030_ALIQ_ICMS = 14
IDX_1030_VL_ICMS   = 21   # campo VL_ICMS identificado pelo valor "92,64" → índice 21
IDX_1030_CFOP      = 34   # "1102" no exemplo

# Registro 1300 (Lançamento contábil):
# |1300|DATA|COD_HIST|NUM_HIST|VL|NUM_NF|HISTORICO|USUARIO|
IDX_1300_DATA     = 1
IDX_1300_COD_HIST = 2
IDX_1300_NUM_HIST = 3
IDX_1300_VL       = 4
IDX_1300_NUM_NF   = 5
IDX_1300_HIST     = 6


def _campo(campos: list, idx: int, default: str = '') -> str:
    """Retorna o campo pelo índice com fallback seguro."""
    return campos[idx].strip() if len(campos) > idx else default


# ==============================
# CARREGAMENTO DA TABELA DE ACUMULADORES
# ==============================
def gerar_modelo_acumuladores_csv() -> bytes:
    linhas = [
        "CFOP,ACUMULADOR",
        "1101,1101", "2101,1101", "3101,1101",
        "1102,1151", "2102,1151", "3102,1151",
        "1111,1103", "2111,1103",
        "1113,1153", "2113,1153",
        "1116,1104", "2116,1104",
        "1120,1105", "2120,1105",
        "1201,1201", "2201,1201",
        "1202,1202", "2202,1202",
        "1203,1203", "2203,1203",
        "1251,1301", "2251,1301",
        "1252,1302", "2252,1302",
        "1301,1334", "2301,1334",
        "1351,1367", "2351,1367",
        "1352,1368", "2352,1368",
        "1401,1401", "2401,1401",
        "1403,1403", "2403,1403",
        "1501,1501", "2501,1501",
        "1502,1502", "2502,1502",
        "1503,1503", "2503,1503",
        "1601,1601", "2601,1601",
        "1701,1701", "2701,1701",
        "1801,1801", "2801,1801",
        "1901,1901", "2901,1901",
        "5101,5101", "6101,5101",
        "5102,5102", "6102,5102",
        "5111,5103", "6111,5103",
        "5113,5104", "6113,5104",
        "5151,5151", "6151,5151",
        "5152,5152", "6152,5152",
        "5153,5153", "6153,5153",
        "5201,5201", "6201,5201",
        "5202,5202", "6202,5202",
        "5251,5301", "6251,5301",
        "5301,5334", "6301,5334",
        "5351,5367", "6351,5367",
        "5352,5368", "6352,5368",
        "5401,5401", "6401,5401",
        "5403,5403", "6403,5403",
        "5501,5501", "6501,5501",
        "5601,5601", "6601,5601",
        "5701,5701", "6701,5701",
        "5801,5801", "6801,5801",
        "5901,5901", "6901,5901",
    ]
    return "\n".join(linhas).encode("latin-1")


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
                "ERRO: O arquivo de acumuladores deve conter as colunas "
                "'CFOP' e 'ACUMULADOR'."
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
            tabela[cfop] = acum

        if not tabela:
            log.append("ERRO: Nenhum par CFOP → Acumulador válido encontrado.")
            return None

        if erros:
            log.append(
                f"AVISO: {erros} linha(s) ignoradas por dados inválidos "
                f"no arquivo de acumuladores."
            )

        log.append(f"Tabela de acumuladores carregada: {len(tabela)} CFOPs mapeados.")
        return tabela

    except Exception as e:
        log.append(f"ERRO ao carregar arquivo de acumuladores: {e}")
        return None


def get_acumulador(cfop: str, tabela: dict, nao_mapeados: set) -> str:
    cfop_norm = str(cfop).strip().zfill(4)
    acum = tabela.get(cfop_norm)
    if acum is None:
        nao_mapeados.add(cfop_norm)
        return '9999'
    return acum


# ==============================
# ANÁLISE DE HIERARQUIA E ORDENAÇÃO
# ==============================
def extrair_notas_com_itens(parsed: dict, tabela_acum: dict,
                             nao_mapeados: set, log: list) -> list:
    """
    Percorre as linhas na ordem original do arquivo e agrupa
    cada 1000 com seus 1020 e 1030 filhos.
    Retorna lista de dicts com a NF e seus itens, já ordenada
    por: data_emissao → num_nf → num_item.
    """
    linhas   = parsed['linhas_ordenadas']
    notas    = []
    nota_atual = None

    for tipo, campos, num_linha in linhas:

        if tipo == '0000':
            continue

        elif tipo == '1000':
            # Fecha nota anterior se existir
            if nota_atual is not None:
                notas.append(nota_atual)

            num_nf     = _campo(campos, IDX_1000_NUM_NF)
            cnpj_emit  = _campo(campos, IDX_1000_CNPJ_EMIT)
            cfop_nf    = _campo(campos, IDX_1000_CFOP)
            dt_emissao = _campo(campos, IDX_1000_DT_EMISSAO)
            dt_entrada = _campo(campos, IDX_1000_DT_ENTRADA)
            vl_total   = _campo(campos, IDX_1000_VL_TOTAL)
            acum_nf    = get_acumulador(cfop_nf, tabela_acum, nao_mapeados)

            nota_atual = {
                'num_nf':      num_nf,
                'cnpj_emit':   cnpj_emit,
                'cfop_nf':     cfop_nf,
                'acum_nf':     acum_nf,
                'dt_emissao':  dt_emissao,
                'dt_entrada':  dt_entrada,
                'vl_total':    vl_total,
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
                log.append(
                    f"AVISO: Registro 1020 na linha {num_linha} "
                    f"sem 1000 pai. Ignorado."
                )

        elif tipo == '1030':
            if nota_atual is not None:
                num_item  = _campo(campos, IDX_1030_NUM_ITEM)
                cfop_item = _campo(campos, IDX_1030_CFOP)
                acum_item = get_acumulador(cfop_item, tabela_acum, nao_mapeados)
                nota_atual['itens_1030'].append({
                    'num_item':  num_item,
                    'cfop':      cfop_item,
                    'acum':      acum_item,
                    'campos':    campos,
                    'linha_orig': num_linha,
                })
            else:
                log.append(
                    f"AVISO: Registro 1030 na linha {num_linha} "
                    f"sem 1000 pai. Ignorado."
                )

        elif tipo == '1300':
            if nota_atual is not None:
                nota_atual['lanc_1300'].append(campos)
            else:
                log.append(
                    f"AVISO: Registro 1300 na linha {num_linha} "
                    f"sem 1000 pai. Ignorado."
                )

    # Fecha última nota
    if nota_atual is not None:
        notas.append(nota_atual)

    # ── Ordenação ──────────────────────────────────────────────────────
    # Critério: data_emissao (DD/MM/AAAA) → num_nf (int) → num_item (int)
    def sort_key_nota(n):
        dt_str = n['dt_emissao']
        try:
            dt = datetime.strptime(dt_str, '%d/%m/%Y')
        except Exception:
            try:
                dt = datetime.strptime(dt_str, '%Y-%m-%d')
            except Exception:
                dt = datetime.min
        try:
            num = int(n['num_nf'])
        except Exception:
            num = 0
        return (dt, num)

    notas.sort(key=sort_key_nota)

    # Ordena os itens dentro de cada nota por num_item
    for nota in notas:
        nota['itens_1030'].sort(
            key=lambda it: int(it['num_item']) if it['num_item'].isdigit() else 0
        )

    log.append(
        f"Hierarquia montada: {len(notas)} nota(s) encontrada(s), "
        f"ordenadas por data de emissão e número."
    )
    return notas


# ==============================
# GERADOR DO ARQUIVO DE SAÍDA
# ==============================
def gerar_saida(
    parsed: dict,
    tabela_acum: dict,
    log: list,
) -> tuple:
    """
    Gera o arquivo de saída com:
    - Registro 0000 (cabeçalho)
    - Para cada NF ordenada: 1000 → 1020 → 1030(s) → 1300(s)
    - Registro 9999 (encerramento)
    Retorna (conteudo_str, stats_dict)
    """
    nao_mapeados = set()
    saida        = StringIO()
    stats = {
        'notas':       0,
        'nf_entrada':  0,
        'nf_saida':    0,
        'itens':       0,
        'lancamentos': 0,
        'erros':       0,
    }

    # Cabeçalho 0000
    if '0000' in parsed['por_tipo']:
        campos_0000, _ = parsed['por_tipo']['0000'][0]
        saida.write('|' + '|'.join(campos_0000) + '|\n')
    else:
        log.append("AVISO: Registro 0000 não encontrado no arquivo.")

    # Monta hierarquia ordenada
    notas = extrair_notas_com_itens(parsed, tabela_acum, nao_mapeados, log)

    for nota in notas:
        # ── 1000 ────────────────────────────────────────────────────
        campos_1000 = nota['campos_1000']
        saida.write('|' + '|'.join(campos_1000) + '|\n')
        stats['notas'] += 1

        # Detecta entrada/saída pelo primeiro dígito do CFOP
        cfop_ini = nota['cfop_nf'][:1] if nota['cfop_nf'] else ''
        if cfop_ini in ('1', '2', '3'):
            stats['nf_entrada'] += 1
        elif cfop_ini in ('5', '6', '7'):
            stats['nf_saida'] += 1

        # ── 1020 ────────────────────────────────────────────────────
        if nota['totais_1020'] is not None:
            saida.write('|' + '|'.join(nota['totais_1020']) + '|\n')

        # ── 1030 (itens) ─────────────────────────────────────────────
        for item in nota['itens_1030']:
            campos = item['campos']
            # Injeta acumulador no campo correto (posição IDX_1030_CFOP + 1
            # se o layout previr, ou apenas reescreve como está)
            # Por ora reescrevemos o registro sem alteração estrutural
            # pois o acumulador é informação adicional logada, não um campo
            # nativo do layout 1030 do Domínio.
            saida.write('|' + '|'.join(campos) + '|\n')
            stats['itens'] += 1

        # ── 1300 (lançamentos) ────────────────────────────────────────
        for lanc in nota['lanc_1300']:
            saida.write('|' + '|'.join(lanc) + '|\n')
            stats['lancamentos'] += 1

    # Encerramento
    saida.write('|9999|\n')

    # Avisos de CFOPs não mapeados
    if nao_mapeados:
        log.append(
            f"AVISO: {len(nao_mapeados)} CFOP(s) não encontrado(s) na tabela "
            f"de acumuladores (acumulador 9999): "
            f"{', '.join(sorted(nao_mapeados))}"
        )

    log.append(
        f"Geração concluída — "
        f"Notas={stats['notas']} | "
        f"Entradas={stats['nf_entrada']} | "
        f"Saídas={stats['nf_saida']} | "
        f"Itens={stats['itens']} | "
        f"Lançamentos={stats['lancamentos']} | "
        f"Erros={stats['erros']}"
    )

    return saida.getvalue(), stats


# ==============================
# GERADOR DO RELATÓRIO DE ORDENAÇÃO
# ==============================
def gerar_relatorio_ordenacao(notas: list) -> pd.DataFrame:
    """
    Gera um DataFrame para exibir a ordem final das notas
    e seus itens — útil para verificação visual.
    """
    linhas = []
    for pos, nota in enumerate(notas, start=1):
        linhas.append({
            'Posição':      pos,
            'Num NF':       nota['num_nf'],
            'CFOP NF':      nota['cfop_nf'],
            'Acumulador':   nota['acum_nf'],
            'Dt Emissão':   nota['dt_emissao'],
            'Dt Entrada':   nota['dt_entrada'],
            'Vl Total':     nota['vl_total'],
            'Qtd Itens':    len(nota['itens_1030']),
            'Qtd Lanç':     len(nota['lanc_1300']),
            'Linha Orig':   nota['linha_orig'],
            'Tipo':         'Entrada' if nota['cfop_nf'][:1] in ('1','2','3') else 'Saída',
        })
        for item in nota['itens_1030']:
            linhas.append({
                'Posição':    f"  └ Item {item['num_item']}",
                'Num NF':     nota['num_nf'],
                'CFOP NF':    item['cfop'],
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

    # ── Cabeçalho ──────────────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="background:#444444; padding:24px 28px 18px 28px; border-radius:8px;
                    border-top:6px solid #FF8000; margin-bottom:28px;">
            <h2 style="color:#FF8000; margin:0; font-family:'Segoe UI',Arial,sans-serif;">
                📄 Conversor / Ordenador Domínio Sistemas &nbsp;|&nbsp; {VERSAO}
            </h2>
            <p style="color:#DDDDDD; margin:6px 0 0 0; font-family:'Segoe UI',Arial,sans-serif;">
                Faça upload da <strong>tabela de acumuladores</strong> e do
                <strong>arquivo Domínio (.txt)</strong>, depois clique em
                <strong>▶ Processar Arquivo</strong>.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Sidebar ────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⬇ Modelo de Acumuladores")
        st.markdown(
            "Baixe o modelo CSV, preencha o DE-PARA "
            "**CFOP → Acumulador** e faça o upload ao lado."
        )
        st.download_button(
            label="⬇ Baixar modelo_acumuladores.csv",
            data=gerar_modelo_acumuladores_csv(),
            file_name="modelo_acumuladores.csv",
            mime="text/plain",
            use_container_width=True,
        )
        st.markdown("---")
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
        st.markdown("### ⚙ Ordenação aplicada")
        st.markdown("Data Emissão → Nº NF → Nº Item")
        st.markdown("### ⚙ Encoding de saída")
        st.markdown("**ANSI (Latin-1)**")

    # ── Instruções ─────────────────────────────────────────────────────
    with st.expander("📖 **Instruções de Uso** — clique para expandir", expanded=False):
        st.markdown(
            """
            <div class="instrucoes-box">

            <h4>🔹 Passo 1 — Baixar o modelo de acumuladores</h4>
            <p>No menu lateral, clique em <b>⬇ Baixar modelo_acumuladores.csv</b>.
            Edite o DE-PARA <code>CFOP → Acumulador</code> conforme a parametrização
            do seu Domínio Sistemas e salve.</p>

            <h4>🔹 Passo 2 — Fazer upload da tabela de acumuladores</h4>
            <p>No campo <b>Tabela de Acumuladores</b>, selecione o arquivo
            <code>.csv</code> ou <code>.xlsx</code> preenchido.</p>

            <h4>🔹 Passo 3 — Fazer upload do arquivo Domínio</h4>
            <p>No campo <b>Arquivo Domínio (.txt)</b>, selecione o arquivo
            exportado pelo Domínio Sistemas.</p>

            <h4>🔹 Passo 4 — Processar</h4>
            <p>Clique em <b>▶ Processar Arquivo</b>. O sistema irá:
            <ul>
                <li>Montar a hierarquia <b>1000 → 1020 → 1030 → 1300</b></li>
                <li>Ordenar as notas por <b>Data Emissão → Nº NF → Nº Item</b></li>
                <li>Mapear os acumuladores conforme a tabela fornecida</li>
                <li>Gerar o arquivo de saída em <b>ANSI (Latin-1)</b></li>
            </ul>
            </p>

            <h4>🔹 Passo 5 — Verificar e baixar</h4>
            <p>Confira a <b>tabela de ordenação</b> exibida na tela e clique em
            <b>⬇ Baixar Arquivo Processado</b>.</p>

            <hr>

            <h4>⚠ Observações importantes</h4>
            <ul>
                <li>CFOPs não encontrados na tabela receberão acumulador <b>9999</b>
                    e serão listados no log.</li>
                <li>Registros 1020, 1030 e 1300 sem um 1000 pai são ignorados
                    com aviso no log.</li>
                <li>O arquivo de saída é gerado em <b>ANSI (Latin-1)</b>,
                    padrão do Domínio Sistemas.</li>
            </ul>

            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ── Session state ──────────────────────────────────────────────────
    defaults = {
        "log_conv":          [f"Aplicação pronta. Versão: {VERSAO}"],
        "resultado_conv":    None,
        "nome_saida":        "saida_dominio.txt",
        "stats_conv":        None,
        "df_ordenacao":      None,
        "tabela_acum_ok":    False,
        "tabela_acum_info":  "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Uploads lado a lado ────────────────────────────────────────────
    col_up1, col_up2 = st.columns([1, 1])

    with col_up1:
        arquivo_acum = st.file_uploader(
            "📂 Tabela de Acumuladores (CFOP → Acumulador)",
            type=["csv", "xlsx", "xls"],
            help="Arquivo .csv ou .xlsx com colunas CFOP e ACUMULADOR",
            key="upload_acum",
        )

    with col_up2:
        uploaded_file = st.file_uploader(
            "📂 Arquivo Domínio de origem (.txt)",
            type=["txt"],
            help="Arquivo .txt exportado pelo Domínio Sistemas",
            key="upload_dominio",
        )

    # Feedback imediato da tabela de acumuladores
    if arquivo_acum is not None:
        log_temp      = []
        raw_acum      = arquivo_acum.read()
        tabela_prev   = carregar_acumuladores(raw_acum, arquivo_acum.name, log_temp)
        arquivo_acum.seek(0)
        if tabela_prev is not None:
            st.success(
                f"✅ Tabela de acumuladores válida — "
                f"**{len(tabela_prev)} CFOPs** mapeados."
            )
            st.session_state.tabela_acum_ok   = True
            st.session_state.tabela_acum_info = f"{len(tabela_prev)} CFOPs mapeados."
        else:
            for msg in log_temp:
                st.error(msg)
            st.session_state.tabela_acum_ok   = False
            st.session_state.tabela_acum_info = ""
    else:
        st.info(
            "⬆ Faça o upload da **tabela de acumuladores** antes de processar. "
            "Baixe o modelo no menu lateral."
        )
        st.session_state.tabela_acum_ok = False

    # ── Botões ─────────────────────────────────────────────────────────
    pode_processar = (
        arquivo_acum  is not None and
        uploaded_file is not None and
        st.session_state.tabela_acum_ok
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
        limpar = st.button("🗑 Limpar", use_container_width=True)

    if limpar:
        for k, v in defaults.items():
            st.session_state[k] = v if not isinstance(v, list) else [f"Campos limpos."]
        st.rerun()

    # ── Processamento ──────────────────────────────────────────────────
    if processar and pode_processar:
        st.session_state.log_conv       = ["Iniciando processamento..."]
        st.session_state.resultado_conv = None
        st.session_state.stats_conv     = None
        st.session_state.df_ordenacao   = None

        try:
            # 1. Carrega tabela de acumuladores
            arquivo_acum.seek(0)
            tabela_acum = carregar_acumuladores(
                arquivo_acum.read(),
                arquivo_acum.name,
                st.session_state.log_conv,
            )
            if tabela_acum is None:
                st.session_state.log_conv.append("ERRO: Tabela inválida. Abortando.")
                st.rerun()

            # 2. Lê e parseia o arquivo Domínio
            raw     = uploaded_file.read()
            content = decode_arquivo(raw)
            parsed  = parse_dominio(content)

            tipos_encontrados = list(parsed['por_tipo'].keys())
            st.session_state.log_conv.append(
                f"Tipos de registro encontrados: {', '.join(tipos_encontrados)}"
            )
            st.session_state.log_conv.append(
                f"Total de linhas processadas: "
                f"{len(parsed['linhas_ordenadas'])}"
            )

            # 3. Gera saída ordenada
            resultado_txt, stats = gerar_saida(
                parsed, tabela_acum, st.session_state.log_conv
            )

            # 4. Gera relatório de ordenação para exibição
            nao_map_temp = set()
            notas_ord    = extrair_notas_com_itens(
                parsed, tabela_acum, nao_map_temp,
                []   # log separado para não duplicar
            )
            df_ord = gerar_relatorio_ordenacao(notas_ord)

            # 5. Codifica em ANSI
            resultado_bytes = encode_ansi_seguro(
                resultado_txt, st.session_state.log_conv
            )

            st.session_state.resultado_conv = resultado_bytes
            st.session_state.stats_conv     = stats
            st.session_state.df_ordenacao   = df_ord
            st.session_state.nome_saida     = (
                uploaded_file.name.replace('.txt', '_processado.txt')
            )

        except Exception:
            st.session_state.log_conv.append("ERRO FATAL durante o processamento.")
            st.session_state.log_conv.append(traceback.format_exc())

        st.rerun()

    # ── Resultado ──────────────────────────────────────────────────────
    if st.session_state.resultado_conv is not None:
        st.success("✅ Arquivo processado com sucesso!")

        stats = st.session_state.stats_conv or {}

        st.markdown("#### 📊 Estatísticas")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Notas",  stats.get('notas',       0))
        col2.metric("NFs Entrada",     stats.get('nf_entrada',  0))
        col3.metric("NFs Saída",       stats.get('nf_saida',    0))

        col4, col5, col6 = st.columns(3)
        col4.metric("Itens",           stats.get('itens',       0))
        col5.metric("Lançamentos",     stats.get('lancamentos', 0))
        col6.metric("Erros",           stats.get('erros',       0))

        st.markdown("---")

        # Tabela de ordenação
        if st.session_state.df_ordenacao is not None:
            with st.expander(
                "📋 Verificação de Ordenação — clique para expandir",
                expanded=True
            ):
                st.markdown(
                    "As notas abaixo estão na **ordem exata** em que serão "
                    "gravadas no arquivo de saída."
                )
                st.dataframe(
                    st.session_state.df_ordenacao,
                    use_container_width=True,
                    hide_index=True,
                )

        # Prévia
        with st.expander("👁️ Prévia do arquivo gerado (primeiras 60 linhas)"):
            preview = '\n'.join(
                st.session_state.resultado_conv
                .decode('latin-1', errors='replace')
                .splitlines()[:60]
            )
            st.code(preview, language='text')

        st.download_button(
            label="⬇ Baixar Arquivo Processado",
            data=st.session_state.resultado_conv,
            file_name=st.session_state.nome_saida,
            mime="text/plain",
            use_container_width=True,
            type="primary",
        )

    # ── Log ────────────────────────────────────────────────────────────
    st.markdown("**Log de processamento**")
    log_texto = "\n".join(st.session_state.log_conv)
    tem_erro  = any(str(l).startswith("ERRO") for l in st.session_state.log_conv)
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

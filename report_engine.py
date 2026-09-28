import glob
import json
import logging
import unicodedata
import datetime as dt
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================
#  CONFIGURACAO - edite aqui
# =============================================================
BASE = Path(__file__).parent

CONFIG = {
    "pasta_entrada":          BASE,
    "carteiras_file":         BASE / "Carteiras Gerenciadoras.xlsx",
    "aprovadas_file":         BASE / "PRIME - APROVADAS SEM MOVIMENTAÇÃO 01-06-2026.xlsx",
    "pasta_saida":            BASE / "relatorios_os",
    "lideres_ordem_os": [
        "Equipe DF", "Vinicius Raimundo Navakoski", "Jose Almy",
        "NF", "NF/Problema", "Willians Castanheira", "Paulo videira",
        "BOP", "Anderson de Oliveira", "Rafaela Rocha", "Cliente",
    ],
}
# =============================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

THIN   = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
C = {
    "dark_blue":  "1F3864", "mid_blue":   "2E75B6", "light_blue": "BDD7EE",
    "header_txt": "FFFFFF", "alt":        "EBF0FA", "red":        "FFC7CE",
    "green":      "C6EFCE", "saida_bg":   "00B050", "entrada_bg": "FF0000",
    "gray_tot":   "D9D9D9",    "yellow":     "FFEB9C",
}

def fill(color):                    return PatternFill("solid", fgColor=color)
def bfont(color="000000", size=10): return Font(name="Arial", bold=True, color=color, size=size)


# =============================================================
#  NORMALIZACAO (resolve variantes com/sem acento)
# =============================================================

def _norm(s):
    if not isinstance(s, str):
        return s
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode().strip().lower()

def _isin_norm(series, values):
    norm_set = {_norm(v) for v in values}
    return series.map(lambda x: _norm(x) in norm_set)

def _eq_norm(series, value):
    norm_val = _norm(value)
    return series.map(lambda x: _norm(x) == norm_val)

def _get_col_norm(df, target):
    """Retorna o nome real da coluna cujo nome normalizado coincide com target."""
    t = _norm(target)
    return next((c for c in df.columns if _norm(c) == t), None)


# =============================================================
#  SELECAO DO ARQUIVO DE OS
# =============================================================

def escolher_arquivo():
    padrao   = str(CONFIG["pasta_entrada"] / "*.xlsx")
    arquivos = sorted(glob.glob(padrao), key=lambda a: Path(a).stat().st_mtime, reverse=True)
    arquivos = [a for a in arquivos
                if not Path(a).name.startswith("OS_")
                and "Carteiras" not in Path(a).name
                and "APROVADAS" not in Path(a).name]

    if not arquivos:
        print("\nNenhum arquivo .xlsx de OS encontrado em:")
        print(CONFIG["pasta_entrada"])
        return None

    if len(arquivos) == 1:
        print(f"\nArquivo OS encontrado: {Path(arquivos[0]).name}")
        return arquivos[0]

    print("\nArquivos encontrados (mais recente primeiro):")
    for i, a in enumerate(arquivos, 1):
        mtime = dt.datetime.fromtimestamp(Path(a).stat().st_mtime).strftime("%d/%m/%Y %H:%M")
        print(f"  [{i}] {Path(a).name}  ({mtime})")

    while True:
        escolha = input("\nDigite o numero do arquivo (Enter = mais recente): ").strip()
        if escolha == "":
            return arquivos[0]
        if escolha.isdigit() and 1 <= int(escolha) <= len(arquivos):
            return arquivos[int(escolha) - 1]
        print("Opcao invalida.")


# =============================================================
#  CONSTANTES — OS (sem duplicatas acentuadas; _isin_norm cuida das variantes)
# =============================================================

STATUS_INDESEJADOS        = ["APROVADA", "CANCELADA", "SERVICO REJEITADO"]
AUTOGESTAO                = ["Auto Gestao", "Autogestao"]
ATENDIMENTOS_NF           = ["Equipe Especializada", "Validacao de NF", "Equipe Orcamentista"]
ABAS_CLIENTE              = [
    "Aguard. Aprovacao", "Aguard. Envio Veiculo",
    "NF Validadas", "Saldo Insuf. Lib. Aprovacao",
    "Fim Servico", "Inicio Servico",
]
ABAS_ORCAMENTISTA_CLIENTE = [
    "Liberacao Aprovacao", "NF Validadas",
    "Saldo Insuf. Lib. Aprovacao", "Validacao 1 Orcamento", "OS Plataforma",
]
ABAS_SO_NF = [
    "OS Plataforma", "Aguard. 1 Orcamento", "Aguard. Reavaliacao",
    "Cotacoes", "Liberacao Aprovacao", "NF Validadas",
    "Saldo Insuf. Lib. Aprovacao", "Validacao 1 Orcamento", "Aguard. Envio Veiculo",
]
ABAS_COLUNAS_OS = [
    "Aguard. 1 Orcamento", "Aguard. Reavaliacao", "Cotacoes",
    "Liberacao Aprovacao", "NF c/ Problema", "OS Plataforma",
    "Validacao 1 Orcamento", "Validacao NF",
]


# =============================================================
#  CARREGAMENTO BASE — OS (logica comum extraida)
# =============================================================

def _carregar_carteiras():
    df = pd.read_excel(CONFIG["carteiras_file"], sheet_name=0)
    df.columns = df.columns.str.strip()
    return df

def _carregar_os_base(caminho_arquivo, tabela_carteira):
    """Lê o arquivo OS, faz merge com carteiras, dedup e remove status indesejados."""
    tabela = pd.read_excel(caminho_arquivo, sheet_name=0)
    tabela.columns = tabela.columns.str.strip()

    for col in ["Validação NF", "Validacao NF", "Atendimento", "Gerenciadora"]:
        if col in tabela.columns:
            tabela.drop(columns=col, inplace=True)

    cart_cols = ["Cod", "Lider", "Status", "Atendimento"]
    vnf_col_cart = _get_col_norm(tabela_carteira, "Validacao NF")
    if vnf_col_cart:
        cart_cols.append(vnf_col_cart)
    if "Gerenciadora" in tabela_carteira.columns:
        cart_cols.append("Gerenciadora")

    tabela_final = pd.merge(
        tabela, tabela_carteira[cart_cols],
        left_on="Cod Cliente", right_on="Cod", how="left",
    )
    tabela_final.drop(columns="Cod", inplace=True)
    tabela_final["cliente_os"] = (
        tabela_final["Cod Cliente"].astype(str) + "-" + tabela_final["Cod OS"].astype(str)
    )
    tabela_final = tabela_final[~tabela_final["Status OS"].isin(STATUS_INDESEJADOS)]
    tabela_final = tabela_final.drop_duplicates(subset="cliente_os")

    # Recalcula "Qtd Dias em Atraso" a partir da Data Abertura, em vez de confiar
    # no valor que vem pronto do arquivo (que pode vir errado/inflado na origem).
    data_abertura = pd.to_datetime(tabela_final["Data Abertura"], errors="coerce")
    hoje = pd.Timestamp(dt.datetime.now().date())
    tabela_final["Qtd Dias em Atraso"] = (hoje - data_abertura.dt.normalize()).dt.days

    sem_data = data_abertura.isna().sum()
    if sem_data:
        log.warning(f"[OS] {sem_data} linha(s) sem 'Data Abertura' valida — "
                    f"'Qtd Dias em Atraso' ficou em branco para essas linhas.")

    return tabela_final


# =============================================================
#  TRATAMENTO — OS (com filtro 60+)
# =============================================================

def tratar_dados_os(caminho_arquivo, tabela_carteira):
    tabela_final = _carregar_os_base(caminho_arquivo, tabela_carteira)

    tabela_final = tabela_final[tabela_final["Qtd Dias em Atraso"] >= 60]
    tabela_final = tabela_final[~_isin_norm(tabela_final["Atendimento"], AUTOGESTAO)]
    tabela_final = tabela_final[tabela_final["Status"] != "Inativo"]

    eh_prime = tabela_final["Gerenciadora"].str.strip().str.upper() == "PRIME"

    tabela_final.loc[
        eh_prime &
        (tabela_final["Lider"] == "Anderson de Oliveira") &
        _eq_norm(tabela_final["Atendimento"], "Equipe Especializada"),
        "Lider"] = "BOP"

    # vnf_col resolvido uma unica vez
    vnf_col = _get_col_norm(tabela_final, "Validacao NF")

    tabela_final.loc[
        eh_prime &
        _eq_norm(tabela_final["Aba OS"], "Validacao NF") &
        _eq_norm(tabela_final[vnf_col], "EQUIPE PLATAFORMA") &
        _isin_norm(tabela_final["Atendimento"], ATENDIMENTOS_NF),
        "Lider"] = "NF"
    tabela_final.loc[
        eh_prime &
        (tabela_final["Aba OS"] == "NF c/ Problema") &
        _eq_norm(tabela_final[vnf_col], "EQUIPE PLATAFORMA") &
        _isin_norm(tabela_final["Atendimento"], ATENDIMENTOS_NF),
        "Lider"] = "NF/Problema"

    mask_vnf_nao_plat = (
        _isin_norm(tabela_final["Aba OS"], ["Validacao NF", "NF c/ Problema"]) &
        ~_eq_norm(tabela_final[vnf_col], "EQUIPE PLATAFORMA")
    )
    tabela_final.loc[mask_vnf_nao_plat, "Lider"] = "Cliente"

    mask_orc = (
        _isin_norm(tabela_final["Aba OS"], ABAS_ORCAMENTISTA_CLIENTE) &
        _isin_norm(tabela_final["Atendimento"], ["Equipe Orcamentista"])
    )
    tabela_final.loc[mask_orc, "Lider"] = "Cliente"

    mask_vnf_orc = (
        _eq_norm(tabela_final["Aba OS"], "Validacao NF") &
        _isin_norm(tabela_final["Atendimento"], ["Equipe Orcamentista"]) &
        _eq_norm(tabela_final[vnf_col], "CLIENTE")
    )
    tabela_final.loc[mask_vnf_orc, "Lider"] = "Cliente"
    tabela_final.loc[_isin_norm(tabela_final["Aba OS"], ABAS_CLIENTE), "Lider"] = "Cliente"

    tabela_final = tabela_final[~(
        _eq_norm(tabela_final[vnf_col], "Cliente") &
        _isin_norm(tabela_final["Aba OS"], ["NF c/ Problema", "Validacao NF"])
    )]
    tabela_final = tabela_final[~(
        _isin_norm(tabela_final["Atendimento"].fillna(""), ["Validacao de NF"]) &
        _isin_norm(tabela_final["Aba OS"], ABAS_SO_NF)
    )]

    log.info(f"[OS] Total final (60+): {len(tabela_final)} linhas")
    return tabela_final


# =============================================================
#  TRATAMENTO — OS SEM FILTRO 60d (para aba "Todos OS")
# =============================================================

def tratar_dados_os_sem_filtro(caminho_arquivo, tabela_carteira):
    tabela_final = _carregar_os_base(caminho_arquivo, tabela_carteira)
    log.info(f"[OS] Total sem filtro (todos): {len(tabela_final)} linhas")
    return tabela_final


# =============================================================
#  TRATAMENTO — APROVADAS
#  Retorna (df_60plus, base_por_gerenciadora) para evitar releitura do arquivo
# =============================================================

def tratar_dados_aprovadas(tabela_carteira):
    aprovadas_path = CONFIG["aprovadas_file"]
    if not aprovadas_path.exists():
        log.warning("Arquivo APROVADAS nao encontrado. Pulando.")
        return pd.DataFrame(), {}

    tabela = pd.read_excel(aprovadas_path, sheet_name=0)
    tabela.columns = tabela.columns.str.strip()

    tabela_final = pd.merge(
        tabela,
        tabela_carteira[["Cod", "Lider"]],
        left_on="Código Cliente", right_on="Cod",
        how="left",
    )
    tabela_final.drop(columns="Cod", inplace=True)

    tabela_final["cliente_os"] = (
        tabela_final["Código Cliente"].astype(str) + "-" +
        tabela_final["Código OS"].astype(str)
    )
    tabela_final = tabela_final.drop_duplicates(subset="cliente_os")

    tabela_final["Receita Parada"] = (
        tabela_final["Valor Total de Itens"].fillna(0) +
        tabela_final["Valor Total de MDO"].fillna(0)
    )

    hoje = pd.to_datetime(dt.datetime.today().date())
    tabela_final["Dias em aberto"] = (
        hoje - pd.to_datetime(tabela_final["Data de Aprovação"])
    ).dt.days

    # Calcula base bruta por gerenciadora ANTES do filtro 60d
    base_aprov_por_ger = {}
    if "Gerenciadora" in tabela_final.columns:
        for ger_k, grp in tabela_final.groupby("Gerenciadora"):
            base_aprov_por_ger[str(ger_k).strip()] = {
                "qtd":     int(len(grp)),
                "receita": float(grp["Receita Parada"].sum()),
            }

    tabela_final = tabela_final[tabela_final["Dias em aberto"] >= 60].copy()

    eh_prime = tabela_final["Gerenciadora"].str.strip().str.upper() == "PRIME"
    tabela_final.loc[eh_prime, "Lider"] = "Equipe Aprovadas"

    log.info(f"[APROVADAS] Total 60+: {len(tabela_final)} linhas")
    log.info(f"[APROVADAS] Gerenciadoras: {sorted(tabela_final['Gerenciadora'].dropna().unique().tolist())}")
    return tabela_final, base_aprov_por_ger


# =============================================================
#  HISTORICO
# =============================================================

def _path_historico(ger, fonte):
    nome = ger.replace("/", "-").replace(" ", "_")
    return CONFIG["pasta_saida"] / f"historico_{fonte}_{nome}.json"

def carregar_historico(ger, fonte):
    p = _path_historico(ger, fonte)
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    return {}

def salvar_historico_os(df_ger, ger, historico):
    CONFIG["pasta_saida"].mkdir(parents=True, exist_ok=True)
    hoje     = dt.datetime.now().strftime("%d/%m/%y")
    clientes = int((df_ger["Lider"] == "Cliente").sum())
    plat     = len(df_ger) - clientes
    historico[hoje] = {
        "plataforma": plat,
        "cliente":    clientes,
        "por_lider":  {k: int(v) for k, v in df_ger.groupby("Lider").size().items()},
    }
    _path_historico(ger, "os").write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info(f"[OS] Historico [{ger}] salvo — {hoje}")

def salvar_historico_aprov(df_ger, ger, historico):
    CONFIG["pasta_saida"].mkdir(parents=True, exist_ok=True)
    hoje = dt.datetime.now().strftime("%d/%m/%y")
    por_lider = {
        lider: {"qtd": int(len(grp)), "receita": float(grp["Receita Parada"].sum())}
        for lider, grp in df_ger.groupby("Lider")
    }
    historico[hoje] = {
        "total_qtd":     int(len(df_ger)),
        "total_receita": float(df_ger["Receita Parada"].sum()),
        "por_lider":     por_lider,
    }
    _path_historico(ger, "aprov").write_text(
        json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    log.info(f"[APROV] Historico [{ger}] salvo — {hoje}")


# =============================================================
#  ACOMPANHAMENTO (aba "Acao") — mapeamento OS a OS entre execucoes
#
#  A cada execucao guardamos um "retrato" (snapshot) das OS 60+ daquele dia
#  em acompanhamento_<fonte>_<gerenciadora>.json:
#       { "2026-09-21": { "123-456": {"Lider": "...", ...}, ... }, ... }
#
#  A partir dos retratos calculamos, para cada OS:
#    - Situacao:  "Permanece" (estava na execucao anterior e continua),
#                 "Nova" (nao estava na execucao anterior),
#                 "Saiu" (estava na execucao anterior e nao esta mais hoje)
#    - Data 1a Entrada: inicio da sequencia ININTERRUPTA de execucoes em que
#      a OS aparece (se sair e voltar, a contagem recomeca)
#    - Dias no Acompanhamento: hoje - Data 1a Entrada
#  Reexecutar no mesmo dia sobrescreve o retrato de hoje (nao duplica).
# =============================================================

COLS_ACAO_OS = ["Cod Cliente", "Cod OS", "Lider", "Aba OS", "Atendimento",
                "Status OS", "Data Abertura", "Qtd Dias em Atraso"]
SNAP_ACAO_OS = ["Cod Cliente", "Cod OS", "Lider", "Aba OS"]

COLS_ACAO_APROV = ["Código Cliente", "Código OS", "Lider", "Data de Aprovação",
                   "Dias em aberto", "Receita Parada"]
SNAP_ACAO_APROV = ["Código Cliente", "Código OS", "Lider", "Receita Parada"]

COLS_ACOMP = ["Cliente-OS", "Situação", "No Relatório Hoje?",
              "Data 1ª Entrada", "Dias no Acompanhamento", "Data Saída"]
_ORDEM_SITUACAO = {"Permanece": 0, "Nova": 1, "Saiu": 2}

def _path_acomp(ger, fonte):
    nome = ger.replace("/", "-").replace(" ", "_")
    return CONFIG["pasta_saida"] / f"acompanhamento_{fonte}_{nome}.json"

def _carregar_acomp(ger, fonte):
    p = _path_acomp(ger, fonte)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        # Nao sobrescreve um arquivo corrompido: guarda de lado e recomeca.
        bkp = p.with_suffix(".json.corrompido")
        p.rename(bkp)
        log.warning(f"[ACAO] {p.name} ilegivel — movido para {bkp.name}; "
                    f"acompanhamento reiniciado.")
        return {}

def _chave_cliente_os(series):
    """'Cliente-OS' estavel entre execucoes (evita '123.0-456' vs '123-456')."""
    return series.astype(str).str.strip().str.replace(r"\.0(?=-|$)", "", regex=True)

def _json_val(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if hasattr(v, "item"):          # escalar numpy -> python
        v = v.item()
    return v if isinstance(v, (str, int, float, bool)) else str(v)

def _inicio_sequencia(chave, datas_ant, hist, padrao):
    """Data mais antiga da sequencia ininterrupta de retratos que contem a chave."""
    inicio = padrao
    for d in reversed(datas_ant):
        if chave not in hist[d]:
            break
        inicio = d
    return inicio

def atualizar_acompanhamento(df_ger, ger, fonte, colunas, campos_snapshot, hoje=None):
    """Atualiza o JSON de acompanhamento e devolve o DataFrame da aba "Acao".

    df_ger           OS 60+ de hoje (precisa da coluna 'cliente_os')
    colunas          colunas de contexto exibidas na aba (as que existirem em df_ger)
    campos_snapshot  colunas guardadas no retrato, para poder exibir as OS que
                     SAIRAM (elas nao estao mais em df_ger)
    """
    hoje     = hoje or dt.date.today()
    hoje_iso = hoje.isoformat()
    hist     = _carregar_acomp(ger, fonte)
    hist.pop(hoje_iso, None)                      # reexecucao no mesmo dia
    tinha_hist = bool(hist)

    colunas = [c for c in colunas if c in df_ger.columns]
    campos  = [c for c in campos_snapshot if c in df_ger.columns]

    atual, ctx_atual = {}, {}
    if len(df_ger):
        chaves = _chave_cliente_os(df_ger["cliente_os"]).tolist()
        for chave, (_, row) in zip(chaves, df_ger.iterrows()):
            atual[chave]     = {c: _json_val(row[c]) for c in campos}
            ctx_atual[chave] = {c: row[c] for c in colunas}

    if not atual and not tinha_hist:
        return pd.DataFrame()                     # nada a acompanhar ainda

    datas_ant = sorted(d for d in hist if d < hoje_iso)
    data_ant  = datas_ant[-1] if datas_ant else None
    snap_ant  = hist[data_ant] if data_ant else {}

    linhas = []
    for chave, ctx in ctx_atual.items():
        inicio = _inicio_sequencia(chave, datas_ant, hist, hoje_iso)
        d_ini  = dt.date.fromisoformat(inicio)
        linhas.append({
            "Cliente-OS": chave,
            "Situação": "Permanece" if chave in snap_ant else "Nova",
            "No Relatório Hoje?": "Sim",
            "Data 1ª Entrada": d_ini,
            "Dias no Acompanhamento": (hoje - d_ini).days,
            "Data Saída": None,
            **ctx,
        })
    for chave, snap in snap_ant.items():
        if chave in atual:
            continue
        d_ini = dt.date.fromisoformat(_inicio_sequencia(chave, datas_ant, hist, data_ant))
        linhas.append({
            "Cliente-OS": chave,
            "Situação": "Saiu",
            "No Relatório Hoje?": "Não",
            "Data 1ª Entrada": d_ini,
            "Dias no Acompanhamento": (hoje - d_ini).days,
            "Data Saída": hoje,
            **{c: snap.get(c) for c in colunas if c in snap},
        })

    hist[hoje_iso] = atual
    CONFIG["pasta_saida"].mkdir(parents=True, exist_ok=True)
    _path_acomp(ger, fonte).write_text(
        json.dumps(hist, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8"
    )

    df = pd.DataFrame(linhas, columns=COLS_ACOMP + colunas)
    df["_o"] = df["Situação"].map(_ORDEM_SITUACAO)
    df = (df.sort_values(["_o", "Dias no Acompanhamento"], ascending=[True, False])
            .drop(columns="_o").reset_index(drop=True))

    cont = df["Situação"].value_counts()
    log.info(f"[ACAO {fonte.upper()}] {ger}: Permanece={cont.get('Permanece', 0)} | "
             f"Nova={cont.get('Nova', 0)} | Saiu={cont.get('Saiu', 0)}"
             + ("" if data_ant else "  (1ª execucao: sem base anterior, tudo entra como Nova)"))
    return df


# =============================================================
#  HELPERS DE CELULA
# =============================================================

def _hdr(ws, r, c, v, bg=None, fg="FFFFFF", size=10):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=True, color=fg, size=size)
    cell.fill      = fill(bg or C["dark_blue"])
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = BORDER

def _dat(ws, r, c, v, bg=None, bold=False, fmt=None, align="center"):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font      = Font(name="Arial", bold=bold, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border    = BORDER
    if bg:  cell.fill          = fill(bg)
    if fmt: cell.number_format = fmt

def _titulo(ws, r, c1, c2, texto, bg=None):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(row=r, column=c1, value=texto)
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    cell.fill      = fill(bg or C["mid_blue"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 22

def _mesclar_hdr(ws, r, c1, c2, texto, bg=None):
    ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    cell = ws.cell(row=r, column=c1, value=texto)
    cell.font      = bfont("FFFFFF")
    cell.fill      = fill(bg or C["mid_blue"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = BORDER
    for c in range(c1 + 1, c2 + 1):
        ws.cell(row=r, column=c).border = BORDER

def _auto_width(ws):
    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, 10), 42)


# =============================================================
#  TABELAS — OS
# =============================================================

def _ord_os(df, col="Lider"):
    ordem = {l: i for i, l in enumerate(CONFIG["lideres_ordem_os"])}
    df = df.copy()
    df["_o"] = df[col].map(lambda x: ordem.get(x, 999))
    return df.sort_values("_o").drop(columns="_o")

def _arvore(ws, df, r0):
    lt = df.groupby("Lider").size().reset_index(name="Q")
    la = df.groupby(["Lider", "Aba OS"]).size().reset_index(name="Q")
    _titulo(ws, r0, 1, 3, "Lider / Aba OS / Quantidade")
    for c, h in enumerate(["Lider", "Aba OS", "Qtd"], 1):
        _hdr(ws, r0+1, c, h)
    r = r0 + 2
    for _, lr in _ord_os(lt).iterrows():
        lider = lr["Lider"]
        _dat(ws, r, 1, lider, bg=C["light_blue"], bold=True, align="left")
        _dat(ws, r, 2, "",    bg=C["light_blue"])
        _dat(ws, r, 3, int(lr["Q"]), bg=C["light_blue"], bold=True)
        r += 1
        for _, sr in la[la["Lider"] == lider].sort_values("Aba OS").iterrows():
            _dat(ws, r, 1, "",           bg=C["alt"])
            _dat(ws, r, 2, sr["Aba OS"], bg=C["alt"], align="left")
            _dat(ws, r, 3, int(sr["Q"]), bg=C["alt"])
            r += 1
    _dat(ws, r, 1, "Total Geral", bold=True, bg=C["dark_blue"])
    ws.cell(row=r, column=1).font = bfont("FFFFFF")
    _dat(ws, r, 2, "", bg=C["dark_blue"])
    _dat(ws, r, 3, int(lt["Q"].sum()), bold=True, bg=C["dark_blue"])
    ws.cell(row=r, column=3).font = bfont("FFFFFF")
    return r + 2

def _datas(ws, df_sem_cli, df_cli, hist, cs=5, r0=1):
    # Mostra so a data do relatorio atual e a do relatorio anterior (nao o
    # historico inteiro) — hist ja inclui a entrada de hoje, pois
    # salvar_historico_os() roda antes de build_sheet().
    todas_datas = list(hist.keys())
    datas = todas_datas[-2:] if len(todas_datas) >= 2 else todas_datas
    ce = cs + len(datas)
    _titulo(ws, r0, cs, ce, "60 Dias Todas Abas (Exceto Aprovadas)")
    for c, h in enumerate(["Data"] + datas, cs):
        _hdr(ws, r0+1, c, h)
    for lbl, row_off in [("Atuacao Plataforma", 2), ("Atuacao Cliente", 3)]:
        cell = ws.cell(row=r0+row_off, column=cs, value=lbl)
        cell.font = bfont(); cell.border = BORDER
    for i, d in enumerate(datas):
        v = hist[d]
        _dat(ws, r0+2, cs+1+i, v.get("plataforma", ""))
        _dat(ws, r0+3, cs+1+i, v.get("cliente", ""))
    # Destaca a coluna de hoje (a ultima) em negrito, para deixar claro qual
    # e o relatorio atual.
    if datas:
        c_hoje = ce
        for row_off in (2, 3):
            ws.cell(row=r0+row_off, column=c_hoje).font = Font(name="Arial", bold=True, size=10)
    return r0 + 5

def _total_os(ws, df_todos_ger, cs=5, r0=8):
    # Os numeros aqui tem que bater exatamente com a aba "Todos OS" dessa
    # gerenciadora: total geral = todas as linhas dessa aba; 60+ / 59- vem
    # direto da coluna "Qtd Dias em Atraso" da mesma base, sem nenhum filtro
    # de negocio extra (autogestao, inativos, etc.) que so vale para a aba
    # operacional de 60+.
    ce = cs + 3
    _titulo(ws, r0, cs, ce, "Total - OS")
    for c, h in enumerate(["Totais", "Qtde OS", "Porcentagem", "Total geral"], cs):
        _hdr(ws, r0+1, c, h)
    base = int(len(df_todos_ger))
    t60  = int((df_todos_ger["Qtd Dias em Atraso"] >= 60).sum()) if base else 0
    t59  = max(0, base - t60)
    for i, (lbl, qty, pct) in enumerate([
        ("Total 60+",   t60,  t60  / base if base else 0),
        ("Total 59-",   t59,  t59  / base if base else 0),
        ("Total geral", base, 1.0),
    ]):
        rr = r0 + 2 + i; bold = (i == 2)
        _dat(ws, rr, cs,   lbl,  bold=bold)
        _dat(ws, rr, cs+1, qty,  bold=bold)
        _dat(ws, rr, cs+2, pct,  bold=bold, fmt="0.00%")
        if i == 0: _dat(ws, rr, ce, base, bold=True)
        else:      ws.cell(row=rr, column=ce).border = BORDER
    return r0 + 6

def _lider_data(ws, df, hist, cs=5, r0=15):
    # Mostra so as ultimas 5 execucoes (~ultimo mes / ultimas 5 semanas), nao
    # o historico inteiro. O calculo de saida/entrada continua comparando
    # sempre hoje x execucao imediatamente anterior, usando o historico
    # completo — mesmo que a coluna "anterior" caia fora da janela exibida.
    todas_datas = list(hist.keys())
    datas = todas_datas[-5:]
    ce    = cs + len(datas) + 2
    _titulo(ws, r0, cs, ce, "60+ Todas as Abas - Quantidade - Lider x Por Data")
    _mesclar_hdr(ws, r0+1, ce-1, ce,
                 dt.datetime.now().strftime("%d/%m/%Y") + " (Qtde OS)",
                 bg=C["mid_blue"])
    for c, h in enumerate(
        ["Lideres"] + [d + " (Qtde OS)" for d in datas] + ["saidas", "entrada"], cs
    ):
        bg = C["saida_bg"] if h == "saidas" else C["entrada_bg"] if h == "entrada" else C["dark_blue"]
        _hdr(ws, r0+2, c, h, bg=bg)
    lh = _ord_os(df.groupby("Lider").size().reset_index(name="Q"))
    tl = {d: hist[d].get("por_lider", {}) for d in todas_datas}
    # Penultima data (do historico completo) = execucao anterior; ultima = hoje
    # (recem salva, igual ao dado vivo)
    data_ant = todas_datas[-2] if len(todas_datas) >= 2 else None
    # Apenas lideres visiveis (nao-cliente) para calcular totais consistentes
    lideres_visiveis = set(lh["Lider"].tolist())
    r = r0 + 3; ss = se = 0
    for _, lr in lh.iterrows():
        lider = lr["Lider"]; hv = int(lr["Q"])
        _dat(ws, r, cs, lider, align="left")
        for i, d in enumerate(datas):
            _dat(ws, r, cs+1+i, int(tl[d].get(lider, 0)))
        ant = int(tl[data_ant].get(lider, 0)) if data_ant else 0
        s = max(0, ant - hv); e = max(0, hv - ant); ss += s; se += e
        _dat(ws, r, ce-1, s, bg=C["green"] if s else None)
        _dat(ws, r, ce,   e, bg=C["red"]   if e else None)
        r += 1
    _dat(ws, r, cs, "Total", bold=True, bg=C["dark_blue"])
    ws.cell(row=r, column=cs).font = bfont("FFFFFF")
    # Total por data: soma so lideres visiveis para alinhar com as linhas acima
    for i, d in enumerate(datas):
        t = int(sum(tl[d].get(ldr, 0) for ldr in lideres_visiveis))
        _dat(ws, r, cs+1+i, t, bold=True, bg=C["dark_blue"])
        ws.cell(row=r, column=cs+1+i).font = bfont("FFFFFF")
    # Garante que a coluna de hoje mostre o total vivo (nao o do historico)
    _dat(ws, r, cs + len(datas), int(lh["Q"].sum()), bold=True, bg=C["dark_blue"])
    ws.cell(row=r, column=cs + len(datas)).font = bfont("FFFFFF")
    _dat(ws, r, ce-1, ss, bold=True, bg=C["saida_bg"])
    _dat(ws, r, ce,   se, bold=True, bg=C["entrada_bg"])
    ws.cell(row=r, column=ce-1).font = bfont("FFFFFF")
    ws.cell(row=r, column=ce).font   = bfont("FFFFFF")
    return r + 2

def _lider_aba(ws, df, cs=5, r0=30):
    pivot = df.groupby(["Lider", "Aba OS"]).size().unstack(fill_value=0).reset_index()
    # Mapeia nomes normalizados para nomes reais presentes no pivot (preserva acentos no output)
    norm_to_actual = {_norm(col): col for col in pivot.columns if col != "Lider"}
    abas = [norm_to_actual[_norm(a)] for a in ABAS_COLUNAS_OS if _norm(a) in norm_to_actual]

    ce = cs + len(abas)
    _titulo(ws, r0, cs, ce, "60+ Todas as Abas - Quantidade - Lider x Por Aba")
    for c, h in enumerate(["Lideres"] + abas, cs):
        _hdr(ws, r0+1, c, h)
    r = r0 + 2
    for _, pr in _ord_os(pivot).iterrows():
        _dat(ws, r, cs, pr["Lider"], align="left")
        for i, a in enumerate(abas):
            _dat(ws, r, cs+1+i, int(pr.get(a, 0)))
        r += 1
    _dat(ws, r, cs, "Total", bold=True, bg=C["dark_blue"])
    ws.cell(row=r, column=cs).font = bfont("FFFFFF")
    for i, a in enumerate(abas):
        t = int(pivot[a].sum()) if a in pivot.columns else 0
        _dat(ws, r, cs+1+i, t, bold=True, bg=C["dark_blue"])
        ws.cell(row=r, column=cs+1+i).font = bfont("FFFFFF")
    return r + 2


# =============================================================
#  TABELAS — APROVADAS
# =============================================================

def _aprov_lider_data(ws, df_ger, hist_aprov, cs=5, r0=1):
    datas     = list(hist_aprov.keys())
    data_hoje = datas[-1] if datas else dt.datetime.now().strftime("%d/%m/%y")
    data_ant  = datas[-2] if len(datas) >= 2 else None

    C_LID  = cs
    C_QANT = cs + 1;  C_RANT = cs + 2
    C_QHOJ = cs + 3;  C_RHOJ = cs + 4;  C_PCT = cs + 5
    C_SAID = cs + 6;  C_SREC = cs + 7
    C_ENT  = cs + 8;  C_EREC = cs + 9
    CE     = C_EREC

    _titulo(ws, r0, C_LID, CE, "60 Dias - APROVADAS")
    ws.cell(row=r0+1, column=C_LID).border = BORDER
    label_ant = f"Qtde.OS {data_ant}" if data_ant else "Qtde.OS (anterior)"
    _mesclar_hdr(ws, r0+1, C_QANT, C_RANT, label_ant, bg=C["mid_blue"])
    _mesclar_hdr(ws, r0+1, C_QHOJ, CE, f"Qtde.OS {data_hoje}", bg=C["mid_blue"])

    for col, txt, bg in [
        (C_LID,  "Líderes",          C["dark_blue"]),
        (C_QANT, "Qtde.OS",          C["dark_blue"]),
        (C_RANT, "Receita Parada $",  C["dark_blue"]),
        (C_QHOJ, "Qtde.OS",          C["dark_blue"]),
        (C_RHOJ, "Receita Parada $",  C["dark_blue"]),
        (C_PCT,  "Porcentagem",       C["dark_blue"]),
        (C_SAID, "Saida",             C["saida_bg"]),
        (C_SREC, "Saida receita",     C["saida_bg"]),
        (C_ENT,  "Entrada",           C["entrada_bg"]),
        (C_EREC, "Entrada receita",   C["entrada_bg"]),
    ]:
        _hdr(ws, r0+2, col, txt, bg=bg)

    grp_hoje = (
        df_ger.groupby("Lider")
        .agg(qtd=("Receita Parada", "count"), rec=("Receita Parada", "sum"))
        .reset_index()
        .sort_values("qtd", ascending=False)
    )
    hist_ant = hist_aprov.get(data_ant, {}).get("por_lider", {}) if data_ant else {}

    total_hoje_qtd = int(grp_hoje["qtd"].sum())
    total_hoje_rec = float(grp_hoje["rec"].sum())
    total_ant_qtd  = int(sum(v.get("qtd", 0)      for v in hist_ant.values()))
    total_ant_rec  = float(sum(v.get("receita", 0) for v in hist_ant.values()))

    r = r0 + 3
    tot_said_q = tot_said_r = tot_ent_q = tot_ent_r = 0.0

    for _, row in grp_hoje.iterrows():
        lider = row["Lider"]
        q_hoj = int(row["qtd"]); r_hoj = float(row["rec"])
        pct   = q_hoj / total_hoje_qtd if total_hoje_qtd else 0.0
        ant   = hist_ant.get(lider, {})
        q_ant = int(ant.get("qtd", 0)); r_ant = float(ant.get("receita", 0.0))
        s_q = max(0, q_ant - q_hoj); s_r = max(0.0, r_ant - r_hoj)
        e_q = max(0, q_hoj - q_ant); e_r = max(0.0, r_hoj - r_ant)
        tot_said_q += s_q; tot_said_r += s_r
        tot_ent_q  += e_q; tot_ent_r  += e_r

        _dat(ws, r, C_LID,  lider, align="left")
        _dat(ws, r, C_QANT, q_ant)
        _dat(ws, r, C_RANT, r_ant, fmt='#,##0.00')
        _dat(ws, r, C_QHOJ, q_hoj)
        _dat(ws, r, C_RHOJ, r_hoj, fmt='#,##0.00')
        _dat(ws, r, C_PCT,  pct,   fmt="0.00%")
        _dat(ws, r, C_SAID, s_q,   bg=C["red"]   if s_q else None)
        _dat(ws, r, C_SREC, s_r,   bg=C["red"]   if s_r else None, fmt='#,##0.00')
        _dat(ws, r, C_ENT,  e_q,   bg=C["green"] if e_q else None)
        _dat(ws, r, C_EREC, e_r,   bg=C["green"] if e_r else None, fmt='#,##0.00')
        r += 1

    _dat(ws, r, C_LID,  "Total geral",      bold=True, bg=C["gray_tot"])
    _dat(ws, r, C_QANT, int(total_ant_qtd), bold=True, bg=C["gray_tot"])
    _dat(ws, r, C_RANT, total_ant_rec,      bold=True, bg=C["gray_tot"], fmt='#,##0.00')
    _dat(ws, r, C_QHOJ, total_hoje_qtd,     bold=True, bg=C["gray_tot"])
    _dat(ws, r, C_RHOJ, total_hoje_rec,     bold=True, bg=C["gray_tot"], fmt='#,##0.00')
    _dat(ws, r, C_PCT,  1.0,                bold=True, bg=C["gray_tot"], fmt="0.00%")
    _dat(ws, r, C_SAID, int(tot_said_q),    bold=True, bg=C["saida_bg"])
    _dat(ws, r, C_SREC, tot_said_r,         bold=True, bg=C["saida_bg"], fmt='#,##0.00')
    _dat(ws, r, C_ENT,  int(tot_ent_q),     bold=True, bg=C["entrada_bg"])
    _dat(ws, r, C_EREC, tot_ent_r,          bold=True, bg=C["entrada_bg"], fmt='#,##0.00')
    for cc in [C_SAID, C_SREC, C_ENT, C_EREC]:
        ws.cell(row=r, column=cc).font = bfont("FFFFFF")
    return r + 2


def _aprov_totais(ws, df_ger, base_aprov_ger, cs=5, r0=1):
    CE = cs + 3
    _titulo(ws, r0, cs, CE, "APROVADAS - Total")
    hoje_label = dt.datetime.now().strftime("%d/%m")
    for c, h in enumerate(
        ["Totais", f"Qtde.OS {hoje_label}", "Receita Parada $", "Porcentagem"], cs
    ):
        _hdr(ws, r0+1, c, h)

    t60_q  = int(len(df_ger))
    t60_r  = float(df_ger["Receita Parada"].sum())
    base_q = int(base_aprov_ger["qtd"])
    base_r = float(base_aprov_ger["receita"])
    t59_q  = max(0, base_q - t60_q)
    t59_r  = max(0.0, base_r - t60_r)

    for i, (lbl, qty, rec, pct) in enumerate([
        ("Total 60+",   t60_q,  t60_r,  t60_q / base_q if base_q else 0),
        ("Total 59-",   t59_q,  t59_r,  t59_q / base_q if base_q else 0),
        ("Total geral", base_q, base_r, 1.0),
    ]):
        rr = r0 + 2 + i; bold = (i == 2)
        _dat(ws, rr, cs,   lbl,  bold=bold)
        _dat(ws, rr, cs+1, qty,  bold=bold)
        _dat(ws, rr, cs+2, rec,  bold=bold, fmt='#,##0.00')
        _dat(ws, rr, cs+3, pct,  bold=bold, fmt="0.00%")
    return r0 + 6


# =============================================================
#  ABA DE DADOS GENERICA
# =============================================================

def _aba_dados(wb, df, titulo, cabecalho=None, formatos=None, cores=None):
    """formatos: {coluna: number_format}; cores: {coluna: {valor: cor_hex}}."""
    ws = wb.create_sheet(title=titulo[:31])
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ncols = len(df.columns)
    idx = {c: i for i, c in enumerate(df.columns, 1)}
    fmt_col   = {idx[c]: f for c, f in (formatos or {}).items() if c in idx}
    cores_col = {idx[c]: m for c, m in (cores or {}).items() if c in idx}

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=cabecalho or titulo)
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    cell.fill      = fill(C["dark_blue"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for c, col in enumerate(df.columns, 1):
        _hdr(ws, 2, c, col)
    ws.row_dimensions[2].height = 28

    # Objetos de estilo criados uma vez fora do loop (melhora performance)
    font_data      = Font(name="Arial", size=10)
    align_center   = Alignment(vertical="center", horizontal="center")
    fill_alt       = fill(C["alt"])

    for r_idx, row in enumerate(df.itertuples(index=False), start=3):
        row_vals = [None if (v is pd.NaT or (isinstance(v, float) and pd.isna(v))) else v
                    for v in row]
        ws.append(row_vals)
        alt = (r_idx % 2 == 0)
        for c_idx in range(1, ncols + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font      = font_data
            cell.border    = BORDER
            cell.alignment = align_center
            if alt:
                cell.fill = fill_alt
            if c_idx in fmt_col:
                cell.number_format = fmt_col[c_idx]
            if c_idx in cores_col:
                cor = cores_col[c_idx].get(cell.value)
                if cor:
                    cell.fill = fill(cor)

    for c_idx, col in enumerate(df.columns, 1):
        max_len = max(len(str(col)), df[col].astype(str).str.len().max())
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 3, 40)

    ws.auto_filter.ref = f"A2:{get_column_letter(ncols)}{len(df)+2}"


# =============================================================
#  MONTAGEM DO WORKBOOK
# =============================================================

def _aba_acao(wb, df_acao, titulo):
    """Aba de acompanhamento OS a OS (Cliente-OS, permanencia e dias)."""
    hoje = dt.datetime.now().strftime("%d/%m/%Y")
    _aba_dados(
        wb, df_acao, titulo,
        cabecalho=f"{titulo} — acompanhamento 60+ em {hoje}",
        formatos={"Data 1ª Entrada": "dd/mm/yyyy", "Data Saída": "dd/mm/yyyy"},
        cores={"Situação": {"Permanece": C["yellow"], "Nova": C["red"], "Saiu": C["green"]}},
    )

def build_sheet(wb, ger, df_os_ger, hist_os,
                df_aprov_ger, hist_aprov, base_aprov_ger,
                df_os_todos=None, df_acao_os=None, df_acao_aprov=None):

    df_os_sem_cli = df_os_ger[df_os_ger["Lider"] != "Cliente"]
    df_os_cli     = df_os_ger[df_os_ger["Lider"] == "Cliente"]
    # Base unica usada tanto pelo bloco "Total - OS" quanto pela aba "Todos
    # OS", pra garantir que os dois batam sempre.
    df_base_todos = df_os_todos if df_os_todos is not None else df_os_ger

    ws = wb.create_sheet(title="Resumo")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    ws.merge_cells("A1:D1")
    cell = ws["A1"]
    cell.value     = f"Gerenciadora: {ger}  |  {dt.datetime.now():%d/%m/%Y %H:%M}"
    cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=12)
    cell.fill      = fill(C["dark_blue"])
    cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 25

    _arvore(ws, df_os_sem_cli, r0=3)
    r = _datas(ws,    df_os_sem_cli, df_os_cli, hist_os, cs=5, r0=3)
    r = _total_os(ws, df_base_todos,          cs=5, r0=r)
    r = _lider_data(ws, df_os_sem_cli, hist_os, cs=5, r0=r)
    r = _lider_aba(ws,  df_os_sem_cli,          cs=5, r0=r)

    r_aprov = r + 2
    if not df_aprov_ger.empty:
        r = _aprov_lider_data(ws, df_aprov_ger, hist_aprov, cs=5, r0=r_aprov)
        _aprov_totais(ws, df_aprov_ger, base_aprov_ger, cs=5, r0=r)
    else:
        ws.cell(row=r_aprov, column=5,
                value=f"Sem dados APROVADAS 60+ para {ger}").font = bfont()

    _auto_width(ws)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 10

    if df_acao_os is not None and not df_acao_os.empty:
        _aba_acao(wb, df_acao_os, "Ação OS")
    if df_acao_aprov is not None and not df_acao_aprov.empty:
        _aba_acao(wb, df_acao_aprov, "Ação Aprovadas")

    cols_os = [c for c in df_os_sem_cli.columns if c != "cliente_os"]
    _aba_dados(wb, df_os_sem_cli[cols_os].reset_index(drop=True), "Dados OS")

    if not df_os_cli.empty:
        cols_cli = [c for c in df_os_cli.columns if c != "cliente_os"]
        _aba_dados(wb, df_os_cli[cols_cli].reset_index(drop=True), "Cliente OS")

    cols_all = [c for c in df_base_todos.columns if c != "cliente_os"]
    _aba_dados(wb, df_base_todos[cols_all].reset_index(drop=True), "Todos OS")

    if not df_aprov_ger.empty:
        cols_aprov = [c for c in df_aprov_ger.columns if c != "cliente_os"]
        _aba_dados(wb, df_aprov_ger[cols_aprov].reset_index(drop=True), "Aprovadas")


# =============================================================
#  MAIN
# =============================================================

def main():
    print("=" * 60)
    print("  Processador OS + APROVADAS por Gerenciadora")
    print("=" * 60)

    arquivo_os = escolher_arquivo()
    if not arquivo_os:
        input("\nPressione Enter para sair...")
        return
    print(f"\nOS:        {Path(arquivo_os).name}")

    # Carteiras lidas uma unica vez e passadas como parametro
    tabela_carteira = _carregar_carteiras()

    try:
        df_os = tratar_dados_os(arquivo_os, tabela_carteira)
    except (KeyError, ValueError) as e:
        print(f"\nERRO ao tratar OS: {e}")
        input("\nPressione Enter para sair...")
        return

    try:
        df_os_todos_raw = tratar_dados_os_sem_filtro(arquivo_os, tabela_carteira)
    except (KeyError, ValueError) as e:
        log.warning(f"Nao foi possivel carregar base completa de OS: {e}")
        df_os_todos_raw = df_os.copy()

    try:
        # base_aprov_por_ger vem junto, sem releitura do arquivo
        df_aprov, base_aprov_por_ger = tratar_dados_aprovadas(tabela_carteira)
        if not df_aprov.empty:
            print(f"APROVADAS: {CONFIG['aprovadas_file'].name}")
    except (KeyError, ValueError) as e:
        print(f"\nERRO ao tratar APROVADAS: {e}")
        df_aprov          = pd.DataFrame()
        base_aprov_por_ger = {}

    gers_os    = set(df_os["Gerenciadora"].dropna().unique()) if "Gerenciadora" in df_os.columns else set()
    gers_aprov = (set(df_aprov["Gerenciadora"].dropna().unique())
                  if not df_aprov.empty and "Gerenciadora" in df_aprov.columns else set())
    geradoras  = sorted(gers_os | gers_aprov) or ["Geral"]

    CONFIG["pasta_saida"].mkdir(parents=True, exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")

    gerados = []
    for ger in geradoras:
        ger_upper = str(ger).strip().upper()

        df_os_ger = (
            df_os.copy() if "Gerenciadora" not in df_os.columns
            else df_os[df_os["Gerenciadora"].str.strip().str.upper() == ger_upper].copy()
        )
        df_aprov_ger = (
            pd.DataFrame() if df_aprov.empty or "Gerenciadora" not in df_aprov.columns
            else df_aprov[df_aprov["Gerenciadora"].str.strip().str.upper() == ger_upper].copy()
        )
        base_aprov_ger = next(
            (v for k, v in base_aprov_por_ger.items() if k.strip().upper() == ger_upper),
            {"qtd": 0, "receita": 0.0},
        )
        df_todos_ger = (
            df_os_todos_raw[
                df_os_todos_raw["Gerenciadora"].str.strip().str.upper() == ger_upper
            ].copy()
            if "Gerenciadora" in df_os_todos_raw.columns
            else df_os_todos_raw.copy()
        )

        hist_os    = carregar_historico(ger, "os")
        hist_aprov = carregar_historico(ger, "aprov")

        salvar_historico_os(df_os_ger, ger, hist_os)
        if not df_aprov_ger.empty:
            salvar_historico_aprov(df_aprov_ger, ger, hist_aprov)

        # Acompanhamento OS a OS (abas "Acao"). Para APROVADAS so atualiza se o
        # arquivo foi carregado (df_aprov nao vazio); se falhou/faltou, nao grava
        # um retrato vazio — senao todas as OS apareceriam como "Saiu".
        # Aba "Acao OS" acompanha exatamente as OS da aba "Dados OS" (60+ ja
        # tratadas, sem as de lider "Cliente", que ficam na aba "Cliente OS").
        df_os_dados = df_os_ger[df_os_ger["Lider"] != "Cliente"]
        df_acao_os = atualizar_acompanhamento(
            df_os_dados, ger, "os", COLS_ACAO_OS, SNAP_ACAO_OS)
        df_acao_aprov = (
            atualizar_acompanhamento(
                df_aprov_ger, ger, "aprov", COLS_ACAO_APROV, SNAP_ACAO_APROV)
            if not df_aprov.empty else pd.DataFrame()
        )

        nome = f"OS_{ger}_{ts}.xlsx".replace("/", "-").replace(" ", "_")
        path = CONFIG["pasta_saida"] / nome
        wb   = Workbook()
        wb.remove(wb.active)
        build_sheet(
            wb, ger,
            df_os_ger,    hist_os,
            df_aprov_ger, hist_aprov, base_aprov_ger,
            df_os_todos=df_todos_ger,
            df_acao_os=df_acao_os,
            df_acao_aprov=df_acao_aprov,
        )
        wb.save(path)
        gerados.append(path)
        print(f"  Gerado: {nome}  (OS 60+={len(df_os_ger)} | Todos OS={len(df_todos_ger)} | Aprov={len(df_aprov_ger)})")

    print(f"\n{len(gerados)} relatorio(s) salvo(s) em:")
    print(f"   {CONFIG['pasta_saida']}")
    input("\nPressione Enter para sair...")


if __name__ == "__main__":
    main()
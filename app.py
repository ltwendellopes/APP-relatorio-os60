"""
OS 60+ — aplicativo de mesa (interface grafica) para gerar os relatorios de
OS/Aprovadas por gerenciadora, sem precisar de Python instalado.

Este arquivo NAO contem a regra de negocio — ela mora inteira em
report_engine.py (o "codigo_final_60_.py" original, sem alteracoes).
Aqui so tem: janela, seletores de arquivo, fio de execucao e otimizacoes de
leitura/escrita de Excel iguais as usadas na versao web (leitor calamine,
escrita "preguicosa" de linhas grandes).
"""
import sys
import os
import io
import json
import queue
import shutil
import threading
import traceback
import datetime as dt
from pathlib import Path
from copy import copy

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------------------------------------------------------------------
#  Config persistida entre execucoes (ultimos arquivos/pasta usados)
# ---------------------------------------------------------------------------
CONFIG_PATH = Path.home() / ".os60_gui_config.json"


def carregar_config():
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def salvar_config(d):
    try:
        CONFIG_PATH.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------------------
#  Importa a regra de negocio original
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).resolve().parent))
import report_engine as motor  # noqa: E402

import pandas as pd  # noqa: E402
from openpyxl.cell.cell import Cell as _Cell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE as _ILLEGAL_RE
from openpyxl.styles import Font as _Font, Alignment as _Alignment
from openpyxl.utils import get_column_letter as _gcl

# ---------------------------------------------------------------------------
#  Otimizacoes de desempenho aplicadas POR FORA (mesma ideia da versao web):
#  leitura de cada planilha uma unica vez (com leitor rapido quando possivel)
#  e escrita das abas de dados sem recriar estilo celula a celula.
#  A regra de negocio do report_engine nao muda.
# ---------------------------------------------------------------------------
_orig_read_excel = pd.read_excel
_read_cache = {}
_LIMITE_CALAMINE = 60 * 1024 * 1024
_COLS_DATA_TEXTO = ("Data Abertura", "Data de Aprovação")


def _tem_calamine():
    try:
        import python_calamine  # noqa: F401
        return True
    except Exception:
        return False


def _read_excel_cache(path, *a, **k):
    key = (str(path), repr(a), repr(sorted(k.items())))
    if key not in _read_cache:
        print(f"  Lendo {Path(str(path)).name} ...")
        df = None
        try:
            tam = os.path.getsize(str(path))
        except OSError:
            tam = 0
        if "engine" not in k and _tem_calamine() and tam <= _LIMITE_CALAMINE:
            try:
                df = _orig_read_excel(path, *a, engine="calamine", **k)
            except Exception as e:
                print(f"  (leitor rápido falhou — {type(e).__name__}; usando o leitor padrão)")
                df = None
        if df is None:
            df = _orig_read_excel(path, *a, **k)
        for col in df.columns:
            if str(col).strip() in _COLS_DATA_TEXTO and not pd.api.types.is_datetime64_any_dtype(df[col]):
                conv = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
                if conv.notna().sum() >= df[col].notna().sum() * 0.5:
                    df[col] = conv
        _read_cache[key] = df
        print(f"  {Path(str(path)).name}: {len(df):,} linhas x {len(df.columns)} colunas".replace(",", "."))
    return _read_cache[key].copy(deep=False)


motor.pd.read_excel = _read_excel_cache

_orig_isin, _orig_eq = motor._isin_norm, motor._eq_norm
motor._isin_norm = lambda s, v: _orig_isin(s, v).astype(bool)
motor._eq_norm = lambda s, v: _orig_eq(s, v).astype(bool)


def _aba_dados_rapido(wb, df, titulo, cabecalho=None, formatos=None, cores=None):
    C, fill, BORDER = motor.C, motor.fill, motor.BORDER
    ws = wb.create_sheet(title=titulo[:31])
    ws.freeze_panes = "A3"
    ws.sheet_view.showGridLines = False
    ncols = len(df.columns)
    print(f"  Montando aba '{titulo}' ({len(df):,} linhas) ...".replace(",", "."))

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=cabecalho or titulo)
    cell.font = _Font(name="Arial", bold=True, color="FFFFFF", size=12)
    cell.fill = fill(C["dark_blue"])
    cell.alignment = _Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for c, col in enumerate(df.columns, 1):
        motor._hdr(ws, 2, c, col)
    ws.row_dimensions[2].height = 28

    idx = {col: i for i, col in enumerate(df.columns, 1)}
    fmt_col = {idx[c]: f for c, f in (formatos or {}).items() if c in idx}
    cores_col = {idx[c]: mapa for c, mapa in (cores or {}).items() if c in idx}

    tmp = wb.create_sheet("__tmp_estilo__")
    tc = tmp.cell(row=1, column=1)
    tc.font = _Font(name="Arial", size=10)
    tc.border = BORDER
    tc.alignment = _Alignment(vertical="center", horizontal="center")
    base = {False: copy(tc._style)}
    tc.fill = fill(C["alt"])
    base[True] = copy(tc._style)

    numfmt_id = {}
    for c_idx, fmt_str in fmt_col.items():
        tc.fill = fill(C["alt"])
        tc.number_format = fmt_str
        numfmt_id[c_idx] = tc._style.numFmtId

    cor_estilo = {}
    for c_idx, mapa in cores_col.items():
        fmt_str = fmt_col.get(c_idx)
        for valor, cor_hex in mapa.items():
            tc.fill = fill(cor_hex)
            tc.number_format = fmt_str if fmt_str else "General"
            cor_estilo[(c_idx, valor)] = copy(tc._style)

    wb.remove(tmp)
    cache = {}

    def estilo(alt, numfmt):
        k = (alt, numfmt)
        s = cache.get(k)
        if s is None:
            s = copy(base[alt])
            s.numFmtId = numfmt
            cache[k] = s
        return s

    ws._o60_df = df
    ws._o60_estilo = estilo
    ws._o60_fmtcol = numfmt_id
    ws._o60_corestilo = cor_estilo
    _dim = f"A1:{_gcl(ncols)}{max(2, len(df) + 2)}"
    ws.calculate_dimension = lambda: _dim

    for c_idx, col in enumerate(df.columns, 1):
        max_len = max(len(str(col)), df[col].astype(str).str.len().max())
        ws.column_dimensions[_gcl(c_idx)].width = min(max_len + 3, 40)

    ws.auto_filter.ref = f"A2:{_gcl(ncols)}{len(df)+2}"


motor._aba_dados = _aba_dados_rapido


def _linhas_lazy(ws):
    df, estilo = ws._o60_df, ws._o60_estilo
    fmtcol = getattr(ws, "_o60_fmtcol", {})
    cor_estilo = getattr(ws, "_o60_corestilo", {})
    for r_idx, row in enumerate(df.itertuples(index=False), start=3):
        alt = (r_idx % 2 == 0)
        linha = []
        for c_idx, v in enumerate(row, 1):
            if v is pd.NaT or (isinstance(v, float) and v != v):
                v = None
            elif isinstance(v, str) and _ILLEGAL_RE.search(v):
                v = _ILLEGAL_RE.sub("", v)
            cell = _Cell(ws, row=r_idx, column=c_idx, value=v)
            if cell.data_type == "f":
                cell.data_type = "s"
            st_fixo = cor_estilo.get((c_idx, v))
            if st_fixo is not None:
                cell._style = st_fixo
            else:
                st = cell._style
                numfmt = fmtcol.get(c_idx, st.numFmtId if st is not None else 0)
                cell._style = estilo(alt, numfmt)
            linha.append(cell)
        yield r_idx, linha


from itertools import chain as _chain
from openpyxl.worksheet._writer import WorksheetWriter as _WW
_orig_rows = _WW.rows


def _rows_lazy(self):
    base = _orig_rows(self)
    if getattr(self.ws, "_o60_df", None) is None:
        return base
    return _chain(base, _linhas_lazy(self.ws))


_WW.rows = _rows_lazy


# ---------------------------------------------------------------------------
#  Execucao do motor original numa thread, com print() indo pro log da tela
# ---------------------------------------------------------------------------
class FilaEscrita(io.TextIOBase):
    """Redireciona print()/logging pra uma fila que a interface le e mostra."""
    def __init__(self, fila):
        self.fila = fila

    def write(self, s):
        if s and s.strip():
            self.fila.put(s)
        return len(s)

    def flush(self):
        pass


def rodar_relatorios(os_path, cart_path, aprov_path, saida_dir, fila, callback_fim):
    stdout_original = sys.stdout
    sys.stdout = FilaEscrita(fila)
    erro = None
    try:
        _read_cache.clear()
        saida_dir.mkdir(parents=True, exist_ok=True)

        motor.CONFIG["pasta_entrada"] = Path(os_path).parent
        motor.CONFIG["carteiras_file"] = Path(cart_path)
        motor.CONFIG["aprovadas_file"] = Path(aprov_path) if aprov_path else Path("__sem_aprovadas__.xlsx")
        motor.CONFIG["pasta_saida"] = saida_dir
        motor.escolher_arquivo = lambda: str(os_path)
        motor.input = lambda *a, **k: ""

        motor.main()
    except Exception:
        erro = traceback.format_exc()
        print(erro)
    finally:
        _read_cache.clear()
        sys.stdout = stdout_original
        fila.put(("__FIM__", erro))
        callback_fim(erro)


# ---------------------------------------------------------------------------
#  Interface
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Relatório OS 60+ — Prime")
        self.geometry("760x560")
        self.minsize(640, 480)

        self.cfg = carregar_config()
        self.var_os = tk.StringVar(value=self.cfg.get("os", ""))
        self.var_cart = tk.StringVar(value=self.cfg.get("cart", ""))
        self.var_aprov = tk.StringVar(value=self.cfg.get("aprov", ""))
        self.var_saida = tk.StringVar(
            value=self.cfg.get("saida", str(Path.home() / "Documents" / "Relatorios_OS60"))
        )

        self.fila = queue.Queue()
        self.rodando = False

        self._montar_ui()
        self.after(150, self._checar_fila)

    # -- UI -----------------------------------------------------------
    def _montar_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        self._linha_arquivo(frm, 0, "Planilha de OS (.xlsx) *", self.var_os, self._escolher_os)
        self._linha_arquivo(frm, 1, "Carteiras Gerenciadoras (.xlsx) *", self.var_cart, self._escolher_cart)
        self._linha_arquivo(frm, 2, "Aprovadas sem movimentação (.xlsx)", self.var_aprov, self._escolher_aprov, opcional=True)
        self._linha_arquivo(frm, 3, "Pasta de saída dos relatórios", self.var_saida, self._escolher_saida, pasta=True)

        frm.columnconfigure(1, weight=1)

        botoes = ttk.Frame(self)
        botoes.pack(fill="x", **pad)
        self.btn_gerar = ttk.Button(botoes, text="Gerar relatórios", command=self._gerar)
        self.btn_gerar.pack(side="left")
        self.btn_abrir = ttk.Button(botoes, text="Abrir pasta de saída", command=self._abrir_pasta, state="disabled")
        self.btn_abrir.pack(side="left", padx=8)

        self.status = ttk.Label(self, text="Pronto.", foreground="#444")
        self.status.pack(fill="x", padx=12)

        self.log = tk.Text(self, height=20, font=("Consolas", 9), bg="#0d0d0d", fg="#d8d8d8", insertbackground="#fff")
        self.log.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log.configure(state="disabled")

    def _linha_arquivo(self, parent, row, label, var, comando, opcional=False, pasta=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ent = ttk.Entry(parent, textvariable=var)
        ent.grid(row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="Escolher pasta..." if pasta else "Escolher...", command=comando).grid(row=row, column=2)
        if opcional:
            ttk.Button(parent, text="Limpar", command=lambda: var.set("")).grid(row=row, column=3, padx=(4, 0))

    # -- seletores ------------------------------------------------------
    def _escolher_os(self):
        p = filedialog.askopenfilename(title="Planilha de OS", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_os.set(p)

    def _escolher_cart(self):
        p = filedialog.askopenfilename(title="Carteiras Gerenciadoras", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_cart.set(p)

    def _escolher_aprov(self):
        p = filedialog.askopenfilename(title="Aprovadas sem movimentação", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_aprov.set(p)

    def _escolher_saida(self):
        p = filedialog.askdirectory(title="Pasta de saída dos relatórios")
        if p:
            self.var_saida.set(p)

    def _abrir_pasta(self):
        caminho = self.var_saida.get()
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)
            elif sys.platform == "darwin":
                os.system(f'open "{caminho}"')
            else:
                os.system(f'xdg-open "{caminho}"')
        except Exception as e:
            messagebox.showerror("Erro", f"Não consegui abrir a pasta:\n{e}")

    # -- geração ----------------------------------------------------
    def _gerar(self):
        if self.rodando:
            return
        os_path = self.var_os.get().strip()
        cart_path = self.var_cart.get().strip()
        aprov_path = self.var_aprov.get().strip()
        saida = self.var_saida.get().strip()

        if not os_path or not Path(os_path).exists():
            messagebox.showwarning("Faltou arquivo", "Escolha a planilha de OS.")
            return
        if not cart_path or not Path(cart_path).exists():
            messagebox.showwarning("Faltou arquivo", "Escolha a planilha de Carteiras Gerenciadoras.")
            return
        if aprov_path and not Path(aprov_path).exists():
            messagebox.showwarning("Arquivo não encontrado", "O arquivo de Aprovadas indicado não existe.")
            return
        if not saida:
            messagebox.showwarning("Faltou pasta", "Escolha a pasta de saída dos relatórios.")
            return

        salvar_config({"os": os_path, "cart": cart_path, "aprov": aprov_path, "saida": saida})

        self.rodando = True
        self.btn_gerar.configure(state="disabled")
        self.btn_abrir.configure(state="disabled")
        self.status.configure(text="Gerando relatórios... isso pode levar alguns minutos.")
        self._limpar_log()

        t = threading.Thread(
            target=rodar_relatorios,
            args=(os_path, cart_path, aprov_path or None, Path(saida), self.fila, self._ao_terminar),
            daemon=True,
        )
        t.start()

    def _ao_terminar(self, erro):
        pass  # o encerramento real é tratado em _checar_fila (thread principal)

    def _limpar_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _escrever_log(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto if texto.endswith("\n") else texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _checar_fila(self):
        try:
            while True:
                item = self.fila.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__FIM__":
                    self._finalizar(item[1])
                else:
                    self._escrever_log(str(item))
        except queue.Empty:
            pass
        self.after(150, self._checar_fila)

    def _finalizar(self, erro):
        self.rodando = False
        self.btn_gerar.configure(state="normal")
        if erro:
            self.status.configure(text="Deu erro — veja o final do log.")
            messagebox.showerror("Erro ao gerar relatórios", "Confira o log na tela para os detalhes.")
        else:
            self.status.configure(text=f"Concluído — {dt.datetime.now():%d/%m/%Y %H:%M}.")
            self.btn_abrir.configure(state="normal")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

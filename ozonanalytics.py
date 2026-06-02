from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# ══════════════════════════════════════════════════════════════════════════════
# НАСТРОЙКА ДИНАМИЧЕСКИХ ПУТЕЙ (Для совместимости с Windows/macOS/GitHub)
# ══════════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).resolve().parent
COSTS_FILE = BASE_DIR / "costs_db.json"

def load_costs() -> dict:
    if COSTS_FILE.exists():
        try:
            with open(COSTS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_costs(costs: dict):
    try:
        with open(COSTS_FILE, "w", encoding="utf-8") as f:
            json.dump(costs, f, ensure_ascii=False, indent=4)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось сохранить себестоимость: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# ПАРСЕРЫ
# ══════════════════════════════════════════════════════════════════════════════

def parse_accrual_excel(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str)
    header_row = None
    for i, row in raw.iterrows():
        if row.astype(str).str.contains("ID начисления", na=False).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("Не найдена строка с заголовками ('ID начисления')")
    df = pd.read_excel(path, header=header_row, dtype=str)
    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()
    return df


def parse_goods_excel(path: str) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, dtype=str, nrows=10)
    header_row = 0
    for i, row in raw.iterrows():
        non_empty = row.dropna().astype(str).str.strip()
        non_empty = non_empty[non_empty != ""]
        if len(non_empty) >= 3:
            header_row = i
            break
    df = pd.read_excel(path, header=header_row, dtype=str)
    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()
    return df


# ══════════════════════════════════════════════════════════════════════════════
# ЛОГИКА ГРУППИРОВКИ И РАСЧЕТОВ
# ══════════════════════════════════════════════════════════════════════════════

def parse_amount(s) -> float:
    if pd.isna(s):
        return 0.0
    s = str(s).replace("₽", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def build_accrual_summary(df: pd.DataFrame) -> pd.DataFrame:
    amount_col = "Сумма итого, руб."
    id_col     = "ID начисления"
    group_col  = "Группа услуг"
    sku_col    = "SKU"
    name_col   = "Название товара"
    art_col    = "Артикул"
    date_col   = "Дата начисления"
    schema_col = "Схема работы"
    qty_col    = "Количество"

    df = df.copy()
    df[id_col]    = df[id_col].fillna("").astype(str).str.strip()
    df["_amount"] = df[amount_col].apply(parse_amount)
    df["_qty"]    = df[qty_col].apply(lambda x: int(parse_amount(x)))

    # Динамические колонки услуг
    pivot = (
        df[df[id_col] != ""]
        .pivot_table(index=id_col, columns=group_col, values="_amount",
                     aggfunc="sum", fill_value=0.0)
        .reset_index()
    )

    def first_val(x):
        v = x[x.notna() & (x.astype(str).str.strip() != "")]
        return v.iloc[0] if len(v) else ""

    meta = (
        df[df[id_col] != ""]
        .sort_values([id_col, sku_col])
        .groupby(id_col, sort=False)
        .agg(**{
            date_col:   (date_col,   "first"),
            art_col:    (art_col,    first_val),
            sku_col:    (sku_col,    first_val),
            name_col:   (name_col,   first_val),
            schema_col: (schema_col, lambda x: "/".join(sorted(x.dropna().unique()))),
            qty_col:    ("_qty",     "max"),
        })
        .reset_index()
    )

    result = meta.merge(pivot, on=id_col, how="left")
    result["Выплата Ozon"] = result[[c for c in pivot.columns if c != id_col]].sum(axis=1)
    
    costs_db = load_costs()
    
    def calculate_total_cost(row):
        art = str(row[art_col]).strip()
        sku = str(row[sku_col]).strip()
        qty = row[qty_col]
        unit_cost = 0.0
        if art in costs_db:
            unit_cost = costs_db[art]
        elif sku in costs_db:
            unit_cost = costs_db[sku]
        return unit_cost * qty

    result["Себестоимость (всего)"] = result.apply(calculate_total_cost, axis=1)
    
    if "Продажи" in result.columns:
        result.loc[result["Продажи"] == 0, "Себестоимость (всего)"] = 0.0
    
    result.loc[result["Выплата Ozon"] <= 0, "Себестоимость (всего)"] = 0.0
    result["ЧИСТАЯ ПРИБЫЛЬ"] = result["Выплата Ozon"] - result["Себестоимость (всего)"]
    
    def calculate_margin_pct(row):
        payout = row["Выплата Ozon"]
        profit = row["ЧИСТАЯ ПРИБЫЛЬ"]
        if payout > 0:
            return (profit / payout) * 100
        return 0.0

    result["Прибыль, %"] = result.apply(calculate_margin_pct, axis=1)
    return result.sort_values(date_col).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════════════════
# КОНСТАНТЫ И СТИЛИ ИНТЕРФЕЙСА
# ══════════════════════════════════════════════════════════════════════════════

LEFT_COLS = {
    "Группировка / SKU", "ID начисления", "Дата начисления", "Артикул", "SKU",
    "Название товара", "Схема работы", "Группа услуг",
}
NON_MONEY = {
    "Группировка / SKU", "ID начисления", "Дата начисления", "Артикул", "SKU",
    "Название товара", "Схема работы", "Количество", "Группа услуг", "Прибыль, %",
}

def create_button(master, text, command, **kwargs) -> tk.Button:
    bg_color = kwargs.pop("bg", "#e1e1e1")
    fg_color = kwargs.pop("fg", "black")
    font = kwargs.pop("font", ("", 10))
    
    btn = tk.Button(
        master, 
        text=text, 
        command=command,
        bg=bg_color,
        fg=fg_color,
        font=font,
        relief="flat",
        bd=0,
        highlightthickness=0,
        activebackground="#cccccc",
        **kwargs
    )
    return btn


# ══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА НАЧИСЛЕНИЙ (С ГРУППИРОВКОЙ, АВТОШИРИНОЙ И СТРОКОЙ ИТОГО)
# ══════════════════════════════════════════════════════════════════════════════

class AccrualTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)
        self._df_raw: pd.DataFrame | None = None
        self._df_summary: pd.DataFrame | None = None
        self._sort_col: str | None = None
        self._sort_asc: bool = True
        self._expanded_skus = set()
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        tb = ttk.Frame(self)
        tb.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        
        create_button(tb, text="Экспорт CSV", command=self._export_csv, width=12, padx=5, pady=2).pack(side="left", padx=2)
        create_button(tb, text="🖨️ Печать", command=self._open_print_form, width=12, padx=5, pady=2, bg="#3498db", fg="white").pack(side="left", padx=2)
        
        ttk.Label(tb, text="Фильтр:").pack(side="left", padx=(10, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter())
        ttk.Entry(tb, textvariable=self._search_var, width=28).pack(side="left")
        self._status = tk.StringVar(value="")
        ttk.Label(tb, textvariable=self._status, foreground="gray").pack(side="right", padx=6)

        paned = ttk.PanedWindow(self, orient="vertical")
        paned.grid(row=1, column=0, sticky="nsew")

        top = ttk.LabelFrame(paned, text="Сводка по начислениям (Группировка по SKU)")
        paned.add(top, weight=3)
        top.columnconfigure(0, weight=1)
        top.rowconfigure(0, weight=1)

        self._tree_summary = ttk.Treeview(top, show="headings", selectmode="browse")
        vsb1 = ttk.Scrollbar(top, orient="vertical",   command=self._tree_summary.yview)
        hsb1 = ttk.Scrollbar(top, orient="horizontal", command=self._tree_summary.xview)
        self._tree_summary.configure(yscrollcommand=vsb1.set, xscrollcommand=hsb1.set)
        self._tree_summary.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")
        hsb1.grid(row=1, column=0, sticky="ew")
        
        self._tree_summary.bind("<ButtonRelease-1>", self._on_click)

        bot_nb = ttk.Notebook(paned)
        paned.add(bot_nb, weight=2)

        detail_frame = ttk.Frame(bot_nb)
        bot_nb.add(detail_frame, text="Детали начислений")
        detail_frame.columnconfigure(0, weight=1)
        detail_frame.rowconfigure(0, weight=1)

        self._tree_detail = ttk.Treeview(detail_frame, show="headings", selectmode="browse")
        vsb2 = ttk.Scrollbar(detail_frame, orient="vertical",   command=self._tree_detail.yview)
        hsb2 = ttk.Scrollbar(detail_frame, orient="horizontal", command=self._tree_detail.xview)
        self._tree_detail.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        self._tree_detail.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")
        hsb2.grid(row=1, column=0, sticky="ew")

        chart_frame = ttk.Frame(bot_nb)
        bot_nb.add(chart_frame, text="Графики")
        self._fig = Figure(figsize=(10, 3), dpi=90)
        self._canvas = FigureCanvasTkAgg(self._fig, master=chart_frame)
        self._canvas.get_tk_widget().pack(fill="both", expand=True)

    def load(self, df_raw: pd.DataFrame):
        self._df_raw = df_raw
        self._df_summary = build_accrual_summary(df_raw)
        self._render_summary()
        self._render_charts(self._df_summary)

    def _render_summary(self):
        if self._df_summary is None:
            return
            
        tree = self._tree_summary
        tree.delete(*tree.get_children())
        
        base_cols = list(self._df_summary.columns)
        if "ID начисления" in base_cols:
            base_cols.remove("ID начисления")
        
        cols = ["Группировка / SKU"] + base_cols
        tree["columns"] = cols

        for c in cols:
            tree.heading(c, text=c, command=lambda _c=c: self._sort(_c))
        
        grouped = self._df_summary.groupby("SKU")
        money_cols = {c for c in cols if c not in NON_MONEY}
        col_widths = {c: len(str(c)) for c in cols}

        total_all_payout = 0.0
        total_all_cost = 0.0
        total_all_profit = 0.0

        for sku, group in grouped:
            total_qty = group["Количество"].sum()
            total_payout = group["Выплата Ozon"].sum()
            total_cost = group["Себестоимость (всего)"].sum()
            total_profit = group["ЧИСТАЯ ПРИБЫЛЬ"].sum()
            margin = (total_profit / total_payout * 100) if total_payout > 0 else 0.0
            
            total_all_payout += total_payout
            total_all_cost += total_cost
            total_all_profit += total_profit
            
            first_row = group.iloc[0]
            prefix = "▼ " if sku in self._expanded_skus else "▶ "
            
            sku_values = {
                "Группировка / SKU": f"{prefix}{sku}",
                "Дата начисления": "—",
                "Артикул": first_row.get("Артикул", ""),
                "SKU": sku,
                "Название товара": first_row.get("Название товара", ""),
                "Схема работы": "—",
                "Количество": total_qty,
                "Выплата Ozon": total_payout,
                "Себестоимость (всего)": total_cost,
                "ЧИСТАЯ ПРИБЫЛЬ": total_profit,
                "Прибыль, %": margin
            }
            
            for c in cols:
                if c not in sku_values and c in group.columns:
                    sku_values[c] = group[c].sum()

            row_values = []
            for c in cols:
                v = sku_values.get(c, "")
                if c in money_cols and isinstance(v, (int, float)):
                    f_val = f"{v:,.2f}"
                elif c == "Прибыль, %" and isinstance(v, (int, float)):
                    f_val = f"{v:.2f}%"
                else:
                    f_val = "" if pd.isna(v) else str(v)
                
                row_values.append(f_val)
                if len(f_val) > col_widths[c]:
                    col_widths[c] = len(f_val)

            sku_tag = "sku_neg" if total_profit < 0 else "sku_group"
            parent_id = f"group_{sku}"
            tree.insert("", "end", iid=parent_id, values=row_values, tags=(sku_tag,))

            if sku in self._expanded_skus:
                for _, child_row in group.iterrows():
                    child_values = {
                        "Группировка / SKU": f"    ID: {child_row['ID начисления']}",
                        "Дата начисления": child_row.get("Дата начисления", ""),
                        "Артикул": child_row.get("Артикул", ""),
                        "SKU": sku,
                        "Название товара": child_row.get("Название товара", ""),
                        "Схема работы": child_row.get("Схема работы", ""),
                        "Количество": child_row["Количество"],
                        "Выплата Ozon": child_row["Выплата Ozon"],
                        "Себестоимость (всего)": child_row["Себестоимость (всего)"],
                        "ЧИСТАЯ ПРИБЫЛЬ": child_row["ЧИСТАЯ ПРИБЫЛЬ"],
                        "Прибыль, %": child_row["Прибыль, %"]
                    }
                    for c in cols:
                        if c not in child_values and c in child_row:
                            child_values[c] = child_row[c]

                    c_row_values = []
                    for c in cols:
                        v = child_values.get(c, "")
                        if c in money_cols and isinstance(v, (int, float)):
                            f_val = f"{v:,.2f}"
                        elif c == "Прибыль, %" and isinstance(v, (int, float)):
                            f_val = f"{v:.2f}%"
                        else:
                            f_val = "" if pd.isna(v) else str(v)
                        c_row_values.append(f_val)
                        if len(f_val) > col_widths[c]:
                            col_widths[c] = len(f_val)

                    c_tag = "neg" if child_row["ЧИСТАЯ ПРИБЫЛЬ"] < 0 else "child"
                    tree.insert("", "end", iid=str(child_row["ID начисления"]), values=c_row_values, tags=(c_tag,))

        if not self._df_summary.empty:
            all_margin = (total_all_profit / total_all_payout * 100) if total_all_payout > 0 else 0.0
            
            total_values = {c: "" for c in cols}
            total_values["Группировка / SKU"] = "ИТОГО ПО ВСЕМ ТОВАРAM:"
            total_values["Выплата Ozon"] = total_all_payout
            total_values["Себестоимость (всего)"] = total_all_cost
            total_values["ЧИСТАЯ ПРИБЫЛЬ"] = total_all_profit
            total_values["Прибыль, %"] = all_margin
            
            if "Продажи" in self._df_summary.columns:
                total_values["Продажи"] = self._df_summary["Продажи"].sum()

            total_row_vals = []
            for c in cols:
                v = total_values.get(c, "")
                if c in ["Выплата Ozon", "Себестоимость (всего)", "ЧИСТАЯ ПРИБЫЛЬ", "Продажи"] and isinstance(v, (int, float)):
                    f_val = f"{v:,.2f}"
                elif c == "Прибыль, %" and isinstance(v, (int, float)):
                    f_val = f"{v:.2f}%"
                else:
                    f_val = str(v)
                total_row_vals.append(f_val)
            
            tree.insert("", "end", iid="total_row_summary", values=total_row_vals, tags=("total_summary",))

        for c in cols:
            calculated_width = max(col_widths[c] * 8 + 20, 100)
            if c == "Название товара":
                calculated_width = min(calculated_width, 400)
            tree.column(c, width=calculated_width, anchor="w" if c in LEFT_COLS else "e", stretch=False)

        tree.tag_configure("sku_group", font=("", 10, "bold"), background="#f5f6fa")
        tree.tag_configure("sku_neg", font=("", 10, "bold"), background="#f5f6fa", foreground="#c0392b")
        tree.tag_configure("neg", foreground="#c0392b")
        tree.tag_configure("child", foreground="#2f3640")
        tree.tag_configure("total_summary", font=("", 10, "bold"), background="#dcdde1", foreground="black")
        
        self._status.set(f"{len(self._df_summary)} уникальных ID")

    def _on_click(self, event):
        item_id = self._tree_summary.focus()
        if not item_id or self._df_raw is None:
            return

        if item_id == "total_row_summary":
            return

        if item_id.startswith("group_"):
            sku = item_id.replace("group_", "")
            if sku in self._expanded_skus:
                self._expanded_skus.remove(sku)
            else:
                self._expanded_skus.add(sku)
            self._render_summary()
        else:
            df_detail = self._df_raw[
                self._df_raw["ID начисления"].astype(str).str.strip() == item_id
            ].copy()
            self._render_detail(df_detail)

    def _render_detail(self, df: pd.DataFrame):
        tree = self._tree_detail
        tree.delete(*tree.get_children())
        non_empty = [c for c in df.columns
                     if df[c].notna().any() and (df[c].astype(str).str.strip() != "").any()]
        df = df[non_empty]
        cols = list(df.columns)
        tree["columns"] = cols
        
        col_widths = {c: len(str(c)) for c in cols}
        for c in cols:
            tree.heading(c, text=c)

        for _, row in df.iterrows():
            vals = []
            for c in cols:
                v = "" if pd.isna(row[c]) else str(row[c])
                vals.append(v)
                if len(v) > col_widths[c]:
                    col_widths[c] = len(v)
            tree.insert("", "end", values=vals)

        for c in cols:
            tree.column(c, width=max(col_widths[c] * 8 + 20, 110), anchor="w" if c in LEFT_COLS else "e", stretch=False)

    def _render_charts(self, df: pd.DataFrame):
        if df is None or df.empty:
            return
        self._fig.clear()
        
        exclude = {
            "ID начисления", "Дата начисления", "Артикул", "SKU", "Группировка / SKU",
            "Название товара", "Схема работы", "Количество", "Выплата Ozon", 
            "Себестоимость (всего)", "ЧИСТАЯ ПРИБЫЛЬ", "Прибыль, %"
        }
        group_cols = [c for c in df.columns if c not in exclude]
        
        ax1 = self._fig.add_subplot(1, 2, 1)
        if group_cols:
            totals = df[group_cols].sum().sort_values()
            ax1.barh(range(len(totals)), totals.values,
                     color=["#e74c3c" if v < 0 else "#2ecc71" for v in totals.values])
            ax1.set_yticks(range(len(totals)))
            ax1.set_yticklabels(totals.index, fontsize=7)
            ax1.set_title("Суммы по группам услуг", fontsize=9)
            ax1.axvline(0, color="black", linewidth=0.5)

        ax2 = self._fig.add_subplot(1, 2, 2)
        if "Дата начисления" in df.columns:
            by_date = df.groupby("Дата начисления")["ЧИСТАЯ ПРИБЫЛЬ"].sum().reset_index().sort_values("Дата начисления")
            ax2.bar(range(len(by_date)), by_date["ЧИСТАЯ ПРИБЫЛЬ"],
                    color=["#e74c3c" if v < 0 else "#3498db" for v in by_date["ЧИСТАЯ ПРИБЫЛЬ"]])
            ax2.set_xticks(range(len(by_date)))
            ax2.set_xticklabels(by_date["Дата начисления"], rotation=45, ha="right", fontsize=6)
            ax2.set_title("ЧИСТАЯ ПРИБЫЛЬ по дате", fontsize=9)
            ax2.axhline(0, color="black", linewidth=0.5)

        self._fig.tight_layout()
        self._canvas.draw()

    def _sort(self, col: str):
        if self._df_summary is None:
            return
        asc = True if self._sort_col != col else not self._sort_asc
        df = self._df_summary.copy()
        try:
            df = df.sort_values(col, ascending=asc)
        except TypeError:
            df = df.sort_values(col, ascending=asc, key=lambda x: x.astype(str))
        self._sort_col = col
        self._sort_asc = asc
        self._df_summary = df
        self._render_summary()

    def _filter(self):
        if self._df_summary is None:
            return
        q = self._search_var.get().strip().lower()
        if not q:
            self._df_summary = build_accrual_summary(self._df_raw)
            self._render_summary()
            return
        mask = self._df_summary.apply(
            lambda r: r.astype(str).str.lower().str.contains(q).any(), axis=1)
        self._df_summary = self._df_summary[mask]
        self._render_summary()

    def _export_csv(self):
        if self._df_summary is None:
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV", "*.csv")],
                                         initialfile="accruals.csv")
        if p:
            self._df_summary.to_csv(p, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Экспорт", f"Сохранено: {p}")

    # ══════════════════════════════════════════════════════════════════════════
    # КРОССПЛАТФОРМЕННАЯ ФОРМА ДЛЯ ПЕЧАТИ (ОКНА WINDOWS / MAC С СОХРАНЕНИЕМ В PDF)
    # ══════════════════════════════════════════════════════════════════════════
    def _open_print_form(self):
        if self._df_summary is None or self._df_summary.empty:
            messagebox.showwarning("Внимание", "Нет данных для формирования печатной формы.")
            return

        print_window = tk.Toplevel(self)
        print_window.title("Форма для печати отчёта")
        print_window.geometry("1100x600")
        print_window.transient(self)
        print_window.grab_set()

        ptb = ttk.Frame(print_window)
        ptb.pack(fill="x", padx=10, pady=5)
        
        required_cols = ["Название товара", "Количество", "Выплата Ozon", "Себестоимость (всего)", "ЧИСТАЯ ПРИБЫЛЬ", "Прибыль, %"]
        has_sales = "Продажи" in self._df_summary.columns
        if has_sales:
            required_cols.insert(2, "Продажи")

        p_tree = ttk.Treeview(print_window, columns=required_cols, show="headings", selectmode="none")
        vsb = ttk.Scrollbar(print_window, orient="vertical", command=p_tree.yview)
        hsb = ttk.Scrollbar(print_window, orient="horizontal", command=p_tree.xview)
        p_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        p_tree.pack(fill="both", expand=True, padx=10, pady=(0, 5))
        vsb.pack(side="right", fill="y", before=p_tree)
        hsb.pack(fill="x", padx=10)

        for c in required_cols:
            p_tree.heading(c, text=c)
            p_tree.column(c, anchor="w" if c == "Название товара" else "e", stretch=True if c == "Название товара" else False, width=130)
        p_tree.column("Название товара", width=350)

        grouped = self._df_summary.groupby("SKU")
        
        sum_qty = 0
        sum_sales = 0.0
        sum_payout = 0.0
        sum_cost = 0.0
        sum_profit = 0.0

        html_rows = []

        for sku, group in grouped:
            q = group["Количество"].sum()
            p = group["Выплата Ozon"].sum()
            c = group["Себестоимость (всего)"].sum()
            pr = group["ЧИСТАЯ ПРИБЫЛЬ"].sum()
            m = (pr / p * 100) if p > 0 else 0.0

            sum_qty += q
            sum_payout += p
            sum_cost += c
            sum_profit += pr

            name_str = str(group.iloc[0]["Название товара"])
            vals = [name_str, str(q)]
            
            html_tds = f"<td>{name_str}</td><td>{q}</td>"
            
            if has_sales:
                s_val = group["Продажи"].sum()
                sum_sales += s_val
                vals.append(f"{s_val:,.2f}")
                html_tds += f"<td>{s_val:,.2f}</td>"
                
            vals.extend([f"{p:,.2f}", f"{c:,.2f}", f"{pr:,.2f}", f"{m:.2f}%"])
            html_tds += f"<td>{p:,.2f}</td><td>{c:,.2f}</td><td>{pr:,.2f}</td><td>{m:.2f}%</td>"
            
            p_tree.insert("", "end", values=vals)
            html_rows.append(f"<tr>{html_tds}</tr>")

        total_margin = (sum_profit / sum_payout * 100) if sum_payout > 0 else 0.0
        total_vals = ["ИТОГО ПО ВСЕМ ТОВАРAM:", str(sum_qty)]
        
        html_total_tds = f"<td class='bold'>ИТОГО ПО ВСЕМ ТОВАРAM:</td><td class='bold'>{sum_qty}</td>"
        
        if has_sales:
            total_vals.append(f"{sum_sales:,.2f}")
            html_total_tds += f"<td class='bold'>{sum_sales:,.2f}</td>"
            
        total_vals.extend([f"{sum_payout:,.2f}", f"{sum_cost:,.2f}", f"{sum_profit:,.2f}", f"{total_margin:.2f}%"])
        html_total_tds += f"<td class='bold'>{sum_payout:,.2f}</td><td class='bold'>{sum_cost:,.2f}</td><td class='bold'>{sum_profit:,.2f}</td><td class='bold'>{total_margin:.2f}%</td>"
        
        p_tree.insert("", "end", values=total_vals, tags=("total_print",))
        p_tree.column("Название товара", width=350)
        p_tree.tag_configure("total_print", font=("", 10, "bold"), background="#dcdde1")

        def sys_print():
            import tempfile
            import os
            import platform
            import subprocess
            
            th_elements = "".join([f"<th>{col}</th>" for col in required_cols])
            
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <title>Печать отчёта Ozon</title>
                <style>
                    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 20px; }}
                    h2 {{ text-align: center; margin-bottom: 20px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }}
                    th, td {{ border: 1px solid #111; padding: 6px 8px; text-align: left; }}
                    th {{ background-color: #f2f2f2; }}
                    td:not(:first-child), th:not(:first-child) {{ text-align: right; }}
                    .bold {{ font-weight: bold; background-color: #eaeaea; }}
                    @media print {{
                        button {{ display: none; }}
                    }}
                </style>
            </head>
            <body>
                <h2>Отчёт по начислениям Ozon</h2>
                <table>
                    <thead><tr>{th_elements}</tr></thead>
                    <tbody>
                        {"".join(html_rows)}
                        <tr class="bold">{html_total_tds}</tr>
                    </tbody>
                </table>
                <script>
                    window.onload = function() {{ 
                        window.print(); 
                    }}
                </script>
            </body>
            </html>
            """
            
            with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
                f.write(html_content)
                temp_path = f.name
                
            current_os = platform.system()
            if current_os == "Darwin":
                subprocess.run(["open", temp_path])
            elif current_os == "Windows":
                os.startfile(temp_path)
            else:
                subprocess.run(["xdg-open", temp_path])

        create_button(ptb, text="🖨️ Отправить на печать", command=sys_print, width=22, pady=4, bg="#2ecc71", fg="white").pack(side="left")
        create_button(ptb, text="Закрыть", command=print_window.destroy, width=12, pady=4).pack(side="right")

        lbl_info = ttk.Label(print_window, text="Печатный вид документа (Выводятся только основные экономические показатели)", font=("", 10, "italic"), foreground="gray")
        lbl_info.pack(anchor="w", padx=10, pady=(0, 5))


# ══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА ТОВАРОВ (С АВТОШИРИНОЙ СТОЛБЦОВ)
# ══════════════════════════════════════════════════════════════════════════════

class GoodsTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)
        self._df_full: pd.DataFrame | None = None
        self._sort_col: str | None = None
        self._sort_asc: bool = True
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        tb = ttk.Frame(self)
        tb.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        
        create_button(tb, text="Экспорт CSV", command=self._export_csv, width=12, padx=5, pady=2).pack(side="left", padx=2)
        
        ttk.Label(tb, text="Фильтр:").pack(side="left", padx=(10, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter())
        ttk.Entry(tb, textvariable=self._search_var, width=28).pack(side="left")
        self._status = tk.StringVar(value="")
        ttk.Label(tb, textvariable=self._status, foreground="gray").pack(side="right", padx=6)

        frame = ttk.Frame(self)
        frame.grid(row=1, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

    def load(self, df: pd.DataFrame):
        self._df_full = df
        self._render(df)

    def _render(self, df: pd.DataFrame):
        tree = self._tree
        tree.delete(*tree.get_children())
        cols = list(df.columns)
        tree["columns"] = cols
        
        col_widths = {c: len(str(c)) for c in cols}
        for c in cols:
            tree.heading(c, text=c, command=lambda _c=c: self._sort(_c))
            
        for _, row in df.iterrows():
            vals = []
            for c in cols:
                v = "" if pd.isna(row[c]) else str(row[c])
                vals.append(v)
                if len(v) > col_widths[c]:
                    col_widths[c] = len(v)
            tree.insert("", "end", values=vals)
            
        for c in cols:
            calculated_width = max(col_widths[c] * 8 + 20, 110)
            if "название" in c.lower():
                calculated_width = min(calculated_width, 400)
            tree.column(c, width=calculated_width, anchor="w", stretch=False)
            
        self._status.set(f"{len(df)} строк")

    def _sort(self, col: str):
        if self._df_full is None:
            return
        asc = True if self._sort_col != col else not self._sort_asc
        df = self._df_full.copy()
        try:
            df = df.sort_values(col, ascending=asc)
        except TypeError:
            df = df.sort_values(col, ascending=asc, key=lambda x: x.astype(str))
        self._sort_col = col
        self._sort_asc = asc
        self._render(df)

    def _filter(self):
        if self._df_full is None:
            return
        q = self._search_var.get().strip().lower()
        if not q:
            self._render(self._df_full)
            return
        mask = self._df_full.apply(
            lambda r: r.astype(str).str.lower().str.contains(q).any(), axis=1)
        self._render(self._df_full[mask])

    def _export_csv(self):
        if self._df_full is None:
            return
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV", "*.csv")],
                                         initialfile="goods.csv")
        if p:
            self._df_full.to_csv(p, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Экспорт", f"Сохранено: {p}")


# ══════════════════════════════════════════════════════════════════════════════
# ВКЛАДКА НАСТРОЙКИ СЕБЕСТОИМОСТИ
# ══════════════════════════════════════════════════════════════════════════════

class CostsTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)
        self._build_ui()
        self._reload_data()

    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        form = ttk.LabelFrame(self, text="Управление базой", padding=10)
        form.grid(row=0, column=0, sticky="nsw", padx=6, pady=6)

        create_button(form, text="📥   Импорт из Excel / CSV", command=self._import_from_file, width=25, pady=6, bg="#2ecc71", fg="white").pack(anchor="w", pady=(5, 20))
        ttk.Separator(form, orient="horizontal").pack(fill="x", pady=(0, 15))

        ttk.Label(form, text="Артикул или SKU:").pack(anchor="w", pady=(0, 2))
        self._sku_var = tk.StringVar()
        ttk.Entry(form, textvariable=self._sku_var, width=25).pack(anchor="w", pady=(0, 10))

        ttk.Label(form, text="Себестоимость (руб):").pack(anchor="w", pady=(0, 2))
        self._cost_var = tk.StringVar()
        ttk.Entry(form, textvariable=self._cost_var, width=25).pack(anchor="w", pady=(0, 15))

        create_button(form, text="Сохранить", command=self._save_entry, width=25, pady=4, bg="#3498db", fg="white").pack(anchor="w", pady=2)
        create_button(form, text="Удалить выбранное", command=self._delete_entry, width=25, pady=4, bg="#e74c3c", fg="white").pack(anchor="w", pady=10)

        table_frame = ttk.LabelFrame(self, text="Текущая база данных себестоимости", padding=6)
        table_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(table_frame, columns=("sku", "cost"), show="headings", selectmode="browse")
        self._tree.heading("sku", text="Артикул / SKU")
        self._tree.heading("cost", text="Себестоимость, руб.")
        self._tree.column("sku", width=250, anchor="w")
        self._tree.column("cost", width=150, anchor="e")
        
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

    def _reload_data(self):
        self._tree.delete(*self._tree.get_children())
        db = load_costs()
        for sku, cost in sorted(db.items()):
            self._tree.insert("", "end", iid=sku, values=(sku, f"{cost:,.2f}"))

    def _on_select(self, _event):
        sel = self._tree.selection()
        if sel:
            sku = sel[0]
            db = load_costs()
            if sku in db:
                self._sku_var.set(sku)
                self._cost_var.set(str(db[sku]))

    def _save_entry(self):
        sku = self._sku_var.get().strip()
        cost_str = self._cost_var.get().strip().replace(" ", "").replace(",", ".")
        
        if not sku:
            messagebox.showwarning("Внимание", "Введите артикул или SKU")
            return
        try:
            cost = float(cost_str)
            if cost < 0: raise ValueError
        except ValueError:
            messagebox.showwarning("Внимание", "Введите корректную цену")
            return

        db = load_costs()
        db[sku] = cost
        save_costs(db)
        
        self._sku_var.set("")
        self._cost_var.set("")
        self._reload_data()

    def _delete_entry(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите товар для удаления")
            return
        sku = sel[0]
        if messagebox.askyesno("Удаление", f"Удалить себестоимость для {sku}?"):
            db = load_costs()
            if sku in db:
                del db[sku]
                save_costs(db)
                self._sku_var.set("")
                self._cost_var.set("")
                self._reload_data()

    def _import_from_file(self):
        path = filedialog.askopenfilename(
            title="Выбрать файл себестоимости",
            filetypes=[("Excel/CSV файлы", "*.xlsx *.xls *.csv"), ("Все файлы", "*.*")]
        )
        if not path:
            return

        try:
            if path.endswith(".csv"):
                df = pd.read_csv(path, dtype=str)
            else:
                df = pd.read_excel(path, dtype=str)

            if df.empty:
                raise ValueError("Файл пуст")

            df.columns = [str(c).strip().lower() for c in df.columns]
            
            sku_col = None
            cost_col = None

            for c in df.columns:
                if any(x in c for x in ["sku", "артикул", "id"]):
                    sku_col = c
                    break
            for c in df.columns:
                if any(x in c for x in ["себестоимость", "цена", "cost"]):
                    cost_col = c
                    break

            if sku_col is None or cost_col is None:
                sku_col = df.columns[0]
                cost_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

            db = load_costs()
            count = 0

            for _, row in df.iterrows():
                sku_val = str(row[sku_col]).strip()
                cost_val_str = str(row[cost_col]).strip()

                if pd.isna(row[sku_col]) or sku_val == "" or sku_val.lower() == "nan":
                    continue

                cost_val_str = cost_val_str.replace("₽", "").replace("\xa0", "").replace(" ", "").replace(",", ".")
                try:
                    cost_val = float(cost_val_str)
                    db[sku_val] = cost_val
                    count += 1
                except ValueError:
                    continue

            save_costs(db)
            self._reload_data()
            messagebox.showinfo("Успех", f"Успешно импортировано товаров: {count}")

        except Exception as e:
            messagebox.showerror("Ошибка импорта", f"Не удалось прочитать файл:\n{e}")


# ══════════════════════════════════════════════════════════════════════════════
# СТАРТОВЫЙ ЭКРАН
# ══════════════════════════════════════════════════════════════════════════════

class ReportTypeSelector(ttk.Frame):

    def __init__(self, parent, on_select: callable):
        super().__init__(parent)
        self._on_select = on_select
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        center = ttk.Frame(self)
        center.grid(row=0, column=0)

        ttk.Label(center, text="Аналитика Ozon", font=("", 22, "bold")).pack(pady=(40, 8))
        ttk.Label(center, text="Выберите тип отчёта или настройте базу данных", font=("", 12),
                  foreground="gray").pack(pady=(0, 40))

        create_button(center, text="📦   По товарам", command=lambda: self._pick("goods"),
                      width=30, pady=12, font=("", 12), bg="#34495e", fg="white").pack(pady=8)

        create_button(center, text="💰   По начислениям", command=lambda: self._pick("accruals"),
                      width=30, pady=12, font=("", 12), bg="#2c3e50", fg="white").pack(pady=8)

        ttk.Separator(center, orient="horizontal").pack(fill="x", pady=20)

        create_button(center, text="⚙️   Настройка себестоимости", command=lambda: self._pick("costs"),
                      width=30, pady=10, font=("", 11), bg="#7f8c8d", fg="white").pack(pady=5)

    def _pick(self, report_type: str):
        if report_type == "costs":
            self._on_select(report_type, "")
            return

        filetypes = [("Excel 2007+", "*.xlsx"), ("Excel 97-2003", "*.xls"), ("Все файлы", "*.*")]
        if report_type == "goods":
            filetypes.insert(2, ("CSV", "*.csv"))
        path = filedialog.askopenfilename(title="Открыть отчёт", filetypes=filetypes)
        if path:
            self._on_select(report_type, path)


# ══════════════════════════════════════════════════════════════════════════════
# ГЛАВНОЕ ОКНО
# ══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Аналитика Ozon")
        self.geometry("1400x800")
        self.minsize(900, 500)
        self._show_selector()

    def _show_selector(self):
        self._clear()
        ReportTypeSelector(self, on_select=self._on_selected).pack(fill="both", expand=True)

    def _on_selected(self, report_type: str, path: str):
        self._clear()

        header = ttk.Frame(self)
        header.pack(fill="x", padx=6, pady=(6, 0))
        
        create_button(header, text="← Назад", command=self._show_selector, width=10, pady=3).pack(side="left")
        
        if report_type == "costs":
            label = "Редактирование себестоимости"
            ttk.Label(header, text=label, font=("", 10, "bold")).pack(side="left", padx=12)
        else:
            label = "По товарам" if report_type == "goods" else "По начислениям"
            fname = path.replace("\\", "/").split("/")[-1]
            ttk.Label(header, text=f"{label}  ·  {fname}", font=("", 10)).pack(side="left", padx=12)
            self._loading_label = ttk.Label(header, text="⏳ Загрузка...", foreground="gray")
            self._loading_label.pack(side="right", padx=10)

        self._content_frame = ttk.Frame(self)
        self._content_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self._content_frame.columnconfigure(0, weight=1)
        self._content_frame.rowconfigure(0, weight=1)

        if report_type == "costs":
            tab = CostsTab(self._content_frame)
            tab.grid(row=0, column=0, sticky="nsew")
        else:
            import threading
            threading.Thread(target=self._load_worker, args=(report_type, path), daemon=True).start()

    def _load_worker(self, report_type: str, path: str):
        try:
            if report_type == "accruals":
                df_raw = parse_accrual_excel(path)
                self.after(0, lambda: self._show_accruals(df_raw))
            else:
                df = parse_goods_excel(path)
                self.after(0, lambda: self._show_goods(df))
        except Exception as e:
            self.after(0, lambda error=e: self._on_load_error(str(error)))

    def _show_accruals(self, df_raw: pd.DataFrame):
        if hasattr(self, '_loading_label') and self._loading_label.winfo_exists():
            self._loading_label.config(text="")
        tab = AccrualTab(self._content_frame)
        tab.grid(row=0, column=0, sticky="nsew")
        tab.load(df_raw)

    def _show_goods(self, df: pd.DataFrame):
        if hasattr(self, '_loading_label') and self._loading_label.winfo_exists():
            self._loading_label.config(text="")
        tab = GoodsTab(self._content_frame)
        tab.grid(row=0, column=0, sticky="nsew")
        tab.load(df)

    def _on_load_error(self, msg: str):
        messagebox.showerror("Ошибка загрузки", msg)
        self._show_selector()

    def _clear(self):
        for w in self.winfo_children():
            w.destroy()


if __name__ == "__main__":
    App().mainloop()
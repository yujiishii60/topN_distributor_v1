# -*- coding: utf-8 -*-
"""GUI launcher (完全版) for scripts.make_topn_simple_refactor

前提:
- 実行モジュールは `scripts.make_topn_simple_refactor` に一元化済み。
- テンプレ指定は CLI 側に非公開（内部で既定テンプレを使用する想定）。

GUIで設定できる引数:
- 必須:  --category, --dates, --out
- 任意:  --event-name, --title-template, --no-date-in-title, --split-by-store, --split-dir

付加機能:
- 事前チェック: dates から必要な CSV (data/material/YYYY/IT_YYYYMM.csv) の存在確認
- 実行ログ（UTF-8強制）
- 完了後のポストチェック（本体xlsx存在 / split件数）
- 便利: 完了後に split / 本体 xlsx を自動オープン
"""
from __future__ import annotations
import os, sys, subprocess, threading, queue, re
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from string import Template
import json
from datetime import datetime
try:
    from tkcalendar import Calendar
except Exception:
    Calendar = None

# ← ここまではOK（インポートだけ）
# --- add once at the very top ---
import sys, pathlib
sys.path.append(str(pathlib.Path(__file__).resolve().parents[1]))
# ---------------------------------

from styles import theme_cyber, theme_pastel
from styles.apply_ttk_min import apply_theme
from styles.widgets import make_calendar, style_toplevel

APP_PATH = Path(__file__).resolve()
REPO_ROOT = APP_PATH.parent.parent
CLI_SIMPLE = 'scripts.make_topn_simple_refactor'
MATERIAL_DIR = REPO_ROOT / 'data' / 'material'
STORE_MASTER = MATERIAL_DIR / 'master' / 'store_master.xlsx'

class TopNGuiApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title('topN 配布ツール（simple_refactor 完全版）')
        self.root.geometry('980x700')
        # 入力
        self.var_event = tk.StringVar()
        self.var_category = tk.StringVar(value='4')
        self.var_dates = tk.StringVar()
        self.var_out = tk.StringVar(value=str(REPO_ROOT / 'data/output/topN_冷総菜.xlsx'))
        self.var_split = tk.BooleanVar(value=True)
        self.var_split_dir = tk.StringVar(value=str(REPO_ROOT / 'data/output/split'))
        self.var_title_template = tk.StringVar(
            value='{yy}年 {range} {cat}単品データ（{page}）'
        )
        self.var_no_date_in_title = tk.BooleanVar(value=False)
        # カテゴリマスタ読込
        self.category_map: dict[str, str] = self._load_category_map()

        # 既定コードがマスタに無ければ先頭に寄せる
        if self.var_category.get() not in self.category_map:
            try:
                first_code = sorted(self.category_map.keys(), key=lambda x: int(x))[0]
            except Exception:
                first_code = next(iter(self.category_map.keys()))
            self.var_category.set(first_code)
        # 便利
        self.var_open_after_split = tk.BooleanVar(value=True)
        self.var_open_after_main = tk.BooleanVar(value=False)
        # 実行
        self.proc: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        # UI
        self._build_ui()
        self._poll_log_queue()

        # __init__ の最後あたりに追記（ボタン押下なしでも即反映）
        for v in (self.var_event, self.var_category, self.var_dates,
                  self.var_title_template, self.var_no_date_in_title):
            v.trace_add("write", lambda *_: self.on_preview_title())
        self.on_preview_title()

    # ===== UI構築 =====
    def _build_ui(self):
        pad = {'padx':10, 'pady':6}
        frm = ttk.Frame(self.root); frm.pack(fill=tk.BOTH, expand=True)

        def row(label, var, browse=None, width=14, entry_w=None):
            f = ttk.Frame(frm); f.pack(fill=tk.X, **pad)
            ttk.Label(f, text=label, width=width).pack(side=tk.LEFT)
            e = ttk.Entry(f, textvariable=var, width=entry_w)
            e.pack(side=tk.LEFT, fill=tk.X, expand=True)
            if browse: ttk.Button(f, text='参照…', command=browse).pack(side=tk.LEFT, padx=6)
            return f

        
        # --- ここを変更（★イベント名の行にボタン追加） ---
        f_event = row('イベント名', self.var_event)
        ttk.Button(f_event, text='候補プレビュー', command=self.on_preview_title)\
        .pack(side=tk.LEFT, padx=6)
        ttk.Button(f_event, text='候補→イベント名へ', command=self.on_use_preview)\
        .pack(side=tk.LEFT)

        # プレビュー表示（イベント名の“次の行”）
        self.var_title_preview = tk.StringVar(value="")
        f_prev = ttk.Frame(frm); f_prev.pack(fill=tk.X, **pad)
        ttk.Label(f_prev, text='プレビュー', width=14).pack(side=tk.LEFT)
        ttk.Label(f_prev, textvariable=self.var_title_preview, anchor='w', foreground='#555')\
        .pack(side=tk.LEFT, fill=tk.X, expand=True)

        # --- 大分類（JSONマスタ由来の Combobox） ---
        f_cat = ttk.Frame(frm); f_cat.pack(fill=tk.X, **pad)
        ttk.Label(f_cat, text='大分類', width=14).pack(side=tk.LEFT)

        choices = [f"{k}：{self.category_map.get(k, k)}"
                  for k in sorted(self.category_map.keys(), key=lambda x: int(x) if str(x).isdigit() else str(x))]
        self.cmb_category = ttk.Combobox(f_cat, values=choices, state='readonly', width=20)

        init_code = self.var_category.get()
        self.cmb_category.set(f"{init_code}：{self.category_map.get(init_code, init_code)}")
        self.cmb_category.pack(side=tk.LEFT)

        def _on_cat_changed(event=None):
            sel = self.cmb_category.get()
            code = sel.split('：', 1)[0].strip() if '：' in sel else sel.strip()
            self.var_category.set(code)

        self.cmb_category.bind("<<ComboboxSelected>>", _on_cat_changed)

        f_date = row('対象日(,区切り)', self.var_dates)
        ttk.Button(f_date, text='サンプル挿入', command=lambda: self.var_dates.set('2024-12-24,2024-12-25,2025-01-02,2025-01-03')).pack(side=tk.LEFT, padx=6)
        ttk.Button(f_date, text='カレンダー…', command=self._open_date_picker).pack(side=tk.LEFT, padx=6)

        row('出力ファイル', self.var_out, browse=lambda: self._browse_save_xlsx(self.var_out))

        # オプション群
        opt = ttk.LabelFrame(frm, text='オプション'); opt.pack(fill=tk.X, **pad)
        f1 = ttk.Frame(opt); f1.pack(fill=tk.X, **pad)
        ttk.Checkbutton(f1, text='店舗別 split を出力 (--split-by-store)', variable=self.var_split).pack(side=tk.LEFT)
        ttk.Label(f1, text='split出力先').pack(side=tk.LEFT, padx=(12,6))
        ttk.Entry(f1, textvariable=self.var_split_dir, width=56).pack(side=tk.LEFT)
        ttk.Button(f1, text='参照…', command=lambda: self._browse_dir(self.var_split_dir)).pack(side=tk.LEFT, padx=6)

        f2 = ttk.Frame(opt); f2.pack(fill=tk.X, **pad)
        ttk.Label(f2, text='タイトルテンプレ').pack(side=tk.LEFT)
        ttk.Entry(f2, textvariable=self.var_title_template, width=50).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(f2, text='タイトルに日付を含めない (--no-date-in-title)', variable=self.var_no_date_in_title).pack(side=tk.LEFT, padx=12)

        # 完了後の挙動
        done = ttk.LabelFrame(frm, text='完了後の動作'); done.pack(fill=tk.X, **pad)
        ttk.Checkbutton(done, text='splitフォルダを開く', variable=self.var_open_after_split).pack(side=tk.LEFT)
        ttk.Checkbutton(done, text='本体xlsxを開く', variable=self.var_open_after_main).pack(side=tk.LEFT, padx=12)

        # 実行ボタン
        btn = ttk.Frame(frm); btn.pack(fill=tk.X, **pad)
        ttk.Button(btn, text='事前チェック', command=self._precheck_dialog).pack(side=tk.LEFT)
        ttk.Button(btn, text='実行', command=self._on_run).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn, text='停止', command=self._on_stop).pack(side=tk.LEFT)

        # ログ
        lf = ttk.LabelFrame(frm, text='ログ'); lf.pack(fill=tk.BOTH, expand=True, **pad)
        self.txt = tk.Text(lf, wrap=tk.NONE, height=18); self.txt.pack(fill=tk.BOTH, expand=True)

    # ===== ヘルパ群 =====
    def _load_category_map(self) -> dict[str, str]:
        """config/category_map.json を優先。無ければビルトイン。"""
        cfg = (REPO_ROOT / "config" / "category_map.json")
        if cfg.exists():
            try:
                data = json.loads(cfg.read_text(encoding="utf-8"))
                return {str(k): str(v) for k, v in data.items()}
            except Exception:
                pass
        return {"1":"寿司","2":"弁当","3":"温総菜","4":"冷総菜","5":"軽食","6":"魚惣菜"}
    
    def _cat_name_from_code(self, code: str | int) -> str:
        return self.category_map.get(str(code), str(code))

    def _dates_to_range(self, dates_list: list[str]) -> str:
        """['2024-12-24','2025-01-03'] -> '2024-12–2025-01'（同月なら 'YYYY-MM'）"""
        try:
            ds = sorted(datetime.strptime(str(d).strip(), "%Y-%m-%d") for d in dates_list if str(d).strip())
        except Exception:
            return ",".join(map(str, dates_list))
        if not ds:
            return ""
        a, b = ds[0], ds[-1]
        if a.year == b.year and a.month == b.month:
            return f"{a.year}-{a.month:02d}"
        return f"{a.year}-{a.month:02d}–{b.year}-{b.month:02d}"

    def _read_dates_list(self) -> list[str]:
        """GUI入力の 'YYYY-MM-DD,YYYY-MM-DD,...' を list[str] 化（/ 混在も '-' に統一）"""
        raw = (self.var_dates.get() or "").strip()
        if not raw:
            return []
        parts = [p.strip().replace("/", "-") for p in raw.split(",") if p.strip()]
        out: list[str] = []
        for p in parts:
            try:
                dt = datetime.strptime(p, "%Y-%m-%d")
                out.append(dt.strftime("%Y-%m-%d"))
            except Exception:
                # 無効な要素はスキップ
                pass
        return out

    def _build_title_preview(self, event_name: str, category_code: str | int,
                            dates_list: list[str], page_no: int,
                            tmpl: str | None, no_date_in_title: bool) -> str:
        """スクリプトと同等のタイトル構築。イベント名が空ならテンプレを適用。年は末尾日基準。"""
        cat = self._cat_name_from_code(category_code)
        ev = (event_name or "").strip()

        dlist = [] if no_date_in_title else [str(d) for d in dates_list]

        if dlist:
            ds_sorted = sorted(dlist)
            last_date = ds_sorted[-1]
            last_date_short = last_date[5:].replace("-", "/")
            yy, yyyy = last_date[2:4], last_date[:4]
        else:
            last_date = last_date_short = yy = yyyy = ""

        if ev:
            return f"{ev} {cat}単品データ（{page_no}）"

        default_tmpl = "{yy}年 {range} {cat}単品データ（{page}）"
        tmpl = (tmpl or default_tmpl).strip()

        values = {
            "event": ev,
            "cat": cat,
            "category": str(category_code),
            "dates": ",".join(dlist),
            "dates_short": ",".join(s[5:].replace("-", "/") for s in dlist),
            "range": self._dates_to_range(dlist),
            "year": yyyy,
            "yy":   yy,
            "date": last_date,
            "date_short": last_date_short,
            "page": str(page_no),
        }

        # 1) $var → 2) {var}
        s = Template(tmpl).safe_substitute(values)
        try:
            return s.format(**values)
        except Exception:
            return s

    # ===== コールバック群 =====
    def on_preview_title(self) -> None:
        event_name = self.var_event.get()
        category   = self.var_category.get()
        dates_list = self._read_dates_list()
        tmpl       = self.var_title_template.get()
        no_date    = bool(self.var_no_date_in_title.get())
        preview = self._build_title_preview(event_name, category, dates_list, page_no=1,
                                            tmpl=tmpl, no_date_in_title=no_date)
        self.var_title_preview.set(preview)

    def on_use_preview(self) -> None:
        s = (self.var_title_preview.get() or "").strip()
        if s:
            self.var_event.set(s)

    def _open_date_picker(self):
        # tkcalendar 無い時のフォールバックはそのまま
        if Calendar is None:
            messagebox.showinfo(
                "カレンダー未導入",
                "tkcalendar が見つかりませんでした。\n\npip install tkcalendar\n\n"
                "当面はテキスト入力（YYYY-MM-DD をカンマ区切り）で指定してください。"
            )
            return

        theme = self.theme         # ← いまのテーマを使う

        # ダイアログ
        top = tk.Toplevel(self.root)
        top.title("対象日を追加")
        top.grab_set()
        style_toplevel(top, theme)  # ← 地の色をテーマに合わせる

        # ダイアログ内は“カード面”で統一するためコンテナFrameを1枚かます
        container = ttk.Frame(top, padding=8)
        container.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(0, weight=1)
        top.rowconfigure(0, weight=1)

        pad = {"padx": 8, "pady": 6}

        # 既存値を正規化して読み込み
        def _parse_dates(s: str) -> list[str]:
            out = []
            for p in (s or "").split(","):
                p = p.strip().replace("/", "-")
                if not p:
                    continue
                try:
                    d = datetime.strptime(p, "%Y-%m-%d").date()
                    out.append(d.strftime("%Y-%m-%d"))
                except Exception:
                    pass
            return sorted(set(out))

        current = _parse_dates(self.var_dates.get())

        # 左: カレンダー（テーマ連動）
        cal = make_calendar(container, theme)
        if cal is None:
            messagebox.showinfo("カレンダー未導入", "tkcalendar が見つかりませんでした。")
            return
        cal.configure(date_pattern="yyyy-mm-dd")  # 表示/取得フォーマット
        cal.grid(row=0, column=0, rowspan=4, sticky="nsew", **pad)

        # 右: 選択済みリスト（tk.Listbox をテーマ色で）
        ttk.Label(container, text="選択済み").grid(row=0, column=1, sticky="w", **pad)
        lb = tk.Listbox(
            container, height=10, exportselection=False,
            bg=theme.surface, fg=theme.text,
            selectbackground=theme.primary, selectforeground=theme.bg,
            highlightthickness=0, bd=0, relief="flat"
        )
        lb.grid(row=1, column=1, sticky="nsew", **pad)
        for d in current:
            lb.insert(tk.END, d)

        def add_date():
            d = cal.get_date()
            if d and d not in current:
                current.append(d); current.sort()
                lb.delete(0, tk.END)
                for x in current:
                    lb.insert(tk.END, x)

        def remove_selected():
            sel = list(lb.curselection())
            if not sel:
                return
            for idx in reversed(sel):
                val = lb.get(idx)
                if val in current:
                    current.remove(val)
                lb.delete(idx)

        # 右: 追加/削除ボタン（ttk → 自動でテーマ適用）
        btns = ttk.Frame(container); btns.grid(row=2, column=1, sticky="w", **pad)
        ttk.Button(btns, text="追加", command=add_date).pack(side=tk.LEFT)
        ttk.Button(btns, text="削除", command=remove_selected).pack(side=tk.LEFT, padx=6)

        # OK/キャンセル
        def on_ok():
            self.var_dates.set(",".join(current))
            top.destroy()

        def on_cancel():
            top.destroy()

        cmd = ttk.Frame(container); cmd.grid(row=3, column=1, sticky="e", **pad)
        ttk.Button(cmd, text="OK", command=on_ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(cmd, text="キャンセル", command=on_cancel).pack(side=tk.LEFT)

        # 伸縮レイアウト
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

    # ===== ファイル選択（Browse系） =====
    def _browse_save_xlsx(self, var):
        path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel','*.xlsx'),('All','*.*')], initialdir=str((REPO_ROOT/'data'/'output').resolve()))
        if path: var.set(path)

    def _browse_dir(self, var):
        path = filedialog.askdirectory(initialdir=str((REPO_ROOT/'data'/'output').resolve()))
        if path: var.set(path)

    # ===== チェック・事前処理 =====
    def _parse_dates(self) -> list[str]:
        raw = (self.var_dates.get() or '').strip()
        raw = raw.replace('，', ',').replace(' ', '').replace('/', '-')
        if not raw:
            return []
        return [d for d in raw.split(',') if d]

    def _collect_needed_csv(self) -> list[Path]:
        csvs: set[Path] = set()
        pat = re.compile(r'^(\d{4})-(\d{2})-(\d{2})$')
        for d in self._parse_dates():
            m = pat.match(d)
            if not m: continue
            yyyy, mm = m.group(1), m.group(2)
            csvs.add(MATERIAL_DIR / yyyy / f'IT_{yyyy}{mm}.csv')
        return sorted(csvs)

    def _precheck(self) -> tuple[bool, str]:
        # category
        try:
            int(self.var_category.get())
        except Exception:
            return False, '大分類は整数コードで選択してください'
        # dates
        dates = self._parse_dates()
        if not dates:
            return False, '対象日をカンマ区切りで入力してください (YYYY-MM-DD)'
        for d in dates:
            try:
                datetime.strptime(d, '%Y-%m-%d')
            except Exception:
                return False, f'日付形式が不正です: {d}'
        # out
        if not (self.var_out.get() or '').strip():
            return False, '出力ファイル(.xlsx)を指定してください'
        # materials
        missing = [str(p) for p in self._collect_needed_csv() if not p.exists()]
        if missing:
            return False, '必要なCSVが見つかりません:\n' + '\n'.join(missing)
        # store master（あればチェック）
        if not STORE_MASTER.exists():
            # 厳密必須としないが警告に含める
            return True, f'警告: 店舗マスタが見つかりません: {STORE_MASTER}'
        return True, 'OK'

    def _precheck_dialog(self):
        ok, msg = self._precheck()
        if ok:
            messagebox.showinfo('事前チェック', f'事前チェックOK\n{msg}')
        else:
            messagebox.showerror('事前チェック', msg)

    
    # ===== 実行・プロセス制御 =====
    def _on_run(self):
        ok, msg = self._precheck()
        if not ok:
            messagebox.showerror('事前チェック', msg)
            return
        args = [
            sys.executable, '-m', CLI_SIMPLE,
            '--event-name', self.var_event.get() or '',
            '--category', self.var_category.get(),
            '--dates', ','.join(self._parse_dates()),
            '--out', self.var_out.get(),
        ]
        if self.var_title_template.get().strip():
            args += ['--title-template', self.var_title_template.get().strip()]
        if self.var_no_date_in_title.get():
            args += ['--no-date-in-title']
        if self.var_split.get():
            # split-dir は既存でも未作成でもOK（作成は CLI/GUI 側で実施）
            args += ['--split-by-store', '--split-dir', self.var_split_dir.get()]
        self._append_log(f"[gui] 実行コマンド:\n  {' '.join(args)}\n\n")
        t = threading.Thread(target=self._run_proc, args=(args,), daemon=True)
        t.start()

    def _run_proc(self, args):
        env = os.environ.copy(); env['PYTHONIOENCODING']='utf-8'; env['PYTHONUTF8']='1'
        self.proc = subprocess.Popen(
            args, cwd=str(REPO_ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding='utf-8', errors='replace', env=env)
        assert self.proc and self.proc.stdout is not None
        for line in self.proc.stdout:
            self.log_queue.put(line)
        code = self.proc.wait()
        self.log_queue.put(f"[done] code={code}\n")
        self._postcheck_and_notify(code)

    # ===== 後処理・通知 =====
    def _postcheck_and_notify(self, code: int):
        try:
            out_file = Path(self.var_out.get())
            split_dir = Path(self.var_split_dir.get())
            exists_main = out_file.exists()
            split_count = 0
            if self.var_split.get() and split_dir.exists():
                split_count = len(list(split_dir.rglob('*.xlsx')))
            self._append_log(f"[postcheck] main_exists={exists_main} split_count={split_count}\n")
            if code == 0:
                msg = f"処理完了\n本体: {'あり' if exists_main else 'なし'}\nsplit: {split_count} 件"
                messagebox.showinfo('完了', msg)
                if self.var_open_after_split.get() and self.var_split.get() and split_dir.exists():
                    try: os.startfile(split_dir)  # type: ignore[attr-defined]
                    except Exception: pass
                if self.var_open_after_main.get() and exists_main:
                    try: os.startfile(out_file)  # type: ignore[attr-defined]
                    except Exception: pass
            else:
                messagebox.showerror('失敗', '処理がエラー終了しました')
        except Exception as e:
            self._append_log(f"[postcheck-error] {e}\n")

    def _on_stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate(); self._append_log('[gui] 停止\n')

    # ===== ログ更新・ポーリング =====
    def _append_log(self, s: str):
        self.txt.insert(tk.END, s)
        self.txt.see(tk.END)

    def _poll_log_queue(self):
        try:
            while True:
                self._append_log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(80, self._poll_log_queue)

def main():
    root = tk.Tk()
    root.title("GUI Launcher")

    # 高DPIスケーリング（先にやる）
    try:
        root.call('tk', 'scaling', 1.2)
    except Exception:
        pass

    # 現在テーマの単一ソース
    current = {"t": theme_cyber.theme}           # ← デフォルトはサイバー
    apply_theme(root, current["t"])              # ← 一括適用

    # アプリ本体を生成し、テーマを共有（self.theme を参照できるように）
    app = TopNGuiApp(root)
    app.theme = current["t"]                     # ← 重要：ダイアログ等から参照

    # Ctrl+T で パステル ↔ サイバー
    def _toggle(_=None):
        current["t"] = (theme_pastel.theme
                        if current["t"].name == "cyber"
                        else theme_cyber.theme)
        app.theme = current["t"]                 # ← アプリ側の現在テーマも更新
        apply_theme(root, current["t"])          # ← 再適用（即時着せ替え）
    root.bind("<Control-t>", _toggle)

    # ※ 旧来の 'vista' / 'xpnative' 切替は撤去（apply_theme を上書きするため）
    # try:
    #     style = ttk.Style(root)
    #     if 'vista' in style.theme_names():
    #         style.theme_use('vista')
    #     elif 'xpnative' in style.theme_names():
    #         style.theme_use('xpnative')
    # except Exception:
    #     pass

    root.mainloop()

if __name__ == '__main__':
    main()

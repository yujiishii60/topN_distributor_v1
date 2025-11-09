# styles/apply_ttk_min.py
from tkinter import ttk, font

def apply_theme(root, theme):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    # 親ウィンドウの地の色
    try:
        root.configure(background=theme.bg)
    except Exception:
        pass

    # ベースフォント
    font.nametofont("TkDefaultFont").configure(
        family=theme.font_family, size=theme.font_size_base
    )

    # 共通デフォルト
    style.configure(".", background=theme.bg, foreground=theme.text)

    # --- Frame / Label / Button -------------------------
    style.configure("TFrame", background=theme.surface, borderwidth=0, relief="flat")
    style.configure("TLabel", background=theme.surface, foreground=theme.text)

    style.configure(
        "TButton",
        background=theme.surface_alt, foreground=theme.text,
        borderwidth=0, padding=(theme.padding, theme.padding),
        focusthickness=1, focuscolor=theme.accent
    )
    style.map(
        "TButton",
        background=[("active", theme.primary), ("pressed", theme.primary), ("disabled", theme.surface_alt)],
        foreground=[("active", theme.bg), ("pressed", theme.bg), ("disabled", theme.text_muted)]
    )

    # --- Entry / Combobox -------------------------------
    # Entry の fieldbackground が地の色と近いと文字が見えにくいので明示
    style.configure("TEntry",
        fieldbackground=theme.surface, foreground=theme.text,
        insertcolor=theme.text
    )
    # Combobox 本体
    style.configure("TCombobox",
        fieldbackground=theme.surface, foreground=theme.text
    )
    # ドロップダウン（リスト部分）の色（ttkの制限があるため map のみ）
    style.map("TCombobox",
        fieldbackground=[("readonly", theme.surface)],
        foreground=[("readonly", theme.text)]
    )

    # --- Check/Radiobutton ------------------------------
    style.configure("TCheckbutton",
        background=theme.surface, foreground=theme.text
    )
    style.map("TCheckbutton",
        foreground=[("selected", theme.primary)],
    )
    style.configure("TRadiobutton",
        background=theme.surface, foreground=theme.text
    )
    style.map("TRadiobutton",
        foreground=[("selected", theme.primary)],
    )

    # --- Notebook（使っていれば） -----------------------
    style.configure("TNotebook", background=theme.surface, borderwidth=0)
    style.configure("TNotebook.Tab",
        background=theme.surface_alt, foreground=theme.text, padding=(10, 6)
    )
    style.map("TNotebook.Tab",
        background=[("selected", theme.primary), ("active", theme.surface_alt)],
        foreground=[("selected", theme.bg), ("active", theme.text)]
    )

    # --- Treeview ---------------------------------------
    style.configure("Treeview",
        background=theme.surface,
        fieldbackground=theme.surface,
        foreground=theme.text, rowheight=24, borderwidth=0
    )
    style.configure("Treeview.Heading",
        background=theme.surface_alt, foreground=theme.text
    )
    style.map("Treeview.Heading",
        background=[("active", theme.primary)],
        foreground=[("active", theme.bg)]
    )

    # 交互の縞（オプション）
    style.map("Treeview",
        background=[("selected", theme.primary)],
        foreground=[("selected", theme.bg)]
    )
    try:
        # striped 行を有効化（clam では option で行う）
        root.tk.call("ttk::style", "configure", "Treeview", "-striped", "1")
    except Exception:
        pass

    # --- Labelframe（“オプション”枠など） ---------------
    style.configure("TLabelframe", background=theme.surface, borderwidth=1, relief="flat")
    style.configure("TLabelframe.Label",
        background=theme.surface, foreground=theme.text_muted
    )

    return style

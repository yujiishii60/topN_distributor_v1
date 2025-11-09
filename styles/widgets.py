# styles/widgets.py
try:
    from tkcalendar import Calendar
except Exception:
    Calendar = None

def style_toplevel(top, theme):
    """Toplevel ダイアログの地の色をテーマに合わせる（任意で呼ぶ）"""
    try:
        top.configure(background=theme.bg)
    except Exception:
        pass

def make_calendar(parent, theme):
    """
    テーマと親ウィジェットを受け取って tkcalendar.Calendar を返す。
    tkcalendar が無ければ None を返す（呼び出し側で分岐）。
    """
    if Calendar is None:
        return None

    return Calendar(
        parent,
        selectmode="day",
        # ベース
        background=theme.surface,
        foreground=theme.text,
        bordercolor=theme.surface_alt,
        # ヘッダ
        headersbackground=theme.surface_alt,
        headersforeground=theme.text,
        # 通常日
        normalbackground=theme.surface,
        normalforeground=theme.text,
        # 週末
        weekendbackground=theme.surface,
        weekendforeground=theme.text,
        # 前後月
        othermonthbackground=theme.surface_alt,
        othermonthforeground=theme.text_muted,
        # 選択色
        selectbackground=theme.primary,
        selectforeground=theme.bg,
        disabledbackground=theme.surface_alt,
    )

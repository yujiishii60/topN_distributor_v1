
# How to apply the Cyber theme (minimal change)

In your `gui_launcher.py` (or your main entry), add:

```python
from styles import theme_cyber
from styles.apply_ttk_min import apply_theme

# after creating root = tk.Tk()
apply_theme(root, theme_cyber.theme)

# (Optional) toggle with Ctrl+T
from styles import theme_pastel
current = {"t": theme_cyber.theme}
def _toggle(_=None):
    current["t"] = theme_pastel.theme if current["t"].name == "cyber" else theme_cyber.theme
    apply_theme(root, current["t"])
root.bind("<Control-t>", _toggle)
```

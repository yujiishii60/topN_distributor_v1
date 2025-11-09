
from dataclasses import dataclass

@dataclass
class Theme:
    name: str
    bg: str
    surface: str
    surface_alt: str
    text: str
    text_muted: str
    primary: str
    accent: str
    radius: int = 10
    padding: int = 8
    font_family: str = "Meiryo UI"
    font_size_base: int = 11

#!/usr/bin/env python3
"""Собирает самодостаточный HTML офера: внедряет шрифты (base64) и логотип."""
import base64, pathlib, sys

HERE = pathlib.Path(__file__).parent
template = (HERE / "offer.template.html").read_text(encoding="utf-8")

# --- шрифты ---
fonts_css = (HERE / "assets" / "fonts_embedded.css").read_text(encoding="utf-8")

# --- логотип ---
logo_b64 = base64.b64encode((HERE / "assets" / "logo.png").read_bytes()).decode()

html = template.replace("/*__FONTS__*/", fonts_css).replace("__LOGO__", logo_b64)

out = HERE / "Коммерческое-предложение-Факторинг.html"
out.write_text(html, encoding="utf-8")
print(f"OK -> {out}  ({out.stat().st_size//1024} KB)")

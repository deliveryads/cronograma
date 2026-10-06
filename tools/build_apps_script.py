"""Gera apps-script/Index.html a partir de eventos/index.html (logo embutido).

Uso: python3 tools/build_apps_script.py
"""
import base64
import pathlib

root = pathlib.Path(__file__).resolve().parent.parent
src = (root / "eventos" / "index.html").read_text(encoding="utf-8")
logo = base64.b64encode((root / "eventos" / "logo.png").read_bytes()).decode()
out = src.replace('src="logo.png"', 'src="data:image/png;base64,' + logo + '"')
assert 'src="logo.png"' not in out
(root / "apps-script" / "Index.html").write_text(out, encoding="utf-8")
print("apps-script/Index.html gerado (%d KB)" % (len(out) // 1024))

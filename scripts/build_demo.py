# 本番版 src/arm-kensa.html からデモ版 dist/arm-kensa-demo.html を作る
import json, base64, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
src = (root/"src/arm-kensa.html").read_text(encoding="utf-8")
rows = []
for line in (root/"data/shipping_oct_44.txt").read_text(encoding="utf-8").splitlines():
    p = line.split("|")
    if len(p) >= 5: rows.append({"z": p[0], "g": p[1], "d": p[2], "to": p[3], "sp": p[4]})
img = "data:image/jpeg;base64," + base64.b64encode((root/"data/demo_kokuin.jpg").read_bytes()).decode()
for a, b in [("const DEMO = false;", "const DEMO = true;"),
             ("const DEMO_ROWS = null;", "const DEMO_ROWS = " + json.dumps(rows, ensure_ascii=False) + ";"),
             ("const DEMO_IMG = null;", "const DEMO_IMG = " + json.dumps(img) + ";")]:
    assert a in src, a
    src = src.replace(a, b)
(root/"dist").mkdir(exist_ok=True)
(root/"dist/arm-kensa-demo.html").write_text(src, encoding="utf-8")
print("built", len(rows), "rows")

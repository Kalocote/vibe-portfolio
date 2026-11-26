from pathlib import Path
from html import escape

log = Path.home() / "vibe" / "organize.log"
text = log.read_text(encoding="utf-8")

moves = []
for line in text.splitlines():
    if line.startswith("已移动"):
        name, kind = line.split(" -> ", 1)
        name = name.replace("已移动:", "").strip()
        kind = kind.strip()
        moves.append((name, kind))

# 脱敏：真实文件名换成编号，只保留扩展名和分类
safe_moves = []
for i, (name, kind) in enumerate(moves, start=1):
    suffix = Path(name).suffix or ""
    safe_moves.append((f"文件_{i:02d}{suffix}", kind))

counts = {}
for name, kind in safe_moves:
    counts[kind] = counts.get(kind, 0) + 1

rows = ""
for name, kind in safe_moves:
    rows += f"<tr><td>{escape(name)}</td><td>{escape(kind)}</td></tr>"

summary = ""
for kind, n in counts.items():
    summary += f'<span class="tag">{escape(kind)} · {n}</span> '

html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>整理报告</title>
<link rel="stylesheet" href="style.css">
</head>
<body>

<nav class="topbar">
    <div class="topbar-inner">
        <a href="home.html">我的小作品</a>
        <button id="themeBtn">深色</button>
    </div>
</nav>

<div class="wrap-narrow page-body">
    <div class="page-head">
        <h1>整理报告</h1>
        <p>下载文件夹的自动整理记录</p>
    </div>
    <div class="card">
        <p class="muted" style="margin: 0;">共整理 {len(safe_moves)} 个文件</p>
        <div class="stats">{summary}</div>
        <table class="data-table">
        <tr><th>文件名</th><th>归档到</th></tr>
        {rows}
        </table>
    </div>
</div>

<script src="theme.js"></script>
</body>
</html>
"""

Path.home().joinpath("vibe", "report.html").write_text(html, encoding="utf-8")
print("已生成脱敏版 report.html")

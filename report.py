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

css = """
body {
    margin: 0;
    padding: 40px;
    background: #f5f3fa;
    font-family: -apple-system, "PingFang SC", sans-serif;
    color: #232129;
}
.card {
    max-width: 900px;
    margin: 0 auto;
    background: white;
    border-radius: 16px;
    padding: 32px;
    box-shadow: 0 12px 40px rgba(80, 60, 120, 0.1);
}
h1 { margin-top: 0; color: #7b4bc4; }
.home { display: inline-block; margin-bottom: 16px; color: #7b4bc4; font-size: 14px; text-decoration: none; }
.home:hover { text-decoration: underline; }
table { width: 100%; border-collapse: collapse; }
th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; }
th { color: #6b6875; font-size: 13px; }
.tag {
    display: inline-block;
    margin: 16px 6px 0 0;
    padding: 6px 14px;
    border-radius: 999px;
    background: #efe8fa;
    color: #6a3fa8;
    font-size: 13px;
}
"""

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
<title>整理报告</title>
<style>{css}</style>
</head>
<body>
<div class="card">
<a class="home" href="home.html">← 返回首页</a>
<h1>我的下载文件夹整理记录</h1>
<p>共整理 {len(safe_moves)} 个文件</p>
{summary}
<table>
<tr><th>文件名</th><th>归档到</th></tr>
{rows}
</table>
</div>
</body>
</html>
"""

Path.home().joinpath("vibe", "report.html").write_text(html, encoding="utf-8")
print("已生成脱敏版 report.html")

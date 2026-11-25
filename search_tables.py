from pathlib import Path
import csv
import sys
from html import escape

from openpyxl import load_workbook

CSS = """
body {
    margin: 0;
    padding: 40px 20px;
    background: #f5f3fa;
    font-family: -apple-system, "PingFang SC", sans-serif;
    color: #232129;
}
.wrap { max-width: 1000px; margin: 0 auto; }
h1 { color: #7b4bc4; margin-top: 0; }
.info { color: #6b6875; font-size: 14px; margin-bottom: 24px; }
.hit {
    background: white;
    border-radius: 16px;
    padding: 20px 24px;
    margin-bottom: 16px;
    box-shadow: 0 8px 30px rgba(80, 60, 120, 0.08);
}
.where { color: #8b8798; font-size: 13px; margin: 0 0 12px; }
.where strong { color: #6a3fa8; font-size: 15px; }
mark { background: #ffe9a8; color: #5a4200; padding: 1px 3px; border-radius: 3px; }
table { width: 100%; border-collapse: collapse; }
th, td { padding: 8px 12px; text-align: left; border-bottom: 1px solid #eee; font-size: 14px; }
th { color: #6b6875; width: 30%; }
.empty {
    background: white;
    border-radius: 16px;
    padding: 40px;
    text-align: center;
    color: #8b8798;
}
"""


def make_hit(file, sheet, row_no, header, values, keyword):
    """把一条匹配记录整理成统一格式"""
    pairs = []
    for key, value in zip(header, values):
        if value is None or str(value) == "":
            continue
        pairs.append((str(key), str(value)))
    return {
        "file": file,
        "sheet": sheet,
        "row": row_no,
        "pairs": pairs,
    }


def search_xlsx(path, keyword):
    """在一个 xlsx 里搜索关键词"""
    hits = []
    wb = load_workbook(path, read_only=True, data_only=True)
    kw = keyword.lower()

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = rows[0]

        for i, values in enumerate(rows[1:], start=2):
            if any(kw in str(v).lower() for v in values if v is not None):
                hits.append(make_hit(path.name, ws.title, i, header, values, keyword))

    wb.close()
    return hits


def search_csv(path, keyword):
    """在一个 csv 里搜索关键词"""
    hits = []
    kw = keyword.lower()

    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    if not rows:
        return hits
    header = rows[0]

    for i, values in enumerate(rows[1:], start=2):
        if any(kw in v.lower() for v in values):
            hits.append(make_hit(path.name, "单表", i, header, values, keyword))

    return hits


def scan(folder, keyword):
    """扫描文件夹里所有表格，返回匹配结果"""
    hits = []

    for path in sorted(folder.iterdir()):
        if path.name.startswith("."):
            continue
        suffix = path.suffix.lower()
        if suffix == ".xlsx":
            hits.extend(search_xlsx(path, keyword))
        elif suffix == ".csv":
            hits.extend(search_csv(path, keyword))

    return hits


def highlight(text, keyword):
    """给匹配到的关键词加高亮（HTML 里用 <mark>）"""
    i = text.lower().find(keyword.lower())
    if i == -1:
        return escape(text)
    before = escape(text[:i])
    matched = escape(text[i:i + len(keyword)])
    after = escape(text[i + len(keyword):])
    return f"{before}<mark>{matched}</mark>{after}"


def build_html(hits, keyword, folder):
    """数据 + 模板 → 搜索报告"""
    if hits:
        cards = ""
        for hit in hits:
            rows = ""
            for key, value in hit["pairs"]:
                rows += (
                    "<tr>"
                    f"<th>{escape(key)}</th>"
                    f"<td>{highlight(value, keyword)}</td>"
                    "</tr>"
                )
            cards += (
                "<div class='hit'>"
                "<p class='where'>📄 <strong>"
                f"{escape(hit['file'])}</strong> · 工作表：{escape(hit['sheet'])} · 第 {hit['row']} 行"
                "</p>"
                "<table>" + rows + "</table>"
                "</div>"
            )
        body = cards
    else:
        body = f"<div class='empty'>没有找到包含 “{escape(keyword)}” 的内容</div>"

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>表格搜索结果</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
<h1>表格搜索结果</h1>
<p class="info">关键词：{escape(keyword)} · 找到 {len(hits)} 条 · 文件夹：{escape(str(folder))}</p>
{body}
</div>
</body>
</html>
"""
    return html


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: .venv/bin/python search_tables.py 关键词 [文件夹]")
        sys.exit(1)

    keyword = sys.argv[1]
    folder = Path.home() / "vibe" / "测试下载" / "表格样例"

    if len(sys.argv) > 2:
        folder = Path(sys.argv[2]).expanduser()

    if not folder.exists():
        print("文件夹不存在:", folder)
        sys.exit(1)

    hits = scan(folder, keyword)
    report = folder / "搜索结果.html"
    report.write_text(build_html(hits, keyword, folder), encoding="utf-8")

    print(f"关键词: {keyword} · 找到 {len(hits)} 条")
    for hit in hits:
        print(f"  📄 {hit['file']} / {hit['sheet']} / 第 {hit['row']} 行")
    print("已生成报告:", report)

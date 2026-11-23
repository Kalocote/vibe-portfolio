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
.wrap { max-width: 900px; margin: 0 auto; }
h1 { color: #7b4bc4; margin-top: 0; }
.stats span {
    display: inline-block;
    margin: 0 8px 16px 0;
    padding: 8px 16px;
    background: #efe8fa;
    color: #6a3fa8;
    border-radius: 999px;
    font-size: 14px;
}
.file {
    background: white;
    border-radius: 16px;
    padding: 24px;
    margin-bottom: 20px;
    box-shadow: 0 8px 30px rgba(80, 60, 120, 0.08);
}
.file h2 { margin: 0 0 16px; font-size: 18px; color: #4a3858; }
table { width: 100%; border-collapse: collapse; }
th, td { padding: 10px 12px; text-align: left; border-bottom: 1px solid #eee; }
th { color: #6b6875; font-size: 13px; }
.header-cell { color: #7b4bc4; font-size: 13px; word-break: break-all; }
details {
    margin-top: 16px;
    padding-top: 12px;
    border-top: 1px dashed #e0dae8;
}
summary {
    cursor: pointer;
    color: #8b8798;
    font-size: 13px;
}
summary:hover { color: #7b4bc4; }
details table { margin-top: 12px; }
"""


def read_xlsx(path):
    """读一个 xlsx 文件，返回每个工作表的信息"""
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = []

    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        header = rows[0] if rows else ()
        sheets.append({
            "name": ws.title,
            "state": ws.sheet_state,
            "rows": max(len(rows) - 1, 0),
            "cols": len(header),
            "header": [str(x) if x is not None else "" for x in header],
        })

    wb.close()
    return sheets


def read_csv(path):
    """读一个 csv 文件，当成单表处理"""
    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    header = rows[0] if rows else []
    return [{
        "name": "单表",
        "state": "visible",
        "rows": max(len(rows) - 1, 0),
        "cols": len(header),
        "header": header,
    }]


def scan(folder):
    """扫描文件夹，返回所有表格文件的信息"""
    results = []

    for path in sorted(folder.iterdir()):
        if path.name.startswith("."):
            continue

        suffix = path.suffix.lower()
        if suffix == ".xlsx":
            sheets = read_xlsx(path)
        elif suffix == ".csv":
            sheets = read_csv(path)
        else:
            continue

        results.append({"name": path.name, "sheets": sheets})

    return results


TABLE_HEAD = "<tr><th>工作表</th><th>数据行</th><th>列数</th><th>表头</th></tr>"


def make_sheet_rows(sheets):
    """把工作表列表变成表格的行"""
    rows = ""
    for sheet in sheets:
        header = " | ".join(escape(h) for h in sheet["header"])
        rows += (
            "<tr>"
            f"<td>{escape(sheet['name'])}</td>"
            f"<td>{sheet['rows']}</td>"
            f"<td>{sheet['cols']}</td>"
            f"<td class='header-cell'>{header}</td>"
            "</tr>"
        )
    return rows


def build_html(data):
    """把数据 + 模板拼成完整的 HTML 报告"""
    total_files = len(data)
    total_sheets = sum(len(item["sheets"]) for item in data)
    total_rows = sum(sheet["rows"] for item in data for sheet in item["sheets"])

    cards = ""
    for item in data:
        visible = [s for s in item["sheets"] if s["state"] == "visible"]
        hidden = [s for s in item["sheets"] if s["state"] != "visible"]

        card = "<div class='file'>"
        card += f"<h2>📄 {escape(item['name'])}</h2>"

        if visible:
            card += "<table>" + TABLE_HEAD + make_sheet_rows(visible) + "</table>"

        if hidden:
            card += (
                "<details>"
                f"<summary>🙈 {len(hidden)} 个隐藏工作表（点击展开）</summary>"
                "<table>" + TABLE_HEAD + make_sheet_rows(hidden) + "</table>"
                "</details>"
            )

        card += "</div>"
        cards += card

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>表格统计报告</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
<h1>表格统计报告</h1>
<div class="stats">
<span>文件 {total_files} 个</span>
<span>工作表 {total_sheets} 个</span>
<span>数据行 {total_rows} 行</span>
</div>
{cards}
</div>
</body>
</html>
"""
    return html


if __name__ == "__main__":
    folder = Path.home() / "vibe" / "测试下载" / "表格样例"

    if len(sys.argv) > 1:
        folder = Path(sys.argv[1]).expanduser()

    if not folder.exists():
        print("文件夹不存在:", folder)
        sys.exit(1)

    data = scan(folder)
    report = folder / "表格统计报告.html"
    report.write_text(build_html(data), encoding="utf-8")

    print("扫描文件夹:", folder)
    print("找到", len(data), "个表格文件")
    print("已生成报告:", report)

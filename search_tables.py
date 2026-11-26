from pathlib import Path
import csv
import sys
from html import escape

from openpyxl import load_workbook

CSS = """
:root {
    --bg: #f5f5f7;
    --surface: #ffffff;
    --text: #1d1d1f;
    --text-2: #6e6e73;
    --text-3: #86868b;
    --blue: #0071e3;
    --line: rgba(0, 0, 0, 0.1);
    --ease: cubic-bezier(0.28, 0.11, 0.32, 1);
}
* { box-sizing: border-box; }
body {
    margin: 0;
    padding: 56px 20px 80px;
    background: var(--bg);
    font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang SC", sans-serif;
    color: var(--text);
    -webkit-font-smoothing: antialiased;
    line-height: 1.47;
}
.wrap { max-width: 980px; margin: 0 auto; }
h1 {
    margin: 0 0 8px;
    font-size: clamp(32px, 5vw, 48px);
    font-weight: 600;
    letter-spacing: -0.03em;
}
.info { color: var(--text-2); font-size: 15px; margin-bottom: 28px; }

/* 顶部两页切换 */
.jump { display: flex; gap: 10px; margin: 0 0 28px; flex-wrap: wrap; }
.jump a {
    padding: 10px 18px;
    background: var(--surface);
    color: var(--text-2);
    border-radius: 999px;
    font-size: 14px;
    text-decoration: none;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04), 0 6px 20px rgba(0,0,0,0.05);
    transition: background 0.2s var(--ease), transform 0.2s var(--ease);
}
.jump a:hover { color: var(--blue); transform: translateY(-1px); }
.jump a.current { background: var(--blue); color: #fff; }

/* 命中卡片 */
.hit {
    background: var(--surface);
    border-radius: 18px;
    padding: 22px 26px;
    margin-bottom: 14px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04), 0 8px 30px rgba(0,0,0,0.05);
}
.where { color: var(--text-3); font-size: 13px; margin: 0 0 14px; }
.where strong { color: var(--blue); font-size: 15px; font-weight: 600; }
mark { background: rgba(255, 149, 0, 0.28); color: var(--text); padding: 1px 4px; border-radius: 4px; }
table { width: 100%; border-collapse: collapse; }
th, td { padding: 11px 14px; text-align: left; border-bottom: 1px solid var(--line); font-size: 15px; }
th { color: var(--text-2); width: 30%; font-size: 13px; font-weight: 500; }
tr:last-child th, tr:last-child td { border-bottom: none; }

.empty {
    background: var(--surface);
    border-radius: 18px;
    padding: 56px 24px;
    text-align: center;
    color: var(--text-2);
    box-shadow: 0 1px 2px rgba(0,0,0,0.04), 0 8px 30px rgba(0,0,0,0.05);
}
.no-match { color: var(--text-2); font-size: 15px; padding: 20px 0; }

/* 分页 */
.pager {
    display: flex;
    justify-content: center;
    align-items: center;
    gap: 8px;
    margin-top: 28px;
    flex-wrap: wrap;
}
.pager button {
    min-width: 40px;
    height: 40px;
    padding: 0 12px;
    border: none;
    border-radius: 12px;
    background: rgba(0,0,0,0.05);
    color: var(--text);
    font-size: 14px;
    font-family: inherit;
    cursor: pointer;
    transition: background 0.2s var(--ease), transform 0.2s var(--ease);
}
.pager button:hover:not(:disabled):not(.active) { background: rgba(0,0,0,0.09); transform: translateY(-1px); }
.pager button:disabled { opacity: 0.35; cursor: not-allowed; }
.pager button.active { background: var(--blue); color: #fff; }
.pager .arrow { font-size: 20px; line-height: 1; padding: 0 14px; }
.page-numbers { display: flex; gap: 8px; flex-wrap: wrap; justify-content: center; }
.page-numbers .dots {
    min-width: 24px; height: 40px;
    display: flex; align-items: center; justify-content: center;
    color: var(--text-3); font-size: 14px;
}
.jump-page {
    display: flex; justify-content: center; align-items: center; gap: 8px;
    margin-top: 16px; color: var(--text-2); font-size: 14px;
}
.jump-page input {
    width: 54px; height: 36px; text-align: center;
    border: 1px solid var(--line); border-radius: 10px;
    font-size: 14px; font-family: inherit; color: var(--text); background: var(--surface);
}
.jump-page input:focus { outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px rgba(0,113,227,0.18); }
.jump-page button {
    height: 36px; padding: 0 16px; border: none; border-radius: 10px;
    background: rgba(0,0,0,0.05); color: var(--text);
    font-size: 13px; font-family: inherit; cursor: pointer;
}
.jump-page button:hover { background: rgba(0,0,0,0.09); }
"""
JS = """
const PAGE_SIZE = 10;

// 生成要显示的页码：页数少就全显示，多则只显示当前页附近 5 个 + 首尾
function pageList(current, pages) {
    if (pages <= 7) {
        const all = [];
        for (let n = 1; n <= pages; n++) all.push(n);
        return all;
    }

    let start = current - 2;
    let end = current + 2;
    if (start < 1) { end += 1 - start; start = 1; }
    if (end > pages) { start -= end - pages; end = pages; }

    const items = [];
    if (start > 1) {
        items.push(1);
        if (start > 2) items.push("…");
    }
    for (let n = start; n <= end; n++) items.push(n);
    if (end < pages) {
        if (end < pages - 1) items.push("…");
        items.push(pages);
    }
    return items;
}

function setupPager(group) {
    const cards = group.querySelectorAll(".hit");
    const pager = group.querySelector(".pager");
    const jump = group.querySelector(".jump-page");

    if (cards.length <= PAGE_SIZE) {
        pager.remove();
        jump.remove();
        return;
    }

    let page = 1;
    const pages = Math.ceil(cards.length / PAGE_SIZE);
    const numbers = pager.querySelector(".page-numbers");
    const prev = pager.querySelector(".prev");
    const next = pager.querySelector(".next");
    const input = jump.querySelector(".page-input");
    const go = jump.querySelector(".go");

    function goTo(target) {
        if (!Number.isInteger(target) || target < 1 || target > pages) return;
        page = target;
        show();
    }

    function show() {
        cards.forEach(function (card, i) {
            const start = (page - 1) * PAGE_SIZE;
            card.style.display = (i >= start && i < start + PAGE_SIZE) ? "" : "none";
        });

        numbers.innerHTML = "";
        for (const item of pageList(page, pages)) {
            if (item === "…") {
                const dots = document.createElement("span");
                dots.className = "dots";
                dots.textContent = "…";
                numbers.appendChild(dots);
                continue;
            }
            const btn = document.createElement("button");
            btn.textContent = item;
            if (item === page) btn.className = "active";
            btn.addEventListener("click", function () { goTo(item); });
            numbers.appendChild(btn);
        }

        prev.disabled = page === 1;
        next.disabled = page === pages;
        input.max = pages;
        input.value = page;
    }

    prev.addEventListener("click", function () { goTo(page - 1); });
    next.addEventListener("click", function () { goTo(page + 1); });
    go.addEventListener("click", function () { goTo(parseInt(input.value, 10)); });
    input.addEventListener("keydown", function (e) {
        if (e.key === "Enter") goTo(parseInt(input.value, 10));
    });

    show();
}

document.querySelectorAll(".result-group").forEach(setupPager);
"""


def make_hit(file, sheet, row_no, header, values, keyword, hidden):
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
        "hidden": hidden,
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

        hidden = ws.sheet_state != "visible"
        for i, values in enumerate(rows[1:], start=2):
            if any(kw in str(v).lower() for v in values if v is not None):
                hits.append(make_hit(path.name, ws.title, i, header, values, keyword, hidden))

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
            hits.append(make_hit(path.name, "单表", i, header, values, keyword, False))

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


def make_card(hit, keyword):
    """把一条命中渲染成一张卡片"""
    rows = ""
    for key, value in hit["pairs"]:
        rows += (
            "<tr>"
            f"<th>{escape(key)}</th>"
            f"<td>{highlight(value, keyword)}</td>"
            "</tr>"
        )
    return (
        "<div class='hit'>"
        "<p class='where'>📄 <strong>"
        f"{escape(hit['file'])}</strong> · 工作表：{escape(hit['sheet'])} · 第 {hit['row']} 行"
        "</p>"
        "<table>" + rows + "</table>"
        "</div>"
    )


def build_html(hits, keyword, folder, current):
    """生成单个结果页：current 是 "normal" 或 "hidden" """
    visible = [hit for hit in hits if not hit["hidden"]]
    hidden = [hit for hit in hits if hit["hidden"]]
    group = visible if current == "normal" else hidden
    title = "表格搜索结果 · 正常工作表" if current == "normal" else "表格搜索结果 · 隐藏工作表"

    pager_html = (
        "<div class='pager'>"
        "<button class='arrow prev'>‹</button>"
        "<span class='page-numbers'></span>"
        "<button class='arrow next'>›</button>"
        "</div>"
        "<div class='jump-page'>"
        f"共 {len(group)} 条 · 跳至 <input type='number' class='page-input' min='1' /> 页"
        "<button class='go'>跳转</button>"
        "</div>"
    )

    if group:
        cards = "".join(make_card(hit, keyword) for hit in group)
        body = "<div class='result-group'>" + cards + pager_html + "</div>"
    else:
        body = f"<div class='empty'>这一组没有找到包含 “{escape(keyword)}” 的内容</div>"

    normal_cls = "current" if current == "normal" else ""
    hidden_cls = "current" if current == "hidden" else ""
    jump = (
        "<nav class='jump'>"
        f"<a class='{normal_cls}' href='搜索结果.html'>📄 正常工作表 · {len(visible)} 条</a>"
        f"<a class='{hidden_cls}' href='隐藏结果.html'>🙈 隐藏工作表 · {len(hidden)} 条</a>"
        "</nav>"
    )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>{title}</title>
<style>{CSS}</style>
</head>
<body>
<div class="wrap">
<h1>{title}</h1>
<p class="info">关键词：{escape(keyword)} · {len(group)} 条 · 文件夹：{escape(str(folder))}</p>
{jump}
{body}
</div>
<script>{JS}</script>
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
    visible = [hit for hit in hits if not hit["hidden"]]
    hidden = [hit for hit in hits if hit["hidden"]]

    normal_report = folder / "搜索结果.html"
    hidden_report = folder / "隐藏结果.html"
    normal_report.write_text(build_html(hits, keyword, folder, "normal"), encoding="utf-8")
    hidden_report.write_text(build_html(hits, keyword, folder, "hidden"), encoding="utf-8")

    print(f"关键词: {keyword} · 找到 {len(hits)} 条（正常 {len(visible)} · 隐藏 {len(hidden)}）")
    for hit in visible:
        print(f"  📄 {hit['file']} / {hit['sheet']} / 第 {hit['row']} 行")
    if hidden:
        print("  🙈 隐藏工作表：")
        for hit in hidden:
            print(f"     {hit['file']} / {hit['sheet']} / 第 {hit['row']} 行")
    print("已生成报告:", normal_report)
    print("           ", hidden_report)

# vibe-portfolio

从零开始学 vibe coding 的记录。这里放着我这段时间做的网页和 Python 小工具。

在线预览：https://kalocote.github.io/vibe-portfolio/

## 网页

- **home.html** — 导航首页，苹果风格，深浅色切换
- **index.html** — 我的第一个网页，记录最初学到的东西
- **todo.html** — 待办清单：本地保存、优先级、筛选、导出 JSON
- **report.html** — 下载文件夹的整理报告（由 report.py 生成）

## Python 工具

- **organize.py** — 按扩展名整理下载文件夹，launchd 定时运行
- **summarize.py** — 批量统计 Excel/CSV 的行列数和表头
- **search_tables.py** — 在表格里搜关键词，隐藏工作表单独归类
- **report.py** — 读整理日志生成脱敏报告
- **guess.py / hello.py / trytest.py** — 早期练习

## 运行

Excel 相关工具依赖 openpyxl，用虚拟环境：

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python search_tables.py 关键词 [文件夹]
.venv/bin/python summarize.py [文件夹]
```

## 一些习惯

- 一个脚本一个职责，Excel 统计和搜索分成两个工具
- 真实数据不进仓库：日志、测试文件夹、生成的报告都在 .gitignore 里
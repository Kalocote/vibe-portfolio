from pathlib import Path
from datetime import datetime
import shutil

print("=== 运行时间:", datetime.now(), "===")

folder = Path.home() / "Downloads"
rules = {
    ".png": "图片",
    ".heic":"图片",
    ".jpg": "图片",
    ".jpeg":"图片",
    ".pdf": "文档",
    ".pptx":"PPT",
    ".xlsx":"表格",
    ".csv":"表格",
    ".zip": "压缩包",
    ".app":"应用",
}

for f in list (folder.iterdir()):
    if f.name.startswith("."):
        continue
    if f.is_dir() and f.suffix.lower() != ".app":
        continue
    kind = rules.get(f.suffix.lower(),"其他")
    target_dir = folder / kind
    target_dir.mkdir(exist_ok=True)

    target_file = target_dir / f.name
    if target_file.exists():
        print("跳过，(已存在):", f.name)
        continue

    try:
        shutil.move(f, target_file)
        print("已移动: ", f.name,"->",kind)
    except OSError as e:
        print("移动失败：", f.name,"，原因：",e)
# -*- coding: utf-8 -*-
"""
根据文件夹内容给 Excel 的"核心文献本名"列上色
四类：
  绿色 - 已有文件夹且里面是 txt
  红色 - 是 pdf
  黄色 - 是截图
  灰色 - 还没有文件夹
"""
import re
import shutil
from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ================= 配置 =================
ROOT = Path(r"D:\project\软件系统分析与设计综合\14组团队沟通与激励\3z - 副本\_按负责人")                     # 主文件夹路径
EXCEL_PATH = Path(r"副本第14组.xlsx")        # Excel 路径    # Excel 路径
EXCEL_SHEET = 0                                 # 工作表索引
HEADER_ROW  = 3                                 # 表头在第几行（1 起算）
SAVE_AS_NEW = True   # True: 另存为 xxx_已上色.xlsx；False: 覆盖原文件(会先备份)

IMG_EXTS = {".png",".jpg",".jpeg",".bmp",".tif",".tiff",".gif",".webp"}
PDF_EXTS = {".pdf"}
TXT_EXTS = {".txt"}

# 颜色 (RGB hex，不带 #)
COLOR_TXT  = "C6EFCE"   # 浅绿
COLOR_PDF  = "FFC7CE"   # 浅红
COLOR_IMG  = "FFEB9C"   # 浅黄
COLOR_NONE = "D9D9D9"   # 灰色

FILL_TXT  = PatternFill("solid", fgColor=COLOR_TXT)
FILL_PDF  = PatternFill("solid", fgColor=COLOR_PDF)
FILL_IMG  = PatternFill("solid", fgColor=COLOR_IMG)
FILL_NONE = PatternFill("solid", fgColor=COLOR_NONE)

# 扫描时跳过这些辅助目录
SKIP_TOP_DIRS = {"_重复备份", "_需OCR_pdf", "_需OCR_截图"}
# 允许扫描的辅助目录（按负责人分组后文件夹会进到这里）
KEEP_TOP_DIRS = {"_按负责人"}

FOLDER_PAT = re.compile(r"^(\d+)(.+)$")

# ================= 工具函数 =================
def strip_marks(s):
    return str(s).replace("《", "").replace("》", "").strip()

def normalize(s):
    if s is None:
        return ""
    s = str(s).replace("《","").replace("》","")
    s = re.sub(r"[\s\u3000\-_()（）\[\]【】<>〈〉,，.。、;；:：'\"`]+", "", s)
    return s.lower()

def scan_folders():
    """收集所有 '编号+名字' 的文件夹 → {规范化名: 路径}"""
    result = {}
    if not ROOT.exists():
        return result
    for p in ROOT.rglob("*"):
        if not p.is_dir():
            continue
        rel = p.relative_to(ROOT)
        top = rel.parts[0]
        if top in SKIP_TOP_DIRS:
            continue
        # 深度限制：ROOT/xxx  或  _按负责人/负责人/xxx
        if len(rel.parts) > 3:
            continue
        m = FOLDER_PAT.match(p.name)
        if not m:
            continue
        key = normalize(m.group(2))
        if not key:
            continue
        result.setdefault(key, p)   # 若有重名，先出现的优先
    return result

def classify_folder(folder: Path) -> str:
    """
    返回 'txt' / 'pdf' / 'img' / 'empty'
    优先级：pdf > img > txt
    （如果想让"只要有 txt 就绿"，把下面三个 if 的顺序改成 txt 优先即可）
    """
    has_txt = has_pdf = has_img = False
    for p in folder.rglob("*"):
        if not p.is_file():
            continue
        ext = p.suffix.lower()
        if   ext in PDF_EXTS: has_pdf = True
        elif ext in IMG_EXTS: has_img = True
        elif ext in TXT_EXTS: has_txt = True
    if has_txt: return "txt"
    if has_pdf: return "pdf"
    if has_img: return "img"

    return "empty"

def find_folder(key: str, index: dict):
    """先精确匹配，再包含匹配"""
    if not key:
        return None
    if key in index:
        return index[key]
    for k, v in index.items():
        if k and (k in key or key in k):
            return v
    return None

# ================= 主流程 =================
print("扫描文件夹 ...")
folder_index = scan_folders()
print(f"  共发现 {len(folder_index)} 个规范文件夹")

# 处理输出路径
if SAVE_AS_NEW:
    out_path = EXCEL_PATH.with_name(EXCEL_PATH.stem + "_已上色" + EXCEL_PATH.suffix)
else:
    backup = EXCEL_PATH.with_name(EXCEL_PATH.stem + "_备份" + EXCEL_PATH.suffix)
    if not backup.exists():
        shutil.copy2(EXCEL_PATH, backup)
        print(f"已备份原文件 → {backup.name}")
    out_path = EXCEL_PATH

print(f"加载 Excel：{EXCEL_PATH.name}")
wb = load_workbook(EXCEL_PATH)
ws = wb[wb.sheetnames[EXCEL_SHEET]] if isinstance(EXCEL_SHEET, int) else wb[EXCEL_SHEET]

# 定位"核心文献本名"列
headers = {}
for col in range(1, ws.max_column + 1):
    v = ws.cell(HEADER_ROW, col).value
    if v is not None:
        headers[str(v).strip()] = col

need = "核心文献本名"
if need not in headers:
    raise SystemExit(f"第 {HEADER_ROW} 行没找到列 '{need}'，实际列名：{list(headers.keys())}")
name_col = headers[need]
print(f"  '核心文献本名' 位于第 {name_col} 列")

# 逐行上色
stats = defaultdict(int)
unmatched = []

for row in range(HEADER_ROW + 1, ws.max_row + 1):
    cell = ws.cell(row, name_col)
    val = cell.value
    if val is None or str(val).strip() == "":
        continue

    name = strip_marks(val)
    key = normalize(name)
    folder = find_folder(key, folder_index)

    if folder is None:
        cell.fill = FILL_NONE
        stats["④ 还没有文件夹"] += 1
        unmatched.append((row, name))
        continue

    kind = classify_folder(folder)
    if kind == "txt":
        cell.fill = FILL_TXT;  stats["① 有文件夹且是txt"] += 1
    elif kind == "pdf":
        cell.fill = FILL_PDF;  stats["② 是pdf"] += 1
    elif kind == "img":
        cell.fill = FILL_IMG;  stats["③ 是截图"] += 1
    else:
        cell.fill = FILL_NONE; stats["④ 文件夹为空"] += 1

wb.save(out_path)
print(f"\n已保存到：{out_path}")
print("\n===== 统计 =====")
for k in ["① 有文件夹且是txt", "② 是pdf", "③ 是截图", "④ 还没有文件夹", "④ 文件夹为空"]:
    if k in stats:
        print(f"  {k}: {stats[k]} 行")

if unmatched:
    print(f"\n未匹配到文件夹的前 20 条（共 {len(unmatched)} 条）：")
    for r, n in unmatched[:20]:
        print(f"  第 {r} 行: {n}")
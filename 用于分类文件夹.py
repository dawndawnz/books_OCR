# -*- coding: utf-8 -*-
"""
文献文件夹批量整理
功能：
1. 闲散文件 -> 按 Excel 核心文献本名(去书名号)归入"编号+名字"规范文件夹
2. 找出重复编号的文件夹，把多出来的移入 _重复备份
3. 找出内部含 pdf / 截图 的非 txt 文件夹
4. 把这些文件夹里的 pdf、截图分别拷到 _需OCR_pdf 和 _需OCR_截图
5. 按"负责人"建文件夹，把编号+名字的规范文件夹移进对应负责人目录
"""
import re
import shutil
from pathlib import Path
from collections import defaultdict

import pandas as pd

# ================= 1. 用户配置 =================
ROOT = Path(r"D:\project\软件系统分析与设计综合\14组团队沟通与激励\3z - 副本")                     # 主文件夹路径
EXCEL_PATH = Path(r"副本第14组.xlsx")        # Excel 路径
EXCEL_SHEET = 0                                # 工作表索引
EXCEL_HEADER_ROW = 2                           # 表头所在行(0 起算)

DIR_DUP   = ROOT / "_重复备份"        # 重复文件夹移到这里（安全删除）
DIR_PDF   = ROOT / "_需OCR_pdf"       # 非 txt 中的 pdf 拷贝
DIR_IMG   = ROOT / "_需OCR_截图"      # 非 txt 中的截图拷贝
DIR_OWNER = ROOT / "_按负责人"        # 负责人分组目录

IMG_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif", ".webp"}
PDF_EXTS = {".pdf"}
TXT_EXTS = {".txt"}
OTHER_EXTS = {".doc", ".docx", ".caj", ".wps", ".rtf"}

SKIP_DIRS = {DIR_DUP.name, DIR_PDF.name, DIR_IMG.name, DIR_OWNER.name}

# ================= 2. 工具函数 =================
FOLDER_PAT = re.compile(r"^(\d+)(.*)$")

def parse_folder(name: str):
    """'6267沙面罢工' -> ('6267', '沙面罢工')"""
    m = FOLDER_PAT.match(name)
    return (m.group(1), m.group(2).strip()) if m else (None, None)

def normalize(s):
    """归一化字符串用于匹配：去书名号、空格、标点、大小写"""
    if s is None:
        return ""
    s = str(s).replace("《", "").replace("》", "")
    s = re.sub(r"[\s\u3000\-_()（）\[\]【】<>〈〉,，.。、;；:：]+", "", s)
    return s.lower()

def strip_book_marks(s):
    return str(s).replace("《", "").replace("》", "").strip()

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def list_top():
    """返回主文件夹下尚未处理的 (文件夹列表, 文件列表)"""
    folders, files = [], []
    if not ROOT.exists():
        return folders, files
    for p in ROOT.iterdir():
        if p.name in SKIP_DIRS or p.name.startswith("_"):
            continue
        (folders if p.is_dir() else files).append(p)
    return folders, files

def safe_dst(folder: Path, filename: str) -> Path:
    """同名时自动加后缀避免覆盖"""
    dst = folder / filename
    if not dst.exists():
        return dst
    stem, suf = Path(filename).stem, Path(filename).suffix
    i = 1
    while True:
        cand = folder / f"{stem}_dup{i}{suf}"
        if not cand.exists():
            return cand
        i += 1

# ================= 3. 读取 Excel =================
print(f"读取 Excel：{EXCEL_PATH}")
df = pd.read_excel(EXCEL_PATH, sheet_name=EXCEL_SHEET, header=EXCEL_HEADER_ROW)
df.columns = [str(c).strip() for c in df.columns]
print("识别到的列：", list(df.columns))

for col in ("负责人", "序号", "核心文献本名"):
    if col not in df.columns:
        raise SystemExit(f"Excel 缺少列 '{col}'。请检查 EXCEL_HEADER_ROW 是否指对了表头行。")

df = df.dropna(subset=["核心文献本名"]).copy()
df["_name"]  = df["核心文献本名"].apply(strip_book_marks)
df["_num"]   = df["序号"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
df["_owner"] = df["负责人"].astype(str).str.strip()

# 名称 -> [(序号, 负责人), ...]
name_info = defaultdict(list)
for _, r in df.iterrows():
    name_info[r["_name"]].append((r["_num"], r["_owner"]))

# 编号 -> 负责人
num_to_owner = {}
for _, infos in name_info.items():
    for num, owner in infos:
        num_to_owner.setdefault(num, owner)

# 归一化索引
name_norm = {normalize(n): n for n in name_info}

def match_name(text):
    if not text:
        return None
    key = normalize(text)
    if not key:
        return None
    if key in name_norm:
        return name_norm[key]
    # 包含匹配（处理尾部的 _dup1、_1 等）
    for k, v in name_norm.items():
        if k and (k in key or key in k):
            return v
    return None

# ================= 任务 1：闲散文件归入规范文件夹 =================
print("\n=== 任务 1：闲散文件归入规范文件夹 ===")
_, files = list_top()
moved = 0
for f in files:
    ext = f.suffix.lower()
    if ext not in (TXT_EXTS | PDF_EXTS | IMG_EXTS | OTHER_EXTS):
        print(f"  [跳过未知类型] {f.name}")
        continue
    name = match_name(f.stem)
    if name is None:
        print(f"  [无匹配] {f.name}")
        continue
    num, owner = name_info[name][0]
    target = ROOT / f"{num}{name}"
    ensure_dir(target)
    shutil.move(str(f), str(safe_dst(target, f.name)))
    moved += 1
    print(f"  [移动] {f.name}  ->  {target.name}/")
print(f"任务 1 完成：共移动 {moved} 个文件")

# ================= 任务 2：重复编号文件夹 =================
print("\n=== 任务 2：重复编号文件夹 ===")
ensure_dir(DIR_DUP)
folders, _ = list_top()
by_num = defaultdict(list)
for d in folders:
    n, _ = parse_folder(d.name)
    if n is not None:
        by_num[n].append(d)

dup_count = 0
for num, lst in by_num.items():
    if len(lst) <= 1:
        continue
    lst.sort(key=lambda p: sum(1 for _ in p.rglob("*")), reverse=True)  # 文件最多的保留
    keep, dups = lst[0], lst[1:]
    for dup in dups:
        dst = DIR_DUP / dup.name
        i = 1
        while dst.exists():
            dst = DIR_DUP / f"{dup.name}_{i}"
            i += 1
        shutil.move(str(dup), str(dst))
        dup_count += 1
        print(f"  [重复编号 {num}] 保留: {keep.name}  → 移走: {dup.name}")
print(f"任务 2 完成：共处理 {dup_count} 个重复文件夹（已备份到 {DIR_DUP.name}/）")

# ================= 任务 3：找出非 txt 文件夹 =================
print("\n=== 任务 3：找出非 txt（含 pdf / 截图）的文件夹 ===")
folders, _ = list_top()
need_ocr = []   # [(path, has_pdf, has_img, exts)]
for d in folders:
    n, _ = parse_folder(d.name)
    if n is None:
        continue
    exts = set()
    for p in d.rglob("*"):
        if p.is_file():
            exts.add(p.suffix.lower())
    has_pdf = bool(exts & PDF_EXTS)
    has_img = bool(exts & IMG_EXTS)
    has_txt = bool(exts & TXT_EXTS)
    if has_pdf or has_img:
        need_ocr.append((d, has_pdf, has_img, exts))
        print(f"  [需 OCR] {d.name}  内容: {sorted(exts)}  txt={'有' if has_txt else '无'}")
print(f"任务 3 完成：共找到 {len(need_ocr)} 个非 txt 文件夹")

# ================= 任务 4：拷贝到 pdf / 截图 文件夹 =================
print("\n=== 任务 4：拷贝到 _需OCR_pdf / _需OCR_截图 ===")
ensure_dir(DIR_PDF)
ensure_dir(DIR_IMG)
n_pdf = n_img = 0
for d, has_pdf, has_img, _ in need_ocr:
    if has_pdf:
        tgt = DIR_PDF / d.name
        for p in d.rglob("*"):
            if p.is_file() and p.suffix.lower() in PDF_EXTS:
                rel = p.relative_to(d)
                (tgt / rel).parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, tgt / rel)
                n_pdf += 1
    if has_img:
        tgt = DIR_IMG / d.name
        for p in d.rglob("*"):
            if p.is_file() and p.suffix.lower() in IMG_EXTS:
                rel = p.relative_to(d)
                (tgt / rel).parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(p, tgt / rel)
                n_img += 1
print(f"任务 4 完成：拷贝 pdf {n_pdf} 个，截图 {n_img} 个")

# ================= 任务 5：按负责人分组 =================
print("\n=== 任务 5：按负责人分组 ===")
ensure_dir(DIR_OWNER)
folders, _ = list_top()
owner_cnt = defaultdict(int)
for d in folders:
    n, _ = parse_folder(d.name)
    if n is None:
        print(f"  [跳过非规范文件夹] {d.name}")
        continue
    owner = num_to_owner.get(n)
    if not owner or owner.lower() in ("nan", "none", "null", ""):
        print(f"  [找不到负责人] {d.name}  (编号 {n})")
        continue
    tgt_root = DIR_OWNER / owner
    ensure_dir(tgt_root)
    dst = tgt_root / d.name
    if dst.exists():
        print(f"  [已存在跳过] {d.name}")
        continue
    shutil.move(str(d), str(dst))
    owner_cnt[owner] += 1
    print(f"  [移入] {d.name}  ->  {owner}/")
print("\n各负责人文件夹数量：")
for o, c in owner_cnt.items():
    print(f"  {o}: {c} 个")
print("\n全部完成 ✅")
# -*- coding: utf-8 -*-
"""
PDF -> 分章节 TXT  批量傻瓜版 (递归扫描)
- 自动单/双栏检测
- 自动分离页底脚注
- 递归处理当前目录及所有子文件夹里的 PDF
- 输出结构镜像输入结构
直接运行: python ocr_book.py
"""

import re, time
from pathlib import Path

import numpy as np
import fitz  # PyMuPDF


# ==================== 傻瓜配置 (一般不用改) ====================
DPI             = 300     # 图片精度, 越高越准但越慢
FORCE_COLUMNS   = 0       # 0=自动检测, 1=强制单栏, 2=强制双栏
FOOTNOTE_RATIO  = 0.88    # 页面下方超过此比例的区域视为脚注
MIN_SCORE       = 0.5     # 识别置信度阈值
SHORT_RATIO     = 0.8     # 合并断行的松紧, 越小段落越少
OUT_DIR         = "output"
# =============================================================


# -------------------------- OCR 引擎 --------------------------
class RapidEngine:
    """兼容 rapidocr 各版本返回格式"""
    def __init__(self, min_score=0.5):
        from rapidocr_onnxruntime import RapidOCR
        self.min_score = float(min_score)
        self.ocr = RapidOCR()

    def run(self, img_bgr):
        raw = self.ocr(img_bgr)
        result = raw[0] if isinstance(raw, tuple) else raw
        if not result:
            return []

        if hasattr(result, "boxes") and hasattr(result, "txts"):
            boxes  = list(result.boxes)
            txts   = list(result.txts)
            scores = list(getattr(result, "scores", [1.0] * len(txts)))
            return self._filter(boxes, txts, scores)

        boxes, txts, scores = [], [], []
        for item in result:
            b, t, s = self._parse_item(item)
            if b is None:
                continue
            boxes.append(b); txts.append(t); scores.append(s)
        return self._filter(boxes, txts, scores)

    def _filter(self, boxes, txts, scores):
        lines = []
        for b, t, s in zip(boxes, txts, scores):
            try:
                s = float(s)
            except (TypeError, ValueError):
                s = 1.0
            if s >= self.min_score:
                lines.append((b, t))
        return lines

    @staticmethod
    def _parse_item(item):
        try:
            if isinstance(item, dict):
                return (item.get("box") or item.get("dt_polys"),
                        item.get("text") or item.get("rec_text"),
                        item.get("score", item.get("rec_score", 1.0)))
            n = len(item)
            if n >= 3:
                a, b, c = item[0], item[1], item[2]
                if isinstance(b, str):
                    return a, b, c
                if isinstance(b, (list, tuple)) and len(b) >= 2:
                    return a, b[0], b[1]
            if n == 2:
                a, b = item
                if isinstance(b, (list, tuple)) and len(b) >= 2:
                    return a, b[0], b[1]
                if isinstance(b, str):
                    return a, b, 1.0
        except Exception:
            pass
        return None, None, None


# -------------------------- 自动检测栏数 --------------------------
def detect_columns(lines, page_width, force=0):
    if force in (1, 2):
        return force
    if len(lines) < 4:
        return 1
    centers = sorted(sum(p[0] for p in box) / len(box) for box, _ in lines)
    max_gap, gap_center = 0, 0
    for a, b in zip(centers, centers[1:]):
        g = b - a
        if g > max_gap:
            max_gap, gap_center = g, (a + b) / 2
    if max_gap > page_width * 0.10 and page_width * 0.35 < gap_center < page_width * 0.65:
        left  = sum(1 for c in centers if c <  gap_center)
        right = len(centers) - left
        if left >= 2 and right >= 2:
            return 2
    return 1


# -------------------------- 脚注分离 --------------------------
def split_footnotes(lines, page_height, ratio=0.88):
    if not lines:
        return [], []
    y_thr = page_height * ratio
    body, fn = [], []
    for box, text in lines:
        top = min(p[1] for p in box)
        (fn if top >= y_thr else body).append((box, text))
    if len(fn) < 2:
        return lines, []
    return body, fn


# -------------------------- 按栏重排 --------------------------
def sort_lines(lines, columns=1):
    if not lines:
        return []
    top  = lambda b: min(p[1] for p in b)
    left = lambda b: min(p[0] for p in b)

    if columns <= 1:
        return [t for _, t in sorted(lines, key=lambda x: (top(x[0]), left(x[0])))]

    xs = [sum(p[0] for p in b) / len(b) for b, _ in lines]
    mid = (min(xs) + max(xs)) / 2.0
    L = [t for t in lines if sum(p[0] for p in t[0]) / len(t[0]) <  mid]
    R = [t for t in lines if sum(p[0] for p in t[0]) / len(t[0]) >= mid]
    if not L or not R:
        return [t for _, t in sorted(lines, key=lambda x: (top(x[0]), left(x[0])))]
    L.sort(key=lambda x: (top(x[0]), left(x[0])))
    R.sort(key=lambda x: (top(x[0]), left(x[0])))
    return [t for _, t in L] + [t for _, t in R]


# -------------------------- 逐页 OCR --------------------------
def ocr_pdf(pdf_path, cache_dir, engine):
    pdf_path  = Path(pdf_path)
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    n = doc.page_count
    print(f"  共 {n} 页, 开始识别...\n")

    for i in range(n):
        cf = cache_dir / f"p{i+1:04d}.txt"
        if cf.exists():
            print(f"    [{i+1}/{n}] 已缓存, 跳过")
            continue

        t0 = time.time()
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=DPI, alpha=False)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = img[:, :, :3]
        img_bgr = np.ascontiguousarray(img[:, :, ::-1])

        try:
            lines = engine.run(img_bgr)
        except Exception as e:
            print(f"    [!] 第 {i+1} 页出错: {e}")
            lines = []

        body, fns = split_footnotes(lines, pix.height, FOOTNOTE_RATIO)
        cols = detect_columns(body, pix.width, FORCE_COLUMNS)

        body_txt = sort_lines(body, cols)
        fn_txt   = sort_lines(fns, 1)

        content = "\n".join(body_txt)
        if fn_txt:
            content += "\n\n=====FOOTNOTES=====\n" + "\n".join(fn_txt)
        cf.write_text(content, encoding="utf-8")

        flag = f"双栏" if cols == 2 else f"单栏"
        fn_flag = f" 脚注{len(fn_txt)}行" if fn_txt else ""
        print(f"    [{i+1}/{n}] {flag} {len(body_txt):>3}行{fn_flag} | {time.time()-t0:.1f}s")

    doc.close()


# -------------------------- 合并断行 --------------------------
HEADING_RE = re.compile(
    r'^\s*('
    r'第\s*[0-9零一二三四五六七八九十百千两]+\s*[章回节卷篇部]'
    r'|Chapter\s+\d+'
    r'|CHAPTER\s+\d+'
    r'|序\s*[言章]?\s*$|前言\s*$|后记\s*$|楔子\s*$|尾声\s*$|附录[A-Za-z]?\s*$'
    r')',
    re.UNICODE
)
SENT_END = re.compile(r'[。！？!?…”』」》〉）\)】\]；;]\s*$')
PAGENUM  = re.compile(
    r'^[-—\s]*\d{1,4}[-—\s]*$'
    r'|^第\s*\d+\s*页.*$'
    r'|^[-—\s]*[ivxlcdm]{1,7}[-—\s]*$',
    re.IGNORECASE
)


def merge_lines(raw, short_ratio=0.8):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return ""
    lens = sorted(len(l) for l in lines)
    full = max(lens[int(len(lens) * 0.8)], 15)

    out, buf = [], []
    def flush():
        if buf:
            out.append("".join(buf)); buf.clear()

    for s in lines:
        if PAGENUM.match(s):
            continue
        if HEADING_RE.match(s) and len(s) <= 40:
            flush(); out.append(s); continue
        buf.append(s)
        if SENT_END.search(s) and len(s) < full * short_ratio:
            flush()
    flush()
    return "\n\n".join(out)


# -------------------------- 分章节 --------------------------
def safe_name(name, n=40):
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip()
    return name[:n] or "untitled"


def split_chapters(clean_text, out_dir):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    paras = clean_text.split("\n\n")
    marks = [i for i, p in enumerate(paras)
             if HEADING_RE.match(p.strip()) and len(p.strip()) <= 40]

    print(f"    检测到 {len(marks)} 个章节标题")

    if not marks:
        (out_dir / "00_full.txt").write_text(clean_text, encoding="utf-8")
        print("    未找到章节标题, 整体输出到 00_full.txt")
        return

    if marks[0] > 0:
        head = "\n\n".join(paras[:marks[0]]).strip()
        if head:
            (out_dir / "000_前言.txt").write_text(head, encoding="utf-8")

    for k, idx in enumerate(marks):
        end = marks[k+1] if k+1 < len(marks) else len(paras)
        body = "\n\n".join(paras[idx:end]).strip()
        title = paras[idx].strip().splitlines()[0]
        fn = out_dir / f"{k+1:03d}_{safe_name(title)}.txt"
        fn.write_text(body, encoding="utf-8")


# -------------------------- 处理单个 PDF --------------------------
def process_one(pdf, engine, out_root, root):
    """pdf: PDF 路径; out_root: 输出根目录; root: 扫描根, 用于算相对路径"""
    try:
        rel = pdf.relative_to(root)
    except ValueError:
        rel = Path(pdf.name)
    # 输出目录: output/<相对路径去掉.pdf>
    book_dir = out_root / rel.with_suffix("")
    cache    = book_dir / "cache"
    chapters = book_dir / "chapters"
    book_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"处理: {rel}")
    print(f"输出: {book_dir}")
    print(f"{'='*60}")

    ocr_pdf(pdf, cache, engine)

    raw_body, raw_fn = [], []
    for f in sorted(cache.glob("p*.txt")):
        txt = f.read_text(encoding="utf-8")
        if not txt.strip():
            continue
        if "=====FOOTNOTES=====" in txt:
            b, fn = txt.split("=====FOOTNOTES=====", 1)
            raw_body.append(b.strip())
            pg = int(f.stem[1:])
            raw_fn.append(f"--- 第 {pg} 页 ---\n{fn.strip()}")
        else:
            raw_body.append(txt.strip())

    raw = "\n".join(raw_body)
    (book_dir / "raw.txt").write_text(raw, encoding="utf-8")
    print(f"  raw.txt        ({len(raw)} 字)")

    if raw_fn:
        fn_all = "\n\n".join(raw_fn)
        (book_dir / "footnotes.txt").write_text(fn_all, encoding="utf-8")
        print(f"  footnotes.txt  ({len(raw_fn)} 页有脚注)")

    clean = merge_lines(raw, SHORT_RATIO)
    (book_dir / "clean.txt").write_text(clean, encoding="utf-8")
    print(f"  clean.txt      ({len(clean)} 字)")

    split_chapters(clean, chapters)


# -------------------------- 主流程 --------------------------
def main():
    print("=" * 60)
    print("  PDF -> 分章节 TXT  (递归扫描, 自动分栏, 自动分离脚注)")
    print("=" * 60)

    root = Path(".").resolve()
    out_root = (root / OUT_DIR).resolve()

    # 递归找所有 PDF, 排除 output 目录里的
    all_pdfs = []
    for p in sorted(root.rglob("*.pdf")):
        if out_root in p.parents or p.parent == out_root:
            continue
        all_pdfs.append(p)

    if not all_pdfs:
        print("\n!! 当前目录(含子目录)下没找到 PDF")
        input("\n按回车退出...")
        return

    # 按相对路径分两组展示
    top_pdfs = [p for p in all_pdfs if p.parent == root]
    sub_pdfs = [p for p in all_pdfs if p.parent != root]

    print(f"\n找到 {len(all_pdfs)} 个 PDF:")
    if top_pdfs:
        print(f"\n  [当前目录] {len(top_pdfs)} 个:")
        for p in top_pdfs:
            print(f"    - {p.name}")
    if sub_pdfs:
        print(f"\n  [子文件夹] {len(sub_pdfs)} 个:")
        # 按文件夹分组展示
        groups = {}
        for p in sub_pdfs:
            groups.setdefault(p.parent, []).append(p)
        for folder in sorted(groups.keys()):
            try:
                relf = folder.relative_to(root)
            except ValueError:
                relf = folder
            print(f"    {relf}/")
            for p in groups[folder]:
                print(f"      - {p.name}")

    print(f"\n按回车开始处理 (Ctrl+C 取消)...")
    try:
        input()
    except (KeyboardInterrupt, EOFError):
        print("已取消")
        return

    t0 = time.time()
    engine = RapidEngine(MIN_SCORE)
    done, failed = 0, 0

    for pdf in all_pdfs:
        try:
            process_one(pdf, engine, out_root, root)
            done += 1
        except KeyboardInterrupt:
            print("\n!! 用户中断")
            break
        except Exception as e:
            failed += 1
            print(f"\n!! 处理 {pdf.name} 失败: {e}")
            import traceback; traceback.print_exc()

    print(f"\n{'='*60}")
    print(f"完成! 成功 {done}, 失败 {failed}, 耗时 {time.time()-t0:.1f}s")
    print(f"结果在: {out_root}")
    print(f"输出结构与输入结构一致")
    print(f"{'='*60}")
    input("\n按回车退出...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n!! 出错了: {e}")
        import traceback; traceback.print_exc()
        input("\n按回车退出...")
#!/usr/bin/env python
"""并行解码证件图片/PDF 中的二维码。

输入 manifest.jsonl，每行至少含：savedPath（文件路径）、pid、name、certLabel
输出 qr.jsonl，每行在原有字段基础上增加：qr / certNo / sign / shortUrl / err

解码策略（依次尝试，命中即停）：
  1. 整图（含 PDF 内嵌 JPEG 抽取）
  2. 切块扫描 3×3、4×4（重叠 15%）——大图里二维码占比小时唯一有效的手段
  3. 灰度 + 自动对比度

注意：整图放大 2~3 倍在这套场景下是反效果（3737px 的图再放大会 OOM），不要用。

依赖：pip install zxing-cpp pillow numpy
"""

from __future__ import annotations

import argparse
import io
import json
import multiprocessing as mp
import pathlib
import sys


def extract_jpegs(data: bytes) -> list[bytes]:
    """从 PDF 字节流里抠出内嵌 JPEG（扫描件常见）。"""
    out, i = [], 0
    while True:
        s = data.find(b"\xff\xd8\xff", i)
        if s < 0:
            break
        e = data.find(b"\xff\xd9", s)
        if e < 0:
            break
        out.append(data[s:e + 2])
        i = e + 2
    return out


def load_images(path: pathlib.Path, max_pdf_pages: int = 3):
    from PIL import Image

    imgs = []
    if path.suffix.lower() == ".pdf":
        for jb in extract_jpegs(path.read_bytes())[:max_pdf_pages]:
            if len(jb) < 3000:
                continue
            try:
                imgs.append(Image.open(io.BytesIO(jb)).convert("RGB"))
            except Exception:
                continue
    else:
        try:
            imgs.append(Image.open(path).convert("RGB"))
        except Exception:
            pass
    return imgs


def decode_tiles(img, n: int, overlap: float = 0.15):
    """把图切成 n×n 带重叠的块，逐块解码。"""
    import zxingcpp

    w, h = img.size
    tw, th = max(w // n, 1), max(h // n, 1)
    ox, oy = int(tw * overlap), int(th * overlap)
    for j in range(n):
        for i in range(n):
            box = (max(0, i * tw - ox), max(0, j * th - oy),
                   min(w, (i + 1) * tw + ox), min(h, (j + 1) * th + oy))
            tile = img.crop(box)
            try:
                for r in zxingcpp.read_barcodes(tile):
                    if r.text:
                        return r.text
            except Exception:
                continue
    return ""


def decode_image(img) -> str:
    import zxingcpp
    from PIL import ImageOps

    variants = [img, img.convert("L"), ImageOps.autocontrast(img.convert("L"))]
    for im in variants:
        try:
            for r in zxingcpp.read_barcodes(im):
                if r.text:
                    return r.text
        except Exception:
            pass
    for n in (3, 4):
        t = decode_tiles(img, n)
        if t:
            return t
    return ""


def split_payload(text: str) -> dict:
    """形如 https://w.url.cn/s/XXXX#<编号>_<签名> 的通用拆解。"""
    base, _, frag = text.partition("#")
    no, _, sign = frag.partition("_")
    return {"shortUrl": base, "certNo": no, "sign": sign}


def work(item: dict) -> dict:
    rec = dict(item)
    rec.update({"qr": "", "certNo": "", "sign": "", "shortUrl": ""})
    path = pathlib.Path(item.get("savedPath", ""))
    if not path.exists():
        rec["err"] = "文件不存在"
        return rec
    try:
        imgs = load_images(path)
    except Exception as exc:
        rec["err"] = f"读取失败: {exc}"
        return rec
    if not imgs:
        rec["err"] = "无可解码图像"
        return rec
    for img in imgs:
        t = decode_image(img)
        if t:
            rec["qr"] = t
            rec.update(split_payload(t))
            break
    return rec


def main() -> int:
    ap = argparse.ArgumentParser(description="并行解码证件二维码")
    ap.add_argument("--manifest", required=True, help="manifest.jsonl，每行含 savedPath")
    ap.add_argument("--out", required=True, help="输出 qr.jsonl")
    ap.add_argument("--workers", type=int, default=4, help="进程数（4 即可，多了会内存吃紧）")
    ap.add_argument("--retry-failed", action="store_true", help="对已存在但没解出的记录重试")
    args = ap.parse_args()

    rows = [json.loads(l) for l in
            pathlib.Path(args.manifest).read_text(encoding="utf-8").splitlines() if l.strip()]

    done = {}
    out_path = pathlib.Path(args.out)
    if out_path.exists():
        for l in out_path.read_text(encoding="utf-8").splitlines():
            if l.strip():
                r = json.loads(l)
                k = r.get("savedPath", "")
                if k not in done or (not done[k].get("certNo") and r.get("certNo")):
                    done[k] = r

    todo = []
    for r in rows:
        k = r.get("savedPath", "")
        if not k:
            continue
        if k in done:
            if args.retry_failed and not done[k].get("certNo"):
                todo.append(r)
        else:
            todo.append(r)

    print(f"共 {len(rows)} 个文件，已处理 {len(done)}，本轮 {len(todo)}，进程 {args.workers}", flush=True)
    if not todo:
        return 0

    results = dict(done)
    ok = 0
    with mp.Pool(max(1, args.workers)) as pool:
        for i, rec in enumerate(pool.imap_unordered(work, todo, chunksize=2), 1):
            results[rec["savedPath"]] = rec
            if rec.get("certNo"):
                ok += 1
            if i % 50 == 0:
                print(f"  {i}/{len(todo)}  解出 {ok}", flush=True)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as fh:
        for rec in results.values():
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")

    total = sum(1 for r in results.values() if r.get("certNo"))
    print(f"完成：本轮解出 {ok}，累计 {total}/{len(results)}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())

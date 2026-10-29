#!/usr/bin/env python
"""身份证号合法性校验（GB 11643）。

用途：判"证件非本人"之前必须先验号。系统里登记的号码如果本身不合法，
那是录入错误，不能当成借证上报。

命令行：
    python id_check.py 110101199002300012
    python id_check.py --file ids.txt          # 一行一个
    python id_check.py --csv 名单.csv --col 身份证号码

作为模块：
    from id_check import id_valid
    ok, why = id_valid("110101199002300012")   # -> (False, "出生日期不合法（2000-10-62）")
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys

WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
CHECK = "10X98765432"


def id_valid(s: str) -> tuple[bool, str]:
    """返回 (是否合法, 原因)。不合法时原因为空表示合法。"""
    s = (s or "").strip().upper().replace(" ", "")
    if not s:
        return False, "为空"
    if len(s) != 18:
        return False, f"长度 {len(s)} 位（应为 18）"

    try:
        dt.date(int(s[6:10]), int(s[10:12]), int(s[12:14]))
    except ValueError:
        return False, f"出生日期不合法（{s[6:10]}-{s[10:12]}-{s[12:14]}）"

    if not s[:17].isdigit():
        return False, "前 17 位含非数字"

    expect = CHECK[sum(int(s[i]) * WEIGHTS[i] for i in range(17)) % 11]
    if s[17] != expect:
        return False, f"校验位不符（应为 {expect}，实为 {s[17]}）"

    return True, ""


def main() -> int:
    ap = argparse.ArgumentParser(description="身份证号合法性校验（GB 11643）")
    ap.add_argument("ids", nargs="*", help="待校验的身份证号")
    ap.add_argument("--file", help="每行一个号码的文本文件")
    ap.add_argument("--csv", help="CSV 文件")
    ap.add_argument("--col", help="CSV 中的列名（与 --csv 搭配）")
    ap.add_argument("--quiet", action="store_true", help="只输出不合法的")
    args = ap.parse_args()

    items: list[str] = list(args.ids)
    if args.file:
        items += [l.strip() for l in open(args.file, encoding="utf-8") if l.strip()]
    if args.csv and args.col:
        import csv
        with open(args.csv, encoding="utf-8-sig", newline="") as fh:
            items += [row.get(args.col, "").strip() for row in csv.DictReader(fh)]

    if not items:
        ap.print_help()
        return 1

    bad = 0
    for s in items:
        ok, why = id_valid(s)
        if ok:
            if not args.quiet:
                print(f"OK    {s}")
        else:
            bad += 1
            print(f"BAD   {s}  ——  {why}")

    print(f"\n共 {len(items)} 条，不合法 {bad} 条", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())

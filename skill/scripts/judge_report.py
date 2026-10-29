#!/usr/bin/env python
"""核验判定与出表（整条流水线里最需要确定性的一段）。

把"判定口径"从模型手里拿回来，写成代码。输入官方核验结果与业务系统名单，
输出与人工资讯同构的 Excel。中间不做任何自由发挥。

用法：
    python judge_report.py \
        --verified  official_verify.jsonl \
        --staff     allstaff.xls \
        --manual    jining.xlsx linyi.xlsx \
        --contaminated contaminated_pids.json \
        --out       查验结果.xlsx

输入说明：
    verified.jsonl  每行 {"pid","certLabel","archivesNum","result":"found","official":{...}}
                    official 至少含 name / idcardNum / jobTypeName / operItemName /
                    certStatus / validBeginDate / validEndDate / shouldReviewDate / issueOrgName
                    （字段名不同时改 OFFICIAL_KEYS）
    staff.xls       业务系统全量导出，需含列：登录账号、代维人员姓名、真实姓名、
                    身份证号码、所属省份、所属地市、所属区域、所属代维专业
    manual          已人工查验过的名单（可省略），按第 6 列账号剔除
    contaminated    数据串位人员 pid 列表（JSON 数组），这些人结论作废

判定规则见 references/judgement-rules.md，本脚本与其一一对应。
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import pathlib
import re
import sys
from collections import defaultdict

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 官方字段名映射（换个平台只改这里）
OFFICIAL_KEYS = {
    "name": "name", "idcard": "idcardNum", "jobType": "jobTypeName",
    "operItem": "operItemName", "status": "certStatus",
    "validFrom": "validBeginDate", "validTo": "validEndDate",
    "review": "shouldReviewDate", "issuer": "issueOrgName",
}

# 证件类型 -> (作业类别, 准操项目)。准操项目为 None 表示不校验
EXPECT = {
    "低压电工证": ("电工作业", "低压电工作业"),
    "高压电工证": ("电工作业", "高压电工作业"),
    "登高证": ("高处作业", "高处安装、维护、拆除作业"),
}

HEAD = ["序号", "省份", "地市", "区县", "维护人员姓名", "维护系统账号", "代维专业",
        "证件", "问题分类", "问题备注说明", "其他发现的问题", "检查人员"]

HEADERS_NEED = ["登录账号", "代维人员姓名", "真实姓名", "身份证号码",
                "所属省份", "所属地市", "所属区域", "所属代维专业"]

_W = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
_CHK = "10X98765432"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def id_valid(s: str) -> tuple[bool, str]:
    s = (s or "").strip().upper()
    if len(s) != 18:
        return False, f"长度{len(s)}位"
    try:
        dt.date(int(s[6:10]), int(s[10:12]), int(s[12:14]))
    except ValueError:
        return False, f"出生日期不合法（{s[6:10]}-{s[10:12]}-{s[12:14]}）"
    if _CHK[sum(int(s[i]) * _W[i] for i in range(17)) % 11] != s[17]:
        return False, "校验位不符"
    return True, ""


def name_anomaly(disp: str, real: str) -> str:
    """系统自身姓名字段的问题。区域前缀属命名惯例，不算。"""
    if not disp or not real:
        return ""
    if re.search(r"\d", disp):
        return f"系统显示名「{disp}」含数字编号"
    if real * 2 in disp:
        return f"系统真实姓名栏「{real}」重复填写"
    if real not in disp:
        return f"系统显示名「{disp}」与真实姓名「{real}」无对应关系"
    prefix = disp.replace(real, "", 1)
    if len(prefix) > 4:
        return f"系统显示名「{disp}」前缀过长（{prefix}）"
    return ""


def _quality(paths: list[str]) -> tuple[bool, str]:
    """粗判图片是否真的不可用（区别于"我们读不出"）。"""
    import io
    import numpy as np
    from PIL import Image, ImageFilter

    best = None
    for p in paths:
        fp = pathlib.Path(p)
        if not fp.exists():
            continue
        try:
            if fp.suffix.lower() == ".pdf":
                data = fp.read_bytes()
                i = data.find(b"\xff\xd8\xff")
                if i < 0:
                    continue
                j = data.find(b"\xff\xd9", i)
                img = Image.open(io.BytesIO(data[i:j + 2])).convert("L")
            else:
                img = Image.open(fp).convert("L")
        except Exception:
            continue
        m = {"w": img.width, "h": img.height, "kb": fp.stat().st_size // 1024,
             "sharp": float(np.asarray(img.filter(ImageFilter.FIND_EDGES), dtype=np.float32).var())}
        if best is None or m["kb"] > best["kb"]:
            best = m
    if best is None:
        return True, "证件文件无法打开"
    if best["kb"] < 80:
        return True, f"文件仅 {best['kb']}KB，疑似空白或截图"
    if max(best["w"], best["h"]) < 900:
        return True, f"分辨率仅 {best['w']}x{best['h']}，不足以辨认"
    return False, ""


def classify_missing(acc, files, qr_by_path, done_pids, acc_pids) -> tuple[str, str]:
    pid = acc_pids.get(acc)
    fs = files.get(acc, [])
    if done_pids is not None and pid and pid not in done_pids:
        return "未核验", "证件尚未完成下载，需补查"
    if not fs:
        return ("未核验", "系统内无特种作业证附件") if done_pids is not None else ("未取得核验", "名单无附件记录")
    if any((qr_by_path.get(f.get("savedPath", "")) or {}).get("certNo") for f in fs):
        return "未取得核验", "已取得二维码但未完成权威核验"
    poor, why = _quality([f.get("savedPath", "") for f in fs])
    if poor:
        return "无法识别", f"{why}，需人工查看原件"
    return "未取得核验", "证件图片清晰但二维码无法读取，属核验手段限制，需人工确认"


def load_staff(path: pathlib.Path) -> dict:
    wb = xlrd.open_workbook(str(path))
    sh = wb.sheet_by_index(0)
    col = {str(sh.cell_value(0, c)).strip(): c for c in range(sh.ncols)}
    missing = [h for h in HEADERS_NEED if h not in col]
    if missing:
        raise SystemExit(f"名单缺少必要列: {missing}")

    def v(r, c):
        x = sh.cell_value(r, c)
        if isinstance(x, float) and x == int(x):
            return str(int(x))
        return "" if x is None else str(x).strip()

    out = {}
    for r in range(1, sh.nrows):
        acc = v(r, col["登录账号"])
        if acc:
            out[acc] = {k: v(r, col[k]) for k in HEADERS_NEED}
    return out


def judge(rec: dict, s: dict, today: dt.date) -> tuple[list, list, str]:
    o = rec.get("official", {})
    g = lambda k: o.get(OFFICIAL_KEYS[k], "") or ""
    disp, real, oname = s.get("代维人员姓名", ""), s.get("真实姓名", ""), g("name")
    probs, others = [], []

    sysid, offid = s.get("身份证号码", "").upper(), g("idcard").upper()
    if sysid and offid and sysid != offid:
        ok, why = id_valid(sysid)
        if ok:
            probs.append(("证件非本人",
                          f"{rec['certLabel']}：档案编码 {rec.get('archivesNum','')} 经官方平台核验，"
                          f"官方身份证 {offid} 与系统登记 {sysid} 不符，证件可能非持证人所有"))
        else:
            probs.append(("系统身份证号有误",
                          f"{rec['certLabel']}：系统登记的身份证号 {sysid} 不合法（{why}），"
                          f"官方核验为 {offid}，应为系统录入错误，证件本身无误"))

    name_ok = bool(oname) and oname in (disp, real)
    if oname and not name_ok:
        probs.append(("姓名不符", f"{rec['certLabel']}：官方姓名 {oname} 与系统登记 {real or disp} 不一致"))
    elif name_ok:
        n = name_anomaly(disp, real)
        if n:
            others.append(n)

    if str(g("status")) != "1":
        probs.append(("证件状态异常",
                      f"{rec['certLabel']}：证件状态非正常（状态码 {g('status')}，"
                      f"有效期 {g('validFrom')} 至 {g('validTo')}）"))
    if g("validTo") and g("validTo") < str(today):
        probs.append(("证件已过期", f"{rec['certLabel']}：有效期止 {g('validTo')}，已过期"))
    if g("review") and g("review") < str(today):
        probs.append(("复审逾期", f"{rec['certLabel']}：应复审 {g('review')}，已逾期"))

    e = EXPECT.get(rec["certLabel"])
    if e:
        if g("jobType") != e[0]:
            probs.append(("证类不符", f"{rec['certLabel']}：系统登记{rec['certLabel']}，官方作业类别为 {g('jobType')}"))
        elif e[1] and g("operItem") != e[1]:
            probs.append(("操作项目不符", f"{rec['certLabel']}：系统登记{rec['certLabel']}，官方准操项目为 {g('operItem')}"))

    shown = oname if name_ok else (real or disp)
    return probs, others, shown


def main() -> int:
    ap = argparse.ArgumentParser(description="核验判定与出表")
    ap.add_argument("--verified", required=True, help="官方核验结果 jsonl")
    ap.add_argument("--staff", required=True, help="业务系统名单导出 xls/xlsx 首表")
    ap.add_argument("--manual", nargs="*", default=[], help="已人工查验的名单（取第6列账号剔除）")
    ap.add_argument("--contaminated", help="数据串位人员 pid JSON 数组")
    ap.add_argument("--pidmap", help="pid→账号 的 jsonl（含 pid 与 account），用于把串位名单对应到人")
    ap.add_argument("--out", required=True)
    ap.add_argument("--sheet2", help="可选：把名单原始数据放到 Sheet2（传原表路径）")
    ap.add_argument("--manifest", help="可选：附件清单 jsonl，用于细分未核验成因")
    ap.add_argument("--qr", help="可选：二维码解码结果 jsonl")
    ap.add_argument("--done", help="可选：下载进度 progress.json")
    args = ap.parse_args()

    staff = load_staff(pathlib.Path(args.staff))
    today = dt.date.today()

    skip = set()
    for f in args.manual:
        ws = load_workbook(f)["Sheet1"]
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=6).value
            if v:
                skip.add(str(v).strip())

    contam = set()
    if args.contaminated and pathlib.Path(args.contaminated).exists():
        contam = set(json.loads(pathlib.Path(args.contaminated).read_text(encoding="utf-8")))

    pid2acc = {}
    if args.pidmap and pathlib.Path(args.pidmap).exists():
        for line in pathlib.Path(args.pidmap).read_text(encoding="utf-8").splitlines():
            if line.strip():
                r = json.loads(line)
                pid2acc[r["pid"]] = r.get("account", "")

    best = {}
    for line in pathlib.Path(args.verified).read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue
        r = json.loads(line)
        k = (r.get("pid"), r.get("archivesNum"), r.get("certLabel"))
        if k not in best or (best[k].get("result") != "found" and r.get("result") == "found"):
            best[k] = r

    contam_accs = {pid2acc[p] for p in contam if p in pid2acc} if pid2acc else set()
    acc_pids = {v: k for k, v in pid2acc.items()}

    files_by_acc = defaultdict(list)
    if args.manifest and pathlib.Path(args.manifest).exists():
        for line in pathlib.Path(args.manifest).read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            m = json.loads(line)
            lab = str(m.get("certLabel", "")).replace("查看", "").strip()
            if lab in EXPECT:
                files_by_acc[pid2acc.get(m.get("pid", ""), "")].append({**m, "certLabel": lab})

    qr_by_path = {}
    if args.qr and pathlib.Path(args.qr).exists():
        for line in pathlib.Path(args.qr).read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            q = json.loads(line)
            k = q.get("savedPath", "")
            if k not in qr_by_path or (not qr_by_path[k].get("certNo") and q.get("certNo")):
                qr_by_path[k] = q

    done_pids = None
    if args.done and pathlib.Path(args.done).exists():
        done_pids = set(json.loads(pathlib.Path(args.done).read_text(encoding="utf-8")).get("done", []))

    by_acc = defaultdict(list)
    for r in best.values():
        if r.get("result") != "found":
            continue
        by_acc[pid2acc.get(r["pid"], "")].append(r)

    rows = []
    for acc, s in staff.items():
        if acc in skip:
            continue
        recs = by_acc.get(acc, [])

        # ① 数据串位优先判：这类人的全部结论作废
        if acc in contam_accs:
            rows.append({**s, "账号": acc,
                         "姓名": s.get("真实姓名") or s.get("代维人员姓名", ""),
                         "证件": ",".join(dict.fromkeys(f["certLabel"] for f in files_by_acc.get(acc, []))),
                         "分类": "需重核",
                         "备注": "详情面板未刷新导致证件文件错位，结论作废，需重新抓取核验",
                         "其他": ""})
            continue

        # ② 没有权威数据：按成因细分，绝不写成"正常"
        if not recs:
            kind, note = classify_missing(acc, files_by_acc, qr_by_path, done_pids, acc_pids)
            rows.append({**s, "账号": acc,
                         "姓名": s.get("真实姓名") or s.get("代维人员姓名", ""),
                         "证件": ",".join(dict.fromkeys(f["certLabel"] for f in files_by_acc.get(acc, []))),
                         "分类": kind, "备注": note, "其他": ""})
            continue

        kinds, notes, others, names, shown = [], [], [], [], ""
        for r in recs:
            p_, o_, nm = judge(r, s, today)
            kinds += [x[0] for x in p_]
            notes += [x[1] for x in p_]
            others += [x for x in o_ if x not in others]
            names.append(r["certLabel"])
            shown = shown or nm

        uniq = list(dict.fromkeys(kinds))
        rows.append({**s, "账号": acc,
                     "姓名": shown or s.get("真实姓名") or s.get("代维人员姓名"),
                     "证件": ",".join(dict.fromkeys(names)),
                     "分类": "、".join(uniq) if uniq else "无问题",
                     "备注": "；".join(notes), "其他": "；".join(others)})

    rows.sort(key=lambda x: (x.get("所属地市", ""), x.get("所属区域", ""),
                             x.get("姓名") or x.get("代维人员姓名", "")))
    write(rows, args)

    from collections import Counter
    c = Counter()
    for x in rows:
        for k in (x["分类"] or "无问题").split("、"):
            c[k] += 1
    print(f"已生成 {args.out}｜{len(rows)} 人")
    print("分类分布:", dict(c))
    return 0


def write(rows: list[dict], args) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEAD)
    for i, x in enumerate(rows, 1):
        ws.append([i, x.get("所属省份", ""), x.get("所属地市", ""), x.get("所属区域", ""),
                   x.get("姓名") or x.get("代维人员姓名", ""), x["账号"],
                   x.get("所属代维专业", ""), x["证件"], x["分类"], x["备注"], x["其他"], ""])

    for i, w in enumerate([6, 13, 13, 13, 14, 17, 26, 24, 22, 70, 34, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    fills = {"red": "FFC7CE", "orange": "FFE0B2", "grey": "EDEDED", "green": "E2EFDA"}
    for r in range(2, ws.max_row + 1):
        k = ws.cell(row=r, column=9).value or ""
        for c in range(1, len(HEAD) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="微软雅黑", size=10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=c in (7, 8, 9, 10, 11),
                                       horizontal="center" if c in (1, 2, 3, 4, 12) else "left")
        if k == "无问题":
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["green"])
        elif k in ("无法识别", "未取得核验", "未核验", "需重核"):
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["grey"])
        elif any(x in k for x in ("非本人", "过期", "状态异常")):
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["red"])
        else:
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["orange"])
        n = max(math.ceil(len(str(ws.cell(row=r, column=10).value or "")) / 40),
                math.ceil(len(str(ws.cell(row=r, column=11).value or "")) / 20), 1)
        ws.row_dimensions[r].height = max(20, 14 * n + 6)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    if args.sheet2 and pathlib.Path(args.sheet2).exists():
        p = pathlib.Path(args.sheet2)
        if p.suffix.lower() == ".xls":
            sh = xlrd.open_workbook(str(p)).sheet_by_index(0)
            ws2 = wb.create_sheet("Sheet2")
            ws2.append([str(sh.cell_value(0, c)).strip() for c in range(sh.ncols)])
            keep = {x["账号"] for x in rows}
            arow = [i for i in range(sh.ncols) if str(sh.cell_value(0, i)).strip() == "登录账号"][0]
            for r in range(1, sh.nrows):
                if str(sh.cell_value(r, arow)).strip() in keep:
                    ws2.append([sh.cell_value(r, c) for c in range(sh.ncols)])
            ws2.freeze_panes = "A2"

    pathlib.Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)


if __name__ == "__main__":
    sys.exit(main())

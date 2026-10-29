"""证件核验报表工具（本地网页版，双击启动）。

给不懂技术的同事用：打开网页 → 选文件 → 点生成 → 下载 Excel。
不联网、不上传云端，所有处理都在本机完成。

启动：双击「启动工具.bat」，或 python app.py
"""

from __future__ import annotations

import base64
import datetime as dt
import io
import json
import pathlib
import re
import sys
import tempfile
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = pathlib.Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

PORT = 8765

# ---------------- 判定规则（与 vendor-cert-audit 技能一致）----------------

EXPECT = {
    "低压电工证": ("电工作业", "低压电工作业"),
    "高压电工证": ("电工作业", "高压电工作业"),
    "登高证": ("高处作业", "高处安装、维护、拆除作业"),
}
_W = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
_CHK = "10X98765432"
HEAD = ["序号", "省份", "地市", "区县", "维护人员姓名", "维护系统账号", "代维专业",
        "证件", "问题分类", "问题备注说明", "其他发现的问题", "检查人员"]
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def id_valid(s: str):
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


def name_anomaly(disp, real):
    if not disp or not real:
        return ""
    if re.search(r"\d", disp):
        return f"系统显示名「{disp}」含数字编号"
    if real * 2 in disp:
        return f"系统真实姓名栏「{real}」重复填写"
    if real not in disp:
        return f"系统显示名「{disp}」与真实姓名「{real}」无对应关系"
    if len(disp.replace(real, "", 1)) > 4:
        return f"系统显示名「{disp}」前缀过长"
    return ""


# ---------------- 读名单 ----------------

NEED = ["登录账号", "代维人员姓名", "真实姓名", "身份证号码",
        "所属省份", "所属地市", "所属区域", "所属代维专业"]


def read_staff(data: bytes, fname: str) -> list[dict]:
    """读业务系统导出的名单（.xls 或 .xlsx，取第一个工作表）。"""
    if fname.lower().endswith(".xls"):
        sh = xlrd.open_workbook(file_contents=data).sheet_by_index(0)
        heads = [str(sh.cell_value(0, c)).strip() for c in range(sh.ncols)]

        def val(r, c):
            x = sh.cell_value(r, c)
            if isinstance(x, float) and x == int(x):
                return str(int(x))
            return "" if x is None else str(x).strip()

        rows = [{heads[c]: val(r, c) for c in range(sh.ncols)} for r in range(1, sh.nrows)]
    else:
        ws = load_workbook(io.BytesIO(data), read_only=True, data_only=True).worksheets[0]
        it = ws.iter_rows(values_only=True)
        heads = [str(h or "").strip() for h in next(it)]
        rows = []
        for vals in it:
            if vals is None:
                continue
            rows.append({heads[i]: ("" if v is None else str(v).strip())
                         for i, v in enumerate(vals) if i < len(heads)})

    missing = [h for h in NEED if h not in (rows[0] if rows else {})]
    if missing:
        raise ValueError(f"名单缺少必要列：{'、'.join(missing)}")
    return [r for r in rows if r.get("登录账号")]


def read_verified(data: bytes) -> dict:
    """读官方核验结果 jsonl；也兼容「已生成的 xlsx」。"""
    out = {}
    text = data.decode("utf-8", "replace").strip()
    if text.startswith("{"):
        for line in text.splitlines():
            if not line.strip():
                continue
            r = json.loads(line)
            if r.get("result") != "found":
                continue
            out[r.get("account") or r.get("账号") or ""] = r
        return out
    return out


def detect_contaminated(manifest_rows: list[dict], name2acc: dict) -> set[str]:
    """检测跨人重复的附件（详情面板未刷新导致文件错位）。

    判据：同一 fileId 出现在两个以上人名下。若文件名里含某人姓名，那个人是真主，
    其余人判为错位。文件名不含姓名时（如"意外险.pdf"），全部判为错位。
    """
    by_fid = {}
    for m in manifest_rows:
        lab = str(m.get("certLabel", "")).replace("查看", "").strip()
        if lab not in EXPECT:
            continue
        by_fid.setdefault(m.get("fileId", ""), []).append(m)

    bad_pids = set()
    for items in by_fid.values():
        if len({i.get("pid") for i in items}) < 2:
            continue
        owner = next((i.get("pid") for i in items if i.get("name") and i["name"] in i.get("fileName", "")), None)
        for i in items:
            if i.get("pid") != owner:
                bad_pids.add(i.get("pid"))

    bad_accs = set()
    for m in manifest_rows:
        if m.get("pid") in bad_pids:
            acc = m.get("account") or ""
            if not acc:
                for k in {m.get("name", "")}:
                    cands = name2acc.get(k, [])
                    acc = cands[0] if len(cands) == 1 else ""
            if acc:
                bad_accs.add(acc)
    return bad_accs


# ---------------- 生成报告 ----------------

def build(rows: list[dict], verified: dict, skip: set[str], contaminated: set[str] | None = None) -> bytes:
    today = dt.date.today()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(HEAD)

    n = 0
    for s in rows:
        acc = s["登录账号"]
        if acc in skip:
            continue
        n += 1
        if contaminated and acc in contaminated:
            ws.append([n, s.get("所属省份", ""), s.get("所属地市", ""), s.get("所属区域", ""),
                       s.get("真实姓名") or s.get("代维人员姓名", ""), acc,
                       s.get("所属代维专业", ""), "", "需重核",
                       "附件与其他人重复，判断为详情页未刷新导致文件错位，结论作废，需重新抓取", "", ""])
            continue

        recs = verified.get(acc, [])
        kinds, notes, others, names, shown = [], [], [], [], ""

        if isinstance(recs, dict):
            recs = [recs]
        for r in recs:
            o = r.get("official", {})
            lab = r.get("certLabel", "")
            disp, real, oname = s.get("代维人员姓名", ""), s.get("真实姓名", ""), o.get("name", "")
            sysid = (s.get("身份证号码") or "").upper()
            offid = (o.get("idcardNum") or "").upper()
            if sysid and offid and sysid != offid:
                ok, why = id_valid(sysid)
                if ok:
                    kinds.append("证件非本人")
                    notes.append(f"{lab}：档案编码 {r.get('archivesNum','')} 经官方核验，官方身份证 {offid} 与系统 {sysid} 不符")
                else:
                    kinds.append("系统身份证号有误")
                    notes.append(f"{lab}：系统身份证号 {sysid} 不合法（{why}），官方为 {offid}，应为录入错误")
            name_ok = bool(oname) and oname in (disp, real)
            if oname and not name_ok:
                kinds.append("姓名不符")
                notes.append(f"{lab}：官方姓名 {oname} 与系统 {real or disp} 不一致")
            elif name_ok:
                a = name_anomaly(disp, real)
                if a and a not in others:
                    others.append(a)
            if str(o.get("certStatus", "1")) != "1":
                kinds.append("证件状态异常")
                notes.append(f"{lab}：状态码 {o.get('certStatus')}，有效期 {o.get('validBeginDate')} 至 {o.get('validEndDate')}")
            if o.get("validEndDate") and o["validEndDate"] < str(today):
                kinds.append("证件已过期")
                notes.append(f"{lab}：有效期止 {o['validEndDate']}")
            rev = o.get("shouldReviewDate") or o.get("reviewDate") or ""
            if rev and rev < str(today):
                kinds.append("复审逾期")
                notes.append(f"{lab}：应复审 {rev}")
            e = EXPECT.get(lab)
            if e:
                if o.get("jobTypeName") != e[0]:
                    kinds.append("证类不符")
                    notes.append(f"{lab}：官方作业类别为 {o.get('jobTypeName')}")
                elif e[1] and o.get("operItemName") != e[1]:
                    kinds.append("操作项目不符")
                    notes.append(f"{lab}：官方准操项目为 {o.get('operItemName')}")
            names.append(lab)
            shown = shown or (oname if name_ok else (real or disp))

        kind = "、".join(dict.fromkeys(kinds)) if kinds else ("无问题" if recs else "未取得核验")
        note = "；".join(notes) if recs else "未提供该人员的官方核验结果"
        ws.append([n, s.get("所属省份", ""), s.get("所属地市", ""), s.get("所属区域", ""),
                   shown or s.get("真实姓名") or s.get("代维人员姓名", ""), acc,
                   s.get("所属代维专业", ""), ",".join(dict.fromkeys(names)),
                   kind, note, "；".join(others), ""])

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
        elif k in ("未取得核验", "无法识别", "未核验"):
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["grey"])
        elif any(x in k for x in ("非本人", "过期", "状态异常")):
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["red"])
        else:
            ws.cell(row=r, column=9).fill = PatternFill("solid", fgColor=fills["orange"])
        ln = max(len(str(ws.cell(row=r, column=10).value or "")) // 40 + 1, 1)
        ws.row_dimensions[r].height = max(20, 14 * ln + 6)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------- 网页 ----------------

PAGE = """<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>证件核验报表工具</title>
<style>
:root{--bg:#0f1115;--card:#181b21;--line:#2a2f38;--fg:#e6e8ec;--dim:#8b93a1;--acc:#4c8dff;--ok:#3fb950;--warn:#d29922}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.6 "Microsoft YaHei",system-ui,sans-serif;padding:32px}
.wrap{max-width:760px;margin:0 auto}
h1{font-size:20px;margin:0 0 4px}
.sub{color:var(--dim);font-size:13px;margin-bottom:24px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px;margin-bottom:14px}
.card h2{font-size:14px;margin:0 0 10px;display:flex;align-items:center;gap:8px}
.tag{font-size:11px;padding:1px 7px;border-radius:20px;border:1px solid var(--line);color:var(--dim);font-weight:400}
.tag.req{color:var(--acc);border-color:var(--acc)}
.tag.opt{color:var(--warn);border-color:var(--warn)}
input[type=file]{width:100%;padding:10px;background:#12151a;border:1px dashed var(--line);border-radius:8px;color:var(--fg);font-size:13px}
.hint{color:var(--dim);font-size:12px;margin-top:6px}
button{width:100%;padding:13px;border:0;border-radius:9px;background:var(--acc);color:#fff;font-size:15px;font-weight:600;cursor:pointer;margin-top:6px}
button:disabled{background:#2a2f38;color:#6b7381;cursor:not-allowed}
#msg{margin-top:14px;padding:12px;border-radius:8px;display:none;font-size:13px;white-space:pre-wrap}
.ok{background:#0f2a17;border:1px solid #1f6f37;color:#7ee2a8}
.err{background:#2a1013;border:1px solid #7d2b35;color:#ff9aa5}
.rules{color:var(--dim);font-size:12px;line-height:1.9}
.rules b{color:#c9d1d9;font-weight:600}
</style></head><body><div class="wrap">
<h1>证件核验报表工具</h1>
<div class="sub">所有处理都在本机完成，不联网、不上传。</div>

<div class="card"><h2>① 人员名单 <span class="tag req">必填</span></h2>
<input type="file" id="staff" accept=".xls,.xlsx">
<div class="hint">业务系统导出的代维人员名单（.xls / .xlsx），需含：登录账号、代维人员姓名、真实姓名、身份证号码、所属省份/地市/区域/代维专业</div></div>

<div class="card"><h2>② 官方核验结果 <span class="tag opt">可选</span></h2>
<input type="file" id="verified" accept=".jsonl,.json">
<div class="hint">不提供也能生成：所有人会标为「未取得核验」，可用于先出一份催办清单。</div></div>

<div class="card"><h2>③ 附件清单 <span class="tag opt">强烈建议</span></h2>
<input type="file" id="manifest" accept=".jsonl,.json">
<div class="hint">证件附件的清单（manifest.jsonl）。勾上后会自动检测"详情面板未刷新导致的文件错位"，把受影响的标为「需重核」，避免把别人的证算到本人头上。</div></div>

<div class="card"><h2>④ 已人工查验名单 <span class="tag opt">可选</span></h2>
<input type="file" id="manual" accept=".xlsx" multiple>
<div class="hint">这些人的结果会被剔除，不重复出表。</div></div>

<button id="go" onclick="run()">生成核验报告</button>
<div id="msg"></div>

<div class="card" style="margin-top:22px"><h2>判定口径</h2><div class="rules">
<b>证件非本人</b>：官方身份证号与系统登记不符，且系统号本身合法<br>
<b>系统身份证号有误</b>：系统号本身不合法（校验位/出生日期错），属录入错误<br>
<b>姓名不符</b>：官方姓名既不等于显示名、也不等于真实姓名<br>
<b>证类不符 / 操作项目不符</b>：系统挂的证与官方作业类别或准操项目不一致（登高证只认「高处安装、维护、拆除作业」）<br>
<b>状态异常 / 已过期 / 复审逾期</b>：以官方返回的证件状态、有效期、应复审日期为准<br>
<b>未取得核验</b>：没有官方数据，不作结论
</div></div>
</div>
<script>
function b64(buf){let s="",u=new Uint8Array(buf);const CH=0x8000;
 for(let i=0;i<u.length;i+=CH)s+=String.fromCharCode.apply(null,u.subarray(i,i+CH));
 return btoa(s);}
async function readFile(f){const b=await f.arrayBuffer();return {name:f.name,data:b64(b)};}
function show(t,cls){const m=document.getElementById("msg");m.className=cls;m.textContent=t;m.style.display="block";}
async function run(){
  const st=document.getElementById("staff").files[0];
  if(!st){show("请先选择人员名单","err");return;}
  const btn=document.getElementById("go");btn.disabled=true;show("正在生成，请稍候…","");
  try{
    const payload={staff:await readFile(st),verified:null,manual:[]};
    const v=document.getElementById("verified").files[0];
    if(v)payload.verified=await readFile(v);
    const mf=document.getElementById("manifest").files[0];
    if(mf)payload.manifest=await readFile(mf);
    for(const f of document.getElementById("manual").files)payload.manual.push(await readFile(f));
    const r=await fetch("/generate",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
    const j=await r.json();
    if(!j.ok){show("生成失败：\\n"+j.error,"err");btn.disabled=false;return;}
    const bin=atob(j.file),arr=new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
    const a=document.createElement("a");
    a.href=URL.createObjectURL(new Blob([arr],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
    a.download=j.name;a.click();
    show(`完成：共 ${j.count} 人\\n`+j.summary+"\\n\\n文件已保存到浏览器下载目录","ok");
  }catch(e){show("出错："+e,"err");}
  btn.disabled=false;
}
</script></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        body = PAGE.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path != "/generate":
            self.send_error(404)
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))

            staff = read_staff(base64.b64decode(payload["staff"]["data"]), payload["staff"]["name"])

            # 核验结果可能按账号、也可能按姓名归档，两种都支持
            name2acc = {}
            for r in staff:
                for key in {r.get("代维人员姓名", ""), r.get("真实姓名", "")}:
                    if key:
                        name2acc.setdefault(key, []).append(r["登录账号"])

            verified = {}
            if payload.get("verified"):
                vdata = base64.b64decode(payload["verified"]["data"])
                used = set()
                for line in vdata.decode("utf-8", "replace").splitlines():
                    if not line.strip():
                        continue
                    r = json.loads(line)
                    if r.get("result") != "found":
                        continue
                    acc = r.get("account") or r.get("账号") or ""
                    if not acc:
                        cands = [a for a in name2acc.get(r.get("name", ""), []) if a not in used]
                        acc = cands[0] if cands else ""
                    if acc:
                        used.add(acc)
                        verified.setdefault(acc, []).append(r)

            skip = set()
            for m in payload.get("manual", []):
                tmp = pathlib.Path(tempfile.mkdtemp()) / "m.xlsx"
                tmp.write_bytes(base64.b64decode(m["data"]))
                ws = load_workbook(tmp)["Sheet1"]
                for r in range(2, ws.max_row + 1):
                    v = ws.cell(row=r, column=6).value
                    if v:
                        skip.add(str(v).strip())

            contaminated = set()
            if payload.get("manifest"):
                mdata = base64.b64decode(payload["manifest"]["data"]).decode("utf-8", "replace")
                mrows = [json.loads(l) for l in mdata.splitlines() if l.strip()]
                contaminated = detect_contaminated(mrows, name2acc)

            blob = build(staff, verified, skip, contaminated)
            name = f"证件核验报告{dt.datetime.now():%Y%m%d_%H%M}.xlsx"

            from collections import Counter
            ws = load_workbook(io.BytesIO(blob))["Sheet1"]
            c = Counter()
            for r in range(2, ws.max_row + 1):
                for k in str(ws.cell(row=r, column=9).value or "").split("、"):
                    c[k] += 1
            summary = "\n".join(f"  {k}：{v}" for k, v in c.most_common())

            body = json.dumps({"ok": True, "file": base64.b64encode(blob).decode(),
                               "name": name, "count": ws.max_row - 1, "summary": summary},
                              ensure_ascii=False).encode()
        except Exception as exc:
            body = json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"},
                              ensure_ascii=False).encode()

        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    url = f"http://127.0.0.1:{PORT}"
    print(f"证件核验报表工具已启动：{url}")
    print("关闭本窗口即退出。\n")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()


if __name__ == "__main__":
    try:
        main()
    except OSError as exc:
        print(f"启动失败：{exc}")
        print("端口 8765 可能被占用，或工具已在运行。")
        input("按回车退出…")
    except Exception:
        traceback.print_exc()
        input("按回车退出…")

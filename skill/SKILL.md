---
name: vendor-cert-audit
description: 批量核验人员证件（特种作业证、职业资格证等）的真伪与一致性。用于从业务系统抓取人员附件、解码证件二维码、调用官方平台核验、交叉比对后按客户模板成表的全流程。适用场景：代维/施工/外包人员证件审核、证书真伪核验、人员与证件是否一致的核查、证件有效期与复审到期排查、与人工查验结果交叉印证。
agent_created: true
---

# 代维人员证件核验流水线

## 用途

把「一大堆人的证件照片/PDF」变成「一张能拿去整改的问题清单」。核心不是 OCR 认字，而是拿到**第三方权威数据**做交叉比对——只有权威源能回答"证是真的吗""证是这个人的吗"。

## 何时使用

- 需要对几十到上千人的持证情况做合规排查
- 手上有业务系统导出的名单，系统里挂着证件附件
- 需要判断：证件真伪、证件与人是否一致、是否过期、复审是否逾期、证类是否匹配岗位
- 已有同事的人工查验结果，需要与自动化结果交叉印证

## 快速开始

首次使用先装依赖：

```bash
python -m pip install zxing-cpp pypdf pillow numpy openpyxl xlrd
```

最小可用路径（不接业务系统，直接处理已下载的证件文件）：

1. 准备 `manifest.jsonl`，每行至少含 `savedPath`（证件文件路径）、`pid`、`name`、`certLabel`
2. 解码二维码：`python scripts/qr_decode.py --manifest manifest.jsonl --out qr.jsonl`
3. 拿二维码里的编号+签名去权威平台核验（这一步要按目标平台单独逆向，见下）
4. 验身份证号：`python scripts/id_check.py --csv 名单.csv --col 身份证号码`
5. 出表：`python scripts/xlsx_fill.py --csv rows.csv --spec columns.json --out 结果.xlsx`

**动手前先读 `references/judgement-rules.md` 第七节**——数据串位的排查必须前置，否则整份结论不可用。

## 六段流水线

### ① 会话接管

业务系统通常有登录态、单点登录、iframe 嵌套。**不要重新登录，复用已有浏览器会话**。

- 通过 CDP（`connect_over_cdp`）接管一个已登录的 Chrome/Edge
- 页面在 iframe 里时，按 URL 特征 + 关键元素存在性双重判断目标 frame；**已分离（detached）的 frame 上调用 `wait_for_selector` 会永久阻塞**，必须先用 `is_detached()` 过滤
- 直接 GET 页面 URL 往往 404——只能从门户菜单点进去
- 同一长连接里多次导航后，frame 检测会失效并静默卡死。**重载页面必须放到新进程里做**
- **打开详情后必须校验面板上显示的是目标人**。详情面板多为异步局部刷新，上一次请求慢或失败时面板仍停在上一个人，此时读到的证件文件属于上一位——会把 A 的证算到 B 头上，误判成"借证"。校验方式：读面板里的姓名字段与目标姓名比对，不一致就重试，三次不对就跳过该人
- 自检信号：若某人的低压证/登高证/保险照片三个格子指向**同一个 fileId**，几乎必然是面板未刷新

### ② 列表全量抓取

- **翻页用「填页码 + 回车」，不要连点「下页」**。datascroller 状态跨会话残留，连点会在中途被误判为最后一页而禁用
- 判断"页真的翻了"用首行主键变化，不要读页码输入框
- 表头是双行合并时**不要按下标取列**，改用值的模式匹配（如身份证 `\d{6}\*+\d{3,4}`）
- 一页一进程：每页跑完就退出，状态坏了交给驱动重开，比在进程内自愈可靠

### ③ 附件批量下载

- 从网络请求里抓真实提交参数（表单字段常是 `xxxForm:xxxButton` 这种 JSF 风格）
- **用页面内 fetch 而不是表单提交**：服务器返回 HTML 错误页时，表单提交会把整个页面顶掉、JS 全失效；fetch 只是一个返回值
- 二进制用 base64 回传；文件名优先取 `Content-Disposition`（通常是干净的 UTF-8，比页面里的 `%uXXXX` 转义可靠）
- 同一文件被多人引用（如公司级保单）按 fileId 缓存，本地复制不再走网络
- 大批量时批内并发（一次 evaluate 里并发 8 条），比逐条快 5–8 倍

### ④ 内容解码

- **PDF 先读文字层**（`pypdf`），电子证书的文字层字段 100% 准确，比 OCR 强得多
- 扫描版 PDF 内部通常是 JPEG，按 `FFD8FF`…`FFD9` 抠出来直接解码，不需要 PDF 渲染库
- 二维码用 `zxing-cpp`（比 opencv 识别率高，有 cp313 轮子）
- **大图不要整图放大**：3737px 的证件图再放大只会爆内存且无效；改用切块扫描（3×3、4×4 重叠）
- 解码失败 ≠ 证件有问题。先量图片质量（尺寸/文件大小/边缘锐度），区分"图片真的糊"和"我读不出"

### ⑤ 权威核验

用凭据换权威数据，这是整条流水线唯一能回答"真伪"的环节。

- 常见凭据是「档案编号 + 签名」或「姓名 + 身份证号 + 证书编号」
- 小程序类平台：抓 XHR 看接口，通常就是普通 GET/POST，凭据在 query 里
- 政府平台常有 WAF：**不要破解验证码**。借用用户浏览器已通过验证的会话——复制一份浏览器配置目录（含 cookie、Local State、localStorage）带调试端口启动
- 跨浏览器搬 cookie 通常无效（WAF 绑定指纹与会话），必须用同一份配置目录
- 返回数据通常含：姓名、身份证号、作业类别、操作项目、发证机关、有效期、证件状态、官方证件照

### ⑥ 交叉比对与出表

判定口径见 `references/judgement-rules.md`，**务必先读**——错误的判定比漏判更伤。

## 脚本

| 脚本 | 用途 |
|---|---|
| `scripts/qr_decode.py` | 并行解码图片/PDF 里的二维码，失败自动切块重试 |
| `scripts/id_check.py` | 身份证号合法性校验（GB 11643） |
| `scripts/xlsx_fill.py` | 按给定列定义把数据写成 Excel，带分级配色与筛选 |
| `scripts/judge_report.py` | **判定与出表的一站式入口**：吃官方核验结果 + 名单，输出与人工资讯同构的 Excel。判定口径全部固化在代码里，不依赖模型临场判断 |

```bash
python scripts/qr_decode.py --manifest manifest.jsonl --out qr.jsonl --workers 4
python scripts/id_check.py 110101199002300012
python scripts/xlsx_fill.py --csv rows.csv --columns columns.json --out 结果.xlsx
```

**判定阶段一律用 `judge_report.py`，不要自己写判定逻辑。** 它的规则与 `references/judgement-rules.md` 一一对应，且已内置串位优先判定：

```bash
python scripts/judge_report.py \
    --verified official_verify.jsonl --staff allstaff.xls \
    --manual jining.xlsx linyi.xlsx \
    --contaminated contaminated_pids.json --pidmap omms_people.jsonl \
    --manifest manifest.jsonl --qr qr_scan.jsonl --done progress.json \
    --sheet2 allstaff.xls --out 查验结果.xlsx
```

## 交付物建议

三层结构最实用：

1. **问题清单**——直接拿去整改
2. **正常人员**——列出核验通过的证据（持有哪些证、有效期），便于复核"凭什么说正常"
3. **未核验人员**——按原因分类（图片模糊/未下载/无附件/手段受限），这是催办名单

若客户已有手工查验模板，**照抄其结构**（列名、行粒度、Sheet 布局、分类用词），不要自创格式。

## 硬性提醒

- 每条结论必须可追溯：说清依据哪个编号、比对了哪两个值
- 拿不到权威数据时，结论写"未核验"，不能写成"正常"或"有问题"
- 与人工核对结果冲突时，先归类再改规则：**叫法差异 / 我方判定错误 / 对方漏项 / 数据缺失**，不要一律迁就一方
- 一个文件写不进去大概率是用户正开着，改个文件名，别覆盖

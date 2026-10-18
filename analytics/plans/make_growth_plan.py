"""新客增长计划（Excel）生成脚本。用法：python3 -I analytics/plans/make_growth_plan.py reports/新客增长计划_草稿vN.xlsx
原稿黑字，原稿红字/黄底照搬，Claude 补充 = 蓝字，待决定 = 黄底。数字来自两店 2026-10 售出明细（只含汇总，不含客户信息）。"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

FONT = "微软雅黑"
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", fgColor="D9E1F2")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
SECTION = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="center")
BLACK, BLUE, RED = "000000", "1F5FBF", "C00000"


def _fill(template, rows, col):
    """'=C{套餐目标}*C{xx}' → '=D12*D7'（col = C 或 D）。"""
    import re
    return re.sub(r"C\{([^}]+)\}", lambda m: f"{col}{rows[m.group(1)]}", template)


def rt(*parts):
    """parts: (text, color) 或 纯文本(黑)。返回富文本。"""
    blocks = []
    for p in parts:
        text, color = (p, BLACK) if isinstance(p, str) else p
        blocks.append(TextBlock(InlineFont(rFont=FONT, sz=10, color=color), text))
    return CellRichText(blocks)


def B(text):  # Claude 补充
    return (text, BLUE)


def R(text):  # 原稿红字
    return (text, RED)


wb = Workbook()
ws = wb.active
ws.title = "新客计划"
gs = wb.create_sheet("目标倒推")
INPUT = PatternFill("solid", fgColor="FFF2CC")
gs["A1"] = "目标倒推：从月权益金目标算出每月需要多少新客"
gs["A1"].font = Font(name=FONT, size=14, bold=True)
gs["A2"] = rt(B("黄底格子是可以改的假设（默认用 10 月实际数据），其余都是公式。"
                "逻辑：权益金 × 套餐占比 = 套餐目标 → 扣掉老客续卡 → ÷ 新客成交客单 = 需要成交的新客 → ÷ 成交率 = 需要的新客数。"))
for j, h in enumerate(["项目", "说明", "1店 Morumbi", "2店 Ibirapuera"], 1):
    c = gs.cell(4, j, h)
    c.font, c.fill, c.border = Font(name=FONT, size=10, bold=True), HEAD, BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
G = {}  # 名称 -> 行号
goal_rows = [
    ("假设", None, None, None, None),
    ("月权益金目标", "你们定的月目标（11月和10月一样）", 140000, 260000, '"R$ "#,##0'),
    ("套餐占权益金比例", "10月实际", 0.647, 0.769, "0%"),
    ("老客续卡占套餐比例", "10月实际（1店只有 3 个成交样本，偏差大）", 0.52, 0.19, "0%"),
    ("新客成交客单", "10月新客买套餐的平均金额（1店只有 1 人，可改成你们的主推套餐价）", 2430, 3238, '"R$ "#,##0'),
    ("新客成交率目标", "你们定的目标", 0.20, 0.20, "0%"),
    ("线下新客占比目标", "建议值，可改", 0.30, 0.30, "0%"),
    ("每月营业日", "按周一~周六估算，节假日自行调整", 25, 25, "0"),
    ("倒推结果", None, None, None, None),
    ("套餐目标", "权益金 × 套餐占比", "=C{月权益金目标}*C{套餐占权益金比例}", None, '"R$ "#,##0'),
    ("其中靠新客", "套餐目标 × (1 − 老客续卡占比)", "=C{套餐目标}*(1-C{老客续卡占套餐比例})", None, '"R$ "#,##0'),
    ("其中靠老客续卡", "套餐目标 × 老客续卡占比", "=C{套餐目标}*C{老客续卡占套餐比例}", None, '"R$ "#,##0'),
    ("需要新客成交人数", "靠新客的金额 ÷ 新客成交客单", "=C{其中靠新客}/C{新客成交客单}", None, "0"),
    ("需要新客人数", "新客成交人数 ÷ 成交率", "=C{需要新客成交人数}/C{新客成交率目标}", None, "0"),
    ("其中线上", "新客 × (1 − 线下占比)", "=C{需要新客人数}*(1-C{线下新客占比目标})", None, "0"),
    ("其中线下", "新客 × 线下占比", "=C{需要新客人数}*C{线下新客占比目标}", None, "0"),
    ("每营业日需要新客", "新客 ÷ 营业日", "=C{需要新客人数}/C{每月营业日}", None, "0.0"),
    ("10月现状", None, None, None, None),
    ("10月新客（人）", "1店 10/1–10/17，2店 10/1–10/18", 15, 179, "0"),
    ("10月营业日", "同一区间去掉周日", 15, 15, "0"),
    ("现在每营业日新客", "", "=C{10月新客（人）}/C{10月营业日}", None, "0.0"),
    ("现在新客成交率", "", 0.067, 0.084, "0.0%"),
    ("现在线下占比", "", 0.667, 0.168, "0%"),
    ("差距：每日新客 需要 ÷ 现在", "> 1 表示新客量不够", "=C{每营业日需要新客}/C{现在每营业日新客}", None, "0.0"),
    ("差距：成交率 目标 ÷ 现在", "> 1 表示转化不够", "=C{新客成交率目标}/C{现在新客成交率}", None, "0.0"),
]
r0 = 5
for i, (name, _, *_rest) in enumerate(goal_rows):
    G[name] = r0 + i
for i, (name, note, v1, v2, fmt) in enumerate(goal_rows):
    row = r0 + i
    if note is None:
        c = gs.cell(row, 1, name)
        c.font, c.fill = Font(name=FONT, size=10, bold=True), SECTION
        for j in range(1, 5):
            gs.cell(row, j).fill = SECTION
        continue
    gs.cell(row, 1, name).font = Font(name=FONT, size=10, bold=True)
    gs.cell(row, 2, note).font = Font(name=FONT, size=9, color="7F7F7F")
    for col, v in ((3, v1), (4, v2 if v2 is not None else v1)):
        L = "C" if col == 3 else "D"
        if isinstance(v, str):
            v = _fill(v, G, L)
        c = gs.cell(row, col, v)
        c.number_format, c.border = fmt, BORDER
        is_input = not isinstance(v, str)
        c.font = Font(name=FONT, size=10, color=BLUE if is_input else BLACK, bold=not is_input)
        if is_input and name not in ("10月新客（人）", "10月营业日", "现在新客成交率", "现在线下占比"):
            c.fill = INPUT
    for j in (1, 2):
        gs.cell(row, j).border = BORDER
last_goal = r0 + len(goal_rows)
gs.cell(last_goal + 1, 1, "读法").font = Font(name=FONT, size=10, bold=True)
gs.cell(last_goal + 2, 1, rt(B("1店 Morumbi：每天新客只有 1 人，要到约 3.6 人（差距约 3.6 倍）；另外一半套餐金额要靠老客续卡。")))
gs.cell(last_goal + 3, 1, rt(B("2店 Ibirapuera：现在每天新客已经够（差距 < 1），关键是成交率 8% → 20%（差距约 2.4 倍）。")))
for j, w in enumerate([24, 46, 16, 16], 1):
    gs.column_dimensions[get_column_letter(j)].width = w


def goal_formula(kind):
    """计划表「目标」格：引用目标倒推页，假设改了会跟着变。"""
    row = G["其中线上"] if kind == "线上" else G["其中线下"]
    share = f"TEXT('目标倒推'!C{G['线下新客占比目标']},\"0%\")"
    head = f'"线下新客占比 "&{share}&CHAR(10)' if kind == "线下" else '""'
    return (f"={head}&\"每月{kind}新客：\"&CHAR(10)&\"1店 \"&TEXT('目标倒推'!C{row},\"0\")&\" 人\"&CHAR(10)"
            f"&\"2店 \"&TEXT('目标倒推'!D{row},\"0\")&\" 人\"&CHAR(10)&CHAR(10)&\"（从11月起，按「目标倒推」）\"")
cols = [("板块", 7), ("当前结果", 30), ("目标\n（什么时候做到什么）", 22), ("问题思考维度参考", 20),
        ("核心问题", 34), ("关键动作", 40), ("衡量指标", 34), ("时间节点", 11), ("责任人", 16)]
for j, (h, w) in enumerate(cols, 1):
    ws.column_dimensions[get_column_letter(j)].width = w

ws["A1"] = "新客增长计划（含10/29沙龙、11月活动）— 草稿 v5"
ws["A1"].font = Font(name=FONT, size=15, bold=True)
ws["A2"] = rt("黑字 = 原稿；", R("红字 = 原稿重点；"), B("蓝字 = Claude 补充的数据和建议（待确认）；"),
              "黄底 = 需要拍板 / 原稿标黄。1店 = Morumbi（街边店），2店 = Ibirapuera（Shopping Ibirapuera 商场店）。"
              "目标列由「目标倒推」页自动算出，改那页的假设这里会跟着变。")
ws["A3"] = rt(B("数据来源：两店 10 月售出明细的「进店渠道」（1店 10/1–10/17，2店 10/1–10/18）。"
                "新客 = 做项目且不是套餐/免单核销的客户（与周报口径一致），按人去重。"))

r = 5
for j, (h, _) in enumerate(cols, 1):
    c = ws.cell(r, j, h)
    c.font, c.fill, c.border = Font(name=FONT, size=10, bold=True), HEAD, BORDER
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[r].height = 32

def put(row, col, value, fill=None):
    c = ws.cell(row, col, value)
    if isinstance(value, str):
        c.font = Font(name=FONT, size=10)
    c.alignment, c.border = WRAP, BORDER
    if fill:
        c.fill = fill
    return c

def block(first, label, current, goal, rows, owner_merge=None):
    """rows: [(维度, 核心问题, 关键动作, 衡量指标, 时间, 责任人, {col: fill})]"""
    last = first + len(rows) - 1
    for i, row in enumerate(rows):
        rr = first + i
        fills = row[6] if len(row) > 6 else {}
        for j, v in enumerate(row[:6], 4):
            put(rr, j, v, fills.get(j))
        for j in (1, 2, 3):
            ws.cell(rr, j).border = BORDER
    for j, v in [(1, label), (2, current), (3, goal)]:
        put(first, j, v)
        ws.merge_cells(start_row=first, start_column=j, end_row=last, end_column=j)
    ws.cell(first, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return last + 1

# ---------------- 线上
online_current = rt(B("10月实际（按进店渠道，新客按人去重）：\n"
                      "1店 Morumbi（10/1–10/17）线上新客 5 人\n· Instagram 4 · 官网 Site 1\n\n"
                      "2店 Ibirapuera（10/1–10/18）线上新客 149 人\n· 投流 Anúncio online 90\n· Instagram 58\n· TikTok 1"))
online_goal = goal_formula("线上")
online_rows = [
    ("内容（KOL内容、自制内容）",
     rt(B("哪类内容真能带来到店？10月 UGC/KOL 渠道新客 0 人")),
     rt(B("固定每周自制内容条数 + 每月 KOL 数；KOL 带来的客人收银时「进店渠道」选 UGC\n"
          "11月内容主线：Morumbi 一周年、Reset Capilar 上新、Black Friday")),
     rt(B("UGC/KOL 渠道新客数（从收银明细统计）")), None, "市场部"),
    ("投流",
     "投流回报比？",
     rt(B("每周把各渠道投放金额发给我，周报里直接算获客成本和回报比\n"
          "11月两条投放：Morumbi 一周年 + Reset Capilar 上新；Black Friday（见「11月活动决策」）")),
     rt(B("投流新客数；获客成本 = 投放 ÷ 新客；回报比 = 投流客户成交金额 ÷ 投放。"
                                     "10月2店投流新客 90 人 → 成交 10 人（11%），套餐 R$ 41,730")),
     None, "市场部（含每周投放金额）"),
    ("私信转化",
     "1、客户反馈：想约，回复慢，约不到；\n2、客户问了一句后就没消息了；\n3、数据如何衡量追踪",
     rt(R("对信息回复的前台给奖励，不从信息上做标记，而是从到店客户反查是哪个渠道"),
        B("\n补充：①定回复时效（例如 15 分钟内）和预约话术；②收银时「进店渠道」必填——"
          "这就是「到店反查」，可以直接从收银明细统计")),
     rt(B("Instagram 渠道新客数、新客→成交率（10月2店 58 人→成交 4 人，7%）；渠道漏填条数")), None, "市场部（回复）\n前台奖励：ISA"),
    ("社媒人员\n（会不会、愿不愿、让不让）", None,
     rt(B("明确社媒负责人和每周 KPI（内容条数、回复时效、私信→预约数）")), None, None, "市场部"),
    (None,
     "2店商场店对1店街边店的影响？\n（物理距离5公里）",
     rt(B("见待办 2")),
     rt(B("10月：1店新客只有 15 人（每营业日 1 人）；两店重叠客户 3 人；1店套餐客户去2店核销 8 次（6 人），"
          "2店客户来1店核销 2 次。要判断影响，需要2店开业前后1店的明细")), None, None, {5: YELLOW}),
]
r = block(r + 1, "线上", online_current, online_goal, online_rows)

# ---------------- 线下
offline_current = rt(B("10月实际：\n1店 Morumbi 线下新客 10 人（占 67%）\n· 未预约到店 7 · 老带新 3\n\n"
                       "2店 Ibirapuera 线下新客 30 人（占 17%）\n· 老带新 21 · 商场 Shopping 9"))
offline_goal = goal_formula("线下")
offline_rows = [
    ("人员意愿", "1、1店护理师不派单，招聘？",
     rt("1、职责：目前已解决，招聘时说清楚；\n", R("2、拓客激励方案"),
        "\n3、下周出 1店护理师内部激励方案：Reset Capilar 11月上线后，让 1店护理师走出店门拓新客"),
     None, "方案：下周（10/19–10/25）\n执行：11月", "营运 Isa"),
    ("拓客人员", "2、招聘节点，培训时间不够，没时间做拓客", "组建拓客组", None, None, None),
    ("拓客方式", "3、有过专员做线下活动，但是效果很差，且带来的客人转换率非常低", None,
     rt(B("每次出店拓客 / 活动单独登记渠道，算到店人数和成交率，效果差的停掉")), None, None),
    ("地推", "周边容易到达的用户没有积淀，“就近原则”没有发挥出来；",
     rt("营运和总部聊线下动作，确定方案\n1、卡券设计\n2、意愿：①激励方案非常重要；②目标设置逐级递增，先设小目标\n"
        "3、能力（话术、如何拓）\n4、地点、频率", B("\n建议：地推卡券就用 Reset Capilar R$199（11月送抹茶饮品）")),
     rt(B("地推渠道：发券数 → 到店数 → 成交数")), "11月", "Isa负责，和总部运营团队定当下，宜昌扣细节",
     {6: YELLOW}),
    ("异业合作", None, rt("进军巴西 Gympass（Wellhub）", B("\n建议：收银「进店渠道」加一个 Gympass 选项，才能看到效果")),
     rt(B("Gympass 渠道新客数")), "11月", "市场部"),
    ("老带新", rt(B("10月老带新：1店 3 人，2店 21 人（2店线下最大来源）；目前没有激励")),
     "11月起，只在 1店 Morumbi：带 1 位客人来，送 1 次快速洗头（Reset Capilar）",
     rt(B("老带新渠道新客数（Indicado por um amigo）")), "11月（暂定只做11月）", None),
    ("到店赠送", None, "11月：两店（1店 Morumbi、2店 Ibirapuera）所有客人，新客老客都一样，到店就送快速肩颈", None, "11月（暂定只做11月）", None),
    ("商场广告置换", rt(B("2店在 Shopping Ibirapuera 商场里：10月「商场」渠道新客 9 人，成交 0")), None, None, None, None),
    ("大客户", None, "写字楼午休快洗：市场部做方案，由市场部人员在 1店 Morumbi 周边写字楼谈判",
     rt(B("签约公司数 → 员工到店数")), "11月", "市场部"),
    ("VIP 沙龙\n（1店 Morumbi）",
     rt(B("Morumbi 新客少、会员多：上午场给卡快用完的会员续卡，下午场给体验过没买的客人首次开卡")),
     rt("按《VIP 沙龙方案 v3》：报名费 R$119；上午场续卡（正常套餐 9 折，满 R$5,000 送居家精粹液）；"
        "下午场首次开卡（R$1,680 / 4 次）；每场 6 人，共发 60 条邀约；护理师到场 R$50 + 成交 R$100（待财务审批）",
        B("\n注意：①离活动只剩约 10 天，方案里的 T-20 邀约期要压缩，本周内发完两波邀约；"
          "②10/29 是周四（方案建议周一 / 周三）；③下午场 10月只有 13 位体验未成交的客人，要把 9 月及以前的也拉进名单")),
     "邀约 60 → 到场 12 → 成交 6；当天业绩", "10/29（周四）", "Morumbi 店长", {6: YELLOW}),
    ("……", None, None, None, None, None),
]
r = block(r, "线下", offline_current, offline_goal, offline_rows)

# ---------------- 其他关键动作（原稿表格下方）
others = ["给员工做客户资源分配",
          "市场部和营运部门的协作关系：",
          "门店：核心是客户为中心，客户数据分析，客户分层，不同渠道客户数据；",
          "社媒：GEO优化，信息“植入”权威网站做运营；植入哪些信息，按“客户生命周期”管理"]
first = r
for i, text in enumerate(others):
    for j in range(1, 10):
        ws.cell(r + i, j).border = BORDER
    put(r + i, 6, text)
put(first, 1, "其他")
ws.merge_cells(start_row=first, start_column=1, end_row=r + len(others) - 1, end_column=1)
ws.cell(first, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
r += len(others) + 2

# ---------------- 待办
ws.cell(r, 1, "三、待办").font = Font(name=FONT, size=13, bold=True)
r += 1
todo_cols = ["序号", "板块", "事项", "负责人", "时间节点"]
spans = [(1, 1), (2, 3), (4, 6), (7, 7), (8, 9)]
for (a, b), h in zip(spans, todo_cols):
    c = ws.cell(r, a, h)
    c.font, c.fill = Font(name=FONT, size=10, bold=True), HEAD
    c.alignment = Alignment(horizontal="center", vertical="center")
    for j in range(a, b + 1):
        ws.cell(r, j).border = BORDER
    if b > a:
        ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
todos = [
    ("1", "线下拓客方案", "和总部营运聊线下方案（含那4点，重点做激励方案）", "营运Isa", None),
    ("2", None, "从数据上分析和看2店开业后对1店的影响", "ISA/廖磊", None),
    ("3", None, "（1）线上：对信息回复的前台给奖励，不从信息上做标记，而是从到店客户反查是哪个渠道", "ISA", None),
    ("4", "线下拓客", "1店 Morumbi 护理师内部激励方案（Reset Capilar 上线后出店拓新客）", "营运 Isa", "下周（10/19–10/25）"),
    ("5", "VIP 沙龙", "1店 Morumbi 沙龙：本周内发完两波邀约（60 条），10/28 完成逐个客户分析 + 彩排，10/29 执行，11/5 前复盘",
     "Morumbi 店长", "10/29"),
    ("6", "11月活动", "回复市场部 11月方案（Reset Capilar、体验券、促销套餐、阶梯券；葡语回复草稿已备好）", "Isa", "尽快"),
    ("7", "大客户", "写字楼午休快洗方案 + 1店 Morumbi 周边写字楼谈判", "市场部", "11月"),
    ("8", "异业合作", "入驻 Gympass（Wellhub）", "市场部", "11月"),
    (rt(B("9")), rt(B("数据")), rt(B("提供2店开业前后1店的售出明细（例如开业前3个月 + 开业后），用于待办 2")), None, None),
    (rt(B("10")), rt(B("投流")), rt(B("每周提供各渠道投放金额，用于算获客成本和回报比")), "市场部", "每周"),
]
for t in todos:
    r += 1
    for (a, b), v in zip(spans, t):
        put(r, a, v)
        for j in range(a, b + 1):
            ws.cell(r, j).border = BORDER
        if b > a:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
    ws.row_dimensions[r].height = 30

# ---------------- 待确认
r += 2
ws.cell(r, 1, "四、确认记录").font = Font(name=FONT, size=13, bold=True)
questions = [
    "以上已于 10/18–10/19 和 Isa 逐项确认：11月目标同10月；老带新只在 1店 Morumbi；CRM 分三类人群发；"
    "Gympass、投流（含每周投放金额）、社媒由市场部负责；护理师拓客激励方案由营运 Isa 负责。新的想法随时补充。",
]
for q in questions:
    r += 1
    ws.cell(r, 1, q).font = Font(name=FONT, size=10)
    ws.cell(r, 1).fill = YELLOW
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)

ws.freeze_panes = "A6"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- 11月活动决策
cs = wb.create_sheet("11月活动决策", 1)
cs["A1"] = "11月活动（市场部方案 + Isa 补充）— 决定"
cs["A1"].font = Font(name=FONT, size=14, bold=True)
cs["A2"] = rt(B("「数据参考」来自两店 10 月明细；「建议」是 Claude 的看法；最后一列是 Isa 已定的（黄底空着的还没定）。"))
heads = [("#", 4), ("事项", 16), ("市场部的提议 / 问题", 36), ("数据参考", 40), ("建议", 46), ("Isa 决定", 26)]
for j, (h, w) in enumerate(heads, 1):
    c = cs.cell(4, j, h)
    c.font, c.fill, c.border = Font(name=FONT, size=10, bold=True), HEAD, BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
    cs.column_dimensions[get_column_letter(j)].width = w
decisions = [
    ("Express 流程", "新项目 Reset capilar / Lavagem terapêutica；流程里只有 5 分钟快速评估，鼓励客人再做完整评估",
     "2店新客成交率 8.4%（目标 20%）；成交基本发生在体验后的方案推荐",
     "保留 5 分钟评估，但结束时帮客人约好「完整评估」；收银系统里给 Reset Capilar 单独建项目名，才能追踪它带来的成交",
     "5 分钟快速评估，鼓励客人另约完整评估"),
    ("Express 名称", "Reset capilar 或 Lavagem terapêutica", "—", "二选一；收银系统、官网、广告用同一个名字", "Reset Capilar"),
    ("Express 正式价", "待定（A definir）",
     "现在新客入门价：体验券 R$188 / 288 / 388；2店新客实付平均 R$194（标准价 360 的项目）/ R$281（标准价 530 的项目）",
     "正式价放在入门券和 R$360 标准项目之间，接替最便宜的体验券做低价入口", "R$199"),
    ("Express 11月优惠", "打折，或者送颈部按摩 / 刮痧（guasha）", "—", "倾向送赠品，不打折，保住正式价",
     "不打折，送抹茶饮品，主打「快速恢复元气」"),
    ("官网体验券", "下架 R$188/288/388 折扣券，恢复原价；改用阶梯券：满 300 减 100，500–1000 减 200，1000 以上减 300",
     "恢复原价后入门价从 188→260、288→430（+38% / +49%）；2店新客主要做的就是这些项目（Desmaquilhagem 82 次、Mio Relax 44 次）",
     "先有 Reset Capilar 做低价入口再下架；阶梯有断层：499 只减 100（20%），500 减 200（40%），999 减 200（20%），1000 减 300（30%）",
     "Reset Capilar 上线即取消 R$188 券；R$288/388 先保留，新客量稳住再下架；阶梯券不能和促销套餐叠加"),
    ("促销套餐 Antiqueda", "原价 1950 → 1280（Detox、Scalp Care、Antiqueda ×2）",
     "现在同类「1680 防脱」卖 R$1,680；2店10月成交客单 R$3,338（多数人一次买 2 个套餐）",
     "1280 比现在卖价低 24%，客单会下降：如果多数成交走促销套餐，2店每月要成交约 127 个新客才够 26 万"
     "（「目标倒推」页把新客成交客单改成 1280 就能看到）",
     "1950 → 1280，不限制，照方案做"),
    ("促销套餐 Primavera/Verão", "Spa essencial、Antioleosidade、Nutrição essencial ×2；价格未定", "—", "需要定价",
     "和 Antiqueda 一样：1950 → 1280"),
    ("两店通用", "Disponível nas duas unidades", "1店 Morumbi 10月新客只有 15 人，客流是最大问题", "活动两店都上", "两店都上（照方案）"),
    ("CRM", "WhatsApp 广播、邮件、分层推促销套餐",
     "2店10月体验后没买套餐的新客约 164 人；1店 80% 的客次是会员",
     "分三层推：体验未成交 → 促销套餐；会员 → 老带新送快速洗头；沉睡客户 → Reset Capilar 回店。名单可以从收银明细导出",
     "分三类人发：体验过没买 → 促销套餐 R$1,280；1店会员 → 老带新送快速洗头；很久没来 → Reset Capilar R$199 回店"),
    ("追踪", "—", "—", "Reset Capilar、促销套餐、阶梯券、Gympass 在收银系统里单独命名 / 记录，进店渠道照填", None),
    ("到店赠送", "（Isa 补充）", "—", "对外宣传时和 Reset Capilar、Black Friday 一起讲", "11月两店所有客人（新老）到店就送快速肩颈"),
    ("老带新", "（Isa 补充）", "10月老带新新客：1店 3 人、2店 21 人", "收银时进店渠道选 Indicado por um amigo，才能统计效果",
     "11月起，只在 1店 Morumbi：带 1 位客人来，送 1 次快速洗头（Reset Capilar）"),
    ("写字楼午休快洗", "（Isa 补充）", "—", "Reset Capilar 适合午休时段；可以做企业团购价 / 员工福利",
     "市场部做方案，由市场部人员在 1店 Morumbi 周边写字楼谈判（11月）"),
    ("Gympass", "（Isa 补充）", "—", "收银「进店渠道」加 Gympass 选项", "11月入驻 Gympass（Wellhub），市场部负责"),
]
for i, (topic, ask, data_ref, advice, decided) in enumerate(decisions, 1):
    row = 4 + i
    values = [i, topic, ask, rt(B(data_ref)) if data_ref != "—" else "—", rt(B(advice)), decided]
    for j, v in enumerate(values, 1):
        c = cs.cell(row, j, v)
        if isinstance(v, (str, int)):
            c.font = Font(name=FONT, size=10, bold=j == 2)
        c.alignment, c.border = WRAP, BORDER
    if decided is None:
        cs.cell(row, 6).fill = YELLOW
    else:
        cs.cell(row, 6).font = Font(name=FONT, size=10, bold=True)
cs.freeze_panes = "C5"
cs.page_setup.orientation = "landscape"
cs.page_setup.fitToWidth, cs.page_setup.fitToHeight = 1, 0
cs.sheet_properties.pageSetUpPr.fitToPage = True

# ---------------- 渠道数据
ds = wb.create_sheet("渠道数据")
ds["A1"] = "10月新客渠道（来自售出明细「进店渠道」）"
ds["A1"].font = Font(name=FONT, size=13, bold=True)
data = [
    ("1店 Morumbi（10/1–10/17）", [
        ("Sem hora marcada（未预约到店）", "线下", 7, 7, 0),
        ("Instagram", "线上", 4, 6, 0),
        ("Indicado por um amigo（老带新）", "线下", 3, 4, 1),
        ("Site（官网）", "线上", 1, 2, 0)]),
    ("2店 Ibirapuera（10/1–10/18）", [
        ("Anúncio online（投流）", "线上", 90, 109, 10),
        ("Instagram", "线上", 58, 67, 4),
        ("TikTok", "线上", 1, 1, 0),
        ("Indicado por um amigo（老带新）", "线下", 21, 24, 1),
        ("Shopping Ibirapuera（商场）", "线下", 9, 11, 0)]),
]
r = 3
for title, rows in data:
    ds.cell(r, 1, title).font = Font(name=FONT, size=11, bold=True)
    r += 1
    for j, h in enumerate(["进店渠道", "线上/线下（建议）", "新客人数", "新客客次", "其中买套餐", "新客→成交率"], 1):
        c = ds.cell(r, j, h)
        c.font, c.fill, c.border = Font(name=FONT, size=10, bold=True), HEAD, BORDER
    first = r + 1
    for row in rows:
        r += 1
        for j, v in enumerate(row, 1):
            c = ds.cell(r, j, v)
            c.font, c.border = Font(name=FONT, size=10), BORDER
        c = ds.cell(r, 6, f"=IF(C{r}=0,\"\",E{r}/C{r})")
        c.number_format, c.font, c.border = "0%", Font(name=FONT, size=10), BORDER
    r += 1
    ds.cell(r, 1, "合计").font = Font(name=FONT, size=10, bold=True)
    for j in (3, 4, 5):
        L = get_column_letter(j)
        c = ds.cell(r, j, f"=SUM({L}{first}:{L}{r - 1})")
        c.font, c.border = Font(name=FONT, size=10, bold=True), BORDER
    c = ds.cell(r, 6, f"=IF(C{r}=0,\"\",E{r}/C{r})")
    c.number_format, c.font, c.border = "0%", Font(name=FONT, size=10, bold=True), BORDER
    r += 1
    ds.cell(r, 1, "线下占比").font = Font(name=FONT, size=10, bold=True)
    c = ds.cell(r, 3, f'=SUMIF(B{first}:B{r - 2},"线下",C{first}:C{r - 2})/C{r - 1}')
    c.number_format, c.font = "0%", Font(name=FONT, size=10, bold=True)
    r += 3
ds.cell(r, 1, "注：新客 = 做项目且不是套餐/免单核销的客户，按人去重；「其中买套餐」= 这些新客里本月买了套餐的人数。"
              "两店重叠客户 3 人；1店套餐客户在2店核销 8 次（6 人），2店客户在1店核销 2 次。").font = \
    Font(name=FONT, size=9, color="7F7F7F")
for j, w in enumerate([34, 18, 12, 12, 12, 14], 1):
    ds.column_dimensions[get_column_letter(j)].width = w

wb.save(sys.argv[1])
print("saved", sys.argv[1])

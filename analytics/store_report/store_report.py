#!/usr/bin/env python3
"""门店经营报表：读取收银系统导出的「售出明细表」，生成 看板 / 月度总数据 / 周度数据 / 员工达成 / 项目分类 报表。

用法:
    python3 store_report.py 售出明细.xlsx                      # 输出到当前目录
    python3 store_report.py 售出明细.xlsx -o 报表.xlsx
    python3 store_report.py 售出明细.xlsx --month 2026-10      # 文件含多个月时只取某月
    python3 store_report.py 售出明细.xlsx --weeks natural      # 周按 周一~周日 划分（默认 1-7日=W1）
    python3 store_report.py 售出明细.xlsx --targets 目标.json   # 指定目标文件（默认自动用 targets/<YYYY-MM>.json）

指标口径见 README.md，与 2026-10 手工透视表结果逐项核对一致。
"""

import argparse
import calendar
import json
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["类型", "分类", "名称", "数量", "支付方式", "支付金额", "现金类支付", "员工", "客户名称", "账务日期"]

# 支付方式归类。新出现的支付方式默认按「现金类」处理，并在「数据检查」中提示。
KNOWN_PAY_METHODS = {
    "代金券", "开业券 Voucher comprado online", "套餐", "免单类",
    "现金", "信用卡 Crédito à vista", "信用卡分期 Crédito parcelado", "银行卡 Débito", "PIX",
}

MONTH_METRICS = ["权益金", "营业额", "总客次", "会员客次", "新客客次", "网红/邀约",
                 "成交人数", "成交客单", "二次券成交数", "体验客单价", "叠加率", "新客成交率"]
HELPER_METRICS = ["套餐成交金额", "数量合计"]
STAFF_METRICS = ["业绩", "营业额", "总客次", "成交人数", "二次券成交人数", "新客成交率"]
STAFF_HELPERS = ["新客客次"]
CUSTOMER_TYPES = ["新客客次", "会员客次", "网红/邀约"]
UNASSIGNED = "未指定员工（线上购券）"
TARGETS_DIR = Path(__file__).resolve().parent / "targets"
RATE_METRICS = {"新客成交率", "体验客单价", "成交客单", "叠加率"}  # 比率型目标：任何时间段的目标都一样
METRIC_ALIASES = {"客单价": "体验客单价"}
WEEKDAYS = "一二三四五六日"
UNCATEGORIZED = "（未分类）"

METRIC_FORMATS = {
    "权益金": "#,##0.00", "营业额": "#,##0.00", "成交客单": "#,##0", "体验客单价": "#,##0.00",
    "叠加率": "0.00", "套餐成交金额": "#,##0.00", "业绩": "#,##0.00", "平均单价": "#,##0.00",
    "客次占比": "0.0%", "营业额占比": "0.0%", "新客成交率": "0.00%", "月达成率": "0.0%", "累计达成率": "0.0%",
}

DEFINITIONS = [
    ("权益金", "类型=权益 的「支付金额」合计，排除支付方式=开业券(线上已付)。即门店当期卖出套餐+代金券收到的钱，退款为负数自动抵扣。"),
    ("营业额", "类型=项目 的「支付金额」合计（含代金券/套餐/免单核销的项目价值）。"),
    ("总客次", "类型=项目 的行数（每做一个项目算 1 客次；退款单扣减）。"),
    ("会员客次", "类型=项目 且 支付方式=套餐 的行数。"),
    ("新客客次", "总客次 − 会员客次 − 网红/邀约（表内公式）。"),
    ("网红/邀约", "类型=项目 且 支付方式=免单类 的行数。"),
    ("成交人数", "类型=权益、分类=套餐，现金类支付合计>0 的客户人数（按客户去重）。"),
    ("成交客单", "套餐成交金额 ÷ 成交人数（表内公式）。套餐成交金额 = 类型=权益、分类=套餐 的「现金类支付」合计。"),
    ("二次券成交数", "类型=权益、分类=代金券，且支付方式不是 代金券/开业券/套餐/免单类（即客人现场掏钱买券）的「数量」合计，单位：张。"),
    ("体验客单价", "营业额 ÷ 总客次（表内公式）。"),
    ("叠加率", "数量合计 ÷ 总客次（表内公式）。数量合计 = 明细表全部行（项目+权益）的「数量」合计。"),
    ("新客成交率", "成交人数 ÷ 新客客次 × 100%（表内公式）。新客客次为 0 时留空。"),
    ("员工·业绩", "该员工名下的权益金（卖套餐+卖券收到的钱，口径同「权益金」）。明细里「实业绩」列全为 0，故不用。"),
    ("员工·营业额/总客次", "该员工名下的项目营业额 / 项目行数，口径同上。"),
    ("员工·成交人数", "该员工名下买套餐的客户人数（去重）。"),
    ("员工·二次券成交人数", "该员工名下现场掏钱买代金券的客户人数（去重）。"),
    ("员工·新客成交率", "该员工名下成交人数 ÷ 该员工服务的新客客次（灰色「计算用:新客客次」列）。"
                     "只卖不做项目的员工新客客次为 0，成交率留空。"),
    ("项目分类", "类型=项目 的行按收银系统「分类」汇总；新客/会员/网红 的划分与上面客次口径相同；"
                "客户人数=做过该分类项目的去重客户数；平均单价=营业额÷总客次。"),
    ("月达成率", "实际 ÷ 月目标。目标来自 targets/<YYYY-MM>.json（用户每月提供）。"),
    ("目标进度", "按「每日目标」累计到报表最后一天占全月的比例（休息日权重为 0），比单纯按天数算更准。"),
    ("同期目标 / 同期达成率", "累加型指标（权益金、营业额、总客次等）：同期目标 = 月目标 × 目标进度；"
                       "比率型指标（新客成交率、体验客单价）：同期目标 = 月目标。同期达成率 ≥ 100% 记 ✓，否则 ✗。"),
    ("周目标 / 日目标", "累加型指标按「每日目标」的比例拆到每天，再按周相加；比率型指标每周、每天的目标与月目标相同。"),
    ("周度划分", "默认：每月 1-7日=W1、8-14日=W2、15-21日=W3、22-28日=W4、29日-月底=W5（按「账务日期」）。"
                "可用 --weeks natural 改成 周一~周日 的自然周。"),
]


# ---------------------------------------------------------------- 读取与清洗

def find_sheet(path):
    """返回 (sheet 名, 表头行号)。优先找名为「售出明细表」的 sheet，否则找第一个含必需列的 sheet。"""
    xls = pd.ExcelFile(path)
    names = sorted(xls.sheet_names, key=lambda n: n != "售出明细表")
    for name in names:
        head = pd.read_excel(xls, sheet_name=name, header=None, nrows=15, dtype=str)
        for i, row in head.iterrows():
            cells = {str(v).strip() for v in row.values if pd.notna(v)}
            if {"类型", "支付方式", "支付金额"} <= cells:
                return name, i
    raise SystemExit(f"在 {path} 中没找到「售出明细表」（需要包含列：{'、'.join(REQUIRED_COLUMNS)}）")


def clean_text(v):
    if pd.isna(v):
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def pay_kind(method):
    m = method or ""
    if "开业券" in m:
        return "开业券"
    if "免单" in m:
        return "免单"
    if m.startswith("代金券"):
        return "代金券"
    if m.startswith("套餐"):
        return "套餐"
    return "现金类"


def load(path, month=None):
    sheet, header_row = find_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet, header=header_row,
                       dtype={"单据编号": str, "客户编号": str, "编号": str})
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise SystemExit(f"「{sheet}」缺少列：{'、'.join(missing)}")

    df = df[df["类型"].notna()].copy()
    for col in ["数量", "支付金额", "现金类支付"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    for col in ["类型", "分类", "名称", "支付方式", "员工", "客户名称"]:
        df[col] = df[col].map(clean_text)
    df["日期"] = pd.to_datetime(df["账务日期"], errors="coerce").dt.normalize()
    if df["日期"].isna().any() and "操作时间" in df.columns:
        df["日期"] = df["日期"].fillna(pd.to_datetime(df["操作时间"], errors="coerce").dt.normalize())
    df["支付归类"] = df["支付方式"].map(pay_kind)
    is_refund = df["单据类型"].astype(str).str.contains("退款") if "单据类型" in df.columns else False
    df["客次符号"] = 1
    df.loc[is_refund, "客次符号"] = -1
    cust_id = df["客户编号"].map(clean_text) if "客户编号" in df.columns else None
    df["客户"] = cust_id.fillna(df["客户名称"]) if cust_id is not None else df["客户名称"]

    months = sorted(df["日期"].dropna().dt.to_period("M").astype(str).unique())
    if month:
        df = df[df["日期"].dt.to_period("M").astype(str) == month]
        if df.empty:
            raise SystemExit(f"文件里没有 {month} 的数据（有：{', '.join(months)}）")
    return df, sheet, months


# ---------------------------------------------------------------- 指标计算

def paying_customers(rows, value_col):
    """按客户去重，净额 > 0 的客户数。"""
    if rows.empty:
        return 0
    net = rows.groupby("客户")[value_col].sum()
    return int((net > 0).sum())


def compute(df):
    proj = df[df["类型"] == "项目"]
    right = df[df["类型"] == "权益"]
    package = right[right["分类"] == "套餐"]
    voucher_sale = right[(right["分类"] == "代金券") & (right["支付归类"] == "现金类")]
    m = {
        "权益金": right.loc[right["支付归类"] != "开业券", "支付金额"].sum(),
        "营业额": proj["支付金额"].sum(),
        "总客次": int(proj["客次符号"].sum()),
        "会员客次": int(proj.loc[proj["支付归类"] == "套餐", "客次符号"].sum()),
        "网红/邀约": int(proj.loc[proj["支付归类"] == "免单", "客次符号"].sum()),
        "成交人数": paying_customers(package, "现金类支付"),
        "二次券成交数": int(voucher_sale["数量"].sum()),
        "二次券成交人数": paying_customers(voucher_sale, "数量"),
        "套餐成交金额": package["现金类支付"].sum(),
        "数量合计": df["数量"].sum(),
    }
    m["新客客次"] = m["总客次"] - m["会员客次"] - m["网红/邀约"]
    m["成交客单"] = m["套餐成交金额"] / m["成交人数"] if m["成交人数"] else 0
    m["体验客单价"] = m["营业额"] / m["总客次"] if m["总客次"] else 0
    m["叠加率"] = m["数量合计"] / m["总客次"] if m["总客次"] else 0
    m["新客成交率"] = m["成交人数"] / m["新客客次"] if m["新客客次"] else None
    m["业绩"] = m["权益金"]
    return m


def mix_table(proj, keys):
    """项目行按 keys 汇总：总客次 / 新客 / 会员 / 网红 / 客户人数 / 营业额，按总客次降序。"""
    cols = ["总客次", *CUSTOMER_TYPES, "客户人数", "营业额"]
    if proj.empty:
        return pd.DataFrame(columns=cols)
    sign = proj["客次符号"]
    p = proj.assign(总客次=sign,
                    会员客次=sign.where(proj["支付归类"] == "套餐", 0),
                    **{"网红/邀约": sign.where(proj["支付归类"] == "免单", 0)})
    p["新客客次"] = p["总客次"] - p["会员客次"] - p["网红/邀约"]
    t = p.groupby(keys)[["总客次", *CUSTOMER_TYPES, "支付金额"]].sum().rename(columns={"支付金额": "营业额"})
    per_customer = p.groupby([*keys, "客户"])["客次符号"].sum().gt(0)
    t["客户人数"] = per_customer.groupby(level=list(range(len(keys)))).sum()
    return t[cols].sort_values("总客次", ascending=False)


def week_ranges(year, month, mode):
    """返回 [(标签, 开始日, 结束日)]，覆盖整个自然月。"""
    last = date(year, month, calendar.monthrange(year, month)[1])
    if mode == "natural":
        bounds, s = [], date(year, month, 1)
        while s <= last:
            e = min(s + timedelta(days=6 - s.weekday()), last)
            bounds.append((s, e))
            s = e + timedelta(days=1)
    else:
        bounds = [(date(year, month, d), min(date(year, month, d) + timedelta(days=6), last))
                  for d in range(1, last.day + 1, 7)]
    return [(f"W{i}", s, e) for i, (s, e) in enumerate(bounds, 1)]


@dataclass
class Report:
    start: date
    end: date
    month: dict
    weeks: list        # [(标签, 指标 dict 或 None)]，整月所有周，还没到的周为 None
    staff: list        # [(员工, 月度指标, [(周标签, 指标)])]，周只含已有数据的周
    categories: pd.DataFrame       # 按分类（月度）
    items: pd.DataFrame            # 按 分类+项目名称
    category_weeks: dict           # {"总客次"/"营业额": DataFrame(分类 × 周)}
    week_defs: list                # [(标签, 开始日, 结束日)]，整月所有周
    daily: list                    # [(日期, 指标)]，从月初到最后一天
    targets: object = None         # Targets 或 None


def build(df, week_mode):
    start, end = df["日期"].min().date(), df["日期"].max().date()
    week_defs = []
    for label, ws, we in week_ranges(start.year, start.month, week_mode):
        text = f"{label} {ws:%m/%d}-{we:%m/%d}"
        if ws <= end < we:
            text += f"（截至{end:%m/%d}）"
        week_defs.append((text, ws, we))
    active_weeks = [w for w in week_defs if w[1] <= end]

    def in_week(rows, ws, we):
        return rows[(rows["日期"] >= pd.Timestamp(ws)) & (rows["日期"] <= pd.Timestamp(we))]

    weeks = [(text, compute(in_week(df, ws, we)) if ws <= end else None) for text, ws, we in week_defs]

    staff = []
    for name, rows in df.groupby(df["员工"].fillna(UNASSIGNED)):
        staff.append((name, compute(rows), [(text, compute(in_week(rows, ws, we))) for text, ws, we in active_weeks]))
    staff.sort(key=lambda x: (x[0] == UNASSIGNED, -x[1]["业绩"], -x[1]["营业额"]))
    staff = [s for s in staff if s[0] != UNASSIGNED or any(s[1][k] for k in STAFF_METRICS)]

    proj = df[df["类型"] == "项目"].assign(分类=lambda x: x["分类"].fillna(UNCATEGORIZED))
    categories = mix_table(proj, ["分类"])
    items = mix_table(proj, ["分类", "名称"])
    order = {c: i for i, c in enumerate(categories.index)}
    items = items.sort_values(["总客次"], ascending=False).sort_index(
        level=0, key=lambda idx: idx.map(order), sort_remaining=False, kind="stable")
    category_weeks = {
        metric: pd.DataFrame({text: in_week(proj, ws, we).groupby("分类")[col].sum()
                              for text, ws, we in active_weeks}).reindex(categories.index).fillna(0)
        for metric, col in [("总客次", "客次符号"), ("营业额", "支付金额")]
    }
    first_day = date(start.year, start.month, 1)
    days = [first_day + timedelta(days=n) for n in range((end - first_day).days + 1)]
    daily = [(d, compute(df[df["日期"] == pd.Timestamp(d)])) for d in days]
    return Report(start, end, compute(df), weeks, staff, categories, items, category_weeks, week_defs, daily)


def checks(df, months, month_arg, sheet, staff):
    out = [f"读取 sheet「{sheet}」共 {len(df)} 行，账务日期 {df['日期'].min():%Y-%m-%d} ~ {df['日期'].max():%Y-%m-%d}。"]
    if len(months) > 1 and not month_arg:
        out.append(f"⚠ 文件包含多个月份（{', '.join(months)}），已合并统计；单月请加 --month。")
    unknown = sorted(set(df["支付方式"].dropna()) - KNOWN_PAY_METHODS)
    if unknown:
        out.append(f"⚠ 出现新的支付方式，已按「现金类」计入权益金/二次券：{'、'.join(unknown)}")
    types = sorted(set(df["类型"].dropna()) - {"项目", "权益"})
    if types:
        out.append(f"⚠ 出现未识别的类型（只计入叠加率的数量合计）：{'、'.join(types)}")
    no_date = int(df["日期"].isna().sum())
    if no_date:
        out.append(f"⚠ {no_date} 行没有账务日期，月度和员工数据已计入，周度数据未计入。")
    for name, m, _ in staff:
        if name == UNASSIGNED:
            out.append(f"⚠ 有记录没填员工：业绩 {m['业绩']:,.2f}、营业额 {m['营业额']:,.2f}、"
                       f"客次 {m['总客次']}，在员工表单列为「{UNASSIGNED}」。")
    refunds = df[df["客次符号"] < 0]
    if len(refunds):
        out.append(f"退款单 {refunds['单据编号'].nunique() if '单据编号' in df.columns else len(refunds)} 张"
                   f"（{len(refunds)} 行，支付金额合计 {refunds['支付金额'].sum():,.2f}），已按负数抵扣。")
    if "实业绩" in df.columns and pd.to_numeric(df["实业绩"], errors="coerce").fillna(0).eq(0).all():
        out.append("明细里「实业绩」列全为 0，员工「业绩」按权益金口径计算。")
    return out


# ---------------------------------------------------------------- 目标

@dataclass
class Targets:
    goals: dict      # 指标 -> 月目标
    weights: dict    # 日期 -> 当天占月目标的比例（合计 = 1），累加型指标按它拆到每天
    source: str
    month: str


def load_targets(path, start):
    """读目标 JSON（格式见 targets/2026-10.json）；没指定路径时找 targets/<YYYY-MM>.json，找不到返回 None。"""
    path = Path(path) if path else TARGETS_DIR / f"{start:%Y-%m}.json"
    if not path.exists():
        return None
    doc = json.loads(path.read_text(encoding="utf-8"))
    goals = {METRIC_ALIASES.get(k, k): float(v) for k, v in doc["月目标"].items()}
    unknown = sorted(set(goals) - set(MONTH_METRICS))
    if unknown:
        raise SystemExit(f"目标文件 {path.name} 里有报表不认识的指标：{'、'.join(unknown)}")
    last = calendar.monthrange(start.year, start.month)[1]
    daily = {date.fromisoformat(d): float(v) for d, v in doc.get("每日目标", {}).items()}
    if not daily:  # 没给每日目标就按天平均
        daily = {date(start.year, start.month, d): 1.0 for d in range(1, last + 1)}
    total = sum(daily.values())
    return Targets(goals, {d: v / total for d, v in daily.items()}, path.name, doc.get("月份", f"{start:%Y-%m}"))


def period_goal(tg, metric, first, last):
    """某段日期的目标：累加型 = 月目标 × 这段日期的权重之和；比率型 = 月目标。"""
    if metric in RATE_METRICS:
        return tg.goals[metric]
    return tg.goals[metric] * sum(w for d, w in tg.weights.items() if first <= d <= last)


def target_summary(rep):
    """返回 (目标进度, 剩余营业日列表, [每个目标指标的达成情况 dict])。"""
    tg = rep.targets
    first = date(rep.start.year, rep.start.month, 1)
    progress = sum(w for d, w in tg.weights.items() if d <= rep.end)
    remaining = sorted(d for d, w in tg.weights.items() if d > rep.end and w > 0)
    rows = []
    for k, goal in tg.goals.items():
        actual = rep.month[k] or 0
        to_date = period_goal(tg, k, first, rep.end)
        rows.append({
            "指标": k, "月目标": goal, "实际": actual,
            "月达成率": actual / goal if goal else 0,
            "同期目标": to_date, "同期达成率": actual / to_date if to_date else 0,
            "应达进度": 1.0 if k in RATE_METRICS else progress,
        })
    return progress, remaining, rows


# ---------------------------------------------------------------- 写 Excel：通用样式

FONT = "微软雅黑"
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
HELPER_FILL = PatternFill("solid", fgColor="F2F2F2")
SUBTOTAL_FILL = PatternFill("solid", fgColor="EEF3FA")
TILE_FILL = PatternFill("solid", fgColor="F5F7FA")
CENTER = Alignment(horizontal="center", vertical="center")
MUTED = "7F7F7F"

# 图表配色（dataviz 参考调色板）：同一个对象在所有图里用同一个颜色。
COLORS = {"权益金": "6250D6", "业绩": "6250D6", "营业额": "008300",
          "新客客次": "2A78D6", "会员客次": "EB6834", "网红/邀约": "1BAF7A"}
BAR_COLOR = "2A78D6"
TARGET_GREY = "A6A6A6"  # 目标 / 应达进度 一律用灰色
GOOD, BAD = "008300", "C00000"
STATUS_FORMAT = '[>=1]"✓ "0.0%;"✗ "0.0%'


def font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


def style_header(cell, helper=False):
    cell.font = font(bold=not helper, italic=helper, color=MUTED if helper else "000000")
    cell.fill = HELPER_FILL if helper else HEAD_FILL
    cell.border = BORDER
    cell.alignment = CENTER


def style_value(cell, metric, helper=False):
    cell.font = font(color=MUTED if helper else "000000", italic=helper)
    cell.border = BORDER
    cell.number_format = METRIC_FORMATS.get(metric, "0")
    if helper:
        cell.fill = HELPER_FILL


def label_cell(ws, row, col, text, bold=False, fill=None):
    c = ws.cell(row, col, text)
    c.font, c.border = font(bold=bold), BORDER
    if fill:
        c.fill = fill
    return c


def data_bar(ws, ref):
    ws.conditional_formatting.add(ref, DataBarRule(start_type="num", start_value=0, end_type="max",
                                                   color=BAR_COLOR))


def note(ws, row, col, text):
    ws.cell(row, col, text).font = font(size=9, color=MUTED)


def status_format(ws, ref):
    """达成率 ≥ 100% 绿色 ✓，否则红色 ✗（数字格式负责符号，条件格式负责颜色）。"""
    ws.conditional_formatting.add(ref, CellIsRule(operator="greaterThanOrEqual", formula=["1"],
                                                  font=Font(color=GOOD, bold=True)))
    ws.conditional_formatting.add(ref, CellIsRule(operator="lessThan", formula=["1"],
                                                  font=Font(color=BAD, bold=True)))


def progress_bar(ws, ref):
    ws.conditional_formatting.add(ref, DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                   color=BAR_COLOR))


def ratio_formula(num, den):
    return f'=IF(OR(NOT(ISNUMBER({num})),NOT(ISNUMBER({den})),{den}=0),"",{num}/{den})'


def goal_text(metric, v):
    if metric == "新客成交率":
        return f"{v:.0%}"
    if metric in ("总客次", "会员客次", "新客客次", "网红/邀约", "成交人数", "二次券成交数"):
        return f"{v:,.0f}"
    return f"R$ {v:,.0f}"


# ---------------------------------------------------------------- 经营报表 sheet

def rate_formula(num_col, den_col, row):
    return f'=IF({den_col}{row}=0,"",{num_col}{row}/{den_col}{row})'


def write_metric_row(ws, row, col0, m):
    """月度/周度数据行。派生指标写公式，引用同行的基础数与右侧「计算用」列。"""
    cols = {k: get_column_letter(col0 + i) for i, k in enumerate(MONTH_METRICS)}
    hcols = {k: get_column_letter(col0 + len(MONTH_METRICS) + 1 + i) for i, k in enumerate(HELPER_METRICS)}
    formulas = {
        "新客客次": f"={cols['总客次']}{row}-{cols['会员客次']}{row}-{cols['网红/邀约']}{row}",
        "成交客单": f"=IF({cols['成交人数']}{row}=0,0,{hcols['套餐成交金额']}{row}/{cols['成交人数']}{row})",
        "体验客单价": f"=IF({cols['总客次']}{row}=0,0,{cols['营业额']}{row}/{cols['总客次']}{row})",
        "叠加率": f"=IF({cols['总客次']}{row}=0,0,{hcols['数量合计']}{row}/{cols['总客次']}{row})",
        "新客成交率": rate_formula(cols["成交人数"], cols["新客客次"], row),
    }
    for k, letter in cols.items():
        cell = ws[f"{letter}{row}"]
        if m is not None:
            cell.value = formulas.get(k, round(m[k], 2))
        style_value(cell, k)
    for k, letter in hcols.items():
        cell = ws[f"{letter}{row}"]
        if m is not None:
            cell.value = round(m[k], 2)
        style_value(cell, k, helper=True)


def write_metric_header(ws, row, col0):
    for i, k in enumerate(MONTH_METRICS):
        style_header(ws.cell(row, col0 + i, k))
    for i, k in enumerate(HELPER_METRICS):
        style_header(ws.cell(row, col0 + len(MONTH_METRICS) + 1 + i, f"计算用:{k}"), helper=True)


def staff_columns(col0):
    return {k: get_column_letter(col0 + j) for j, k in enumerate(STAFF_METRICS + STAFF_HELPERS)}


def write_staff_header(ws, row, col0):
    for k, letter in staff_columns(col0).items():
        helper = k in STAFF_HELPERS
        style_header(ws[f"{letter}{row}"], helper=helper)
        ws[f"{letter}{row}"] = f"计算用:{k}" if helper else k


def write_month_targets(ws, row, col0, rep):
    """月度实际行下方的 4 行目标对比，只填有目标的列，返回下一个空行。"""
    _, _, rows = target_summary(rep)
    labels = ["月目标", "月达成率", f"同期目标（截至{rep.end:%m/%d}）", "同期达成率"]
    for i, text in enumerate(labels):
        label_cell(ws, row + i, 2, text, bold=i == 3)
        for j in range(len(MONTH_METRICS)):
            ws.cell(row + i, col0 + j).border = BORDER
    for t in rows:
        k = t["指标"]
        letter = get_column_letter(col0 + MONTH_METRICS.index(k))
        actual, goal, rate, to_date, to_rate = (f"{letter}{row + i}" for i in range(-1, 4))
        ws[goal] = round(t["月目标"], 6)
        ws[rate] = ratio_formula(actual, goal)
        ws[to_date] = round(t["同期目标"], 6)
        ws[to_rate] = ratio_formula(actual, to_date)
        for ref, fmt_key in [(goal, k), (rate, "月达成率"), (to_date, k)]:
            style_value(ws[ref], fmt_key)
        ws[to_rate].number_format = STATUS_FORMAT
    first, last = get_column_letter(col0), get_column_letter(col0 + len(MONTH_METRICS) - 1)
    progress_bar(ws, f"{first}{row + 1}:{last}{row + 1}")
    status_format(ws, f"{first}{row + 3}:{last}{row + 3}")
    return row + 4


def write_staff_values(ws, row, col0, m, bold=False, fill=None):
    """员工行。新客成交率写公式：成交人数 ÷ 右侧灰色「计算用:新客客次」。m=None 时写合计行（上方各行求和）。"""
    col = staff_columns(col0)
    for k, letter in col.items():
        c = ws[f"{letter}{row}"]
        helper = k in STAFF_HELPERS
        if k == "新客成交率":
            c.value = rate_formula(col["成交人数"], col["新客客次"], row)
        elif isinstance(m, tuple):
            c.value = f"=SUM({letter}{m[0]}:{letter}{m[1]})"
        else:
            c.value = round(m[k], 2)
        style_value(c, k, helper=helper)
        if bold:
            c.font = font(bold=True, italic=helper, color=MUTED if helper else "000000")
        if fill and not helper:
            c.fill = fill


def write_main_sheet(ws, rep, source_name):
    col0 = 3  # C 列开始放指标，与手工模板一致
    ws["A1"] = f"门店经营数据  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, f"数据来源：{source_name}　生成时间：{datetime.now():%Y-%m-%d %H:%M}　"
                   f"灰色「计算用」列为公式的底数；口径见「指标说明」")

    # 月度总数据（有目标时下面加 月目标 / 月达成率 / 同期目标 / 同期达成率）
    r = 4
    ws.cell(r, 1, "月度总数据").font = font(bold=True)
    write_metric_header(ws, r, col0)
    write_metric_row(ws, r + 1, col0, rep.month)
    r += 2
    if rep.targets:
        label_cell(ws, r - 1, 2, "实际")
        r = write_month_targets(ws, r, col0, rep)

    # 周度数据
    r += 1
    write_metric_header(ws, r, col0)
    ws.cell(r + 1, 1, "周度数据").font = font(bold=True)
    for i, (label, m) in enumerate(rep.weeks):
        row = r + 1 + i
        label_cell(ws, row, 2, label)
        ws.cell(row, 1).border = BORDER
        write_metric_row(ws, row, col0, m)
    last_week_row = r + len(rep.weeks)
    for k in ["权益金", "营业额", "总客次"]:
        letter = get_column_letter(col0 + MONTH_METRICS.index(k))
        data_bar(ws, f"{letter}{r + 1}:{letter}{last_week_row}")

    # 员工达成 · 月度数据（与手工模板一致）
    r = last_week_row + 3
    write_staff_header(ws, r, 4)
    ws.cell(r + 1, 1, "员工达成").font = font(bold=True)
    ws.cell(r + 1, 2, "月度数据").font = font(bold=True)
    for i, (name, m, _) in enumerate(rep.staff):
        row = r + 1 + i
        ws.cell(row, 3, name).border = BORDER
        ws.cell(row, 3).font = font(italic=name == UNASSIGNED)
        write_staff_values(ws, row, 4, m)
    total_row = r + 1 + len(rep.staff)
    label_cell(ws, total_row, 3, "合计", bold=True)
    write_staff_values(ws, total_row, 4, (r + 1, total_row - 1), bold=True)
    for letter in "DE":
        data_bar(ws, f"{letter}{r + 1}:{letter}{total_row - 1}")
    note(ws, total_row + 1, 3, "注：同一客户找过多位员工成交时会各算一次，故人数合计（及合计行的新客成交率）可能略高于月度总数据。")

    # 员工达成 · 月总周分：每位员工一块，先月度合计，再逐周
    r = total_row + 3
    ws.cell(r, 1, "员工周度数据").font = font(bold=True)
    style_header(ws.cell(r, 2, "员工"))
    style_header(ws.cell(r, 3, "期间"))
    write_staff_header(ws, r, 4)
    row = r + 1
    for name, m, weeks in rep.staff:
        first = row
        label_cell(ws, row, 3, "月度合计", bold=True, fill=SUBTOTAL_FILL)
        write_staff_values(ws, row, 4, m, bold=True, fill=SUBTOTAL_FILL)
        for label, wm in weeks:
            row += 1
            label_cell(ws, row, 3, label)
            write_staff_values(ws, row, 4, wm)
        ws.cell(first, 2, name).font = font(bold=True, italic=name == UNASSIGNED)
        ws.cell(first, 2).alignment = Alignment(vertical="center", wrap_text=True)
        for rr in range(first, row + 1):
            ws.cell(rr, 2).border = BORDER
        if row > first:
            ws.merge_cells(start_row=first, start_column=2, end_row=row, end_column=2)
        row += 1

    helper_start = col0 + len(MONTH_METRICS) + 1
    for i in range(col0, helper_start + len(HELPER_METRICS)):
        ws.column_dimensions[get_column_letter(i)].width = 20 if i >= helper_start else 14
    for letter, w in {"A": 12, "B": 26, "C": 26}.items():
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "C1"
    landscape(ws)


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------------------------------------------------------- 项目分类 sheet

def write_mix_block(ws, r, title, table, key_headers, with_share):
    """写一张 分类(或 分类+项目) 汇总表，返回下一个空行。占比、平均单价用公式。"""
    ws.cell(r, 1, title).font = font(bold=True, size=11)
    r += 1
    metrics = ["总客次", *CUSTOMER_TYPES, "客户人数", "营业额"]
    headers = [*key_headers, "总客次", *(["客次占比"] if with_share else []), *CUSTOMER_TYPES, "客户人数",
               "营业额", *(["营业额占比"] if with_share else []), "平均单价"]
    for j, h in enumerate(headers, 1):
        style_header(ws.cell(r, j, h))
    col = {h: get_column_letter(j) for j, h in enumerate(headers, 1)}
    first = r + 1
    last = r + len(table)
    total = last + 1
    prev_key = None
    for i, (key, x) in enumerate(table.iterrows()):
        row = first + i
        keys = key if isinstance(key, tuple) else (key,)
        for j, k in enumerate(keys, 1):
            shown = "" if (j == 1 and len(keys) > 1 and k == prev_key) else k
            label_cell(ws, row, j, shown, bold=j == 1)
        prev_key = keys[0]
        for k in metrics:
            style_value(ws[f"{col[k]}{row}"], k)
            ws[f"{col[k]}{row}"] = round(float(x[k]), 2)
        if with_share:
            for share, base in [("客次占比", "总客次"), ("营业额占比", "营业额")]:
                ws[f"{col[share]}{row}"] = f"=IF({col[base]}${total}=0,0,{col[base]}{row}/{col[base]}${total})"
                style_value(ws[f"{col[share]}{row}"], share)
        ws[f"{col['平均单价']}{row}"] = f"=IF({col['总客次']}{row}=0,0,{col['营业额']}{row}/{col['总客次']}{row})"
        style_value(ws[f"{col['平均单价']}{row}"], "平均单价")
    if with_share:
        label_cell(ws, total, 1, "合计", bold=True)
        for h in headers[1:]:
            c = ws[f"{col[h]}{total}"]
            if h in ("客次占比", "营业额占比"):
                c.value = f"=SUM({col[h]}{first}:{col[h]}{last})"
            elif h == "平均单价":
                c.value = f"=IF({col['总客次']}{total}=0,0,{col['营业额']}{total}/{col['总客次']}{total})"
            elif h == "客户人数":
                c.value = None  # 客户可能做过多个分类，人数不能直接相加
            else:
                c.value = f"=SUM({col[h]}{first}:{col[h]}{last})"
            style_value(c, h)
            c.font = font(bold=True)
    for k in ["总客次", "营业额"]:
        data_bar(ws, f"{col[k]}{first}:{col[k]}{last}")
    return (total if with_share else last) + 2


def write_week_matrix(ws, r, title, matrix, metric):
    ws.cell(r, 1, title).font = font(bold=True, size=11)
    r += 1
    headers = ["分类", *map(short_week, matrix.columns), "月度合计"]
    for j, h in enumerate(headers, 1):
        style_header(ws.cell(r, j, h))
    total_col = get_column_letter(len(headers))
    for i, (cat, x) in enumerate(matrix.iterrows(), r + 1):
        label_cell(ws, i, 1, cat, bold=True)
        for j, v in enumerate(x.values, 2):
            style_value(ws.cell(i, j, round(float(v), 2)), metric)
        c = ws[f"{total_col}{i}"]
        c.value = f"=SUM(B{i}:{get_column_letter(len(headers) - 1)}{i})"
        style_value(c, metric)
        c.font = font(bold=True)
    return r + len(matrix) + 2


def short_week(label):
    """「W3 10/15-10/21（截至10/18）」→「W3（截至10/18）」，用于表头和图表坐标轴。"""
    return label.split()[0] + (label[label.index("（"):] if "（" in label else "")


def write_category_sheet(wb, rep):
    ws = wb.create_sheet("项目分类")
    ws["A1"] = f"项目分类统计  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, "只统计类型=项目（客人实际做的项目）；新客/会员/网红 的划分与经营报表一致；平均单价=营业额÷总客次。"
                   "各项目明细见「项目明细」。")
    r = write_mix_block(ws, 4, "按分类（月度）", rep.categories, ["分类"], with_share=True)
    r = write_week_matrix(ws, r, "按分类 · 周度客次", rep.category_weeks["总客次"], "总客次")
    write_week_matrix(ws, r, "按分类 · 周度营业额", rep.category_weeks["营业额"], "营业额")
    ws.column_dimensions["A"].width = 24
    for j in range(2, 12):
        ws.column_dimensions[get_column_letter(j)].width = 15
    landscape(ws)

    ws = wb.create_sheet("项目明细")
    ws["A1"] = f"项目明细  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, "每个项目做了多少次、谁在做（新客/会员/网红）、带来多少营业额；按分类分组，组内按客次排序。")
    write_mix_block(ws, 4, "按项目（月度）", rep.items, ["分类", "项目名称"], with_share=False)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 46
    for j in range(3, 10):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = "C6"
    landscape(ws)


# ---------------------------------------------------------------- 目标达成 sheet

def write_target_sheet(wb, rep):
    ws = wb.create_sheet("目标达成")
    tg = rep.targets
    progress, remaining, rows = target_summary(rep)
    ws["A1"] = f"目标达成  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, f"目标来源：{tg.source}。✓ = 达到同期目标（按每日目标累计到 {rep.end:%m/%d}），✗ = 落后。"
                   f"口径见「指标说明」。")
    for col, (label, value, fmt_) in enumerate([("目标进度", progress, "0.0%"), ("剩余营业日", len(remaining), "0")]):
        label_cell(ws, 3, 1 + col * 2, label, bold=True, fill=HEAD_FILL)
        c = ws.cell(3, 2 + col * 2, value)
        c.number_format, c.font, c.border = fmt_, font(bold=True), BORDER
    days_left = "$D$3"

    # 月度
    r = 5
    ws.cell(r, 1, "月度目标达成").font = font(bold=True, size=11)
    headers = ["指标", "月目标", "实际", "月达成率", "同期目标", "同期达成率", "剩余需完成", "剩余日均需完成"]
    for j, h in enumerate(headers, 1):
        style_header(ws.cell(r + 1, j, h))
    for i, t in enumerate(rows):
        row, k = r + 2 + i, t["指标"]
        label_cell(ws, row, 1, k, bold=True)
        cells = {"月目标": round(t["月目标"], 6), "实际": round(t["实际"], 6),
                 "月达成率": f"=IF(B{row}=0,\"\",C{row}/B{row})", "同期目标": round(t["同期目标"], 6),
                 "同期达成率": f"=IF(E{row}=0,\"\",C{row}/E{row})"}
        if k not in RATE_METRICS:
            cells["剩余需完成"] = f"=MAX(B{row}-C{row},0)"
            cells["剩余日均需完成"] = f"=IF({days_left}=0,\"\",G{row}/{days_left})"
        for j, h in enumerate(headers[1:], 2):
            c = ws.cell(row, j, cells.get(h))
            style_value(c, "月达成率" if h == "月达成率" else k)
        ws[f"F{row}"].number_format = STATUS_FORMAT
    last = r + 1 + len(rows)
    progress_bar(ws, f"D{r + 2}:D{last}")
    status_format(ws, f"F{r + 2}:F{last}")

    # 周度
    r = last + 3
    ws.cell(r, 1, "周度目标达成").font = font(bold=True, size=11)
    note(ws, r, 3, "进行中的周用整周目标比较，达成率会偏低；还没到的周只列目标。")
    goals = list(tg.goals)
    style_header(ws.cell(r + 1, 1, "周"))
    for g, k in enumerate(goals):
        for j, h in enumerate([f"{k}目标", f"{k}实际", "达成率"]):
            style_header(ws.cell(r + 1, 2 + g * 3 + j, h))
    for i, ((label, ws_, we), (_, m)) in enumerate(zip(rep.week_defs, rep.weeks)):
        row = r + 2 + i
        label_cell(ws, row, 1, label)
        for g, k in enumerate(goals):
            tcol, acol, rcol = (get_column_letter(2 + g * 3 + j) for j in range(3))
            ws[f"{tcol}{row}"] = round(period_goal(tg, k, ws_, we), 6)
            ws[f"{acol}{row}"] = None if m is None or m[k] is None else round(m[k], 6)
            ws[f"{rcol}{row}"] = ratio_formula(f"{acol}{row}", f"{tcol}{row}") if m is not None else None
            style_value(ws[f"{tcol}{row}"], k)
            style_value(ws[f"{acol}{row}"], k)
            style_value(ws[f"{rcol}{row}"], "月达成率")
            ws[f"{rcol}{row}"].number_format = STATUS_FORMAT
    last = r + 1 + len(rep.week_defs)
    for g in range(len(goals)):
        letter = get_column_letter(4 + g * 3)
        status_format(ws, f"{letter}{r + 2}:{letter}{last}")

    # 每日（只列累加型目标）
    r = last + 3
    additive = [k for k in goals if k not in RATE_METRICS]
    ws.cell(r, 1, "每日目标 vs 实际").font = font(bold=True, size=11)
    note(ws, r, 3, "累计达成率 = 月初到当天的累计实际 ÷ 累计目标。")
    style_header(ws.cell(r + 1, 1, "日期"))
    style_header(ws.cell(r + 1, 2, "星期"))
    sub = ["目标", "实际", "累计目标", "累计实际", "累计达成率"]
    for g, k in enumerate(additive):
        for j, h in enumerate(sub):
            style_header(ws.cell(r + 1, 3 + g * 5 + j, f"{k}{h}" if j < 4 else h))
    actual_by_day = dict(rep.daily)
    first_row = r + 2
    month_days = sorted(tg.weights)
    for i, d in enumerate(month_days):
        row = first_row + i
        c = ws.cell(row, 1, d)
        c.number_format, c.font, c.border = "mm/dd", font(), BORDER
        label_cell(ws, row, 2, f"周{WEEKDAYS[d.weekday()]}" + ("（休）" if tg.weights[d] == 0 else ""))
        m = actual_by_day.get(d) if d <= rep.end else None
        for g, k in enumerate(additive):
            t, a, ct, ca, cr = (get_column_letter(3 + g * 5 + j) for j in range(5))
            ws[f"{t}{row}"] = round(tg.goals[k] * tg.weights[d], 6)
            ws[f"{ct}{row}"] = f"=SUM({t}${first_row}:{t}{row})"
            if m is not None:
                ws[f"{a}{row}"] = round(m[k], 6)
                ws[f"{ca}{row}"] = f"=SUM({a}${first_row}:{a}{row})"
                ws[f"{cr}{row}"] = ratio_formula(f"{ca}{row}", f"{ct}{row}")
            for ref in (t, a, ct, ca):
                style_value(ws[f"{ref}{row}"], k)
            style_value(ws[f"{cr}{row}"], "累计达成率")
            ws[f"{cr}{row}"].number_format = STATUS_FORMAT
    last = first_row + len(month_days) - 1
    for g in range(len(additive)):
        letter = get_column_letter(7 + g * 5)
        status_format(ws, f"{letter}{first_row}:{letter}{last}")

    ws.column_dimensions["A"].width = 26
    for j in range(2, 3 + 5 * max(len(additive), 1) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    landscape(ws)


# ---------------------------------------------------------------- 看板 sheet（KPI + 图表）

def write_block(ds, r, headers, rows, fmts):
    """在「图表数据」写一块表：表头 + 行，返回 (表头行, 最后一行)。"""
    for j, h in enumerate(headers, 1):
        style_header(ds.cell(r, j, h))
    for i, values in enumerate(rows, r + 1):
        for j, v in enumerate(values, 1):
            c = ds.cell(i, j, v)
            if j > 1:
                c.number_format = fmts[j - 2] if isinstance(fmts, list) else fmts
    return r, r + len(rows)


def write_chart_data(wb, rep):
    """图表引用的数据（纯数值），返回各块的位置。"""
    ds = wb.create_sheet("图表数据")
    note(ds, 1, 1, "看板图表的数据来源，由脚本自动生成。")
    pos = {}

    weeks = [(label, m) for label, m in rep.weeks if m is not None]
    headers = ["周", "权益金", "营业额", *CUSTOMER_TYPES]
    pos["weeks"] = write_block(ds, 3, headers, [[short_week(label), *(round(m[k], 2) for k in headers[1:])]
                                                for label, m in weeks], "#,##0")

    # 横向条形图从下往上画，所以按升序写，最大的排在最上面
    staff = sorted(rep.staff, key=lambda s: (s[1]["业绩"], s[1]["营业额"]))
    pos["staff"] = write_block(ds, pos["weeks"][1] + 3, ["员工", "业绩", "营业额"],
                               [[name, round(m["业绩"], 2), round(m["营业额"], 2)] for name, m, _ in staff], "#,##0")

    headers = ["分类", *CUSTOMER_TYPES, "营业额"]
    pos["categories"] = write_block(ds, pos["staff"][1] + 3, headers,
                                    [[cat, *(round(float(x[k]), 2) for k in headers[1:])]
                                     for cat, x in rep.categories.iloc[::-1].iterrows()], "#,##0")

    if rep.targets:
        tg = rep.targets
        money = [k for k in ("权益金", "营业额") if k in tg.goals]
        _, _, rows = target_summary(rep)
        pos["achieve"] = write_block(ds, pos["categories"][1] + 3, ["指标", "月达成率", "应达进度"],
                                     [[t["指标"], round(t["月达成率"], 4), round(t["应达进度"], 4)]
                                      for t in reversed(rows)], "0%")
        actual_by_day = dict(rep.daily)
        cum_rows, cum = [], {k: [0.0, 0.0] for k in money}
        for d in sorted(tg.weights):
            row = [f"{d:%m/%d}"]
            for k in money:
                cum[k][0] += tg.goals[k] * tg.weights[d]
                if d <= rep.end:
                    cum[k][1] += actual_by_day[d][k] if d in actual_by_day else 0
                row += [round(cum[k][0], 2), round(cum[k][1], 2) if d <= rep.end else None]
            cum_rows.append(row)
        pos["cumulative"] = write_block(ds, pos["achieve"][1] + 3,
                                        ["日期", *(h for k in money for h in (f"{k}累计目标", f"{k}累计"))],
                                        cum_rows, "#,##0")
        week_rows = [[short_week(label), *(v for k in money for v in (
            round(period_goal(tg, k, ws_, we), 2), round(m[k], 2) if m else None))]
            for (label, ws_, we), (_, m) in zip(rep.week_defs, rep.weeks)]
        pos["week_targets"] = write_block(ds, pos["cumulative"][1] + 3,
                                          ["周", *(h for k in money for h in (f"{k}目标", k))], week_rows, "#,##0")
        pos["money"] = money

    ds.column_dimensions["A"].width = 30
    return ds, pos


def series_color(name):
    if name.endswith("目标") or name == "应达进度":
        return TARGET_GREY
    return COLORS.get(name.replace("累计", ""), BAR_COLOR)


def finish_chart(ch, n_series, num_fmt):
    if n_series > 1:
        ch.legend.position = "b"
    else:
        ch.legend = None
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = num_fmt
    return ch


def make_chart(ds, header_row, last_row, cols, title, horizontal=False, stacked=False, labels=False,
               width=16.5, height=7.5, num_fmt="#,##0"):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.title = title
    ch.style = 2
    ch.width, ch.height = width, height
    if stacked:
        ch.grouping, ch.overlap = "stacked", 100
    ch.gapWidth = 60
    for c in cols:
        ch.add_data(Reference(ds, min_col=c, min_row=header_row, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(ds, min_col=1, min_row=header_row + 1, max_row=last_row))
    for s, c in zip(ch.series, cols):
        s.graphicalProperties.solidFill = series_color(ds.cell(header_row, c).value)
        s.graphicalProperties.line.solidFill = "FFFFFF"
    if labels:
        ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                      showLegendKey=False, showPercent=False)
    return finish_chart(ch, len(cols), num_fmt)


def make_line_chart(ds, header_row, last_row, cols, title, width=16.5, height=7.5):
    """累计 vs 目标：目标是灰色虚线，实际是该指标的颜色；还没到的日子留空（断开）。"""
    ch = LineChart()
    ch.title = title
    ch.style = 2
    ch.width, ch.height = width, height
    for c in cols:
        ch.add_data(Reference(ds, min_col=c, min_row=header_row, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(ds, min_col=1, min_row=header_row + 1, max_row=last_row))
    for s, c in zip(ch.series, cols):
        name = ds.cell(header_row, c).value
        s.graphicalProperties.line.solidFill = series_color(name)
        s.graphicalProperties.line.width = 28575
        if name.endswith("目标"):
            s.graphicalProperties.line.dashStyle = "dash"
        s.marker.symbol = "none"
        s.smooth = False
    ch.display_blanks = "gap"
    ch.x_axis.tickLblSkip = 3
    return finish_chart(ch, len(cols), "#,##0")


def write_tiles(ws, rep):
    """月度 12 个指标的 KPI 卡片（3 行 × 4 个）。有目标的卡片加「目标 · 同期 ✓/✗」和月达成进度条。返回下一个空行。"""
    m = rep.month
    goals = {t["指标"]: t for t in target_summary(rep)[2]} if rep.targets else {}
    voucher_etc = m["权益金"] - m["套餐成交金额"]
    rate = m["新客成交率"]
    tiles = [
        ("权益金", '"R$ "#,##0', m["权益金"], f"套餐 R$ {m['套餐成交金额']:,.0f} · 券及其他 R$ {voucher_etc:,.0f}"),
        ("营业额", '"R$ "#,##0', m["营业额"], "项目核销价值"),
        ("总客次", "#,##0", m["总客次"], "做项目的人次"),
        ("新客客次", "#,##0", m["新客客次"],
         f"占总客次 {m['新客客次'] / m['总客次']:.0%}" if m["总客次"] else ""),
        ("会员客次", "#,##0", m["会员客次"], "套餐核销"),
        ("网红/邀约", "#,##0", m["网红/邀约"], "免单类"),
        ("成交人数", "#,##0", m["成交人数"], "买套餐的客户（去重）"),
        ("新客成交率", "0.00%", rate if rate is not None else "-",
         f"成交人数 {m['成交人数']} ÷ 新客客次 {m['新客客次']}"),
        ("成交客单", '"R$ "#,##0', m["成交客单"], "套餐金额 ÷ 成交人数"),
        ("二次券成交数", "#,##0", m["二次券成交数"], f"现场买券的客户 {m['二次券成交人数']} 人"),
        ("体验客单价", '"R$ "#,##0.00', m["体验客单价"], "营业额 ÷ 总客次"),
        ("叠加率", "0.00", m["叠加率"], "数量合计 ÷ 总客次"),
    ]
    height = 4 if goals else 3
    for i, (label, fmt_, value, sub) in enumerate(tiles):
        row = 4 + (i // 4) * (height + 1)
        col = 2 + (i % 4) * 4
        for rr in range(row, row + height):
            ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 2)
            for cc in range(col, col + 3):
                ws.cell(rr, cc).fill = TILE_FILL
            ws.cell(rr, col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.cell(row, col, label).font = font(size=10, color=MUTED)
        v = ws.cell(row + 1, col, round(value, 6) if isinstance(value, (int, float)) else value)
        v.font, v.number_format = font(size=18, bold=True), fmt_
        ws.row_dimensions[row + 1].height = 28
        t = goals.get(label)
        if t:
            ok = t["同期达成率"] >= 1
            sub = f"目标 {goal_text(label, t['月目标'])} · 同期 {'✓' if ok else '✗'} {t['同期达成率']:.0%}"
            ws.cell(row + 2, col, sub).font = font(size=9, bold=True, color=GOOD if ok else BAD)
            bar = ws.cell(row + 3, col, round(t["月达成率"], 4))
            bar.number_format, bar.font = '"月达成 "0%', font(size=8, color=MUTED)
            progress_bar(ws, bar.coordinate)
        else:
            ws.cell(row + 2, col, sub).font = font(size=8, color=MUTED)
    return 4 + -(-len(tiles) // 4) * (height + 1) + 1


def write_dashboard(wb, rep):
    ws = wb.create_sheet("看板", 0)
    ws.sheet_view.showGridLines = False
    for j in range(1, 18):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions["A"].width = 2
    ws["B1"] = f"门店经营看板  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["B1"].font = font(size=16, bold=True)
    hint = "数值明细见「经营报表」「项目分类」；口径见「指标说明」。"
    if rep.targets:
        progress = target_summary(rep)[0]
        hint = (f"目标进度 {progress:.1%}（按每日目标累计到 {rep.end:%m/%d}）；✓/✗ = 是否达到同期目标，"
                f"进度条 = 月目标完成比例。详见「目标达成」。") + hint
    note(ws, 2, 2, hint)
    top = write_tiles(ws, rep)

    ds, pos = write_chart_data(wb, rep)
    w0, w1 = pos["weeks"]
    s0, s1 = pos["staff"]
    c0, c1 = pos["categories"]
    staff_height = max(7.5, 0.75 * (s1 - s0) + 2.5)
    cat_height = max(7.5, 0.75 * (c1 - c0) + 2.5)
    charts = []
    if rep.targets:
        d0, d1 = pos["cumulative"]
        t0, t1 = pos["week_targets"]
        a0, a1 = pos["achieve"]
        for i, k in enumerate(pos["money"]):
            charts.append(make_line_chart(ds, d0, d1, [2 + i * 2, 3 + i * 2], f"{k} 累计 vs 目标（R$）"))
        charts.append(make_chart(ds, a0, a1, [2, 3], "目标达成率 vs 应达进度", horizontal=True, labels=True,
                                 height=max(7.5, 1.1 * (a1 - a0) + 2.5), num_fmt="0%"))
        for i, k in enumerate(pos["money"]):
            charts.append(make_chart(ds, t0, t1, [2 + i * 2, 3 + i * 2], f"周度 {k} 实际 vs 目标（R$）"))
    else:
        charts.append(make_chart(ds, w0, w1, [2, 3], "周度 权益金 vs 营业额（R$）", labels=True))
    charts += [
        make_chart(ds, w0, w1, [4, 5, 6], "周度客次构成（新客/会员/网红）", stacked=True),
        make_chart(ds, s0, s1, [2, 3], "员工 业绩 vs 营业额（R$）", horizontal=True, height=staff_height),
        make_chart(ds, c0, c1, [2, 3, 4], "项目分类客次（按客户类型）", horizontal=True, stacked=True,
                   height=cat_height),
        make_chart(ds, c0, c1, [5], "项目分类营业额（R$）", horizontal=True, labels=True, height=cat_height),
    ]
    row = top
    for i in range(0, len(charts), 2):  # 两列排布，每行高度取两张图中较高的
        pair = charts[i:i + 2]
        for ch, col in zip(pair, "BJ"):
            ws.add_chart(ch, f"{col}{row}")
        row += int(max(ch.height for ch in pair) / 0.53) + 2
    landscape(ws)


# ---------------------------------------------------------------- 其他 sheet 与入口

def write_detail(wb, title, rows, amount_col):
    ws = wb.create_sheet(title)
    cols = [("日期", 12), ("单据编号", 16), ("客户名称", 30), ("员工", 24), ("名称", 36),
            ("数量", 8), ("支付方式", 28), ("金额", 12)]
    for j, (h, w) in enumerate(cols, 1):
        style_header(ws.cell(1, j, h))
        ws.column_dimensions[get_column_letter(j)].width = w
    rows = rows.sort_values(["日期", "客户名称"])
    for i, (_, x) in enumerate(rows.iterrows(), 2):
        values = [x["日期"].date() if pd.notna(x["日期"]) else None, x.get("单据编号"), x["客户名称"],
                  x["员工"] or UNASSIGNED, x["名称"], x["数量"], x["支付方式"], x[amount_col]]
        for j, v in enumerate(values, 1):
            c = ws.cell(i, j, None if pd.isna(v) else v)
            c.font, c.border = font(), BORDER
        ws.cell(i, 1).number_format = "yyyy-mm-dd"
        ws.cell(i, 8).number_format = "#,##0.00"
    n = len(rows) + 2
    ws.cell(n, 5, "合计").font = font(bold=True)
    for j in (6, 8):
        letter = get_column_letter(j)
        c = ws.cell(n, j, f"=SUM({letter}2:{letter}{n - 1})")
        c.font = font(bold=True)
        c.number_format = "#,##0.00" if j == 8 else "0"
    ws.freeze_panes = "A2"


def write_report(path, df, rep, notes, source_name, week_mode):
    wb = Workbook()
    write_main_sheet(wb.active, rep, source_name)
    wb.active.title = "经营报表"
    if rep.targets:
        write_target_sheet(wb, rep)
    write_category_sheet(wb, rep)

    ds = wb.create_sheet("指标说明")
    for j, h in enumerate(["指标", "计算口径", "本期数值"], 1):
        style_header(ds.cell(1, j, h))
    for i, (k, text) in enumerate(DEFINITIONS, 2):
        ds.cell(i, 1, k).font = font(bold=True)
        ds.cell(i, 2, text).font = font()
        ds.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
        if rep.month.get(k) is not None:
            style_value(ds.cell(i, 3, round(rep.month[k], 6)), k)
        for j in (1, 2, 3):
            ds.cell(i, j).border = BORDER
    note(ds, len(DEFINITIONS) + 2, 1,
         f"周度划分方式：{'自然周(周一~周日)' if week_mode == 'natural' else '每月按 7 天一段'}")
    ds.column_dimensions["A"].width = 20
    ds.column_dimensions["B"].width = 90
    ds.column_dimensions["C"].width = 14

    # 核对明细（相当于原来的透视表）
    right = df[df["类型"] == "权益"]
    package = right[(right["分类"] == "套餐") & (right["现金类支付"] != 0)]
    voucher_sale = right[(right["分类"] == "代金券") & (right["支付归类"] == "现金类")]
    write_detail(wb, "套餐成交明细", package, "现金类支付")
    write_detail(wb, "二次券成交明细", voucher_sale, "支付金额")

    cs = wb.create_sheet("数据检查")
    style_header(cs.cell(1, 1, "数据检查 / 提示"))
    for i, line in enumerate(notes, 2):
        cs.cell(i, 1, line).font = font(color="C00000" if line.startswith("⚠") else "000000")
    cs.column_dimensions["A"].width = 120

    write_dashboard(wb, rep)  # 放在第一个 sheet；图表数据 sheet 排在最后
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def fmt(k, v):
    if v is None:
        return "-"
    if k == "新客成交率":
        return f"{v:.2%}"
    return f"{v:,.2f}" if isinstance(v, float) else str(v)


def main(argv=None):
    ap = argparse.ArgumentParser(description="售出明细表 → 门店经营报表")
    ap.add_argument("input", help="收银系统导出的 Excel（含「售出明细表」）")
    ap.add_argument("-o", "--output", help="输出路径，默认 ./门店经营报表_<开始>_<结束>.xlsx")
    ap.add_argument("--month", help="只统计某个月，例如 2026-10")
    ap.add_argument("--weeks", choices=["fixed", "natural"], default="fixed",
                    help="fixed: 1-7日=W1…（默认）；natural: 周一~周日")
    ap.add_argument("--targets", help="目标 JSON，默认自动找 targets/<YYYY-MM>.json")
    args = ap.parse_args(argv)

    df, sheet, months = load(args.input, args.month)
    rep = build(df, args.weeks)
    rep.targets = load_targets(args.targets, rep.start)
    notes = checks(df, months, args.month, sheet, rep.staff)
    if rep.targets:
        notes.append(f"目标文件：{rep.targets.source}（{'、'.join(rep.targets.goals)}）。")
        if rep.targets.month != f"{rep.start:%Y-%m}":
            notes.append(f"⚠ 目标文件写的是 {rep.targets.month}，数据是 {rep.start:%Y-%m}，请确认目标是否用对。")
    else:
        notes.append(f"没有 {rep.start:%Y-%m} 的目标文件（targets/{rep.start:%Y-%m}.json），本次报表不含目标达成。")
    out = Path(args.output or f"门店经营报表_{rep.start:%Y%m%d}_{rep.end:%Y%m%d}.xlsx")
    write_report(out, df, rep, notes, Path(args.input).name, args.weeks)

    print(f"已生成：{out}")
    print(f"区间：{rep.start} ~ {rep.end}")
    print("月度总数据：" + "，".join(f"{k} {fmt(k, rep.month[k])}" for k in MONTH_METRICS))
    for label, m in rep.weeks:
        if m:
            print(f"  {label}：权益金 {m['权益金']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
                  f"成交人数 {m['成交人数']}，二次券 {m['二次券成交数']}，新客成交率 {fmt('新客成交率', m['新客成交率'])}")
    print("员工达成（月度｜各周业绩）：")
    for name, m, weeks in rep.staff:
        by_week = " / ".join(f"{label.split()[0]} {wm['业绩']:,.0f}" for label, wm in weeks)
        print(f"  {name}：业绩 {m['业绩']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
              f"成交人数 {m['成交人数']}，二次券成交人数 {m['二次券成交人数']}，"
              f"新客成交率 {fmt('新客成交率', m['新客成交率'])}｜{by_week}")
    if rep.targets:
        progress, remaining, rows = target_summary(rep)
        print(f"目标达成（目标进度 {progress:.1%}，剩余营业日 {len(remaining)} 天）：")
        for t in rows:
            k = t["指标"]
            print(f"  {k}：实际 {fmt(k, t['实际'])} / 月目标 {fmt(k, t['月目标'])} = {t['月达成率']:.1%}；"
                  f"同期目标 {fmt(k, t['同期目标'])}，同期达成 {t['同期达成率']:.1%} {'✓' if t['同期达成率'] >= 1 else '✗'}")
    print("项目分类（总客次｜新客/会员/网红｜营业额）：")
    total = rep.categories["总客次"].sum()
    for cat, x in rep.categories.iterrows():
        print(f"  {cat}：{x['总客次']:.0f}（{x['总客次'] / total:.0%}）｜{x['新客客次']:.0f}/{x['会员客次']:.0f}/"
              f"{x['网红/邀约']:.0f}｜{x['营业额']:,.2f}")
    for line in notes:
        print(line)


if __name__ == "__main__":
    sys.exit(main())

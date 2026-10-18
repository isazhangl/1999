#!/usr/bin/env python3
"""门店经营报表：读取收银系统导出的「售出明细表」，生成 看板 / 月度总数据 / 周度数据 / 员工达成 / 项目分类 报表。

用法:
    python3 store_report.py 售出明细.xlsx                      # 输出到当前目录
    python3 store_report.py 售出明细.xlsx -o 报表.xlsx
    python3 store_report.py 售出明细.xlsx --month 2026-10      # 文件含多个月时只取某月
    python3 store_report.py 售出明细.xlsx --weeks natural      # 周按 周一~周日 划分（默认 1-7日=W1）

指标口径见 README.md，与 2026-10 手工透视表结果逐项核对一致。
"""

import argparse
import calendar
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["类型", "分类", "名称", "数量", "支付方式", "支付金额", "现金类支付", "员工", "客户名称", "账务日期"]

# 支付方式归类。新出现的支付方式默认按「现金类」处理，并在「数据检查」中提示。
KNOWN_PAY_METHODS = {
    "代金券", "开业券 Voucher comprado online", "套餐", "免单类",
    "现金", "信用卡 Crédito à vista", "信用卡分期 Crédito parcelado", "银行卡 Débito", "PIX",
}

MONTH_METRICS = ["权益金", "营业额", "总客次", "会员客次", "新客客次", "网红/邀约",
                 "成交人数", "成交客单", "二次券成交数", "体验客单价", "叠加率"]
HELPER_METRICS = ["套餐成交金额", "数量合计"]
STAFF_METRICS = ["业绩", "营业额", "总客次", "成交人数", "二次券成交人数"]
CUSTOMER_TYPES = ["新客客次", "会员客次", "网红/邀约"]
UNASSIGNED = "未指定员工（线上购券）"
UNCATEGORIZED = "（未分类）"

METRIC_FORMATS = {
    "权益金": "#,##0.00", "营业额": "#,##0.00", "成交客单": "#,##0", "体验客单价": "#,##0.00",
    "叠加率": "0.00", "套餐成交金额": "#,##0.00", "业绩": "#,##0.00", "平均单价": "#,##0.00",
    "客次占比": "0.0%", "营业额占比": "0.0%",
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
    ("员工·业绩", "该员工名下的权益金（卖套餐+卖券收到的钱，口径同「权益金」）。明细里「实业绩」列全为 0，故不用。"),
    ("员工·营业额/总客次", "该员工名下的项目营业额 / 项目行数，口径同上。"),
    ("员工·成交人数", "该员工名下买套餐的客户人数（去重）。"),
    ("员工·二次券成交人数", "该员工名下现场掏钱买代金券的客户人数（去重）。"),
    ("项目分类", "类型=项目 的行按收银系统「分类」汇总；新客/会员/网红 的划分与上面客次口径相同；"
                "客户人数=做过该分类项目的去重客户数；平均单价=营业额÷总客次。"),
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
    return Report(start, end, compute(df), weeks, staff, categories, items, category_weeks)


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


# ---------------------------------------------------------------- 经营报表 sheet

def write_metric_row(ws, row, col0, m):
    """月度/周度数据行。派生指标写公式，引用同行的基础数与右侧「计算用」列。"""
    cols = {k: get_column_letter(col0 + i) for i, k in enumerate(MONTH_METRICS)}
    hcols = {k: get_column_letter(col0 + len(MONTH_METRICS) + 1 + i) for i, k in enumerate(HELPER_METRICS)}
    formulas = {
        "新客客次": f"={cols['总客次']}{row}-{cols['会员客次']}{row}-{cols['网红/邀约']}{row}",
        "成交客单": f"=IF({cols['成交人数']}{row}=0,0,{hcols['套餐成交金额']}{row}/{cols['成交人数']}{row})",
        "体验客单价": f"=IF({cols['总客次']}{row}=0,0,{cols['营业额']}{row}/{cols['总客次']}{row})",
        "叠加率": f"=IF({cols['总客次']}{row}=0,0,{hcols['数量合计']}{row}/{cols['总客次']}{row})",
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


def write_staff_values(ws, row, col0, m, bold=False, fill=None):
    for j, k in enumerate(STAFF_METRICS):
        c = ws.cell(row, col0 + j, round(m[k], 2))
        style_value(c, k)
        if bold:
            c.font = font(bold=True)
        if fill:
            c.fill = fill


def write_main_sheet(ws, rep, source_name):
    col0 = 3  # C 列开始放指标，与手工模板一致
    ws["A1"] = f"门店经营数据  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, f"数据来源：{source_name}　生成时间：{datetime.now():%Y-%m-%d %H:%M}　"
                   f"灰色「计算用」列为公式的底数；口径见「指标说明」")

    # 月度总数据
    r = 4
    ws.cell(r, 1, "月度总数据").font = font(bold=True)
    write_metric_header(ws, r, col0)
    write_metric_row(ws, r + 1, col0, rep.month)

    # 周度数据
    r = 7
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
    for i, k in enumerate(STAFF_METRICS):
        style_header(ws.cell(r, 4 + i, k))
    ws.cell(r + 1, 1, "员工达成").font = font(bold=True)
    ws.cell(r + 1, 2, "月度数据").font = font(bold=True)
    for i, (name, m, _) in enumerate(rep.staff):
        row = r + 1 + i
        ws.cell(row, 3, name).border = BORDER
        ws.cell(row, 3).font = font(italic=name == UNASSIGNED)
        write_staff_values(ws, row, 4, m)
    total_row = r + 1 + len(rep.staff)
    label_cell(ws, total_row, 3, "合计", bold=True)
    for j, k in enumerate(STAFF_METRICS):
        letter = get_column_letter(4 + j)
        cell = ws.cell(total_row, 4 + j, f"=SUM({letter}{r + 1}:{letter}{total_row - 1})")
        style_value(cell, k)
        cell.font = font(bold=True)
    for letter in "DE":
        data_bar(ws, f"{letter}{r + 1}:{letter}{total_row - 1}")
    note(ws, total_row + 1, 3, "注：同一客户找过多位员工成交时会各算一次，故人数合计可能大于月度去重人数。")

    # 员工达成 · 月总周分：每位员工一块，先月度合计，再逐周
    r = total_row + 3
    ws.cell(r, 1, "员工周度数据").font = font(bold=True)
    style_header(ws.cell(r, 2, "员工"))
    style_header(ws.cell(r, 3, "期间"))
    for j, k in enumerate(STAFF_METRICS):
        style_header(ws.cell(r, 4 + j, k))
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


# ---------------------------------------------------------------- 看板 sheet（KPI + 图表）

def write_chart_data(wb, rep):
    """图表引用的数据（纯数值），返回各块的位置。"""
    ds = wb.create_sheet("图表数据")
    note(ds, 1, 1, "看板图表的数据来源，由脚本自动生成。")
    pos = {}

    weeks = [(label, m) for label, m in rep.weeks if m is not None]
    r = 3
    headers = ["周", "权益金", "营业额", *CUSTOMER_TYPES]
    for j, h in enumerate(headers, 1):
        style_header(ds.cell(r, j, h))
    for i, (label, m) in enumerate(weeks, r + 1):
        ds.cell(i, 1, short_week(label))
        for j, k in enumerate(headers[1:], 2):
            ds.cell(i, j, round(m[k], 2)).number_format = "#,##0"
    pos["weeks"] = (r, r + len(weeks))

    # 横向条形图从下往上画，所以按升序写，最大的排在最上面
    staff = sorted(rep.staff, key=lambda s: (s[1]["业绩"], s[1]["营业额"]))
    r = pos["weeks"][1] + 3
    for j, h in enumerate(["员工", "业绩", "营业额"], 1):
        style_header(ds.cell(r, j, h))
    for i, (name, m, _) in enumerate(staff, r + 1):
        ds.cell(i, 1, name)
        ds.cell(i, 2, round(m["业绩"], 2)).number_format = "#,##0"
        ds.cell(i, 3, round(m["营业额"], 2)).number_format = "#,##0"
    pos["staff"] = (r, r + len(staff))

    cats = rep.categories.iloc[::-1]
    r = pos["staff"][1] + 3
    headers = ["分类", *CUSTOMER_TYPES, "营业额"]
    for j, h in enumerate(headers, 1):
        style_header(ds.cell(r, j, h))
    for i, (cat, x) in enumerate(cats.iterrows(), r + 1):
        ds.cell(i, 1, cat)
        for j, k in enumerate(headers[1:], 2):
            ds.cell(i, j, round(float(x[k]), 2)).number_format = "#,##0"
    pos["categories"] = (r, r + len(cats))

    ds.column_dimensions["A"].width = 30
    return ds, pos


def make_chart(ds, header_row, last_row, cols, title, horizontal=False, stacked=False, labels=False,
               width=16.5, height=7.5):
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
        color = COLORS.get(ds.cell(header_row, c).value, BAR_COLOR)
        s.graphicalProperties.solidFill = color
        s.graphicalProperties.line.solidFill = "FFFFFF"
    if labels:
        ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                      showLegendKey=False, showPercent=False)
    if len(cols) > 1:
        ch.legend.position = "b"
    else:
        ch.legend = None
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.numFmt = "#,##0"
    ch.y_axis.majorGridlines.spPr = None
    return ch


def write_dashboard(wb, rep):
    ws = wb.create_sheet("看板", 0)
    ws.sheet_view.showGridLines = False
    for j in range(1, 18):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions["A"].width = 2
    ws["B1"] = f"门店经营看板  {rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"
    ws["B1"].font = font(size=16, bold=True)
    note(ws, 2, 2, "数值明细见「经营报表」「项目分类」；口径见「指标说明」。")

    m = rep.month
    voucher_etc = m["权益金"] - m["套餐成交金额"]
    tiles = [
        ("权益金", f'"R$ "#,##0', m["权益金"], f"套餐 R$ {m['套餐成交金额']:,.0f} · 券及其他 R$ {voucher_etc:,.0f}"),
        ("营业额", f'"R$ "#,##0', m["营业额"], "项目核销价值"),
        ("总客次", "#,##0", m["总客次"], f"新客 {m['新客客次']} · 会员 {m['会员客次']} · 网红/邀约 {m['网红/邀约']}"),
        ("成交人数", "#,##0", m["成交人数"], "买套餐的客户（去重）"),
        ("成交客单", f'"R$ "#,##0', m["成交客单"], "套餐金额 ÷ 成交人数"),
        ("二次券成交数", "#,##0", m["二次券成交数"], f"现场买券的客户 {m['二次券成交人数']} 人"),
        ("体验客单价", f'"R$ "#,##0.00', m["体验客单价"], "营业额 ÷ 总客次"),
        ("叠加率", "0.00", m["叠加率"], "数量合计 ÷ 总客次"),
    ]
    for i, (label, fmt, value, sub) in enumerate(tiles):
        row = 4 + (i // 4) * 4
        col = 2 + (i % 4) * 4
        for rr in range(row, row + 3):
            ws.merge_cells(start_row=rr, start_column=col, end_row=rr, end_column=col + 2)
            for cc in range(col, col + 3):
                ws.cell(rr, cc).fill = TILE_FILL
        ws.cell(row, col, label).font = font(size=10, color=MUTED)
        v = ws.cell(row + 1, col, round(value, 2))
        v.font, v.number_format = font(size=18, bold=True), fmt
        ws.cell(row + 2, col, sub).font = font(size=8, color=MUTED)
        for rr in range(row, row + 3):
            ws.cell(rr, col).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[row + 1].height = 28

    ds, pos = write_chart_data(wb, rep)
    w0, w1 = pos["weeks"]
    s0, s1 = pos["staff"]
    c0, c1 = pos["categories"]
    staff_height = max(7.5, 0.75 * (s1 - s0) + 2.5)
    cat_height = max(7.5, 0.75 * (c1 - c0) + 2.5)
    ws.add_chart(make_chart(ds, w0, w1, [2, 3], "周度 权益金 vs 营业额（R$）", labels=True), "B13")
    ws.add_chart(make_chart(ds, w0, w1, [4, 5, 6], "周度客次构成（新客/会员/网红）", stacked=True), "J13")
    ws.add_chart(make_chart(ds, s0, s1, [2, 3], "员工 业绩 vs 营业额（R$）", horizontal=True,
                            height=staff_height), "B29")
    ws.add_chart(make_chart(ds, c0, c1, [2, 3, 4], "项目分类客次（按客户类型）", horizontal=True,
                            stacked=True, height=cat_height), "J29")
    below = 29 + int(max(staff_height, cat_height) / 0.53) + 2
    ws.add_chart(make_chart(ds, c0, c1, [5], "项目分类营业额（R$）", horizontal=True, labels=True,
                            height=cat_height), f"B{below}")
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
    write_category_sheet(wb, rep)

    ds = wb.create_sheet("指标说明")
    for j, h in enumerate(["指标", "计算口径", "本期数值"], 1):
        style_header(ds.cell(1, j, h))
    for i, (k, text) in enumerate(DEFINITIONS, 2):
        ds.cell(i, 1, k).font = font(bold=True)
        ds.cell(i, 2, text).font = font()
        ds.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
        if k in rep.month:
            style_value(ds.cell(i, 3, round(rep.month[k], 2)), k)
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


def fmt(v):
    return f"{v:,.2f}" if isinstance(v, float) else str(v)


def main(argv=None):
    ap = argparse.ArgumentParser(description="售出明细表 → 门店经营报表")
    ap.add_argument("input", help="收银系统导出的 Excel（含「售出明细表」）")
    ap.add_argument("-o", "--output", help="输出路径，默认 ./门店经营报表_<开始>_<结束>.xlsx")
    ap.add_argument("--month", help="只统计某个月，例如 2026-10")
    ap.add_argument("--weeks", choices=["fixed", "natural"], default="fixed",
                    help="fixed: 1-7日=W1…（默认）；natural: 周一~周日")
    args = ap.parse_args(argv)

    df, sheet, months = load(args.input, args.month)
    rep = build(df, args.weeks)
    notes = checks(df, months, args.month, sheet, rep.staff)
    out = Path(args.output or f"门店经营报表_{rep.start:%Y%m%d}_{rep.end:%Y%m%d}.xlsx")
    write_report(out, df, rep, notes, Path(args.input).name, args.weeks)

    print(f"已生成：{out}")
    print(f"区间：{rep.start} ~ {rep.end}")
    print("月度总数据：" + "，".join(f"{k} {fmt(rep.month[k])}" for k in MONTH_METRICS))
    for label, m in rep.weeks:
        if m:
            print(f"  {label}：权益金 {m['权益金']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
                  f"成交人数 {m['成交人数']}，二次券 {m['二次券成交数']}")
    print("员工达成（月度｜各周业绩）：")
    for name, m, weeks in rep.staff:
        by_week = " / ".join(f"{label.split()[0]} {wm['业绩']:,.0f}" for label, wm in weeks)
        print(f"  {name}：业绩 {m['业绩']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
              f"成交人数 {m['成交人数']}，二次券成交人数 {m['二次券成交人数']}｜{by_week}")
    print("项目分类（总客次｜新客/会员/网红｜营业额）：")
    total = rep.categories["总客次"].sum()
    for cat, x in rep.categories.iterrows():
        print(f"  {cat}：{x['总客次']:.0f}（{x['总客次'] / total:.0%}）｜{x['新客客次']:.0f}/{x['会员客次']:.0f}/"
              f"{x['网红/邀约']:.0f}｜{x['营业额']:,.2f}")
    for line in notes:
        print(line)


if __name__ == "__main__":
    sys.exit(main())

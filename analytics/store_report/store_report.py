#!/usr/bin/env python3
"""门店经营报表：读取收银系统导出的「售出明细表」，生成 月度总数据 / 周度数据 / 员工达成 报表。

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
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
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
UNASSIGNED = "未指定员工（线上购券）"

METRIC_FORMATS = {
    "权益金": "#,##0.00", "营业额": "#,##0.00", "成交客单": "#,##0", "体验客单价": "#,##0.00",
    "叠加率": "0.00", "套餐成交金额": "#,##0.00", "业绩": "#,##0.00",
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


def build(df, week_mode):
    month_m = compute(df)
    start, end = df["日期"].min().date(), df["日期"].max().date()

    weeks = []
    for label, ws, we in week_ranges(start.year, start.month, week_mode):
        rows = df[(df["日期"] >= pd.Timestamp(ws)) & (df["日期"] <= pd.Timestamp(we))]
        text = f"{label} {ws:%m/%d}-{we:%m/%d}"
        if ws > end:
            weeks.append((text, None))
            continue
        if we > end:
            text += f"（截至{end:%m/%d}）"
        weeks.append((text, compute(rows)))

    staff = []
    for name, rows in df.groupby(df["员工"].fillna(UNASSIGNED)):
        staff.append((name, compute(rows)))
    staff.sort(key=lambda x: (x[0] == UNASSIGNED, -x[1]["业绩"], -x[1]["营业额"]))
    staff = [s for s in staff if s[0] != UNASSIGNED or any(s[1][k] for k in STAFF_METRICS)]
    return month_m, weeks, staff, start, end


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
    for name, m in staff:
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


# ---------------------------------------------------------------- 写 Excel

FONT = "微软雅黑"
THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
HELPER_FILL = PatternFill("solid", fgColor="F2F2F2")
CENTER = Alignment(horizontal="center", vertical="center")


def font(**kw):
    return Font(name=FONT, size=kw.pop("size", 10), **kw)


def style_header(cell, helper=False):
    cell.font = font(bold=not helper, italic=helper, color="7F7F7F" if helper else "000000")
    cell.fill = HELPER_FILL if helper else HEAD_FILL
    cell.border = BORDER
    cell.alignment = CENTER


def style_value(cell, metric, helper=False):
    cell.font = font(color="7F7F7F" if helper else "000000", italic=helper)
    cell.border = BORDER
    cell.number_format = METRIC_FORMATS.get(metric, "0")
    if helper:
        cell.fill = HELPER_FILL


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


def write_report(path, df, month_m, weeks, staff, start, end, notes, source_name, week_mode):
    wb = Workbook()
    ws = wb.active
    ws.title = "经营报表"
    col0 = 3  # C 列开始放指标，与手工模板一致

    ws["A1"] = f"门店经营数据  {start:%Y-%m-%d} ~ {end:%Y-%m-%d}"
    ws["A1"].font = font(size=14, bold=True)
    ws["A2"] = (f"数据来源：{source_name}　生成时间：{datetime.now():%Y-%m-%d %H:%M}　"
                f"灰色「计算用」列为公式的底数；口径见「指标说明」")
    ws["A2"].font = font(size=9, color="7F7F7F")

    # 月度总数据
    r = 4
    ws.cell(r, 1, "月度总数据").font = font(bold=True)
    write_metric_header(ws, r, col0)
    write_metric_row(ws, r + 1, col0, month_m)

    # 周度数据
    r = 7
    write_metric_header(ws, r, col0)
    ws.cell(r + 1, 1, "周度数据").font = font(bold=True)
    for i, (label, m) in enumerate(weeks):
        row = r + 1 + i
        c = ws.cell(row, 2, label)
        c.font, c.border = font(), BORDER
        ws.cell(row, 1).border = BORDER
        write_metric_row(ws, row, col0, m)

    # 员工达成
    r = r + 1 + len(weeks) + 2
    for i, k in enumerate(STAFF_METRICS):
        style_header(ws.cell(r, 4 + i, k))
    ws.cell(r + 1, 1, "员工达成").font = font(bold=True)
    ws.cell(r + 1, 2, "月度数据").font = font(bold=True)
    for i, (name, m) in enumerate(staff):
        row = r + 1 + i
        c = ws.cell(row, 3, name)
        c.font, c.border = font(italic=name == UNASSIGNED), BORDER
        for j, k in enumerate(STAFF_METRICS):
            style_value(ws.cell(row, 4 + j, round(m[k], 2)), k)
    total_row = r + 1 + len(staff)
    c = ws.cell(total_row, 3, "合计")
    c.font, c.border = font(bold=True), BORDER
    for j, k in enumerate(STAFF_METRICS):
        letter = get_column_letter(4 + j)
        cell = ws.cell(total_row, 4 + j, f"=SUM({letter}{r + 1}:{letter}{total_row - 1})")
        style_value(cell, k)
        cell.font = font(bold=True)
    ws.cell(total_row + 1, 3, "注：同一客户找过多位员工成交时会各算一次，故人数合计可能大于月度去重人数。").font = \
        font(size=9, color="7F7F7F")

    helper_start = col0 + len(MONTH_METRICS) + 1
    for i in range(col0, helper_start + len(HELPER_METRICS)):
        ws.column_dimensions[get_column_letter(i)].width = 20 if i >= helper_start else 14
    for letter, w in {"A": 12, "B": 26, "C": 26}.items():
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "C1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 指标说明
    ds = wb.create_sheet("指标说明")
    for j, h in enumerate(["指标", "计算口径", "本期数值"], 1):
        style_header(ds.cell(1, j, h))
    for i, (k, text) in enumerate(DEFINITIONS, 2):
        ds.cell(i, 1, k).font = font(bold=True)
        ds.cell(i, 2, text).font = font()
        ds.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
        if k in month_m:
            style_value(ds.cell(i, 3, round(month_m[k], 2)), k)
        for j in (1, 2, 3):
            ds.cell(i, j).border = BORDER
    ds.cell(len(DEFINITIONS) + 2, 1, f"周度划分方式：{'自然周(周一~周日)' if week_mode == 'natural' else '每月按 7 天一段'}").font = \
        font(size=9, color="7F7F7F")
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

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


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


# ---------------------------------------------------------------- 入口

def main(argv=None):
    ap = argparse.ArgumentParser(description="售出明细表 → 门店经营报表")
    ap.add_argument("input", help="收银系统导出的 Excel（含「售出明细表」）")
    ap.add_argument("-o", "--output", help="输出路径，默认 ./门店经营报表_<开始>_<结束>.xlsx")
    ap.add_argument("--month", help="只统计某个月，例如 2026-10")
    ap.add_argument("--weeks", choices=["fixed", "natural"], default="fixed",
                    help="fixed: 1-7日=W1…（默认）；natural: 周一~周日")
    args = ap.parse_args(argv)

    df, sheet, months = load(args.input, args.month)
    month_m, weeks, staff, start, end = build(df, args.weeks)
    notes = checks(df, months, args.month, sheet, staff)
    out = Path(args.output or f"门店经营报表_{start:%Y%m%d}_{end:%Y%m%d}.xlsx")
    write_report(out, df, month_m, weeks, staff, start, end, notes, Path(args.input).name, args.weeks)

    print(f"已生成：{out}")
    print(f"区间：{start} ~ {end}")
    print("月度总数据：" + "，".join(
        f"{k} {month_m[k]:,.2f}" if isinstance(month_m[k], float) else f"{k} {month_m[k]}"
        for k in MONTH_METRICS))
    for label, m in weeks:
        if m:
            print(f"  {label}：权益金 {m['权益金']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
                  f"成交人数 {m['成交人数']}，二次券 {m['二次券成交数']}")
    print("员工达成：")
    for name, m in staff:
        print(f"  {name}：业绩 {m['业绩']:,.2f}，营业额 {m['营业额']:,.2f}，总客次 {m['总客次']}，"
              f"成交人数 {m['成交人数']}，二次券成交人数 {m['二次券成交人数']}")
    for line in notes:
        print(line)


if __name__ == "__main__":
    sys.exit(main())

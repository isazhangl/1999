#!/usr/bin/env python3
"""门店经营报表：读取收银系统导出的「售出明细表」，生成 看板 / 月度总数据 / 周度数据 / 员工达成 / 项目分类 报表。

用法:
    python3 store_report.py 售出明细.xlsx                      # 输出到当前目录
    python3 store_report.py 售出明细.xlsx -o 报表.xlsx
    python3 store_report.py 售出明细.xlsx --month 2026-10      # 指定报表月份（默认取文件里数据最多的月份）
    python3 store_report.py 售出明细.xlsx --weeks fixed        # 周按 1-7日=W1… 划分（默认 周一~周日，跨月的周算满 7 天）
    python3 store_report.py 售出明细.xlsx --targets 目标.json   # 指定目标文件（默认 targets/<门店>/<YYYY-MM>.json）
    python3 store_report.py 售出明细.xlsx --store Morumbi      # 指定门店（默认按「权益归属门店」自动识别）

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
    "代金券", "开业券 Voucher comprado online", "套餐", "免单类", "品项",
    "现金", "信用卡 Crédito à vista", "信用卡分期 Crédito parcelado", "银行卡 Débito", "PIX",
}

MONTH_METRICS = ["权益金", "营业额", "总客次", "会员客次", "新客客次", "网红/邀约",
                 "成交人数", "新客成交人数", "成交客单", "二次券成交数", "体验客单价", "叠加率", "新客成交率",
                 "续卡人数", "续卡金额", "会员复购率", "人效"]
HELPER_METRICS = ["套餐成交金额", "项目次数", "护理师人天", "会员到店人数"]
STAFF_METRICS = ["业绩", "营业额", "总客次", "成交人数", "新客成交人数", "二次券成交人数", "新客成交率"]
STAFF_HELPERS = ["新客客次"]
CUSTOMER_TYPES = ["新客客次", "会员客次", "网红/邀约"]
MIX_TYPES = ["新客项目次数", "会员项目次数", "网红/邀约项目次数"]   # 项目分类页按项目次数统计
VISIT_TYPE_COL = {"会员": "会员项目次数", "网红": "网红/邀约项目次数", "新客": "新客项目次数"}
UNASSIGNED = "未指定员工"
# 权益金拆解（用户 2026-10-25 定）：卖套餐 / 二次券 / 体验 / 线上卖券，四项相加 = 权益金
BREAKDOWN = ["卖套餐", "二次券", "体验", "线上卖券"]
BREAKDOWN_COLS = {"卖套餐": "套餐额", "二次券": "二次券额", "体验": "体验额", "线上卖券": "线上额"}
TARGETS_DIR = Path(__file__).resolve().parent / "targets"  # targets/<门店>/<YYYY-MM>.json
RATE_METRICS = {"新客成交率", "体验客单价", "成交客单", "叠加率", "人效", "会员复购率"}  # 比率型目标：任何时间段的目标都一样
METRIC_ALIASES = {"客单价": "体验客单价"}
WEEKDAYS = "一二三四五六日"
UNCATEGORIZED = "（未分类）"

METRIC_FORMATS = {
    "权益金": "#,##0.00", "营业额": "#,##0.00", "成交客单": "#,##0", "体验客单价": "#,##0.00",
    "叠加率": "0.00", "套餐成交金额": "#,##0.00", "业绩": "#,##0.00", "平均单价": "#,##0.00",
    **{k: "#,##0.00" for k in BREAKDOWN},
    "人效": "0.00", "日均上班护理师": "0.0", "会员复购率": "0.00%", "续卡金额": "#,##0.00",
    "客次占比": "0.0%", "项目次数占比": "0.0%", "营业额占比": "0.0%", "新客成交率": "0.00%", "月达成率": "0.0%", "累计达成率": "0.0%",
}

# ---------------------------------------------------------------- 中葡双语
# 表头、固定标签在保存前统一翻译（见 translate_workbook）；带数字/日期的句子在生成时直接写双语。
PT = {
    # 指标
    "权益金": "Vendas Pacotes/Vouchers", "营业额": "Receita de Serviços", "总客次": "Total de Atendimentos",
    "会员客次": "Atend. Clientes Pacote", "新客客次": "Atend. Novos Clientes", "网红/邀约": "Influencers/Convites",
    "成交人数": "Fechamentos de Pacote", "成交客单": "Ticket Médio Pacote", "二次券成交数": "Vouchers Retorno Vendidos",
    "体验客单价": "Ticket Médio Atendimento", "叠加率": "Itens por Atendimento", "新客成交率": "Conversão Novos Clientes",
    "业绩": "Vendas", "二次券成交人数": "Clientes Voucher Retorno", "套餐成交金额": "Valor em Pacotes",
    "数量合计": "Total de Itens", "客户人数": "Clientes Únicos", "客次占比": "% Atendimentos",
    "项目次数": "Serviços Realizados", "项目次数占比": "% Serviços", "新客项目次数": "Serviços Novos Clientes",
    "会员项目次数": "Serviços Clientes Pacote", "网红/邀约项目次数": "Serviços Influencers/Convites",
    "按分类 · 周度项目次数": "Serviços por Categoria e Semana",
    "营业额占比": "% Receita", "平均单价": "Preço Médio",
    "卖套餐": "Pacotes", "二次券": "Vouchers de Retorno", "体验": "Experiência", "线上卖券": "Vouchers Online",
    "权益金拆解": "Composição das Vendas", "占比": "% do Total",
    "新客成交人数": "Fechamentos Novos Clientes", "续卡人数": "Renovações de Pacote", "续卡金额": "Valor de Renovação",
    "会员复购率": "Recompra Clientes Pacote", "会员到店人数": "Clientes Pacote Atendidos",
    "人效": "Atend. por Terapeuta/Dia", "护理师人天": "Terapeuta-Dias", "营业天数": "Dias com Atendimento",
    "日均上班护理师": "Terapeutas por Dia", "上班护理师": "Terapeutas no Dia",
    "当天做项目的护理师": "Terapeutas que Atenderam", "月度 / 周度人效": "Produtividade por Mês e Semana",
    "每日人效": "Produtividade por Dia",
    # 表头 / 标签
    "分类": "Categoria", "项目名称": "Serviço", "月度合计": "Total do Mês", "合计": "Total", "员工": "Colaborador(a)",
    "期间": "Período", "周": "Semana", "日期": "Data", "星期": "Dia", "指标": "Indicador", "名称": "Item",
    "数量": "Qtd.", "单据编号": "Nº Documento", "客户名称": "Cliente", "支付方式": "Forma de Pagamento",
    "金额": "Valor", "计算口径": "Como é calculado", "本期数值": "Valor do Período",
    "数据检查 / 提示": "Verificação / Avisos", "月目标": "Meta do Mês", "实际": "Realizado",
    "月达成率": "% Meta do Mês", "同期目标": "Meta até Hoje", "同期达成率": "% Meta até Hoje",
    "剩余需完成": "Falta Atingir", "剩余日均需完成": "Necessário por Dia", "目标进度": "Progresso Esperado",
    "剩余营业日": "Dias Abertos Restantes", "达成率": "% Atingido", "累计达成率": "% Acumulado",
    "应达进度": "Progresso Esperado", "未指定员工": "Sem colaborador(a)",
    "（未分类）": "(Sem categoria)",
    # 分区标题
    "月度总数据": "Resumo do Mês", "周度数据": "Dados Semanais", "员工达成": "Desempenho da Equipe",
    "月度数据": "Mensal", "员工周度数据": "Equipe por Semana", "月度目标达成": "Metas do Mês",
    "周度目标达成": "Metas por Semana", "每日目标 vs 实际": "Meta Diária vs Realizado",
    "按分类（月度）": "Por Categoria (Mês)", "按分类 · 周度客次": "Atendimentos por Categoria e Semana",
    "按分类 · 周度营业额": "Receita por Categoria e Semana", "按项目（月度）": "Por Serviço (Mês)",
    # 指标说明里的条目名
    "员工·业绩": "Equipe · Vendas", "员工·营业额/总客次": "Equipe · Receita / Atendimentos",
    "员工·成交人数": "Equipe · Fechamentos", "员工·二次券成交人数": "Equipe · Clientes Voucher Retorno",
    "员工·新客成交率": "Equipe · Conversão Novos Clientes", "项目分类": "Categorias de Serviço",
    "同期目标 / 同期达成率": "Meta até Hoje / % Meta até Hoje", "续卡人数 / 续卡金额": "Renovações / Valor", "周目标 / 日目标": "Meta Semanal / Diária",
    "周度划分": "Divisão das Semanas",
}
PT_SUFFIX = {"累计目标": "Meta Acumulada", "累计实际": "Realizado Acumulado", "累计": "Acumulado",
             "目标": "Meta", "实际": "Realizado"}
PT_WEEKDAYS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
PT_MONTHS = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
SHEETS = {"看板": "看板 Painel", "经营报表": "经营报表 Relatório", "目标达成": "目标达成 Metas",
          "项目分类": "项目分类 Categorias", "项目明细": "项目明细 Serviços", "指标说明": "指标说明 Definições",
          "套餐成交明细": "套餐成交明细 Pacotes", "二次券成交明细": "二次券成交明细 Vouchers",
          "数据检查": "数据检查 Verificação", "图表数据": "图表数据 Dados Gráficos", "人效": "人效 Produtividade"}


def pt_of(text):
    """中文标签 → 葡语；支持「计算用:指标」「指标+目标/实际/累计…」这类组合，查不到返回 None。"""
    if text in PT:
        return PT[text]
    if text.startswith("计算用:") and text[4:] in PT:
        return f"Base: {PT[text[4:]]}"
    for suffix, word in PT_SUFFIX.items():
        if text.endswith(suffix) and text[:-len(suffix)] in PT:
            return f"{PT[text[:-len(suffix)]]} – {word}"
    return None


def dm(d):
    """日期写成「18/out」：中葡读者都不会把日和月看反。"""
    return f"{d.day:02d}/{PT_MONTHS[d.month - 1]}"


def bi(zh, pt):
    return f"{zh} | {pt}"


DEFINITIONS = [
    ("权益金", "卖套餐 + 二次券 + 体验 + 线上卖券（2026-10-25 起含线上开业券）。即门店当期卖套餐、卖券、做项目直接收的钱，"
             "加上客人线上买券的钱；用代金券/套餐抵扣的部分不算，退款为负数自动抵扣。"
             "比收银系统里的权益金多出「做项目时直接付钱」的部分（系统只统计类型=权益），差额见「数据检查」。"),
    ("卖套餐", "类型=权益、分类=套餐 的「支付金额」，不含用代金券/套餐抵扣的部分。"),
    ("人效", "总客次 ÷ 护理师人天（客次/人/天）。护理师人天 = 每天有做项目的员工人数相加，来自售出明细的「员工」，不是排班表："
            "上班但当天一个客人都没做的人不算，所以实际人效会略低。每日 / 每周 / 每月明细见「人效」页。"),
    ("二次券", "店里卖出、当天没有用掉的券（类型=权益、分类=代金券，现金/卡/PIX 支付）；券的退款按负数记在这里。"),
    ("体验", "做项目时直接付钱（类型=项目、现金/卡/PIX 支付）＋ 店里现买、当天就用掉的券。当天是否用掉按同一客户同一天估算："
            "当天用券做的项目先抵之前买的券、再抵当天登记的线上券，剩下的才抵店里现买的券。"),
    ("线上卖券", "支付方式 = 开业券（客人线上已付）的 权益行和项目行。线上卖券没有销售员工，不算员工业绩。"),
    ("营业额", "类型=项目 的「支付金额」合计（含代金券/套餐/免单核销的项目价值）。"),
    ("总客次", "到店次数：同一客户同一天来店算 1 客次，不管做了几个项目（2026-10-25 起；之前是按项目行数算）。"
              "整次到店都退款的不算。"),
    ("会员客次", "会员的到店次数：那次到店有套餐核销，或客户编号是字母+数字（如 B14，会员用券或直接付钱也算会员）。"),
    ("新客客次", "总客次 − 会员客次 − 网红/邀约（表内公式），即客户编号是纯数字、没用套餐也没免单的到店。"),
    ("网红/邀约", "那次到店有免单类付款（没有套餐核销）的到店次数。"),
    ("成交人数", "类型=权益、分类=套餐，现金类支付合计>0 的客户人数（按客户去重）。"),
    ("成交客单", "套餐成交金额 ÷ 成交人数（表内公式）。套餐成交金额 = 类型=权益、分类=套餐 的「现金类支付」合计。"),
    ("二次券成交数", "「二次券」的张数：店里现场掏钱买的代金券（现金/卡/PIX，或组合支付里有现金部分），去掉当天就用掉的（算体验）。"),
    ("体验客单价", "营业额 ÷ 总客次（表内公式）。"),
    ("叠加率", "项目次数 ÷ 总客次（表内公式）：平均每次到店做几个项目。项目次数 = 类型=项目 的行数（组合支付的拆分行不算，退款扣减）。"),
    ("新客成交人数", "买套餐的客户里，客户编号是纯数字的（新客）人数。"),
    ("新客成交率", "新客成交人数 ÷ 新客客次 × 100%（表内公式；2026-10-25 起分子只算新客，会员续卡算进会员复购率）。新客客次为 0 时留空。"),
    ("续卡人数 / 续卡金额", "买套餐的客户里，客户编号是字母+数字（会员）的人数和套餐金额（现金类支付）。"),
    ("会员复购率", "有效会员复购率 = 统计周期内复购（续卡）的会员数 ÷ 统计周期内到店的会员总客数（去重；续卡但没做项目的会员也算到店）。"),
    ("员工·业绩", "该员工名下的 卖套餐 + 二次券 + 体验（不含线上卖券）。明细里「实业绩」列全为 0，故不用。"),
    ("员工·营业额/总客次", "该员工名下的项目营业额 / 该员工服务的到店次数（一次到店有两位员工服务时，两人各算 1 次）。"),
    ("员工·成交人数", "该员工名下买套餐的客户人数（去重）。"),
    ("员工·二次券成交人数", "该员工名下买「二次券」的客户人数（去重，当天就用掉的券不算）。"),
    ("员工·新客成交率", "该员工名下新客成交人数 ÷ 该员工服务的新客客次（灰色「计算用:新客客次」列）。"
                     "只卖不做项目的员工新客客次为 0，成交率留空。"),
    ("项目分类", "类型=项目 的行按收银系统「分类」汇总，按项目次数算（一次到店做两个分类的项目各算 1 次）；"
                "新客/会员/网红 按那次到店的类型；客户人数=做过该分类项目的去重客户数；平均单价=营业额÷项目次数。"),
    ("月达成率", "实际 ÷ 月目标。目标来自 targets/<YYYY-MM>.json（用户每月提供）。"),
    ("目标进度", "按「每日目标」累计到报表最后一天占全月的比例（休息日权重为 0），比单纯按天数算更准。"),
    ("同期目标 / 同期达成率", "累加型指标（权益金、营业额、总客次等）：同期目标 = 月目标 × 目标进度；"
                       "比率型指标（新客成交率、体验客单价）：同期目标 = 月目标。同期达成率 ≥ 100% 记 ✓，否则 ✗。"),
    ("周目标 / 日目标", "累加型指标按「每日目标」的比例拆到每天，再按周相加；比率型指标每周、每天的目标与月目标相同。"),
    ("周度划分", "默认：周一~周日（按「账务日期」）。W1 = 包含 1 号的那一周，从那周的周一开始，跨月的周算满 7 天"
                "（例如 2026-10 的 W1 = 9/28–10/4）；月度、员工月度、每日和目标只算本月。"
                "导出明细时从 W1 的周一开始导，第一周才完整。可用 --weeks fixed 改回 1-7日=W1…"),
]


DEFINITIONS_PT = {
    "权益金": "Pacotes + vouchers de retorno + experiência + vouchers online (desde 25/10/2026 inclui os vouchers comprados online, 开业券). É o dinheiro recebido por pacotes, vouchers e serviços pagos na hora, mais as compras online; o que foi pago com voucher/pacote não conta e estornos entram negativos. Fica acima do valor do sistema pela parte de serviços pagos na hora (o sistema só soma linhas 权益); a diferença aparece em «Verificação».",
    "卖套餐": "«Valor pago» das linhas 权益 da categoria 套餐 (pacotes), sem a parte paga com voucher/pacote.",
    "人效": "Total de atendimentos ÷ terapeuta-dias (atendimentos por terapeuta por dia). Terapeuta-dias = soma, dia a dia, das colaboradoras que fizeram ao menos um serviço (pela coluna «员工» das vendas, não pela escala): quem trabalhou e não atendeu ninguém não conta, então a produtividade real é um pouco menor. Detalhe por dia/semana/mês na aba «人效 Produtividade».",
    "二次券": "Vouchers vendidos na loja (dinheiro/cartão/PIX) e não usados no mesmo dia; estornos de voucher entram aqui negativos.",
    "体验": "Serviços pagos na hora (linhas 项目 pagas com dinheiro/cartão/PIX) + vouchers comprados na loja e usados no mesmo dia. «Mesmo dia» é estimado por cliente e dia: os serviços pagos com voucher consomem primeiro vouchers comprados antes, depois vouchers online do dia, e só então os comprados na loja.",
    "线上卖券": "Linhas 权益 e 项目 pagas com 开业券 (voucher comprado online). Não têm vendedora, por isso não entram nas vendas da equipe.",
    "营业额": "Soma do «valor pago» das linhas tipo 项目 (serviços), incluindo serviços pagos com voucher, pacote ou cortesia.",
    "总客次": "Número de visitas: o mesmo cliente no mesmo dia = 1 atendimento, não importa quantos serviços fez (desde 25/10/2026; antes contava cada serviço). Visita totalmente estornada não conta.",
    "会员客次": "Visitas de clientes de pacote: houve uso de pacote na visita ou o código do cliente tem letra + número (ex.: B14) — cliente de pacote que pagou com voucher ou na hora também conta aqui.",
    "新客客次": "Total de atendimentos − atendimentos de pacote − influencers/convites (fórmula na planilha).",
    "网红/邀约": "Visitas com pagamento de cortesia (免单类) e sem uso de pacote.",
    "成交人数": "Número de clientes (sem repetir) que compraram pacote com pagamento em dinheiro/cartão/PIX > 0.",
    "成交客单": "Valor em pacotes ÷ fechamentos de pacote (fórmula). Valor em pacotes = «pagamento em dinheiro» das vendas de pacote.",
    "二次券成交数": "Quantidade de «vouchers de retorno»: vendidos na loja (dinheiro/cartão/PIX ou parte em dinheiro), sem os usados no mesmo dia (esses contam como experiência).",
    "体验客单价": "Receita de serviços ÷ total de atendimentos (fórmula).",
    "叠加率": "Serviços ÷ visitas (fórmula): quantos serviços por visita, em média. Serviços = linhas tipo 项目 (sem linhas de pagamento dividido; estornos descontam).",
    "新客成交人数": "Clientes que compraram pacote com código só de números (clientes novos).",
    "新客成交率": "Fechamentos de novos clientes ÷ visitas de novos clientes × 100% (fórmula; desde 25/10/2026 só conta novos — renovações de clientes de pacote vão para a recompra). Fica em branco se não houver visitas de novos clientes.",
    "续卡人数 / 续卡金额": "Clientes de pacote (código com letra + número) que compraram pacote de novo, e o valor pago.",
    "会员复购率": "Recompra de clientes de pacote = clientes de pacote que renovaram no período ÷ clientes de pacote atendidos no período (sem repetir; quem renovou sem fazer serviço também conta).",
    "员工·业绩": "Pacotes + vouchers de retorno + experiência no nome da colaboradora (sem vouchers online). A coluna 实业绩 do sistema está toda zerada, por isso não é usada.",
    "员工·营业额/总客次": "Receita de serviços / visitas atendidas pela colaboradora (se duas atenderam a mesma visita, conta para as duas).",
    "员工·成交人数": "Clientes (sem repetir) que compraram pacote com a colaboradora.",
    "员工·二次券成交人数": "Clientes (sem repetir) que compraram voucher de retorno com a colaboradora (vouchers usados no mesmo dia não contam).",
    "员工·新客成交率": "Fechamentos de novos clientes da colaboradora ÷ visitas de novos clientes que ela atendeu (coluna cinza «Base»). Quem só vende e não atende fica em branco.",
    "项目分类": "Serviços (linhas tipo 项目) por «categoria» do sistema, contados por serviço (uma visita com duas categorias conta em cada uma); novos/pacote/influencers pelo tipo da visita; clientes únicos = clientes sem repetir; preço médio = receita ÷ serviços.",
    "月达成率": "Realizado ÷ meta do mês. Metas vêm de targets/<AAAA-MM>.json (enviadas todo mês).",
    "目标进度": "Quanto da meta do mês deveria estar feito até o último dia do relatório, somando as metas diárias (dias fechados = 0).",
    "同期目标 / 同期达成率": "Indicadores somáveis (vendas, receita, atendimentos): meta até hoje = meta do mês × progresso esperado. Indicadores de taxa (conversão, ticket médio): meta até hoje = meta do mês. ≥ 100% = ✓, abaixo = ✗.",
    "周目标 / 日目标": "Indicadores somáveis são divididos por dia na proporção das metas diárias e somados por semana; indicadores de taxa têm a mesma meta todo dia/semana.",
    "周度划分": "Padrão: semana de segunda a domingo (pela data contábil). W1 = a semana que contém o dia 1, desde a segunda-feira; semanas que cruzam o mês contam os 7 dias (ex.: em 2026-10, W1 = 28/09–04/10). Mês, equipe, dias e metas contam só o mês do relatório. Exporte a partir da segunda-feira da W1 para a 1ª semana ficar completa. Com --weeks fixed volta a dias 1-7 = W1…",
}


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
    m = method if isinstance(method, str) else ""   # 退款单里有支付方式为空的行（读进来是 NaN）
    if "开业券" in m:
        return "开业券"
    if "免单" in m:
        return "免单"
    if m.startswith("代金券") or m.startswith("品项"):  # 品项 = 体验券里包含的单个项目的核销
        return "代金券"
    if m.startswith("套餐"):
        return "套餐"
    return "现金类"


def detect_store(df):
    """按「权益归属门店」出现最多的值识别门店：「São Paulo - Unidade Morumbi」→「Morumbi」。"""
    if "权益归属门店" not in df.columns or df["权益归属门店"].dropna().empty:
        return None
    full = df["权益归属门店"].dropna().map(clean_text).mode().iloc[0]
    return full.split("Unidade", 1)[-1].strip() if "Unidade" in full else full


def classify_rights(df):
    """权益金拆解，每行记到 套餐额 / 二次券额 / 体验额 / 线上额 四列（compute 直接相加）：
    - 线上卖券：支付方式 = 开业券（客人线上已付）的 权益行和项目行；
    - 卖套餐：分类 = 套餐 的权益行，用代金券 / 套餐抵扣的部分不算（不是新收的钱）；
    - 体验：做项目时直接付钱（项目行、现金类支付）＋ 店里现买、当天就用掉的券；
    - 二次券：店里卖出、当天没用掉的券（券的退款按负数记在这里）。
    「当天用掉」按客户逐日估算：当天用券做的项目先抵之前买的券，再抵当天登记的线上券，剩下的才抵店里现买的券
    （按操作时间先后）。文件开始前买的券看不到，所以老客当天先用旧券再买新券时会被算成体验。
    另记 二次券数 / 体验券数（张），用于二次券成交数。"""
    right = df["类型"] == "权益"
    paid_new = ~df["支付归类"].isin(["开业券", "代金券", "套餐"])   # 店里新收的钱
    online = df["支付归类"] == "开业券"
    vouchers = right & (df["分类"] != "套餐")
    sold = vouchers & (df["分类"] == "代金券") & ((df["支付归类"] == "现金类") | (df["现金类支付"] > 0))
    online_v = vouchers & online & (df["数量"] > 0)
    redeem = (df["类型"] == "项目") & (df["支付归类"] == "代金券") & (df["客次"] > 0)
    order = df["操作时间"].astype(str) if "操作时间" in df.columns else pd.Series(df.index.astype(str), index=df.index)

    frac = pd.Series(0.0, index=df.index)   # 每行算作「体验」的比例
    pool = df[df["日期"].notna() & ((sold & (df["数量"] > 0)) | online_v | redeem)]
    for _, rows in pool.groupby("客户"):
        stock = 0.0   # 之前买了还没用的券
        for _, d in rows.groupby("日期"):
            left = d.loc[redeem[d.index], "客次"].sum()
            from_stock = min(stock, left)
            left -= from_stock
            online_qty = d.loc[online_v[d.index], "数量"].sum()
            from_online = min(online_qty, left)
            left -= from_online
            bought = taken = 0.0
            cash = d[sold[d.index] & (d["数量"] > 0)]
            for idx in order[cash.index].sort_values().index:
                q = df.at[idx, "数量"]
                take = min(q, left)
                frac[idx] = take / q
                left -= take
                bought += q
                taken += take
            stock += -from_stock + (online_qty - from_online) + (bought - taken)
    if "单据编号" in df.columns:  # 组合支付拆出来的行跟原行走
        main = df[vouchers & ~df["拆分行"]].assign(_f=frac).drop_duplicates(["单据编号", "名称"])
        keys = pd.MultiIndex.from_frame(df[["单据编号", "名称"]])
        mapped = pd.Series(main.set_index(["单据编号", "名称"])["_f"].reindex(keys).values, index=df.index)
        frac = frac.where(~(vouchers & df["拆分行"]), mapped.fillna(0.0))

    amount = df["支付金额"]
    voucher_money = amount.where(vouchers & paid_new, 0.0)
    qty = df["数量"].where(sold, 0)
    df["线上额"] = amount.where(online & df["类型"].isin(["权益", "项目"]), 0.0)
    df["套餐额"] = amount.where(right & (df["分类"] == "套餐") & paid_new, 0.0)
    df["体验额"] = amount.where((df["类型"] == "项目") & (df["支付归类"] == "现金类"), 0.0) + voucher_money * frac
    df["二次券额"] = voucher_money * (1 - frac)
    df["体验券数"] = qty * frac
    df["二次券数"] = qty * (1 - frac)
    return df


def is_member_code(codes):
    """会员编号是字母+数字（如 B14），新客是纯数字（用户 2026-10-25 定）。"""
    return codes.astype(str).str.contains(r"[A-Za-z]", regex=True)


def classify_visits(df):
    """客次 = 到店次数：同一客户同一天来店算 1 次，不管做几个项目（做了几个项目看叠加率 = 项目次数 ÷ 客次）。
    每次到店的类型：有套餐核销 = 会员；否则有免单 = 网红/邀约；否则看客户编号——会员编号是字母+数字（如 B14），
    新客是纯数字（用户 2026-10-25 定）。"""
    proj = (df["类型"] == "项目") & df["日期"].notna()
    who = df["客户"].astype(object).where(df["客户"].notna(), "第" + df.index.astype(str) + "行")
    df["到店键"] = None
    df.loc[proj, "到店键"] = who[proj].astype(str) + "|" + df.loc[proj, "日期"].dt.strftime("%Y-%m-%d")
    p = df[proj]
    visit = p.groupby("到店键")
    has = lambda kind: visit["支付归类"].transform(lambda x: (x == kind).any())
    code = p["客户编号"].astype(str) if "客户编号" in p.columns else pd.Series("", index=p.index)
    member_code = is_member_code(code)
    kind = pd.Series("新客", index=p.index)
    kind[member_code] = "会员"
    kind[has("免单")] = "网红"
    kind[has("套餐")] = "会员"
    df["访客类型"] = None
    df.loc[p.index, "访客类型"] = kind
    df["到店有效"] = False
    df.loc[p.index, "到店有效"] = visit["客次"].transform("sum") > 0   # 整次到店都退款了就不算
    return df


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
    # 组合支付：系统把同一行的第二种付款拆成「数量 0」的行（员工常为空）。这种行不算客次，员工跟原行走。
    df["拆分行"] = (df["数量"] == 0) & ~is_refund
    if "单据编号" in df.columns:
        owner = (df[~df["拆分行"]].dropna(subset=["员工"])
                 .drop_duplicates(["单据编号", "名称"]).set_index(["单据编号", "名称"])["员工"])
        keys = pd.MultiIndex.from_frame(df[["单据编号", "名称"]])
        df["员工"] = df["员工"].fillna(pd.Series(owner.reindex(keys).values, index=df.index).where(df["拆分行"]))
    df["客次"] = df["客次符号"].where(~df["拆分行"], 0)
    cust_id = df["客户编号"].map(clean_text) if "客户编号" in df.columns else None
    df["客户"] = cust_id.fillna(df["客户名称"]) if cust_id is not None else df["客户名称"]
    classify_rights(df)
    classify_visits(df)

    months = sorted(df["日期"].dropna().dt.to_period("M").astype(str).unique())
    df.attrs["store"] = detect_store(df)
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
    visits = proj[(proj["客次"] > 0) & proj["到店有效"].astype(bool)].drop_duplicates("到店键")   # 见 classify_visits
    m = {k: df[col].sum() for k, col in BREAKDOWN_COLS.items()}   # 权益金拆解，见 classify_rights
    m.update({
        "权益金": sum(m[k] for k in BREAKDOWN),
        "营业额": proj["支付金额"].sum(),
        "总客次": len(visits),
        "会员客次": int((visits["访客类型"] == "会员").sum()),
        "网红/邀约": int((visits["访客类型"] == "网红").sum()),
        "项目次数": int(proj["客次"].sum()),
        "成交人数": paying_customers(package, "现金类支付"),
        "二次券成交数": int(round(df["二次券数"].sum())),
        "二次券成交人数": paying_customers(df[df["二次券数"] != 0], "二次券数"),
        "套餐成交金额": package["现金类支付"].sum(),
    })
    m["新客客次"] = m["总客次"] - m["会员客次"] - m["网红/邀约"]
    m["成交客单"] = m["套餐成交金额"] / m["成交人数"] if m["成交人数"] else 0
    m["体验客单价"] = m["营业额"] / m["总客次"] if m["总客次"] else 0
    m["叠加率"] = m["项目次数"] / m["总客次"] if m["总客次"] else 0
    # 买套餐的人按客户编号分：字母+数字 = 会员续卡，纯数字 = 新客成交
    codes = package["客户编号"] if "客户编号" in package.columns else pd.Series("", index=package.index)
    paid = package.assign(_编号=codes).groupby("客户").agg(现金=("现金类支付", "sum"), 编号=("_编号", "first"))
    paid = paid[paid["现金"] > 0]
    renew = is_member_code(paid["编号"])
    m["新客成交人数"] = int((~renew).sum())
    m["续卡人数"] = int(renew.sum())
    m["续卡金额"] = float(paid.loc[renew, "现金"].sum())
    members = set(visits.loc[visits["访客类型"] == "会员", "客户"]) | set(paid.index[renew])
    m["会员到店人数"] = len(members)
    m["会员复购率"] = m["续卡人数"] / m["会员到店人数"] if m["会员到店人数"] else None
    m["新客成交率"] = m["新客成交人数"] / m["新客客次"] if m["新客客次"] else None
    m["业绩"] = m["权益金"] - m["线上卖券"]   # 线上卖券没有销售员工，不算员工业绩
    # 人效 = 总客次 ÷ 护理师人天（每天有做项目的员工人数相加；来自售出明细，不是排班表）
    per_day = proj[proj["员工"].notna()].groupby(["日期", "员工"])["客次"].sum()
    m["护理师人天"] = int((per_day > 0).sum())
    m["营业天数"] = int((proj.groupby("日期")["客次"].sum() > 0).sum())
    m["人效"] = m["总客次"] / m["护理师人天"] if m["护理师人天"] else None
    return m


def mix_table(proj, keys):
    """项目行按 keys 汇总：项目次数 / 新客·会员·网红的项目次数（按到店类型）/ 客户人数 / 营业额，按项目次数降序。"""
    cols = ["项目次数", *MIX_TYPES, "客户人数", "营业额"]
    if proj.empty:
        return pd.DataFrame(columns=cols)
    sign = proj["客次"]
    p = proj.assign(项目次数=sign, **{col: sign.where(proj["访客类型"] == kind, 0) for kind, col in VISIT_TYPE_COL.items()})
    t = p.groupby(keys)[["项目次数", *MIX_TYPES, "支付金额"]].sum().rename(columns={"支付金额": "营业额"})
    per_customer = p.groupby([*keys, "客户"])["客次"].sum().gt(0)
    t["客户人数"] = per_customer.groupby(level=list(range(len(keys)))).sum()
    return t[cols].sort_values("项目次数", ascending=False)


def week_ranges(year, month, mode):
    """返回 [(标签, 开始日, 结束日)]，覆盖整个自然月。
    natural = 周一~周日，跨月的周算满 7 天（W1 从包含 1 号那周的周一开始，最后一周到包含月底那周的周日）；
    fixed = 1-7日=W1、8-14日=W2…，只在本月内。"""
    first = date(year, month, 1)
    last = date(year, month, calendar.monthrange(year, month)[1])
    if mode == "natural":
        bounds, s = [], first - timedelta(days=first.weekday())
        while s <= last:
            bounds.append((s, s + timedelta(days=6)))
            s += timedelta(days=7)
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
    context: object = None         # 周度用到的全部行（跨月的周会含上月/下月的天）
    data_span: tuple = None        # 周度数据实际覆盖的 (第一天, 最后一天)
    targets: object = None         # Targets 或 None
    neighbor_targets: dict = None  # 跨月的周用到的相邻月份目标 {"YYYY-MM": Targets}
    store: str = None              # 门店名，例如 Ibirapuera


def in_days(rows, first, last):
    return rows[(rows["日期"] >= pd.Timestamp(first)) & (rows["日期"] <= pd.Timestamp(last))]


def build(df, week_mode, context=None):
    """df = 报表月份的行；context = 整个文件（跨月的周用它补上月/下月的天，月度数据只看 df）。"""
    start, end = df["日期"].min().date(), df["日期"].max().date()
    defs = week_ranges(start.year, start.month, week_mode)
    ctx = in_days(df if context is None else context, defs[0][1], defs[-1][2])
    data_first, data_last = ctx["日期"].min().date(), ctx["日期"].max().date()
    week_defs = []
    for label, ws, we in defs:
        text = f"{label} {dm(ws)}-{dm(we)}"
        if ws < data_first <= we:
            text += f"（从/desde {dm(data_first)}）"
        if ws <= data_last < we:
            text += f"（截至/até {dm(data_last)}）"
        week_defs.append((text, ws, we))
    active_weeks = [w for w in week_defs if w[1] <= data_last]

    weeks = [(text, compute(in_days(ctx, ws, we)) if ws <= data_last else None) for text, ws, we in week_defs]

    staff = []
    ctx_staff = ctx["员工"].fillna(UNASSIGNED)
    for name, rows in df.groupby(df["员工"].fillna(UNASSIGNED)):
        wrows = ctx[ctx_staff == name]
        staff.append((name, compute(rows), [(text, compute(in_days(wrows, ws, we))) for text, ws, we in active_weeks]))
    staff.sort(key=lambda x: (x[0] == UNASSIGNED, -x[1]["业绩"], -x[1]["营业额"]))
    staff = [s for s in staff if s[0] != UNASSIGNED or any(s[1][k] for k in STAFF_METRICS)]

    proj = df[df["类型"] == "项目"].assign(分类=lambda x: x["分类"].fillna(UNCATEGORIZED))
    categories = mix_table(proj, ["分类"])
    items = mix_table(proj, ["分类", "名称"])
    order = {c: i for i, c in enumerate(categories.index)}
    items = items.sort_values(["项目次数"], ascending=False).sort_index(
        level=0, key=lambda idx: idx.map(order), sort_remaining=False, kind="stable")
    ctx_proj = ctx[ctx["类型"] == "项目"].assign(分类=lambda x: x["分类"].fillna(UNCATEGORIZED))
    category_weeks = {
        metric: pd.DataFrame({text: in_days(ctx_proj, ws, we).groupby("分类")[col].sum()
                              for text, ws, we in active_weeks}).reindex(categories.index).fillna(0)
        for metric, col in [("项目次数", "客次"), ("营业额", "支付金额")]
    }
    first_day = date(start.year, start.month, 1)
    days = [first_day + timedelta(days=n) for n in range((end - first_day).days + 1)]
    daily = [(d, compute(df[df["日期"] == pd.Timestamp(d)])) for d in days]
    return Report(start, end, compute(df), weeks, staff, categories, items, category_weeks, week_defs, daily,
                  context=ctx, data_span=(data_first, data_last))


def checks(df, df_all, months, month, sheet, rep):
    """数据检查提示（中葡双语）。以 ⚠ 开头的会标红。df = 报表月份的行，df_all = 整个文件。"""
    staff = rep.staff
    first, last = df_all["日期"].min(), df_all["日期"].max()
    out = [bi(f"读取 sheet「{sheet}」共 {len(df_all)} 行，账务日期 {first:%Y-%m-%d} ~ {last:%Y-%m-%d}。",
              f"Aba «{sheet}»: {len(df_all)} linhas, de {first:%d/%m/%Y} a {last:%d/%m/%Y}.")]
    if len(months) > 1:
        out.append(bi(f"报表月份 {month}：月度、员工月度、每日和目标只算 {month}；其他月份的天只用于跨月的周（周一~周日）。",
                      f"Mês do relatório: {month}. Mês, equipe, dias e metas contam só {month}; "
                      f"dias de outros meses entram apenas nas semanas que cruzam o mês."))
    outside = len(df_all) - len(df) - len(rep.context[rep.context["日期"].dt.to_period("M").astype(str) != month])
    if outside > 0:
        out.append(bi(f"⚠ {outside} 行既不在 {month}、也不在跨月的周里，已忽略。",
                      f"{outside} linhas fora de {month} e das semanas do relatório foram ignoradas."))
    (_, w1_start, w1_end), (_, wn_start, wn_end) = rep.week_defs[0], rep.week_defs[-1]
    data_first, data_last = rep.data_span
    if data_first > w1_start:
        out.append(bi(f"⚠ 第一周是 {dm(w1_start)}–{dm(w1_end)}，但文件从 {dm(data_first)} 才有数据，第一周只含 "
                      f"{dm(data_first)}–{dm(w1_end)}。要完整的第一周，导出时从 {w1_start:%Y-%m-%d} 开始导。",
                      f"A 1ª semana vai de {dm(w1_start)} a {dm(w1_end)}, mas o arquivo começa em {dm(data_first)}. "
                      f"Para a semana completa, exporte a partir de {w1_start:%d/%m/%Y}."))
    month_last = date(rep.start.year, rep.start.month, calendar.monthrange(rep.start.year, rep.start.month)[1])
    if rep.end == month_last and data_last < wn_end:
        out.append(bi(f"最后一周到 {dm(wn_end)}（下个月），文件只到 {dm(data_last)}；下个月的天导出后会补进这一周。",
                      f"A última semana vai até {dm(wn_end)} (mês seguinte); o arquivo vai só até {dm(data_last)}."))
    unknown = sorted(set(df["支付方式"].dropna()) - KNOWN_PAY_METHODS)
    if unknown:
        out.append(bi(f"⚠ 出现新的支付方式，已按「现金类」计入权益金/二次券：{'、'.join(unknown)}",
                      f"Nova forma de pagamento, contada como dinheiro/cartão: {', '.join(unknown)}"))
    products = df[df["类型"] == "产品"]
    if len(products):
        out.append(bi(f"产品销售 {products['数量'].sum():.0f} 件，合计 R$ {products['支付金额'].sum():,.2f}"
                      f"（{'、'.join(products['名称'].dropna().unique())}），不计入权益金和营业额。",
                      f"Venda de produtos: {products['数量'].sum():.0f} un., R$ {products['支付金额'].sum():,.2f} "
                      f"({', '.join(products['名称'].dropna().unique())}) — fora de vendas e receita."))
    splits = int(df["拆分行"].sum())
    if splits:
        out.append(bi(f"{splits} 行是组合支付拆出来的「数量 0」行：金额照算，不算客次，员工跟原行。",
                      f"{splits} linhas de pagamento dividido (qtd. 0): valor conta, não contam como atendimento."))
    types = sorted(set(df["类型"].dropna()) - {"项目", "权益", "产品"})
    if types:
        out.append(bi(f"⚠ 出现未识别的类型（不计入任何指标）：{'、'.join(types)}",
                      f"Tipo desconhecido (fora de todos os indicadores): {', '.join(types)}"))
    no_date = int(df["日期"].isna().sum())
    if no_date:
        out.append(bi(f"⚠ {no_date} 行没有账务日期，月度和员工数据已计入，周度数据未计入。",
                      f"{no_date} linhas sem data: entram no mês e na equipe, mas não nas semanas."))
    direct = df.loc[df["类型"] == "项目", ["体验额", "线上额"]].sum()
    if direct.sum():
        out.append(bi(f"权益金比收银系统多 R$ {direct.sum():,.2f}：做项目时直接付钱 R$ {direct['体验额']:,.2f}（算体验）"
                      f"＋ 做项目时直接用开业券 R$ {direct['线上额']:,.2f}（算线上卖券）；系统的权益金只算类型=权益。",
                      f"Vendas acima do sistema em R$ {direct.sum():,.2f}: serviços pagos na hora R$ {direct['体验额']:,.2f} "
                      f"+ serviços pagos com 开业券 R$ {direct['线上额']:,.2f} (o sistema só soma linhas 权益)."))
    for name, m, _ in staff:
        if name == UNASSIGNED:
            out.append(bi(f"⚠ 有记录没填员工：业绩 {m['业绩']:,.2f}、营业额 {m['营业额']:,.2f}、"
                          f"客次 {m['总客次']}，在员工表单列为「{UNASSIGNED}」。",
                          f"Registros sem colaborador(a): vendas R$ {m['业绩']:,.2f}, receita R$ {m['营业额']:,.2f}, "
                          f"{m['总客次']} atendimentos — aparecem como «{PT[UNASSIGNED]}»."))
    refunds = df[df["客次符号"] < 0]
    if len(refunds):
        docs = refunds["单据编号"].nunique() if "单据编号" in df.columns else len(refunds)
        total = refunds["支付金额"].sum()
        out.append(bi(f"退款单 {docs} 张（{len(refunds)} 行，支付金额合计 {total:,.2f}），已按负数抵扣。",
                      f"{docs} estorno(s) ({len(refunds)} linhas, R$ {total:,.2f}) descontados."))
    if "实业绩" in df.columns and pd.to_numeric(df["实业绩"], errors="coerce").fillna(0).eq(0).all():
        out.append(bi("明细里「实业绩」列全为 0，员工「业绩」按权益金口径计算。",
                      "A coluna 实业绩 está zerada; as vendas da equipe usam o critério «Vendas Pacotes/Vouchers»."))
    return out


# ---------------------------------------------------------------- 目标

@dataclass
class Targets:
    goals: dict      # 指标 -> 月目标
    weights: dict    # 日期 -> 当天占月目标的比例（合计 = 1），累加型指标按它拆到每天
    source: str
    month: str


def load_targets(path, start, store):
    """读目标 JSON（格式见 targets/Ibirapuera/2026-10.json）；没指定路径时找 targets/<门店>/<YYYY-MM>.json。
    找不到返回 None（不同门店目标不同，绝不拿别的门店的目标凑数）。"""
    if not path:
        if not store:
            return None
        path = TARGETS_DIR / store / f"{start:%Y-%m}.json"
    path = Path(path)
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
    source = f"{path.parent.name}/{path.name}"
    return Targets(goals, {d: v / total for d, v in daily.items()}, source, doc.get("月份", f"{start:%Y-%m}"))


def period_goal(tg, metric, first, last):
    """某段日期的目标：累加型 = 月目标 × 这段日期的权重之和；比率型 = 月目标。"""
    if metric in RATE_METRICS:
        return tg.goals[metric]
    return tg.goals[metric] * sum(w for d, w in tg.weights.items() if first <= d <= last)


def week_targets(rep, label, ws, we, m):
    """一周的 {指标: (目标, 实际)} 和提示。跨月的周用相邻月份的目标文件补上那几天；
    补不上时，累加型指标只比较有目标的那几天（实际也只算那几天），并返回一句提示。"""
    by_month = {f"{rep.start:%Y-%m}": rep.targets, **(rep.neighbor_targets or {})}
    days = [ws + timedelta(days=n) for n in range((we - ws).days + 1)]
    covered = [d for d in days if f"{d:%Y-%m}" in by_month]   # 跨月的周里缺的天只会在头或尾，所以是连续的
    missing = [d for d in days if d not in covered]
    sub = compute(in_days(rep.context, covered[0], covered[-1])) if missing and m is not None else None
    out = {}
    for k, goal in rep.targets.goals.items():
        if k in RATE_METRICS:
            out[k] = (goal, None if m is None else m[k])
        else:
            tg_of = lambda d: by_month[f"{d:%Y-%m}"]
            g = sum(tg_of(d).goals.get(k, 0) * tg_of(d).weights.get(d, 0) for d in covered)
            out[k] = (g, None if m is None else (sub if sub else m)[k])
    hint = None
    if missing:
        hint = bi(f"{label.split()[0]}：{dm(missing[0])}–{dm(missing[-1])} 没有每日目标（缺 {missing[0]:%Y-%m} 的目标文件），"
                  f"累加型指标的周目标和实际都只算 {dm(covered[0])}–{dm(covered[-1])}。",
                  f"{label.split()[0]}: sem metas diárias de {dm(missing[0])} a {dm(missing[-1])}; "
                  f"indicadores somáveis comparam só {dm(covered[0])}–{dm(covered[-1])}.")
    return out, hint


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
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
MUTED = "7F7F7F"

# 图表配色（dataviz 参考调色板）：同一个对象在所有图里用同一个颜色。
COLORS = {"新客项目次数": "2A78D6", "会员项目次数": "EB6834", "网红/邀约项目次数": "1BAF7A", "卖套餐": "6250D6", "二次券": "D69E00", "体验": "E05A8A", "线上卖券": "4BA3C3",
          "权益金": "6250D6", "业绩": "6250D6", "营业额": "008300",
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
    if metric in ("新客成交率", "会员复购率"):
        return f"{v:.0%}"
    if metric == "人效":
        return f"{v:.1f}"
    if metric in ("总客次", "会员客次", "新客客次", "网红/邀约", "成交人数", "新客成交人数", "续卡人数", "二次券成交数"):
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
        "叠加率": f"=IF({cols['总客次']}{row}=0,0,{hcols['项目次数']}{row}/{cols['总客次']}{row})",
        "新客成交率": rate_formula(cols["新客成交人数"], cols["新客客次"], row),
        "会员复购率": rate_formula(cols["续卡人数"], hcols["会员到店人数"], row),
        "人效": rate_formula(cols["总客次"], hcols["护理师人天"], row),
    }
    for k, letter in cols.items():
        cell = ws[f"{letter}{row}"]
        if m is not None:
            cell.value = formulas[k] if k in formulas else round(m[k], 2)
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
    labels = ["月目标", "月达成率", f"同期目标 / Meta até {dm(rep.end)}", "同期达成率"]
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
    """员工行。新客成交率写公式：新客成交人数 ÷ 右侧灰色「计算用:新客客次」。m=None 时写合计行（上方各行求和）。"""
    col = staff_columns(col0)
    for k, letter in col.items():
        c = ws[f"{letter}{row}"]
        helper = k in STAFF_HELPERS
        if k == "新客成交率":
            c.value = rate_formula(col["新客成交人数"], col["新客客次"], row)
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
    ws["A1"] = f"门店经营数据 Dados da Loja  {period(rep)}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, bi(f"数据来源：{source_name}　生成时间：{datetime.now():%Y-%m-%d %H:%M}　"
                      f"灰色「计算用」列为公式的底数；口径见「{SHEETS['指标说明']}」",
                      f"Fonte: {source_name} · gerado em {datetime.now():%d/%m/%Y %H:%M} · colunas cinza «Base» "
                      f"alimentam as fórmulas; definições na aba «{SHEETS['指标说明']}»"))

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

    # 权益金拆解：月度 + 各周，合计 = 权益金，最后一行是月度占比
    r = last_week_row + 3
    ws.cell(r, 1, "权益金拆解").font = font(bold=True)
    for j, h in enumerate(["期间", *BREAKDOWN, "权益金"], 2):
        style_header(ws.cell(r, j, h))
    rows = [("月度合计", rep.month), *rep.weeks]
    for i, (label, m) in enumerate(rows):
        row = r + 1 + i
        label_cell(ws, row, 2, label, bold=i == 0, fill=SUBTOTAL_FILL if i == 0 else None)
        for j, k in enumerate(BREAKDOWN, 3):
            c = ws.cell(row, j, None if m is None else round(m[k], 6))
            style_value(c, k)
        last_col = get_column_letter(2 + len(BREAKDOWN))
        total = ws.cell(row, 3 + len(BREAKDOWN), f"=SUM(C{row}:{last_col}{row})" if m is not None else None)
        style_value(total, "权益金")
        total.font = font(bold=True)
    share_row = r + 1 + len(rows)
    label_cell(ws, share_row, 2, "占比", bold=True)
    total_letter = get_column_letter(3 + len(BREAKDOWN))
    for j in range(3, 3 + len(BREAKDOWN)):
        letter = get_column_letter(j)
        c = ws.cell(share_row, j, ratio_formula(f"{letter}{r + 1}", f"{total_letter}{r + 1}"))
        style_value(c, "客次占比")
    note(ws, share_row + 1, 2, bi("权益金 = 卖套餐 + 二次券 + 体验 + 线上卖券。线上卖券 = 开业券支付；体验 = 做项目直接付钱 + 店里现买当天用掉的券；"
                                  "二次券 = 店里卖出、当天没用掉的券。员工业绩不含线上卖券。",
                                  "Vendas = pacotes + vouchers de retorno + experiência + vouchers online. Online = pago com 开业券; "
                                  "experiência = serviço pago na hora + voucher comprado e usado no mesmo dia; retorno = voucher "
                                  "comprado na loja e não usado no dia. Vendas da equipe não incluem online."))

    # 员工达成 · 月度数据（与手工模板一致）
    r = share_row + 4
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
    note(ws, total_row + 1, 3, bi("注：同一客户找过多位员工成交时会各算一次，故人数合计（及合计行的新客成交率）可能略高于月度总数据。",
                                  "Obs.: cliente que fechou com mais de uma colaboradora conta para cada uma, então o total "
                                  "de clientes (e a conversão da linha Total) pode ficar um pouco acima do resumo do mês."))

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
        ws.column_dimensions[get_column_letter(i)].width = 20 if i >= helper_start else 15
    for letter, w in {"A": 30, "B": 32, "C": 28}.items():
        ws.column_dimensions[letter].width = w
    ws.freeze_panes = "C1"
    landscape(ws)


def period(rep):
    """标题里的「门店 · 日期区间」。"""
    store = f"{rep.store} · " if rep.store else ""
    return f"{store}{rep.start:%Y-%m-%d} ~ {rep.end:%Y-%m-%d}"


def landscape(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------------------------------------------------------- 项目分类 sheet

def write_mix_block(ws, r, title, table, key_headers, with_share):
    """写一张 分类(或 分类+项目) 汇总表，返回下一个空行。占比、平均单价用公式。"""
    ws.cell(r, 1, title).font = font(bold=True, size=11)
    r += 1
    metrics = ["项目次数", *MIX_TYPES, "客户人数", "营业额"]
    headers = [*key_headers, "项目次数", *(["项目次数占比"] if with_share else []), *MIX_TYPES, "客户人数",
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
            for share, base in [("项目次数占比", "项目次数"), ("营业额占比", "营业额")]:
                ws[f"{col[share]}{row}"] = f"=IF({col[base]}${total}=0,0,{col[base]}{row}/{col[base]}${total})"
                style_value(ws[f"{col[share]}{row}"], share)
        ws[f"{col['平均单价']}{row}"] = f"=IF({col['项目次数']}{row}=0,0,{col['营业额']}{row}/{col['项目次数']}{row})"
        style_value(ws[f"{col['平均单价']}{row}"], "平均单价")
    if with_share:
        label_cell(ws, total, 1, "合计", bold=True)
        for h in headers[1:]:
            c = ws[f"{col[h]}{total}"]
            if h in ("项目次数占比", "营业额占比"):
                c.value = f"=SUM({col[h]}{first}:{col[h]}{last})"
            elif h == "平均单价":
                c.value = f"=IF({col['项目次数']}{total}=0,0,{col['营业额']}{total}/{col['项目次数']}{total})"
            elif h == "客户人数":
                c.value = None  # 客户可能做过多个分类，人数不能直接相加
            else:
                c.value = f"=SUM({col[h]}{first}:{col[h]}{last})"
            style_value(c, h)
            c.font = font(bold=True)
    for k in ["项目次数", "营业额"]:
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
    """「W3 15/out-21/out（截至/até 18/out）」→「W3（截至/até 18/out）」，用于表头和图表坐标轴。"""
    return label.split()[0] + (label[label.index("（"):] if "（" in label else "")


def write_efficiency_sheet(wb, rep):
    """人效：月度 + 各周（总客次、护理师人天、营业天数、日均上班护理师、人效），再列每日明细和每日人效图。"""
    ws = wb.create_sheet(SHEETS["人效"])
    ws["A1"] = f"人效 Produtividade  {period(rep)}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, bi("人效 = 总客次 ÷ 护理师人天。护理师人天 = 每天有做项目的员工人数相加（来自售出明细，不是排班表）。",
                      "Produtividade = atendimentos ÷ terapeuta-dias (colaboradoras que atenderam em cada dia, pelas vendas, "
                      "não pela escala)."))
    r = 4
    ws.cell(r, 1, "月度 / 周度人效").font = font(bold=True, size=11)
    heads = ["期间", "总客次", "护理师人天", "营业天数", "日均上班护理师", "人效"]
    for j, h in enumerate(heads, 1):
        style_header(ws.cell(r + 1, j, h))
    for i, (label, m) in enumerate([("月度合计", rep.month), *rep.weeks]):
        row = r + 2 + i
        label_cell(ws, row, 1, label, bold=i == 0, fill=SUBTOTAL_FILL if i == 0 else None)
        for j, k in enumerate(["总客次", "护理师人天", "营业天数"], 2):
            style_value(ws.cell(row, j, None if m is None else m[k]), k)
        style_value(ws.cell(row, 5, ratio_formula(f"C{row}", f"D{row}") if m else None), "日均上班护理师")
        style_value(ws.cell(row, 6, ratio_formula(f"B{row}", f"C{row}") if m else None), "人效")
    r = r + 2 + 1 + len(rep.weeks) + 2

    ws.cell(r, 1, "每日人效").font = font(bold=True, size=11)
    heads = ["日期", "星期", "总客次", "上班护理师", "人效", "当天做项目的护理师"]
    for j, h in enumerate(heads, 1):
        style_header(ws.cell(r + 1, j, h))
    proj = rep.context[(rep.context["类型"] == "项目") & rep.context["员工"].notna()]
    per_day = proj.groupby(["日期", "员工"])["客次"].sum()
    first = r + 2
    for i, (d, m) in enumerate(rep.daily):
        row = first + i
        names = sorted(n for (day, n), v in per_day.items() if day.date() == d and v > 0)
        values = [dm(d), f"周{WEEKDAYS[d.weekday()]} {PT_WEEKDAYS[d.weekday()]}", m["总客次"], m["护理师人天"],
                  ratio_formula(f"C{row}", f"D{row}"), "、".join(names)]
        for j, v in enumerate(values, 1):
            c = ws.cell(row, j, v)
            c.font, c.border = font(), BORDER
            if not m["总客次"]:
                c.fill = HELPER_FILL
        ws.cell(row, 5).number_format = METRIC_FORMATS["人效"]
    last = first + len(rep.daily) - 1
    data_bar(ws, f"E{first}:E{last}")
    for letter, w in zip("ABCDEF", [30, 14, 12, 16, 12, 80]):
        ws.column_dimensions[letter].width = w
    chart = make_chart(ws, r + 1, last, [5], "每日人效（客次/人）\nAtendimentos por terapeuta por dia",
                       width=22, height=8, num_fmt="0.0")
    ws.add_chart(chart, "H4")
    landscape(ws)


def write_category_sheet(wb, rep):
    ws = wb.create_sheet(SHEETS["项目分类"])
    ws["A1"] = f"项目分类统计 Categorias de Serviço  {period(rep)}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, bi("只统计类型=项目（客人实际做的项目），按项目次数算（一次到店做两个项目算 2 次）；"
                      "新客/会员/网红 按那次到店的类型；平均单价=营业额÷项目次数。"
                      f"各项目明细见「{SHEETS['项目明细']}」。",
                      "Só serviços realizados, contados por serviço; novos/pacote/influencers pelo tipo da visita; "
                      f"preço médio = receita ÷ serviços. Detalhe por serviço na aba «{SHEETS['项目明细']}»."))
    r = write_mix_block(ws, 4, "按分类（月度）", rep.categories, ["分类"], with_share=True)
    r = write_week_matrix(ws, r, "按分类 · 周度项目次数", rep.category_weeks["项目次数"], "项目次数")
    write_week_matrix(ws, r, "按分类 · 周度营业额", rep.category_weeks["营业额"], "营业额")
    ws.column_dimensions["A"].width = 24
    for j in range(2, 12):
        ws.column_dimensions[get_column_letter(j)].width = 15
    landscape(ws)

    ws = wb.create_sheet(SHEETS["项目明细"])
    ws["A1"] = f"项目明细 Serviços  {period(rep)}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, bi("每个项目做了多少次、谁在做（新客/会员/网红）、带来多少营业额；按分类分组，组内按项目次数排序。",
                      "Quantas vezes cada serviço foi feito, para quem (novos/pacote/influencers) e quanto rendeu; "
                      "agrupado por categoria, ordenado por atendimentos."))
    write_mix_block(ws, 4, "按项目（月度）", rep.items, ["分类", "项目名称"], with_share=False)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 46
    for j in range(3, 10):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = "C6"
    landscape(ws)


# ---------------------------------------------------------------- 目标达成 sheet

def write_target_sheet(wb, rep):
    ws = wb.create_sheet(SHEETS["目标达成"])
    tg = rep.targets
    progress, remaining, rows = target_summary(rep)
    ws["A1"] = f"目标达成 Metas  {period(rep)}"
    ws["A1"].font = font(size=14, bold=True)
    note(ws, 2, 1, bi(f"目标来源：{tg.source}。✓ = 达到同期目标（按每日目标累计到 {dm(rep.end)}），✗ = 落后。"
                      f"口径见「{SHEETS['指标说明']}」。",
                      f"Metas: {tg.source}. ✓ = atingiu a meta até {dm(rep.end)} (soma das metas diárias), "
                      f"✗ = abaixo. Definições na aba «{SHEETS['指标说明']}»."))
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
    note(ws, r, 3, bi("进行中的周用整周目标比较，达成率会偏低；还没到的周只列目标。",
                      "A semana em andamento é comparada com a meta da semana inteira; semanas futuras só mostram a meta."))
    goals = list(tg.goals)
    style_header(ws.cell(r + 1, 1, "周"))
    for g, k in enumerate(goals):
        for j, h in enumerate([f"{k}目标", f"{k}实际", "达成率"]):
            style_header(ws.cell(r + 1, 2 + g * 3 + j, h))
    hints = []
    for i, ((label, ws_, we), (_, m)) in enumerate(zip(rep.week_defs, rep.weeks)):
        row = r + 2 + i
        label_cell(ws, row, 1, label)
        values, hint = week_targets(rep, label, ws_, we, m)
        if hint:
            hints.append(hint)
        for g, k in enumerate(goals):
            tcol, acol, rcol = (get_column_letter(2 + g * 3 + j) for j in range(3))
            goal, actual = values[k]
            ws[f"{tcol}{row}"] = round(goal, 6)
            ws[f"{acol}{row}"] = None if actual is None else round(actual, 6)
            ws[f"{rcol}{row}"] = ratio_formula(f"{acol}{row}", f"{tcol}{row}") if m is not None else None
            style_value(ws[f"{tcol}{row}"], k)
            style_value(ws[f"{acol}{row}"], k)
            style_value(ws[f"{rcol}{row}"], "月达成率")
            ws[f"{rcol}{row}"].number_format = STATUS_FORMAT
    last = r + 1 + len(rep.week_defs)
    for g in range(len(goals)):
        letter = get_column_letter(4 + g * 3)
        status_format(ws, f"{letter}{r + 2}:{letter}{last}")
    for hint in hints:
        last += 1
        note(ws, last, 1, "⚠ " + hint)
        ws.cell(last, 1).font = font(color=BAD, size=9)

    # 每日（只列累加型目标）
    r = last + 3
    additive = [k for k in goals if k not in RATE_METRICS]
    ws.cell(r, 1, "每日目标 vs 实际").font = font(bold=True, size=11)
    note(ws, r, 3, bi("累计达成率 = 月初到当天的累计实际 ÷ 累计目标。",
                      "% acumulado = realizado acumulado desde o dia 1 ÷ meta acumulada."))
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
        c.number_format, c.font, c.border = "dd/mm", font(), BORDER
        closed = "（休 Fechado）" if tg.weights[d] == 0 else ""
        label_cell(ws, row, 2, f"周{WEEKDAYS[d.weekday()]} {PT_WEEKDAYS[d.weekday()]}{closed}")
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

    for j in range(2, 3 + 5 * max(len(additive), 1) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 22
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
    ds = wb.create_sheet(SHEETS["图表数据"])
    note(ds, 1, 1, bi("看板图表的数据来源，由脚本自动生成。", "Dados dos gráficos do painel, gerados automaticamente."))
    pos = {}

    weeks = [(label, m) for label, m in rep.weeks if m is not None]
    headers = ["周", "权益金", "营业额", *CUSTOMER_TYPES, *BREAKDOWN]
    pos["weeks"] = write_block(ds, 3, headers, [[short_week(label), *(round(m[k], 2) for k in headers[1:])]
                                                for label, m in weeks], "#,##0")
    pos["breakdown_cols"] = [headers.index(k) + 1 for k in BREAKDOWN]

    # 横向条形图从下往上画，所以按升序写，最大的排在最上面
    staff = sorted(rep.staff, key=lambda s: (s[1]["业绩"], s[1]["营业额"]))
    pos["staff"] = write_block(ds, pos["weeks"][1] + 3, ["员工", "业绩", "营业额"],
                               [[name, round(m["业绩"], 2), round(m["营业额"], 2)] for name, m, _ in staff], "#,##0")

    headers = ["分类", *MIX_TYPES, "营业额"]
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
            row = [dm(d)]
            for k in money:
                cum[k][0] += tg.goals[k] * tg.weights[d]
                if d <= rep.end:
                    cum[k][1] += actual_by_day[d][k] if d in actual_by_day else 0
                row += [round(cum[k][0], 2), round(cum[k][1], 2) if d <= rep.end else None]
            cum_rows.append(row)
        pos["cumulative"] = write_block(ds, pos["achieve"][1] + 3,
                                        ["日期", *(h for k in money for h in (f"{k}累计目标", f"{k}累计"))],
                                        cum_rows, "#,##0")
        week_rows = []
        for (label, ws_, we), (_, m) in zip(rep.week_defs, rep.weeks):
            values, _ = week_targets(rep, label, ws_, we, m)
            week_rows.append([short_week(label), *(v for k in money for v in (
                round(values[k][0], 2), None if values[k][1] is None else round(values[k][1], 2)))])
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
    rate = m["新客成交率"]
    tiles = [
        ("权益金", '"R$ "#,##0', m["权益金"],
         f"含线上卖券 Online R$ {m['线上卖券']:,.0f}（{m['线上卖券'] / m['权益金']:.0%}）" if m["权益金"] else ""),
        ("营业额", '"R$ "#,##0', m["营业额"], "项目核销价值 · Valor dos serviços realizados"),
        ("总客次", "#,##0", m["总客次"],
         f"人效 {m['人效']:.1f} 客次/人/天 · por terapeuta/dia" if m["人效"] else "做项目的人次 · Serviços realizados"),
        ("新客客次", "#,##0", m["新客客次"],
         f"占总客次 / do total {m['新客客次'] / m['总客次']:.0%}" if m["总客次"] else ""),
        ("会员客次", "#,##0", m["会员客次"], "套餐核销 · Pagos com pacote"),
        ("网红/邀约", "#,##0", m["网红/邀约"], "免单类 · Cortesias"),
        ("成交人数", "#,##0", m["成交人数"],
         f"新客 {m['新客成交人数']} · 续卡 {m['续卡人数']}（复购率 {m['会员复购率']:.0%}）" if m["会员复购率"] is not None
         else f"新客 {m['新客成交人数']} · 续卡 {m['续卡人数']}"),
        ("新客成交率", "0.00%", rate if rate is not None else "-",
         f"{m['新客成交人数']} ÷ {m['新客客次']}（新客成交 ÷ 新客客次 · fechamentos ÷ visitas novas）"),
        ("成交客单", '"R$ "#,##0', m["成交客单"], "套餐金额 ÷ 成交人数 · Pacotes ÷ fechamentos"),
        ("二次券成交数", "#,##0", m["二次券成交数"], f"{m['二次券成交人数']} 位客户 · clientes"),
        ("体验客单价", '"R$ "#,##0.00', m["体验客单价"], "营业额 ÷ 总客次 · Receita ÷ atendimentos"),
        ("叠加率", "0.00", m["叠加率"], f"{m['项目次数']} 个项目 ÷ {m['总客次']} 客次 · serviços ÷ visitas"),
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
        title = ws.cell(row, col, f"{label}\n{PT[label]}")  # 卡片标题两行，不走 translate_workbook
        title.font = font(size=9, color=MUTED)
        title.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        ws.row_dimensions[row].height = 26
        v = ws.cell(row + 1, col, round(value, 6) if isinstance(value, (int, float)) else value)
        v.font, v.number_format = font(size=18, bold=True), fmt_
        ws.row_dimensions[row + 1].height = 28
        t = goals.get(label)
        if t:
            ok = t["同期达成率"] >= 1
            sub = f"目标 Meta {goal_text(label, t['月目标'])} · {'✓' if ok else '✗'} {t['同期达成率']:.0%}"
            ws.cell(row + 2, col, sub).font = font(size=8, bold=True, color=GOOD if ok else BAD)
            bar = ws.cell(row + 3, col, round(t["月达成率"], 4))
            bar.number_format, bar.font = '"月达成 Mês "0%', font(size=8, color=MUTED)
            progress_bar(ws, bar.coordinate)
        else:
            ws.cell(row + 2, col, sub).font = font(size=8, color=MUTED)
    return 4 + -(-len(tiles) // 4) * (height + 1) + 1


def write_dashboard(wb, rep):
    ws = wb.create_sheet(SHEETS["看板"], 0)
    ws.sheet_view.showGridLines = False
    for j in range(1, 18):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions["A"].width = 2
    ws["B1"] = f"门店经营看板 Painel da Loja  {period(rep)}"
    ws["B1"].font = font(size=16, bold=True)
    hint_zh = f"数值明细见「{SHEETS['经营报表']}」「{SHEETS['项目分类']}」；口径见「{SHEETS['指标说明']}」。"
    hint_pt = f"Detalhes nas abas «{SHEETS['经营报表']}» e «{SHEETS['项目分类']}»; definições em «{SHEETS['指标说明']}»."
    if rep.targets:
        progress = target_summary(rep)[0]
        hint_zh = (f"目标进度 {progress:.1%}（按每日目标累计到 {dm(rep.end)}）；✓/✗ = 是否达到同期目标，"
                   f"进度条 = 月目标完成比例，详见「{SHEETS['目标达成']}」。") + hint_zh
        hint_pt = (f"Progresso esperado {progress:.1%} (metas diárias até {dm(rep.end)}); ✓/✗ = meta até hoje "
                   f"atingida ou não; barra = % da meta do mês (aba «{SHEETS['目标达成']}»). ") + hint_pt
    note(ws, 2, 2, hint_zh)
    note(ws, 3, 2, hint_pt)
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
            charts.append(make_line_chart(ds, d0, d1, [2 + i * 2, 3 + i * 2],
                                          f"{k} 累计 vs 目标（R$）\n{PT[k]}: acumulado vs meta"))
        charts.append(make_chart(ds, a0, a1, [2, 3], "目标达成率 vs 应达进度\n% atingido vs progresso esperado",
                                 horizontal=True, labels=True,
                                 height=max(7.5, 1.1 * (a1 - a0) + 2.5), num_fmt="0%"))
        for i, k in enumerate(pos["money"]):
            charts.append(make_chart(ds, t0, t1, [2 + i * 2, 3 + i * 2],
                                     f"周度 {k} 实际 vs 目标（R$）\n{PT[k]}: semanal vs meta"))
    else:
        charts.append(make_chart(ds, w0, w1, [2, 3], "周度 权益金 vs 营业额（R$）\nVendas vs receita por semana",
                                 labels=True))
    charts += [
        make_chart(ds, w0, w1, pos["breakdown_cols"], "周度权益金构成（R$）\nComposição das vendas por semana",
                   stacked=True),
        make_chart(ds, w0, w1, [4, 5, 6], "周度客次构成（新客/会员/网红）\nAtendimentos por semana e tipo de cliente",
                   stacked=True),
        make_chart(ds, s0, s1, [2, 3], "员工 业绩 vs 营业额（R$）\nEquipe: vendas vs receita", horizontal=True,
                   height=staff_height),
        make_chart(ds, c0, c1, [2, 3, 4], "项目分类项目次数（按客户类型）\nServiços por categoria e tipo de cliente",
                   horizontal=True, stacked=True,
                   height=cat_height),
        make_chart(ds, c0, c1, [5], "项目分类营业额（R$）\nReceita por categoria", horizontal=True, labels=True,
                   height=cat_height),
    ]
    row = top
    for i in range(0, len(charts), 2):  # 两列排布，每行高度取两张图中较高的
        pair = charts[i:i + 2]
        for ch, col in zip(pair, "BJ"):
            ws.add_chart(ch, f"{col}{row}")
        row += int(max(ch.height for ch in pair) / 0.53) + 2
    landscape(ws)


# ---------------------------------------------------------------- 其他 sheet 与入口

def wrapped_lines(text, width):
    """估算一段文字在给定列宽里自动换行后的行数（中文按 2 个字符宽，按词折行）。"""
    cap = max(int(width) - 2, 4)
    lines, used = 1, 0
    for word in text.split(" "):
        w = sum(2 if ord(ch) > 0x2E80 else 1 for ch in word)
        if used and used + 1 + w > cap:
            lines, used = lines + 1, 0
        used += w + (1 if used else 0)
        while used > cap:
            lines, used = lines + 1, used - cap
    return lines


def translate_workbook(wb):
    """保存前把表头和固定标签换成中葡双语：表头（蓝/灰底）上下两行，其他标签「中文 / Português」。
    图表数据 sheet 一律单行（它的表头就是图例和坐标轴文字）。"""
    header_fills = {HEAD_FILL.fgColor.rgb[-6:], HELPER_FILL.fgColor.rgb[-6:]}
    for ws in wb.worksheets:
        inline_only = ws.title == SHEETS["图表数据"]
        heights = {}
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value, str) or c.value.startswith("="):
                    continue
                p = pt_of(c.value)
                if p is None:
                    continue
                if inline_only or (c.fill.fgColor.rgb or "")[-6:] not in header_fills:
                    c.value = f"{c.value} / {p}"
                    continue
                width = ws.column_dimensions[c.column_letter].width or 8.43
                lines = wrapped_lines(c.value, width) + wrapped_lines(p, width)
                c.value = f"{c.value}\n{p}"
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                heights[c.row] = max(heights.get(c.row, 15), 13 * lines + 4)
        for r, h in heights.items():
            ws.row_dimensions[r].height = h


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
        ws.cell(i, 1).number_format = "dd/mm/yyyy"
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
    wb.active.title = SHEETS["经营报表"]
    if rep.targets:
        write_target_sheet(wb, rep)
    write_efficiency_sheet(wb, rep)
    write_category_sheet(wb, rep)

    ds = wb.create_sheet(SHEETS["指标说明"])
    for j, h in enumerate(["指标", "计算口径", "Definição (português)", "本期数值"], 1):
        style_header(ds.cell(1, j, h))
    for i, (k, text) in enumerate(DEFINITIONS, 2):
        ds.cell(i, 1, k).font = font(bold=True)
        ds.cell(i, 2, text).font = font()
        ds.cell(i, 3, DEFINITIONS_PT.get(k)).font = font()
        for j in (1, 2, 3):
            ds.cell(i, j).alignment = Alignment(wrap_text=True, vertical="top")
        if rep.month.get(k) is not None:
            style_value(ds.cell(i, 4, round(rep.month[k], 6)), k)
        for j in (1, 2, 3, 4):
            ds.cell(i, j).border = BORDER
    natural = week_mode == "natural"
    note(ds, len(DEFINITIONS) + 2, 1, bi(f"周度划分方式：{'周一~周日（跨月的周算满 7 天）' if natural else '每月按 7 天一段'}",
                                         f"Semanas: {'segunda a domingo (semanas que cruzam o mês com 7 dias)' if natural else 'blocos de 7 dias a partir do dia 1'}"))
    ds.column_dimensions["A"].width = 24
    ds.column_dimensions["B"].width = 60
    ds.column_dimensions["C"].width = 60
    ds.column_dimensions["D"].width = 14
    landscape(ds)

    # 核对明细（相当于原来的透视表）
    right = df[df["类型"] == "权益"]
    package = right[(right["分类"] == "套餐") & (right["现金类支付"] != 0)]
    retorno = df[(df["二次券数"] != 0) | (df["二次券额"] != 0)].assign(数量=lambda x: x["二次券数"].round(2))
    write_detail(wb, SHEETS["套餐成交明细"], package, "现金类支付")
    write_detail(wb, SHEETS["二次券成交明细"], retorno, "二次券额")

    cs = wb.create_sheet(SHEETS["数据检查"])
    style_header(cs.cell(1, 1, "数据检查 / 提示"))
    for i, line in enumerate(notes, 2):
        cs.cell(i, 1, line).font = font(color="C00000" if line.startswith("⚠") else "000000")
    cs.column_dimensions["A"].width = 120

    write_dashboard(wb, rep)  # 放在第一个 sheet；图表数据 sheet 排在最后
    translate_workbook(wb)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def fmt(k, v):
    if v is None:
        return "-"
    if k in ("新客成交率", "会员复购率"):
        return f"{v:.2%}"
    return f"{v:,.2f}" if isinstance(v, float) else str(v)


def main(argv=None):
    ap = argparse.ArgumentParser(description="售出明细表 → 门店经营报表")
    ap.add_argument("input", help="收银系统导出的 Excel（含「售出明细表」）")
    ap.add_argument("-o", "--output", help="输出路径，默认 ./门店经营报表_<开始>_<结束>.xlsx")
    ap.add_argument("--month", help="报表月份，例如 2026-10（默认取文件里数据最多的月份）")
    ap.add_argument("--weeks", choices=["natural", "fixed"], default="natural",
                    help="natural: 周一~周日，跨月的周算满 7 天（默认）；fixed: 1-7日=W1…")
    ap.add_argument("--targets", help="目标 JSON，默认 targets/<门店>/<YYYY-MM>.json")
    ap.add_argument("--store", help="门店名（例如 Ibirapuera、Morumbi），默认按「权益归属门店」识别")
    args = ap.parse_args(argv)

    df_all, sheet, months = load(args.input)
    period_of = df_all["日期"].dt.to_period("M").astype(str)
    month = args.month or period_of[df_all["日期"].notna()].value_counts().idxmax()
    df = df_all[(period_of == month) | df_all["日期"].isna()].copy()
    if df["日期"].isna().all():
        raise SystemExit(f"文件里没有 {month} 的数据（有：{', '.join(months)}）")
    df.attrs["store"] = df_all.attrs.get("store")
    rep = build(df, args.weeks, df_all)
    rep.store = args.store or df.attrs.get("store")
    rep.targets = load_targets(args.targets, rep.start, rep.store)
    if rep.targets:  # 跨月的周：相邻月份有目标文件就用它补上那几天
        rep.neighbor_targets = {}
        for d in (rep.week_defs[0][1], rep.week_defs[-1][2]):
            key = f"{d:%Y-%m}"
            if key != month and key not in rep.neighbor_targets:
                tg = load_targets(None, date(d.year, d.month, 1), rep.store)
                if tg:
                    rep.neighbor_targets[key] = tg
    notes = checks(df, df_all, months, month, sheet, rep)
    notes.insert(0, bi(f"门店：{rep.store or '未识别'}" + ("（--store 指定）" if args.store else "（按权益归属门店识别）"),
                       f"Loja: {rep.store or 'não identificada'}"))
    if rep.targets:
        notes.append(bi(f"目标文件：{rep.targets.source}（{'、'.join(rep.targets.goals)}）。",
                        f"Metas: {rep.targets.source} ({', '.join(pt_of(k) or k for k in rep.targets.goals)})."))
        if rep.targets.month != f"{rep.start:%Y-%m}":
            notes.append(bi(f"⚠ 目标文件写的是 {rep.targets.month}，数据是 {rep.start:%Y-%m}，请确认目标是否用对。",
                            f"O arquivo de metas é de {rep.targets.month} e os dados de {rep.start:%Y-%m}."))
    else:
        notes.append(bi(f"没有 {rep.store} {rep.start:%Y-%m} 的目标文件（targets/{rep.store}/{rep.start:%Y-%m}.json），"
                        f"本次报表不含目标达成。",
                        f"Sem metas para {rep.store} {rep.start:%Y-%m}; relatório sem acompanhamento de metas."))
    store_tag = f"{rep.store}_" if rep.store else ""
    out = Path(args.output or f"门店经营报表_{store_tag}{rep.start:%Y%m%d}_{rep.end:%Y%m%d}.xlsx")
    write_report(out, df, rep, notes, Path(args.input).name, args.weeks)

    print(f"已生成：{out}（门店 {rep.store}）")
    print(f"区间：{rep.start} ~ {rep.end}")
    print("月度总数据：" + "，".join(f"{k} {fmt(k, rep.month[k])}" for k in MONTH_METRICS))
    print("权益金拆解：" + "，".join(f"{k} {rep.month[k]:,.2f}" for k in BREAKDOWN))
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
    print("项目分类（项目次数｜新客/会员/网红｜营业额）：")
    total = rep.categories["项目次数"].sum()
    for cat, x in rep.categories.iterrows():
        print(f"  {cat}：{x['项目次数']:.0f}（{x['项目次数'] / total:.0%}）｜{x['新客项目次数']:.0f}/{x['会员项目次数']:.0f}/"
              f"{x['网红/邀约项目次数']:.0f}｜{x['营业额']:,.2f}")
    for line in notes:
        print(line)


if __name__ == "__main__":
    sys.exit(main())

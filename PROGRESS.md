# 1999 Diamonds — Project Progress

Running log for the 1999 lab-grown diamonds business. Maintained by the
**Diamond Studio** agent (`.claude/agents/diamond-studio.md`). Newest entries
go at the top of each section.

## Overview

- **Business:** 1999 — lab-grown diamonds.
- **Riverflow team:** "1999 diamonds".
- **Assistant:** Diamond Studio agent handles image creation (Riverflow) and
  keeps this log up to date.

## Milestones

- **2026-08-09** — 新增样品文档:`docs/样品清单.md`(约16件精简清单)+ `docs/样品拿货思路.md`(拿货原则/取舍/实际进货记录待填 —— 等创始人发戒指照片登记)。
- **2026-08-09** — 说明书 v1.2:补齐三块调研短板 —— 巴西钻戒需求数据(结婚95万对/年、钻戒均价R$4,900、节点体量)、同行价格全面对比(培育 vs 天然 vs 莫桑,三方同规格表)、本地培育钻品牌全景(GAEM/Naïve/Panna/Vivara/Pandora 等 + 修正 Only Diamonds/Ayaris 引用)。关键结论:价格是入场券非护城河,差异化靠 IP+情感+证书透明。
- **2026-08-09** — 说明书 v1.1:敲定价格带(求婚 R$2,800–20,000/1ct 18K≈R$8,500;日常 R$800起)、英雄单品(单钻耳钉 ponto de luz)、主理人本人出镜、出图主力用 Riverflow(Lovart 待网络放行)。
- **2026-08-09** — 接入 Lovart AI skill(`.claude/skills/lovart-api/`),密钥走 `.env`(gitignore),配置见 `docs/lovart-setup.md`;端到端测试待重开 session + 填新密钥后进行。
- **2026-08-09** — 完成《项目说明书 v1》(`docs/项目说明书.md`),含巴西/美国市场调研 + 竞品对标 + 执行步骤 + 时间线建议。
- **2026-08-09** — Set up the Diamond Studio agent and this progress log.

## 业务现状快照(2026-08-09)

- ✅ 已完成:市场初步调研、公司注册+开户、LOVART 初步品牌设计、供应商确认
- 🔄 进行中:国内样品采购
- 🔜 待办:包装盒供应商 → 物料设计 → 进出口&资金流转 → 素材 → 网站 → 社媒 → 内容 → Campaign → 开售
- 关键定位:巴西培育钻,"真钻但够得着";求婚系列(核心)+ 个人佩戴系列;TikTok引流+IG品牌+葡语独立站
- 建议时间线:圣诞季(2026.11-12)软启动首销 → 2027 母亲节+情人节大 Campaign

## Image assets log

| Date | Purpose | Type | Riverflow generation ID | Link |
|------|---------|------|-------------------------|------|
| _(none yet)_ | | | | |

## Decisions

- **2026-08-09** — Track all project progress in `PROGRESS.md` at the repo root.

## Open tasks

- [ ] Add products / product images to Riverflow to enable product photoshoots.
- [ ] Decide on brand style (logo, fonts, colours) and set up brand style rules.
- [ ] Generate a first set of hero product images.

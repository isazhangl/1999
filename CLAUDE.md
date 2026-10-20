# 用户与工作背景（给以后的会话看）

- 用户 Isa：**KOMO 的 CEO**（2026-10-20 升任，之前负责营运）；同时是 **1999**（培育钻首饰品牌，本仓库最早的项目）的创始人，
  另外协助「海外仓」「电商」两个项目。中文沟通；有疑问时**用弹窗（AskUserQuestion）问**，用户明确要求过。

## 工作管理面板「项目总控台」

- 链接：https://claude.ai/artifact/Veaga83VgWD2yyhtiNt1Km （用户自己的 artifact；能力：db、user、mcp Windsor.ai get_data 拉 Meta 投放数据）
- 四个视图：今天（所有公司当天待办）/ 规划（每家公司：目标 → KR → 小项目 → 待办）/ 投放（Meta 广告）/ 全部公司。
- 数据在 `projects` 集合，一家公司一个文档：
  | 公司 | doc_id | 我的身份 |
  |---|---|---|
  | KOMO | `kf38vrinms2eljqqupjz` | 任职（exec），现为 CEO |
  | 1999 | `7amuvgf3ei7p22m5tb9o` | 创始人（lead） |
  | 电商 | `9b8aexb0cikmqfna68x9` | 协助 |
  | 海外仓 | `3axnuicqlbqf9e1lauif` | 协助 |
- 文档字段：`objectives[{id,title,period,status,krs[{id,title,target,unit,current}]}]`、`initiatives[{id,title,due,objectiveId}]`、
  `tasks[{id,text,done,due(YYYY-MM-DD),initiativeId,createdAt,doneAt,today}]`，以及 name/stage/roleType/focus/health/milestone/notes 等基本信息。
  有 due 的待办到那天会自动出现在「今天」。
- 往面板里加/改事项：用 ArtifactData 先 `get` 该公司文档拿 version，在原数组上追加（id 用 8 位随机串），
  `update` 时带 `if_version`；只改 objectives/initiatives/tasks，不要覆盖别的字段。改之前先跟用户确认。

## KOMO（巴西圣保罗头皮护理门店）

- **1店 = Morumbi（街边店，先开，11月满一周年）**，**2店 = Ibirapuera（Shopping Ibirapuera 商场店，后开）**。
- 周报：用户说「做周报」→ 按 `.claude/skills/store-report/SKILL.md` 做（脚本 `analytics/store_report/store_report.py`，
  目标在 `analytics/store_report/targets/<门店>/<YYYY-MM>.json`）。
- 新客增长计划：`analytics/plans/make_growth_plan.py` 生成 Excel（当前草稿 v7，含 10/29 Morumbi VIP 沙龙和 11 月活动决定）。
- 收银明细含客户姓名/电话：输入和输出 Excel 都不入库（`reports/` 已 gitignore）。

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

## KOMO Q4 OKR 面板（单独一个，和总控台分开）

- 链接：https://claude.ai/artifact/3gjAZkHcw4bZ3wC75EETsW （能力：db、user；中/PT 切换；给店长、市场部一起更新）
- 2026 Q4 = 11–12 月，KR 按月、按店（`m` = 1店 Morumbi，`i` = 2店 Ibirapuera）看。O1 业绩达标（O 本身就是权益金+营业额，
  KR：新客数 / 新客成交率 / 老客续卡（先金额后人数）/ 客次，**具体数字由 Isa 拆**）；O2 搭建加盟体系（打好基础，含 COF）；
  O3 组织架构和流程优化（11 月底前；KR 写法还是草稿）。总控台里的 KOMO 待办保持不动，人事类待办（升职、开除）不放进这个共享面板。
- 集合（一项一个文档，文字字段 `zh`/`pt`）：`objectives{kind:okr|ops,noteZh,notePt}`、
  `krs{objectiveId,type:metric|milestone,kind?:result,lines[{key,zh,pt,unit}],values{<line>:{<YYYY-MM>:{m|i:{t,a}}}},status,due,ownerIds}`、
  `projects{objectiveId,krId,ownerIds,due}`、`tasks{projectId,krId,objectiveId,done,due,ownerIds,doneAt,doneBy}`、`members{name,pt,role,rolePt}`。
  负责人存 `members` 的 id（组织架构图来了再把「1店店长/2店店长」改成名字）。用 ArtifactData 改之前先跟用户确认，写已有文档要带 if_version。

## KOMO（巴西圣保罗头皮护理门店）

- **1店 = Morumbi（街边店，先开，11月满一周年）**，**2店 = Ibirapuera（Shopping Ibirapuera 商场店，后开）**。
- 周报：用户说「做周报」→ 按 `.claude/skills/store-report/SKILL.md` 做（脚本 `analytics/store_report/store_report.py`，
  目标在 `analytics/store_report/targets/<门店>/<YYYY-MM>.json`）。
  2026-10-25 定稿：周一~周日分周（导出从 W1 的周一开始）、客次 = 到店次数、权益金含线上卖券并拆四类、人效、新客成交率只算新客 / 会员复购率。会员 = 客户编号字母+数字，新客 = 纯数字。
- 新客增长计划：`analytics/plans/make_growth_plan.py` 生成 Excel（当前草稿 v7，含 10/29 Morumbi VIP 沙龙和 11 月活动决定）。
- 收银明细含客户姓名/电话：输入和输出 Excel 都不入库（`reports/` 已 gitignore）。

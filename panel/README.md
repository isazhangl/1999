# 项目总控台（panel）

`project-hub.html` 是已发布页面的源码：https://claude.ai/artifact/Veaga83VgWD2yyhtiNt1Km
（`#today` / `#plan` / `#all` 直达三个视图）。改完用 Artifact 工具以同一文件路径重新发布。

## 数据

数据存在该 artifact 的数据库里（不在这个文件里），集合 `projects`，每家公司一个文档：

- 基本信息：`name` `tagline` `stage` `stageNote` `roleType` `role` `roleNote` `focus` `health` `milestone{title,date}` `links[]` `notes`
- 规划：`objectives[{id,title,period,status,krs[{id,title,current,target,unit}]}]`
  `initiatives[{id,title,due,objectiveId}]`（小项目，`objectiveId` 为空 = 没挂目标）
- 待办：`tasks[{id,text,done,doneAt,due,today,initiativeId,obs,createdAt}]`（`obs` = 用户写的结论/备注）

读写用 `ArtifactData`，写之前先读，带 `if_version`；用户会直接在页面里改，不要覆盖。

## 每日提醒

定时任务「每日待办提醒」（`trig_01Ro5yroGaBwAbRr8BLxwjY9`），每天 07:53 圣保罗时间，
只读数据库，把当天清单推送到手机。

## Meta 投放（「投放」页）

- 公司文档里的 `metaAccount`（广告账户 ID，1999 = `970593382756180`）决定哪些公司出现在「投放」页。
- 页面用 `mcp` capability 以查看者身份调用 Windsor.ai 的 `get_data`（manifest 只声明这一个工具）：
  `connector: "facebook"`、`date_preset: "last_30d"`、按天 × 广告系列取
  `spend` `impressions` `clicks` `link_clicks` `actions_omni_purchase` `action_values_omni_purchase` `campaign_status` `account_currency`，
  7 / 14 / 30 天在页面里自己汇总。返回 `{status, data[], total_rows}`；`status: "pending"` 时按 `poll_after_seconds` 再取。
- 每日提醒的定时任务不带连接器，所以推送里没有投放数据。

# 项目总控台（panel）

`project-hub.html` 是已发布页面的源码：https://claude.ai/artifact/Veaga83VgWD2yyhtiNt1Km
（`#today` / `#plan` / `#all` 直达三个视图）。改完用 Artifact 工具以同一文件路径重新发布。

## 数据

数据存在该 artifact 的数据库里（不在这个文件里），集合 `projects`，每家公司一个文档：

- 基本信息：`name` `tagline` `stage` `stageNote` `roleType` `role` `roleNote` `focus` `health` `milestone{title,date}` `links[]` `notes`
- 规划：`objectives[{id,title,period,status,krs[{id,title,current,target,unit}]}]`
  `initiatives[{id,title,due,objectiveId}]`（小项目，`objectiveId` 为空 = 没挂目标）
- 待办：`tasks[{id,text,done,doneAt,due,today,initiativeId,createdAt}]`

读写用 `ArtifactData`，写之前先读，带 `if_version`；用户会直接在页面里改，不要覆盖。

## 每日提醒

定时任务「每日待办提醒」（`trig_01Ro5yroGaBwAbRr8BLxwjY9`），每天 07:53 圣保罗时间，
只读数据库，把当天清单推送到手机。

## 待做：接入 Meta Ads 投放数据

用户想在面板上看 Meta Ads 的投放结果（1999 的广告账户在 Ads Manager 链接里，`act=970593382756180`）。
当时 claude.ai 里的 Windsor.ai 连接器没有完成授权，所以还没做。授权完成后在新会话里：

1. 先读一次 Windsor.ai 工具（`get_accounts`、`get_fields`、`get_data`）的输入结构，并用只读调用看清返回格式。
2. 页面用 `mcp` capability 以查看者身份调用 Windsor.ai（manifest 里只列用到的工具），
   在「全部公司」和「规划」的 1999 下显示近 7 天 / 30 天花费、展示、点击、CTR、CPC、购买、ROAS，以及按广告系列的明细，并标明数据更新时间。
3. 不要把真实返回值写进页面源码。

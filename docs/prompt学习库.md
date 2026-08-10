# 1999 · Prompt 学习库（出图复盘）

收录**验证成功**的出图 prompt,持续积累 + 定期复盘,沉淀成 1999 自己的出图配方。

**用法**:每当一张图创始人确认"成功",就把它的 prompt 按下方模板收进"成功案例";失败的教训进"避坑";我定期把共性规律提炼进"复盘·规律"。

> 通用参数:平台 **Riverflow freestyle** · 场景图模型 **`nano-banana-2`** · 参考图 `reference_asset_ids`(用真实商品照)· 竖版社媒 `4:5`。
> #1 祖母绿三石戒参考图 asset:`6bf16af5-66f4-403e-9cfd-4eb416a50fc8`。

---

## 一、成功案例

### ✅ #001 · 粉色马蹄莲 × #1 祖母绿三石戒(社媒场景图)

- **日期**:2026-08-10
- **用途**:社媒/场景图(生活方式)
- **模型 / 画幅**:nano-banana-2 · 4:5 · 2K
- **参考图**:#1 真实商品照(`6bf16af5-…`)
- **成图 asset**:`329877bf-e99c-4754-8987-475994289a24`

**Prompt(原文)**
```
Editorial macro jewelry photograph. A single large pink calla lily fills most of the frame — its curled petal and thick green stem are life-size and dominant. The diamond ring from the reference image rests threaded naturally around the curving stem, but it is REALISTICALLY SMALL in scale: about the true size of a real engagement ring against a real full-size flower, only roughly one-sixth the height of the bloom — do NOT oversize the ring. The ring is tilted at a casual, slightly off-center angle, nestled where the petal meets the stem, NOT perfectly centered or standing upright. Natural asymmetric composition, shallow depth of field: the ring in crisp focus while the petals fall softly out of focus. Soft directional natural window light, gentle realistic shadows, a hint of dew, subtle organic imperfections. Pure white background, high-end editorial jewelry photography, hyper-realistic textures, cinematic yet candid — captured, not staged. Portrait orientation.
```

**为什么成功(可复用要点)**
- **花型简单**:马蹄莲单瓣 + 光滑花茎,模型好控制,不易糊。
- **戒指绕茎摆放**:比"平放/悬浮"更自然,读起来像随手一放。
- **真实小比例**:`REALISTICALLY SMALL … one-sixth the height` 有效压住了戒指过大。
- **反 AI 感三件套**:歪着偏心(off-center/tilted/NOT centered)+ 浅景深 + 自然光影/露珠/小瑕疵。
- **大理石台面 + 纯白底**:高级、干净。

---

## 二、避坑（失败教训)

| 现象 | 出现在 | 原因 | 对策 |
|---|---|---|---|
| 戒指**乱加钻**(3 石变 4–5 石) | 玫瑰 v3/v4 | nano-banana 对 reference"再创作",prompt 数量约束压不住 | 用**多角度真图**做 reference;或走商品图 product edit 线;复杂花更易触发 |
| 花瓣**变布料/绸缎质感** | 玫瑰 v4 | 玫瑰多层花瓣复杂,模型易糊成织物 | 优先**造型简洁的花**(马蹄莲/郁金香/兰花);玫瑰慎用 |
| 变成**上手双戒** | 玫瑰 v2 | "nestled/cradled" + reference 把模型带向手部 | 明确 `NO hands, NO fingers, only the flower and one ring` |
| 戒指**过大**(比例失真) | 第一批 4 张 | 未约束比例 | 加 `REALISTICALLY SMALL, ~1/6–1/7 花尺寸, do NOT oversize` |
| 戒指**立起来 / 无支撑** | 桃子 v2、宽景 v1 | 未指定卧姿与支撑 | `the ring lies reclined and FLAT, resting on its side against the fruit with a clear point of support, NOT upright, NOT floating` |
| 产品**直面镜头**(呆板) | 宽景 v2 | 未指定角度 | 场景图一律 `turned at a three-quarter angle to camera, do NOT face the camera dead-on` |
| 道具**不连贯**(桌上一滩巧克力,桃子上却没有) | 宽景 v1 | 未约束物理来源 | `chocolate drips down the side of the peach onto the surface forming a CONNECTED puddle (physically continuous, flows from the fruit), not a separate pool` |
| 酱汁/巧克力**铺太满太刻意** | 桃子 v2 | 未限量 | `chocolate applied partially and naturally, not covering everything` |

---

## 三、复盘 · 规律（定期更新)

> 每积累几条成功案例,我把共性规律提炼到这里。

- **花型复杂度 = 成败关键**:简洁花(马蹄莲)成功率高;复杂多瓣花(玫瑰)易糊 + 易乱加钻。
- **戒指忠实度**:单张 reference 压不住镶嵌细节 → 重点款用**多角度真图**参考。
- **去 AI 感通用杠杆**:真实比例 + 不对称摆放 + 浅景深 + 自然光影/瑕疵(详见 `出图流程与环境配置.md` 第四之三节配方)。

---

_收录规则:创始人确认"成功"→ 收进"成功案例";我定期复盘更新"规律"。_

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

### ✅ #002 · 桃子 + 巧克力"感官系" × #1 祖母绿三石戒(生活场景图)

- **日期**:2026-08-10
- **用途**:社媒/生活场景图(高级感官系,换调性)
- **模型 / 画幅**:nano-banana-2 · 4:5 · 2K
- **参考图**:#1 真实商品照(`6bf16af5-…`)
- **成图 asset**:`2ddb4432-4467-4080-b22d-b4d0e5702dbf`(v5)

**Prompt(原文)**
```
High-end gourmet still-life, WIDE medium shot with lots of clean negative space (NOT a close-up) — the whole tabletop arrangement is visible and the ring is just ONE SMALL element within it. SCALE IS CRITICAL: a real engagement ring is only about 2 cm wide and the peach is about 7 cm, so the ring must appear that small — roughly one-third the width of the peach half and occupying only about 10% of the frame width. Do NOT oversize the ring; if in doubt make it smaller. SCENE: a halved ripe peach with glossy dark chocolate over its cut face, dripping down the side into a CONNECTED puddle on the white marble surface (the puddle clearly flows from the fruit, physically continuous). The diamond ring from the reference lies RECLINED on its side, rotated about 30 degrees back on its horizontal axis so it leans back and lies low — NOT standing upright, NOT facing the camera dead-on — resting with a clear support point against the base of the peach, turned to a three-quarter angle. Scale-anchor / styling props: one whole ripe peach and a few shards of dark chocolate nearby. RING design IDENTICAL to reference: a THREE-STONE ring, EXACTLY three diamonds total (one emerald-cut step-cut center + exactly ONE tapered baguette side stone on each side), double claw prongs, white gold band; do NOT add extra stones. Cinematic dramatic lighting, hyper-realistic glossy wet textures, sharp reflections, shallow depth of field with the small ring in crisp focus. Pure white background. Portrait orientation.
```

**为什么成功(可复用要点)** —— 经过 v1→v5 五轮迭代才收敛,关键四招:
- **比例**:`WIDE shot` + 尺度锚点(整桃)+ 明确尺寸/占比(`~2cm ring, ~7cm peach, ~10% frame width, do NOT oversize`)→ 戒指终于显小。
- **平放**:`RECLINED, rotated ~30° back on its horizontal axis, lies low, NOT upright`→ 压住"立正"倾向。
- **不正对**:`three-quarter angle, NOT facing the camera dead-on`。
- **物理连贯**:`chocolate drips down into a CONNECTED puddle that flows from the fruit`→ 桌上那滩有来源。
- **3 石锁定**:`EXACTLY three diamonds … do NOT add extra stones`(拉宽后模型更稳,数量也更容易对)。

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

- **花型/物体复杂度 = 成败关键**:简洁物体(马蹄莲、桃子)成功率高;复杂多瓣花(玫瑰)易糊成布料 + 易乱加钻。
- **戒指忠实度**:单张 reference 压不住镶嵌细节(常见"3 石→5 石")→ 重点款用**多角度真图**参考;每张必做**数钻自检**。
- **比例控制(生活场景图核心)**:AI 默认把首饰画大,靠 ①画面拉宽 ②尺度锚点(已知尺寸日常物)③写实际尺寸+占比(`~2cm, ~10% frame width`)。拉宽同时还顺带**降低乱加钻**——远景下模型更稳。
- **摆放真实感**:戒指要**侧躺/横轴后仰(~30°)有支撑**,别立正、别悬浮、别正对镜头(场景图一律 three-quarter)。
- **物理连贯**:溢出物(巧克力/液体)必须**有来源**,桌面的一滩要从主体滴落而来。
- **去 AI 感通用杠杆**:真实比例 + 不对称摆放 + 浅景深 + 自然光影/瑕疵(配方详见 `出图流程与环境配置.md` 第四之三节)。
- **迭代规律**:一个新场景常需 3–5 轮收敛(比例→摆放→连贯→数钻),成功后即收录本库,避免重复踩坑。

---

_收录规则:创始人确认"成功"→ 收进"成功案例";我定期复盘更新"规律"。_

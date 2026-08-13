# 1999 · Midjourney 速用卡

MJ 折腾,这张卡照着做就行:**5 步流程 + 安全词表 + 3 个模板(改词即用)**。
配套:忠实商品图/场景图走 Riverflow(见 [`prompt学习库.md`](./prompt学习库.md)),MJ 只做**上手/模特 campaign**。

---

## 何时用 MJ / 何时用 Riverflow
| 用途 | 用哪个 |
|---|---|
| 商品图(白底 3 角度)、花/水果场景静物 | **Riverflow**(简单、忠实、无审核) |
| 上手 / 模特佩戴 campaign 大片 | **Midjourney**(折腾但这类最强) |

---

## 死记 5 步
1. **设置里一次选好**:Version **V8.2** + **Style Raw**(以后不用每次写)。
2. **戒指图拖进「Omni Reference」框**(最右;不是 Image Prompts)。
3. **文字只写场景描述** —— **别打 `--oref`、别打 `--v`**(槽和设置已管)。
4. **结尾固定加**:`--ar 2:3 --ow 300 --no deformed hands, extra fingers`
5. **避雷词**(见下表),别写 sensual / nude 之类。

---

## 安全词表(避审核)
| 你想表达 | ❌ 别用(易被拦) | ✅ 换成 |
|---|---|---|
| 性感 / 亲密 | sensual, intimate, seductive, body | elegant, serene, refined, quiet luxury |
| 裸色 / 裸 | nude | soft-pink, blush, warm beige |
| 皮肤大特写堆料 | face+skin+close 堆一起 | 强调 "**jewelry** photograph",多描述戒指少描述脸 |
> 被拦了先**再点一次**(审核是概率性的),还不行就换词 / 改用下面**模板 C(无脸)**。

---

## 3 个模板(把 `[方括号]` 换成你的款)

### 模板 A · 上手贴脸(暖调美妆感)
```
close-up editorial jewelry photograph, a model's hand raised gently to her face, fingertips near her temple, part of her face softly framed by her hand, one eye and natural brow visible, warm golden-brown skin with realistic texture, glossy soft-pink manicure on short almond nails, wearing [a slim tapered gold band with a horizontal east-west marquise diamond], elegant serene quiet-luxury mood, soft warm golden-hour light, deep warm background in shadow, cinematic shallow depth of field focused on the ring, 85mm f1.4, subtle film grain, hyper-realistic --ar 9:16 --style raw --ow 300 --no deformed hands, extra fingers, text, watermark
```

### 模板 B · 黑西装(安静奢华编辑风)
```
editorial jewelry campaign, upper body of an elegant model in a structured black blazer, one hand resting flat over the black fabric showing [a delicate gold ring with a warm champagne cushion-cut diamond], other hand raised with fingertips near her lips wearing a fine gold two-stone ring, warm natural skin, short natural nails, brown hair pulled back, clean light grey background, soft even daylight, quiet-luxury minimalist mood, shallow depth of field on the rings, 50mm, film grain, hyper-realistic --ar 2:3 --style raw --ow 300 --no deformed hands, extra fingers, text, watermark
```

### 模板 C · 纯手部特写(最安全:无脸,避审核 + 避画坏脸)
```
close-up luxury jewelry campaign, an elegant hand against [a structured black blazer / warm skin / cream silk], showing [a slim gold band with a round solitaire diamond], short natural nails, soft warm daylight, minimalist quiet-luxury, cinematic shallow depth of field on the ring, 85mm f1.4, film grain, hyper-realistic --ar 4:5 --style raw --ow 300 --no deformed hands, extra fingers
```

**换款子弹**(填进 `[戒款]`):
- `a slim tapered gold band with a horizontal east-west marquise diamond`
- `a gold six-prong solitaire ring with a round diamond`
- `a delicate gold ring with a warm champagne cushion-cut diamond`
- `a gold three-stone ring with an emerald-cut center and tapered baguette sides`
- `fine gold huggie hoop earrings with pavé diamonds`(上耳:把 hand 换成 ear/neck 场景)

---

## 参数速查
| 参数 | 作用 |
|---|---|
| **Omni Reference 槽** | 放戒指图 = 锁"物体的形"(**你的戒指靠它**) |
| `--ar` | 画幅:`9:16` 竖 / `2:3` 竖 / `4:5` 帖 / `1:1` 方 |
| `--style raw` | 去 MJ 滤镜,更像真实摄影 |
| `--ow 100–1000` | Omni 权重(默认100);戒指要像调 **200–400**,场景崩了就调低 |
| `--no …` | 排除项:`deformed hands, extra fingers, text, watermark` |
| 别写 | `--oref` / `--v`(用槽 + 设置,别在文字里写) |

---

## 翻车急救
| 报错 / 现象 | 原因 | 解 |
|---|---|---|
| **Character Reference is missing** | 文字写了 `--oref` 但 Omni 槽没放图 | 删掉文字里的 `--oref`,把戒指图**拖进 Omni Reference 槽** |
| **AI Moderator is unsure** | 有触发词 | 换安全词 / 再点一次 / 改用**模板 C 无脸** |
| **手畸形 / 多指** | MJ 老难点 | 加 `--no deformed hands, extra fingers`;多抽几张挑;或用**模板 C** |
| 戒指不够像 | Omni 权重低 | `--ow` 调到 300–500;戒指图拍清楚、单枚特写当参考 |

> ⚠️ MJ 是"神似不是复刻"。要**精确某个 SKU** → 用 Riverflow;MJ 负责**氛围大片**。

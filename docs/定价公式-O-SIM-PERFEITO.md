# O SIM PERFEITO · 定价公式（草案 · 待创始人微调）

配置线钻戒定价 = 「镶嵌底价 + 主石(按克拉) 」×「金属系数」，再做心理取整并设最低价。
所有产品先以 **DRAFT** 上架，价格随时可批量重算调整。

---

## 公式

```
Preço (R$) = arredondar( ( Base_engaste + Pedra_central[quilate] ) × Fator_metal )
piso mínimo = R$ 1.990
arredondar → 最近的 R$100，再减 10（价格尾数统一为 X.X90）
```

- **Base_engaste**：镶嵌款式底价（含工费/金重/碎钻·melee/侧石），以 18K 金为基准。
- **Pedra_central**：主石零售贡献，按克拉档。
- **Fator_metal**：金属材料系数（银最低 → 铂金最高）。

---

## 参数表（可调）

### 1) 主石 · 按克拉（R$）
| 克拉 | 0.5 | 0.7 | 1.0 | 1.5 | 2.0 | 2.5 | 3.0 |
|---|---|---|---|---|---|---|---|
| Pedra_central | 1.400 | 2.000 | 3.200 | 5.400 | 8.000 | 11.500 | 15.500 |

### 2) 金属系数
| 代码 | 金属 | Fator |
|---|---|---|
| `AG` | Prata 925 | 0.55 |
| `OB` | Ouro 18K branco | 1.00 |
| `OA` | Ouro 18K amarelo | 1.00 |
| `OR` | Ouro 18K rosé | 1.00 |
| `PL` | Platina Pt950 | 1.35 |

### 3) 镶嵌底价 Base_engaste（18K 基准，R$）
| 款式码 | 名称 | Base |
|---|---|---|
| `SOL` | Solitário Puro | 1.600 |
| `ENL` | Enlace | 1.900 |
| `BZL` | Aura (bezel) | 1.900 |
| `CAT` | Catedral | 1.900 |
| `INF` | Infinito | 2.000 |
| `TRG` | Trilogia Gota | 3.200 |
| `TRE` | Trilogia Esmeralda | 3.200 |
| `TRB` | Trilogia Baguete | 3.200 |
| `TRR` | Trilogia Redonda | 3.200 |
| `HAL` | Halo | 2.600 |
| `HAF` | Halo Flor | 2.800 |
| `HAP` | Halo Pavé | 3.000 |
| `HAV` | Halo Vinha | 3.000 |
| `HHL` | Hidden Halo | 2.600 |
| `HHP` | Halo Oculto Pavé | 3.000 |
| `HHC` | Halo Oculto Catedral | 3.000 |
| `PAV` | Pavé | 2.400 |
| `VIN` | Vinha | 2.600 |
| `CRO` | Coroa | 2.600 |

---

## 价格锚点（抽样验算）
| SKU 组合 | 计算 | 价格 R$ |
|---|---|---|
| SOL · 0.5ct · Prata (AG) | (1600+1400)×0.55=1650 → 取整→ **piso** | 1.990 |
| SOL · 1.0ct · Ouro branco (OB) | (1600+3200)×1.0=4800 | 4.790 |
| SOL · 1.0ct · Platina (PL) | 4800×1.35=6480 | 6.490 |
| HAP · 2.0ct · Ouro amarelo (OA) | (3000+8000)×1.0=11000 | 10.990 |
| TRG · 3.0ct · Platina (PL) | (3200+15500)×1.35=25245 | 25.190 |

> 与 1999 入门线（0.5ct 银戒固定 R$1.999）自然衔接：O Sim Perfeito 的 0.5ct 银款落在最低价 R$1.990，往上按克拉/金属/款式递增。

---

## 使用说明
- 生成脚本：`scripts/gen_osp_variants.py`（读取本表参数 → 输出各款式 350 个变体：形状10 × 克拉7 × 金属5）。
- 微调只需改本文件的三张参数表 → 重跑脚本 → 批量更新 Shopify 价格。

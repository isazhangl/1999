# 1999 · 产品架构与 SKU 体系

顶层 = **系列 × 品类** 双维度。本文件定义三大系列、品类、变体维度与 **SKU 编码规则**。

---

## 一、系列 × 品类

| 系列 | 性质 | 品类 | 配置维度 |
|---|---|---|---|
| **O SIM PERFEITO**(钻戒) | 定制 · 可配置 | **只戒指** | 款式 × 形状(10) × 克拉(0.5–3ct,7 档) × 金属(5) |
| **FOREVER MINE**(日常) | 成品 + 极少可选 | 耳饰 / 项链·吊坠 / 手链 / 戒指 | 逐款独立 SKU;个别款可选金色或尺寸 |
| **1999**(入门) | 固定套 · **只银 S925 · 圆** | 耳饰 / 吊坠 / 戒指 / 手链 | 固定规格,不配置 |

---

## 二、SKU 编码规则

**通用格式**:`[系列]-[品类]-[款式]-[形状]-[克拉]-[金属]`(用不到的段省略)

### 码表
| 维度 | 代码 |
|---|---|
| **系列** | `SP` O Sim Perfeito · `FM` Forever Mine · `99` 1999 入门 |
| **品类** | `AN` anel(戒)· `BR` brinco(耳)· `CO` colar(项链)· `PI` pingente(吊坠)· `PU` pulseira(手链) |
| **形状** | `RD` 圆 · `OV` 椭圆 · `EM` 祖母绿 · `MQ` 马眼 · `GO` 水滴 · `PR` 公主 · `RA` 雷迪恩 · `AL` 垫形 · `AS` asscher · `CR` 心形 |
| **克拉** | ×100:`030` `050` `070` `100` `150` `200` `250` `300` |
| **金属** | `AG` 银925 · `OB` 18K白 · `OA` 18K黄 · `OR` 18K玫瑰 · `PL` 铂金Pt950 |
| **钻戒款式** | 见 §三 款式表(SOL/TRS/HAL/HHL/PAV… 持续补充) |

> 💡 **金属出图**:白 18K(`OB`)、铂金(`PL`)、银(`AG`)视觉同为白色系,**共用 1 张「白色金属」图**;黄金(`OA`)、玫瑰金(`OR`)各 1 张 → 每款只需 **3 张金属图**(白 / 黄 / 玫瑰),不是 5 张。

### 各系列用法
- **O SIM PERFEITO(配置线)**:`SP-AN-[款式]-[形状]-[克拉]-[金属]`
  - 例:单石·圆·1ct·18K白 → `SP-AN-SOL-RD-100-OB`;三石·椭圆·1.5ct·铂金 → `SP-AN-TRS-OV-150-PL`
- **1999(固定线)**:`99-[品类]-[克拉]`(金属恒 AG、形状恒 RD,省略)
- **FOREVER MINE(成品线)**:`FM-[品类]-[款号3位]-[可选金属/尺寸]`
  - 例:`FM-BR-001`(第 1 款耳饰)、`FM-CO-014-OA`(第 14 款项链·黄金变体)

---

## 三、O SIM PERFEITO · 钻戒镶嵌款式表（持续补充）

> 创始人发款式图 → 命名(葡/中)+ 分配款式码。

| 款式码 | 名称(PT / 中) | 说明 | 状态 |
|---|---|---|---|
| `SOL` | Solitário Puro / 至纯单石 | 细圈平臂 · 四爪 · 最经典 | ✅ |
| `ENL` | Enlace / 缠绕 | 戒臂不对称环绕主石(bypass/扭臂) | ✅ |
| `BZL` | Aura / 包边光环 | 全金属包边镶(bezel)· 平臂 | ✅ |
| `CAT` | Catedral / 教堂拱肩 | 戒臂上扬拱肩(cathedral)· 四爪 | ✅ |
| `INF` | Infinito / 无限交织 | 戒臂分叉交叉于底(crossover/split) | ✅ |
| `TRS` | Trilogia / 三石(总称) | 主石 + 两侧石(下列按侧石形状细分) | ✅ |
| `TRG` | Trilogia Gota / 三石·梨形侧石 | 侧石梨形(pear) | ✅ |
| `TRE` | Trilogia Esmeralda / 三石·祖母绿侧石 | 侧石祖母绿/长方阶梯 | ✅ |
| `TRB` | Trilogia Baguete / 三石·锥形长方侧石 | 侧石锥形长方(tapered baguette) | ✅ |
| `TRR` | Trilogia Redonda / 三石·圆形侧石 | 侧石圆形(round) | ✅ |
| `HAL` | Halo / 光环 | 主石周围一圈碎钻光环 | ✅ |
| `HAF` | Halo Flor / 花形光环 | 花瓣/扇贝形复古光环 | ✅ |
| `HAP` | Halo Pavé / 光环+碎钻臂 | 完整光环 + 戒臂 pavé | ✅ |
| `HAV` | Halo Vinha / 光环+缠枝臂 | 完整光环 + 扭转藤蔓 pavé 臂 | ✅ |
| `HHL` | Hidden Halo / 隐藏光环 | 侧面才可见的隐藏光环 · 平臂 | ✅ |
| `HHP` | Halo Oculto Pavé / 隐藏光环+碎钻臂 | 隐藏光环 + 戒臂 pavé | ✅ |
| `HHC` | Halo Oculto Catedral / 隐藏光环·拱肩 | 隐藏光环 + 拱肩(cathedral) | ✅ |
| `PAV` | Pavé / 戒臂镶钻 | 戒臂排镶碎钻(无光环) | ✅ |
| `VIN` | Vinha / 缠枝碎钻 | 戒臂扭转如藤蔓 + pavé | ✅ |
| `CRO` | Coroa / 皇冠(V形) | pavé 呈 V/chevron 上扬拥托主石 | ✅ |
| … | (待命名) | 发图后补 | ⏳ |

> `SOL`/`ENL`/`BZL`/`CAT`/`INF` 同属**单石家族**,区别在戒臂/镶法;形状与金属为独立 SKU 维度。

---

## 四、1999 入门线 · 固定 SKU(银 S925 · 圆)

| SKU | 品类 | 规格 |
|---|---|---|
| `99-BR-030` | 耳钉(一对) | 银 · 圆 · **0.3ct/颗** |
| `99-PI-050` | 吊坠(无链) | 银 · 圆 · 0.5ct |
| `99-AN-050` | 戒指 | 银 · 圆 · 0.5ct |
| `99-PU-050` | 手链 | 银 · 圆 · 0.5ct |

---

---

## 五、FOREVER MINE · 成品样板（已在 Shopify 建档,DRAFT）

> 逐款独立 SKU,少量可选(金属/尺寸)。以下为首批样板,后续按真实成品清单扩充。

| SKU | 品类 | 名称(PT) | 可选项 | 价 R$ |
|---|---|---|---|---|
| `FM-CO-001` | 项链 | Colar Ponto de Luz | — | 2.490 |
| `FM-BR-001` | 耳饰 | Brincos Argola Cravejados | 金属(白/黄/玫瑰) | 1.290 |
| `FM-PU-001` | 手链 | Pulseira Tennis Clássica | — | 3.990 |
| `FM-AN-001` | 戒指 | Aliança Meia-Eternidade | 尺寸(14/16/18/20/22) | 2.490 |

---

## 六、Shopify 登记状态

- **定价公式**:见 [`定价公式-O-SIM-PERFEITO.md`](./定价公式-O-SIM-PERFEITO.md)(创始人可微调三张参数表 → 重跑 `scripts/gen_osp_variants.py` → 批量改价)。
- **系列 = Coleção(智能集合,按 TAG 归类)**:
  - `Coleção 1999`(tag `colecao-1999`)· `O Sim Perfeito`(tag `o-sim-perfeito`)· `Forever Mine`(tag `forever-mine`)。
- **O SIM PERFEITO**:19 个镶嵌款式 = 19 个商品,每个 **形状10 × 克拉7 × 金属5 = 350 变体**,合计 **6.650 变体**,状态 DRAFT,按公式定价。图片后续生成再上传。
- **1999 入门线**:4 个固定商品(耳钉/吊坠/戒指/手链),全 **R$1.999**。
- 详见 [`Shopify登记状态.md`](./Shopify登记状态.md)。

---

## 待补 / 待办
- [ ] **钻戒款式**:创始人陆续发图 → 命名 + 配码,补进 §三。
- [ ] **FOREVER MINE**:按真实成品清单继续扩充逐款 SKU(品类/款号/可选项/价格)。
- [ ] 视觉图更正:1999 耳钉图注 0.5ct → **0.3ct**(见 `docs/images/1999-series/`)。
- [ ] **图片**:各 SKU 电商效果图生成后,回传 Shopify(需公网 HTTPS 图址)。
- [ ] 旧演示商品清理:Shopify 上有 9 个早期 demo 戒指(智能集合 `o sim perfeito`,按 TYPE=Anel de Noivado),待创始人确认是否删除/归档。

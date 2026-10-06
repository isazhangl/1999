# 1999 · SEO 指南（关键词 + 分类 + Shopify 设置）

> 市场：巴西 pt-BR｜品类：培育钻订婚戒（anéis de noivado com diamante cultivado）
> 目标：用「培育钻 + 款式/切工」长尾词抢转化，用「订婚戒」头部词做品牌曝光。
> 研究来源：Bloomberg Línea、Forbes BR、Casamentos.com.br、77 Diamonds PT、Bulgari PT、Swarovski PT、metadatareactor（Shopify SEO 2026）等。

---

## 一、关键词体系（三层）

### 1) 头部词（量大、竞争大）— 用于首页 / 主集合 / 品牌
- `anel de noivado`（★核心）
- `anel de noivado de diamante`
- `anel de compromisso`
- `aliança de noivado`（注意：巴西「aliança」多指对戒/婚戒，语义略不同，次要）

### 2) 材质差异化词（品牌命脉、必抢、竞争中）— 几乎每页都要带
- `diamante cultivado em laboratório`（★主）
- `diamante de laboratório`
- `diamante cultivado`
- `anel de noivado diamante cultivado`（★高意图长尾）
- `diamante sintético`（有人搜，调性低，只放描述里做兜底，**不进标题**）
- 信任词：`certificado IGI` / `com laudo` / `IGI`

### 3) 款式 / 切工词（转化最高、竞争最低）— 用于产品 / 款式集合 / 形状集合
| 中文 | pt-BR 搜索词（放 meta） | 备注 |
|---|---|---|
| 单钻 | `anel solitário` / `anel de noivado solitário` | 最经典、搜量大 |
| 光环 | `anel halo` / `anel de noivado halo` | vintage 风 |
| 隐藏光环 | `halo oculto` / `hidden halo` | 两个都放 |
| 三石 | `anel três pedras` + `anel trilogia` | **两词都要**：三石=大众词，trilogia=行业词 |
| 戒臂满钻 | `aro cravejado` / `pavê` | pavê 指密钉 |
| 圆形 | `diamante redondo` / `lapidação brilhante` | |
| 椭圆 | `diamante oval` | 显大、搜量高 |
| 祖母绿 | `diamante esmeralda` / `lapidação esmeralda` | |
| 公主方 | `diamante princesa` | |
| 马眼 | `diamante marquesa` / `marquise` | ⚠️ **不是 "navete"**，巴西搜 marquesa |
| 水滴/梨形 | `diamante gota` / `pera` | |
| 枕形 | `almofada` / `cushion` | 两个都放 |
| 雷迪恩 | `diamante radiante` | |
| Asscher | `asscher` | 小众，长尾 |
| 心形 | `diamante coração` | |

---

## 二、分类（SEO 信息架构）

SEO 友好的三级结构（和导航一致）：

```
anel de noivado (首页/品牌词)
├── 按款式 (estilo-*)  ← 主力 SEO 集合页
│   ├── Solitário      → anel de noivado solitário
│   ├── Halo           → anel de noivado halo
│   ├── Halo Oculto    → halo oculto / hidden halo
│   ├── Três Pedras    → anel três pedras / trilogia
│   └── Aro Trabalhado → aro cravejado / pavê
└── 按形状 (formato-*) ← 长尾 SEO 集合页
    ├── Redondo, Oval, Esmeralda, Marquesa(Navete), Gota,
        Princesa, Radiante, Almofada, Asscher, Coração
```

**原则**：集合页（category）比单品更容易排名——因为聚合了同类词。所以**集合页的 SEO 要最用心**。

> 命名提醒：「Navete」这个集合/产品，**显示名可保留 Navete**，但 **handle 和 meta 用 marquesa**（`/collections/marquesa`、meta 里写 "marquesa / marquise"），否则搜不到。

---

## 三、Shopify 到底怎么设（操作层）

每个产品 / 集合页，Shopify 后台底部都有 **「Listagem do mecanismo de pesquisa / Edição de SEO」**，三个字段：

1. **Título da página（Meta title）** — ≤ **60 字符**。公式：
   `[款式/形状] + [材质] + 品牌`
   例：`Anel de Noivado Solitário Oval · Diamante Cultivado`
   - 把最想排的词放最前面，不要用内部型号名。
2. **Descrição（Meta description）** — **150–160 字符**。公式：
   `一句卖点（切工+材质+laudo）+ 信任（IGI/garantia）+ CTA（fale no WhatsApp / sob encomenda）`
3. **URL e identificador（handle）** — 短、含关键词、用连字符：
   `anel-de-noivado-solitario-oval`（不要用随机 ID / 不要带克拉数字）

另外两处（常被忽略、很加分）：
- **图片 alt text**：每张产品图写 `anel de noivado [formato] diamante cultivado`（既帮 SEO 又帮无障碍）。
- **集合描述正文**：集合页顶部放 1–2 段含关键词的介绍文字（Google 读得到）。

---

## 四、每个集合页的 SEO（★优先做这个）

> Meta title 已控制在 ~60 字符内；描述 ~150–160。品牌后缀 `| 1999` 视长度加。

### 款式集合（estilo-*）
| 集合 | handle | Meta title | Meta description |
|---|---|---|---|
| Solitário | `estilo-solitario` | `Anel de Noivado Solitário · Diamante Cultivado` | `Anéis de noivado solitário com diamante cultivado em laboratório, cor D/E/F e laudo IGI. Feitos sob encomenda em São Paulo. Fale no WhatsApp.` |
| Halo | `estilo-halo` | `Anel de Noivado Halo · Diamante Cultivado em Lab` | `Anéis de noivado halo com diamante cultivado em laboratório, mais brilho e presença. Com laudo IGI, feitos sob encomenda. Monte o seu.` |
| Halo Oculto | `estilo-halo-oculto` | `Anel de Noivado Halo Oculto · Diamante Cultivado` | `Anéis de noivado com halo oculto (hidden halo) e diamante cultivado em laboratório, com laudo IGI. Sob encomenda em São Paulo. Fale no WhatsApp.` |
| Três Pedras / Trilogia | `estilo-trilogia` | `Anel de Noivado Três Pedras (Trilogia) · Diamante` | `Anéis de noivado três pedras (trilogia) com diamante cultivado em laboratório e laudo IGI. Passado, presente e futuro. Sob encomenda.` |
| Aro Trabalhado | `estilo-aro-trabalhado` | `Anel de Noivado Pavê e Aro Cravejado · Diamante` | `Anéis de noivado com aro cravejado e pavê de diamante cultivado em laboratório, com laudo IGI. Brilho em toda a peça. Sob encomenda.` |

### 形状集合（formato-*）
| 集合 | handle 建议 | Meta title | Meta description（模板，替换 [X]） |
|---|---|---|---|
| Redondo | `formato-redondo` | `Anel de Noivado Diamante Redondo · Cultivado` | `Anel de noivado com diamante redondo cultivado em laboratório, o clássico mais brilhante. Cor D/E/F, laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Oval | `formato-oval` | `Anel de Noivado Diamante Oval · Cultivado` | `Anel de noivado com diamante oval cultivado em laboratório — alonga os dedos. Cor D/E/F, laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Esmeralda | `formato-esmeralda` | `Anel de Noivado Diamante Esmeralda · Cultivado` | `Anel de noivado com diamante esmeralda cultivado em laboratório, brilho sereno e minimalista. Laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Marquesa (Navete) | `marquesa`（★改） | `Anel de Noivado Diamante Marquesa · Cultivado` | `Anel de noivado com diamante marquesa (marquise) cultivado em laboratório, alongado e marcante. Laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Gota | `formato-gota` | `Anel de Noivado Diamante Gota · Cultivado` | `Anel de noivado com diamante gota (pera) cultivado em laboratório. Laudo IGI, sob encomenda em São Paulo. Fale no WhatsApp.` |
| Princesa | `formato-princesa` | `Anel de Noivado Diamante Princesa · Cultivado` | `Anel de noivado com diamante princesa cultivado em laboratório, quadrado e de brilho intenso. Laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Radiante | `formato-radiante` | `Anel de Noivado Diamante Radiante · Cultivado` | `Anel de noivado com diamante radiante cultivado em laboratório, brilho vibrante. Laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Almofada | `formato-almofada` | `Anel de Noivado Diamante Almofada · Cultivado` | `Anel de noivado com diamante almofada (cushion) cultivado em laboratório. Laudo IGI, sob encomenda em São Paulo. Fale no WhatsApp.` |
| Asscher | `formato-asscher` | `Anel de Noivado Diamante Asscher · Cultivado` | `Anel de noivado com diamante Asscher cultivado em laboratório, art déco geométrico. Laudo IGI, sob encomenda. Fale no WhatsApp.` |
| Coração | `formato-coracao` | `Anel de Noivado Diamante Coração · Cultivado` | `Anel de noivado com diamante coração cultivado em laboratório, símbolo romântico. Laudo IGI, sob encomenda. Fale no WhatsApp.` |

---

## 五、单品 SEO（上线 6 款优先）

| 产品 | Meta title (≤60) | Meta description (150–160) |
|---|---|---|
| Solitário Redondo | `Anel de Noivado Solitário Redondo · Diamante Lab` | `Anel de noivado solitário com diamante redondo de 1 ct cultivado em laboratório, cor D/E/F, laudo IGI. Platina 950, sob encomenda. Fale no WhatsApp.` |
| Solitário Oval | `Anel de Noivado Solitário Oval · Diamante Lab` | `Anel de noivado solitário com diamante oval de 1 ct cultivado em laboratório, cor D/E/F, laudo IGI. Ouro 18K, sob encomenda. Fale no WhatsApp.` |
| Solitário Esmeralda | `Anel de Noivado Solitário Esmeralda · Diamante Lab` | `Anel de noivado solitário com diamante esmeralda de 1 ct cultivado em laboratório, laudo IGI. Ouro 18K branco, sob encomenda. Fale no WhatsApp.` |
| Solitário Coração | `Anel de Noivado Solitário Coração · Diamante Lab` | `Anel de noivado solitário com diamante coração de 1,5 ct cultivado em laboratório, cor D, laudo IGI. Platina 950, sob encomenda. Fale no WhatsApp.` |
| Radiante Aro Cravejado | `Anel de Noivado Radiante Aro Cravejado · Diamante` | `Anel de noivado com diamante radiante de 2 ct e aro cravejado, cultivado em laboratório, halo oculto, laudo IGI. Platina 950. Fale no WhatsApp.` |
| Oval Aro Cravejado | `Anel de Noivado Oval Aro Cravejado · Diamante Lab` | `Anel de noivado com diamante oval de 1 ct e aro cravejado, cultivado em laboratório, cor D, laudo IGI. Ouro 18K. Sob encomenda. Fale no WhatsApp.` |
| Três Pedras Esmeralda | `Anel de Noivado Três Pedras Esmeralda · Diamante` | `Anel de noivado três pedras (trilogia) com diamante esmeralda de 2 ct cultivado em laboratório, laudo IGI. Ouro 18K branco. Fale no WhatsApp.` |

> 其余款式戒（Halo、Pavé、Vinha 等）上架前按同公式补：`Anel de Noivado [款式] · Diamante Cultivado` + 描述。

---

## 六、落地优先级

1. ★★★ **款式集合 + 形状集合**（16 个页面）——最容易排名，先做。
2. ★★ **上线的 6 款单品** meta。
3. ★ 图片 alt text（全产品）。
4. Navete 集合 handle 改 `marquesa`（并在 meta 保留「marquesa / marquise」）。
5. 建一个 **blog/FAQ 页**答高搜问题（抢信息流量）：
   - "Diamante cultivado é de verdade?"
   - "Diamante cultivado vs natural: diferença e preço"
   - "Quanto custa um anel de noivado de diamante cultivado?"
   - "Como escolher o formato do diamante"
6. 技术面：确保每页 title/description 唯一；handle 不带随机 ID；提交 sitemap（Shopify 自动生成 `/sitemap.xml`）到 Google Search Console。

---

## 七、避坑清单
- ❌ 标题堆砌关键词（`anel anel noivado diamante barato...`）→ Google 降权。
- ❌ 多页用同一个 meta → 自相竞争。
- ❌ handle 里放克拉/随机码。
- ✅ 一页一个主词；描述自然、含 1 个信任点 + 1 个 CTA。
- ✅ `diamante cultivado em laboratório` 贯穿全站（这是你和天然钻店的最大区隔词）。

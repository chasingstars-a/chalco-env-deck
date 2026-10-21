---
name: chalco-env-deck
description: 用「绿色环保企业」商务模板的视觉体系制作中文 PPT：品牌绿渐变 + 品牌蓝点缀 + 微软雅黑，含封面、目录、章节页、内容页（要点/卡片/指标/表格/流程/时间线）、结尾页，以及成品自检脚本。当用户要做环保 / 节能 / 环境服务类公司的公司简介、对外汇报、市场材料、项目汇报，或要求"按绿色环保风格做 PPT""套这套绿色模板"时使用。
whenToUse: 需要产出或改写一份绿色环保企业风格的 .pptx（16:9、绿色主题、可替换徽标、微软雅黑），或需要复用该模板的配色、版式、素材结构时加载。做成 Office 文档之外的纯文字方案时不适用。
---

# 绿色环保企业风格 PPT

本技能把一版「绿色环保企业」公司简介模板**固化成了代码**：一套 python-pptx 版式库 + 占位素材 + 自检脚本。
不再手动改模板文件，而是写一个 30 行的构建脚本，把内容填进既定版式，产出同色调、同排版的成品。

素材全部是程序生成的占位图，换成自己公司的徽标/字标/照片只需覆盖 `assets/` 同名文件。

## 0. 铁律

1. **只用这套视觉**：品牌绿渐变 + 品牌蓝 + 白/浅蓝底 + 微软雅黑；**不要**引入红色强调、其他字体、其他色系（旧版模板里的红色加粗是历史包袱，不复刻）。
2. **logo 只能是素材里的徽标**：写 logo 用 `logo(slide)` / `asset("logo_emblem.png")`，不重绘、不变形、不加边框、不调色；位置固定在右上角（内容页）或中上（封面）。
3. **16:9 固定**：13.333 × 7.5 英寸，所有坐标单位为英寸，内容不出血、不越界。
4. **先读设计规范再动手**：配色、字号、栅格、版式图鉴见 `references/design-spec.md` 与 `references/layouts.md`，改动版式前先看它们。
5. **交付前必须自检**：`scripts/check_deck.py` 必须 0 error；有视觉要求时再渲染抽查（见第 5 步）。

## 1. 什么时候用哪种页面

| 需求 | 调用 |
|---|---|
| 封面（有主标题） | `Deck.cover(title=..., subtitle=..., location=..., date=...)` |
| 封面（纯公司形象，无标题） | `Deck.cover()` |
| 目录（3~6 条） | `Deck.toc([(标题, 英文), ...], subtitle=...)` |
| 章节过渡页 | `Deck.section("01", "公司简介", en="COMPANY PROFILE", photo=...)` |
| 普通内容页 | `p = Deck.page("标题")` 后用 `p.bullets / p.cards / p.table / ...` |
| 结尾页 | `Deck.closing(company=..., address=...)` |

内容页可用的块（详见 `references/layouts.md`）：
`bullets` 要点、`text` 段落、`image` 配图、`cards` 三卡（可带图）、`card` 单卡、
`kpis` 指标卡、`table` 表格、`process` 箭头流程、`banner` 尖头横幅、`tags` 标签、
`note` 提示条、`timeline` 时间线、`chart` 原生图表。

## 2. 标准工作流

1. **问清三件事**（若用户没给）：这份 PPT 给谁看、需要几页/几个章节、有没有必须出现的数据与图片。
2. **搭脚本**：在用户工作目录新建 `build_xxx.py`，按 `examples/build_demo.py` 的结构改。
   路径写法（Windows）：

   ```python
   import sys
   sys.path.insert(0, r"C:\Users\Lin\.dsh\skills\chalco-env-deck\scripts")
   from chalco_deck import Deck, C, asset, FONT
   ```

3. **逐页填内容**：先定骨架（封面 → 目录 → 每章 1 张章节页 → 内容页 3~6 张 → 结尾），再填字。
   内容页标题 4~10 个汉字；每页只讲一件事；每页要点 ≤ 5 条、每条 ≤ 2 行。
4. **跑脚本**：`python build_xxx.py`（用 `load_workspace_dependencies` 给的 Python），产出 .pptx。
5. **自检 + 视觉抽查**：

   ```powershell
   python "<skill>\scripts\check_deck.py" "<输出.pptx>"
   ```

   0 error 之后，若用户在意视觉，用 office-render 的 LibreOffice Kit 渲染 1~3 页抽查（图片、表格、文字是否被挤压）。
6. **交付**：`present` 交付 .pptx，并在回复里说明页数、章节结构、需要用户确认的数据项。

## 3. 关键 API 速查

```python
d = Deck()                                   # 16:9；主题字体已改为微软雅黑
d.cover(title="公司简介", subtitle="副标题", location="中国 · 北京", date="2026.10")
d.toc([("公司简介", "COMPANY PROFILE"), ("科技立企", "TECHNOLOGY")])
d.section("01", "公司简介", en="COMPANY PROFILE", photo=asset("photo_leaf.jpg"))
p = d.page("公司概况")                        # 背景 + 绿色箭头标题 + 右上 logo + 页码
p.bullets([["加粗开头", " 后续文字"], "第二条"], box=(0.95, 1.45, 6.5, 4.2), size=15)
p.image(asset("photo_leaf.jpg"), box=(7.9, 1.45, 4.4, 4.2), radius=0.14)
p.cards([dict(title="价值观", lines=["创新 共赢"], image=asset("bg_section.jpg"))],
        box=(0.95, 1.4, 11.43, 4.9))
p.kpis([("10", "亿元", "注册资本"), ("68", "家", "所属骨干企业")],
       box=(0.95, 1.4, 11.43, 1.95))
p.table(["能力", "说明"], [["技术研发", "……"]], box=(0.95, 3.6, 11.43, 2.5),
        col_widths=[2.0, 6.0, 3.2], col_align=["c", "l", "l"])
p.process([("技术研发", "标准与专利"), ("工程总包", "EPC 交付")], box=(0.95, 1.9, 11.43, 1.0))
# 5 列以上箭头把文字区让宽，否则标题会折行：
# p.process(steps, box=..., pad_left=0.16, pad_right=0.04, adj=0.28, title_size=13.5)
# desc 里可用 "\n" 手动断行，避免出现单字孤行
p.banner("示例集团「三大板块」高质量发展战略", box=(0.95, 3.35, 11.43, 0.52))
p.tags(["铝", "铜", "铅锌"], box=(0.95, 1.9))
p.note([[("说明  ", {"bold": True, "color": C.GREEN}), "提示文字"]], box=(0.95, 6.3, 11.43, 0.45))
p.timeline([("2019.01", "宁夏丰昊配售电", "配售电、新能源")], box=(0.95, 2.6, 5.6, 3.5))
d.closing()                                  # 致辞字标 + 底栏 + 二维码
d.save("output.pptx")
```

`C` 是调色板常量（`C.GREEN`、`C.MINT`、`C.BLUE`、`C.INK` …），**不要写死别的色号**；
确需新色先加进 `C` 与 `PALETTE`，否则自检会告警。

字号基准（内容页）：页标题 28、小标题 18~20、正文 15~16、注释 12~13、页码 14。

## 4. 素材（assets/）

| 文件 | 用途 |
|---|---|
| `logo_emblem.png` / `logo_emblem_white.png` | 占位徽标（彩色 / 反白），换成自家 logo 即可 |
| `logo_wordmark.png` | 「示例环保集团」绿色渐变字标 |
| `wordmark_mulu.png` / `wordmark_thanks.png` | 「目录」「分享结束 感谢聆听」字标 |
| `bg_cover_full.jpg` | 封面底图（天空+草地+地球，已预合成 16:9） |
| `bg_light.jpg` | 内页浅蓝光效底 |
| `bg_section.jpg` / `photo_leaf.jpg` | 章节页/配图用实景图 |
| `bg_closing.jpg` | 结尾页城市底图 |
| `qr_wechat.png` | 公司二维码（结尾页使用） |

素材均由原模板抽取（`tools/prepare_assets.py` 可复现），**不要**用网络图替换封面/结尾底图，以免破坏色调。

## 5. 自检与渲染

```powershell
# 结构 / 字体 / 配色 / logo / 页码 / 溢出
python "<skill>\scripts\check_deck.py" "out.pptx"

# 视觉抽查（LibreOffice Kit，注意 --pages 用逗号串）
node "<libreoffice-kit cli>" render --input "out.pptx" --output-dir ".preview" --pages "1,2,4"
```

`check_deck.py` 报 error 必须修；warning 里"文字可能溢出"要靠调小字号或加宽文本框解决，不要用 `spAutoFit` 糊过去。

## 6. 常见坑

- **图片带透明通道**：直接用 `pic()`，它会把 PNG 的 alpha 保留（不要自己 open/save 成 JPG，会变黑底）。
- **表格**：只用 `p.table()`，竖线会自动与单元格底色同色；自己插原生表格会出现深浅不一的框线。
- **文字换行**：不要在 run 里塞 `\n`，把多行交给 `text()` 的 list 或 `\n` 字符串（库会拆成独立段落，避免居中错位）。
- **徽标尺寸**：内容页固定 0.687×0.703 英寸、右上角 (11.91, 0.374)；放大或挪动都会破坏版式一致性。
- **旧模板的红色强调**：不要沿用。需要强调用 `C.GREEN` 加粗，或 `note()` 提示条。

# go-docx 功能全景分析与路线图

## 一、DOCX 格式常用功能全景

DOCX 文件（ECMA-376 / Office Open XML）是一个 zip 包，内部由多个 XML part 组成。以下按功能域列出最常用的能力：

---

### 1. 文本与段落（Paragraph / Run）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 段落创建 | `<w:p>` | ✅ 已支持 |
| 段落对齐（Justification） | `<w:jc>` start/center/end/both/distribute | ✅ 已支持 |
| 段落缩进（Indentation） | `<w:ind>` 首行/悬挂/左右缩进 | ✅ 结构已有，API 未封装 |
| 段落间距（Spacing） | `<w:spacing>` 段前/段后/行距 | ✅ 结构已有，API 未封装 |
| 段落样式引用 | `<w:pStyle>` | ✅ 已支持 `Style()` |
| 分页符 | `<w:br w:type="page">` | ✅ 已支持 `AddPageBreaks()` |
| Run 文本 | `<w:t>` | ✅ 已支持 `AddText()` |
| 粗体 / 斜体 | `<w:b>` / `<w:i>` | ✅ 已支持 |
| 下划线 | `<w:u>` | ✅ 已支持 |
| 删除线 | `<w:strike>` | ✅ 已支持 |
| 字体设置 | `<w:rFonts>` | ✅ 已支持 |
| 字号 | `<w:sz>` / `<w:szCs>` | ✅ 已支持 |
| 字体颜色 | `<w:color>` | ✅ 已支持 |
| 高亮 | `<w:highlight>` | ✅ 已支持 |
| 底纹 | `<w:shd>` | ✅ 已支持 |
| 上标/下标 | `<w:vertAlign>` | ✅ 结构已有（解析），API 未封装 |
| 字符间距 | `<w:spacing>` (run level) | ✅ 已支持 |
| Tab 字符 | `<w:tab>` | ✅ 已支持 |
| 换行符 | `<w:br>` | ✅ 已支持 |
| 编号/列表 | `<w:numPr>` | ✅ 已支持 `NumPr()` |
| **跨 Run 文本替换** | 占位符跨多个 run 时的查找替换 | ❌ 缺失 |
| **Run 合并** | 相同属性 run 合并 | ✅ 已支持 `MergeText()` |

### 2. 超链接

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 外部超链接 | `<w:hyperlink r:id>` + relationship | ✅ 已支持 `AddLink()` |
| 内部书签链接 | `<w:hyperlink w:anchor>` | ⚠️ 解析时部分支持，API 未封装 |

### 3. 表格（Table）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 表格创建 | `<w:tbl>` by row×col | ✅ 已支持 |
| 指定行高/列宽创建 | twips 单位 | ✅ 已支持 `AddTableTwips()` |
| 表格宽度 | `<w:tblW>` dxa/auto/pct | ✅ 已支持 |
| 表格对齐 | `<w:jc>` | ✅ 已支持 |
| 表格边框 | `<w:tblBorders>` | ✅ 已支持 |
| 表格定位 | `<w:tblpPr>` | ✅ 已支持（解析+结构） |
| 表格样式 | `<w:tblStyle>` | ✅ 结构支持（解析），API 未封装 |
| 表格外观控制 | `<w:tblLook>` | ✅ 已支持 |
| 单元格宽度 | `<w:tcW>` | ✅ 已支持 |
| 单元格边框 | `<w:tcBorders>` | ✅ 已支持（解析+结构） |
| 单元格底纹 | `<w:shd>` | ✅ 已支持 `Shade()` |
| 单元格内边距 | `<w:tcMar>` | ✅ 已支持 `Padding()` |
| 单元格垂直对齐 | `<w:vAlign>` | ✅ 已支持（解析） |
| **横向合并（GridSpan）** | `<w:gridSpan>` | ⚠️ 解析支持，无创建 API |
| **纵向合并（vMerge）** | `<w:vMerge>` | ⚠️ 解析支持，无创建 API |
| **重复表头** | `<w:tblHeader>` | ❌ 缺失 |
| **嵌套表格** | `<w:tbl>` inside `<w:tc>` | ✅ 已支持（解析） |
| **行高设置** | `<w:trHeight>` | ✅ 已支持 |
| **行对齐** | `<w:jc>` on row | ✅ 已支持 |
| 表格 Grid | `<w:tblGrid>` / `<w:gridCol>` | ✅ 已支持 |

### 4. 图片与绘图（Drawing）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 行内图片 | `<wp:inline>` | ✅ 已支持 `AddInlineDrawing()` |
| 浮动图片 | `<wp:anchor>` | ✅ 已支持 `AddAnchorDrawing()` |
| 图片尺寸控制 | EMU 单位 | ✅ 已支持 `Size()` |
| 图片自动格式检测 | imgsz 库 | ✅ 已支持 |
| 文字环绕 | wrapNone/wrapSquare/wrapTight 等 | ⚠️ 仅支持 wrapNone |
| 图片位置 | relativeFrom + posOffset | ⚠️ 结构有，API 不完整 |
| **图片裁剪** | `<a:srcRect>` | ❌ 缺失 |
| **图片旋转** | `<a:xfrm rot>` | ❌ 缺失 |
| **图片特效** | `<a:effectLst>` | ⚠️ 结构文件已有，但未完整连接 |

### 5. 形状与画布（Shape / Canvas / Group）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 行内形状 | wps:wsp inline | ✅ 已支持 |
| 浮动形状 | wps:wsp anchor | ✅ 已支持 |
| 画布 | wpc | ✅ 结构支持 |
| 分组 | wpg | ✅ 结构支持 |
| 形状文本框 | `<wps:txbx>` | ⚠️ 结构存在，API 有限 |

### 6. 页眉 / 页脚 / 页码

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| **默认页眉** | `word/header1.xml` + sectPr 引用 | ❌ 缺失 |
| **首页页眉** | `w:titlePg` + header 引用 | ❌ 缺失 |
| **奇偶页眉** | `w:evenAndOddHeaders` | ❌ 缺失 |
| **默认页脚** | `word/footer1.xml` + sectPr 引用 | ❌ 缺失 |
| **页码字段** | `<w:fldChar>` + `PAGE` / `NUMPAGES` | ❌ 缺失 |

> [!IMPORTANT]
> 页眉/页脚是业务文档中使用频率极高的功能，当前完全缺失，是 P0 优先级。

### 7. 样式与编号体系（Styles / Numbering）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| styles.xml 解析 | 全局样式定义 | ❌ 不解析（仅模板回写） |
| 样式引用 | `<w:pStyle>` / `<w:rStyle>` | ✅ 已支持 |
| **样式管理 API** | 列出/新增/修改样式 | ❌ 缺失 |
| numbering.xml 解析 | 编号定义 | ❌ 不解析（仅模板回写） |
| 编号引用 | `<w:numPr>` | ✅ 已支持 |
| **编号管理 API** | 新增/复用 abstractNum/num | ❌ 缺失 |

### 8. 文档结构与属性（Section / Document Properties）

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 节属性 | `<w:sectPr>` | ✅ 已支持 |
| 纸张大小 | `<w:pgSz>` | ✅ 已支持 |
| 页边距 | `<w:pgMar>` | ✅ 已支持 |
| 分栏 | `<w:cols>` | ✅ 已支持 |
| 文档网格 | `<w:docGrid>` | ✅ 已支持 |
| 多节支持 | 段落内 `<w:sectPr>` 实现分节 | ❌ 缺失 |
| **核心属性** | `docProps/core.xml` (标题/作者/日期) | ❌ 缺失 |
| **扩展属性** | `docProps/app.xml` | ❌ 缺失 |

### 9. 引用与批注体系

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| **脚注** | `<w:footnoteReference>` + footnotes.xml | ❌ 缺失 |
| **尾注** | `<w:endnoteReference>` + endnotes.xml | ❌ 缺失 |
| **批注/评论** | `<w:commentReference>` + comments.xml | ❌ 缺失 |
| **书签** | `<w:bookmarkStart>` / `<w:bookmarkEnd>` | ❌ 缺失 |
| **域（Field）** | `<w:fldChar>` / `<w:instrText>` (TOC, PAGE, DATE 等) | ❌ 缺失 |
| **目录（TOC）** | 基于域代码的自动目录 | ❌ 缺失 |

### 10. 内容控件 / 修订 / 高级

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| **SDT 内容控件** | `<w:sdt>` (下拉/文本/日期等) | ❌ 缺失 |
| **修订追踪** | `<w:ins>` / `<w:del>` / `<w:rPrChange>` | ❌ 缺失 |
| **图表** | `<c:chart>` (chart part) | ❌ 缺失 |
| **公式（Math）** | `<m:oMath>` | ❌ 缺失 |
| **OLE 对象** | `<o:OLEObject>` | ❌ 缺失 |

### 11. 文档操作

| 子功能 | 说明 | go-docx 状态 |
|:---|:---|:---|
| 解析 docx | zip → xml → struct | ✅ 已支持 |
| 生成 docx | struct → xml → zip | ✅ 已支持 |
| 文档拆分 | 按段落规则拆分 | ✅ 已支持 `SplitByParagraph()` |
| 文档合并 | 追加另一份文档内容 | ✅ 已支持 `AppendFile()` |
| 模板支持 | 使用已有 docx 作为模板 | ✅ 已支持 `UseTemplate()` |
| 媒体迁移 | 拆分/合并时自动迁移图片等 | ✅ 已支持 |
| 未知节点保留 | `RawXMLNode` round-trip | ✅ 已支持（Phase 0 已实现） |
| **跨 Run 文本替换** | 模板占位符替换 | ❌ 缺失 |

---

## 二、能力矩阵汇总

| 状态 | 数量 | 说明 |
|:---|:---:|:---|
| ✅ 已支持 | ~40 | 具备可用 API 或完整结构 |
| ⚠️ 部分支持 | ~8 | 结构/解析存在，但缺少创建或完整 API |
| ❌ 缺失 | ~20 | 完全未实现 |

### 按功能域覆盖率

```mermaid
graph LR
    A["文本/段落<br>90%"] --> B["超链接<br>80%"]
    B --> C["表格<br>80%"]
    C --> D["图片/绘图<br>70%"]
    D --> E["形状/画布<br>70%"]
    E --> F["页眉/页脚<br>0%"]
    F --> G["样式/编号<br>30%"]
    G --> H["文档结构<br>60%"]
    H --> I["引用/批注<br>0%"]
    I --> J["内容控件/高级<br>0%"]
    J --> K["文档操作<br>85%"]
```

---

## 三、优先级排序

`Priority Score = 频率(40%) + 价值(40%) + 落地可行性(20%)`，满分 100。

| 排名 | 能力域 | 当前状态 | 频率 | 价值 | 可行性 | Score |
|:---:|:---|:---|:---:|:---:|:---:|:---:|
| 1 | 模板文本替换（含跨 Run） | ❌ 缺失 | 5 | 5 | 4 | **96** |
| 2 | 页眉/页脚/页码 | ❌ 缺失 | 5 | 5 | 3 | **92** |
| 3 | 表格合并单元格 API（GridSpan/vMerge） | ⚠️ 解析有 | 5 | 4 | 5 | **92** |
| 4 | 段落缩进/间距 API 封装 | ⚠️ 结构有 | 5 | 3 | 5 | **84** |
| 5 | 样式管理 API | ❌ 缺失 | 4 | 5 | 3 | **84** |
| 6 | 编号管理 API | ❌ 缺失 | 4 | 4 | 3 | **76** |
| 7 | 书签与域（含 TOC） | ❌ 缺失 | 3 | 5 | 3 | **76** |
| 8 | 脚注/尾注 | ❌ 缺失 | 3 | 4 | 3 | **68** |
| 9 | 批注/评论 | ❌ 缺失 | 3 | 4 | 3 | **68** |
| 10 | 文档属性（core/app） | ❌ 缺失 | 3 | 3 | 5 | **64** |
| 11 | 图片增强（裁剪/旋转/环绕） | ❌/⚠️ | 3 | 3 | 3 | **60** |
| 12 | SDT 内容控件 | ❌ 缺失 | 3 | 4 | 2 | **60** |
| 13 | 多节支持 | ❌ 缺失 | 2 | 4 | 3 | **56** |
| 14 | 修订追踪 | ❌ 缺失 | 2 | 3 | 2 | **44** |
| 15 | 图表/公式 | ❌ 缺失 | 2 | 3 | 1 | **40** |

---

## 四、12 个月 Roadmap（4 阶段）

### Phase 0：高频 API 补全与稳定性（Q1 · 月 1-3）

**目标**：把已有结构但缺少 API 的高频功能补上，并强化 round-trip 保真。

| 交付物 | 优先级 | 说明 |
|:---|:---:|:---|
| `ReplaceText` / `ReplacePlaceholder` | P0 | 支持跨 run 文本查找替换，样式保持 |
| 表格合并 API（`SetColSpan` / `SetRowSpan`） | P0 | 基于已有 GridSpan/vMerge 结构封装 |
| 段落 `Indent()` / `LineSpacing()` API | P0 | 封装已有 `Ind` / `Spacing` 结构 |
| 上标/下标 `VertAlign()` API | P1 | 封装已有结构 |
| 单元格边框 API | P1 | 封装已有 `tcBorders` |
| 重复表头 `SetRepeatHeader()` | P1 | 新增 `<w:tblHeader>` |
| round-trip golden 测试扩充 | P0 | 覆盖更多真实文档样本 |

**验收标准**：
- 占位符替换在跨 run、含样式场景正确率 ≥ 99%
- 表格合并功能可通过 Word 打开验证
- 已有测试全部通过（无回归）

---

### Phase 1：页眉/页脚与样式体系（Q2 · 月 4-6）

**目标**：完成页眉/页脚/页码的完整读写，建立样式管理入口。

| 交付物 | 优先级 | 说明 |
|:---|:---:|:---|
| Header/Footer 读写 | P0 | 新增 header/footer XML part 解析与写入 |
| `GetHeader` / `SetHeader` API | P0 | 支持 default / first / evenOdd |
| `GetFooter` / `SetFooter` API | P0 | 同上 |
| 页码字段 `AddPageNumber()` | P0 | `<w:fldChar>` + `PAGE` 域代码 |
| styles.xml 解析 | P1 | 读取全局样式定义 |
| `ListStyles()` / `EnsureStyle()` | P1 | 样式管理 API |
| 文档属性读写 | P2 | core.xml / app.xml |

**验收标准**：
- 页眉/页脚在新建与解析文档均可稳定读写
- 页码在 Word 中正确显示
- 样式可列出、应用、新增基础样式

---

### Phase 2：编号与引用体系（Q3 · 月 7-9）

**目标**：建立编号管理 + 引用体系的可编程入口。

| 交付物 | 优先级 | 说明 |
|:---|:---:|:---|
| numbering.xml 解析 | P1 | abstractNum / num 定义解析 |
| `EnsureNumbering()` API | P1 | 编号定义创建与复用 |
| 书签 API 读写 | P1 | `<w:bookmarkStart>` / `<w:bookmarkEnd>` |
| 域代码基础框架 | P1 | `<w:fldChar>` / `<w:instrText>` |
| 脚注/尾注 API | P2 | footnotes.xml / endnotes.xml |
| 批注 API | P2 | comments.xml |
| 图片增强（环绕/旋转） | P2 | 扩展 Drawing API |

**验收标准**：
- 编号在 Word 中显示与预期一致
- 书签可创建、定位、通过内部超链接跳转
- 脚注/尾注/批注可创建、解析、回写

---

### Phase 3：高级对象与性能（Q4 · 月 10-12）

**目标**：具备高级对象策略能力，改善大文档可用性。

| 交付物 | 优先级 | 说明 |
|:---|:---:|:---|
| SDT 内容控件读写 | P2 | 结构化文档标签 |
| 多节支持 | P2 | 段落级 sectPr 实现分节 |
| 目录（TOC）支持 | P2 | 基于域代码的 TOC 结构 |
| 图表/公式"保留"策略 | P3 | round-trip 不破坏 |
| 修订追踪"保留"策略 | P3 | round-trip 不破坏 |
| 大文档性能优化 | P3 | 内存峰值/吞吐改善 |

**验收标准**：
- SDT 在主流样本上可稳定提取/填充
- 含图表/公式文档 round-trip 不损坏
- 基准样本内存/耗时相较基线显著改善

---

## 五、拟新增 API 草案

> [!NOTE]
> 以下为方向性设计草案，不代表最终签名。

### 文本替换

```go
// 跨 run 文本替换
func (f *Docx) ReplaceText(find, replace string, opts ...ReplaceOption) error

// 占位符批量替换（如 {{name}} → 张三）
func (f *Docx) ReplacePlaceholder(data map[string]string, opts ...ReplaceOption) error
```

### 段落增强

```go
func (p *Paragraph) Indent(left, right, firstLine, hanging int) *Paragraph
func (p *Paragraph) LineSpacing(before, after, line int) *Paragraph
func (r *Run) VertAlign(val string) *Run  // superscript / subscript
```

### 表格增强

```go
func (c *WTableCell) SetColSpan(span int) *WTableCell
func (c *WTableCell) SetRowSpan(val string) *WTableCell  // "restart" / "continue"
func (c *WTableCell) SetBorder(border CellBorder) *WTableCell
func (r *WTableRow) SetRepeatHeader(enable bool) *WTableRow
```

### 页眉/页脚

```go
func (f *Docx) GetHeader(kind HeaderKind) (*Header, error)
func (f *Docx) SetHeader(kind HeaderKind, header *Header) error
func (f *Docx) GetFooter(kind FooterKind) (*Footer, error)
func (f *Docx) SetFooter(kind FooterKind, footer *Footer) error
func (p *Paragraph) AddPageNumber(style PageNumberStyle) *Run
```

### 样式与编号

```go
func (f *Docx) ListStyles() []StyleInfo
func (f *Docx) EnsureStyle(style StyleDef) (styleID string, err error)
func (f *Docx) EnsureNumbering(def NumberingDef) (numID string, err error)
func (p *Paragraph) ApplyStyle(styleID string) *Paragraph
```

### 引用体系

```go
func (p *Paragraph) AddFootnote(text string) *Run
func (p *Paragraph) AddEndnote(text string) *Run
func (p *Paragraph) AddComment(author, text string) *Run
func (f *Docx) AddBookmark(name string, p *Paragraph) error
```

---

## 六、现有 ROADMAP.md 对比

仓库中已有 [ROADMAP.md](file:///Users/yuez/Workspace/Projects/event-insight-center/go-docx/ROADMAP.md)，本分析在其基础上做了以下补充和细化：

| 维度 | 已有 ROADMAP | 本分析补充 |
|:---|:---|:---|
| 功能全景 | 9 个能力域 | 11 个能力域、60+ 子功能逐项对标 |
| Phase 0 聚焦 | 保真回写 | 保真回写已部分落地（RawXMLNode），重心转向 API 补全 |
| 表格增强粒度 | ColSpan + 边框 + 重复表头 | 增加 RowSpan、嵌套表格确认、单元格级边框 API |
| 段落增强 | 未提及 | 增加 Indent/LineSpacing/VertAlign 等常用 API 封装 |
| 文档属性 | 未提及 | 增加 core.xml / app.xml 读写 |
| 图片增强 | 未提及 | 增加裁剪、旋转、多种文字环绕模式 |
| API 草案 | 4 个域 | 6 个域（新增引用体系、段落增强） |

---

## 七、风险与建议

> [!WARNING]
> - **OOXML 规范复杂**：不同 Word 版本（2007/2010/2013/2016/365）存在兼容差异，需要用多版本样本验证
> - **样式继承链**：样式名和编号定义依赖复杂的继承链（basedOn / link / numStyleLink），实现时需考虑完整解析
> - **域代码复杂性**：`<w:fldChar>` 是状态机结构（begin → instrText → separate → result → end），横跨多个 run，实现难度较大

> [!TIP]
> - **优先从"已有结构→封装 API"的路径入手**，投入少、见效快（如段落缩进、表格合并）
> - **页眉/页脚 是 ROI 最高的新增功能**，实现后可覆盖大量业务场景
> - **建议每个 Phase 结束前进行 round-trip 回归测试**，防止新功能引入结构破坏

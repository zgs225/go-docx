# go-docx 路线图（阶段制）

## 背景与目标

### 库定位
`go-docx` 是一个面向 Go 的 `.docx` 读写库，目标是在开源场景下提供可控、可扩展、可自动化的 Word 文档处理能力，覆盖“解析、编辑、生成、回写”全链路。

### 目标用户
- 需要批量生成合同/报告/通知等文档的后端服务团队
- 需要在保留原文档结构前提下进行自动化修改的文档平台团队
- 需要在 CI/服务端环境执行文档处理流水线的工程团队

### 总体目标
- 建立“高保真回写”能力：常见复杂文档在 parse/write 后不丢结构
- 完成高频业务能力闭环：模板替换、页眉页脚页码、表格高级操作
- 建立样式/编号/引用等文档体系化能力
- 为 SDT、图表/公式与大文档性能优化打下可迭代基础

## 当前能力盘点

### 已支持能力
- 文档解析与打包回写
- 段落与 run 级文本编辑（颜色、字号、下划线、高亮、字体等）
- 超链接编辑
- 图片（inline/anchor）插入与基础尺寸控制
- 表格创建与基础设置（宽度、对齐、底纹、padding）
- 形状、画布、分组相关结构的解析与部分生成能力
- 文档切分、段落/图形元素过滤等工具能力
- `ReplaceText/ReplacePlaceholder`（支持跨 run）
- `Hyperlink` 内文本替换与样式保真策略（MVP）
- `instrText` 可控替换（白名单 + 字段边界安全）
- 表格增强第一批：`ColSpan` / `RowSpan(restart|continue)` / `CellBorder` API
- 页眉/页脚/页码（单节 + 多 section 定位写入）
- first/even 组合开关（`titlePg` + `evenAndOddHeaders`）
- Header/Footer 共享对象 COW 隔离 + part 去重
- 表格增强后续项：`SetRepeatHeader` + 统一布局控制 API（表格级/行级）
- `rtcheck` 结构等价 CLI 与 CI job（GitHub Actions）

### 主要缺口
- 表格能力仍有缺口（自动跨行/跨列重排、`cantSplit` 等高级行属性）
- 缺少样式体系与编号体系的管理 API（style/numbering）
- 缺少脚注/尾注/批注/书签与域等引用体系能力
- 缺少 SDT（内容控件）读写与数据填充能力
- 缺少图表/公式等高级对象的可控策略（至少“保留 + 有限编辑”）

### 兼容性风险（重点）
- OOXML 细节复杂，不同 Word 版本可能存在兼容差异
- 真实业务样本覆盖不足时，测试通过不代表生产安全
- 回写保真与功能扩展存在结构冲突风险

## docx 功能全景与优先级排序

### 排序规则
`Priority Score = 频率(40%) + 价值(40%) + 落地可行性(20%)`

评分说明：
- 频率：业务中出现频率（1-5）
- 价值：对业务结果与工程成本影响（1-5）
- 落地可行性：在当前代码基线下的实现可达性（1-5）
- 最终分数：`(频率*0.4 + 价值*0.4 + 可行性*0.2) * 20`，满分 100

### 能力清单（按分数排序）

| 排名 | 能力域 | 状态 | 频率 | 价值 | 可行性 | Priority Score |
|---|---|---|---:|---:|---:|---:|
| 1 | 文档保真回写（未知节点保留 + round-trip） | 已完成 | 5 | 5 | 4 | 96 |
| 2 | 模板文本替换（含跨 run） | 已完成 | 5 | 5 | 4 | 96 |
| 3 | 页眉/页脚/页码 | 已完成 | 5 | 4 | 4 | 88 |
| 4 | 表格高级能力（合并/边框/布局） | 已完成（Phase 1 范围） | 4 | 5 | 4 | 88 |
| 5 | 样式与编号管理（styles/numbering） | 部分支持 | 4 | 5 | 3 | 84 |
| 6 | 引用体系（脚注/尾注/批注/书签/域） | 部分支持（仅字段代码替换） | 3 | 5 | 3 | 76 |
| 7 | SDT 内容控件 | 缺失 | 3 | 4 | 3 | 68 |
| 8 | 图表/公式策略（保留与有限编辑） | 缺失 | 2 | 4 | 2 | 56 |
| 9 | 大文档性能优化（内存/吞吐） | 缺失 | 3 | 3 | 2 | 56 |

### Top 优先级列表
1. 样式与编号管理（P1）
2. 引用体系补全（P2）
3. SDT 与高级对象策略（P3）
4. 图表/公式策略（P3）
5. 大文档性能优化（P3）

## 分阶段计划（阶段 + 验收标准）

### Phase 0：保真与稳定性底座

状态：`已完成`

已完成交付：
- [x] 未知节点保留机制（覆盖 body/paragraph/run/table/row/cell 关键层级）
- [x] round-trip 样本与结构等价校验基线
- [x] `rtcheck` CLI（可接入 CI）
- [x] `rtcheck` GitHub Actions job

验收结果：
- [x] 指定样本 parse + write 后关键结构不丢失
- [x] `go test ./...` 持续全绿
- [x] `rtcheck` 可作为 CI gate 使用

### Phase 1：高频业务能力

状态：`已完成`

已完成交付：
- [x] `ReplaceText/ReplacePlaceholder`（跨 run）
- [x] `WithMaxReplacements` / `WithCaseSensitive`
- [x] `Hyperlink` 文本替换
- [x] 样式保真与空 run 清理
- [x] `instrText` 替换开关（默认关闭）
- [x] `instrText` 白名单（`FORMTEXT` / `MERGEFIELD`）
- [x] 字段边界安全（`begin/separate/end`）
- [x] 表格增强第一批：`ColSpan` / `RowSpan` / `CellBorder`
- [x] 表格增强后续项：重复表头（`SetRepeatHeader`）与统一布局控制 API
- [x] Header/Footer 读写入口（单节 + 多 section）
- [x] 页码字段插入 API（含对齐与样式映射）
- [x] first/even 组合开关（`SetSectionTitlePage` / `SetEvenAndOddHeaders`）
- [x] Header/Footer COW 隔离与 part 去重
- [x] `main.go` 示例集成（Span/Border + Replace API）

阶段验收标准：
- 占位符替换在跨 run/混排场景稳定通过
- Header/Footer 在新建与解析后文档均可稳定读写
- 表格增强项具备单测 + 联动 round-trip 验证

### Phase 2：文档体系能力

状态：`未开始`

关键交付物：
- style 管理 API（读取/应用/新增基础样式）
- numbering 管理 API（abstractNum/num 定义与复用）
- 引用体系第一批：脚注、尾注、批注、书签/基础域

验收标准：
- 样式/编号在 Word 中显示与预期一致，且回写可复用
- 脚注/尾注/批注可创建、可解析、可回写
- 新增能力具有完整示例与最小迁移说明

### Phase 3：高级对象与性能

状态：`未开始`

关键交付物：
- SDT（内容控件）读写与基础填充
- 图表/公式的“保留 + 有限编辑”策略（优先不破坏）
- 大文档性能优化（内存峰值、处理吞吐、稳定性）

验收标准：
- SDT 在主流样本上可稳定提取/填充
- 含图表/公式文档 round-trip 不损坏关键结构
- 在基准样本上内存峰值/耗时相较基线显著改进

## 公共 API 草案（规划入口）

> 以下为规划草案，仅定义方向与入口，不代表最终签名。

### 已落地 API（节选）
- `ReplaceText(find, replace string, opts ...ReplaceOption) error`
- `ReplacePlaceholder(data map[string]string, opts ...ReplaceOption) error`
- `WithFieldCodeReplacement(enabled bool)`
- `WithFieldCodeWhitelist(types ...string)`
- `(*WTableCell).SetColSpan(cols int)`
- `(*WTableCell).SetRowSpanRestart()`
- `(*WTableCell).SetRowSpanContinue()`
- `(*WTableCell).ClearRowSpan()`
- `(*WTableCell).SetCellBorderTop/Right/Bottom/Left(...)`
- `(*WTableCell).SetCellBordersSame(...)`
- `(*WTableCell).ClearCellBorders()`
- `SetHeader(kind HeaderKind, header *Header) error`
- `SetFooter(kind FooterKind, footer *Footer) error`
- `AddPageNumber(style PageNumberStyle, kind ...FooterKind) error`
- `AddPageNumberAligned(style PageNumberStyle, align string, kind ...FooterKind) error`
- `SectionCount() int`
- `SetSectionHeaderText(section int, kind HeaderKind, text string) error`
- `SetSectionFooterText(section int, kind FooterKind, text string) error`
- `AddSectionPageNumberAligned(section int, style PageNumberStyle, align string, kind ...FooterKind) error`
- `SetSectionTitlePage(section int, enabled bool) error`
- `SetEvenAndOddHeaders(enabled bool) error`
- `SetRepeatHeader(row int, enable bool) *Table`
- `SetLayout(opts TableLayoutOptions) *Table`
- `SetRowLayout(row int, opts RowLayoutOptions) *Table`

### 待落地 API（规划）
- `ListStyles() []StyleInfo`
- `EnsureStyle(style StyleDef) (styleID string, err error)`
- `EnsureNumbering(def NumberingDef) (numID string, err error)`

## 验收标准（全局）

- 文档完整性：覆盖“功能全景、排序、阶段计划、验收标准、API 草案”
- 一致性：高优先级能力优先进入当前阶段执行队列
- 可执行性：每阶段包含“目标、关键交付物、验收标准”
- 变更范围：路线图文档可独立演进，不与功能提交强绑定

## 风险与依赖

### 关键风险
- OOXML 细节复杂，且不同 Word 版本存在兼容差异
- 缺少真实业务样本时，测试通过不代表生产安全
- 回写保真与功能扩展可能存在结构冲突

### 外部依赖
- 稳定的真实 docx 样本库（目录、域、页眉页脚、引用、控件、图表）
- CI 中可重复执行的 round-trip 对比工具链
- 明确的向后兼容策略（旧 API、旧文档行为）

### 向后兼容策略
- 新 API 以增量方式引入，避免破坏现有入口
- 对已支持功能保持默认行为不变，新增能力通过选项启用
- 关键行为变更提供迁移说明与版本标记

## 附录

### 术语
- Round-trip：文档 `parse -> modify(可选) -> write` 后结构与关键语义保持一致
- 保真：回写后不丢关键节点、不破坏可编辑性与显示语义
- SDT：Structured Document Tag，Word 内容控件

### 优先级评分规则
- Score = `(频率*0.4 + 价值*0.4 + 可行性*0.2) * 20`
- 分数区间解释：
  - 90-100：必须立即推进（P0）
  - 80-89：高优先级（P1）
  - 65-79：中优先级（P2）
  - <65：后续规划（P3）

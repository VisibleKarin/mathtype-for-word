# MathType for Word

## English

The [original project](https://github.com/felimet/mathtype-for-word) lets AI agents use a local tool interface (MCP) to edit MathType equations in Word and PowerPoint quietly, without taking over the mouse or keyboard. This personal fork focuses on Word: replacing selected Word equations (OMML) with editable MathType objects and applying the document preferences below.

### What this fork changes

1. **Native numbering and references through MCP.** The normal render, validation and field-update tools now support continuous, chapter, section, and chapter–section numbering, such as `(1)`, `(2-1)` and `(2.3.1)`. MathType's own equation breaks, numbers and references work together, avoiding a separate hand-built list-numbering scheme. MathType still uses Word fields internally. Equation chapters/sections are independent of Word sections and heading styles; their boundaries must be specified explicitly.
2. **Size based on actual body text.** Unless asked to preserve the size, match MathType's internal Full size (main symbol size) to nearby body text, including directly applied character formatting rather than only the Normal style. Ambiguous context requires an explicit size. Save and reopen to verify the actual internal value. Keep MathType's fraction layout; do not add extra shrinking to imitate OMML's numerator/denominator appearance.
3. **Local edits in an existing document.** Inspect the source equation, review its mathematical content, mark only the selected OMML objects, and render editable MathType replacements in a working copy. Existing MathType objects can also be replaced locally together with their preview, dimensions and baseline. Preserve surrounding text, breaks, comments and native references wherever the selected range permits. A separate output file protects the original; the document does not need to be rebuilt from scratch.
4. **Simple inline mathematics stays editable text.** Prefer ordinary Word characters with true subscript/superscript formatting for short expressions within a sentence. Preserve mathematical roles and punctuation, and check the character baseline and line height. Complex inline formulas may still need MathType. Document grid, paragraph grid snapping and character grid snapping are separate settings: inspect them together instead of globally disabling the grid or applying one baseline offset to every formula.
5. **Displayed equations use MathType and tab stops.** A center tab places the equation and a right tab places its number, using the actual available line width. When an explanation logically continues the same paragraph, ask whether to use a soft break (Shift+Enter) and leave the continuation flush left. Keep an existing mixed paragraph intact; do not automatically split or merge paragraphs. New display markers require a separate paragraph; object-only placement is used when preserving an existing mixed layout.
6. **Check content after formatting.** Compare mathematical structure to catch accidental duplication or changes to script/fraction structure. This is a preservation check, not a proof that a derivation is correct. Formula-writing guidance is planned for the future.
7. **Check the formatting route before first use.** For internal-size work, check the installed SDK, worker architecture, exported functions and actual SDK connection. Start with the verified native SDK route, avoiding the auxiliary VBA/WLL formatting route that failed on the tested machine. Report the failing stage before editing a document. This does not establish the root cause of the old VBA initialization failure.

Illustrated explanations are planned for the future. Detailed inputs and limits are in the [workflow](skills/mathtype-for-word/references/workflow-and-schema.md) and [local-edit guide](skills/mathtype-for-word/references/local-equation-edits.md).

### Credits and caution

Based on **Jia-Ming Zhou (Felimet)**'s [mathtype-for-word](https://github.com/felimet/mathtype-for-word). The original [MIT license](LICENSE) and copyright notices are retained; see the upstream README for its full introduction.

These changes were implemented **with Codex assistance** and remain experimental. **Use cautiously**, test document copies, and review both mathematical content and layout. Local tests use Windows desktop Word **16.0** and MathType **7.8.2.441** with its x86 SDK. Other installations need their own prerequisite checks. The local interface depends on installed desktop Word and MathType.

---

## 中文

[原项目](https://github.com/felimet/mathtype-for-word)的重要特色是让 AI 通过本地工具接口（MCP）静默编辑 Word 和 PowerPoint 中的 MathType 公式，无需接管桌面键鼠。本 fork 主要面向 Word，用于将选定的 Word 自带公式（OMML）替换为可编辑的 MathType 公式，并按以下个人偏好处理格式。

### 改动内容

1. **通过 MCP 使用 MathType 原生编号和引用。** 常规渲染、验证和字段更新接口支持连续编号、按章、按节及按章和节编号，例如 `(1)`、`(2-1)`、`(2.3.1)`，最后一种表示第 2 章第 3 节的第 1 个公式。公式章、节标记与编号、引用由 MathType 配套管理，便于多章节文档统一处理，无需另拼一套手动多级列表。MathType 底层仍使用 Word 字段；它的公式章、节独立于 Word 分节符和标题样式，需要明确指定边界。
2. **依据正文实际字号匹配公式。** 除非要求保留字号，优先读取邻近正文字符的实际字号，包括直接设置的格式，而不只读取“正文／Normal”样式；上下文不明确时需指定字号。将此值写入 MathType 内部 Full（主字号），保存后重新读回核验。分式保留 MathType 自身的布局，不为模仿 OMML 分子、分母的显示效果而额外缩小字号。
3. **在现有文档中局部修改。** 先核对原公式内容，再精确标记选定的 OMML 对象，在工作副本中生成 MathType 公式。已有 MathType 对象也可局部替换，并同步预览、尺寸和基线。在目标范围允许的情况下，保留相邻文字、软硬回车、批注及原生编号引用。另存输出用于保护原稿，不必从头重建整篇文档。
4. **简单行内数学内容优先使用 Word 文字。** 行内（inline）指与正文处于同一句中的变量或表达式。简单内容优先使用普通字符及 Word 上下标，保留数学角色和标点，并检查字符基线、行高；复杂行内公式仍可使用 MathType。文档网格、段落对齐网格和字符对齐网格是不同设置，应结合检查，不全局关闭网格，也不给所有公式套用同一个位置偏移。
5. **独立显示的公式使用 MathType 与制表位。** 显示（display）公式单独占一行，使用居中制表位放公式、右对齐制表位放编号，位置按实际可用行宽计算。若公式后的说明仍属于同一自然段，提醒用户确认是否使用软回车（Shift+Enter），并让续行顶格。已有混排段落优先保留，不自动拆段或合段。新建显示公式的标记须独立成段；保留原有混排时只替换公式对象。
6. **核对数学结构。** 检查格式调整是否造成内容重复、上下标或分式层级变化；这项检查不能证明推导正确。适当提示用户公式撰写规范（功能待未来加入）。
7. **补充首次使用检查。** 处理内部字号前，检查本机 SDK、辅助进程位数、所需接口和实际连接。直接采用已验证的原生 SDK 路线，避免先走本机曾失败的旧 VBA／WLL 格式化链路，并在编辑文档前报告具体失败阶段。这不等于已查明旧宏初始化失败的根因。

（图片说明待未来加入）

具体参数与限制见[工作流程](skills/mathtype-for-word/references/workflow-and-schema.md)及[局部编辑说明](skills/mathtype-for-word/references/local-equation-edits.md)。

### 原项目与致谢

基于 [mathtype-for-word](https://github.com/felimet/mathtype-for-word)，原作者为 **Jia-Ming Zhou（Felimet）**。保留原项目的 [MIT 许可证](LICENSE)及版权声明。原项目的完整介绍请见其 README。

### 使用提示

本 fork 的改进在 **Codex 的帮助下实现**，仍属个人实验性版本，**请谨慎使用**。建议先在文档副本上测试，并人工检查公式内容和排版。本地实测环境为 Windows 桌面版 **Word 16.0、MathType 7.8.2.441（x86 SDK）**；其他环境需先检查。使用时仍需本机安装桌面版 Word 和 MathType。

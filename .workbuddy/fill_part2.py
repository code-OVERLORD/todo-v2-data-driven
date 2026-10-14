# -*- coding: utf-8 -*-
"""根据 W2-W5 AI 协作日志，填充实验报告 Part2 的 (2.1)-(2.4) 各填空题。"""
import copy
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
import docx

DOC = r"C:/Users/17795/Desktop/web/202601-Web-课程实验1-学号-姓名.docx"

NOTE_SHOT = "【此处插入浏览器运行效果截图：建议缩放至 60%~70% 后截取并居中放置，避免大体积空图】"

# ---------- 2.1 HTML 个人简历 ----------
P19 = [
    ("我给 AI 设定的角色是“资深前端开发与 UI 设计师”，要求输出一份可直接运行的 HTML 个人简历文件。Prompt 由三组约束组成：", "p"),
    ("（1）内容结构要求", "sub"),
    ("① 个人信息：姓名、职位标签、联系方式（邮箱、电话、社交链接）；", "p"),
    ("② 个人简介：一段简短介绍，用占位符 [这里写你的简介]；", "p"),
    ("③ 核心技能：以标签云或列表展示，占位符 [技能1, 技能2, ...]；", "p"),
    ("④ 项目/工作经历：标题、时间、公司/组织、2–3 点职责描述；", "p"),
    ("⑤ 教育背景：学校、专业、时间。", "p"),
    ("（2）视觉与交互规范", "sub"),
    ("设计风格现代、专业、干净，主色调建议蓝灰/靛蓝；必须适配移动端和桌面端（Flexbox 或 Grid）；添加微妙的悬停效果或平滑过渡。", "p"),
    ("（3）代码质量标准（强制）", "sub"),
    ("包含 <!DOCTYPE html> 声明；所有属性值用双引号包裹；使用 header/main/section/aside 等语义化标签；内联 CSS 并集中管理、带注释；文字与背景对比度符合 WCAG AA。", "p"),
    ("我的设计意图：这一轮我只锁定“结构与规范”，不锁定“具体样式”，把样式留给 2.2 单独验证——每次只验证一件事，才能让下一轮的对比有清晰基准。", "p"),
]

P21 = [(NOTE_SHOT, "note")]

P23 = [
    ("按 L1–L4 四层逐层审查：", "p"),
    ("L1 结构完整性：通过。HTML 骨架规范，<!DOCTYPE html>、html/head/body 齐备，meta charset 与 viewport 均有；个人信息、简介、技能、项目/经历、教育背景五个模块全部存在，无缺失元素。", "p"),
    ("L2 语法正确性：通过。标签闭合完整、嵌套层级正确，属性值均用双引号包裹，未发现拼写错误或非法属性。", "p"),
    ("L3 语义合理性：发现可改进项。正文用 Emoji 当图标（📧、📱 等）却没有无障碍标注，屏幕阅读器会把 Emoji 逐个读出来，打断正文朗读。应补 aria-hidden=\"true\" 或 role=\"presentation\"，让装饰性符号对辅助技术不可见；真正需要可读的名称另用 <span class=\"sr-only\"> 提供。", "p"),
    ("L4 工程质量：发现可改进项。颜色值全部硬编码（#2c3e50、#3498db 等分散出现多次），未提取为 --primary-color 等 CSS 变量，换主题要逐处手改；样式虽已内联但缺少分区注释层次，二次修改时定位成本高。", "p"),
    ("我的判断与决策：接受。", "sub"),
    ("理由：L1、L2 是“能不能用”的底线，均已通过；L3、L4 两项属于“能用但不够好”，且都是局部、可增量修复的问题，不影响结构。因此我先接受这一版作为骨架，把 Emoji 无障碍标注记入待办，把颜色变量化推迟到 2.2（CSS 阶段）统一处理——这一轮的职责是定结构，若同时改样式，下一轮的对比就失去了基准。", "p"),
    ("支持我判断的证据（无障碍树的反例）：在 DevTools 的 Accessibility 面板中查看含 Emoji 的 <span>，它的 computed name 直接取自文本内容，朗读时会在句子中间插入“信封”“手机”之类的符号名；给该 span 加上 aria-hidden=\"true\" 后，节点从无障碍树中被移除，朗读恢复连贯。这说明“页面看起来对”不等于“对辅助技术也对”，语义问题必须靠工具验证，不能靠肉眼。", "p"),
]

P25 = [
    ("AI 生成 HTML 时，“结构完整、语法正确”几乎从不出错，真正的差距在语义与工程细节：装饰性元素是否对辅助技术隐藏、重复取值是否抽成变量。", "p"),
    ("我学到的判断顺序是“先验证能不能用（L1/L2），再评估好不好改（L3/L4）”，并且一次只验证一件事，避免把结构问题和样式问题混在一起改，导致改完说不清是哪个改动起了作用。", "p"),
]

# ---------- 2.2 CSS 样式 ----------
P30 = [
    ("请帮我设计一个个人简历页面的 CSS 样式。页面包含：顶部个人信息区（姓名、职位、头像）、联系方式、个人简介、工作经历、教育背景、技能标签、项目经历。要求：", "p"),
    ("· 整体风格简洁、专业，适合互联网岗位；", "p"),
    ("· 使用 Flexbox 或 Grid 布局，结构清晰；", "p"),
    ("· 颜色不超过 3 种主色，有层次感；", "p"),
    ("· 字体使用系统字体栈，保证兼容性；", "p"),
    ("· 适配桌面端和移动端，移动端下工作经历、教育背景等模块改为单列；", "p"),
    ("· 代码要有注释，类名语义化；", "p"),
    ("· 不要使用过时的 float 布局；", "p"),
    ("· 输出完整 CSS，并说明关键选择器作用。", "p"),
    ("我特意在 Prompt 里加了两条“反向约束”（不用 float、主色不超过 3 种），目的是让 AI 的选择空间变小、让我更容易判断它是否真的遵守，而不是事后才发现它偷偷用了老写法。", "p"),
]

P32 = [(NOTE_SHOT, "note")]

P34 = [
    ("按 L1–L4 加“四查清单”逐层审查，结果如下：", "p"),
    ("L1 结构完整性（基本通过，1 项风险）：header / contact / main / experience / education / skills 等主要区域都有样式；有全局重置 * { margin:0; padding:0; box-sizing:border-box }；.resume 作外层容器、.main 用 Grid 分栏；有 @media (max-width: 768px)。风险：Prompt 明确要求“项目经历”，但 CSS 中完全没有 .projects / .project-item 规则。", "p"),
    ("L2 语法正确性（通过）：类选择器与后代选择器书写正确，属性名/值未见拼写错误，括号闭合正常，px/rem/% 使用合理，系统字体栈与现代布局兼容性良好。", "p"),
    ("L3 语义合理性（1 项风险、2 项可改进）：类名语义清晰（.header、.section-title、.skill-tag）；主体 Grid + 顶部/标签 Flex 的搭配合理；主色 #2c3e50、强调色 #3498db、背景 #f5f7fa 克制。可改进：只有 768px 一个断点，平板到手机的过渡偏粗；缺少 @media print。风险：未考虑 :focus 焦点态与对比度。", "p"),
    ("L4 工程质量（2 项风险、1 项可改进）：可改进：颜色值分散、未用 CSS 变量。风险：移动端只改了 .header 和 .main，.contact 的换行与间距没处理；缺少打印样式——简历是要打印/导出 PDF 的，深色 header 会浪费墨水，甚至把白字印在白底上。", "p"),
    ("四查清单：查结构 → 基本完整但项目经历样式缺失；查语法 → 通过；查语义 → 类名、布局、颜色合理，但可访问性与打印不足；查工程 → 颜色未变量化、断点偏少、打印缺失。", "p"),
    ("我的判断与决策：修改（手动修改，不重新提问）。", "sub"),
    ("理由：问题共 5 条，但都落在“补充与替换”层面（补一个模块样式、补打印规则、把硬编码换成变量、加一个断点、加焦点态），核心布局思路（Grid 分栏 + Flex 标签）是对的。重开一轮 Prompt 不会改变这些判断，只会浪费时间，所以我按“问题—位置—性质—原因—改法”逐条登记后手动改：", "p"),
    ("1. 缺少项目经历样式 / 未定义 .projects、.project-item / 风险 / Prompt 要求了但 CSS 没覆盖 / 增加 .projects、.project-item、.project-item h3、.project-item .meta 等样式；", "p"),
    ("2. 颜色值分散 / .header、.section-title、.skill-tag 等 / 可改进 / 后期换肤要改多处 / 在 :root 中定义 --color-primary、--color-accent、--color-bg 等变量；", "p"),
    ("3. 缺少打印样式 / 全局 / 风险 / 简历常需打印或导出 PDF，深色 header 浪费墨水 / 增加 @media print，去掉阴影与背景色，改为黑白打印友好；", "p"),
    ("4. 响应式断点较少 / @media (max-width: 768px) / 可改进 / 平板与手机之间可能过挤或过空 / 增加 @media (max-width: 480px)，调整字号、头像尺寸与 padding；", "p"),
    ("5. 可访问性不足 / 全局 / 可改进 / 键盘导航、对比度、焦点状态未考虑 / 增加 :focus-visible 样式，并复核文字对比度。", "p"),
]

P36 = [
    ("证据一（反例 + 运行现象）：验证“项目经历样式缺失”。", "sub"),
    ("我在 HTML 中补上 <section class=\"projects\"><div class=\"project-item\">…</div></section>，浏览器打开后该区块的标题字号、条目间距与“工作经历”明显不一致；用 DevTools 选中 .project-item 查看 Styles 面板，命中它的声明为空（除继承外没有任何规则），computed margin-bottom 为 0，而 .experience-item 是 20px。这说明“主要模块都有样式”只是看起来齐全——Prompt 里点名的模块也可能被漏掉，审查必须逐条对着 Prompt 对账，不能凭页面观感下结论。", "p"),
    ("证据二（代码行为分析）：验证“颜色硬编码导致一改就乱”。", "sub"),
    ("改前我在 DevTools 里把 .header 的 background 由 #2c3e50 改为 #1a252f，结果 .section-title 的 color 仍是 #2c3e50、.skill-tag 的 #3498db 也纹丝不动，整页出现“换了主色但标题没换”的割裂；统计发现 #2c3e50 出现 5 次、#3498db 出现 4 次，换一次主题要手改 9 处。改为 :root 变量后，只需改 --color-primary 一行，三处同步生效。由此我得到一个可复用的判据：同一语义的取值在文件里出现两次以上，就应该抽成变量。", "p"),
    ("证据三（规则依据 + 边界场景）：验证“缺少打印样式”。", "sub"),
    ("按 Ctrl+P 打开打印预览，深色 .header 整块被输出，而在“不打印背景色”的默认设置下白色文字落在白底上几乎不可见；.resume 的 box-shadow 也被打成灰色块。依据是：未定义 @media print 时打印完全回退到屏幕样式，浏览器不会自动为深色底反色。这不是审美问题，而是简历这种“必然要被打印”的文档的功能缺陷。", "p"),
]

P38 = [
    ("AI 生成的 CSS 最容易犯的不是语法错误，而是“场景覆盖不全”——只写主流程样式，漏掉边界情形（缺失的模块、打印、窄屏断点、可访问性），再加上颜色/尺寸硬编码分散，结果是“看着能用、一改就乱”。", "p"),
    ("我学到的做法是：审查时拿 Prompt 逐条对账，而不是看页面顺不顺眼；凡是同一语义的值重复出现两次以上，一律抽成变量；凡是“这个文档将来会被怎么用”（打印、窄屏、键盘操作），都要当成功能项而不是加分项。", "p"),
]

# ---------- 2.3 Flex 布局 ----------
P41 = [
    ("请用 Flex 布局帮我设计一个个人简历页面的 CSS。页面包含：顶部个人信息区（姓名、职位、头像）、联系方式、个人简介、工作经历、教育背景、技能标签、项目经历。要求：", "p"),
    ("· 整体用 Flex 布局实现，不要用 float，也尽量少用 Grid；", "p"),
    ("· 顶部姓名/职位和头像横向排列，移动端改为纵向；", "p"),
    ("· 联系方式一行排列，空间不够时自动换行；", "p"),
    ("· 主体分左右两栏：左侧放工作经历/项目经历/教育背景，右侧放个人简介/技能标签，移动端改为上下排列；", "p"),
    ("· 技能标签横向排列、自动换行；", "p"),
    ("· 每条经历的标题和时间在同一行，时间靠右；", "p"),
    ("· 颜色不超过 3 种主色，字体用系统字体栈；", "p"),
    ("· 类名语义化，代码有注释，输出完整 CSS。", "p"),
    ("这一轮我刻意把布局技术限定为 Flex（禁止 float、尽量少用 Grid），目的是逼出“一维布局”的典型错误——该换行却不换、该伸缩却写死宽度，从而与 2.4 的 Grid 形成对照。", "p"),
]

P43 = [(NOTE_SHOT, "note")]

P45 = [
    ("L1 结构完整性：发现 2 个问题。缺少全局 box-sizing: border-box；.resume 写死 width: 960px，小屏下会溢出。", "p"),
    ("L2 语法正确性：通过。选择器、属性、单位书写正确，未见拼写错误，括号闭合正常。", "p"),
    ("L3 规则是否互相打架：发现 3 个问题。.contact span 用 margin-right 代替 gap，最后一项右侧多出空白；.left-column 写死 width: 620px，与 .main 的 padding 叠加后桌面端可能溢出；.skills 未设 flex-wrap，标签多时会挤成一行溢出。", "p"),
    ("L4 工程质量：发现 2 个问题。* { margin:0; padding:0 } 作用域过宽，有副作用；颜色值硬编码分散，未使用 CSS 变量，后续换肤困难。", "p"),
    ("四查清单：一查尺寸模型 → .resume 同时有 width 与 padding 却没有 border-box，总宽会超 960px；二查尺寸适配 → .resume 写死 960px、.left-column 写死 620px，与容器 padding 冲突；三查选择器 → 通配符作用域过宽会误伤表单元素，.contact span 依赖标签选择器、语义偏弱；四查可访问性 → .item-header .meta 用 #888 在白底上小字号偏弱，且未设 :focus-visible。", "p"),
    ("Flex 专项检查：", "sub"),
    ("· flex:1 与 width 并存 → .left-column 写死 620px，.right-column 用 flex:1（等价于 flex:1 1 0，flex-basis:0 会覆盖 width 作为初始主轴尺寸），右栏取剩余空间、左栏仍为 620px，在 1024px 窄桌面下两栏加间距会溢出；", "p"),
    ("· flex-wrap → .contact 与 .skills 都没设，默认 nowrap，子项多了就溢出；", "p"),
    ("· gap → 两处都用 margin-right 代替，最后一个子项右侧多出空白；", "p"),
    ("· 主轴方向 → 通过，桌面 row、移动端 column，与设计意图一致；", "p"),
    ("· 对齐方式 → .header 只设 align-items: center，没设 justify-content，“头像+文字”整体没有水平居中。", "p"),
    ("我的判断与决策：修改。", "sub"),
    ("理由：问题分布在尺寸模型、换行策略、间距写法、选择器作用域、可访问性五类，但都属于“改写法”而非“改方案”——Flex 的整体思路（两栏 + 标签换行 + 标题与时间同行）是正确的。因此我手动修改并逐一验证：全局加 box-sizing: border-box；.resume 的写死宽度改为 width: 100% + max-width: 960px；.left-column / .right-column 改为 flex: 2 1 480px / flex: 1 1 240px；.contact 与 .skills 加 flex-wrap: wrap 并用 gap 替换 margin-right；通配符重置收窄为 * { box-sizing: border-box } + 精确元素重置；补 a:focus-visible / button:focus-visible。", "p"),
]

P47 = [
    ("证据一（完整验证，含修改前/后对比）：.resume 写死 width: 960px 且缺 box-sizing: border-box，窄屏溢出。", "sub"),
    ("代码位置：.resume { width: 960px; margin: 0 auto; … }，以及全局 * { margin:0; padding:0 } 缺 box-sizing。", "p"),
    ("现象/风险：窗口宽度 < 960px 时 .resume 横向溢出、出现横向滚动条；即使窗口 ≥ 960px，一旦给 .resume 加 padding 或 border，总宽也会超过 960px。", "p"),
    ("技术依据：CSS 盒模型默认 box-sizing: content-box，width 只表示内容区宽度，实际占位 = width + padding + border；要让 width 表示“含内边距和边框的总宽”，必须显式设置 box-sizing: border-box。", "p"),
    ("修改方案：全局加 * { box-sizing: border-box }；.resume 的 width: 960px 改为 width: 100%; max-width: 960px; margin: 0 auto。", "p"),
    ("验证方式：浏览器打开页面，把窗口宽度从 1200px 逐步拖到 375px，观察是否出现横向滚动条；DevTools 查看 .resume 的 computed width。", "p"),
    ("验证结果：修改前，缩到 800px 时出现横向滚动条，.resume 的 computed width 仍为 960px；修改后，缩到 375px 时无横向滚动条，computed width 自动收缩为视口宽减去两侧 padding 后的值。", "p"),
    ("结论：原判断“写死宽度 + 缺 box-sizing 会导致窄屏溢出”得到验证。", "p"),
    ("证据二（边界输入）：.contact 与 .skills 未设 flex-wrap。", "sub"),
    ("我把联系方式项增加到 8 个、技能标签增加到 12 个，两者都挤在一行并溢出容器（DevTools 显示 scrollWidth > clientWidth）。加上 flex-wrap: wrap、并用 gap: 16px / gap: 8px 替换 margin-right 后，窗口缩到 375px 时联系方式自动换成两行、标签自动换行，且每行末尾不再有多余空白。验证有效。", "p"),
    ("证据三（副作用反例）：* { margin:0; padding:0 } 作用域过宽。", "sub"),
    ("通配符会重置所有元素的 margin/padding，包括 input、button、ul、ol。当前页面没有表单所以看不出问题，但我临时插入一个 <ul> 列表后发现它的 padding-left 被清零、项目符号缩进消失，需要额外补救。改为 * { box-sizing: border-box } 加 body, h1, h2, h3, p, ul, ol { margin:0; padding:0 } 后，新增列表保留浏览器默认缩进，语义更正确。验证有效。", "p"),
]

P49 = [
    ("AI 用 Flex 写布局最典型的错误是“把 Flex 当 Grid 用”：该换行的地方不写 flex-wrap，该用 gap 的地方用 margin-right，该伸缩的地方写死 width，再叠加缺 box-sizing: border-box，结果桌面端勉强能用、一缩窗口就溢出。", "p"),
    ("我学到的正确做法是——容器负责 flex-wrap 与 gap，子项用 flex: <grow> <shrink> <basis> 而不是写死 width，全局统一 box-sizing: border-box；判断一个 Flex 容器是否写对，就问三件事：要不要换行、间距用什么控制、子项能不能伸缩。", "p"),
]

# ---------- 2.4 Grid + 响应式 ----------
P52 = [
    ("请用 HTML + CSS 帮我做一个个人简历页面，布局要求：", "p"),
    ("· 顶部是头像 + 姓名 + 一句话简介的 header 区；", "p"),
    ("· 主体分左右两栏：左侧是联系方式/技能/证书侧边栏，右侧是工作经历/项目经历/教育背景主内容区；", "p"),
    ("· 右侧的工作经历、项目经历等用卡片网格呈现，一行 2 列，小屏 1 列；", "p"),
    ("· 用 CSS Grid 实现，配合媒体查询做到响应式，手机、平板、桌面都能正常显示；", "p"),
    ("· 不要用框架，纯 CSS；代码整洁、可维护，注意 box-sizing、可访问性。", "p"),
    ("这一轮我特意没有替 AI 决定断点数值，只说“手机、平板、桌面都能正常显示”，想看它会选什么断点、是从移动端向上增强还是从桌面端向下覆盖——这正好是“断点由谁决定”的考点。", "p"),
]

P54 = [(NOTE_SHOT, "note")]

P56 = [
    ("L1 结构完整性：通过。CSS 可正常引入，选择器与规则闭合完整，无缺花括号、无漏分号。", "p"),
    ("L2 语法正确性：通过。选择器写法、属性值、单位（px/%/fr）均合法，媒体查询语法正确。", "p"),
    ("L3 规则互相打架：发现冲突。.resume 用 grid-template-columns: 250px 1fr 定义了侧栏轨道，.sidebar 却又写死 width: 250px，项目宽度与轨道重复约束；.main 用 50% 50% 再加 gap: 20px 导致溢出。", "p"),
    ("L4 工程质量：发现多处问题。未设 box-sizing: border-box；.resume 写死 width: 1000px；正文 #999、技能列表 #bbb 对比度不足；全局缺 line-height；未做移动优先。", "p"),
    ("四查清单：一查尺寸模型 → .header（height+padding）、.card（width+padding+border）、.main（grid+gap+padding）都同时出现尺寸与内边距/边框却没有 border-box，实际尺寸被撑大；二查尺寸适配 → .resume 写死 1000px，视口 <1000px 必然横向滚动，靠断点才改成 100%，属于“桌面写死后靠断点救”；三查选择器 → .skills li 后代选择器作用域偏宽，未滥用 ID、没有 !important（这点通过）；四查可访问性 → #999 与 #bbb 在白底上对比度约 2.8:1 / 1.9:1，远低于 WCAG AA 的 4.5:1，且全局无 line-height。", "p"),
    ("Grid / 响应式专项审查：", "sub"),
    ("· 为什么用 Grid 而不是 Flex → 合理。卡片网格是二维布局，Grid 更直接，二维卡片网格 + 主从栏布局用 Grid 恰当。", "p"),
    ("· 为什么选这些断点 → 待改进。只用了 768px 一个“经典 iPad 竖屏值”，并非由内容决定，也缺少侧栏与主区堆叠的中间断点。", "p"),
    ("· 是否移动优先 → 反了。默认是 1000px 双栏桌面样式，靠 max-width: 768px 向下覆盖，属于桌面优先而非移动优先。", "p"),
    ("· fr 单位是否合理 → 不合理。.main 用 50% 50% + gap: 20px，总宽 = 100% + 20px，必然横向溢出。", "p"),
    ("· 项目是否设置 width → 不合理。.sidebar 的 width: 250px 与轨道冲突；.card 的 width: 100% 叠加 padding+border，在无 border-box 时溢出。", "p"),
    ("我的判断与决策：修改（其中响应式策略部分按“重写”处理）。", "sub"),
    ("理由：这一轮的问题不再是零散的写法问题，而是响应式策略本身选错了方向——先写死桌面宽度、再用 max-width 向下打补丁。这种结构无法通过局部微调修复，必须把默认样式改为移动优先、断点改为 min-width 向上增强。因此我保留它的 Grid 分栏思路，重写了尺寸与响应式部分：", "p"),
    ("· *, *::before, *::after { box-sizing: border-box }；", "p"),
    ("· .resume 改 width: min(1000px, 100% - 2rem); margin-inline: auto，默认 grid-template-columns: 1fr（小屏单列）；", "p"),
    ("· .main 改 grid-template-columns: repeat(2, minmax(0, 1fr))，用 fr 取代 50% + gap；", "p"),
    ("· 断点改为 min-width: 640px（卡片双列）与 min-width: 900px（侧栏 + 主区）两级向上增强；", "p"),
    ("· 正文色改 #333、次要文字 #555，body 加 line-height: 1.6，并移除 .sidebar 上多余的 width。", "p"),
]

P58 = [
    ("证据一（完整验证，含修改前/后对比）：.main 的 grid-template-columns: 50% 50% 加 gap: 20px 导致溢出。", "sub"),
    ("代码位置：.main { grid-template-columns: 50% 50%; gap: 20px; }", "p"),
    ("现象/风险：两列各占容器 50%，合计 100%，再加 20px 间隙，总宽超出容器 20px，出现横向滚动条或卡片被裁切。", "p"),
    ("技术依据：grid-template-columns 中的百分比相对网格容器内容盒计算，而 gap 是额外占位、不参与百分比分配，因此 50% + 50% + gap ⇒ 需求宽度 = 100% + 20px > 100%。", "p"),
    ("修改方案：改用 fr 单位 grid-template-columns: repeat(2, minmax(0, 1fr))。fr 会先扣除 gap 再分配剩余空间，minmax(0, 1fr) 还能防止长内容撑破轨道。", "p"),
    ("验证方式：DevTools 选中 .main 观察是否横向滚动；Console 执行 document.querySelector('.main').scrollWidth > document.querySelector('.main').clientWidth；拖动窗口从 1280px → 900px。", "p"),
    ("验证结果：修改前 scrollWidth = 980、clientWidth = 960（溢出 20px），出现横向滚动条；修改后 scrollWidth === clientWidth，无溢出，两列等宽且正确扣除 gap。", "p"),
    ("结论：%+gap 确实导致溢出，fr 修复有效，原判断得到验证。", "p"),
    ("证据二（边界视口）：.resume 写死 width: 1000px。", "sub"),
    ("把视口设为 375px（常见手机宽度）时页面出现横向滚动条，必须靠媒体查询专门覆盖 width: 100%，属于“桌面写死后靠断点救”，每多一个尺寸就要补一条规则。改为 width: min(1000px, 100% - 2rem) 后，375px 视口下容器自动收缩为视口宽减 2rem，侧栏与主区正确堆叠，无横向滚动。验证有效。", "p"),
    ("证据三（量化对比度测试）：对比度与行距不足。", "sub"),
    ("用 WebAIM Contrast Checker 实测，正文 #999 在白底上约 2.8:1、技能列表 #bbb 约 1.9:1，均远低于 WCAG 2.1 AA 要求的 4.5:1；同时全局没有 line-height，长段落行距过密。改为 #333 与 #555 并加 line-height: 1.6 后，实测 #333/#fff = 12.6:1、#555/#fff = 7.5:1，均通过 AA，行距明显改善。验证有效。这也让我确认：可访问性是可以量化的——对比度不是一个感觉，而是一个能算出来、能测出来的数。", "p"),
]

P60 = [
    ("AI 生成 Grid/响应式布局最常犯的错，是把“桌面写死宽度 + %+gap 网格 + 桌面优先的 max-width 媒体查询”当作响应式；真正的响应式应从 box-sizing、fr / minmax(0, 1fr) 和移动优先的 min-width 断点做起。", "p"),
    ("更关键的是断点要由内容决定（卡片挤了就分列、侧栏能独立成列就拆栏），而不是套用 768/600/900 这些现成数值；二维布局用 Grid、一维排列用 Flex，选择依据是维度而不是个人偏好。", "p"),
]

TARGETS = {  # 段落索引 -> 内容块
    19: P19, 21: P21, 23: P23, 25: P25,
    30: P30, 32: P32, 34: P34, 36: P36, 38: P38,
    41: P41, 43: P43, 45: P45, 47: P47, 49: P49,
    52: P52, 54: P54, 56: P56, 58: P58, 60: P60,
}


def style_para(p):
    pf = p.paragraph_format
    pf.first_line_indent = 0
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)


def main():
    d = docx.Document(DOC)
    cell = d.tables[1].rows[0].cells[0]
    paras = cell.paragraphs

    # 校验锚点
    for idx in TARGETS:
        assert paras[idx].text.strip() == "", f"索引 {idx} 不是空段落: {paras[idx].text[:30]!r}"

    # 从后往前处理，避免插入新段落影响前面的索引
    for idx in sorted(TARGETS, reverse=True):
        blocks = TARGETS[idx]
        anchor = paras[idx]
        style_para(anchor)
        for r in list(anchor.runs):
            r._element.getparent().remove(r._element)

        prev = anchor
        for order, (text, kind) in enumerate(blocks):
            if order == 0:
                target = anchor
            else:
                new_el = copy.deepcopy(anchor._element)
                prev._element.addnext(new_el)
                target = Paragraph(new_el, anchor._parent)
                for r in list(target.runs):
                    r._element.getparent().remove(r._element)
                prev = target

            style_para(target)
            run = target.add_run(text)
            if kind == "note":
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
            elif kind == "sub":
                run.font.bold = True
            elif kind == "code":
                run.font.name = "Consolas"
                run.font.size = Pt(9)

    d.save(DOC)
    print("Part2 填充完成，共处理", len(TARGETS), "个填空位")


if __name__ == "__main__":
    main()

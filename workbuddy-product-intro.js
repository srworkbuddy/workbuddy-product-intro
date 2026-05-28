const fs = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
        WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
        ExternalHyperlink } = fs;
const fss = require('fs');

const BLUE = "0052D9";
const LIGHT_BLUE = "E8F0FE";
const DARK = "1A1A1A";
const GRAY = "666666";
const LIGHT_GRAY = "F5F5F5";
const WHITE = "FFFFFF";
const ACCENT = "00A870";
const LINK_COLOR = "0052D9";
const GREEN_BG = "F0FFF5";

const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorders = { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } };

function headerCell(text, width) {
  return new TableCell({
    borders: cellBorders, width: { size: width, type: WidthType.DXA },
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, bold: true, size: 20, color: WHITE, font: "Microsoft YaHei" })] })]
  });
}

function dataCell(text, width, opts = {}) {
  return new TableCell({
    borders: cellBorders, width: { size: width, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: { before: 40, after: 40 },
      indent: opts.indent ? { left: 120 } : undefined,
      children: [new TextRun({ text, size: 18, color: opts.bold ? DARK : GRAY, bold: opts.bold, font: "Microsoft YaHei" })] })]
  });
}

function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, size: 32, bold: true, color: BLUE, font: "Microsoft YaHei" })] });
}

function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 300, after: 150 },
    children: [new TextRun({ text, size: 26, bold: true, color: DARK, font: "Microsoft YaHei" })] });
}

function heading3(text) {
  return new Paragraph({ spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, size: 22, bold: true, color: BLUE, font: "Microsoft YaHei" })] });
}

function bodyText(text, opts = {}) {
  return new Paragraph({ spacing: { before: 80, after: 80 }, indent: opts.indent ? { left: 360 } : undefined,
    children: [new TextRun({ text, size: 20, color: opts.color || DARK, font: "Microsoft YaHei", bold: opts.bold, italics: opts.italics })] });
}

function highlightBox(text) {
  return new Paragraph({ spacing: { before: 120, after: 120 }, indent: { left: 360, right: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 6, color: BLUE } },
    children: [new TextRun({ text, size: 20, color: BLUE, italics: true, font: "Microsoft YaHei" })] });
}

function greenBox(text) {
  return new Paragraph({ spacing: { before: 100, after: 100 }, indent: { left: 360, right: 360 },
    shading: { fill: GREEN_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text, size: 19, color: "006644", font: "Microsoft YaHei" })] });
}

function spacer(h = 100) {
  return new Paragraph({ spacing: { before: h, after: 0 }, children: [] });
}

function bullet(text, ref) {
  return new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, font: "Microsoft YaHei" })] });
}

// Material link section with clickable hyperlinks
function materialLinks(links) {
  const children = [
    new Paragraph({ spacing: { before: 160, after: 60 },
      children: [new TextRun({ text: "🔗 素材链接", size: 20, bold: true, color: BLUE, font: "Microsoft YaHei" })] })
  ];
  for (const link of links) {
    children.push(new Paragraph({ spacing: { before: 30, after: 30 }, indent: { left: 360 },
      children: [
        new TextRun({ text: "• ", size: 18, color: GRAY, font: "Microsoft YaHei" }),
        new ExternalHyperlink({ link: link.url, children: [
          new TextRun({ text: link.name, size: 18, color: LINK_COLOR, font: "Microsoft YaHei", underline: { type: "single" } })
        ]})
      ]
    }));
  }
  return children;
}

// Enterprise scene card generator
function enterpriseScene(title, painPoint, steps, customerValue, links) {
  const items = [];
  items.push(heading3(title));
  items.push(bodyText("痛点：\u201c + painPoint, { bold: false }));
  items.push(bodyText(\u201d实践路径：", { bold: true }));
  steps.forEach((step, i) => {
    items.push(new Paragraph({ numbering: { reference: "nl1", level: 0 }, spacing: { before: 30, after: 30 },
      children: [new TextRun({ text: step, size: 19, font: "Microsoft YaHei" })] }));
  });
  items.push(greenBox("客户价值：" + customerValue));
  items.push(...materialLinks(links));
  return items;
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Microsoft YaHei", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, color: BLUE, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: DARK, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 300, after: 150 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bl1", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl2", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl3", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl4", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl5", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl6", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl7", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "bl8", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "nl1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // Cover page
    {
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        spacer(2400),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [new TextRun({ text: "WorkBuddy", size: 72, bold: true, color: BLUE, font: "Microsoft YaHei" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
          children: [new TextRun({ text: "全场景职场 AI 原生智能体桌面工作台", size: 32, color: GRAY, font: "Microsoft YaHei" })] }),
        spacer(400),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
          children: [new TextRun({ text: "产品介绍 · 场景实践 · 使用案例", size: 24, color: DARK, font: "Microsoft YaHei" })] }),
        spacer(200),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "一句话让 AI 替你上班", size: 28, bold: true, color: ACCENT, font: "Microsoft YaHei" })] }),
        spacer(1600),
        new Paragraph({ alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "腾讯 CSIG · 2026年5月", size: 20, color: GRAY, font: "Microsoft YaHei" })] }),
      ]
    },
    // TOC
    {
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "WorkBuddy 产品介绍", size: 16, color: GRAY, font: "Microsoft YaHei" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "第 ", size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY }), new TextRun({ text: " 页\u201c, size: 16, color: GRAY })] })] }) },
      children: [
        new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: \u201d目录", size: 36, bold: true, color: BLUE, font: "Microsoft YaHei" })] }),
        new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-2" }),
      ]
    },
    // Main content
    {
      properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "WorkBuddy 产品介绍", size: 16, color: GRAY, font: "Microsoft YaHei" })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "第 ", size: 16, color: GRAY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GRAY }), new TextRun({ text: " 页\u201c, size: 16, color: GRAY })] })] }) },
      children: [
        // ====== 1. 产品概述 ======
        heading1(\u201d一、产品概述"),
        heading2("1.1 产品定位"),
        bodyText("WorkBuddy 是腾讯推出的桌面级 AI 助手，被媒体叫作\u201c腾讯版小龙虾\u201d。跟只能在对话框里回答问题的传统 AI 不一样，WorkBuddy 能直接操作你电脑上的文件\u2014\u2014你说一句话，它自己去规划、执行，最后把做好的东西交给你。2026 年 3 月 9 日上线。\u201c),
        highlightBox(\u201d产品于 2026 年 3 月 9 日正式上线，基于腾讯 CodeBuddy 同源架构打造"),
        heading2("1.2 核心差异化\u201c),
        bodyText(\u201d传统 AI 聊天工具 vs WorkBuddy 的本质区别：\u201c, { bold: true }),
        new Table({
          columnWidths: [2200, 3580, 3580],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d对比维度\u201c, 2200), headerCell(\u201d传统 AI 聊天工具", 3580), headerCell("WorkBuddy 智能体工作台\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d它的角色\u201c, 2200, {bold:true}), dataCell(\u201d写作顾问 / 建议提供者\u201c, 3580), dataCell(\u201d办公执行者 / 任务完成同事\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d交付结果\u201c, 2200, {bold:true}), dataCell(\u201d方法建议、模板、润色建议\u201c, 3580), dataCell(\u201d分类结果、回复草稿、正式邮件、可发送交付物\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d能否接触现场\u201c, 2200, {bold:true}), dataCell(\u201d不能，需用户复制粘贴\u201c, 3580), dataCell(\u201d能读取文件/邮箱/表格，理解上下文\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d用户要做什么\u201c, 2200, {bold:true}), dataCell(\u201d自己筛选、判断、写作、发送\u201c, 3580), dataCell(\u201d提出目标，检查结果，确认发送\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d核心差别\u201c, 2200, {bold:true, shading: LIGHT_BLUE}), dataCell(\u201d让员工\u201c知道怎么处理\u201d", 3580, {shading: LIGHT_BLUE}), dataCell("让员工\u201c直接拿到处理后的结果\u201d", 3580, {shading: LIGHT_BLUE})] }),
          ]
        }),

        // ====== 2. 核心能力 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1("二、核心能力"),
        heading2("2.1 智能任务自动化执行\u201c),
        bodyText(\u201d你说一句话，它自己把任务拆开、排好步骤、跑完全程。适合周期性、重复性的办公活儿。\u201c),
        bullet(\u201d自动生成日报/周报/季度报告", "bl1"),
        bullet("定时抓取竞品信息和行业动态", "bl1"),
        bullet("整理会议录音与纪要", "bl1"),

        heading2("2.2 本地文件全流程操作\u201c),
        bodyText(\u201d你授权之后，它可以直接读写你电脑上的文件——批量编辑、分类整理、格式转换、归档。\u201c),
        bullet(\u201d整理桌面文件按类型分类", "bl2"),
        bullet("批量重命名、分类归档", "bl2"),
        bullet("处理 Excel 表格、编辑 Word 文档", "bl2"),

        heading2("2.3 多模态办公内容创作\u201c),
        bodyText(\u201d写文案、做 PPT、画海报、出数据图表，风格、框架、细节随你调。"),

        heading2("2.4 专业数据处理与报表生成\u201c),
        bodyText(\u201d数据清洗、统计分析、图表可视化、自动出报表。表格数据跑一遍，核心指标和结论就出来了。"),

        heading2("2.5 跨平台远程电脑控制\u201c),
        bodyText(\u201d支持 MCP 协议，连接企业微信、QQ、飞书、钉钉等办公平台。手机上发一句话，办公室电脑就开始干活，做完把结果给你。\u201c),
        highlightBox(\u201d企微/QQ 三步连上，手机发指令，电脑端自动干活"),

        heading2("2.6 Skills 技能包系统\u201c),
        bodyText(\u201d自带 20 多个 Skills 技能包，不用写代码就能扩展新能力。日常操作可以封装成可复用的技能，也能在团队里分享、上传到开源社区。\u201c),
        bullet(\u201d办公效率全家桶（Word、Excel、PPT、PDF 处理）", "bl3"),
        bullet("内容创作与多平台适配（小红书、公众号、SEO）", "bl3"),
        bullet("垂直行业专业任务（法律、金融、人力资源）", "bl3"),

        heading2("2.7 多模型自由切换\u201c),
        bodyText(\u201d不绑定单一模型，五大主流大模型随意切：\u201c),
        new Table({
          columnWidths: [2200, 3580, 3580],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d模型\u201c, 2200), headerCell(\u201d特点\u201c, 3580), headerCell(\u201d适用场景\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d腾讯混元\u201c, 2200, {bold:true}), dataCell(\u201d强大中文理解、多模态处理\u201c, 3580), dataCell(\u201d报告撰写、文案生成、数据分析", 3580)] }),
            new TableRow({ children: [dataCell("DeepSeek-V3/R1", 2200, {bold:true}), dataCell("推理能力强、精度高\u201c, 3580), dataCell(\u201d复杂架构设计、调试任务", 3580)] }),
            new TableRow({ children: [dataCell("DeepSeek-V3.2", 2200, {bold:true}), dataCell("响应快、成本低\u201c, 3580), dataCell(\u201d日常简单查询、代码补全", 3580)] }),
            new TableRow({ children: [dataCell("GLM-5 系列\u201c, 2200, {bold:true}), dataCell(\u201d中文语义理解优秀\u201c, 3580), dataCell(\u201d中文语境复杂语义处理", 3580)] }),
            new TableRow({ children: [dataCell("Kimi-K2 系列\u201c, 2200, {bold:true}), dataCell(\u201d超长上下文（256K）\u201c, 3580), dataCell(\u201d超长文档阅读、多文档分析", 3580)] }),
            new TableRow({ children: [dataCell("MiniMax-M2 系列\u201c, 2200, {bold:true}), dataCell(\u201d多媒体内容生成\u201c, 3580), dataCell(\u201d内容创作、角色扮演\u201c, 3580)] }),
          ]
        }),

        // ====== 3. 企业版飞轮 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1(\u201d三、企业版核心逻辑：AI 工作飞轮\u201c),
        highlightBox(\u201d企业版不是给员工多一个 AI 聊天框。它跑的是一条飞轮：统一入口 → 超级个体 → 超级团队 → 企业 AI 资产 → 企业数字员工 → 组织生产力持续释放。"),
        ...materialLinks([
          { name: "WorkBuddy企业版主线逻辑说明材料", url: "https://csig.lexiangla.com/pages/dd473bab87ae46ab91a1049a101d35b4" }
        ]),
        heading2("3.1 五大核心概念\u201c),
        new Table({
          columnWidths: [1800, 3780, 3780],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d概念\u201c, 1800), headerCell(\u201d定义\u201c, 3780), headerCell(\u201d价值\u201c, 3780)] }),
            new TableRow({ children: [dataCell(\u201d超级个体\u201c, 1800, {bold:true}), dataCell(\u201d员工用 WorkBuddy 调用组织能力干活\u201c, 3780), dataCell(\u201d个人能力直接放大，普通员工也能做更专业的活儿\u201c, 3780)] }),
            new TableRow({ children: [dataCell(\u201d超级团队\u201c, 1800, {bold:true}), dataCell(\u201d人 + 数字员工组队干活\u201c, 3780), dataCell(\u201d团队围绕任务持续推进，不只是沟通\u201c, 3780)] }),
            new TableRow({ children: [dataCell(\u201d企业 AI 资产\u201c, 1800, {bold:true}), dataCell(\u201d每次任务沉淀知识、技能、经验和案例\u201c, 3780), dataCell(\u201d经验从个人释放，变成组织可复用能力\u201c, 3780)] }),
            new TableRow({ children: [dataCell(\u201d企业数字员工", 1800, {bold:true}), dataCell("7×24 在线，懂岗位、有技能、熟悉企业知识\u201c, 3780), dataCell(\u201d数字员工越来越懂企业，能承担更多专业任务\u201c, 3780)] }),
            new TableRow({ children: [dataCell(\u201d统一治理\u201c, 1800, {bold:true}), dataCell(\u201d管理员统一管控智能体的权限、安全和审计\u201c, 3780), dataCell(\u201d企业敢把更多流程交给 AI，形成规模化落地", 3780)] }),
          ]
        }),

        heading2("3.2 Agent Suite 企业套件\u201c),
        new Table({
          columnWidths: [2200, 3580, 3580],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d套件\u201c, 2200), headerCell(\u201d定位\u201c, 3580), headerCell(\u201d价值\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d腾讯文档\u201c, 2200, {bold:true}), dataCell(\u201d内容生产与协作底座", 3580), dataCell("AI 生成的东西直接进协作体系\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d腾讯网盘\u201c, 2200, {bold:true}), dataCell(\u201d文件和素材管理底座\u201c, 3580), dataCell(\u201d本地产物、团队共享文件都能存、能回溯\u201c, 3580)] }),
            new TableRow({ children: [dataCell(\u201d腾讯乐享\u201c, 2200, {bold:true}), dataCell(\u201d知识治理底座\u201c, 3580), dataCell(\u201d制度、培训、经验变成数字员工的知识\u201c, 3580)] }),
          ]
        }),

        // ====== 4. 使用场景 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1(\u201d四、使用场景"),

        // --- 4.1 办公场景 ---
        heading2("4.1 办公场景：高效信息沟通\u201c),
        heading3(\u201d场景描述\u201c),
        bodyText(\u201d周一早上 9 点，员工打开邮箱发现 38 封未读邮件，包含领导催进度、客户报价、HR 表格、项目协作、会议通知等不同类型事项。"),
        heading3("WorkBuddy 做法\u201c),
        bullet(\u201d读取未读邮件汇总文件，识别每封邮件的业务含义和处理责任", "nl1"),
        bullet("将 38 封邮件分为\u201c必须亲自回/助理代回/仅知会/可忽略\u201d4 类", "nl1"),
        bullet("为\u201c必须亲自回\u201d的邮件生成专业中文回复草稿", "nl1"),
        bullet("处理重点英文客户邮件：先生成中文审稿，再翻译成商务英语版", "nl1"),
        bullet("员工确认后可继续完成发送或回写", "nl1"),
        heading3("提效效果\u201c),
        bodyText(\u201d传统方式 2-3 小时 → WorkBuddy 10-20 分钟，员工从执行者变成审核者。"),
        ...materialLinks([
          { name: "WorkBuddy办公场景01_高效信息沟通_最佳实践案例", url: "https://csig.lexiangla.com/pages/e769ffa538b64e888380bf1659f58e9e" },
          { name: "WorkBuddy&办公场景01：高效信息沟通（视频）", url: "https://csig.lexiangla.com/pages/a7d93805da24427e98c01b1880a3eca4" },
          { name: "传统AI聊天_vs_WorkBuddy干活_最佳实践案例", url: "https://csig.lexiangla.com/pages/b96442a34bef4915b0e42ac078116d29" }
        ]),

        // --- 4.2 PM场景 ---
        heading2("4.2 PM 场景：项目过程管理\u201c),
        new Table({
          columnWidths: [800, 1800, 2480, 2080, 2200],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d序号\u201c, 800), headerCell(\u201d场景\u201c, 1800), headerCell(\u201d典型输入\u201c, 2480), headerCell(\u201d标准输出\u201c, 2080), headerCell(\u201d直接价值", 2200)] }),
            new TableRow({ children: [dataCell("1", 800), dataCell("会议纪要闭环\u201c, 1800, {bold:true}), dataCell(\u201d会议记录/聊天记录\u201c, 2480), dataCell(\u201d结论/行动项/责任人/DDL/风险项\u201c, 2080), dataCell(\u201d会后闭环速度提升", 2200)] }),
            new TableRow({ children: [dataCell("2", 800), dataCell("周报/月报生成\u201c, 1800, {bold:true}), dataCell(\u201d本周事项、风险、里程碑\u201c, 2480), dataCell(\u201d管理层口径周报草稿\u201c, 2080), dataCell(\u201d写作时间下降、口径更统一", 2200)] }),
            new TableRow({ children: [dataCell("3", 800), dataCell("验收材料打包\u201c, 1800, {bold:true}), dataCell(\u201d验收报告、清单、截图\u201c, 2480), dataCell(\u201d可提交材料包+检查清单\u201c, 2080), dataCell(\u201d减少漏项、提升交付质量", 2200)] }),
            new TableRow({ children: [dataCell("4", 800), dataCell("项目知识检索\u201c, 1800, {bold:true}), dataCell(\u201d历史PRD、纪要、周报、合同\u201c, 2480), dataCell(\u201d关键结论与出处\u201c, 2080), dataCell(\u201d减少\u201c找资料\u201d耗时", 2200)] }),
            new TableRow({ children: [dataCell("5", 800), dataCell("远程触发任务\u201c, 1800, {bold:true}), dataCell(\u201d微信/企微/QQ 指令\u201c, 2480), dataCell(\u201d周报草稿、资料归档结果\u201c, 2080), dataCell(\u201d碎片时间也能推进项目", 2200)] }),
          ]
        }),
        ...materialLinks([
          { name: "WorkBuddy-PM在日常工作应用场景实践-v5", url: "https://csig.lexiangla.com/pages/41aa17715ff64c15a8bd662ab8963cff" },
          { name: "TAPD + 乐享 + WorkBuddy 联动方案", url: "https://csig.lexiangla.com/pages/98928c2886ac42e58d289bd6c8878a8d" }
        ]),

        // --- 4.3 企业版场景实践（8个详细场景） ---
        new Paragraph({ children: [new PageBreak()] }),
        heading2("4.3 企业版场景实践\u201c),
        bodyText(\u201d以下场景均来自 WorkBuddy 企业版主线逻辑与 WorkBuddy+乐享联合方案，每个场景包含客户痛点、实践路径和客户可感知价值。\u201c),

        // 场景一：售后/客服
        ...enterpriseScene(
          \u201d场景一：售后 / 客服问题处理\u201c,
          \u201d客服培训慢，复杂问题升级多，历史案例和知识库难以复用，新人不会查也不知道哪个答案是对的，前线问题没有被反哺成组织知识\u201c,
          [
            \u201d售后同学每天打开 WorkBuddy 处理客户问题\u201c,
            \u201d调用乐享知识库中的 FAQ、故障案例、操作手册和话术模板\u201c,
            \u201d售后数字员工整理可能原因、排查顺序和客户回复草稿\u201c,
            \u201d需要升级时，数字员工加入 Teams，与产品、研发和二线支持协同\u201c,
            \u201d最终回复、排查记录和新案例沉淀回企业知识库\u201c,
            \u201d沉淀的案例和 SOP 再赋能下一次售后数字员工\u201c
          ],
          \u201d客户回复更快 | 排查路径更标准 | 新人更容易上手 | 案例持续沉淀 | 售后经验从个人能力变成组织能力",
          [
            { name: "WorkBuddy企业版主线逻辑说明材料", url: "https://csig.lexiangla.com/pages/dd473bab87ae46ab91a1049a101d35b4" },
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // 场景二：销售/售前
        ...enterpriseScene(
          "场景二：销售 / 售前材料快速响应\u201c,
          \u201d客户需求一来资料找不到或版本不统一，不同销售讲法不一致影响专业形象，售前专家被反复拉来回答基础问题，方案、案例、FAQ、话术散在不同地方\u201c,
          [
            \u201d乐享统一沉淀：产品资料、行业方案、成功案例、招投标模板、常见异议回复、实施边界与报价规则\u201c,
            \u201d销售在企微或桌面侧向 WorkBuddy 发起任务",
            "WorkBuddy 调用乐享知识库内容后，输出：客户方案初稿、沟通邮件、拜访前准备提纲、行业案例摘要、会后纪要与跟进清单\u201c,
            \u201d团队在 Teams 中协同修改，产品、售前和销售共同审阅\u201c,
            \u201d定稿保存到腾讯文档，证据来源沉淀到销售弹药库\u201c,
            \u201d乐享负责\u201c\u201c找准内容\u201d\u201d，WorkBuddy 负责"\u201c变成输出\u201d""
          ],
          "客户响应速度提升 | 销售口径一致性提高 | 资深售前重复支持减少 | 方案制作时间缩短 | 销售输出建立在企业知识基础上",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" },
            { name: "workbuddy营销部门或市场部门场景", url: "https://csig.lexiangla.com/pages/ae018f44e3d64f96bfc6425df69463ed" }
          ]
        ),

        // 场景三：会议闭环
        ...enterpriseScene(
          "场景三：会议到任务闭环\u201c,
          \u201d会议纪要有了，但行动项没人跟，责任人和后续进展容易断，会后闭环全靠人追\u201c,
          [
            \u201d会议记录、聊天记录、语音转写进入 WorkBuddy",
            "数字员工提炼结论、行动项、责任人、DDL 和风险项",
            "Teams 把行动项分配给成员和数字员工\u201c,
            \u201d自动化任务定期提醒、跟踪和汇总进展\u201c,
            \u201d后续产出、复盘和决策沉淀为团队资产\u201c,
            \u201d下一次会议可直接复用这些组织记忆\u201c
          ],
          \u201d会议不再止于纪要 | 行动项持续推进 | 任务状态可见 | 复盘资料可沉淀 | 团队协同从沟通走向执行",
          [
            { name: "WorkBuddy-PM在日常工作应用场景实践-v5", url: "https://csig.lexiangla.com/pages/41aa17715ff64c15a8bd662ab8963cff" }
          ]
        ),

        // 场景四：企业知识运营
        ...enterpriseScene(
          "场景四：企业知识运营\u201c,
          \u201d企业不缺知识，但知识散在文档、网盘、乐享、业务系统和群聊里，员工找不到、找不准、也不知道能不能用，对合规敏感的岗位担心 AI 乱答、越权、引用过期内容\u201c,
          [
            \u201d员工在 WorkBuddy 中直接询问制度、流程、产品资料和项目经验",
            "WorkBuddy 通过 Agent Suite 调用腾讯文档、腾讯网盘和腾讯乐享\u201c,
            \u201d乐享负责知识治理，确保知识可信、可控、可追溯（权限、有效期、版本管理）\u201c,
            \u201d数字员工根据企业知识完成问答、总结、分析和材料生成\u201c,
            \u201d新问题、新回答、新材料和知识缺口沉淀为企业 AI 资产\u201c,
            \u201d这些资产继续赋能知识运营数字员工和业务数字员工\u201c
          ],
          \u201d知识找得到 | 回答答得准 | 权限管得住 | 结果可追溯 | 企业知识越用越好",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // 场景五：新人培训
        ...enterpriseScene(
          "场景五：新人培训与岗位辅导\u201c,
          \u201d培训资料很多但不好找不成体系，培训结束后员工仍不会实操，带教大量依赖老员工一对一答疑，新人高频问题重复出现占用管理精力\u201c,
          [
            \u201d乐享沉淀岗位 SOP、流程规范、制度说明、典型案例、常见错误、模板和标准口径、培训资料与题库\u201c,
            \u201d新人遇到问题，通过知识库快速找到答案",
            "WorkBuddy 进一步生成：学习清单、每周学习计划、错题/易错点总结、岗位周报、导师跟进建议\u201c,
            \u201d把培训从静态课件变成动态辅导\u201c,
            \u201d把组织经验留在系统里，而不是留在个人脑子里\u201c
          ],
          \u201d缩短新人上岗周期 | 降低培训和带教成本 | 减少基础问题反复答疑 | 培训从静态课件变动态辅导",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // 场景六：管理层周报
        ...enterpriseScene(
          "场景六：管理层周报 / 例会 / 经营复盘\u201c,
          \u201d信息散落在会议纪要、项目材料、沟通记录里，每周花大量时间收集整理，周报和复盘高度依赖手工搬运，同一件事描述不统一影响决策效率\u201c,
          [
            \u201d乐享沉淀日常经营类内容：项目纪要、复盘文档、制度更新、风险记录、业务分析、运营日报/周报素材\u201c,
            \u201d管理者通过 WorkBuddy 发起：\u201c\u201c整理本周重点事项和风险点\u201d\u201d、"\u201c生成销售周报\u201d"、"\u201c汇总项目例会待办\u201d"",
            "WorkBuddy 生成：周报、例会提纲、跟进行动清单、经营摘要、风险项列表\u201c,
            \u201d接入企微后支持结果推送和远程办公场景\u201c
          ],
          \u201d节省管理层和助理整理时间 | 强化经营信息统一口径 | 提高例会和复盘效率 | 知识沉淀进入经营管理闭环",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // 场景七：移动办公
        ...enterpriseScene(
          "场景七：移动办公 / 出差办公助手\u201c,
          \u201d人在外面资料在公司电脑里，时间碎片化需要用手机快速处理工作，找人代查资料效率低还打断别人\u201c,
          [
            \u201d企业知识统一沉淀在乐享中，供检索和调用\u201c,
            \u201d人在外部通过企微向 WorkBuddy 下发任务",
            "WorkBuddy 调用知识并在本地电脑完成操作，再把结果传回企微\u201c,
            \u201d示例：手机企微发起"\u201c把客户X的最新案例、方案模板和会后纪要发给我，生成跟进邮件\u201d" → 系统调用乐享资料 → WorkBuddy 整理邮件并处理文件 → 结果直接推回企微\u201c
          ],
          \u201d移动办公效率提升 | 减少跨设备跨人协调成本 | 出差岗位随时获得组织知识支持 | "\u201c找资料-处理-回传\u201d"一次做完",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // 场景八：制度/合规
        ...enterpriseScene(
          "场景八：制度 / 合规 / 专业支持\u201c,
          \u201d内容是否准确、权限是否可控、版本是否最新、使用过程是否可追溯。只有 Agent 没有知识治理底座，客户容易担心风险\u201c,
          [
            \u201d乐享沉淀：制度规范、合规要求、审批规则、专业模板、历史案例、审核意见、业务指引\u201c,
            \u201d通过权限和有效期控制，保证不同角色看到的是正确、有效、可授权的内容",
            "WorkBuddy 在受控知识基础上生成：制度答复草稿、审核清单、材料整理初稿、内部通知、执行提醒",
            "AI 使用建立在合规治理基础上\u201c
          ],
          \u201d减少错用制度和版本错误 | 提高专业岗位工作规范性 | 降低业务风险 | AI 使用建立在合规治理基础上",
          [
            { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
          ]
        ),

        // ====== 5. 使用案例 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1("五、使用案例"),

        // --- 案例1: 学术信息助手 ---
        heading2("5.1 案例一：构建个人学术信息助手\u201c),
        heading3(\u201d背景\u201c),
        bodyText(\u201d搜广推领域算法人员，需要持续关注学术前沿，但信息获取主要靠被动刷社媒，碎片化且缺乏针对性。"),
        heading3("WorkBuddy 解决方案\u201c),
        bodyText(\u201d从 0 到 1 搭建了 arXiv 论文每日自动筛选、精排、深度解读、报告生成、知识库沉淀的全流程自动化系统。\u201c),
        bullet(\u201d三路并行召回：关键词 + Embedding 语义 + 高频作者追踪，召回率 97.4%", "bl4"),
        bullet("9 维结构化精读框架（参考 ICML/NeurIPS 审稿标准 + KDD ADS Track 评审标准）", "bl4"),
        bullet("每个工作日早上 9 点自动生成 10 篇精读 + 20 篇速览 + 趋势分析", "bl4"),
        bullet("自动上传乐享知识库 + 邮件通知，每日自动化执行", "bl4"),
        bullet("搭建 2 天，代码约 5000 行全部由 WorkBuddy 生成，外部 API 成本为零", "bl4"),
        heading3("核心成果\u201c),
        new Table({
          columnWidths: [4680, 4680],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d维度\u201c, 4680), headerCell(\u201d数据\u201c, 4680)] }),
            new TableRow({ children: [dataCell(\u201d搭建耗时", 4680, {bold:true}), dataCell("~2 天（从 0 到全自动化）\u201c, 4680)] }),
            new TableRow({ children: [dataCell(\u201d代码量", 4680, {bold:true}), dataCell("~5000 行 Python（全部由 WorkBuddy 生成）\u201c, 4680)] }),
            new TableRow({ children: [dataCell(\u201d每日阅读耗时", 4680, {bold:true}), dataCell("~5 分钟过一遍推送\u201c, 4680)] }),
            new TableRow({ children: [dataCell(\u201d召回率", 4680, {bold:true}), dataCell("97.4%（基于 Zotero 343 篇论文验证）\u201c, 4680)] }),
            new TableRow({ children: [dataCell(\u201d外部 API 成本", 4680, {bold:true}), dataCell("¥0（使用 WorkBuddy 原生 LLM 能力）", 4680)] }),
          ]
        }),
        ...materialLinks([
          { name: "WorkBuddy 提效案例：构建个人学术信息助手", url: "https://csig.lexiangla.com/pages/e2cf5cbc5a8f44d28176829c92dbbe6e" }
        ]),

        // --- 案例2: AI简历筛选 ---
        heading2("5.2 案例二：AI 简历筛选提效 90%"),
        bodyText("新加坡团队使用 WorkBuddy 进行 AI 简历筛选：输入岗位需求和候选人简历文件，WorkBuddy 自动解析简历内容、按岗位要求逐项匹配评分、生成筛选结果和推荐理由。人工筛选时间从几小时压到几分钟，筛选的一致性和准确度也提高了，避免了不同 HR 主观判断导致的标准偏差。\u201c),
        ...materialLinks([
          { name: \u201d【案例】新加坡-AI简历筛选提效90%", url: "https://csig.lexiangla.com/pages/bfce4bba7e754e38826b50396ced359d" }
        ]),

        // --- 案例3: 五粮液 ---
        heading2("5.3 案例三：五粮液全民 AI 办公提效动员方案\u201c),
        bodyText(\u201d基于乐享、企微、WorkBuddy 等工具，给五粮液领导层做了一套全民 AI 办公提效动员方案，从知识管理到智能办公整个链路跑通了。方案覆盖白酒行业典型办公场景：知识检索、会议纪要闭环、周报自动生成、质检提效等。WorkBuddy 在白酒行业的第一个客户案例，验证了制造业+AI办公的落地路径。\u201c),
        ...materialLinks([
          { name: \u201d基于乐享、企微、workbuddy等生成五粮液全民AI办公提效动员方案", url: "https://csig.lexiangla.com/pages/2b1e103954414fedad7c66094930b402" },
          { name: "五粮液：AI助力质检提效解决方案", url: "https://csig.lexiangla.com/pages/13fe256cb6184810a790fc1e045247b1" }
        ]),

        // --- 案例4: 冠盛 ---
        heading2("5.4 案例四：冠盛客户 WorkBuddy 场景沟通实践\u201c),
        bodyText(\u201d冠盛客户团队和 WorkBuddy 团队做了一次场景沟通，把制造业日常办公的痛点和需求梳理了一遍。围绕售后问题处理、销售材料快速响应、新人培训、管理层周报复盘等场景逐一验证，确认了 WorkBuddy 在制造业场景能落地。沟通结论：WorkBuddy + 乐享联合方案可以覆盖制造企业核心办公痛点。\u201c),

        // --- 案例5: WorkBuddy+乐享联合方案 ---
        heading2("5.5 案例五：WorkBuddy + 乐享 AI 知识库联合方案落地"),
        bodyText("WorkBuddy 与乐享 AI 知识库联合方案已在多个客户场景验证：乐享负责知识沉淀、权限治理、搜索问答和知识运营，WorkBuddy 负责桌面 AI 智能体、本地执行、企微接入和结果回传。两者组合后形成\u201c\u201c知识获取+任务执行+结果交付\u201d\u201d完整链路。已覆盖六大场景：销售/售前快速响应、客服/客成/实施支持、新人培训与岗位辅导、管理层周报/例会/经营复盘、制度/合规/专业支持、移动办公/出差办公助手。"),
        ...materialLinks([
          { name: "WorkBuddy + 乐享AI知识库联合方案", url: "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52" }
        ]),

        // --- 案例6: PM全场景提效 ---
        heading2("5.6 案例六：PM 项目过程管理全场景提效\u201c),
        bodyText(\u201d针对项目 PM 的 5 个高频场景（会议纪要闭环、周报/月报生成、验收材料打包、项目知识检索、移动端远程触发），WorkBuddy 提供了从输入到输出的标准化实操路径。核心思路：PM 从\u201c\u201c从零写报告\u201d\u201d转为\u201c\u201c审核和定稿\u201d\u201d，从\u201c\u201c找资料靠记忆\u201d\u201d变成\u201c\u201c按出处检索\u201d\u201d，从\u201c\u201c等回工位再做\u201d\u201d变成\u201c\u201c先启动任务回位后收口\u201d\u201d。建议先从纪要闭环+周报生成开始试点。"),
        ...materialLinks([
          { name: "WorkBuddy-PM在日常工作应用场景实践-v5", url: "https://csig.lexiangla.com/pages/41aa17715ff64c15a8bd662ab8963cff" },
          { name: "TAPD + 乐享 + WorkBuddy 联动方案", url: "https://csig.lexiangla.com/pages/98928c2886ac42e58d289bd6c8878a8d" }
        ]),

        // ====== 6. 安全保障 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1("六、数据安全与合规"),
        heading2("6.1 三重安全机制\u201c),
        bullet(\u201d本地执行：所有文件处理在本地完成，原始文件从不离开用户电脑", "bl5"),
        bullet("权限可控：文件访问需用户授权，企业管理员可统一配置权限策略", "bl5"),
        bullet("审计留痕：每次执行生成唯一 TraceID，绑定真实员工账号，可追溯可审计", "bl5"),
        heading2("6.2 合规认证\u201c),
        bullet(\u201d通过信通院 Claw 可信能力评估（5 大维度 57 项指标）", "bl6"),
        bullet("通过信通院智能原生编码 IDE 评估 4+ 级", "bl6"),
        bullet("同源工具 CodeBuddy 新增代码 100% 由 AI 自举生成（Code 自举）", "bl6"),

        // ====== 7. CodeBuddy 同源架构 ======
        heading1("七、CodeBuddy 同源架构"),
        bodyText("WorkBuddy 基于腾讯云代码助手 CodeBuddy 同源架构打造，CodeBuddy 提供三种形态覆盖不同用户：\u201c),
        new Table({
          columnWidths: [2200, 3580, 3580],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d产品形态\u201c, 2200), headerCell(\u201d适用用户\u201c, 3580), headerCell(\u201d核心特点", 3580)] }),
            new TableRow({ children: [dataCell("CodeBuddy IDE", 2200, {bold:true}), dataCell("产品/设计师/全栈开发/初学者\u201c, 3580), dataCell(\u201d产设研一体化工作台，"对话即编程"", 3580)] }),
            new TableRow({ children: [dataCell("CodeBuddy 插件\u201c, 2200, {bold:true}), dataCell(\u201d日常编码开发者\u201c, 3580), dataCell(\u201d即插即用、融入现有工作流", 3580)] }),
            new TableRow({ children: [dataCell("CodeBuddy Code", 2200, {bold:true}), dataCell("DevOps/运维/SRE/资深开发者\u201c, 3580), dataCell(\u201d命令行工具，任务编排能力强\u201c, 3580)] }),
          ]
        }),
        spacer(100),
        bodyText(\u201d核心认证与指标：\u201c, { bold: true }),
        bullet(\u201d信通院智能原生编码 IDE 评估 4+ 级", "bl7"),
        bullet("3 端覆盖（IDE / 插件 / CLI），数十款主流 IDE 支持", "bl7"),
        bullet("100% Code 自举新增代码由 AI 生成", "bl7"),

        // ====== 8. 内网素材索引 ======
        new Paragraph({ children: [new PageBreak()] }),
        heading1("八、内网素材索引\u201c),
        bodyText(\u201d以下是本次搜索整理的内网核心素材清单，均存储在乐享知识库，点击链接可直达原文：\u201c),
        new Table({
          columnWidths: [600, 3200, 3000, 2560],
          rows: [
            new TableRow({ tableHeader: true, children: [headerCell(\u201d序号\u201c, 600), headerCell(\u201d材料名称\u201c, 3200), headerCell(\u201d所在知识库\u201c, 3000), headerCell(\u201d链接", 2560)] }),
            ...[
              ["1", "WorkBuddy办公场景01_高效信息沟通_最佳实践案例", "WorkBuddy资料库", "https://csig.lexiangla.com/pages/e769ffa538b64e888380bf1659f58e9e"],
              ["2", "WorkBuddy-PM在日常工作应用场景实践-v5", "AI应用", "https://csig.lexiangla.com/pages/41aa17715ff64c15a8bd662ab8963cff"],
              ["3", "WorkBuddy企业版主线逻辑说明材料", "WorkBuddy资料库", "https://csig.lexiangla.com/pages/dd473bab87ae46ab91a1049a101d35b4"],
              ["4", "WorkBuddy + 乐享AI知识库联合方案", "WorkBuddy资料库", "https://csig.lexiangla.com/pages/0be95da3d040443bacab438f7c678c52"],
              ["5", "WorkBuddy 提效案例：构建个人学术信息助手\u201c, \u201d玩转亮点功能", "https://csig.lexiangla.com/pages/e2cf5cbc5a8f44d28176829c92dbbe6e"],
              ["6", "【案例】新加坡-AI简历筛选提效90%", "OHR AI Sharing", "https://csig.lexiangla.com/pages/bfce4bba7e754e38826b50396ced359d"],
              ["7", "五粮液全民AI办公提效动员方案\u201c, \u201d白酒行业AI应用", "https://csig.lexiangla.com/pages/2b1e103954414fedad7c66094930b402"],
              ["8", "五粮液：AI助力质检提效解决方案\u201c, \u201d交付保障中心综合知识库", "https://csig.lexiangla.com/pages/13fe256cb6184810a790fc1e045247b1"],
              ["10", "workbuddy营销部门或市场部门场景\u201c, \u201d培训售前素材", "https://csig.lexiangla.com/pages/ae018f44e3d64f96bfc6425df69463ed"],
              ["11", "TAPD + 乐享 + WorkBuddy 联动方案\u201c, \u201d工作日志", "https://csig.lexiangla.com/pages/98928c2886ac42e58d289bd6c8878a8d"],
              ["12", "腾讯企业版龙虾方案医疗行业", "AI应用", "https://csig.lexiangla.com/pages/07e1a40b226e4a3a98f6daa47a65e1d3"],
              ["13", "Workbuddy 应用场景举例（课纲示例）\u201c, \u201d演示物料", "https://csig.lexiangla.com/pages/a704a6b92f864d40a92eefc363fd3dc6"],
              ["14", "传统AI聊天_vs_WorkBuddy干活_最佳实践案例", "WorkBuddy资料库", "https://csig.lexiangla.com/pages/b96442a34bef4915b0e42ac078116d29"],
              ["15", "WorkBuddy&办公场景01：高效信息沟通（视频）", "WorkBuddy资料库", "https://csig.lexiangla.com/pages/a7d93805da24427e98c01b1880a3eca4"],
            ].map(([num, name, kb, url]) =>
              new TableRow({ children: [
                dataCell(num, 600),
                dataCell(name, 3200),
                dataCell(kb, 3000),
                new TableCell({ borders: cellBorders, width: { size: 2560, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
                  children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [
                    new ExternalHyperlink({ link: url, children: [
                      new TextRun({ text: "查看原文", size: 16, color: LINK_COLOR, font: "Microsoft YaHei", underline: { type: "single" } })
                    ]})
                  ]})]
                })
              ] })
            )
          ]
        }),
        spacer(100),
        highlightBox("所有材料均可直接点击链接访问，\u201cWorkBuddy资料库（持续更新）\u201d含售前弹药库和售中/后材料两个子文件夹"),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fss.writeFileSync("/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/WorkBuddy产品介绍_场景与案例.docx", buffer);
  console.log("Document created successfully!");
});

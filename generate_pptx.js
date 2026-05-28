const pptxgen = require('pptxgenjs');
const path = require('path');

async function createPresentation() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'Tencent CSIG';
  pptx.title = 'WorkBuddy 产品介绍';

  const BLUE = '0052D9';
  const DARK = '1A1A1A';
  const GRAY = '666666';
  const GREEN = '00A870';
  const LIGHT_BLUE = 'E8F0FE';
  const GREEN_BG = 'F0FFF5';
  const LEXIANG_GOLD = 'FFF8E6';
  const LEXIANG_TEXT = '8B6914';

  // ===== Slide 1: Cover =====
  let slide = pptx.addSlide();
  slide.background = { color: BLUE };
  slide.addText('WorkBuddy', { x: 1, y: 1.8, w: 8, h: 1.2, fontSize: 48, bold: true, color: 'FFFFFF', align: 'center' });
  slide.addText('全场景职场 AI 原生智能体桌面工作台', { x: 1, y: 3, w: 8, h: 0.6, fontSize: 18, color: 'FFFFFF', align: 'center' });
  slide.addText('产品介绍 · 场景实践 · 使用案例', { x: 1, y: 3.8, w: 8, h: 0.4, fontSize: 14, color: 'FFFFFF', align: 'center' });
  slide.addText('一句话让 AI 替你上班', { x: 1, y: 4.5, w: 8, h: 0.5, fontSize: 20, bold: true, color: GREEN, align: 'center' });
  slide.addText('腾讯 CSIG · 2026年5月', { x: 1, y: 6.5, w: 8, h: 0.3, fontSize: 12, color: 'FFFFFF', align: 'center' });

  // ===== Slide 2: Overview =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('产品概述', { x: 0.6, y: 0.3, w: 5, h: 0.4, fontSize: 10, color: BLUE });
  slide.addText('产品概述', { x: 0.6, y: 0.6, w: 5, h: 0.6, fontSize: 24, bold: true, color: DARK });
  slide.addText([
    { text: 'WorkBuddy 是腾讯推出的桌面级 AI 助手，被媒体叫作"腾讯版小龙虾"。跟只能在对话框里回答问题的传统 AI 不一样，WorkBuddy 能直接操作你电脑上的文件——你说一句话，它自己去规划、执行，最后把做完的东西交给你。' },
  ], { x: 0.6, y: 1.5, w: 8.8, h: 1, fontSize: 12, color: DARK, lineSpacingMultiple: 1.4 });
  slide.addText('产品于 2026 年 3 月 9 日正式上线，基于腾讯 CodeBuddy 同源架构打造', { x: 0.6, y: 2.6, w: 8.8, h: 0.4, fontSize: 11, color: BLUE, italic: true });

  // Core diff table
  const diffRows = [
    [{ text: '对比维度', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '传统 AI 聊天工具', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: 'WorkBuddy 智能体工作台', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }],
    ['它的角色', '写作顾问 / 建议提供者', '办公执行者 / 任务完成同事'],
    ['交付结果', '方法建议、模板、润色建议', '分类结果、回复草稿、正式邮件、可发送交付物'],
    ['能否接触现场', '不能，需用户复制粘贴', '能读取文件/邮箱/表格，理解上下文'],
    ['用户要做什么', '自己筛选、判断、写作、发送', '提出目标，检查结果，确认发送'],
    [{ text: '核心区别', options: { bold: true } }, '让员工"知道怎么处理"', '让员工"直接拿到处理后的结果"'],
  ];
  slide.addTable(diffRows, { x: 0.6, y: 3.2, w: 8.8, colW: [2, 3.4, 3.4], fontSize: 10, border: { pt: 0.5, color: 'DDDDDD' }, autoPage: false });

  // ===== Slide 3: Core Capabilities =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('核心能力', { x: 0.6, y: 0.6, w: 5, h: 0.6, fontSize: 24, bold: true, color: DARK });

  const capabilities = [
    ['智能任务自动化执行', '你说一句话，它自己把任务拆开、排好步骤、跑完全程'],
    ['本地文件全流程操作', '直接读写电脑文件——批量编辑、分类整理、格式转换、归档'],
    ['多模态办公内容创作', '写文案、做 PPT、画海报、出数据图表'],
    ['专业数据处理与报表生成', '数据清洗、统计分析、图表可视化、自动出报表'],
    ['跨平台远程电脑控制', '连接企微/QQ/飞书/钉钉，手机发指令电脑自动干活'],
    ['Skills 技能包系统', '20+技能包，不用写代码扩展新能力，可封装复用和团队分享'],
  ];
  capabilities.forEach((cap, i) => {
    const y = 1.4 + i * 0.85;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 8.8, h: 0.7, fill: { color: LIGHT_BLUE }, rectRadius: 0.1 });
    slide.addText(cap[0], { x: 0.8, y: y + 0.05, w: 3, h: 0.3, fontSize: 12, bold: true, color: BLUE });
    slide.addText(cap[1], { x: 0.8, y: y + 0.35, w: 8.4, h: 0.3, fontSize: 10, color: GRAY });
  });

  // ===== Slide 4: Multi-model =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('多模型自由切换', { x: 0.6, y: 0.6, w: 5, h: 0.6, fontSize: 24, bold: true, color: DARK });
  const modelRows = [
    [{ text: '模型', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '特点', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '适用场景', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }],
    ['腾讯混元', '强大中文理解、多模态处理', '报告撰写、文案生成、数据分析'],
    ['DeepSeek-V3/R1', '推理能力强、精度高', '复杂架构设计、调试任务'],
    ['GLM-5 系列', '中文语义理解优秀', '中文语境复杂语义处理'],
    ['Kimi-K2 系列', '超长上下文（256K）', '超长文档阅读、多文档分析'],
    ['MiniMax-M2 系列', '多媒体内容生成', '内容创作、角色扮演'],
  ];
  slide.addTable(modelRows, { x: 0.6, y: 1.5, w: 8.8, colW: [2, 3.4, 3.4], fontSize: 10, border: { pt: 0.5, color: 'DDDDDD' }, autoPage: false });

  // ===== Slide 5: Enterprise Flywheel =====
  slide = pptx.addSlide();
  slide.background = { color: BLUE };
  slide.addText('企业版核心逻辑：AI 工作飞轮', { x: 0.6, y: 0.4, w: 8, h: 0.5, fontSize: 10, color: 'FFFFFF' });
  slide.addText('企业版核心逻辑', { x: 0.6, y: 0.7, w: 8, h: 0.7, fontSize: 28, bold: true, color: 'FFFFFF' });
  slide.addText('企业版不是给员工多一个 AI 聊天框。它跑的是一条飞轮：统一入口 → 超级个体 → 超级团队 → 企业 AI 资产 → 企业数字员工 → 组织生产力持续释放', { x: 0.6, y: 1.7, w: 8.8, h: 0.8, fontSize: 13, color: 'FFFFFF', italic: true, lineSpacingMultiple: 1.4 });

  const flywheelItems = [
    ['超级个体', '员工用 WorkBuddy 调用组织能力干活'],
    ['超级团队', '人 + 数字员工组队干活'],
    ['企业 AI 资产', '每次任务沉淀知识、技能、经验和案例'],
    ['企业数字员工', '7×24 在线，懂岗位、有技能、熟悉企业知识'],
    ['统一治理', '管理员统一管控权限、安全和审计'],
  ];
  flywheelItems.forEach((item, i) => {
    const y = 2.8 + i * 0.7;
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 8.8, h: 0.55, fill: { color: 'FFFFFF', transparency: 15 }, rectRadius: 0.08 });
    slide.addText(item[0], { x: 0.8, y, w: 2.5, h: 0.55, fontSize: 12, bold: true, color: GREEN, valign: 'middle' });
    slide.addText(item[1], { x: 3.3, y, w: 5.8, h: 0.55, fontSize: 11, color: 'FFFFFF', valign: 'middle' });
  });

  // Agent Suite
  slide.addText('Agent Suite 企业套件', { x: 0.6, y: 6.3, w: 8.8, h: 0.3, fontSize: 10, bold: true, color: GREEN });
  slide.addText('腾讯文档（内容生产与协作底座） | 腾讯网盘（文件管理底座） | 腾讯乐享（知识治理底座）', { x: 0.6, y: 6.6, w: 8.8, h: 0.3, fontSize: 9, color: 'FFFFFF' });

  // ===== Slide 6: Office Scenario =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('01 / OFFICE', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
  slide.addText('场景一：高效信息沟通', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });
  slide.addText('周一早上 9 点，员工打开邮箱发现 38 封未读邮件，包含领导催进度、客户报价、HR 表格、项目协作、会议通知等不同类型事项。', { x: 0.6, y: 1.3, w: 8.8, h: 0.6, fontSize: 11, color: GRAY, lineSpacingMultiple: 1.3 });

  const steps = [
    '读取未读邮件汇总文件，识别每封邮件的业务含义和处理责任',
    '将 38 封邮件分为"必须亲自回/助理代回/仅知会/可忽略"4 类',
    '为"必须亲自回"的邮件生成专业中文回复草稿',
    '处理重点英文客户邮件：先生成中文审稿，再翻译成商务英语版',
    '员工确认后可继续完成发送或回写',
  ];
  steps.forEach((step, i) => {
    slide.addText(`${i + 1}. ${step}`, { x: 0.6, y: 2.1 + i * 0.5, w: 8.8, h: 0.4, fontSize: 10.5, color: DARK });
  });
  slide.addText('传统方式 2-3 小时 → WorkBuddy 10-20 分钟，员工从执行者变成审核者', { x: 0.6, y: 4.8, w: 8.8, h: 0.4, fontSize: 12, bold: true, color: GREEN });

  // ===== Slide 7: PM Scenario =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('02 / PM', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
  slide.addText('场景二：PM 项目过程管理', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });

  const pmRows = [
    [{ text: '场景', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '典型输入', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '标准输出', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '直接价值', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }],
    ['会议纪要闭环', '会议记录/聊天记录', '结论/行动项/责任人/DDL/风险项', '会后闭环速度提升'],
    ['周报/月报生成', '本周事项/风险/里程碑', '管理层口径周报草稿', '写作时间下降'],
    ['验收材料打包', '验收报告/清单/截图', '可提交材料包+检查清单', '减少漏项'],
    ['项目知识检索', '历史PRD/纪要/周报/合同', '关键结论与出处', '减少"找资料"耗时'],
    ['远程触发任务', '微信/企微/QQ 指令', '周报草稿/资料归档结果', '碎片时间推进项目'],
  ];
  slide.addTable(pmRows, { x: 0.6, y: 1.3, w: 8.8, colW: [2, 2.3, 2.5, 2], fontSize: 9.5, border: { pt: 0.5, color: 'DDDDDD' }, autoPage: false });

  // ===== Slide 8: Enterprise Scenarios (non-乐享) =====
  const entScenes = [
    {
      icon: '会议', title: '会议到任务闭环',
      pain: '会议纪要有了，但行动项没人跟，责任人和后续进展容易断，会后闭环全靠人追',
      steps: '① 会议记录/聊天记录/语音转写进入 WorkBuddy ② 数字员工提炼结论/行动项/责任人/DDL/风险项 ③ Teams 把行动项分配给成员和数字员工 ④ 自动化任务定期提醒、跟踪和汇总进展 ⑤ 复盘和决策沉淀为团队资产 ⑥ 下一次会议可直接复用组织记忆',
      value: '会议不再止于纪要 | 行动项持续推进 | 任务状态可见 | 团队协同从沟通走向执行'
    },
    {
      icon: '管理', title: '管理层周报 / 例会 / 经营复盘',
      pain: '信息散落在会议纪要/项目材料/沟通记录里，每周花大量时间收集整理，周报和复盘高度依赖手工搬运',
      steps: '① WorkBuddy 自动汇总项目纪要/制度更新/风险记录/业务分析/日报周报素材 ② 管理者通过 WorkBuddy 发起任务 ③ WorkBuddy 生成周报/例会提纲/跟进清单/经营摘要/风险项列表 ④ 接入企微后支持结果推送和远程办公',
      value: '节省管理层整理时间 | 强化经营信息统一口径 | 提高例会和复盘效率 | 知识沉淀进入经营管理闭环'
    },
  ];

  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('04 / ENTERPRISE', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
  slide.addText('企业版场景实践', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });

  entScenes.forEach((scene, i) => {
    const x = 0.4 + i * 4.8;
    const y = 1.3;
    // Card
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 4.6, h: 5.5, fill: { color: 'F8FAFF' }, line: { color: 'E0E8F5', width: 0.5 }, rectRadius: 0.1 });
    // Badge
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 0.2, w: 0.8, h: 0.35, fill: { color: BLUE }, rectRadius: 0.05 });
    slide.addText(scene.icon, { x: x + 0.2, y: y + 0.2, w: 0.8, h: 0.35, fontSize: 9, color: 'FFFFFF', align: 'center', valign: 'middle' });
    // Title
    slide.addText(scene.title, { x: x + 1.15, y: y + 0.15, w: 3.3, h: 0.4, fontSize: 13, bold: true, color: DARK });
    // Pain
    slide.addText('痛点：' + scene.pain, { x: x + 0.2, y: y + 0.7, w: 4.2, h: 0.6, fontSize: 9.5, color: GRAY, lineSpacingMultiple: 1.3 });
    // Steps label
    slide.addText('实践路径：', { x: x + 0.2, y: y + 1.4, w: 4.2, h: 0.25, fontSize: 9.5, bold: true, color: DARK });
    // Steps
    slide.addText(scene.steps, { x: x + 0.2, y: y + 1.65, w: 4.2, h: 2.8, fontSize: 9, color: BLUE, lineSpacingMultiple: 1.2 });
    // Value
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: y + 4.9, w: 4.2, h: 0.4, fill: { color: GREEN_BG }, rectRadius: 0.05 });
    slide.addText('客户价值：' + scene.value, { x: x + 0.25, y: y + 4.9, w: 4.1, h: 0.4, fontSize: 8.5, bold: true, color: '006644', valign: 'middle' });
  });

  // ===== Slide 8: 全岗位高频场景 =====
  const roleScenes = [
    { icon: 'HR', title: 'HR · 简历筛选与招聘提效', pain: '逐份打开简历手动比对JD、填表打分干到眼花；招聘文案从头写，入职通知一份份发', steps: '① 批量读取候选人简历，按匹配度自动打分排出面试优先级 ② 生成《候选人匹配度评分报告》含排名/核心亮点/关键提醒 ③ 自动生成岗位招聘文案 ④ 5分钟批量发送100份个性化入职通知', value: '简历筛选半天→15分钟 | 招聘文案一键生成 | 入职通知批量发送' },
    { icon: '财务', title: '财务 · 发票报销与报表汇总', pain: '月底一张张拍照录入发票，金额/日期/项目名错一个字整张表重来；多张Excel报表汇总靠手动', steps: '① 批量导入发票文件，自动提取金额/日期/项目名称/代码 ② 按"客户名字+开具时间+总金额"命名归档 ③ 45张发票5分钟生成报销明细表 ④ 多个Excel报表自动汇总生成可视化图表', value: '发票录入2小时→5分钟 | 报销明细可直接提交 | 报表自动汇总' },
    { icon: '行政', title: '行政 · 会议纪要与文件归档', pain: '会议录音转写后还要手动提炼决议和行动项；文件命名混乱、分类归档靠手工', steps: '① 会议录音/零散笔记进入WorkBuddy，自动提炼关键决策/行动项/责任人/截止时间 ② 生成标准化会议纪要按部门分类 ③ 办公文件自动归档，按类别分类统一重命名 ④ 批量生成各类通知文件', value: '纪要整理30分钟→3分钟 | 文件归档自动化 | 通知格式统一规范' },
    { icon: '运营', title: '运营/市场 · 内容与竞品分析', pain: '搜资料1小时、组织语言1小时、排版半小时，交稿还被说"不够网感"；竞品分析靠手动', steps: '① 输入主题自动搜集最新资料，15分钟出推文初稿 ② 多平台文案一键生成（公众号/小红书/短视频脚本） ③ 分析竞品推广策略整理成复盘报告 ④ 公众号周报自动化：每周一自动拉取阅读数据分析', value: '15分钟3篇小红书图文 | 公众号效率提升300%+ | 竞品分析半天→30分钟' },
    { icon: '法务', title: '法务 · 合同审查与风险排查', pain: '老板给了10份合同逐一审查风险，每次只能上传1份；合同条款多，关键风险点容易遗漏', steps: '① 批量上传合同文件，给一条统一指令系统自动逐个处理 ② 自动识别合同中风险条款/异常约定/合规问题 ③ 生成风险审查报告标注关键提醒和修改建议 ④ 合同版本对比自动标注差异条款', value: '10份合同审查1天→2小时 | 风险点自动标注 | 版本差异一目了然' },
    { icon: '教育', title: '教育 · 出题/成绩/课件', pain: '出阅读理解题太费时间，网上题目千篇一律；考完试成绩分析要用Excel算半天', steps: '① 根据课文自动生成阅读理解题（选择+简答+参考答案）导出Word直接打印 ② 成绩表自动分析：各科平均分/及格率/优秀率+柱状图 ③ Word文档一键生成课件PPT ④ 复旦大学200+师生现场体验"AI百校行"', value: '出题1小时→5分钟 | 成绩分析30分钟→30秒 | 课件1小时→20分钟' },
  ];

  for (let page = 0; page < 2; page++) {
    slide = pptx.addSlide();
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
    slide.addText('03 / ROLE SCENES', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
    slide.addText(page === 0 ? '全岗位高频场景' : '续：全岗位高频场景', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });
    if (page === 0) {
      slide.addText('来自 WorkBuddy 官方公众号和腾讯云开发者社区的真实用户场景', { x: 0.6, y: 1.1, w: 8.8, h: 0.3, fontSize: 10, color: GRAY, italic: true });
    }

    const scenes = roleScenes.slice(page * 3, page * 3 + 3);
    scenes.forEach((scene, i) => {
      const x = 0.25 + i * 3.2;
      const y = page === 0 ? 1.5 : 1.2;

      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 3.05, h: 5.5, fill: { color: 'F8FAFF' }, line: { color: 'E0E8F5', width: 0.5 }, rectRadius: 0.08 });
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 0.1, w: 0.7, h: 0.3, fill: { color: BLUE }, rectRadius: 0.04 });
      slide.addText(scene.icon, { x: x + 0.1, y: y + 0.1, w: 0.7, h: 0.3, fontSize: 8, color: 'FFFFFF', align: 'center', valign: 'middle' });
      slide.addText(scene.title, { x: x + 0.9, y: y + 0.1, w: 2.0, h: 0.3, fontSize: 10, bold: true, color: DARK });
      slide.addText('痛点：' + scene.pain, { x: x + 0.1, y: y + 0.5, w: 2.85, h: 0.7, fontSize: 8.5, color: GRAY, lineSpacingMultiple: 1.2 });
      slide.addText('做法：', { x: x + 0.1, y: y + 1.3, w: 2.85, h: 0.2, fontSize: 8.5, bold: true, color: DARK });
      slide.addText(scene.steps, { x: x + 0.1, y: y + 1.5, w: 2.85, h: 2.8, fontSize: 8, color: BLUE, lineSpacingMultiple: 1.15 });
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 4.6, w: 2.85, h: 0.7, fill: { color: GREEN_BG }, rectRadius: 0.04 });
      slide.addText(scene.value, { x: x + 0.15, y: y + 4.65, w: 2.75, h: 0.6, fontSize: 8, color: '006644', bold: true, lineSpacingMultiple: 1.2 });
    });
  }

  // ===== Slides 11-12: 企业版场景实践 =====
  const lexiangScenes = [
    {
      icon: '售后', title: '售后 / 客服问题处理',
      lexiang: '沉淀 FAQ/故障案例/操作手册/话术模板，排查记录和新案例沉淀回知识库',
      workbuddy: '调用乐享知识 → 整理原因/排查顺序/回复草稿 → 沉淀回知识库',
      steps: '① 售后同学打开 WorkBuddy 处理客户问题 ② 调用乐享知识库 FAQ/故障案例/操作手册/话术模板 ③ 售后数字员工整理可能原因、排查顺序和回复草稿 ④ 需要升级时加入 Teams 协同 ⑤ 排查记录和新案例沉淀回知识库 ⑥ 沉淀的案例和 SOP 再赋能下一次售后数字员工',
      value: '客户回复更快 | 排查路径更标准 | 新人更容易上手 | 售后经验从个人能力变组织能力'
    },
    {
      icon: '销售', title: '销售 / 售前材料快速响应',
      lexiang: '统一沉淀产品资料/行业方案/成功案例/招标模板/异议回复/报价规则',
      workbuddy: '调用乐享素材 → 生成方案初稿/沟通邮件/拜访提纲 → 定稿保存回腾讯文档',
      steps: '① 乐享统一沉淀产品资料/行业方案/成功案例/招标模板/异议回复/报价规则 ② 销售在企微或桌面侧向 WorkBuddy 发起任务 ③ 输出客户方案初稿/沟通邮件/拜访提纲/案例摘要/纪要与跟进清单 ④ Teams 协同修改审阅 ⑤ 定稿保存腾讯文档，沉淀到销售弹药库 ⑥ 乐享"找准内容"，WorkBuddy"变成输出"',
      value: '客户响应速度提升 | 销售口径一致性提高 | 方案制作时间缩短 | 销售输出建立在企业知识基础上'
    },
    {
      icon: '知识', title: '企业知识运营',
      lexiang: '确保知识可信、可控、可追溯，权限和有效期控制',
      workbuddy: 'Agent Suite 调用乐享 → 在受控知识基础上完成问答/总结/分析/材料生成',
      steps: '① 员工在 WorkBuddy 中直接询问制度/流程/产品资料/项目经验 ② Agent Suite 调用腾讯文档/网盘/乐享 ③ 乐享确保知识可信、可控、可追溯 ④ 数字员工完成问答/总结/分析/材料生成 ⑤ 新问题和知识缺口沉淀为企业 AI 资产 ⑥ 资产继续赋能数字员工',
      value: '知识找得到 | 回答答得准 | 权限管得住 | 结果可追溯 | 企业知识越用越好'
    },
    {
      icon: '培训', title: '新人培训与岗位辅导',
      lexiang: '沉淀岗位 SOP/流程规范/典型案例/常见错误/模板口径/培训资料与题库',
      workbuddy: '新人通过知识库快速找答案 → 生成学习清单/学习计划/错题总结/岗位周报',
      steps: '① 乐享沉淀岗位 SOP/流程规范/典型案例/常见错误/模板口径/培训资料与题库 ② 新人通过知识库快速找到答案 ③ WorkBuddy 生成学习清单/学习计划/错题总结/岗位周报/导师跟进建议 ④ 把培训从静态课件变成动态辅导 ⑤ 把组织经验留在系统里',
      value: '缩短新人上岗周期 | 降低培训和带教成本 | 减少基础问题反复答疑 | 培训从静态课件变动态度辅导'
    },
    {
      icon: '远程', title: '移动办公 / 出差办公助手',
      lexiang: '企业知识统一沉淀在乐享中，手机端即可调用',
      workbuddy: '人在外部通过企微下发任务 → 调用知识并在本地电脑完成操作 → 结果传回企微',
      steps: '① 企业知识统一沉淀在乐享中 ② 人在外部通过企微向 WorkBuddy 下发任务 ③ WorkBuddy 调用知识并在本地电脑完成操作，结果传回企微 ④ 示例：企微发起"把客户X的最新案例发给我，生成跟进邮件" → 调用乐享资料 → WorkBuddy 整理邮件处理文件 → 结果推回企微',
      value: '移动办公效率提升 | 减少跨设备跨人协调成本 | "找资料-处理-回传"一次做完'
    },
    {
      icon: '合规', title: '制度 / 合规 / 专业支持',
      lexiang: '沉淀制度规范/合规要求/审批规则/专业模板/历史案例/审核意见',
      workbuddy: '在受控知识基础上生成制度答复草稿/审核清单/材料整理初稿/执行提醒',
      steps: '① 乐享沉淀制度规范/合规要求/审批规则/专业模板/历史案例/审核意见/业务指引 ② 通过权限和有效期控制保证不同角色看到正确、有效、可授权的内容 ③ WorkBuddy 在受控知识基础上生成制度答复草稿/审核清单/材料整理初稿/内部通知/执行提醒 ④ AI 使用建立在合规治理基础上',
      value: '减少错用制度和版本错误 | 提高专业岗位规范性 | 降低业务风险 | AI 建立在合规治理基础上'
    },
  ];

  // 2 slides, 3 scenarios each
  for (let page = 0; page < 2; page++) {
    slide = pptx.addSlide();
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
    slide.addText('05 / LEXIANG + WORKBUDDY', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
    slide.addText('乐享知识库联动场景', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });
    // Subtitle explaining the synergy
    if (page === 0) {
      slide.addText('乐享负责知识沉淀与权限治理 → WorkBuddy 负责智能调用与任务执行 → 结果可交付、可追溯', { x: 0.6, y: 1.1, w: 8.8, h: 0.3, fontSize: 10, color: GRAY, italic: true });
    } else {
      slide.addText('续：乐享知识库联动场景', { x: 0.6, y: 1.1, w: 8.8, h: 0.3, fontSize: 10, color: GRAY, italic: true });
    }

    const scenes = lexiangScenes.slice(page * 3, page * 3 + 3);
    scenes.forEach((scene, i) => {
      const x = 0.25 + i * 3.2;
      const y = 1.5;

      // Card background
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 3.05, h: 5.5, fill: { color: 'F8FAFF' }, line: { color: 'E0E8F5', width: 0.5 }, rectRadius: 0.08 });

      // Icon badge
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 0.1, w: 0.7, h: 0.3, fill: { color: BLUE }, rectRadius: 0.04 });
      slide.addText(scene.icon, { x: x + 0.1, y: y + 0.1, w: 0.7, h: 0.3, fontSize: 8, color: 'FFFFFF', align: 'center', valign: 'middle' });

      // Title
      slide.addText(scene.title, { x: x + 0.9, y: y + 0.05, w: 2, h: 0.35, fontSize: 10, bold: true, color: DARK });

      // 乐享负责 - gold tag
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 0.5, w: 2.85, h: 0.55, fill: { color: LEXIANG_GOLD }, rectRadius: 0.04 });
      slide.addText('乐享：' + scene.lexiang, { x: x + 0.15, y: y + 0.5, w: 2.75, h: 0.55, fontSize: 7, color: LEXIANG_TEXT, valign: 'middle' });

      // WorkBuddy负责 - blue tag
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 1.15, w: 2.85, h: 0.55, fill: { color: LIGHT_BLUE }, rectRadius: 0.04 });
      slide.addText('WorkBuddy：' + scene.workbuddy, { x: x + 0.15, y: y + 1.15, w: 2.75, h: 0.55, fontSize: 7, color: BLUE, valign: 'middle' });

      // Steps
      slide.addText('实践路径：', { x: x + 0.1, y: y + 1.8, w: 2.85, h: 0.2, fontSize: 7, bold: true, color: DARK });
      slide.addText(scene.steps, { x: x + 0.1, y: y + 2.0, w: 2.85, h: 2.6, fontSize: 6.5, color: BLUE, lineSpacingMultiple: 1.1 });

      // Customer value
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: y + 5.0, w: 2.85, h: 0.35, fill: { color: GREEN_BG }, rectRadius: 0.04 });
      slide.addText('价值：' + scene.value, { x: x + 0.13, y: y + 5.0, w: 2.8, h: 0.35, fontSize: 6, bold: true, color: '006644', valign: 'middle' });
    });
  }

  // ===== Slide 11: Use Cases =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('06 / USE CASES', { x: 0.6, y: 0.3, w: 5, h: 0.3, fontSize: 10, color: BLUE });
  slide.addText('使用案例（WorkBuddy 专属）', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });

  const cases = [
    { num: '01', title: '构建个人学术信息助手', desc: '从0到1搭建 arXiv 论文每日自动筛选、精排、深度解读、报告生成全流程自动化。三路并行召回，召回率97.4%，每日自动生成10篇精读+20篇速览', metric: '搭建~2天 | 代码~5000行 | 召回率97.4% | API成本¥0' },
    { num: '02', title: 'AI 简历筛选提效 90%', desc: '新加坡团队用 WorkBuddy 做简历筛选，自动解析简历、按岗位逐项匹配评分、生成筛选结果和推荐理由，人工筛选时间从几小时压到几分钟', metric: '提效90% | 人工筛选几小时→几分钟' },
    { num: '03', title: '五粮液全民AI办公提效', desc: '基于乐享、企微、WorkBuddy 给五粮液做全民AI办公提效动员方案，覆盖白酒行业典型场景：知识检索、会议纪要闭环、周报自动生成、质检提效', metric: '白酒行业首个客户落地 | 知识管理→智能办公全链路' },
    { num: '04', title: '冠盛客户场景沟通实践', desc: '和冠盛团队做场景沟通，围绕售后问题处理、销售材料快速响应、新人培训、管理层周报复盘等场景逐一验证，确认 WorkBuddy+乐享可覆盖制造业核心痛点', metric: '制造业场景验证 | 售后/销售/培训/管理逐一确认' },
    { num: '05', title: 'WorkBuddy+乐享联合方案落地', desc: '乐享负责知识沉淀/权限治理/搜索问答，WorkBuddy 负责桌面AI智能体/本地执行/企微接入/结果回传，形成"知识获取+任务执行+结果交付"完整链路，覆盖6大场景', metric: '6大场景方案 | 知识获取+任务执行+结果交付' },
    { num: '06', title: 'PM项目管理全场景提效', desc: '5个高频场景（纪要闭环/周报月报/验收打包/知识检索/远程触发）标准化实操。PM从"从零写报告"转为"审核和定稿"，从"找资料靠记忆"变成"按出处检索"', metric: '5大PM高频场景 | PM角色从执行者→审核者' },
  ];

  cases.forEach((c, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.15;
    const y = 1.3 + row * 2.7;

    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 3, h: 2.5, fill: { color: 'FAFBFF' }, line: { color: 'E0E8F5', width: 0.5 }, rectRadius: 0.1 });
    slide.addText(c.num, { x: x + 0.15, y: y + 0.1, w: 1, h: 0.4, fontSize: 22, bold: true, color: BLUE });
    slide.addText(c.title, { x: x + 0.15, y: y + 0.5, w: 2.7, h: 0.35, fontSize: 11, bold: true, color: DARK });
    slide.addText(c.desc, { x: x + 0.15, y: y + 0.9, w: 2.7, h: 1.1, fontSize: 8, color: GRAY, lineSpacingMultiple: 1.2 });
    slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.15, y: y + 2.05, w: 2.7, h: 0.35, fill: { color: 'F0F5FF' }, rectRadius: 0.05 });
    slide.addText(c.metric, { x: x + 0.2, y: y + 2.05, w: 2.6, h: 0.35, fontSize: 8, bold: true, color: BLUE, valign: 'middle' });
  });

  // ===== Slide 12: Security =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('数据安全与合规', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });

  const secItems = [
    ['1', '本地执行', '所有文件处理在本地完成，原始文件从不离开用户电脑'],
    ['2', '权限隔离', '沙箱与容器隔离运行，网关防护，传输加密，高危操作拦下来'],
    ['3', '审计留痕', '感知可信过滤、提示词攻击防护、推理过程可视化，每个动作都留痕'],
  ];
  secItems.forEach((item, i) => {
    const y = 1.3 + i * 0.9;
    slide.addShape(pptx.shapes.OVAL, { x: 0.6, y, w: 0.45, h: 0.45, fill: { color: BLUE } });
    slide.addText(item[0], { x: 0.6, y, w: 0.45, h: 0.45, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    slide.addText(item[1], { x: 1.2, y, w: 2, h: 0.45, fontSize: 13, bold: true, color: DARK, valign: 'middle' });
    slide.addText(item[2], { x: 3.2, y, w: 6, h: 0.45, fontSize: 10.5, color: GRAY, valign: 'middle' });
  });

  slide.addText('合规认证', { x: 0.6, y: 4.2, w: 5, h: 0.3, fontSize: 13, bold: true, color: DARK });
  const compItems = [
    '国内首批通过信通院 Claw 可信能力评估（57个具体能力项）',
    '信通院智能原生编码 IDE 评估 4+ 级（177个能力子项测评表现优异）',
    '同源工具 CodeBuddy 自举验证（新增代码 100% 由 AI 自举生成）',
  ];
  compItems.forEach((item, i) => {
    slide.addText('• ' + item, { x: 0.8, y: 4.6 + i * 0.45, w: 8.4, h: 0.35, fontSize: 10.5, color: GRAY });
  });

  // ===== Slide 13: CodeBuddy =====
  slide = pptx.addSlide();
  slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: BLUE } });
  slide.addText('CodeBuddy 同源架构', { x: 0.6, y: 0.6, w: 5, h: 0.5, fontSize: 22, bold: true, color: DARK });
  slide.addText('WorkBuddy 基于腾讯云代码助手 CodeBuddy 同源架构，CodeBuddy 有三种形态覆盖不同用户：', { x: 0.6, y: 1.3, w: 8.8, h: 0.4, fontSize: 11, color: GRAY });

  const cbRows = [
    [{ text: '产品形态', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '适用用户', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }, { text: '核心特点', options: { bold: true, color: 'FFFFFF', fill: { color: BLUE } } }],
    ['CodeBuddy IDE', '产品/设计师/全栈/初学者', '产设研一体化工作台，"对话即编程"'],
    ['CodeBuddy 插件', '日常编码开发者', '即插即用、融入现有工作流'],
    ['CodeBuddy Code', 'DevOps/运维/SRE/资深开发', '命令行工具，任务编排能力强'],
  ];
  slide.addTable(cbRows, { x: 0.6, y: 2, w: 8.8, colW: [2.5, 3.3, 3], fontSize: 11, border: { pt: 0.5, color: 'DDDDDD' }, autoPage: false });

  // ===== Slide 14: End =====
  slide = pptx.addSlide();
  slide.background = { color: BLUE };
  slide.addText('WorkBuddy', { x: 1, y: 2, w: 8, h: 1, fontSize: 42, bold: true, color: 'FFFFFF', align: 'center' });
  slide.addText('一句话让 AI 替你上班', { x: 1, y: 3.2, w: 8, h: 0.6, fontSize: 22, color: GREEN, align: 'center', bold: true });
  slide.addText('腾讯 CSIG · 2026', { x: 1, y: 5, w: 8, h: 0.3, fontSize: 12, color: 'FFFFFF', align: 'center' });

  // Save
  const outputPath = '/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/WorkBuddy产品介绍.pptx';
  await pptx.writeFile({ fileName: outputPath });
  console.log('PPT created successfully!');
}

createPresentation().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});

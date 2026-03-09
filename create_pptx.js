const PptxGenJS = require('pptxgenjs');
const path = require('path');

// 创建演示文稿
let pptx = new PptxGenJS();

// 设置专业配色方案（Ocean Gradient 主题）
const colors = {
  primary: '065A82',      // 深蓝色
  secondary: '1C7293',    // 青色
  tertiary: '2E86AB',     // 中蓝色
  white: 'FFFFFF',
  light: 'F0F8FF',
  dark: '1A1A2E',
  accent: '00D4FF'       // 亮青色强调
};

// ==========================================
// 幻灯片 1: 封面
// ==========================================
let slide1 = pptx.addSlide();
slide1.background = { color: colors.primary };

// 标题
slide1.addText('广播新闻稿与新闻特写\n写作指南', {
  x: 0.5, y: 1.5, w: 9, h: 1.5,
  fontSize: 44,
  bold: true,
  color: colors.white,
  align: 'center',
  fontFace: 'Arial'
});

// 装饰线
slide1.addShape(pptx.ShapeType.line, {
  x: 3, y: 2.8, w: 4, h: 0,
  line: { color: colors.accent, width: 3 }
});

// 副标题
slide1.addText('Broadcast News & Feature Writing Guide', {
  x: 0.5, y: 3.3, w: 9, h: 0.5,
  fontSize: 18,
  color: 'B8E4F0',
  align: 'center',
  fontFace: 'Calibri'
});

// 底部装饰
slide1.addShape(pptx.ShapeType.rect, {
  x: 0, y: 6.5, w: 10, h: 0.5,
  fill: { color: colors.accent }
});

// ==========================================
// 幻灯片 2: 广播新闻稿范文
// ==========================================
let slide2 = pptx.addSlide();
slide2.background = { color: colors.light };

// 标题栏
slide2.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide2.addText('第一部分：广播新闻稿范文', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 新闻标题
slide2.addText('联合国官员和专家对新都中学\n人口教育工作十分赞赏', {
  x: 0.5, y: 1.5, w: 9, h: 1,
  fontSize: 24,
  bold: true,
  color: colors.primary,
  fontFace: 'Arial'
});

// 作者信息
slide2.addText('作者：邱晋南 陈四益 陈至崇', {
  x: 0.5, y: 2.4, w: 9, h: 0.4,
  fontSize: 14,
  color: colors.secondary,
  italic: true,
  fontFace: 'Calibri'
});

// 主要内容 - 使用带图标的列表
const content1 = [
  { icon: '📚', text: '新都中学人口教育五年多来的成果' },
  { icon: '👨‍🎓', text: '1200多名高中学生接受了人口科学教育' },
  { icon: '📄', text: '学生写的5份专论被提交给联合国教科文组织' }
];

let yPos = 3.2;
content1.forEach((item, index) => {
  // 图标圆圈
  slide2.addShape(pptx.ShapeType.ellipse, {
    x: 0.7, y: yPos - 0.1, w: 0.5, h: 0.5,
    fill: { color: colors.secondary }
  });
  slide2.addText(item.icon, {
    x: 0.85, y: yPos, w: 0.3, h: 0.3,
    fontSize: 20,
    align: 'center'
  });

  // 文字内容
  slide2.addText(item.text, {
    x: 1.4, y: yPos, w: 7.5, h: 0.5,
    fontSize: 18,
    color: colors.dark,
    bullet: false,
    fontFace: 'Calibri'
  });

  yPos += 0.6;
});

// ==========================================
// 幻灯片 3: 联合国评价
// ==========================================
let slide3 = pptx.addSlide();
slide3.background = { color: colors.light };

// 标题栏
slide3.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide3.addText('第二部分：联合国评价', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 三个评价卡片
const evaluations = [
  {
    name: '沙尔马博士',
    title: '联合国教科文组织人口教育顾问',
    quote: '"我对你们卓有成效的人口教育表示十分赞赏"'
  },
  {
    name: '海迪·斯温德尔斯女士',
    title: '联合国人口活动基金审评小组成员',
    quote: '"我要把你们写的心得带回纽约，让我的女儿拿到学校去念"'
  },
  {
    name: '联合国审评报告',
    title: '官方评价',
    quote: '"政府官员、教员和学生的高度主动性是非凡的"'
  }
];

evaluations.forEach((eval, index) => {
  const yStart = 1.8 + index * 1.6;

  // 卡片背景
  slide3.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: yStart, w: 9, h: 1.4,
    fill: { color: colors.white },
    line: { color: colors.secondary, width: 2 }
  });

  // 装饰色块
  slide3.addShape(pptx.ShapeType.rect, {
    x: 0.5, y: yStart, w: 0.3, h: 1.4,
    fill: { color: colors.secondary }
  });

  // 姓名
  slide3.addText(eval.name, {
    x: 1, y: yStart + 0.1, w: 4, h: 0.3,
    fontSize: 16,
    bold: true,
    color: colors.primary,
    fontFace: 'Arial'
  });

  // 职位
  slide3.addText(eval.title, {
    x: 1, y: yStart + 0.35, w: 4, h: 0.3,
    fontSize: 12,
    color: colors.secondary,
    italic: true,
    fontFace: 'Calibri'
  });

  // 引用
  slide3.addText(eval.quote, {
    x: 1, y: yStart + 0.75, w: 8, h: 0.5,
    fontSize: 13,
    color: colors.dark,
    fontFace: 'Calibri',
    align: 'left'
  });
});

// ==========================================
// 幻灯片 4: 如何写广播新闻稿 (1-2)
// ==========================================
let slide4 = pptx.addSlide();
slide4.background = { color: colors.light };

// 标题栏
slide4.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide4.addText('第三部分：如何写广播新闻稿（要点 1-2）', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 要点1：通俗化、口语化
slide4.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.4, w: 9, h: 0.6,
  fill: { color: colors.secondary }
});
slide4.addText('要点一：通俗化、口语化', {
  x: 0.7, y: 1.5, w: 8.6, h: 0.4,
  fontSize: 20,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

const point1Details = [
  '✓ 用词要普通',
  '✓ 句子要短',
  '✓ 多用人称名词',
  '✓ 避免同音歧解和同意反复'
];

let yDetail1 = 2.2;
point1Details.forEach(detail => {
  slide4.addText(detail, {
    x: 1, y: yDetail1, w: 8, h: 0.35,
    fontSize: 14,
    color: colors.dark,
    fontFace: 'Calibri'
  });
  yDetail1 += 0.4;
});

// 分隔线
slide4.addShape(pptx.ShapeType.line, {
  x: 1, y: 4, w: 8, h: 0,
  line: { color: colors.accent, width: 2, dash: 'dash' }
});

// 要点2：采写要及时
slide4.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 4.2, w: 9, h: 0.6,
  fill: { color: colors.tertiary }
});
slide4.addText('要点二：采写要及时', {
  x: 0.7, y: 4.3, w: 8.6, h: 0.4,
  fontSize: 20,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

const point2Details = [
  '✓ 发挥快捷特性',
  '✓ 只抓重点，少作深度挖掘'
];

let yDetail2 = 5;
point2Details.forEach(detail => {
  slide4.addText(detail, {
    x: 1, y: yDetail2, w: 8, h: 0.35,
    fontSize: 14,
    color: colors.dark,
    fontFace: 'Calibri'
  });
  yDetail2 += 0.4;
});

// ==========================================
// 幻灯片 5: 如何写广播新闻稿 (3-4)
// ==========================================
let slide5 = pptx.addSlide();
slide5.background = { color: colors.light };

// 标题栏
slide5.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide5.addText('第三部分：如何写广播新闻稿（要点 3-4）', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 要点3：导语技巧
slide5.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 1.4, w: 9, h: 0.6,
  fill: { color: colors.secondary }
});
slide5.addText('要点三：导语技巧', {
  x: 0.7, y: 1.5, w: 8.6, h: 0.4,
  fontSize: 20,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

const point3Details = [
  '✓ 先打招呼，请听众收听',
  '✓ 增强吸引力'
];

let yDetail3 = 2.2;
point3Details.forEach(detail => {
  slide5.addText(detail, {
    x: 1, y: yDetail3, w: 8, h: 0.35,
    fontSize: 14,
    color: colors.dark,
    fontFace: 'Calibri'
  });
  yDetail3 += 0.4;
});

// 分隔线
slide5.addShape(pptx.ShapeType.line, {
  x: 1, y: 4, w: 8, h: 0,
  line: { color: colors.accent, width: 2, dash: 'dash' }
});

// 要点4：音响特点
slide5.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 4.2, w: 9, h: 0.6,
  fill: { color: colors.tertiary }
});
slide5.addText('要点四：音响特点', {
  x: 0.7, y: 4.3, w: 8.6, h: 0.4,
  fontSize: 20,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

const point4Details = [
  '✓ 突出广播的音响特点',
  '✓ 录音讲话、录音新闻、录音通讯等'
];

let yDetail4 = 5;
point4Details.forEach(detail => {
  slide5.addText(detail, {
    x: 1, y: yDetail4, w: 8, h: 0.35,
    fontSize: 14,
    color: colors.dark,
    fontFace: 'Calibri'
  });
  yDetail4 += 0.4;
});

// ==========================================
// 幻灯片 6: 倒三角结构
// ==========================================
let slide6 = pptx.addSlide();
slide6.background = { color: colors.light };

// 标题栏
slide6.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide6.addText('第四部分：倒三角结构', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 倒三角图示
slide6.addShape(pptx.ShapeType.rect, {
  x: 3, y: 1.5, w: 4, h: 1.2,
  fill: { color: colors.primary },
  line: { color: colors.secondary, width: 2 }
});
slide6.addText('标题\n说明新闻卖点', {
  x: 3.5, y: 1.8, w: 3, h: 0.6,
  fontSize: 16,
  bold: true,
  color: colors.white,
  align: 'center',
  fontFace: 'Arial'
});

slide6.addShape(pptx.ShapeType.rect, {
  x: 3.5, y: 2.8, w: 3, h: 1.2,
  fill: { color: colors.secondary },
  line: { color: colors.primary, width: 2 }
});
slide6.addText('第一段\n时间、地点、人物\n起因、经过、结果', {
  x: 3.8, y: 3.1, w: 2.4, h: 0.6,
  fontSize: 14,
  bold: true,
  color: colors.white,
  align: 'center',
  fontFace: 'Arial'
});

slide6.addShape(pptx.ShapeType.rect, {
  x: 4, y: 4.1, w: 2, h: 1.2,
  fill: { color: colors.tertiary },
  line: { color: colors.secondary, width: 2 }
});
slide6.addText('补充\n细节', {
  x: 4.3, y: 4.4, w: 1.4, h: 0.6,
  fontSize: 14,
  bold: true,
  color: colors.white,
  align: 'center',
  fontFace: 'Arial'
});

// 箭头指示
slide6.addText('↓', {
  x: 4.7, y: 2.65, w: 0.6, h: 0.3,
  fontSize: 32,
  color: colors.accent,
  align: 'center'
});

slide6.addText('↓', {
  x: 4.7, y: 3.95, w: 0.6, h: 0.3,
  fontSize: 32,
  color: colors.accent,
  align: 'center'
});

// 核心原则
slide6.addShape(pptx.ShapeType.rect, {
  x: 1, y: 5.6, w: 8, h: 1,
  fill: { color: colors.white },
  line: { color: colors.primary, width: 3 }
});
slide6.addText('核心原则：头最大，越往下越小，最能吸引读者注意', {
  x: 1.3, y: 5.8, w: 7.4, h: 0.6,
  fontSize: 18,
  bold: true,
  color: colors.primary,
  align: 'center',
  fontFace: 'Arial'
});

// ==========================================
// 幻灯片 7: 新闻特写范文
// ==========================================
let slide7 = pptx.addSlide();
slide7.background = { color: colors.light };

// 标题栏
slide7.addShape(pptx.ShapeType.rect, {
  x: 0, y: 0, w: 10, h: 1.2,
  fill: { color: colors.primary }
});
slide7.addText('第五部分：新闻特写范文', {
  x: 0.5, y: 0.3, w: 9, h: 0.8,
  fontSize: 28,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

// 标题
slide7.addText('时代需要最可爱的人', {
  x: 0.5, y: 1.4, w: 9, h: 0.6,
  fontSize: 32,
  bold: true,
  color: colors.primary,
  align: 'center',
  fontFace: 'Arial'
});

// 副标题
slide7.addText('—— 记著名作家魏巍同李国安会见', {
  x: 0.5, y: 2, w: 9, h: 0.4,
  fontSize: 18,
  color: colors.secondary,
  align: 'center',
  italic: true,
  fontFace: 'Calibri'
});

// 内容要点框
slide7.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 2.6, w: 4, h: 3,
  fill: { color: colors.white },
  line: { color: colors.secondary, width: 2 }
});

// 装饰色块
slide7.addShape(pptx.ShapeType.rect, {
  x: 0.5, y: 2.6, w: 0.3, h: 3,
  fill: { color: colors.secondary }
});

slide7.addText('内容要点', {
  x: 1, y: 2.8, w: 3.3, h: 0.4,
  fontSize: 18,
  bold: true,
  color: colors.primary,
  fontFace: 'Arial'
});

const featureDetails = [
  '• 老作家魏巍会见"模范团长"李国安',
  '• 76岁的老作家与军人握手',
  '• 两个时代的英雄对话',
  '• 魏巍称李国安为"和平建设年代最可爱的人"'
];

let yFeature = 3.3;
featureDetails.forEach(detail => {
  slide7.addText(detail, {
    x: 1, y: yFeature, w: 3.3, h: 0.4,
    fontSize: 13,
    color: colors.dark,
    fontFace: 'Calibri'
  });
  yFeature += 0.4;
});

// 主题框
slide7.addShape(pptx.ShapeType.rect, {
  x: 5, y: 2.6, w: 4.5, h: 3,
  fill: { color: colors.primary },
  line: { color: colors.accent, width: 2 }
});

slide7.addText('核心主题', {
  x: 5.3, y: 2.8, w: 3.9, h: 0.4,
  fontSize: 18,
  bold: true,
  color: colors.white,
  fontFace: 'Arial'
});

slide7.addText('两个时代的英雄，都是最可爱的人', {
  x: 5.3, y: 3.3, w: 3.9, h: 1.2,
  fontSize: 20,
  bold: true,
  color: colors.accent,
  align: 'center',
  fontFace: 'Arial'
});

// ==========================================
// 幻灯片 8: 结束页
// ==========================================
let slide8 = pptx.addSlide();
slide8.background = { color: colors.primary };

// 谢谢文字
slide8.addText('谢谢观看', {
  x: 0.5, y: 3, w: 9, h: 0.8,
  fontSize: 54,
  bold: true,
  color: colors.white,
  align: 'center',
  fontFace: 'Arial'
});

slide8.addText('Thank You', {
  x: 0.5, y: 3.8, w: 9, h: 0.5,
  fontSize: 32,
  color: colors.accent,
  align: 'center',
  fontFace: 'Arial'
});

// 装饰线
slide8.addShape(pptx.ShapeType.line, {
  x: 3, y: 4.5, w: 4, h: 0,
  line: { color: colors.accent, width: 3 }
});

// 副标题
slide8.addText('Broadcast News & Feature Writing Guide', {
  x: 0.5, y: 4.8, w: 9, h: 0.4,
  fontSize: 18,
  color: 'B8E4F0',
  align: 'center',
  fontFace: 'Calibri'
});

// 保存文件
const outputPath = path.join(process.env.HOME, 'Documents', '广播新闻稿教学.pptx');
pptx.writeFile({ fileName: outputPath })
  .then(() => {
    console.log(`演示文稿已成功保存到: ${outputPath}`);
  })
  .catch(err => {
    console.error('保存失败:', err);
  });


//wwdadadjsad 

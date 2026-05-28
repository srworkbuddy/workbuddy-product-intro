const pptxgen = require('pptxgenjs');
const html2pptx = require('/Users/annycheng/.workbuddy/skills/pptx/scripts/html2pptx');
const path = require('path');

async function createPresentation() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.author = 'Tencent CSIG';
  pptx.title = 'WorkBuddy 产品介绍';

  const slideDir = '/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/slides';
  const slideFiles = [
    'slide01.html', 'slide02.html', 'slide03.html', 'slide04.html',
    'slide05.html', 'slide06.html', 'slide07.html', 'slide08.html',
    'slide08b.html', 'slide09.html', 'slide10.html', 'slide11.html',
    'slide12.html'
  ];

  for (const file of slideFiles) {
    const filePath = path.join(slideDir, file);
    console.log(`Processing ${file}...`);
    await html2pptx(filePath, pptx);
  }

  const outputPath = '/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/WorkBuddy产品介绍.pptx';
  await pptx.writeFile({ fileName: outputPath });
  console.log('PPT created successfully!');
}

createPresentation().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});

const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-crashpad', '--disable-dev-shm-usage', '--disable-crashpad-for-testing'],
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    ignoreDefaultArgs: ['--enable-automation']
  });
  const page = await browser.newPage();
  await page.setViewport({width: 1200, height: 800, deviceScaleFactor: 3});
  
  const htmlPath = path.resolve('/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/WorkBuddy产品介绍_长图.html');
  await page.goto('file://' + htmlPath, { waitUntil: 'networkidle0' });
  
  const body = await page.$('body');
  await body.screenshot({
    path: '/Users/annycheng/WorkBuddy/2026-05-26-11-02-48/WorkBuddy产品介绍_长图.png',
    type: 'png'
  });
  
  console.log('PNG exported successfully');
  await browser.close();
})();

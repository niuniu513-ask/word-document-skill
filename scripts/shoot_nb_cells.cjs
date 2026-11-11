// 离屏截取 notebook 单元格：一次启动浏览器，整页长图 + 逐单元格元素截图。
//
// 为什么不用 VS Code + CDP：notebook 是虚拟化列表，滚动状态会被记住，
// 大输出单元格的行高与位置需要推算，任何一步不准都会截到别的单元格。
// 这里直接渲染成完整 DOM，元素边界由浏览器给出，不存在滚动与定位问题。
//
// 用法：
//   node scripts/shoot_nb_cells.cjs EVIDENCE/nb.html EVIDENCE/shots
//   node scripts/shoot_nb_cells.cjs EVIDENCE/nb.html EVIDENCE/shots --scale 2 --browser "C:/.../chrome.exe"
//
// 产物（<outdir>）：
//   full.png                 整页长图（留档/总览）
//   012-code.png             第 12 个原始 cell 的代码区
//   012-output.png           同 cell 的输出区（无输出则不生成）
//   012-cell.png             同 cell 的"代码+输出"连体图
//   screenshots-cropped.json 清单，可直接作为 lab_report.py build 的 screenshots 输入
//   manifest.json            附坐标与首行，便于人工核对配对

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function sha256(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function arg(name, dflt) {
  const i = process.argv.indexOf(name);
  return i > -1 && process.argv[i + 1] ? process.argv[i + 1] : dflt;
}

function loadPlaywright() {
  const tries = [
    process.env.PLAYWRIGHT_MODULE,
    'playwright-core',
    'playwright',
    path.join(__dirname, '..', 'node_modules', 'playwright-core'),
    path.join(__dirname, '..', 'node_modules', 'playwright'),
    path.join(process.cwd(), 'node_modules', 'playwright-core'),
  ].filter(Boolean);
  const errors = [];
  for (const t of tries) {
    try {
      return require(t);
    } catch (e) {
      errors.push(`${t}: ${e.message.split('\n')[0]}`);
    }
  }
  throw new Error(
    '找不到 playwright-core。先执行 npm install playwright-core，' +
    '或用 PLAYWRIGHT_MODULE 指向已有安装目录。\n尝试过：\n  ' + errors.join('\n  ')
  );
}

function findBrowser() {
  const cands = [
    process.env.CHROME_PATH,
    'C:/Program Files/Google/Chrome/Application/chrome.exe',
    'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
    'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
    'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
  ].filter(Boolean);
  for (const c of cands) {
    if (fs.existsSync(c)) return c;
  }
  throw new Error('找不到 Chrome/Edge。用 --browser 指定，或设置 CHROME_PATH。');
}

(async () => {
  const htmlPath = path.resolve(process.argv[2] || '');
  const outDir = path.resolve(process.argv[3] || '');
  const scale = Number(arg('--scale', '2'));
  const browserPath = arg('--browser', null) || findBrowser();

  if (!htmlPath || !fs.existsSync(htmlPath)) {
    throw new Error(`HTML 不存在: ${htmlPath}（先跑 render_nb_html.py）`);
  }
  fs.mkdirSync(outDir, { recursive: true });

  const { chromium } = loadPlaywright();
  const browser = await chromium.launch({
    headless: true,
    executablePath: browserPath,
    args: ['--hide-scrollbars', '--font-render-hinting=none'],
  });
  const page = await browser.newPage({
    viewport: { width: 1280, height: 2000 },
    deviceScaleFactor: scale,
  });

  await page.goto('file:///' + htmlPath.replace(/\\/g, '/'), { waitUntil: 'load' });
  // 内嵌 base64 图必须解码完成，否则截到空白
  await page.waitForFunction(
    () => Array.from(document.images).every((i) => i.complete && i.naturalWidth > 0),
    { timeout: 60000 }
  );
  await page.evaluate(() => (document.fonts ? document.fonts.ready : Promise.resolve()));
  await page.waitForTimeout(300);

  const full = path.join(outDir, 'full.png');
  await page.screenshot({ path: full, fullPage: true });

  const cells = await page.$$('.cell');
  const items = [];
  for (const cell of cells) {
    const idx = Number(await cell.getAttribute('data-idx'));
    const hasOut = (await cell.getAttribute('data-has-output')) === '1';
    const tag = String(idx).padStart(3, '0');
    const rec = { cell: idx, hasOutput: hasOut };

    const input = await cell.$('.input');
    if (input) {
      const f = path.join(outDir, `${tag}-code.png`);
      await input.screenshot({ path: f });
      const b = await input.boundingBox();
      rec.codeCrop = f;
      rec.codeBox = b;
      rec.codeAudit = {
        mode: 'element',
        raw_sha256: sha256(f),
        output_sha256: sha256(f),
        crop_ltrb_px: [0, 0, Math.round(b.width * scale), Math.round(b.height * scale)],
      };
    }

    if (hasOut) {
      const output = await cell.$('.output');
      if (output) {
        const f = path.join(outDir, `${tag}-output.png`);
        await output.screenshot({ path: f });
        const b = await output.boundingBox();
        rec.outputCrop = f;
        rec.outputBox = b;
        rec.outputAudit = {
          mode: 'element',
          raw_sha256: sha256(f),
          output_sha256: sha256(f),
          crop_ltrb_px: [0, 0, Math.round(b.width * scale), Math.round(b.height * scale)],
        };
      }
    }

    // 连体图：代码与输出在同一张里，报告排版更紧凑
    const fc = path.join(outDir, `${tag}-cell.png`);
    await cell.screenshot({ path: fc });
    rec.cellCrop = fc;

    rec.sourceFirstLine = (await cell.$eval('.input pre', (e) => e.innerText.split('\n')[0])).trim();
    // 全量源码：供 lab_report.py audit 与 notebook 逐字对账（离屏路线没有证据副本）
    rec.source = await cell.$eval('.input pre', (e) => e.innerText);
    items.push(rec);
    console.log(`cell ${idx}: code${hasOut ? '+output' : ''} -> ${tag}-*.png`);
  }

  const payload = {
    html: htmlPath,
    pageWidth: await page.evaluate(() => document.body.scrollWidth),
    scale,
    fullImage: full,
    cells: items,
  };
  fs.writeFileSync(path.join(outDir, 'screenshots-cropped.json'),
    JSON.stringify(payload, null, 1));
  fs.writeFileSync(path.join(outDir, 'manifest.json'),
    JSON.stringify(payload, null, 1));

  console.log(`\n共 ${items.length} 个 cell，有输出 ${items.filter((i) => i.hasOutput).length} 个`);
  console.log(`清单: ${path.join(outDir, 'screenshots-cropped.json')}`);
  await browser.close();
})().catch((e) => {
  console.error(e.message || e);
  process.exit(1);
});

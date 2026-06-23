// Рендерит самодостаточный HTML офера в PDF формата A4 средствами Chromium (Playwright).
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
// Playwright из глобальной установки (CommonJS) — резолвим через require.
const { chromium } = require(process.env.PLAYWRIGHT_PKG || 'playwright');
import path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
const htmlPath = path.join(here, 'Коммерческое-предложение-Факторинг.html');
const pdfPath  = path.join(here, 'Коммерческое-предложение-Факторинг.pdf');

const browser = await chromium.launch();
const page = await browser.newPage();
await page.goto('file://' + htmlPath, { waitUntil: 'networkidle' });
await page.emulateMedia({ media: 'print' });
await page.pdf({
  path: pdfPath,
  format: 'A4',
  printBackground: true,
  preferCSSPageSize: true,
  margin: { top: '0', right: '0', bottom: '0', left: '0' },
});
await browser.close();
console.log('PDF -> ' + pdfPath);

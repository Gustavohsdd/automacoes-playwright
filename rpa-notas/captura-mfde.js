require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 60 });
  const context = await browser.newContext({ storageState: 'auth.json' }); // usa sua sessão já logada
  const page = await context.newPage();

  // Abre a tela de lançamentos
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // Salva imagem da página inteira e o HTML
  await page.screenshot({ path: 'mfde.png', fullPage: true });
  const html = await page.content();
  fs.writeFileSync('mfde.html', html, 'utf8');

  console.log('✔️ Salvei mfde.png e mfde.html na pasta do projeto.');
  await page.waitForTimeout(2000);
  await browser.close();
})();

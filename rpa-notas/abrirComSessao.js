require('dotenv').config();
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext({ storageState: 'auth.json' }); // reutiliza o login
  const page = await context.newPage();

  // Abre a página inicial do sistema (ou cai no login se a sessão expirou)
  await page.goto(process.env.URL_HOME || process.env.URL_LOGIN, { waitUntil: 'domcontentloaded' });

  console.log('✅ Janela aberta com sessão carregada.');
  console.log('➡️  Navegue MANUALMENTE até a tela onde você lança as notas.');
  console.log('📋 Copie a URL dessa tela e cole aqui no chat.');

  // Mantém a janela aberta por 10 minutos para você navegar
  await page.waitForTimeout(10 * 60 * 1000);

  await browser.close();
})();

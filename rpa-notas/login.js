require('dotenv').config();
const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto(process.env.URL_LOGIN, { waitUntil: 'domcontentloaded' });

  // Campos de código/usuário e senha (pelos placeholders da sua tela)
  await page.getByPlaceholder(/codigo/i).fill(process.env.USUARIO);
  await page.getByPlaceholder(/senha/i).fill(process.env.SENHA);

  // Seleção da Empresa:
  const empresa = process.env.EMPRESA;

  try {
    // 1) Tenta como <select> padrão
    await page.getByRole('combobox', { name: /empresa/i }).selectOption({ label: empresa });
  } catch {
    // 2) Se for menu customizado, clica e escolhe pelo texto visível
    const combo = page.getByRole('combobox', { name: /empresa/i });
    await combo.click();
    // tenta pegar a opção pelo texto
    await page.getByRole('option', { name: new RegExp(empresa, 'i') }).click();
  }

  // Entrar e aguardar carregar
  await Promise.all([
    page.waitForLoadState('networkidle'),
    page.getByRole('button', { name: /entrar/i }).click()
  ]);

  // Salva a sessão para reutilizar depois
  await context.storageState({ path: 'auth.json' });
  console.log('Login ok. Sessão salva em auth.json');

  await browser.close();
})();

const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({
    headless: false
  });
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/login');
  await page.getByRole('textbox', { name: 'Codigo' }).fill('meu codigo');
  await page.getByRole('textbox', { name: 'Codigo' }).press('Tab');
  await page.getByRole('textbox', { name: 'Senha' }).fill('minha senha');
  await page.getByRole('textbox', { name: 'Senha' }).press('Tab');
  await page.getByLabel('Empresa *').press('ArrowDown');
  await page.getByLabel('Empresa *').selectOption('1');
  await page.getByRole('button', { name: 'Entrar' }).click();
  await page.getByRole('button', { name: 'Produto' }).click();
  await page.getByRole('paragraph').filter({ hasText: 'Perda' }).click();
  await page.getByRole('button', { name: 'Incluir' }).click();
  await page.locator('input[type="date"]').fill('2025-09-09');
  await page.getByRole('textbox', { name: 'Código Produto' }).click();
  await page.getByRole('textbox', { name: 'Código Produto' }).click();
  await page.getByRole('textbox', { name: 'Código Produto' }).fill('2004620017491');
  await page.getByRole('textbox', { name: 'Código Produto' }).press('Enter');
  await page.getByRole('textbox', { name: 'Quantidade' }).press('Enter');
  await page.getByRole('button', { name: 'Gravar' }).click();

  // ---------------------
  await context.close();
  await browser.close();
})();
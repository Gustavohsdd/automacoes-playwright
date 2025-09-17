require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext({ storageState: 'auth.json' });
  const page = await context.newPage();

  // Abre a tela de Manifestação
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // Garante que o acordeon de filtros esteja ABERTO
  try {
    const campoChave = page.getByPlaceholder(/chave/i);
    if (!(await campoChave.isVisible())) {
      await page.getByText(/dados de pesquisa/i).click();
      await page.waitForTimeout(300);
    }
  } catch (_) {
    // se já estiver aberto, segue o jogo
  }

  // Lê a PRIMEIRA chave do arquivo
  const linhas = fs.readFileSync('chaves.txt', 'utf8')
    .split(/\r?\n/)
    .map(s => s.trim())
    .filter(Boolean);

  const chave = linhas[0];
  if (!chave) {
    console.log('⚠️ Coloque ao menos 1 chave dentro de chaves.txt (uma por linha).');
    await browser.close();
    return;
  }

  console.log('Usando a chave:', chave);

  // Preenche o campo "Chave"
  await page.getByPlaceholder(/chave/i).fill(chave);

  // Clica no botão "Pesquisar"
  await Promise.all([
    page.waitForLoadState('networkidle'),
    page.getByRole('button', { name: /pesquisar/i }).click()
  ]);

  // Salva um print do resultado para conferência
  await page.screenshot({ path: 'resultado-1.png', fullPage: true });
  console.log('✔️ Pesquisou a chave e salvou resultado-1.png');

  // Deixa a janela aberta 10s para você ver
  await page.waitForTimeout(10000);
  await browser.close();
})();

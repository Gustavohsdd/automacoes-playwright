require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext({ storageState: 'auth.json' });
  const page = await context.newPage();

  page.on('dialog', async d => { console.log('Confirmação:', d.message()); await d.accept(); });

  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // Abre acordeon se precisar
  try {
    const campoChave = page.getByPlaceholder(/chave/i);
    if (!(await campoChave.isVisible())) {
      await page.getByText(/dados de pesquisa/i).click();
      await page.waitForTimeout(300);
    }
  } catch {}

  // Lê a primeira chave
  const linhas = fs.readFileSync('chaves.txt', 'utf8').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const chave = linhas[0];
  if (!chave) { console.log('⚠️ Coloque 1 chave em chaves.txt'); await browser.close(); return; }

  // Pesquisa
  await page.getByPlaceholder(/chave/i).fill(chave);
  await Promise.all([
    page.waitForLoadState('networkidle'),
    page.getByRole('button', { name: /pesquisar/i }).click()
  ]);

  // Localiza tabela e linha
  const tabela = page.locator('table').filter({ has: page.locator('tbody tr') }).first();
  await tabela.waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
  const linhasTabela = tabela.locator('tbody tr');
  const count = await linhasTabela.count();
  if (count === 0) {
    console.log('⚠️ Nenhuma linha encontrada.');
    await page.screenshot({ path: 'resultado-sem-linhas.png', fullPage: true });
    await browser.close();
    return;
  }

  let linha = linhasTabela.filter({ hasText: chave }).first();
  if (await linha.count() === 0) linha = linhasTabela.nth(0);

  // Botão Executar Entrada
  const btnExecutar = page.getByRole('button', { name: /executar entrada/i });

  // Função para esperar botão habilitar
  async function esperarHabilitar(maxMs = 4000) {
    const t0 = Date.now();
    while (Date.now() - t0 < maxMs) {
      const disabled = await btnExecutar.isDisabled().catch(() => false);
      if (!disabled) return true;
      await page.waitForTimeout(150);
    }
    return false;
  }

  // TENTATIVA A: clique na própria linha
  await linha.scrollIntoViewIfNeeded();
  await linha.click();
  await page.waitForTimeout(250);
  if (!(await esperarHabilitar())) {
    // TENTATIVA B: clique na primeira célula (alguns grids exigem clicar na célula)
    const primeiraCelula = linha.locator('td').first();
    if (await primeiraCelula.count()) {
      await primeiraCelula.click();
      await page.waitForTimeout(250);
    }
  }
  if (!(await esperarHabilitar())) {
    // TENTATIVA C: double-click
    await linha.dblclick();
    await page.waitForTimeout(250);
  }
  if (!(await esperarHabilitar())) {
    // TENTATIVA D (fallback): Selecionar Todas
    const selTodas = page.getByRole('button', { name: /selecionar todas/i });
    if (await selTodas.count()) {
      await selTodas.click();
      await page.waitForTimeout(300);
    }
  }

  // Verifica estado do botão
  const habilitado = !(await btnExecutar.isDisabled().catch(() => false));
  console.log('Botão "Executar Entrada" habilitado?', habilitado);

  // Se habilitou, clica
  if (habilitado) {
    await Promise.all([
      page.waitForLoadState('networkidle'),
      btnExecutar.click()
    ]);
    await page.waitForTimeout(1500);
    await page.screenshot({ path: 'apos-executar-entrada.png', fullPage: true });
    console.log('✔️ Cliquei em Executar Entrada. Print salvo.');
  } else {
    console.log('⚠️ Não consegui habilitar o botão. Salvando print para análise.');
    await page.screenshot({ path: 'nao-habilitou.png', fullPage: true });
  }

  await page.waitForTimeout(8000);
  await browser.close();
})();

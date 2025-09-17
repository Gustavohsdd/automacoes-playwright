require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');

// ===== Modal da automação (overlay simples) =====
async function mostrarModal(page, { titulo, mensagem, botoes }) {
  await page.evaluate(({ titulo, mensagem, botoes }) => {
    const old = document.getElementById('__rpa_modal');
    if (old) old.remove();

    const overlay = document.createElement('div');
    overlay.id = '__rpa_modal';
    Object.assign(overlay.style, {
      position: 'fixed', inset: '0', background: 'rgba(0,0,0,.45)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 2147483647
    });

    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;min-width:420px;max-width:760px;padding:20px 24px;border-radius:12px;font-family:system-ui,Segoe UI,Roboto,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.2)';
    const h2 = document.createElement('h2');
    h2.textContent = titulo || 'Atenção';
    h2.style.cssText = 'margin:0 0 8px 0;font-size:20px';
    const p = document.createElement('p');
    p.textContent = mensagem || '';
    p.style.cssText = 'margin:0 0 16px 0;white-space:pre-wrap;line-height:1.4';
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';

    botoes.forEach((b, idx) => {
      const btn = document.createElement('button');
      btn.textContent = b.label;
      btn.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
      if (b.primary) { btn.style.background = '#2563eb'; btn.style.color = '#fff'; btn.style.borderColor = '#2563eb'; }
      btn.onclick = () => { window.__rpaDecision = idx; overlay.remove(); };
      row.appendChild(btn);
    });

    box.append(h2, p, row);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
  }, { titulo, mensagem, botoes });

  const res = await page.waitForFunction(() => window.__rpaDecision !== undefined, { timeout: 0 });
  const idx = await res.jsonValue();
  await page.evaluate(() => { delete window.__rpaDecision; });
  return idx;
}

// ===== Fechar modais do SISTEMA (Bootstrap/Swal/etc.) =====
async function fecharModaisSistema(page, tentativas = 10) {
  for (let i = 0; i < tentativas; i++) {
    // Pressiona ESC (alguns modais fecham assim)
    await page.keyboard.press('Escape').catch(() => {});
    // Tenta clicar em botões de fechar comuns
    await page.evaluate(() => {
      const visivel = (el) => el && el.offsetParent !== null;
      const candidatos = [
        ...document.querySelectorAll('.modal.show, .modal[style*="display: block"], [role="dialog"], .swal2-container')
      ].filter(visivel);

      candidatos.forEach(modal => {
        const btns = [
          modal.querySelector('[data-bs-dismiss="modal"]'),
          modal.querySelector('[data-dismiss="modal"]'),
          modal.querySelector('.btn-close'),
          modal.querySelector('button.close'),
          [...modal.querySelectorAll('button, a')].find(b => /fechar|close|ok|entendi|continuar/i.test(b?.textContent || ''))
        ].filter(Boolean);
        if (btns.length) btns[0].click();
      });
    }).catch(() => {});
    await page.waitForTimeout(250);
    // se não houver mais backdrop/modal, sai
    const aindaTem = await page.evaluate(() => {
      return !!document.querySelector('.modal.show, .modal[style*="display: block"], .modal-backdrop, [role="dialog"], .swal2-container');
    }).catch(() => false);
    if (!aindaTem) return;
  }
}

// ===== Verificar se a tela "Importa Danfe" está aberta =====
async function esperarImportaDanfe(page, timeoutMs = 20000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    await fecharModaisSistema(page, 2);
    const visivel = await page.getByText(/importa danfe/i).first().isVisible().catch(() => false);
    if (visivel) return true;
    await page.waitForTimeout(300);
  }
  return false;
}

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext({ storageState: 'auth.json' });
  const page = await context.newPage();

  page.on('dialog', async d => { await d.accept(); });

  // 1) Abre a lista
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // 2) Abre filtros se preciso
  try {
    const campoChave = page.getByPlaceholder(/chave/i);
    if (!(await campoChave.isVisible())) {
      await page.getByText(/dados de pesquisa/i).click();
      await page.waitForTimeout(300);
    }
  } catch {}

  // 3) Lê 1ª chave
  const chave = fs.readFileSync('chaves.txt', 'utf8')
    .split(/\r?\n/).map(s => s.trim()).filter(Boolean)[0];
  if (!chave) { await mostrarModal(page, { titulo: 'Sem chaves', mensagem: 'Adicione ao menos 1 chave no arquivo chaves.txt', botoes: [{label:'Ok', primary:true}] }); await browser.close(); return; }

  // 4) Pesquisa
  await page.getByPlaceholder(/chave/i).fill(chave);
  await Promise.all([
    page.waitForLoadState('networkidle'),
    page.getByRole('button', { name: /pesquisar/i }).click()
  ]);

  // 5) Seleciona linha e habilita "Executar Entrada"
  const tabela = page.locator('table').filter({ has: page.locator('tbody tr') }).first();
  await tabela.waitFor({ state: 'visible', timeout: 15000 }).catch(() => {});
  let linha = tabela.locator('tbody tr').filter({ hasText: chave }).first();
  if (await linha.count() === 0) linha = tabela.locator('tbody tr').first();

  const btnExecutar = page.getByRole('button', { name: /executar entrada/i });
  async function habilitarExecutar() {
    for (let i = 0; i < 6; i++) {
      await linha.click();
      await page.waitForTimeout(200);
      if (!await btnExecutar.isDisabled().catch(() => false)) return true;
      const td = linha.locator('td').first();
      if (await td.count()) { await td.click(); await page.waitForTimeout(200); }
      if (!await btnExecutar.isDisabled().catch(() => false)) return true;
    }
    return false;
  }
  await habilitarExecutar();

  // 6) Clica em Executar Entrada e ESPERA a tela abrir (fechando modais do sistema)
  await Promise.allSettled([
    page.waitForLoadState('networkidle'),
    btnExecutar.click()
  ]);
  const abriuDireto = await esperarImportaDanfe(page, 12000);

  // 7) Se não abriu, oferece opções em modal da automação (sem fechar o navegador)
  if (!abriuDireto) {
    while (true) {
      const escolha = await mostrarModal(page, {
        titulo: 'Não consegui abrir a nota',
        mensagem: 'Pode ser XML indisponível ou nota cancelada.\nEscolha uma opção:',
        botoes: [
          { label: 'Tentar manualmente', primary: true },
          { label: 'Repetir clique em "Executar Entrada"' },
          { label: 'Pular esta nota' }
        ]
      });

      if (escolha === 0) {
        await mostrarModal(page, {
          titulo: 'Abra manualmente',
          mensagem: 'Abra você mesmo a tela da nota. Quando a tela "Importa Danfe" aparecer, clique em Continuar.',
          botoes: [{ label: 'Continuar', primary: true }, { label: 'Cancelar' }]
        });
        const ok = await esperarImportaDanfe(page, 20000);
        if (ok) break; // segue
      } else if (escolha === 1) {
        await fecharModaisSistema(page, 5);
        await Promise.allSettled([
          page.waitForLoadState('networkidle'),
          btnExecutar.click()
        ]);
        const ok = await esperarImportaDanfe(page, 12000);
        if (ok) break;
      } else {
        await mostrarModal(page, { titulo: 'Pulando', mensagem: 'Esta nota foi pulada.', botoes: [{label:'Ok', primary:true}] });
        await browser.close();
        return;
      }
    }
  }

  // 8) Você preenche Operação e CFOP
  await mostrarModal(page, {
    titulo: 'Operação e CFOP',
    mensagem: 'Digite a Operação (*) e escolha o CFOP (*). Quando terminar, clique em Continuar.',
    botoes: [{ label: 'Continuar', primary: true }]
  });

  // 9) Vai para aba "Dados do Produtos" (tentando variações)
  let clicouAba = false;
  for (const alvo of [
    page.getByRole('tab', { name: /dados do produtos/i }),
    page.getByRole('tab', { name: /dados do produto/i }),
    page.getByText(/dados do produtos/i),
    page.getByText(/dados do produto/i),
  ]) {
    try { if (await alvo.count()) { await alvo.click(); clicouAba = true; break; } } catch {}
  }
  if (!clicouAba) {
    await mostrarModal(page, {
      titulo: 'Clique na aba de produtos',
      mensagem: 'Não consegui clicar na aba. Clique você na aba "Dados do Produtos" e depois em Continuar.',
      botoes: [{ label: 'Continuar', primary: true }]
    });
  }
  await page.waitForTimeout(700);

  // 10) Captura a maior tabela visível e salva CSV (;)
  const tabelas = page.locator('table');
  const n = await tabelas.count();
  let melhor = null, melhorLinhas = 0;
  for (let i = 0; i < n; i++) {
    const t = tabelas.nth(i);
    if (await t.isVisible()) {
      const linhas = await t.locator('tbody tr').count().catch(() => 0);
      if (linhas > melhorLinhas) { melhor = t; melhorLinhas = linhas; }
    }
  }
  if (!melhor || melhorLinhas === 0) {
    await mostrarModal(page, { titulo: 'Sem itens', mensagem: 'Não encontrei a tabela de produtos nesta nota.', botoes: [{label:'Ok', primary:true}] });
    await browser.close();
    return;
  }

  const linhas = melhor.locator('tbody tr');
  const qtd = await linhas.count();
  const dados = [];
  for (let i = 0; i < qtd; i++) {
    const cols = linhas.nth(i).locator('td');
    const qcols = await cols.count();
    const row = [];
    for (let j = 0; j < qcols; j++) {
      const txt = (await cols.nth(j).innerText()).trim().replace(/\s+/g, ' ');
      row.push(txt);
    }
    dados.push(row);
  }
  const csv = dados.map(r => r.map(v => `"${v.replace(/"/g, '""')}"`).join(';')).join('\n');
  fs.writeFileSync('produtos.csv', csv, 'utf8');

  await mostrarModal(page, { titulo: 'Lista capturada', mensagem: `Encontrei ${qtd} item(ns). Salvei "produtos.csv".`, botoes: [{label:'Ok', primary:true}] });
  await browser.close();
})();

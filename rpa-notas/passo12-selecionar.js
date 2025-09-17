require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');
const crypto = require('crypto');

// ======= util: modal simples (mensagem com botões) =======
async function mostrarModal(page, { titulo, mensagem, botoes }) {
  await page.evaluate(({ titulo, mensagem, botoes }) => {
    const old = document.getElementById('__rpa_modal'); if (old) old.remove();
    const overlay = document.createElement('div');
    overlay.id = '__rpa_modal';
    Object.assign(overlay.style, {
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 2147483647
    });
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;min-width:420px;max-width:780px;padding:20px 24px;border-radius:12px;font-family:system-ui,Segoe UI,Roboto,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.2)';
    const h2 = Object.assign(document.createElement('h2'), { textContent: titulo || 'Atenção' });
    h2.style.cssText = 'margin:0 0 8px;font-size:20px';
    const p = Object.assign(document.createElement('p'), { textContent: mensagem || '' });
    p.style.cssText = 'margin:0 0 16px;white-space:pre-wrap;line-height:1.4';
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';
    botoes.forEach((b, i) => {
      const btn = document.createElement('button');
      btn.textContent = b.label;
      btn.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
      if (b.primary) { btn.style.background = '#2563eb'; btn.style.borderColor = '#2563eb'; btn.style.color = '#fff'; }
      btn.onclick = () => { window.__rpaDecision = i; overlay.remove(); };
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

// ======= util: fechar modais nativos do sistema =======
async function fecharModaisSistema(page, tentativas = 8) {
  for (let i = 0; i < tentativas; i++) {
    await page.keyboard.press('Escape').catch(() => {});
    await page.evaluate(() => {
      const vis = (el) => el && el.offsetParent !== null;
      const modais = [
        ...document.querySelectorAll('.modal.show, .modal[style*="display: block"], [role="dialog"], .swal2-container')
      ].filter(vis);
      modais.forEach(m => {
        const btn =
          m.querySelector('[data-bs-dismiss="modal"]') ||
          m.querySelector('[data-dismiss="modal"]') ||
          m.querySelector('.btn-close') ||
          m.querySelector('button.close') ||
          [...m.querySelectorAll('button,a')].find(b => /fechar|close|ok|entendi|continuar/i.test(b?.textContent||''));
        btn?.click();
      });
    }).catch(() => {});
    await page.waitForTimeout(200);
    const aindaTem = await page.evaluate(() =>
      !!document.querySelector('.modal.show, .modal[style*="display: block"], .modal-backdrop, [role="dialog"], .swal2-container')
    ).catch(() => false);
    if (!aindaTem) return;
  }
}

// ======= esperar tela "Importa Danfe" =======
async function esperarImportaDanfe(page, timeoutMs = 20000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    await fecharModaisSistema(page, 2);
    const ok = await page.getByText(/importa danfe/i).first().isVisible().catch(() => false);
    if (ok) return true;
    await page.waitForTimeout(300);
  }
  return false;
}

// ======= modal de seleção com tabela e checkboxes =======
async function modalSelecaoProdutos(page, itens) {
  await page.evaluate((itens) => {
    const old = document.getElementById('__rpa_lista'); if (old) old.remove();

    const overlay = document.createElement('div');
    overlay.id = '__rpa_lista';
    Object.assign(overlay.style, {
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,.45)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2147483647
    });

    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;width:min(980px,96vw);max-height:min(86vh,900px);padding:18px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);font-family:system-ui,Segoe UI,Roboto,sans-serif;display:flex;flex-direction:column';

    const head = document.createElement('div');
    head.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:10px';
    const h = document.createElement('h2'); h.textContent = 'Selecione os produtos para editar';
    h.style.cssText = 'font-size:18px;margin:0';
    const markAll = document.createElement('button');
    markAll.textContent = 'Marcar/Desmarcar todos';
    markAll.style.cssText = 'padding:6px 10px;border-radius:8px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';

    const wrap = document.createElement('div');
    wrap.style.cssText = 'overflow:auto;border:1px solid #e5e7eb;border-radius:10px';
    const table = document.createElement('table');
    table.style.cssText = 'width:100%;border-collapse:separate;border-spacing:0';
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    ['✓','Nome','Cód. Vinc.','Descr. Vinc.','CFOP'].forEach(txt => {
      const th = document.createElement('th'); th.textContent = txt;
      th.style.cssText = 'position:sticky;top:0;background:#f9fafb;text-align:left;padding:8px;border-bottom:1px solid #e5e7eb';
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    const tbody = document.createElement('tbody');

    itens.forEach((it) => {
      const tr = document.createElement('tr');
      tr.style.cssText = 'border-bottom:1px solid #f1f5f9';
      const td0 = document.createElement('td');
      td0.style.cssText = 'padding:6px 8px';
      const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = false;
      cb.dataset.idx = String(it.idx);
      td0.appendChild(cb); tr.appendChild(td0);

      [it.nome, it.codv, it.descv, it.cfop].forEach(v => {
        const td = document.createElement('td');
        td.textContent = v || '';
        td.style.cssText = 'padding:6px 8px';
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    table.append(thead, tbody);
    wrap.appendChild(table);

    const foot = document.createElement('div');
    foot.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;margin-top:10px';
    const btnCancel = document.createElement('button');
    btnCancel.textContent = 'Cancelar';
    btnCancel.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
    const btnOk = document.createElement('button');
    btnOk.textContent = 'Continuar';
    btnOk.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #2563eb;background:#2563eb;color:#fff;cursor:pointer';

    markAll.onclick = () => {
      const boxes = tbody.querySelectorAll('input[type="checkbox"]');
      const temNaoMarcado = Array.from(boxes).some(b => !b.checked);
      boxes.forEach(b => b.checked = temNaoMarcado);
    };
    btnCancel.onclick = () => { window.__rpaSel = null; overlay.remove(); };
    btnOk.onclick = () => {
      const selecionados = [...tbody.querySelectorAll('input[type="checkbox"]:checked')].map(cb => Number(cb.dataset.idx));
      window.__rpaSel = selecionados;
      overlay.remove();
    };

    head.append(h, markAll);
    box.append(head, wrap, foot);
    foot.append(btnCancel, btnOk);
    overlay.appendChild(box);
    document.body.appendChild(overlay);
  }, itens);

  const res = await page.waitForFunction(() => window.__rpaSel !== undefined, { timeout: 0 });
  const selecionados = await res.evaluate(() => window.__rpaSel);
  await page.evaluate(() => { delete window.__rpaSel; });
  return selecionados; // array de índices ou null
}

(async () => {
  const browser = await chromium.launch({ headless: false, slowMo: 80 });
  const context = await browser.newContext({ storageState: 'auth.json' });
  const page = await context.newPage();

  page.on('dialog', async d => { await d.accept(); });

  // 1) Abrir lista
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil: 'domcontentloaded' });
  await page.waitForLoadState('networkidle');

  // 2) Filtros
  try {
    const campoChave = page.getByPlaceholder(/chave/i);
    if (!(await campoChave.isVisible())) { await page.getByText(/dados de pesquisa/i).click(); await page.waitForTimeout(300); }
  } catch {}

  // 3) Chave (pega a primeira de chaves.txt)
  const chave = fs.readFileSync('chaves.txt', 'utf8').split(/\r?\n/).map(s=>s.trim()).filter(Boolean)[0];
  if (!chave) { await mostrarModal(page, { titulo:'Sem chave', mensagem:'Adicione ao menos 1 chave em chaves.txt', botoes:[{label:'Ok', primary:true}] }); await browser.close(); return; }

  await page.getByPlaceholder(/chave/i).fill(chave);
  await Promise.all([ page.waitForLoadState('networkidle'), page.getByRole('button', { name: /pesquisar/i }).click() ]);

  // 4) Seleciona linha e abre nota
  const tabela = page.locator('table').filter({ has: page.locator('tbody tr') }).first();
  await tabela.waitFor({ state:'visible', timeout:15000 }).catch(()=>{});
  let linha = tabela.locator('tbody tr').filter({ hasText: chave }).first();
  if (await linha.count() === 0) linha = tabela.locator('tbody tr').first();
  const btnExecutar = page.getByRole('button', { name: /executar entrada/i });

  async function habilitarExecutar() {
    for (let i=0;i<6;i++){
      await linha.click(); await page.waitForTimeout(200);
      if (!(await btnExecutar.isDisabled().catch(()=>false))) return true;
      const td = linha.locator('td').first(); if (await td.count()) { await td.click(); await page.waitForTimeout(200); }
      if (!(await btnExecutar.isDisabled().catch(()=>false))) return true;
    }
    return false;
  }
  await habilitarExecutar();

  await Promise.allSettled([ page.waitForLoadState('networkidle'), btnExecutar.click() ]);
  let abriu = await esperarImportaDanfe(page, 12000);

  if (!abriu) {
    while (true) {
      const escolha = await mostrarModal(page, {
        titulo: 'Não consegui abrir a nota',
        mensagem: 'Talvez XML indisponível ou nota cancelada.\nEscolha:',
        botoes: [{label:'Tentar manualmente', primary:true}, {label:'Repetir clique'}, {label:'Pular'}]
      });
      if (escolha === 0) {
        await mostrarModal(page, { titulo: 'Abra manualmente', mensagem: 'Abra a tela da nota e clique em Continuar.', botoes: [{label:'Continuar', primary:true}] });
        abriu = await esperarImportaDanfe(page, 20000);
        if (abriu) break;
      } else if (escolha === 1) {
        await fecharModaisSistema(page, 5);
        await Promise.allSettled([ page.waitForLoadState('networkidle'), btnExecutar.click() ]);
        abriu = await esperarImportaDanfe(page, 12000);
        if (abriu) break;
      } else {
        await mostrarModal(page, { titulo: 'Pulado', mensagem: 'Esta nota foi pulada.', botoes: [{label:'Ok', primary:true}] });
        await browser.close(); return;
      }
    }
  }

  // 5) Você preenche Operação e CFOP
  await mostrarModal(page, { titulo: 'Operação e CFOP', mensagem: 'Digite a Operação (*) e escolha o CFOP (*). Depois clique em Continuar.', botoes: [{label:'Continuar', primary:true}] });

  // 6) Ir para aba "Dados do Produtos"
  let clicouAba = false;
  for (const alvo of [
    page.getByRole('tab', { name: /dados do produtos/i }),
    page.getByRole('tab', { name: /dados do produto/i }),
    page.getByText(/dados do produtos/i),
    page.getByText(/dados do produto/i),
  ]) { try { if (await alvo.count()) { await alvo.click(); clicouAba = true; break; } } catch {} }
  if (!clicouAba) {
    await mostrarModal(page, { titulo: 'Clique na aba de produtos', mensagem: 'Clique você na aba "Dados do Produtos" e então Continuar.', botoes: [{label:'Continuar', primary:true}] });
  }
  await page.waitForTimeout(700);

  // 7) Capturar tabela e montar lista para seleção
  const info = await page.evaluate(() => {
    function norm(t){ return (t||'').trim().replace(/\s+/g,' '); }
    const vis = (el)=> el && el.offsetParent !== null;
    const tables = [...document.querySelectorAll('table')].filter(vis);
    let best=null, max=0;
    for (const t of tables) {
      const rows = t.querySelectorAll('tbody tr').length;
      if (rows > max) { best=t; max=rows; }
    }
    if (!best) return null;

    const headers = [...best.querySelectorAll('thead th')].map(th => norm(th.innerText.toLowerCase()));
    const map = {};
    headers.forEach((h,i) => {
      if (map.nome===undefined && /(nome|descri[cç][aã]o|produto)/i.test(h)) map.nome=i;
      if (map.codv===undefined && /c[oó]d.*vinc/i.test(h)) map.codv=i;
      if (map.descv===undefined && /descr.*vinc/i.test(h)) map.descv=i;
      if (map.cfop===undefined && /cfop/i.test(h)) map.cfop=i;
    });

    const rows = [...best.querySelectorAll('tbody tr')];
    const itens = rows.map((tr, idx) => {
      const tds = [...tr.querySelectorAll('td')];
      const g = (i)=> norm(tds[i]?.innerText || '');
      return {
        idx,
        nome: map.nome!=null ? g(map.nome) : g(0),
        codv: map.codv!=null ? g(map.codv) : '',
        descv: map.descv!=null ? g(map.descv) : '',
        cfop: map.cfop!=null ? g(map.cfop) : ''
      };
    });
    return { itens };
  });

  if (!info || !info.itens?.length) {
    await mostrarModal(page, { titulo: 'Sem itens', mensagem: 'Não achei a tabela de produtos.', botoes: [{label:'Ok', primary:true}] });
    await browser.close(); return;
  }

  // 7.1) Salvar produtos.csv (sempre)
  const linhasCSV = [
    'Nome;Cód. Vinc.;Descr. Vinc.;CFOP',
    ...info.itens.map(it => [it.nome, it.codv, it.descv, it.cfop]
      .map(v => `"${String(v||'').replace(/"/g,'""')}"`).join(';'))
  ].join('\n');
  fs.writeFileSync('produtos.csv', linhasCSV, 'utf8');

  // 8) Modal com checkboxes
  const selecionados = await modalSelecaoProdutos(page, info.itens);
  if (!selecionados) {
    await mostrarModal(page, { titulo:'Cancelado', mensagem:'Seleção cancelada.', botoes:[{label:'Ok', primary:true}] });
    await browser.close(); return;
  }

  // 9) Salva seleção SEMPRE (sobrescreve) + fingerprint da tabela
  const fingerprint = crypto
    .createHash('sha1')
    .update(JSON.stringify(info.itens.map(i => [i.nome, i.codv, i.descv, i.cfop])))
    .digest('hex');

  const payload = {
    chave,
    carimbo: new Date().toISOString(),
    total_itens: info.itens.length,
    tabela_hash: fingerprint,
    selecionados
  };
  fs.writeFileSync('selecionados.json', JSON.stringify(payload, null, 2), 'utf8');

  await mostrarModal(page, {
    titulo: 'Selecionados salvos',
    mensagem: `Você marcou ${selecionados.length} item(ns).\nArquivos gerados:\n- produtos.csv\n- selecionados.json\nAgora rode o Passo 13 para editar os itens.`,
    botoes: [{label:'Ok', primary:true}]
  });

  await browser.close();
})();

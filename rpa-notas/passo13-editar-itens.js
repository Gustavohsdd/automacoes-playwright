require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');
const crypto = require('crypto');

/* ======== Modais da automação ======== */
async function modalMsg(page, { titulo, mensagem, botoes }) {
  await page.evaluate(({ titulo, mensagem, botoes }) => {
    const old = document.getElementById('__rpa_modal'); if (old) old.remove();
    const overlay = document.createElement('div');
    overlay.id = '__rpa_modal';
    Object.assign(overlay.style, { position:'fixed', inset:0, background:'rgba(0,0,0,.45)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:2147483647 });
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;min-width:420px;max-width:780px;padding:18px 22px;border-radius:12px;font-family:system-ui,Segoe UI,Roboto,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.2)';
    const h = document.createElement('h2'); h.textContent = titulo || 'Atenção'; h.style.cssText = 'margin:0 0 8px;font-size:20px';
    const p = document.createElement('p'); p.textContent = mensagem || ''; p.style.cssText = 'margin:0 0 16px;white-space:pre-wrap;line-height:1.4';
    const row = document.createElement('div'); row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap';
    botoes.forEach((b,i)=>{
      const btn = document.createElement('button'); btn.textContent = b.label;
      btn.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
      if (b.primary){ btn.style.background='#2563eb'; btn.style.borderColor='#2563eb'; btn.style.color='#fff'; }
      btn.onclick = ()=>{ window.__rpaDecision=i; overlay.remove(); };
      row.appendChild(btn);
    });
    box.append(h,p,row); overlay.appendChild(box); document.body.appendChild(overlay);
  }, { titulo, mensagem, botoes });
  const res = await page.waitForFunction(() => window.__rpaDecision !== undefined, { timeout: 0 });
  const idx = await res.jsonValue();
  await page.evaluate(() => { delete window.__rpaDecision; });
  return idx;
}

async function modalEscolherCFOP(page, { opcoes, atual }) {
  await page.evaluate(({ opcoes, atual }) => {
    const old = document.getElementById('__rpa_cfop'); if (old) old.remove();
    const overlay = document.createElement('div'); overlay.id='__rpa_cfop';
    Object.assign(overlay.style, { position:'fixed', inset:0, background:'rgba(0,0,0,.45)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:2147483647 });
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;min-width:420px;max-width:720px;padding:18px 22px;border-radius:12px;font-family:system-ui,Segoe UI,Roboto,sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.2)';
    const h = document.createElement('h2'); h.textContent = 'Escolher CFOP'; h.style.cssText='margin:0 0 8px;font-size:20px';
    const wrap = document.createElement('div'); wrap.style.cssText='display:flex;flex-direction:column;gap:10px;margin:8px 0 16px';

    let select = null, input = null;
    if (opcoes && opcoes.length) {
      const label = document.createElement('label'); label.textContent = 'Opções disponíveis:';
      select = document.createElement('select'); select.style.cssText='padding:8px;border:1px solid #d0d7de;border-radius:8px';
      const optV = document.createElement('option'); optV.value=''; optV.textContent='(manter atual / escolher manual)'; select.appendChild(optV);
      opcoes.forEach(o => { const opt=document.createElement('option'); opt.value=o; opt.textContent=o; if (o===atual) opt.selected=true; select.appendChild(opt); });
      wrap.append(label, select);
    }
    const lab2 = document.createElement('label'); lab2.textContent = 'Ou digite exatamente como aparece no menu:';
    input = document.createElement('input'); input.type='text'; input.value = atual || ''; input.placeholder='ex.: 5.102 – Venda de produção ...';
    input.style.cssText = 'padding:8px;border:1px solid #d0d7de;border-radius:8px';
    wrap.append(lab2, input);

    const row = document.createElement('div'); row.style.cssText='display:flex;gap:8px;justify-content:flex-end';
    const btnPular = document.createElement('button'); btnPular.textContent='Pular este item'; btnPular.style.cssText='padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
    const btnAplicar = document.createElement('button'); btnAplicar.textContent='Aplicar e continuar'; btnAplicar.style.cssText='padding:8px 14px;border-radius:10px;border:1px solid #2563eb;background:#2563eb;color:#fff;cursor:pointer';

    btnPular.onclick = ()=>{ window.__rpaCFOP = { pular:true }; overlay.remove(); };
    btnAplicar.onclick = ()=>{
      const escolhido = (select && select.value) ? select.value : (input.value || '');
      window.__rpaCFOP = { pular:false, valor: (escolhido||'').trim() };
      overlay.remove();
    };

    box.append(h, wrap, row); row.append(btnPular, btnAplicar);
    overlay.appendChild(box); document.body.appendChild(overlay);
  }, { opcoes, atual });
  const res = await page.waitForFunction(() => window.__rpaCFOP !== undefined, { timeout: 0 });
  const data = await res.jsonValue();
  await page.evaluate(() => { delete window.__rpaCFOP; });
  return data;
}

/* ======== Ferramentas de página ======== */
async function fecharModaisSistema(page, tentativas = 8) {
  for (let i=0;i<tentativas;i++){
    await page.keyboard.press('Escape').catch(()=>{});
    await page.evaluate(() => {
      const vis = el => el && el.offsetParent !== null;
      const modais = [...document.querySelectorAll('.modal.show, .modal[style*="display: block"], [role="dialog"], .swal2-container')].filter(vis);
      modais.forEach(m=>{
        const btn = m.querySelector('[data-bs-dismiss="modal"], [data-dismiss="modal"], .btn-close, button.close') ||
                    [...m.querySelectorAll('button,a')].find(b=>/fechar|close|ok|entendi|continuar/i.test(b?.textContent||''));
        btn?.click();
      });
    }).catch(()=>{});
    await page.waitForTimeout(200);
    const ainda = await page.evaluate(()=>!!document.querySelector('.modal.show, .modal[style*="display: block"], .modal-backdrop, [role="dialog"], .swal2-container')).catch(()=>false);
    if (!ainda) return;
  }
}

async function esperarImportaDanfe(page, timeoutMs=20000){
  const t0=Date.now();
  while(Date.now()-t0<timeoutMs){
    await fecharModaisSistema(page,2);
    const ok = await page.getByText(/importa danfe/i).first().isVisible().catch(()=>false);
    if (ok) return true;
    await page.waitForTimeout(300);
  }
  return false;
}

async function irParaAbaProdutos(page){
  for (const alvo of [
    page.getByRole('tab', { name: /dados do produtos/i }),
    page.getByRole('tab', { name: /dados do produto/i }),
    page.getByText(/dados do produtos/i),
    page.getByText(/dados do produto/i),
  ]) { try { if (await alvo.count()) { await alvo.click(); return true; } } catch {} }
  return false;
}

async function abrirNotaPorChave(page, chave){
  await page.goto('https://araujopatrocinio.inovautomacao.com.br/mfde', { waitUntil:'domcontentloaded' });
  await page.waitForLoadState('networkidle');
  try { const campo = page.getByPlaceholder(/chave/i); if (!(await campo.isVisible())) { await page.getByText(/dados de pesquisa/i).click(); await page.waitForTimeout(300); } } catch {}
  await page.getByPlaceholder(/chave/i).fill(chave);
  await Promise.all([ page.waitForLoadState('networkidle'), page.getByRole('button', { name:/pesquisar/i }).click() ]);

  const tabela = page.locator('table').filter({ has: page.locator('tbody tr') }).first();
  await tabela.waitFor({ state:'visible', timeout:15000 }).catch(()=>{});
  let linha = tabela.locator('tbody tr').filter({ hasText: chave }).first();
  if (await linha.count()===0) linha = tabela.locator('tbody tr').first();

  const btnExec = page.getByRole('button', { name:/executar entrada/i });
  for (let i=0;i<6;i++){
    await linha.click(); await page.waitForTimeout(200);
    if (!(await btnExec.isDisabled().catch(()=>false))) break;
    const td = linha.locator('td').first(); if (await td.count()) { await td.click(); await page.waitForTimeout(200); }
  }
  await Promise.allSettled([ page.waitForLoadState('networkidle'), btnExec.click() ]);
  return await esperarImportaDanfe(page, 15000);
}

async function coletarOpcoesCFOPnoModal(page){
  const opcoes = await page.evaluate(() => {
    const vis = el => el && el.offsetParent !== null;
    const modal = [...document.querySelectorAll('.modal.show, [role="dialog"]')].filter(vis).at(-1);
    if (!modal) return { tipo:'none', opcoes:[], atual:'' };
    const texto = (el)=> (el?.innerText||'').trim();
    let select = null;
    const labels = [...modal.querySelectorAll('label')];
    for(const lb of labels){
      if (/cfop/i.test(texto(lb))){
        select = lb.parentElement?.querySelector('select') || lb.nextElementSibling?.querySelector?.('select');
        if (select) break;
      }
    }
    if (!select){
      select = modal.querySelector('select[id*="cfop" i], select[name*="cfop" i]');
    }
    if (select){
      const opts = [...select.querySelectorAll('option')].map(o=>o.textContent.trim()).filter(Boolean);
      const atual = select.selectedOptions?.[0]?.textContent?.trim() || '';
      return { tipo:'select', opcoes:opts, atual };
    }
    return { tipo:'none', opcoes:[], atual:'' };
  });
  return opcoes;
}

async function aplicarCFOPnoModal(page, label){
  if (!label) return false;
  const okSelect = await page.evaluate((label)=>{
    const vis = el => el && el.offsetParent !== null;
    const modal = [...document.querySelectorAll('.modal.show, [role="dialog"]')].filter(vis).at(-1);
    if (!modal) return false;
    const byLabel = ()=>{
      const labels = [...modal.querySelectorAll('label')];
      for(const lb of labels){
        if (/cfop/i.test((lb.innerText||'').trim())){
          const sel = lb.parentElement?.querySelector('select') || lb.nextElementSibling?.querySelector?.('select');
          if (sel){
            const opt = [...sel.options].find(o => (o.textContent||'').trim() === label);
            if (opt){ sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles:true })); return true; }
          }
        }
      }
      return false;
    };
    const byId = ()=>{
      const sel = modal.querySelector('select[id*="cfop" i], select[name*="cfop" i]');
      if (!sel) return false;
      const opt = [...sel.options].find(o => (o.textContent||'').trim() === label);
      if (opt){ sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles:true })); return true; }
      return false;
    };
    return byLabel() || byId();
  }, label);
  if (okSelect) return true;

  try {
    const combo = page.getByRole('combobox', { name:/cfop/i });
    if (await combo.count()) { await combo.click(); await page.getByRole('option', { name: label }).click(); return true; }
  } catch {}
  try { await page.getByText(label, { exact: true }).click(); return true; } catch {}
  return false;
}

async function esperarModalProdutoFechar(page, timeoutMs=5*60*1000){
  const handle = await page.evaluateHandle(() => {
    const vis = el => el && el.offsetParent !== null;
    return [...document.querySelectorAll('.modal.show, [role="dialog"]')].filter(vis).at(-1) || null;
  });
  if (!handle) return true;
  const ok = await page.waitForFunction((el)=> !el || !document.body.contains(el) || el.offsetParent===null, handle, { timeout: timeoutMs }).catch(()=>false);
  try { await handle.dispose(); } catch {}
  return !!ok;
}

/* ======== Modal de seleção (reaproveitado se a tabela mudar) ======== */
async function modalSelecaoProdutos(page, itens) {
  await page.evaluate((itens) => {
    const old = document.getElementById('__rpa_lista'); if (old) old.remove();
    const overlay = document.createElement('div');
    overlay.id = '__rpa_lista';
    Object.assign(overlay.style, { position:'fixed', inset:0, background:'rgba(0,0,0,.45)', display:'flex', alignItems:'center', justifyContent:'center', zIndex:2147483647 });
    const box = document.createElement('div');
    box.style.cssText = 'background:#fff;width:min(980px,96vw);max-height:min(86vh,900px);padding:18px;border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.2);font-family:system-ui,Segoe UI,Roboto,sans-serif;display:flex;flex-direction:column';
    const head = document.createElement('div'); head.style.cssText = 'display:flex;justify-content:space-between;align-items:center;margin-bottom:10px';
    const h = document.createElement('h2'); h.textContent = 'Selecione os produtos para editar'; h.style.cssText = 'font-size:18px;margin:0';
    const markAll = document.createElement('button'); markAll.textContent = 'Marcar/Desmarcar todos'; markAll.style.cssText = 'padding:6px 10px;border-radius:8px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
    const wrap = document.createElement('div'); wrap.style.cssText = 'overflow:auto;border:1px solid #e5e7eb;border-radius:10px';
    const table = document.createElement('table'); table.style.cssText = 'width:100%;border-collapse:separate;border-spacing:0';
    const thead = document.createElement('thead'); const trh = document.createElement('tr');
    ['✓','Nome','Cód. Vinc.','Descr. Vinc.','CFOP'].forEach(txt => {
      const th = document.createElement('th'); th.textContent = txt;
      th.style.cssText = 'position:sticky;top:0;background:#f9fafb;text-align:left;padding:8px;border-bottom:1px solid #e5e7eb';
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    const tbody = document.createElement('tbody');
    itens.forEach((it) => {
      const tr = document.createElement('tr'); tr.style.cssText = 'border-bottom:1px solid #f1f5f9';
      const td0 = document.createElement('td'); td0.style.cssText = 'padding:6px 8px';
      const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = false; cb.dataset.idx = String(it.idx);
      td0.appendChild(cb); tr.appendChild(td0);
      [it.nome, it.codv, it.descv, it.cfop].forEach(v => { const td = document.createElement('td'); td.textContent = v || ''; td.style.cssText = 'padding:6px 8px'; tr.appendChild(td); });
      tbody.appendChild(tr);
    });
    table.append(thead, tbody); wrap.appendChild(table);
    const foot = document.createElement('div'); foot.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;margin-top:10px';
    const btnCancel = document.createElement('button'); btnCancel.textContent = 'Cancelar'; btnCancel.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #d0d7de;background:#f6f8fa;cursor:pointer';
    const btnOk = document.createElement('button'); btnOk.textContent = 'Continuar'; btnOk.style.cssText = 'padding:8px 14px;border-radius:10px;border:1px solid #2563eb;background:#2563eb;color:#fff;cursor:pointer';
    markAll.onclick = () => { const boxes = tbody.querySelectorAll('input[type="checkbox"]'); const anyOff = Array.from(boxes).some(b => !b.checked); boxes.forEach(b => b.checked = anyOff); };
    btnCancel.onclick = () => { window.__rpaSel = null; overlay.remove(); };
    btnOk.onclick = () => { const selecionados = [...tbody.querySelectorAll('input[type="checkbox"]:checked')].map(cb => Number(cb.dataset.idx)); window.__rpaSel = selecionados; overlay.remove(); };
    head.append(h, markAll); box.append(head, wrap, foot); foot.append(btnCancel, btnOk);
    overlay.appendChild(box); document.body.appendChild(overlay);
  }, itens);
  const res = await page.waitForFunction(() => window.__rpaSel !== undefined, { timeout: 0 });
  const selecionados = await res.evaluate(() => window.__rpaSel);
  await page.evaluate(() => { delete window.__rpaSel; });
  return selecionados;
}

/* ======== Coletar lista e gerar hash ======== */
async function coletarListaProdutos(page){
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
  if (!info || !info.itens?.length) return { itens: [], hash: '' };
  const hash = crypto.createHash('sha1').update(JSON.stringify(info.itens.map(i => [i.nome, i.codv, i.descv, i.cfop]))).digest('hex');
  return { itens: info.itens, hash };
}

/* ======== Fluxo principal ======== */
(async () => {
  // Lê seleção anterior
  let selecaoAnterior = null;
  try { selecaoAnterior = JSON.parse(fs.readFileSync('selecionados.json','utf8')); } catch {}

  if (!selecaoAnterior?.selecionados?.length) {
    console.log('selecionados.json não encontrado ou vazio. Rode o Passo 12 antes.');
    return;
  }
  const chave = (fs.readFileSync('chaves.txt','utf8').split(/\r?\n/).map(s=>s.trim()).filter(Boolean)[0]) || selecaoAnterior.chave;
  if (!chave){ console.log('Sem chave em chaves.txt e no selecionados.json'); return; }

  const browser = await chromium.launch({ headless:false, slowMo:80 });
  const context = await browser.newContext({ storageState:'auth.json' });
  const page = await context.newPage();
  page.on('dialog', async d => { await d.accept(); });

  // Abre a nota e vai para a aba de produtos
  const abriu = await abrirNotaPorChave(page, chave);
  if (!abriu){
    await modalMsg(page, { titulo:'Nota não aberta', mensagem:'Abra manualmente a nota (Importa Danfe) e clique Continuar.', botoes:[{label:'Continuar', primary:true}] });
  }
  let okAba = await irParaAbaProdutos(page);
  if (!okAba){
    await modalMsg(page, { titulo:'Aba de produtos', mensagem:'Clique você na aba "Dados do Produtos" e depois em Continuar.', botoes:[{label:'Continuar', primary:true}] });
  }
  await page.waitForTimeout(600);

  // Coleta lista atual e compara hash
  const { itens: itensAtuais, hash: hashAtual } = await coletarListaProdutos(page);
  if (!itensAtuais.length){
    await modalMsg(page, { titulo:'Sem itens', mensagem:'Não achei a tabela de produtos nesta nota.', botoes:[{label:'Ok', primary:true}] });
    await browser.close(); return;
  }

  if (hashAtual !== selecaoAnterior.tabela_hash || chave !== selecaoAnterior.chave) {
    // Lista mudou: reabrir seleção aqui mesmo
    const decisao = await modalMsg(page, {
      titulo: 'Lista alterada',
      mensagem: 'A lista de produtos desta nota mudou desde a seleção anterior.\nDeseja re-selecionar os itens agora?',
      botoes: [{label:'Cancelar'}, {label:'Re-selecionar agora', primary:true}]
    });
    if (decisao !== 1) { await browser.close(); return; }

    const novaSelecao = await modalSelecaoProdutos(page, itensAtuais);
    if (!novaSelecao) { await browser.close(); return; }

    selecaoAnterior = {
      chave,
      carimbo: new Date().toISOString(),
      total_itens: itensAtuais.length,
      tabela_hash: hashAtual,
      selecionados: novaSelecao
    };
    fs.writeFileSync('selecionados.json', JSON.stringify(selecaoAnterior, null, 2), 'utf8');
  }

  // Edita itens marcados
  const tabela = page.locator('table').filter({ has: page.locator('tbody tr') }).first();
  await tabela.waitFor({ state:'visible', timeout:15000 }).catch(()=>{});
  const linhas = tabela.locator('tbody tr');
  const totalLinhas = await linhas.count();

  for (let k=0; k<selecaoAnterior.selecionados.length; k++){
    const idx = selecaoAnterior.selecionados[k];
    if (idx >= totalLinhas) continue;

    const row = linhas.nth(idx);
    await row.scrollIntoViewIfNeeded();
    await row.dblclick().catch(async ()=>{ await row.click(); });

    await page.waitForTimeout(400);
    const infos = await coletarOpcoesCFOPnoModal(page);
    const escolha = await modalEscolherCFOP(page, { opcoes: infos.opcoes, atual: infos.atual });
    if (escolha.pular){
      await modalMsg(page, { titulo:'Item pulado', mensagem:`Item ${k+1}/${selecaoAnterior.selecionados.length} pulado.`, botoes:[{label:'Ok', primary:true}] });
      await fecharModaisSistema(page, 2);
      continue;
    }
    if (escolha.valor){
      const aplicado = await aplicarCFOPnoModal(page, escolha.valor);
      if (!aplicado){
        await modalMsg(page, { titulo:'Selecione manualmente', mensagem:'Não consegui aplicar o CFOP sozinho. Selecione manualmente no modal do sistema e clique em Continuar.', botoes:[{label:'Continuar', primary:true}] });
      }
    }
    await modalMsg(page, { titulo:'Confirme no sistema', mensagem:'Preencha "Codigo: *" (se faltar) e clique no botão **Confirmar** do modal do sistema. A automação espera o modal fechar.', botoes:[{label:'Ok', primary:true}] });
    await esperarModalProdutoFechar(page);
    await page.waitForTimeout(300);
  }

  const dec = await modalMsg(page, { titulo:'Finalizar', mensagem:'Todos os itens selecionados foram processados.\nDeseja clicar em **Executar Entrada** agora?', botoes:[{label:'Não'}, {label:'Sim', primary:true}] });
  if (dec === 1){
    const btn = page.getByRole('button', { name: /^Executar Entrada$/i });
    if (await btn.count()){
      await Promise.allSettled([ page.waitForLoadState('networkidle'), btn.click() ]);
    }
  }

  await browser.close();
})();

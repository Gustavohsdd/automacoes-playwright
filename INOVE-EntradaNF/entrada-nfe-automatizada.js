/**
 * ===============================================================
 *  BLOCO 0 — IMPORTS, CONFIG BÁSICA E UTILITÁRIOS
 *  - Playwright
 *  - fs para ler Chaves.txt
 *  - readline para pausas e inputs no terminal (sem libs extras)
 * ===============================================================
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// ===== CLR com RGB (Truecolor) e compatibilidade com API antiga =====
globalThis.CLR ??= (() => {
  const ESC = '\x1b[';
  const reset = `${ESC}0m`;
  const bold  = `${ESC}1m`;

  // Detecção de profundidade e fallback automático
  const depth = (() => {
    try { return process.stdout?.getColorDepth?.() ?? 8; } catch { return 8; }
  })();
  const truecolor = depth >= 24 || process.env.COLORTERM === 'truecolor';

  // Fallback p/ 8/16 cores
  const BASIC = [
    { code: 30, rgb: [0,0,0] }, { code: 31, rgb: [205,49,49] },
    { code: 32, rgb: [13,188,121] }, { code: 33, rgb: [229,229,16] },
    { code: 34, rgb: [36,114,200] }, { code: 35, rgb: [188,63,188] },
    { code: 36, rgb: [17,168,205] }, { code: 37, rgb: [229,229,229] },
    { code: 90, rgb: [128,128,128] },
  ];
  const nearestBasicCode = (r,g,b) => {
    let best = BASIC[0], bestD = Infinity;
    for (const c of BASIC) {
      const dr = r - c.rgb[0], dg = g - c.rgb[1], db = b - c.rgb[2];
      const d = dr*dr + dg*dg + db*db;
      if (d < bestD) { bestD = d; best = c; }
    }
    return best.code;
  };

  const fg = (r,g,b) => truecolor ? `${ESC}38;2;${r};${g};${b}m` : `${ESC}${nearestBasicCode(r,g,b)}m`;
  const bg = (r,g,b) => {
    if (truecolor) return `${ESC}48;2;${r};${g};${b}m`;
    const c = nearestBasicCode(r,g,b);
    return `${ESC}${(c >= 90 ? c + 10 : c + 10)}m`; // 30–37->40–47, 90–97->100–107
  };
  const wrap = (txt, r,g,b, opts = {}) =>
    `${opts.bg ? bg(r,g,b) : fg(r,g,b)}${opts.bold ? bold : ''}${txt}${reset}`;

  // === SUA PALETA CENTRAL (troque só estes números para mudar tudo) ===
  // helper: HEX -> [r,g,b]
const hexToRgb = (hex) => {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  if (!m) throw new Error('HEX inválido: ' + hex);
  return [parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16)];
};

// edite aqui com o picker do VS Code (clique na cor)
const PALETTE_HEX = {
  cyan:    '#00B4FF', // perguntas
  green:   '#00C800', // sucesso
  red:     '#DC1E1E', // erro
  yellow:  '#FFC800', // atenção
  blue:    '#5078FF',
  magenta: '#f1a009',
  gray:    '#969696',
};

// converte p/ a paleta que o CLR usa internamente
const PALETTE = Object.fromEntries(
  Object.entries(PALETTE_HEX).map(([k, hex]) => [k, hexToRgb(hex)])
);

  // Gera propriedades compatíveis: CLR.cyan, CLR.red, etc. (foreground)
  const named = Object.fromEntries(
    Object.entries(PALETTE).map(([k, [r,g,b]]) => [k, fg(r,g,b)])
  );
  // Versões de fundo: CLR.bgCyan, CLR.bgRed, etc.
  const bgNamed = Object.fromEntries(
    Object.entries(PALETTE).map(([k, [r,g,b]]) => [`bg${k[0].toUpperCase()+k.slice(1)}`, bg(r,g,b)])
  );

  return { reset, bold, fg, bg, wrap, ...named, ...bgNamed };
})();

const color = {
  ask: (s) => `${CLR.cyan}${s}${CLR.reset}`,        // perguntas (digitação)
  ok: (s) => `${CLR.green}${s}${CLR.reset}`,        // sucesso
  err: (s) => `${CLR.red}${s}${CLR.reset}`,         // erro
  warn: (s) => `${CLR.yellow}${s}${CLR.reset}`,     // analisar navegador / atenção
  info: (s) => `${CLR.blue}${s}${CLR.reset}`,       // informativo
  section: (s) => `${CLR.magenta}${s}${CLR.reset}`, // seções
};

//============================== FIM DO ESQUEMA DE CORES ==============================

// ===== PATCH CMD-SAFE (somente saída no terminal; não mexe na lógica) =====
const EOL = process.platform === 'win32' ? '\r\n' : '\n';
const USE_ASCII = process.platform === 'win32' && !process.env.FORCE_UNICODE;

// Substitui caracteres que bagunçam o CMD por equivalentes ASCII
function sanitizeForCmd(x) {
  if (typeof x !== 'string') return x;
  let s = x;

  if (USE_ASCII) {
    s = s
      // linhas e box-drawing
      .replace(/[\u2500\u2501\u2502\u250c\u2510\u2514\u2518\u251c\u2524\u252c\u2534\u253c]/g, '-')
      // travessões e reticências
      .replace(/—|–/g, '-')
      .replace(/…/g, '...')
      // aspas tipográficas
      .replace(/[“”]/g, '"')
      .replace(/[‘’]/g, "'")
      // setas, marcadores e símbolos
      .replace(/↳/g, '->')
      .replace(/►/g, '>')
      .replace(/•/g, '*')
      // check/warn/x
      .replace(/✔/g, '[OK]')
      .replace(/✖/g, '[X]')
      .replace(/⚠/g, '[!]');
    }

  // normaliza quebras de linha para o Windows (evita duplicidade visual)
  s = s.replace(/\r?\n/g, EOL);
  return s;
}
// Monkey-patch: toda saída passa pelo sanitizador
const _log = console.log.bind(console);
const _error = console.error.bind(console);
console.log = (...args) => _log(...args.map(sanitizeForCmd));
console.error = (...args) => _error(...args.map(sanitizeForCmd));

// Ajuste aqui se quiser manter fixo no código (ou deixe vazio para perguntar no runtime)
const CRED = {
  urlLogin: 'https://araujopatrocinio.inovautomacao.com.br/login',
  urlHome: 'https://araujopatrocinio.inovautomacao.com.br',
  codigo: '86',    // se vazio, pergunta no terminal
  senha: '120718',     // se vazio, pergunta no terminal
  empresaValor: '1', // value do select (ex.: "1")
};

// helper de prompt no terminal (perguntas sempre em CIANO)
function prompt(question) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const q = question.endsWith(' ') ? question : `${question} `;
  return new Promise(res => rl.question(color.ask(q), answer => { rl.close(); res(answer); }));
}

// pausa só apertar Enter
function pause(msg = 'Pressione Enter para continuar...') {
  return prompt(`\n${msg}`);
}

// leitura de chaves no arquivo "Chaves.txt" (mesma pasta do script)
function lerChaves() {
  const p = path.join(process.cwd(), 'Chaves.txt');
  if (!fs.existsSync(p)) {
    throw new Error(`Arquivo "Chaves.txt" não encontrado na pasta atual: ${process.cwd()}`);
  }
  const raw = fs.readFileSync(p, 'utf8');
  return raw.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
}

// imprime tabela simples no terminal (Cód Vinc / Descr / CFOP)
function printTabelaProdutosBasica(produtos) {
  console.log('\nLista de produtos (índice  ->  Cód.Vinc | Descr.Vinc | CFOP)\n');
  produtos.forEach((p, i) => {
    const cod = (p.codVinc ?? '').toString().padEnd(10, ' ');
    const desc = (p.descrVinc ?? '').toString().padEnd(40, ' ');
    const cfop = (p.cfop ?? '').toString().padEnd(6, ' ');
    console.log(`${String(i).padStart(3, ' ')} -> ${cod} | ${desc} | ${cfop}`);
  });
  console.log('');
}

// espera botão por texto (regex) e clica com fallback por contains
async function clickByText(page, textRegex, { timeout = 10000 } = {}) {
  const btn = page.getByRole('button', { name: textRegex });
  await btn.waitFor({ timeout });
  await btn.click();
}

// espera qualquer modal do sistema sumir (tentativas genéricas)
async function waitModalFechar(page, { timeout = 30000 } = {}) {
  try {
    await page.getByRole('button', { name: /Confirmar/i }).waitFor({ state: 'detached', timeout });
  } catch (_) {
    await page.waitForTimeout(800); // grace period
  }
}

// duplo clique estável em célula por texto
async function dblclickCellByText(page, texto) {
  const cell = page.getByRole('cell', { name: texto });
  await cell.waitFor();
  const box = await cell.boundingBox();
  if (!box) throw new Error(`Não encontrei a célula para "${texto}"`);
  await cell.dblclick();
}

// --- HELPER: estou logado?
async function isOnLogin(page) {
  const url = page.url() || '';
  if (/\/login/i.test(url)) return true;
  const hasCodigo = await page.getByRole('textbox', { name: /Codigo/i }).count();
  const hasSenha = await page.getByRole('textbox', { name: /Senha/i }).count();
  return (hasCodigo && hasSenha);
}

// --- HELPER: estou na tela de PESQUISA (Importação NF-e)?
async function isOnPesquisaImportacao(page) {
  const temChave = await page.getByRole('textbox', { name: /Chave/i }).count();
  const temPesq = await page.getByRole('button', { name: /Pesquisar/i }).count();
  return !!(temChave && temPesq);
}


/**
 * ===============================================================
 *  BLOCO 1 — LOGIN
 * ===============================================================
 */
async function fazerLogin(page) {
  const codigo = CRED.codigo || await prompt('Digite seu CÓDIGO de acesso:');
  const senha = CRED.senha || await prompt('Digite sua SENHA:');

  await page.goto(CRED.urlLogin, { waitUntil: 'domcontentloaded' });

  await page.getByRole('textbox', { name: 'Codigo' }).fill(codigo);
  await page.getByRole('textbox', { name: 'Senha' }).fill(senha);

  const empresaSelect = page.getByLabel('Empresa *');
  await empresaSelect.waitFor();
  await empresaSelect.selectOption(CRED.empresaValor);

  await clickByText(page, /Entrar/i);
  await page.waitForLoadState('networkidle');
}

/**
 * ===============================================================
 *  BLOCO 2 — NAVEGAR: COMPRA → IMPORTAÇÃO NF-e (usado só se precisar)
 * ===============================================================
 */
async function irParaImportacao(page) {
  if (await isOnPesquisaImportacao(page)) return;

  const compra = page.getByText(/^Compra$/i).first();
  if (await compra.count()) {
    await compra.hover().catch(() => { });
    await compra.click().catch(() => { });
    await page.waitForTimeout(300);
  }

  let importacaoNFe = page.locator('p:has-text("Importação NF-e")').first();
  if (!(await importacaoNFe.count())) {
    importacaoNFe = page.getByRole('link', { name: /Importação NF-e/i }).first();
  }
  if (!(await importacaoNFe.count())) {
    importacaoNFe = page.getByText(/Importação NF-e/i).first();
  }

  await importacaoNFe.waitFor({ timeout: 15000 });
  await importacaoNFe.click();
  await page.waitForLoadState('networkidle');
}


/**
 * ===============================================================
 *  BLOCO 3 — PESQUISAR NOTA PELA CHAVE (validação XML/Entrada)
 * ===============================================================
 */
async function pesquisarEExecutarEntradaInicial(page, chave) {
  try { await clickByText(page, /Dados de Pesquisa/i); } catch (_) { }

  await page.getByRole('textbox', { name: /Chave/i }).fill(chave);
  await clickByText(page, /Pesquisar/i);
  await page.waitForLoadState('networkidle');

  const norm = (s) => (s || '')
    .toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const row = page.locator('table tbody tr').filter({
    has: page.locator('td', { hasText: chave })
  }).first();

  await row.waitFor({ timeout: 15000 });

  const table = row.locator('xpath=ancestor::table[1]');
  const headers = (await table.locator('thead th').allTextContents())
    .map(t => (t || '').replace(/\s+/g, ' ').trim());

  let idxXML = headers.findIndex(h => norm(h).includes('xml'));
  let idxEntrada = headers.findIndex(h => norm(h).includes('entrada'));

  const tdsCount = await row.locator('td').count();
  if (idxXML < 0 && tdsCount >= 2) idxXML = tdsCount - 2;
  if (idxEntrada < 0 && tdsCount >= 1) idxEntrada = tdsCount - 1;

  const tds = row.locator('td');
  const xmlTxt = (idxXML >= 0) ? (await tds.nth(idxXML).innerText()).trim() : '';
  const entTxt = (idxEntrada >= 0) ? (await tds.nth(idxEntrada).innerText()).trim() : '';

  const xmlOK = /sim/i.test(xmlTxt);
  const entOK = /(não|nao)/i.test(entTxt);

  if (!xmlOK || !entOK) {
    console.log('\n' + color.err('não é possivel executar a entrada desta nota no momento, tente novamente depois'));
    await prompt('Deseja prosseguir para a próxima nota? (Pressione Enter para continuar)');
    throw new Error('PULAR_NOTA');
  }

  await row.click();

  await clickByText(page, /Executar Entrada/i);
  await page.waitForLoadState('networkidle');

  await Promise.race([
    page.waitForSelector('input#Operacao', { state: 'attached' }),
    page.getByLabel(/Operaç[aã]o\s*\*/i).waitFor({ state: 'attached' }),
    page.waitForSelector('input[name*="Operacao"], input[aria-label*="Opera"], input[placeholder*="Opera"]', { state: 'attached' }),
  ]).catch(() => { });
  await page.waitForTimeout(250);

  console.log('\n' + color.section('----------------------------------------------------------------------'));
  console.log(color.section('PASSO: Operação e CFOP (1ª aba) — configurando pelo CMD'));
  console.log(color.section('----------------------------------------------------------------------'));

  // -------- Operacao --------
  let operacaoInput = page.locator('input#Operacao').first();
  if (!(await operacaoInput.count())) operacaoInput = page.getByLabel(/Operaç[aã]o\s*\*/i).first();
  if (!(await operacaoInput.count())) {
    operacaoInput = page.locator('input[placeholder*="Opera"], input[aria-label*="Opera"], input[name*="Operacao"]').first();
  }

  if (await operacaoInput.count()) {
    await operacaoInput.waitFor({ timeout: 15000 });
    const codigoOperacao = await prompt('Digite o CÓDIGO de Operação (#Operacao):');
    if (codigoOperacao?.trim()) {
      await operacaoInput.fill(codigoOperacao.trim());
      await operacaoInput.blur().catch(() => { });
      await page.keyboard.press('Tab').catch(() => { });
      await page.waitForTimeout(700);
      console.log(color.ok('✔ Código de Operação aplicado.'));
    } else {
      console.log(color.info('[AVISO] Código de Operação vazio. Mantendo valor atual.'));
    }
  } else {
    console.log(color.warn('[AVISO] Campo "#Operacao" não localizado.'));
  }

  // -------- CFOP (select dependente) --------
  let cfopSelect = page.locator('select#Cfop').first();
  if (!(await cfopSelect.count())) cfopSelect = page.getByLabel(/CFOP\s*\*/i).first();
  if (!(await cfopSelect.count())) cfopSelect = page.locator('select[name*="Cfop"], select[id*="Cfop"]').first();

  if (await cfopSelect.count()) {
    await cfopSelect.waitFor({ timeout: 15000 });

    for (let i = 0; i < 50; i++) {
      if (!(await cfopSelect.isDisabled().catch(() => false))) break;
      await page.waitForTimeout(120);
    }

    const opts = await cfopSelect.locator('option').all();
    const opcoes = [];
    for (const o of opts) {
      const v = (await o.getAttribute('value')) || '';
      if (!v) continue;
      const t = (await o.textContent())?.trim() || v;
      opcoes.push({ value: v, label: t });
    }

    if (opcoes.length) {
      console.log('\nOpções de CFOP (1ª aba):');
      opcoes.forEach((op, i) => console.log(`${i} → ${op.label} (value=${op.value})`));
      const idxStr = await prompt('Escolha o ÍNDICE da CFOP (1ª aba):');
      const idx = parseInt(idxStr, 10);
      if (Number.isFinite(idx) && idx >= 0 && idx < opcoes.length) {
        await cfopSelect.selectOption(opcoes[idx].value);
        console.log(color.ok(`CFOP selecionada (1ª aba): ${opcoes[idx].label} (value=${opcoes[idx].value})`));
      } else {
        console.log(color.info('Índice inválido. Mantendo CFOP atual na 1ª aba.'));
      }
    } else {
      console.log(color.warn('[AVISO] Nenhuma opção de CFOP encontrada no select da 1ª aba.'));
    }
  } else {
    console.log(color.warn('[AVISO] Campo "#Cfop" não localizado.'));
  }

  console.log(color.section('----------------------------------------------------------------------'));
  console.log(color.section('FIM do passo Operação/CFOP (1ª aba). Prosseguindo…'));
  console.log(color.section('----------------------------------------------------------------------\n'));

  await page.waitForTimeout(100);
}


/**
 * ===============================================================
 *  BLOCO 4 — IR PARA A ABA "Dados do Produtos" (robusto)
 * ===============================================================
 */
async function irParaAbaDadosDoProdutos(page) {
  const tentativas = [
    () => page.getByRole('tab', { name: /Dados do Produto[s]?/i }).first(),
    () => page.getByRole('button', { name: /Dados do Produto[s]?/i }).first(),
    () => page.getByRole('link', { name: /Dados do Produto[s]?/i }).first(),
    () => page.locator('a,button,li,span,div').filter({ hasText: /Dados do Produto[s]?/i }).first()
  ];

  let clicou = false;
  for (const mk of tentativas) {
    const loc = mk();
    if (await loc.count()) {
      await loc.click({ timeout: 15000 }).catch(() => { });
      clicou = true;
      break;
    }
  }

  if (!clicou) throw new Error('Não encontrei a aba "Dados do Produtos" para clicar.');

  await page.waitForSelector('table tbody tr', { state: 'visible', timeout: 20000 });
  await page.waitForTimeout(300);
}


/**
 * ===============================================================
 *  BLOCO 5 — LISTAR PRODUTOS NA TABELA E PERGUNTAR QUAIS EDITAR
 *  (RESPONSIVO À LARGURA DO TERMINAL + [VAZIO] em vermelho/negrito)
 *  (sem redeclarar CLR/color → usa aliases _CLR e _color)
 * ===============================================================
 */

// aliases locais (NÃO redeclaram as globais do projeto)
const _CLR   = (globalThis.CLR   ?? { reset:'\x1b[0m', bold:'\x1b[1m', red:'\x1b[31m' });
const _color = (globalThis.color ?? { section:(s)=>s, warn:(s)=>s });

// Larguras alvo e mínimas por coluna
const _PREFW = { IDX:3, DESC:44, CUSTO:10, QTDE:6, TOTAL:14, FATOR:6, COD:10, DESCVINC:32, CFOP:6 };
const _MINW  = { IDX:3, DESC:22, CUSTO:7,  QTDE:4, TOTAL:10, FATOR:4, COD:6,  DESCVINC:16, CFOP:3  };

// separador entre colunas
const _SEP = ' | ';
const _SEP_COUNT = 8;
const _SEP_LEN = _SEP.length * _SEP_COUNT;

// objeto mutável com as larguras em uso
let _W = { ..._PREFW };

/* ------------------- utils de padding/normalização ------------------- */
function _padVal(v, w) {
  let s = (v ?? '');
  if (!s || String(s).trim() === '') s = '[VAZIO]';
  s = String(s);
  if (s.length > w) return s.slice(0, w - 1) + '…';
  return s + ' '.repeat(w - s.length);
}
function _padLeft(v, w) {
  const s = String(v ?? '');
  if (s.length > w) return s.slice(-w);
  return ' '.repeat(w - s.length) + s;
}
function _normLabel(s) {
  return (s || '')
    .toString()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[.\-_:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
function _isDesc(lbl){ const n=_normLabel(lbl); return n.includes('descr') && !n.includes('vinc'); }
function _isCusto(lbl){ return _normLabel(lbl).includes('custo'); }
function _isQtde(lbl){ const n=_normLabel(lbl); return n.includes('qtde') || n.includes('quant'); }
function _isTotal(lbl){ return _normLabel(lbl).includes('total'); }
function _isFator(lbl){ return _normLabel(lbl).includes('fator'); }
function _isCodVinc(lbl){ const n=_normLabel(lbl); return n.includes('cod') && n.includes('vinc'); }
function _isDescrVinc(lbl){ const n=_normLabel(lbl); return n.includes('descr') && n.includes('vinc'); }
function _isCFOP(lbl){ return _normLabel(lbl).includes('cfop'); }

/* ------------------- largura do terminal e linhas cheias ------------------- */
function _termCols() {
  const c = (typeof process !== 'undefined' && process.stdout && Number.isFinite(process.stdout.columns))
    ? process.stdout.columns
    : 120;
  return Math.max(40, c); // nunca menos que 40
}
function _makeLine(ch='=', margin=1) {
  // mesma largura para "=" e "-" — margin=1 evita quebra por última coluna
  const w = _termCols() - margin;
  return ch.repeat(w);
}

/* ------------------- cálculo de larguras das colunas ------------------- */
function _totalWidth(W) {
  return W.IDX + W.DESC + W.CUSTO + W.QTDE + W.TOTAL + W.FATOR + W.COD + W.DESCVINC + W.CFOP + _SEP_LEN;
}
function _recalcWidths() {
  const target = _termCols() - 1; // margem de segurança
  let W = { ..._PREFW };

  if (_totalWidth(W) <= target) { _W = W; return; }

  const reduce = (key, n) => {
    const min = _MINW[key];
    const can = Math.max(0, W[key] - min);
    const dec = Math.min(can, n);
    W[key] -= dec;
    return dec;
  };

  let guard = 0;
  while (_totalWidth(W) > target && guard++ < 1000) {
    let sobra = _totalWidth(W) - target;

    if (sobra > 0) sobra -= reduce('DESC', Math.ceil(sobra * 0.6));
    if (sobra > 0) sobra -= reduce('DESCVINC', Math.ceil(sobra * 0.4));

    const ordem = ['TOTAL','CUSTO','COD','FATOR','QTDE','CFOP'];
    for (const k of ordem) {
      if (sobra <= 0) break;
      sobra -= reduce(k, Math.min(2, sobra));
    }
    if (sobra > 0) {
      const all = ['DESC','DESCVINC','TOTAL','CUSTO','COD','FATOR','QTDE','CFOP'];
      for (const k of all) {
        if (sobra <= 0) break;
        sobra -= reduce(k, 1);
      }
    }
  }
  _W = W;
}

/* ------------------- localizar tabela e extrair produtos ------------------- */
async function _encontrarTabelaProdutos(page) {
  const tabelas = await page.locator('table:has(thead):has(tbody tr)').all();
  for (const t of tabelas) {
    const ths = t.locator('thead th');
    const cab = (await ths.allTextContents()).map(x => x.replace(/\s+/g, ' ').trim());
    const idx = {
      descricao: cab.findIndex(_isDesc),
      custo: cab.findIndex(_isCusto),
      qtde: cab.findIndex(_isQtde),
      total: cab.findIndex(_isTotal),
      fator: cab.findIndex(_isFator),
      codVinc: cab.findIndex(_isCodVinc),
      descrVinc: cab.findIndex(_isDescrVinc),
      cfop: cab.findIndex(_isCFOP),
    };
    const hits = Object.values(idx).filter(i => i >= 0).length;
    if (idx.descricao >= 0 && idx.cfop >= 0 && hits >= 5) return { tabela: t, idx };
  }
  return { tabela: null, idx: null };
}

async function coletarProdutosDaTabela(page) {
  const { tabela, idx } = await _encontrarTabelaProdutos(page);
  if (!tabela || !idx) {
    const cells = await page.locator('table tbody td').allTextContents().catch(() => []);
    const linhasCruas = (cells || []).map(x => x.trim()).filter(Boolean);
    return { produtos: [], linhasCruas };
  }
  const linhas = await tabela.locator('tbody tr:has(td)').all();
  const produtos = [];

  for (const r of linhas) {
    const tds = r.locator('td');
    const read = async (i) => {
      if (i == null || i < 0) return '';
      const txt = await tds.nth(i).innerText().catch(() => '');
      return (txt || '').replace(/\s+/g, ' ').trim();
    };
    const item = {
      descricao: await read(idx.descricao),
      custo:     await read(idx.custo),
      qtde:      await read(idx.qtde),
      total:     await read(idx.total),
      fator:     await read(idx.fator),
      codVinc:   await read(idx.codVinc),
      descrVinc: await read(idx.descrVinc),
      cfop:      await read(idx.cfop),
      _row: r
    };
    const temAlgo = Object.values(item).some(v => (typeof v === 'string' ? v.trim() !== '' : false));
    if (temAlgo) produtos.push(item);
  }

  if (produtos.length === 0) {
    const cells = await tabela.locator('tbody td').allTextContents().catch(() => []);
    const linhasCruas = (cells || []).map(x => x.trim()).filter(Boolean);
    return { produtos: [], linhasCruas };
  }
  return { produtos, linhasCruas: null };
}

/* ------------------- cabeçalho ------------------- */
function _buildHeader() {
  _recalcWidths();
  const parts = [
    _padVal('idx',       _W.IDX),
    _padVal('Descrição', _W.DESC),
    _padVal('Custo',     _W.CUSTO),
    _padVal('Qtde',      _W.QTDE),
    _padVal('Total',     _W.TOTAL),
    _padVal('Fator',     _W.FATOR),
    _padVal('Cód.Vinc',  _W.COD),
    _padVal('Descr.Vinc',_W.DESCVINC),
    _padVal('CFOP',      _W.CFOP),
  ];
  return parts.join(_SEP);
}

/* ------------------- destaque de [VAZIO] ------------------- */
function _colorizeVazio(str) {
  return String(str).replace(/\[VAZIO\]/g, `${_CLR.bold}${_CLR.red}[VAZIO]${_CLR.reset}`);
}

/* ------------------- impressão da tabela ------------------- */
function printTabelaProdutos(produtos) {
  console.log(_color.section('PRODUTOS — Resumo para conferência\n'));

  const headerStr = _buildHeader();

  // Agora as duas linhas usam a MESMA largura do terminal
  const LINE_EQ   = _makeLine('=');
  const LINE_DASH = _makeLine('-');

  console.log(LINE_EQ);
  console.log(`${_CLR.bold}${headerStr}${_CLR.reset}`);
  console.log(LINE_EQ);

  produtos.forEach((p, i) => {
    let linha = [
      _padLeft(i,           _W.IDX),
      _padVal(p.descricao,  _W.DESC),
      _padVal(p.custo,      _W.CUSTO),
      _padVal(p.qtde,       _W.QTDE),
      _padVal(p.total,      _W.TOTAL),
      _padVal(p.fator,      _W.FATOR),
      _padVal(p.codVinc,    _W.COD),
      _padVal(p.descrVinc,  _W.DESCVINC),
      _padVal(p.cfop,       _W.CFOP),
    ].join(_SEP);

    linha = _colorizeVazio(linha);
    console.log(linha);
    console.log(LINE_DASH); // separador entre produtos — mesmo comprimento da linha de "="
  });

  console.log(LINE_EQ);
  console.log('');
}

/* ------------------- prompt de seleção ------------------- */
async function perguntarIndicesParaEditar(produtos, linhasCruas) {
  console.log('\n' + _color.section('------------------------------------------------------------'));
  console.log(_color.section('PRODUTOS — Resumo para conferência'));
  console.log(_color.section('------------------------------------------------------------'));

  if (linhasCruas) {
    console.log(_color.warn('[MODO BRUTO] Não consegui mapear as colunas do grid.'));
    console.log(linhasCruas.join('\n'));
    console.log('Digite partes do texto para eu localizar e abrir edição.');
    const keys = await prompt('Termos (separados por |):');
    const termos = keys.split('|').map(s => s.trim()).filter(Boolean);
    console.log(_color.section('------------------------------------------------------------\n'));
    return { modo: 'texto', termos };
  }

  printTabelaProdutos(produtos);

  const entrada = await prompt('\nDigite os ÍNDICES dos produtos a editar (ex.: 0,2,5) ou deixe vazio para nenhum:');
  const indices = entrada
    .split(',')
    .map(s => s.trim())
    .filter(Boolean)
    .map(n => parseInt(n, 10))
    .filter(n => Number.isFinite(n) && n >= 0 && n < produtos.length);

  console.log(_color.section('------------------------------------------------------------\n'));
  return { modo: 'indice', indices };
}


/**
 * ===============================================================
 *  BLOCO 6 — EDITAR PRODUTOS SELECIONADOS (Codigo → Fator → CFOP)
 *  Ajustes: perguntas inline (cursor na mesma linha) com ask()
 *           + fator “modo calculadora” (unidade ou kg) com TAB.
 * ===============================================================
 */
async function editarProdutosSelecionados(page, produtos, selecao) {
  // ---- Helper de pergunta inline (cursor na mesma linha) ----
  const tintAsk = (s) => (typeof color?.ask === 'function' ? color.ask(s)
                    : (typeof color?.info === 'function' ? color.info(s) : s));
  async function ask(msg) {
    // imprime a mensagem colorida direto no prompt para ficar inline
    const resp = await prompt(tintAsk(String(msg)) + ' ');
    return (resp ?? '').trim();
  }

  async function abrirModalDoItem(item) {
    if (!item?._row) throw new Error('Linha do produto (_row) não encontrada para abrir o modal.');
    const row = item._row;
    await row.waitFor({ timeout: 15000 });
    await row.scrollIntoViewIfNeeded().catch(() => {});
    try { await row.dblclick(); }
    catch {
      const firstCell = row.locator('td').first();
      await firstCell.waitFor({ timeout: 10000 });
      await firstCell.dblclick();
    }
    await page.waitForSelector('[role="dialog"], .modal.show, .modal-dialog', {
      state: 'visible',
      timeout: 15000
    });
  }

  let listaParaEditar = [];
  if (selecao.modo === 'indice') {
    listaParaEditar = selecao.indices.map(i => produtos[i]).filter(Boolean);
  } else {
    for (const termo of selecao.termos) {
      const hit = produtos.find(p => (p.descrVinc || '').toUpperCase().includes(termo.toUpperCase()));
      if (hit) listaParaEditar.push(hit);
    }
  }

  const lerValorSeguro = async (inp) => {
    try { return (await inp.inputValue())?.trim() ?? ''; } catch { return ''; }
  };

  // ---------- Helpers de input mascarado ----------
  async function limparCampoMascarado(inp, page) {
    try { await inp.click({ force: true }); } catch {}
    try { await inp.focus(); } catch {}
    try { await page.keyboard.press('Control+A'); } catch {}
    for (let i = 0; i < 8; i++) {
      try { await page.keyboard.press('Backspace'); } catch {}
      await page.waitForTimeout(20);
    }
    for (let i = 0; i < 2; i++) {
      try { await page.keyboard.press('Delete'); } catch {}
      await page.waitForTimeout(10);
    }
    await page.waitForTimeout(80);
  }

  async function digitarDigitos(inp, page, digs) {
    for (const ch of digs) {
      await inp.type(ch, { delay: 30 });
      await page.waitForTimeout(18);
    }
  }

  // ---------- Conversões ----------
  function normalizaDecimal3(txt) {
    const t = (txt || '').toString().trim().replace(/\./g, '').replace(',', '.');
    if (!/^\d+(\.\d+)?$/.test(t)) return null;
    const n = Number(t);
    if (Number.isNaN(n)) return null;
    return n.toFixed(3).replace('.', ',');
  }

  // "2,555" → "2555"; "0,500" → "500"; "12" → "12"
  function textoParaDigitos(txt) {
    const s = (txt || '').toString().trim();
    if (s.includes(',')) {
      const dec3 = normalizaDecimal3(s);
      if (!dec3) return null;
      const [i, d] = dec3.split(',');
      return `${i.replace(/[^\d]/g, '')}${(d || '').replace(/[^\d]/g, '')}`;
    }
    if (!/^\d+$/.test(s)) return null; // unidade = só dígitos
    return s;
  }

  function digitosParaDecimal3(digs) {
    const d = (digs || '0').replace(/[^\d]/g, '');
    if (d.length <= 3) return `0,${d.padStart(3, '0')}`;
    const int = d.slice(0, -3).replace(/^0+(?=\d)/, '');
    const dec = d.slice(-3);
    return `${int || '0'},${dec}`;
  }

  const norm = (s) => (s || '').toString().trim().replace(/\./g, '');

  // ---------- Tentativas de escrita ----------
  async function tentarModoDecimal(inp, page, digs) {
    await limparCampoMascarado(inp, page);
    await digitarDigitos(inp, page, digs);      // sem vírgula — máscara desloca
    try { await page.keyboard.press('Tab'); } catch {}
    await page.waitForTimeout(160);
    const exibido = await lerValorSeguro(inp);
    const esperado = digitosParaDecimal3(digs);
    return { ok: norm(exibido) === norm(esperado), exibido, esperado, modo: 'decimal' };
  }

  async function tentarModoInteiro(inp, page, digs) {
    await limparCampoMascarado(inp, page);
    await digitarDigitos(inp, page, digs);      // inteiro, sem vírgula
    try { await page.keyboard.press('Tab'); } catch {}
    await page.waitForTimeout(160);
    const exibido = await lerValorSeguro(inp);
    const esperadoTrim = digs.replace(/^0+(?=\d)/, '');
    const exibidoTrim = exibido.replace(/^0+(?=\d)/, '');
    return { ok: norm(exibidoTrim) === norm(esperadoTrim), exibido, esperado: esperadoTrim, modo: 'inteiro' };
  }

  async function setFatorInteligente(modal, page, textoUsuario) {
    let inp = modal.getByRole('textbox', { name: /Fator/i }).first();
    if (!(await inp.count())) inp = modal.getByLabel(/Fator/i).first();
    if (!(await inp.count())) inp = modal.locator('input#Fator, input[name*="Fator"]').first();
    if (!(await inp.count())) throw new Error('Campo "Fator" não localizado.');

    const raw = (textoUsuario || '').trim();
    const digs = textoParaDigitos(raw);
    if (!digs) throw new Error('Entrada inválida para Fator.');

    // Se usuário usou vírgula → tenta decimal primeiro; senão inteiro primeiro.
    if (raw.includes(',')) {
      const r1 = await tentarModoDecimal(inp, page, digs);
      if (r1.ok) return r1;
      const r2 = await tentarModoInteiro(inp, page, digs);
      return r2.ok ? { ...r2, fallback: 'inteiro' } : r2;
    } else {
      const r1 = await tentarModoInteiro(inp, page, digs);
      if (r1.ok) return r1;
      const r2 = await tentarModoDecimal(inp, page, digs);
      return r2.ok ? { ...r2, fallback: 'decimal' } : r2;
    }
  }

  // --------- Pergunta do Fator (inline) ---------
  const pedirFatorValido = async (valorAtual) => {
    while (true) {
      const entrada = await ask(
        `Fator Atual (deixe em branco para manter): ${valorAtual || '[VAZIO]'}`
      );
      if (!entrada) return null; // manter
      const raw = entrada.replace(/\s+/g, '');
      if (/^\d+$/.test(raw)) return raw;                 // unidade
      if (/^\d+,\d+$/.test(raw.replace('.', ','))) return normalizaDecimal3(raw); // kg
      console.log(color.err('Valor inválido. Use apenas dígitos (unidade) ou decimal com vírgula (3 casas).'));
    }
  };

  for (const item of listaParaEditar) {
    const nomeLog = item.descrVinc || item.descricao || item.codVinc || 'ITEM';
    console.log(`\n>> ${color.warn('Abrindo modal do produto para análise/edição:')} ${nomeLog}`);
    await abrirModalDoItem(item);

    const modal = page.locator('[role="dialog"], .modal.show, .modal-dialog').last();
    await modal.waitFor({ timeout: 15000 });

    // === 1) CODIGO =========================================================
    let inpCodigo = modal.getByRole('textbox', { name: /Codigo/i }).first();
    if (!(await inpCodigo.count())) inpCodigo = modal.getByLabel(/Codigo/i).first();
    if (!(await inpCodigo.count())) inpCodigo = modal.locator('input#Codigo, input[name*="Codigo"]').first();

    if (await inpCodigo.count()) {
      const atual = await lerValorSeguro(inpCodigo);
      const novo = await ask(`Codigo atual: ${atual || '[VAZIO]'} — Digite novo Codigo (Enter para manter):`);
      if (novo) {
        await inpCodigo.fill(novo);
        try { await inpCodigo.focus(); } catch {}
        try { await inpCodigo.press('Enter'); } catch { await page.keyboard.press('Enter').catch(() => {}); }
        await page.waitForTimeout(150);
        console.log(color.ok(`✔ Codigo atualizado e confirmado (Enter): ${novo}`));
      } else {
        console.log(color.info('-> Codigo mantido.'));
      }
    } else {
      console.log(color.warn('[!] Campo "Codigo" não localizado neste modal.'));
    }

    // === 2) FATOR — Unidade (inteiro) ou Kg (decimal com vírgula) =========
    let inpFator = modal.getByRole('textbox', { name: /Fator/i }).first();
    if (!(await inpFator.count())) inpFator = modal.getByLabel(/Fator/i).first();
    if (!(await inpFator.count())) inpFator = modal.locator('input#Fator, input[name*="Fator"]').first();

    if (await inpFator.count()) {
      const atual = await lerValorSeguro(inpFator);
      const novoFator = await pedirFatorValido(atual);
      if (novoFator !== null) {
        try {
          const r = await setFatorInteligente(modal, page, novoFator);
          if (r.ok) {
            if (r.fallback === 'decimal') {
              console.log(color.warn(`⚠ UI não aceitou inteiro; aplicado como decimal (kg). Exibido: "${r.exibido}"`));
            } else if (r.fallback === 'inteiro') {
              console.log(color.warn(`⚠ UI não aceitou decimal; aplicado como inteiro (unidade). Exibido: "${r.exibido}"`));
            } else {
              const labelModo = r.modo === 'inteiro' ? 'unidade (inteiro)' : 'kg (3 casas)';
              console.log(color.ok(`✔ Fator atualizado (${labelModo}) e confirmado (Tab). Exibido: "${r.exibido}"`));
            }
          } else {
            console.log(color.err(`✖ Não foi possível validar o Fator. Exibido: "${r.exibido}" | Esperado: "${r.esperado}"`));
          }
        } catch (e) {
          console.log(color.err(`Falha ao definir Fator: ${e.message}`));
        }
      } else {
        console.log(color.info('-> Fator mantido.'));
      }
    } else {
      console.log(color.warn('[!] Campo "Fator" não localizado neste modal.'));
    }

    // === 3) CFOP ===========================================================
    let sel = modal.getByLabel(/CFOP:\s*\*/i).first();
    if (!(await sel.count())) sel = modal.locator('select#CfopCodigo').first();
    if (!(await sel.count())) sel = modal.locator('select').filter({ hasText: /CFOP/i }).first();

    if (await sel.count()) {
      await sel.waitFor({ timeout: 15000 });
      const opts = await sel.locator('option').all();
      const opcoes = [];
      for (const o of opts) {
        const v = (await o.getAttribute('value')) || '';
        const t = (await o.textContent())?.trim() || v;
        if (v) opcoes.push({ value: v, label: t });
      }

      console.log('\nOpções de CFOP:');
      opcoes.forEach((op, i) => console.log(`${i}: ${op.label} (value=${op.value})`));
      const idxStr = await ask('ÍNDICE da CFOP para este produto:');
      const idx = parseInt(idxStr, 10);
      if (Number.isFinite(idx) && idx >= 0 && idx < opcoes.length) {
        await sel.selectOption(opcoes[idx].value);
        console.log(color.ok(`✔ CFOP selecionada: ${opcoes[idx].label}`));
      } else {
        console.log(color.info('Índice inválido. Mantendo CFOP atual.'));
      }
    } else {
      console.log(color.warn('[!] Campo "CFOP" não localizado neste modal.'));
    }

    // === 4) CONFIRMAR AUTOMÁTICO ==========================================
    let btnConfirmar = modal.getByRole('button', { name: /Confirmar/i }).first();
    if (!(await btnConfirmar.count())) btnConfirmar = modal.getByRole('button', { name: /Salvar/i }).first();

    if (await btnConfirmar.count()) {
      await btnConfirmar.click().catch(() => {});
      await waitModalFechar(page);
      console.log(color.ok('✔ Modal confirmado e fechado.'));
    } else {
      console.log(color.warn('[!] Botão "Confirmar/Salvar" não encontrado; feche manualmente.'));
      await pause('Após fechar o modal no sistema, aperte Enter aqui...');
      await waitModalFechar(page);
    }
  }
}


/**
 * ===============================================================
 *  BLOCO 7 — EXECUTAR ENTRADA (SEGUNDO BOTÃO) E TRÍPLICE MODAL
 * ===============================================================
 */
async function executarEntradaFinal(page) {
  await clickByText(page, /Executar Entrada/i);
  await page.waitForTimeout(300);

  // Modal 1 (OK automático)
  try {
    await clickByText(page, /OK/i, { timeout: 8000 });
  } catch { }

  // Modal 2 — instruções (amarelo) + prompt (ciano)
  console.log('\n-------------------------------------------------------------------------------');
  console.log('\n------------------------------ ANALISE DE MARKUP ------------------------------');
  console.log('\n-------------------------------------------------------------------------------');
  console.log(color.warn('Revise/edite MANUALMENTE este modal no sistema.'));
  console.log(color.warn('Quando terminar a conferência no navegador, volte ao CMD:'));
  await pause('Pressione Enter para o ROBÔ clicar "Confirmar" no modal 2...');

  // tentar localizar o modal 2 visível (com botão Confirmar/Salvar)
  let modal2 = null;
  try {
    await page.waitForSelector('[role="dialog"], .modal.show, .modal-dialog', {
      state: 'visible',
      timeout: 15000
    });

    const candidatos = page
      .locator('[role="dialog"].show, .modal.show, .modal-dialog')
      .filter({ has: page.getByRole('button', { name: /Confirmar|Salvar/i }) });

    if (await candidatos.count()) {
      modal2 = candidatos.last();
    } else {
      modal2 = page.locator('[role="dialog"], .modal.show, .modal-dialog').last();
    }
  } catch { }

  let btnConfirmar = null;
  if (modal2) {
    btnConfirmar = modal2.getByRole('button', { name: /Confirmar/i }).first();
    if (!(await btnConfirmar.count())) {
      btnConfirmar = modal2.getByRole('button', { name: /Salvar/i }).first();
    }
  } else {
    btnConfirmar = page.getByRole('button', { name: /Confirmar/i }).last();
    if (!(await btnConfirmar.count())) {
      btnConfirmar = page.getByRole('button', { name: /Salvar/i }).last();
    }
  }

  if (btnConfirmar && (await btnConfirmar.count())) {
    await btnConfirmar.click().catch(() => { });
    await waitModalFechar(page);
    console.log(color.ok('✔ Modal 2 confirmado pelo robô.'));
  } else {
    console.log(color.warn('[!] Não localizei o botão "Confirmar/Salvar" no modal 2.'));
    console.log(color.warn('   Se já confirmou manualmente no navegador, continue.'));
    await pause('Após confirmar manualmente no sistema, aperte Enter aqui...');
    await waitModalFechar(page);
  }

  // Modal 3 (clicar "Não" sempre)
  try {
    await clickByText(page, /Não/i, { timeout: 8000 });
  } catch { }
}


/**
 * ===============================================================
 *  BLOCO 8 — PROCESSAR UMA CHAVE COMPLETA (PIPELINE)
 * ===============================================================
 */
async function processarChave(page, chave) {
  try {
    console.log(color.section(`\n==============================\nProcessando CHAVE: ${chave}\n==============================`));
    await pesquisarEExecutarEntradaInicial(page, chave);
    await irParaAbaDadosDoProdutos(page);

    const { produtos, linhasCruas } = await coletarProdutosDaTabela(page);
    const selecao = await perguntarIndicesParaEditar(produtos, linhasCruas);
    await editarProdutosSelecionados(page, produtos, selecao);
    await executarEntradaFinal(page);

    console.log('\n' + color.ok(`[OK] Chave ${chave} finalizada com sucesso.`));
    return { sucesso: true };
  } catch (err) {
    if (err && /PULAR_NOTA/.test(err.message)) {
      console.log(color.info('[INFO] Nota pulada: condições "XML=Sim" e "Entrada=Não" não atendidas. Indo para a próxima…'));
      return { sucesso: false, erro: 'PULADA' };
    }
    console.error(color.err(`[ERRO] Falhou na chave ${chave}: ${err.message}`));
    return { sucesso: false, erro: err.message };
  }
}


/**
 * ===============================================================
 *  BLOCO 9 — ROTINA PRINCIPAL (MAIN)
 * ===============================================================
 */
(async () => {
  const browser = await chromium.launch({ headless: false }); // visível
  const context = await browser.newContext();
  const page = await context.newPage();

  // LOGIN
  await fazerLogin(page);

  // COMPRA → IMPORTAÇÃO
  await irParaImportacao(page);

  // Ler chaves do arquivo
  const chaves = lerChaves();
  if (chaves.length === 0) {
    console.log(color.info('Nenhuma chave encontrada em "Chaves.txt". Encerrando.'));
    await browser.close();
    process.exit(0);
  }

  const resultados = [];

  for (const chave of chaves) {
    if (await isOnLogin(page)) {
      console.log('\n' + color.info('[INFO] Sessão expirada/na página de login. Fazendo login novamente...'));
      await fazerLogin(page);
      await irParaImportacao(page);
    } else if (!(await isOnPesquisaImportacao(page))) {
      await irParaImportacao(page);
    }

    const r = await processarChave(page, chave);
    resultados.push({ chave, ...r });

    await page.waitForTimeout(400);
  }

  // RESUMO FINAL
  console.log('\n' + color.section('================ RESUMO FINAL ================'));
  const ok = resultados.filter(x => x.sucesso);
  const fail = resultados.filter(x => !x.sucesso);
  console.log(color.ok(`Sucesso: ${ok.length}`));
  ok.forEach(x => console.log(color.ok(`  - ${x.chave}`)));
  console.log(color.err(`Falhas: ${fail.length}`));
  fail.forEach(x => console.log(color.err(`  - ${x.chave} :: ${x.erro}`)));
  console.log(color.section('=============================================='));

  await context.close();
  await browser.close();
})();

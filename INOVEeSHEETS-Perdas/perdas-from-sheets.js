// INOVEeSHEETS-Perdas — perdas-from-sheets.js
// Pasta sugerida: C:\Users\gusta\Documents\AutomacoesPro\INOVEeSHEETS-Perdas
// Dependências: npm i playwright googleapis dotenv
// .env (exemplo):
// SHEET_ID=1IYKAp2XXTN4ktE4288kGMnzj41xO8TgeTpgobb71hGA
// SHEET_TAB=Produtos
// HEADLESS=false
// GOOGLE_APPLICATION_CREDENTIALS=C:\Users\gusta\Documents\AutomacoesPro\INOVEeSHEETS-Perdas\service-account.json
// URL_LOGIN=https://araujopatrocinio.inovautomacao.com.br/login
// URL_HOME=https://araujopatrocinio.inovautomacao.com.br
// USUARIO="seu usuario"
// SENHA="sua senha"
// EMPRESA_VALUE=1
// LOTE_GRAVAR=50
//
// Novidade: Pré-validação de GTIN (padrão: inicia com "2" e tem 13 dígitos). Itens inválidos não são lançados e são listados no relatório final.

require('dotenv').config();
const { chromium } = require('playwright');
const { google } = require('googleapis');

// ===== Cores ANSI (sem dependências externas) =====
const c = {
  red:    (s) => `\x1b[31m${s}\x1b[0m`,
  yellow: (s) => `\x1b[33m${s}\x1b[0m`,
  green:  (s) => `\x1b[32m${s}\x1b[0m`,
  cyan:   (s) => `\x1b[36m${s}\x1b[0m`,
  blue:   (s) => `\x1b[34m${s}\x1b[0m`,
  gray:   (s) => `\x1b[90m${s}\x1b[0m`,
  white:  (s) => `${s}`,
};

// ==========================
// CONFIG
// ==========================
const SHEET_ID   = process.env.SHEET_ID;
const SHEET_TAB  = process.env.SHEET_TAB || 'Produtos';
const HEADLESS   = String(process.env.HEADLESS || 'false').toLowerCase() === 'true';
const URL_LOGIN  = process.env.URL_LOGIN;
const URL_HOME   = process.env.URL_HOME;
const USUARIO    = process.env.USUARIO || '';
const SENHA      = process.env.SENHA   || '';
const EMPRESA_VALUE = String(process.env.EMPRESA_VALUE || '1');
const LOTE_GRAVAR   = parseInt(process.env.LOTE_GRAVAR || '50', 10);

// Ajustes finos de tempo (ms) após ações críticas
const WAIT_AFTER_GTIN_MS      = parseInt(process.env.WAIT_AFTER_GTIN_MS || '800', 10);   // após Enter no GTIN
const WAIT_AFTER_QTD_ENTER_MS = parseInt(process.env.WAIT_AFTER_QTD_ENTER_MS || '800', 10); // após Enter em Quantidade

// Esperas adicionais para telas lentas / pós-gravação
const INCLUIR_TIMEOUT_MS     = parseInt(process.env.INCLUIR_TIMEOUT_MS || '90000', 10);  // até 90s p/ “Incluir”
const INCLUIR_RETRY          = parseInt(process.env.INCLUIR_RETRY || '2', 10);           // tentativas de recuperar tela “Perdas”
const WAIT_AFTER_GRAVAR_MS   = parseInt(process.env.WAIT_AFTER_GRAVAR_MS || '3000', 10); // respiro após “Gravar”
const WAIT_AFTER_ABRIR_INCLUIR_MS = parseInt(process.env.WAIT_AFTER_ABRIR_INCLUIR_MS || '600', 10);

if (!SHEET_ID) throw new Error('Defina SHEET_ID no .env');
if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) throw new Error('Defina GOOGLE_APPLICATION_CREDENTIALS no .env');
if (!URL_LOGIN) throw new Error('Defina URL_LOGIN no .env');
if (!USUARIO || !SENHA) console.warn(c.yellow('AVISO: USUARIO/SENHA não definidos no .env. Se a tela pedir, preencha manualmente.'));

// ==========================
// HELPERS — Sheets
// ==========================
function normHeader(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[\s\t\n\r]+/g, ' ')
    .replace(/[\.:;,_-]+/g, '')
    .trim()
    .toLowerCase();
}

async function getSheets() {
  const auth = new google.auth.GoogleAuth({ scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  return google.sheets({ version: 'v4', auth });
}

function a1Col(n) { // 1->A
  let s = '';
  while (n > 0) {
    let m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function toISODate(cell) {
  if (cell === undefined || cell === null || cell === '') return '';
  if (typeof cell === 'number') {
    // Serial date do Sheets
    const base = new Date(1899, 11, 30);
    const ms = Math.round(cell * 24 * 60 * 60 * 1000);
    const d = new Date(base.getTime() + ms);
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    return `${yyyy}-${mm}-${dd}`;
  }
  const s = String(cell).trim();
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/aaaa
  if (m) {
    const dd = m[1].padStart(2,'0');
    const mm = m[2].padStart(2,'0');
    const yyyy = m[3];
    return `${yyyy}-${mm}-${dd}`;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s; // yyyy-mm-dd
  return s; // fallback
}

async function lerPlanilhaCompleta(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A:Z`,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });
  const values = res.data.values || [];
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0];
  const rows = values.slice(1).map((r, i) => ({ _rowNumber: i+2, _raw: r }));
  return { headers, rows };
}

function indices(headers) {
  const map = {};
  headers.forEach((h, i) => { map[normHeader(h)] = i; });
  function idx(...cands) {
    for (const c of cands) {
      const k = normHeader(c);
      if (map[k] !== undefined) return map[k];
    }
    // aproximação
    const want = normHeader(cands[0]);
    const entry = Object.entries(map).find(([k]) => k.includes(want));
    return entry ? entry[1] : -1;
  }
  return {
    gtin:   idx('Gtin'),
    data:   idx('DATA','Data'),
    status: idx('Status','Status.'),
  };
}

async function atualizarStatusEmLote(sheets, headers, rowNumbers, novoStatus) {
  if (!rowNumbers.length) return;
  const { status } = indices(headers);
  if (status < 0) throw new Error('Coluna "Status" não encontrada na planilha.');
  const colA1 = a1Col(status + 1);

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      valueInputOption: 'RAW',
      data: rowNumbers.map(rn => ({ range: `${SHEET_TAB}!${colA1}${rn}:${colA1}${rn}`, values: [[novoStatus]] }))
    }
  });
}

// ==========================
// HELPERS — GTIN
// ==========================
function isDigits(str) {
  return /^\d+$/.test(str);
}
function isValidGtinPattern(gtin) {
  // Requisito: começa com '2' e tem 13 dígitos
  return /^2\d{12}$/.test(gtin);
}
function reasonForInvalid(gtin) {
  if (!isDigits(gtin)) return 'contém caracteres não numéricos';
  if (gtin.length !== 13) return `tamanho != 13 (len=${gtin.length})`;
  if (!gtin.startsWith('2')) return 'não começa com 2';
  return 'padrão inválido';
}

// ==========================
// HELPERS — Playwright
// ==========================
async function fazerLogin(page) {
  await page.goto(URL_LOGIN, { waitUntil: 'domcontentloaded' });
  const codigo = page.getByRole('textbox', { name: 'Codigo' });
  const senha  = page.getByRole('textbox', { name: 'Senha' });
  await codigo.fill(USUARIO);
  await codigo.press('Tab');
  await senha.fill(SENHA);
  await senha.press('Tab');
  await page.getByLabel('Empresa *').press('ArrowDown').catch(() => {});
  await page.getByLabel('Empresa *').selectOption(EMPRESA_VALUE);
  await page.getByRole('button', { name: 'Entrar' }).click();
  await page.waitForLoadState('networkidle');
}

async function irParaPerdas(page) {
  const btnProduto = page.getByRole('button', { name: 'Produto' });
  await btnProduto.waitFor({ timeout: 15000 });
  await btnProduto.click();

  let alvoPerdas = page.getByRole('paragraph').filter({ hasText: /Perda[s]?/i }).first();
  if (!(await alvoPerdas.count())) alvoPerdas = page.getByText(/Perda[s]?/i).first();
  await alvoPerdas.click();
  await page.waitForLoadState('networkidle');
}

async function esperarTelaListaPerdas(page, timeoutMs = INCLUIR_TIMEOUT_MS) {
  const btn = page.getByRole('button', { name: 'Incluir' }).first();
  await btn.waitFor({ state: 'visible', timeout: timeoutMs });
  await page.waitForTimeout(200);
  if (await btn.isDisabled().catch(() => false)) {
    await page.waitForTimeout(800);
  }
}

async function abrirInclusaoComRetry(page) {
  for (let tent = 0; tent <= INCLUIR_RETRY; tent++) {
    try {
      await esperarTelaListaPerdas(page);
      const btnIncluir = page.getByRole('button', { name: 'Incluir' }).first();
      await btnIncluir.click();
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(WAIT_AFTER_ABRIR_INCLUIR_MS);
      return; // sucesso
    } catch (e) {
      if (tent === INCLUIR_RETRY) throw e;
      console.log(c.yellow(`Não localizei 'Incluir' a tempo (tentativa ${tent+1}). Reabrindo menu Perdas...`));
      await irParaPerdas(page);
    }
  }
}

async function preencherData(page, isoDate) {
  const dateInput = page.locator('input[type="date"]').first();
  if (await dateInput.count()) {
    await dateInput.fill(isoDate);
    return;
  }
  let dataBox = page.getByRole('textbox', { name: /Data/i }).first();
  if (!(await dataBox.count())) dataBox = page.getByLabel(/Data/i).first();
  if (await dataBox.count()) {
    const [yyyy, mm, dd] = isoDate.split('-');
    await dataBox.click({ clickCount: 3 });
    await page.keyboard.type(dd);
    await page.waitForTimeout(80);
    await page.keyboard.type(mm);
    await page.waitForTimeout(80);
    await page.keyboard.type(yyyy);
    await page.keyboard.press('Enter');
    return;
  }
  console.log(c.yellow('AVISO: campo de Data não encontrado; continue manualmente se necessário.'));
}

async function lançarProduto(page, gtin) {
  const codProd = page.getByRole('textbox', { name: 'Código Produto' }).first();
  await codProd.waitFor({ timeout: 15000 });
  await codProd.click();
  await codProd.fill(String(gtin));
  await codProd.press('Enter');

  await page.waitForTimeout(WAIT_AFTER_GTIN_MS);

  const qtd = page.getByRole('textbox', { name: 'Quantidade' }).first();
  if (await qtd.count()) {
    await qtd.press('Enter').catch(() => {});
  } else {
    await page.keyboard.press('Enter').catch(() => {});
  }
  await page.waitForTimeout(WAIT_AFTER_QTD_ENTER_MS);
}

async function gravarLote(page) {
  const btnGravar = page.getByRole('button', { name: 'Gravar' }).first();
  await btnGravar.waitFor({ timeout: 15000 });
  await btnGravar.click();

  await Promise.race([
    page.getByText(/sucesso|gravado|salvo/i).waitFor({ timeout: 60000 }).catch(() => {}),
    page.waitForLoadState('networkidle'),
    page.waitForTimeout(3000),
  ]);
}

// ==========================
// PIPELINE
// ==========================
(async () => {
  const sheets = await getSheets();
  const { headers, rows } = await lerPlanilhaCompleta(sheets);
  if (!headers.length) { console.log(c.red('Planilha vazia ou sem cabeçalho.')); return; }
  const { gtin: idxGtin, data: idxData, status: idxStatus } = indices(headers);
  if (idxGtin < 0)   throw new Error('Coluna "Gtin" não encontrada.');
  if (idxData < 0)   throw new Error('Coluna "DATA" não encontrada.');
  if (idxStatus < 0) throw new Error('Coluna "Status" não encontrada.');

  // Pega pendentes (Status em branco) com GTIN e DATA presentes
  const pendentes = rows
    .filter(r => !String(r._raw[idxStatus] ?? '').trim())
    .map(r => ({
      rowNumber: r._rowNumber,
      gtin: String(r._raw[idxGtin] ?? '').trim(),
      isoDate: toISODate(r._raw[idxData]),
    }))
    .filter(x => x.gtin && x.isoDate);

  if (!pendentes.length) {
    console.log(c.green('Nenhum item pendente (Status em branco).'));
    return;
  }

  console.log(c.cyan(`Encontrados ${pendentes.length} itens pendentes.`));

  // ===== Pré-análise de GTIN =====
  const invalidGtins = [];
  const itens = [];
  for (const it of pendentes) {
    const g = it.gtin;
    if (isValidGtinPattern(g)) {
      itens.push(it);
    } else {
      invalidGtins.push({
        rowNumber: it.rowNumber,
        gtin: g,
        isoDate: it.isoDate,
        reason: reasonForInvalid(g),
      });
    }
  }

  console.log(c.yellow(`→ ${itens.length} dentro do padrão (^2\\d{12}$).`));
  console.log(
    invalidGtins.length
      ? c.red(`→ ${invalidGtins.length} fora do padrão (não serão lançados).`)
      : c.green('→ 0 fora do padrão.')
  );

  // Se não houver itens válidos, não abre navegador; só imprime relatório e encerra
  if (!itens.length) {
    if (invalidGtins.length) {
      console.log(c.yellow('\nRelatório — GTIN fora do padrão (ignorados):'));
      for (const inv of invalidGtins) {
        console.log(
          c.red(`  Linha ${inv.rowNumber} — DATA ${inv.isoDate} — GTIN "${inv.gtin}" — Motivo: ${inv.reason}`)
        );
      }
    }
    console.log(c.cyan('\nFinalizado (nenhum GTIN válido para lançar).'));
    return;
  }

  // ===== Automação Web para itens válidos =====
  const browser = await chromium.launch({ headless: HEADLESS });
  const context = await browser.newContext();
  const page = await context.newPage();

  // 1) Login
  await fazerLogin(page);

  // 2) Produto → 3) Perdas
  await irParaPerdas(page);

  let processadosNoLote = [];
  let primeiroIsoDaLote = null;

  const fecharLote = async () => {
    if (!processadosNoLote.length) return;

    // Grava no sistema
    await gravarLote(page);

    // Respiro extra
    await page.waitForTimeout(WAIT_AFTER_GRAVAR_MS);

    // Garante que voltamos à tela com o botão Incluir
    try {
      await esperarTelaListaPerdas(page);
    } catch (_) {
      await irParaPerdas(page);
      await esperarTelaListaPerdas(page);
    }

    // Marca no Sheets
    await atualizarStatusEmLote(sheets, headers, processadosNoLote, 'Lançado');
    console.log(c.green(`✔ Lote gravado e ${processadosNoLote.length} item(ns) marcados como Lançado no Sheets.`));

    processadosNoLote = [];
    primeiroIsoDaLote = null;
  };

  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];

    // Abre inclusão e define data para o primeiro do lote (ou quando a data mudar)
    if (processadosNoLote.length === 0) {
      await abrirInclusaoComRetry(page);
      primeiroIsoDaLote = item.isoDate;
      await preencherData(page, item.isoDate);
      await page.waitForTimeout(150);
    }

    if (primeiroIsoDaLote && item.isoDate !== primeiroIsoDaLote) {
      console.log(c.yellow(`↪ Data mudou de ${primeiroIsoDaLote} para ${item.isoDate}. Gravando lote atual e abrindo novo.`));
      await fecharLote();
      await abrirInclusaoComRetry(page);
      primeiroIsoDaLote = item.isoDate;
      await preencherData(page, item.isoDate);
    }

    console.log(c.blue(`→ Lançando item ${i+1}/${itens.length} (linha ${item.rowNumber}) — DATA ${item.isoDate} — GTIN ${item.gtin}`));

    await lançarProduto(page, item.gtin);
    processadosNoLote.push(item.rowNumber);

    if (processadosNoLote.length >= LOTE_GRAVAR) {
      await fecharLote();
    }
  }

  // Fecha o último lote remanescente
  await fecharLote();

  await context.close();
  await browser.close();

  // ===== Relatório final de inválidos =====
  if (invalidGtins.length) {
    console.log(c.yellow('\nRelatório — GTIN fora do padrão (ignorados):'));
    for (const inv of invalidGtins) {
      console.log(
        c.red(`  Linha ${inv.rowNumber} — DATA ${inv.isoDate} — GTIN "${inv.gtin}" — Motivo: ${inv.reason}`)
      );
    }
  }

  console.log(c.cyan('\nFinalizado.'));
})().catch(err => {
  const rawMsg = (err && (err.message || err.status || err.code)) || '';
  const causeMsg = (err && err.cause && (err.cause.message || err.cause.status || err.cause.code)) || '';
  const full = `${rawMsg} ${causeMsg}`.toLowerCase();

  if (full.includes('sheets api has not been used') || full.includes('it is disabled') || full.includes('permission_denied')) {
    console.error(c.red(`\n[Erro] Google Sheets API não habilitada para este projeto GCP.`));
    console.error(`  Projeto (número): 613989212968`);
    console.error(`  Acesse e habilite: https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=613989212968`);
    console.error(`  Depois aguarde 1–3 minutos e rode novamente.`);
    console.error(`\nAlém disso, confira se a planilha foi compartilhada com o e-mail da Service Account (client_email no JSON).`);
    console.error(`Para ver o e-mail da SA rapidamente:`);
    console.error(`  node -e "console.log(require('./service-account.json').client_email)"`);
  } else if (full.includes('the caller does not have permission') || full.includes('permission')) {
    console.error(c.red(`\n[Erro] A Service Account não tem permissão na planilha.`));
    console.error(`Abra a planilha no Google Sheets → Compartilhar → adicione o e-mail da Service Account como Editor.`);
    console.error(`Dica para ver o e-mail da SA: node -e "console.log(require('./service-account.json').client_email)"`);
  } else {
    console.error(c.red('Falha geral:'), err);
  }
  process.exit(1);
});

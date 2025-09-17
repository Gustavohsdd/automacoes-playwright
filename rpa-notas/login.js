// rpa-notas/login.js
// Requisitos: npm i playwright dotenv
// Sessão persistente: usamos launchPersistentContext com um perfil salvo em ./.auth/chromium-profile

const path = require("path");
const fs = require("fs");
const { chromium } = require("playwright");

// Tenta ./.env e, se não existir, tenta ../.env
let loadedFrom = null;
for (const p of [path.join(__dirname, ".env"), path.join(__dirname, "../.env")]) {
  if (fs.existsSync(p)) {
    require("dotenv").config({ path: p });
    loadedFrom = p;
    break;
  }
}
if (!loadedFrom) {
  console.warn("⚠️ Arquivo .env não encontrado em ./ ou ../");
}

const URL_LOGIN = process.env.URL_LOGIN?.trim();
const USUARIO   = process.env.USUARIO?.replaceAll('"', '').trim();
const SENHA     = process.env.SENHA?.replaceAll('"', '').trim();
const EMPRESA   = process.env.EMPRESA?.replaceAll('"', '').trim(); // ex.: 1 - ARAUJO PANIFICADORA 24HORAS

if (!URL_LOGIN || !USUARIO || !SENHA || !EMPRESA) {
  console.error("❌ Faltam variáveis no .env (URL_LOGIN, USUARIO, SENHA, EMPRESA).");
  process.exit(1);
}

const AUTH_DIR = path.resolve(__dirname, "../.auth");
const PROFILE_DIR = path.join(AUTH_DIR, "chromium-profile");
const STORAGE_FILE = path.join(AUTH_DIR, "storageState.json");

fs.mkdirSync(AUTH_DIR, { recursive: true });

(async () => {
  // 1) Abre Chromium com PERFIL PERSISTENTE (cookies e sessão ficam salvos)
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    viewport: null,
    args: ["--start-maximized"],
  });

  const page = context.pages()[0] || await context.newPage();
  page.setDefaultTimeout(20000);

  try {
    // 2) Vai para a página de login
    await page.goto(URL_LOGIN, { waitUntil: "domcontentloaded" });

    // 3) Se já estiver logado (cookie salvo), não faz login de novo
    if (!page.url().includes("/login")) {
      console.log("✅ Já está logado. Sessão recuperada do perfil persistente.");
    } else {
      console.log("🔐 Efetuando login…");

      // Preenche Código e Senha pelos placeholders vistos no print
      await page.getByPlaceholder(/codigo/i).fill(USUARIO, { delay: 40 });
      await page.getByPlaceholder(/senha/i).fill(SENHA, { delay: 40 });

      // Seleciona Empresa (tenta <select>, senão usa combo customizado)
      await selectEmpresa(page, EMPRESA);

      // Clica Entrar e aguarda sair do /login
      await Promise.all([
        page.waitForNavigation({ waitUntil: "networkidle" }),
        page.getByRole("button", { name: /^Entrar$/i }).click()
      ]);

      if (page.url().includes("/login")) {
        throw new Error("Login não saiu da tela de /login (verifique credenciais/empresa).");
      }
      console.log("✅ Login realizado.");
    }

    // 4) Salva storageState (útil se quiser reusar com outro tipo de contexto)
    await context.storageState({ path: STORAGE_FILE });
    console.log(`💾 Estado de sessão salvo em: ${STORAGE_FILE}`);
    console.log("🟢 Sessão mantida. Deixe esta janela aberta enquanto usar o sistema.");
    console.log("   Quando você fechar o Chromium, a sessão continua salva no perfil para a próxima execução.");

    // Mantém o script vivo enquanto o navegador estiver aberto
    context.on("close", () => {
      console.log("👋 Chromium fechado. Encerrando script.");
      process.exit(0);
    });
    await new Promise(() => {}); // mantém rodando
  } catch (err) {
    // Captura evidência se falhar
    const ts = new Date().toISOString().replace(/[:.]/g, "-");
    const shot = path.resolve(__dirname, `../screenshots/login-error-${ts}.png`);
    try { await page.screenshot({ path: shot, fullPage: true }); } catch {}
    console.error("❌ Erro no login:", err.message);
    console.error("📷 Screenshot (se disponível):", shot);
    process.exit(1);
  }
})();

// ---------- helpers ----------
async function selectEmpresa(page, empresaStr) {
  const codigo = empresaStr.match(/^\s*(\d+)/)?.[1] || null;
  const rotulo = empresaStr.replace(/^\s*\d+\s*-\s*/, "").trim();

  // Tenta <select> nativo
  const selects = page.locator("select");
  if (await selects.count()) {
    try {
      await selects.first().selectOption({ label: empresaStr });
      return;
    } catch {}
    if (codigo) {
      try {
        await selects.first().selectOption({ value: codigo });
        return;
      } catch {}
    }
  }

  // Tenta combobox customizada (ex.: Angular/Material)
  try {
    const combo = page.getByRole("combobox").first();
    await combo.click();
    let opt = page.getByRole("option", { name: new RegExp(rotulo, "i") });
    if ((await opt.count()) === 0 && codigo) {
      opt = page.getByRole("option", { name: new RegExp(`^${codigo}\\b`, "i") });
    }
    await opt.first().click();
    return;
  } catch {}

  // Tenta por lista genérica no popup
  try {
    await page.getByText(new RegExp(rotulo, "i"), { exact: false }).first().click();
    return;
  } catch {}

  throw new Error(`Não consegui selecionar a empresa: "${empresaStr}"`);
}

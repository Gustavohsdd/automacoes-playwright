// Filename: RelatorioProdutosporHorario.js

//--------------------------------------------------------
// Bloco 1. Imports
//--------------------------------------------------------

import { chromium } from 'playwright';
import { google } from 'googleapis';
import { JWT } from 'google-auth-library';
import * as dotenv from 'dotenv';
import xlsx from 'xlsx';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import dayjs from 'dayjs';
import customParseFormat from 'dayjs/plugin/customParseFormat.js';
import isSameOrBefore from 'dayjs/plugin/isSameOrBefore.js';
import chalk from 'chalk';
// 'readline/promises' não é mais necessário, foi removido.

//--------------------------------------------------------
// Bloco 2. Configurações Iniciais
//--------------------------------------------------------

dotenv.config();
dayjs.extend(customParseFormat);
dayjs.extend(isSameOrBefore);

//--------------------------------------------------------
// Bloco Constantes (NOVO BLOCO)
//--------------------------------------------------------

// --- PRODUTOS ---
// Adicione ou remova os códigos de produto que devem ser processados
const PRODUTOS_PARA_PROCESSAR = [
  '2', // Pão Francês
  '274', // Pão de Queijo
  // '15', // Ex: Próximo produto
];

// --- Google Sheets ---
const SHEET_ID = process.env.SHEET_ID;
const REGISTRO_TAB = 'Registro';
const VALOR_TAB = 'Valor';
const QUANTIDADE_TAB = 'Quantidade';

// --- Automação ---
const BATCH_SIZE = 30; // Quantidade de dias para salvar em lote

// --- Seletores e Nomes de Botões (Playwright) ---

// Login
const NAME_TEXTBOX_CODIGO = 'Codigo';
const NAME_TEXTBOX_SENHA = 'Senha';
const LABEL_SELECT_EMPRESA = 'Empresa *';
const NAME_BUTTON_ENTRAR = 'Entrar';

// Navegação Menu
const NAME_BUTTON_VENDA = 'Venda';
const TEXT_LINK_COMPARATIVO = 'Comparativo';
const INDEX_LINK_COMPARATIVO = 1; // Usar o segundo link "Comparativo"

// Página de Relatório (Filtros)
const SELECTOR_ACCORDION_PRODUTO = '#mat-expansion-panel-header-2';
const NAME_BUTTON_INCLUIR = 'Incluir';
const NAME_BUTTON_DADOS_PESQUISA = 'Dados de Pesquisa Informações';
const NAME_TEXTBOX_CODIGO_MODAL = 'Código'; // 'Código' com acento, do modal
const NAME_BUTTON_PESQUISAR = 'Pesquisar';
const ROLE_CELL_PRODUTO = 'cell'; // para 'page.getByRole('cell', ...)'
const NAME_BUTTON_CONFIRMAR = 'Confirmar';

// Tipo de Arquivo
const SELECTOR_SELECT_TIPO_ARQUIVO = 'select[name="TipoArquivo"]';
const VALUE_SELECT_TIPO_ARQUIVO = '2'; // '2' = .xlsx

// Filtro de Data
const TEXT_INPUT_DATA_INICIAL = 'Inicial';
const SELECTOR_INPUT_DATA_INICIAL = '#DataInicial';
const TEXT_INPUT_DATA_FINAL = 'Final';
const SELECTOR_INPUT_DATA_FINAL = '#DataFinal';

// Gerar Relatório
const NAME_BUTTON_GERAR = 'Gerar';

// --- Parser do Relatório (Índices de Linhas/Colunas) ---
const INDEX_LINHA_PRODUTO = 6;     // Linha 7 no Excel (info do produto)
const INDEX_LINHA_QUANTIDADE = 7; // Linha 8 no Excel (dados de quantidade)
const INDEX_LINHA_VALOR = 8;      // Linha 9 no Excel (dados de valor)
const START_COL_HORA = 1;         // Coluna B (início dos dados 00:00)
const END_COL_HORA = 25;          // Coluna Y (fim dos dados 23:00)


//--------------------------------------------------------
// Bloco 3. Funções Google Sheets
//--------------------------------------------------------

/**
 * Autoriza a aplicação a usar a API do Google Sheets.
 */
async function authorizeGoogleSheets() {
  const credentialsPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;

  if (!credentialsPath) {
    console.error(chalk.red('Erro Fatal: Variável de ambiente GOOGLE_APPLICATION_CREDENTIALS não encontrada.'));
    console.error(chalk.yellow('Por favor, verifique se o seu arquivo .env está correto e contém o caminho para o service-account.json.'));
    throw new Error('Caminho das credenciais do Google ausente.');
  }

  let credentials;
  try {
    const fileContent = await fs.readFile(credentialsPath, 'utf-8');
    credentials = JSON.parse(fileContent);
  } catch (err) {
    console.error(chalk.red(`Erro Fatal: Não foi possível ler ou parsear o arquivo service-account.json em: ${credentialsPath}`));
    console.error(chalk.red(err.message));
    throw new Error('Falha ao carregar credenciais.');
  }

  const privateKey = credentials.private_key;
  const clientEmail = credentials.client_email;

  if (!privateKey || !clientEmail) {
    console.error(chalk.red('Erro Fatal: O arquivo service-account.json não contém os campos "private_key" ou "client_email".'));
    throw new Error('Arquivo de credenciais inválido.');
  }

  const jwtClient = new JWT({
    email: clientEmail,
    key: privateKey.replace(/\\n/g, '\n'),
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  await jwtClient.authorize();
  return google.sheets({ version: 'v4', auth: jwtClient });
}

/**
 * Busca dados de uma aba específica.
 */
async function getSheetsData(sheets, range) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range,
  });
  return response.data.values || [];
}

/**
 * Adiciona dados ao final de uma aba. (Req. 9.1)
 */
async function appendSheetsData(sheets, range, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED', // Para interpretar "," como decimal
    resource: { values },
  });
}

/**
 * Atualiza uma célula ou intervalo específico. (Req. 11.1)
 */
async function updateSheetsData(sheets, range, values) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED', // Para interpretar "," como decimal
    resource: { values },
  });
}

//--------------------------------------------------------
// Bloco 4. Funções Auxiliares da Automação
//--------------------------------------------------------

/**
 * Função 'promptForProductCode' REMOVIDA.
 */

/**
 * Verifica as abas 'Registro' E 'Valor' para encontrar a última data REAL processada. (Req. 7.1, 7.2)
 * Retorna o dia *seguinte* à última data registrada, ou 01/01/2022 se não houver registro.
 */
async function getNextDate(sheets, productCode) {
  let registroDate = null;
  let valorDate = null;
  const startDate = dayjs('2022-01-01', 'YYYY-MM-DD');

  try {
    // 1. Verifica a aba 'Registro' (nosso "checkpoint" oficial)
    const registroData = await getSheetsData(sheets, `${REGISTRO_TAB}!A:B`);
    const registro = registroData.find(row => row[0] === productCode);
    if (registro && registro[1]) {
      registroDate = dayjs(registro[1], 'DD/MM/YYYY');
    }
  } catch (e) {
    console.warn(chalk.yellow(`Aviso ao ler aba 'Registro': ${e.message}`));
  }

  try {
    // 2. Verifica a aba 'Valor' (os dados reais) para auto-correção
    const valorData = await getSheetsData(sheets, `'${VALOR_TAB}'!A:B`);
    const productDates = valorData
      .filter(row => row[1] === productCode && row[0]) // Filtra pelo código e se a data existe
      .map(row => dayjs(row[0], 'DD/MM/YYYY')) // Converte para datas
      .filter(date => date.isValid()); // Remove datas inválidas

    if (productDates.length > 0) {
      // Encontra a data mais recente
      valorDate = productDates.reduce((max, current) => current.isAfter(max) ? current : max, productDates[0]);
    }
  } catch (e) {
    console.warn(chalk.yellow(`Aviso ao ler aba 'Valor': ${e.message}`));
  }

  // 3. Compara as datas para encontrar a mais recente
  let lastDate = null;
  if (registroDate && valorDate) {
    lastDate = registroDate.isAfter(valorDate) ? registroDate : valorDate;
  } else {
    lastDate = registroDate || valorDate;
  }

  // 4. Decide a data de início
  if (!lastDate) {
    console.log(chalk.yellow(`Nenhum registro encontrado para ${productCode}. Iniciando em 01/01/2022.`));
    return startDate;
  }

  if (lastDate.isBefore(startDate)) {
    console.log(chalk.yellow(`Registro antigo encontrado (${lastDate.format('DD/MM/YYYY')}). Iniciando em 01/01/2022.`));
    return startDate;
  }

  const nextDate = lastDate.add(1, 'day');
  console.log(chalk.blue(`Último dado real encontrado em ${lastDate.format('DD/MM/YYYY')}. Processando a partir de ${nextDate.format('DD/MM/YYYY')}.`));
  return nextDate;
}


/**
 * Atualiza a aba 'Registro' com a última data processada. (Req. 11)
 */
async function updateRegistro(sheets, productCode, formattedDate) {
  const range = `${REGISTRO_TAB}!A:B`;
  const data = await getSheetsData(sheets, range);
  const rowIndex = data.findIndex(row => row[0] === productCode);

  if (rowIndex === -1) {
    // Não encontrou, adiciona nova linha
    await appendSheetsData(sheets, range, [[productCode, formattedDate]]);
  } else {
    // Encontrou, atualiza a data
    const cellToUpdate = `${REGISTRO_TAB}!B${rowIndex + 1}`;
    await updateSheetsData(sheets, cellToUpdate, [[formattedDate]]);
  }
}

/**
 * Processa o arquivo .xlsx baixado e formata os dados para a planilha. (Req. 9)
 */
async function parseDownloadedReport(filePath, formattedDate) {
  const cleanNumber = (str) => (str || '0').toString().replace(/\./g, '');
  const cleanCurrency = (str) => (str || '0').toString().replace('R$', '').replace(/\./g, '').trim();

  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  if (!data || data.length < 9 || !data[INDEX_LINHA_PRODUTO] || !data[INDEX_LINHA_VALOR]) {
    console.error(chalk.red(`  -> Erro no parser: O arquivo baixado (${filePath}) não contém dados ou é um arquivo de erro.`));
    console.error(chalk.red(`  -> Linhas encontradas: ${data ? data.length : 0}. Esperado: 9+.`));
    throw new Error('Arquivo de relatório inválido ou vazio. Provavelmente a sessão caiu.');
  }

  // Linhas baseadas nas constantes
  const productInfoLine = data[INDEX_LINHA_PRODUTO][0]; // "Produto: 2 - PAO FRANCES"
  const productInfo = productInfoLine.split(': ')[1] || ''; // "2 - PAO FRANCES"

  const [code, ...nameParts] = (productInfo || productInfoLine).split(' - ');
  const productName = nameParts.join(' - ');

  const quantidadeRow = data[INDEX_LINHA_QUANTIDADE];
  const valorRow = data[INDEX_LINHA_VALOR];

  // Colunas da Planilha: Data, Codigo, Produto, 00:00, ..., 23:00
  const baseRow = [formattedDate, code.trim(), productName.trim()];

  // Pega os 24 valores horários (baseado nas constantes)
  const quantidadeData = quantidadeRow.slice(START_COL_HORA, END_COL_HORA).map(cleanNumber);
  const valorData = valorRow.slice(START_COL_HORA, END_COL_HORA).map(cleanCurrency);

  const sheetRowQuantidade = [...baseRow, ...quantidadeData];
  const sheetRowValor = [...baseRow, ...valorData];

  return { sheetRowQuantidade, sheetRowValor, reportCode: code.trim() };
}

//--------------------------------------------------------
// Bloco 5. Função Principal (MODIFICADA PARA LOOP DE PRODUTOS)
//--------------------------------------------------------

(async () => {
  let browser;
  let context;
  let sheets;

  try {
    // 1. Inicialização (Executa uma vez)
    console.log(chalk.blue('Iniciando automação...'));
    sheets = await authorizeGoogleSheets();
    browser = await chromium.launch({ headless: (process.env.HEADLESS === 'true') });
    context = await browser.newContext({ acceptDownloads: true });

    // 2. Loop Principal de Produtos (NOVA ESTRUTURA)
    for (const productCode of PRODUTOS_PARA_PROCESSAR) {
      console.log(chalk.magenta.bold(`\n=================================================`));
      console.log(chalk.magenta.bold(`Iniciando processamento para o Produto: ${productCode}`));
      console.log(chalk.magenta.bold(`=================================================\n`));

      // --- Acumuladores e Função de Lote (definidos por produto) ---
      let allValorRows = [];
      let allQuantidadeRows = [];
      let lastProcessedDate = null;
      let lastProcessedReportCode = productCode; // Define um padrão
      let daysProcessedInBatch = 0;
      let page; // Definida aqui para ser acessível no finally

      // Função de salvar o lote (declarada dentro do loop para acessar os acumuladores)
      const saveBatchData = async () => {
        if (allValorRows.length === 0) {
          return; // Não faz nada se não houver dados
        }

        console.log(chalk.blue(`Salvando lote de ${allValorRows.length} dias para o produto ${lastProcessedReportCode}...`));
        try {
          console.log(chalk.blue(`Enviando ${allValorRows.length} registros para a aba '${VALOR_TAB}'...`));
          await appendSheetsData(sheets, `${VALOR_TAB}!A:AA`, allValorRows);

          console.log(chalk.blue(`Enviando ${allQuantidadeRows.length} registros para a aba '${QUANTIDADE_TAB}'...`));
          await appendSheetsData(sheets, `${QUANTIDADE_TAB}!A:AA`, allQuantidadeRows);

          console.log(chalk.blue(`Atualizando registro para o produto ${lastProcessedReportCode} com a data ${lastProcessedDate}...`));
          await updateRegistro(sheets, lastProcessedReportCode, lastProcessedDate);

          console.log(chalk.green('Lote salvo com sucesso na planilha!'));

        } catch (sheetsError) {
          console.error(chalk.red('Erro CRÍTICO ao salvar lote na planilha:'), sheetsError);
          console.error(chalk.yellow('Os dados deste lote NÃO foram salvos. A automação continuará, mas este lote pode precisar ser reprocessado.'));
        } finally {
          // Limpa os acumuladores para o próximo lote
          allValorRows = [];
          allQuantidadeRows = [];
          daysProcessedInBatch = 0;
        }
      };
      // --- Fim Acumuladores e Função ---

      try {
        // 3. Processamento por Produto
        page = await context.newPage();
        let dateToProcess = await getNextDate(sheets, productCode);
        const stopDate = dayjs().subtract(1, 'day'); // A automação para no dia anterior ao atual

        // 4. Iniciar Navegador e Fazer Login (por produto)
        console.log(chalk.blue('Iniciando nova página e fazendo login...'));
        await page.goto(process.env.URL_LOGIN);
        await page.getByRole('textbox', { name: NAME_TEXTBOX_CODIGO }).fill(process.env.USUARIO);
        await page.getByRole('textbox', { name: NAME_TEXTBOX_SENHA }).fill(process.env.SENHA);
        await page.getByLabel(LABEL_SELECT_EMPRESA).selectOption(process.env.EMPRESA_VALUE);
        await page.getByRole('button', { name: NAME_BUTTON_ENTRAR }).click();
        await page.waitForURL(process.env.URL_HOME);
        console.log(chalk.green('Login realizado com sucesso.'));

        // 5. Navegação Inicial e Filtros Permanentes (por produto)
        console.log(chalk.blue('Navegando para o relatório e aplicando filtros...'));
        await page.getByRole('button', { name: NAME_BUTTON_VENDA }).click();
        await page.getByText(TEXT_LINK_COMPARATIVO).nth(INDEX_LINK_COMPARATIVO).click();
        await page.waitForSelector(SELECTOR_ACCORDION_PRODUTO);

        // Filtro de Produto
        await page.locator(SELECTOR_ACCORDION_PRODUTO).click();
        await page.getByRole('button', { name: NAME_BUTTON_INCLUIR }).click();
        await page.getByRole('button', { name: NAME_BUTTON_DADOS_PESQUISA }).click();
        await page.getByRole('textbox', { name: NAME_TEXTBOX_CODIGO_MODAL }).click();
        await page.getByRole('textbox', { name: NAME_TEXTBOX_CODIGO_MODAL }).fill(productCode); // <-- USA O CÓDIGO DO LOOP
        await page.getByRole('button', { name: NAME_BUTTON_PESQUISAR }).click();
        await page.waitForSelector(`role=${ROLE_CELL_PRODUTO}[name="${productCode}"]`);
        await page.getByRole(ROLE_CELL_PRODUTO, { name: productCode }).click();
        await page.getByRole('button', { name: NAME_BUTTON_CONFIRMAR }).click();

        // Filtro de Tipo de Arquivo
        await page.locator(SELECTOR_SELECT_TIPO_ARQUIVO).selectOption(VALUE_SELECT_TIPO_ARQUIVO);
        console.log(chalk.green(`Filtros permanentes aplicados para o produto ${productCode}. Iniciando loop de datas...`));

        // 6. Loop Principal de Processamento de Relatórios (Datas)
        while (dateToProcess.isSameOrBefore(stopDate, 'day')) {
          const dateStr = dateToProcess.format('YYYY-MM-DD');
          const formattedDate = dateToProcess.format('DD/MM/YYYY');
          let tempFilePath = '';

          try {
            console.log(chalk.cyan(`Processando relatório para ${formattedDate} (Prod: ${productCode})...`));

            // Filtros de Data
            await page.locator('ia-input-list').filter({ hasText: TEXT_INPUT_DATA_INICIAL }).locator(SELECTOR_INPUT_DATA_INICIAL).fill(dateStr);
            await page.locator('ia-input-list').filter({ hasText: TEXT_INPUT_DATA_FINAL }).locator(SELECTOR_INPUT_DATA_FINAL).fill(dateStr);

            // Gerar e Baixar Relatório
            const downloadPromise = page.waitForEvent('download');
            await page.getByRole('button', { name: NAME_BUTTON_GERAR }).click();
            const download = await downloadPromise;

            tempFilePath = path.join(os.tmpdir(), `relatorio-${productCode}-${dateStr}-${Date.now()}.xlsx`);
            await download.saveAs(tempFilePath);
            console.log(chalk.blue(`Arquivo temporário salvo em: ${tempFilePath}`));

            // Processar Arquivo
            const { sheetRowQuantidade, sheetRowValor, reportCode } = await parseDownloadedReport(tempFilePath, formattedDate);

            allValorRows.push(sheetRowValor);
            allQuantidadeRows.push(sheetRowQuantidade);
            lastProcessedDate = formattedDate;
            lastProcessedReportCode = reportCode; // Atualiza com o código do relatório

            console.log(chalk.green(`Sucesso: Relatório de ${formattedDate} para o produto ${reportCode} processado.`));
            await fs.unlink(tempFilePath);
            dateToProcess = dateToProcess.add(1, 'day');
            daysProcessedInBatch++;

            // Salvar o lote se atingir o tamanho
            if (daysProcessedInBatch >= BATCH_SIZE) {
              console.log(chalk.cyan(`Atingido o tamanho do lote (${BATCH_SIZE}). Salvando dados...`));
              await saveBatchData();
            }

          } catch (loopError) {
            console.error(chalk.red(`Falha ao processar data ${formattedDate} (Prod: ${productCode}): ${loopError.message}`));
            if (tempFilePath) {
              try { await fs.unlink(tempFilePath); } catch (e) { } // Tenta limpar se falhar
            }

            console.log(chalk.yellow('Tentando recarregar a página e continuar no MESMO dia...'));

            // Tenta recarregar a página e reaplicar filtros
            try {
              await page.reload({ timeout: 15000 });
              console.log(chalk.yellow('Página recarregada. Reaplicando filtros permanentes...'));

              await page.waitForSelector(SELECTOR_ACCORDION_PRODUTO); // Espera a página carregar
              await page.locator(SELECTOR_ACCORDION_PRODUTO).click();
              await page.getByRole('button', { name: NAME_BUTTON_INCLUIR }).click();
              await page.getByRole('button', { name: NAME_BUTTON_DADOS_PESQUISA }).click();
              await page.getByRole('textbox', { name: NAME_TEXTBOX_CODIGO_MODAL }).click();
              await page.getByRole('textbox', { name: NAME_TEXTBOX_CODIGO_MODAL }).fill(productCode);
              await page.getByRole('button', { name: NAME_BUTTON_PESQUISAR }).click();
              await page.waitForSelector(`role=${ROLE_CELL_PRODUTO}[name="${productCode}"]`);
              await page.getByRole(ROLE_CELL_PRODUTO, { name: productCode }).click();
              await page.getByRole('button', { name: NAME_BUTTON_CONFIRMAR }).click();
              await page.locator(SELECTOR_SELECT_TIPO_ARQUIVO).selectOption(VALUE_SELECT_TIPO_ARQUIVO);
              
              console.log(chalk.green('Filtros permanentes reaplicados. Continuando do dia que falhou...'));

            } catch (reloadError) {
              console.error(chalk.red('Falha ao recarregar a página. A sessão pode ter expirado. Encerrando processamento DESTE PRODUTO.'), reloadError);
              console.log(chalk.yellow('Tentando salvar dados acumulados antes de pular para o próximo produto...'));
              await saveBatchData();
              break; // Sai do loop 'while' (datas) e vai para o próximo produto
            }
          }
        } // Fim do loop 'while' (datas)

        // 7. Envio em Lote Final (para dados restantes do produto)
        console.log(chalk.blue(`Processamento de dias concluído para ${productCode}. Salvando lote final...`));
        await saveBatchData();

      } catch (productError) {
        console.error(chalk.red(`Erro fatal ao processar o produto ${productCode}:`), productError);
        console.log(chalk.yellow('Tentando salvar dados parciais deste produto antes de continuar...'));
        await saveBatchData(); // Tenta salvar o que tem
        // O loop 'for' continuará para o próximo produto
      } finally {
        if (page) {
          await page.close();
        }
        console.log(chalk.magenta.bold(`--- Processamento do Produto ${productCode} finalizado ---`));
      }
    } // Fim do loop 'for' (produtos)

  } catch (error) {
    // Captura erros de inicialização (antes do loop de produtos)
    console.error(chalk.red('Erro fatal na inicialização da automação (Planilhas ou Navegador):'), error);
  } finally {
    if (browser) {
      await browser.close();
      console.log(chalk.blue('\nNavegador fechado. Automação concluída.'));
    }
  }
})();

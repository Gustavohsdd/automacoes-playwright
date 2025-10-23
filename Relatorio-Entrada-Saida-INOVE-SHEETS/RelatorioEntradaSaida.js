//--------------------------------------------------------------------------
// Bloco 1 - Importações
//--------------------------------------------------------------------------

import { chromium } from 'playwright-chromium';
import dotenv from 'dotenv';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import * as xlsx from 'xlsx'; // MODIFICADO: Importa a biblioteca 'xlsx'
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';
import { 
    format as formatDate, 
    parse as parseDate, 
    addDays,
    subDays,
    isBefore,
    isAfter,
    startOfToday,
    endOfMonth
} from 'date-fns';

//--------------------------------------------------------------------------
// Bloco 2 - Configuração Inicial
//--------------------------------------------------------------------------
dotenv.config();
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

//--------------------------------------------------------------------------
// Bloco 3 - Constantes e Seletores
//--------------------------------------------------------------------------

// --- Configurações da Planilha ---
const SERVICE_ACCOUNT_FILE = path.join(__dirname, 'service-account.json');
const GOOGLE_SHEET_ID = process.env.SHEET_ID_ENTRADA_SAIDA;
const SHEET_NAME_ENTRADA = 'Entrada';
const DATE_COLUMN_HEADER = 'Data Entrada'; // Nome exato da coluna na Planilha
const DEFAULT_START_DATE = '2022-02-01'; // Data (yyyy-MM-dd) para usar se a planilha estiver vazia

// --- Configurações de Download ---
const DOWNLOAD_PATH = path.join(__dirname, 'downloads');
// MODIFICADO: Alterado para .xlsx
const PRODUTOS_FILE_NAME = 'relatorio_produtos.xlsx';
const ENTRADA_FILE_NAME = 'relatorio_entrada.xlsx';

// --- Seletore do ERP INOVE (Baseado no seu passo-a-passo) ---
const INOVE_URL = 'https://araujopatrocinio.inovautomacao.com.br/login';
const SELECTORS = {
    // Login
    INPUT_CODIGO: "role=textbox[name='Codigo']",
    INPUT_SENHA: "role=textbox[name='Senha']",
    SELECT_EMPRESA: "Empresa *",
    BUTTON_ENTRAR: "role=button[name='Entrar']",
    
    // Navegação Produtos
    BUTTON_PRODUTO: "role=button[name='Produto']",
    LINK_PLANILHA_PRODUTO: "role=paragraph >> text='Planilha Produto'",
    BUTTON_EXPORTAR_PRODUTO: "role=button[name='Exportar']",
    
    // Navegação Entradas (Estes são usados na função de 'baixarRelatorioEntrada')
    // Nota: A lógica de navegação foi corrigida diretamente na função
    //       usando getByRole e getByText, conforme sua sugestão.
    BUTTON_COMPRA: "role=button[name='Compra']",
    LINK_ENTRADA: "role=button[name='Entrada'] >> nth=3", // Este seletor estava causando o erro
    
    // Filtros Entradas
    SELECT_TIPO_ENTRADA: 'select[name="Tipo"]',
    SELECT_TIPO_RELATORIO: 'select[name="TipoRelatorio"]',
    SELECT_TIPO_ARQUIVO: 'select[name="TipoArquivo"]',
    INPUT_DATA_INICIAL: "ia-input-list:has-text('Data Inicial') >> #DataInicial",
    INPUT_DATA_FINAL: "ia-input-list:has-text('Data Final') >> #DataFinal",
    BUTTON_GERAR_ENTRADA: "role=button[name='Gerar']",
};

/**
 * =================================================================================
 * [OBRIGATÓRIO] AJUSTE OS NOMES DAS COLUNAS AQUI
 *
 * Verifique os nomes exatos das colunas nos seus arquivos CSV baixados
 * e ajuste as strings aqui para que o script possa ler os dados corretamente.
 * =================================================================================
 */
const MAPEAMENTO_COLUNAS_CSV = {
    // Colunas do CSV 'relatorio_produtos.csv'
    PRODUTOS: {
        CODIGO: 'Codigo',
        FAMILIA: 'Familia',
        DEPARTAMENTO: 'Departamento'
    },
    // Colunas do CSV 'relatorio_entrada.csv'
    ENTRADA: {
        DATA: 'Data Entrada',
        CODIGO: 'Codigo',
        DESCRICAO: 'Descricao',
        QNTD: 'Quantidade',
        PRECO_UN: 'Valor Unitario',
        DESC: 'Desconto',
        TOTAL: 'Valor Total'
    }
};

//--------------------------------------------------------------------------
// Bloco 4 - Funções Auxiliares
//--------------------------------------------------------------------------

/**
 * Autentica com a API do Google Sheets usando a conta de serviço.
 */
async function authenticateGoogleSheets() {
    console.log('Autenticando com Google Sheets...');
    const serviceAccountInfo = JSON.parse(
        await fs.readFile(SERVICE_ACCOUNT_FILE, { encoding: 'utf8' })
    );

    const serviceAccountAuth = new JWT({
        email: serviceAccountInfo.client_email,
        key: serviceAccountInfo.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(GOOGLE_SHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    console.log(`Planilha "${doc.title}" carregada.`);
    return doc;
}

/**
 * Encontra a última data na planilha e retorna o dia SEGUINTE.
 */
async function getProximaDataInicial(sheet) {
    console.log(`Buscando última data na aba "${SHEET_NAME_ENTRADA}"...`);
    await sheet.loadHeaderRow();
    
    const headerIndex = sheet.headerValues.indexOf(DATE_COLUMN_HEADER);
    if (headerIndex === -1) {
        throw new Error(`Coluna "${DATE_COLUMN_HEADER}" não encontrada na planilha.`);
    }

    const rows = await sheet.getRows();
    if (rows.length === 0) {
        console.log('Planilha vazia. Usando data inicial padrão.');
        return parseDate(DEFAULT_START_DATE, 'yyyy-MM-dd', new Date());
    }

    for (let i = rows.length - 1; i >= 0; i--) {
        const ultimaDataString = rows[i].get(DATE_COLUMN_HEADER);
        if (ultimaDataString) {
            try {
                const ultimaData = parseDate(ultimaDataString, 'dd/MM/yyyy', new Date());
                const proximaData = addDays(ultimaData, 1);
                console.log(`Última data encontrada: ${ultimaDataString}. Próxima busca: ${formatDate(proximaData, 'dd/MM/yyyy')}`);
                return proximaData;
            } catch (error) {
                console.warn(`Data mal formatada ignorada: "${ultimaDataString}"`);
            }
        }
    }

    console.log('Nenhuma data válida encontrada nas linhas. Usando data inicial padrão.');
    return parseDate(DEFAULT_START_DATE, 'yyyy-MM-dd', new Date());
}

/**
 * Formata um objeto Date para o formato 'yyyy-MM-dd' exigido pelo ERP.
 */
function formatDateERP(date) {
    return formatDate(date, 'yyyy-MM-dd');
}

/**
 * Lê e parseia um arquivo XLSX (Excel).
 * Esta função substitui a antiga parseCSV.
 */
async function parseXLSX(filePath) {
    console.log(`Lendo arquivo XLSX: ${filePath}`);
    try {
        // Lê o arquivo como um buffer binário
        const fileBuffer = await fs.readFile(filePath);

        // Parseia o buffer do arquivo
        const workbook = xlsx.read(fileBuffer);
        
        // Pega a primeira planilha do arquivo
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Converte a planilha para um array de objetos JSON
        // dateNF garante que as datas do Excel sejam formatadas como string
        // no padrão da sua planilha (dd/MM/yyyy)
        const jsonData = xlsx.utils.sheet_to_json(worksheet, {
            dateNF: 'dd/mm/yyyy'
        });

        // O antigo 'csv-parse' tinha 'trim: true'. Vamos replicar isso
        // para os nomes das colunas (chaves do objeto), 
        // caso o Excel tenha cabeçalhos com espaços (ex: " Codigo ").
        const trimmedData = jsonData.map(row => {
            const newRow = {};
            for (const key in row) {
                newRow[key.trim()] = row[key];
            }
            return newRow;
        });
        
        return trimmedData;

    } catch (error) {
        console.error(`Erro ao ler o arquivo Excel ${filePath}:`, error);
        throw new Error(`Falha ao processar arquivo XLSX: ${filePath}`);
    }
}


/**
 * Limpa a pasta de downloads.
 */
async function limparDownloads() {
    console.log(`Limpando pasta de downloads: ${DOWNLOAD_PATH}`);
    try {
        const files = await fs.readdir(DOWNLOAD_PATH);
        for (const file of files) {
            await fs.unlink(path.join(DOWNLOAD_PATH, file));
        }
        console.log('Downloads limpos.');
    } catch (error) {
        if (error.code !== 'ENOENT') {
            console.error('Erro ao limpar downloads:', error);
        }
    }
}

/**
 * Processa os dados, cruza as informações e formata para a planilha.
 */
async function processarDados(relatorioEntradaPath, relatorioProdutosPath) {
    console.log('Iniciando processamento e cruzamento de dados...');
    
    // MODIFICADO: Chama a nova função parseXLSX
    const produtosCSV = await parseXLSX(relatorioProdutosPath);
    const mapaProdutos = new Map();
    const mapProd = MAPEAMENTO_COLUNAS_CSV.PRODUTOS;

    for (const produto of produtosCSV) {
        const codigo = produto[mapProd.CODIGO];
        if (codigo) {
            mapaProdutos.set(codigo, {
                familia: produto[mapProd.FAMILIA] || 'N/A',
                depto: produto[mapProd.DEPARTAMENTO] || 'N/A'
            });
        }
    }
    console.log(`${mapaProdutos.size} produtos carregados no mapa.`);

    // MODIFICADO: Chama a nova função parseXLSX
    const entradasCSV = await parseXLSX(relatorioEntradaPath);
    const dadosParaPlanilha = [];
    const mapEntrada = MAPEAMENTO_COLUNAS_CSV.ENTRADA;

    for (const entrada of entradasCSV) {
        const codigo = entrada[mapEntrada.CODIGO];
        const infoProduto = mapaProdutos.get(codigo) || { familia: 'N/A', depto: 'N/A' };

        // A biblioteca xlsx pode retornar números (ex: 123) 
        // onde o csv-parse retornava strings (ex: "123").
        // O código de mapeamento continua funcionando da mesma forma.
        const linha = [
            entrada[mapEntrada.DATA],
            codigo,
            entrada[mapEntrada.DESCRICAO],
            infoProduto.familia,
            infoProduto.depto,
            entrada[mapEntrada.QNTD],
            entrada[mapEntrada.PRECO_UN],
            entrada[mapEntrada.DESC],
            entrada[mapEntrada.TOTAL]
        ];
        
        dadosParaPlanilha.push(linha);
    }

    console.log(`${dadosParaPlanilha.length} linhas de entrada processadas.`);
    return dadosParaPlanilha;
}

/**
 * Adiciona os dados processados na planilha Google Sheets.
 */
async function atualizarPlanilha(sheet, dados) {
    if (dados.length === 0) {
        console.log('Nenhum dado novo para adicionar à planilha.');
        return;
    }
    
    console.log(`Adicionando ${dados.length} linhas à planilha...`);
    const headerValues = sheet.headerValues;
    const dadosComoObjetos = dados.map(linhaArray => {
        const objLinha = {};
        headerValues.forEach((header, index) => {
            objLinha[header] = linhaArray[index];
        });
        return objLinha;
    });

    await sheet.addRows(dadosComoObjetos);
    console.log('Planilha atualizada com sucesso.');
}


// --- NOVAS FUNÇÕES DE AUTOMAÇÃO (PARA O LOOP) ---

/**
 * Faz o login e baixa o relatório de produtos (dicionário).
 */
async function loginEBaixarProdutos(page) {
    // --- Login ---
    console.log('Fazendo login no ERP...');
    await page.goto(INOVE_URL);

    await page.locator(SELECTORS.INPUT_CODIGO).fill(process.env.USUARIO);
    await page.locator(SELECTORS.INPUT_SENHA).fill(process.env.SENHA);
    await page.getByLabel(SELECTORS.SELECT_EMPRESA).selectOption(process.env.EMPRESA_VALUE);
    await page.locator(SELECTORS.BUTTON_ENTRAR).click();

    console.log('Aguardando página principal carregar (esperando pelo botão "Produto")...');
    await page.locator(SELECTORS.BUTTON_PRODUTO).waitFor({ state: 'visible', timeout: 10000 });
    console.log('Login realizado com sucesso.');

    // --- Download Planilha Produtos (Dicionário) ---
    console.log('Baixando planilha de produtos (dicionário)...');
    await page.locator(SELECTORS.BUTTON_PRODUTO).click();
    await page.locator(SELECTORS.LINK_PLANILHA_PRODUTO).click();
    
    const downloadPromiseProduto = page.waitForEvent('download');
    await page.locator(SELECTORS.BUTTON_EXPORTAR_PRODUTO).click();
    const downloadProduto = await downloadPromiseProduto;
    
    // MODIFICADO: Salva com o nome .xlsx
    const pathProduto = path.join(DOWNLOAD_PATH, PRODUTOS_FILE_NAME);
    await downloadProduto.saveAs(pathProduto);
    console.log(`Planilha de produtos salva em: ${pathProduto}`);
    return pathProduto;
}

/**
 * Baixa o relatório de entradas para um período específico.
 */
async function baixarRelatorioEntrada(page, dataInicial, dataFinal) {
    console.log(`Baixando relatório de entradas de ${dataInicial} a ${dataFinal}...`);
    
    // --- CORREÇÃO APLICADA AQUI ---
    // Navega para a página de relatório usando os seletores que você forneceu,
    // que são mais robustos (getByRole e getByText).
    // O script não precisa voltar para a home, ele navega de onde parou.
    await page.getByRole('button', { name: 'Compra' }).click();
    await page.getByText('Entrada').nth(3).click();
    // --- FIM DA CORREÇÃO ---
    
    // Aplica filtros
    await page.locator(SELECTORS.SELECT_TIPO_ENTRADA).selectOption('1');
    await page.locator(SELECTORS.SELECT_TIPO_RELATORIO).selectOption('A');
    await page.locator(SELECTORS.SELECT_TIPO_ARQUIVO).selectOption('2'); // 2 = CSV (Mantemos isso, pois o ERP pode estar enviando Excel mesmo pedindo CSV)
    
    await page.locator(SELECTORS.INPUT_DATA_INICIAL).fill(dataInicial);
    await page.locator(SELECTORS.INPUT_DATA_FINAL).fill(dataFinal);

    // Gera relatório
    const downloadPromiseEntrada = page.waitForEvent('download', { timeout: 60000 }); // Timeout maior para relatórios
    await page.locator(SELECTORS.BUTTON_GERAR_ENTRADA).click();
    const downloadEntrada = await downloadPromiseEntrada;
    
    // MODIFICADO: Salva com o nome .xlsx
    const pathEntrada = path.join(DOWNLOAD_PATH, ENTRADA_FILE_NAME);
    await downloadEntrada.saveAs(pathEntrada);
    console.log(`Relatório de entradas salvo em: ${pathEntrada}`);
    return pathEntrada;
}

//--------------------------------------------------------------------------
// Bloco 5 - Principal (Função main)
//--------------------------------------------------------------------------

async function main() {
    console.log('--- Iniciando Automação de Relatório de Entradas ---');
    await limparDownloads(); // Limpa execuções anteriores
    await fs.mkdir(DOWNLOAD_PATH, { recursive: true });

    let browser;
    let pathProduto, pathEntrada;

    try {
        // --- 1. Lógica de Datas (Início) ---
        const doc = await authenticateGoogleSheets();
        const sheet = doc.sheetsByTitle[SHEET_NAME_ENTRADA];
        if (!sheet) {
            throw new Error(`Aba "${SHEET_NAME_ENTRADA}" não encontrada!`);
        }

        let dataInicialLoop = await getProximaDataInicial(sheet);
        const dataFinalGeral = subDays(startOfToday(), 1); // Sempre até ontem

        // Validação para não rodar desnecessariamente
        if (!isBefore(dataInicialLoop, dataFinalGeral)) {
            console.log(`Planilha já está atualizada até ${formatDate(dataFinalGeral, 'dd/MM/yyyy')}. Nenhuma ação necessária.`);
            console.log('--- Automação Concluída (Sem Novas Ações) ---');
            return;
        }
        
        console.log(`Planilha está em ${formatDate(dataInicialLoop, 'dd/MM/yyyy')}. Buscando dados até ${formatDate(dataFinalGeral, 'dd/MM/yyyy')}.`);

        // --- 2. Automação Web (Playwright) ---
        console.log('Iniciando navegador...');
        browser = await chromium.launch({ 
            headless: false, // Mude para true para rodar em background
            downloadsPath: DOWNLOAD_PATH
        });
        const context = await browser.newContext({ acceptDownloads: true });
        const page = await context.newPage();

        // --- Login e Download do Dicionário (Feito 1x) ---
        pathProduto = await loginEBaixarProdutos(page);

        // --- 3. Loop de Extração Mês a Mês ---
        while (isBefore(dataInicialLoop, dataFinalGeral)) {
            
            // 3a. Calcular o período deste loop (1 mês)
            let dataFinalLoop = endOfMonth(dataInicialLoop);
            
            // 3b. Garantir que o fim do loop não ultrapasse o dia de "ontem"
            if (isAfter(dataFinalLoop, dataFinalGeral)) {
                dataFinalLoop = dataFinalGeral;
            }

            const dataInicialFormatada = formatDateERP(dataInicialLoop);
            const dataFinalFormatada = formatDateERP(dataFinalLoop);
            
            console.log(`--- [LOOP] Processando período: ${dataInicialFormatada} a ${dataFinalFormatada} ---`);

            // 3c. Baixar o relatório de transações (entradas)
            pathEntrada = await baixarRelatorioEntrada(page, dataInicialFormatada, dataFinalFormatada);

            // 3d. Processar dados
            const dadosProcessados = await processarDados(pathEntrada, pathProduto);

            // 3e. Atualizar planilha
            await atualizarPlanilha(sheet, dadosProcessados);

            // 3f. Limpar download da entrada (para o próximo loop)
            await fs.unlink(pathEntrada);
            console.log(`Arquivo de entrada ${pathEntrada} processado e removido.`);

            // 3g. Preparar para o próximo loop
            dataInicialLoop = addDays(dataFinalLoop, 1);
        }

        await browser.close();
        console.log('Navegador fechado.');

        console.log('--- Automação Concluída com Sucesso ---');

    } catch (error) {
        console.error('!!! ERRO NA AUTOMAÇÃO !!!');
        console.error(error);
        if (browser) {
            await browser.close();
        }
    } finally {
        // --- 5. Limpeza Final (do arquivo de produtos) ---
        if (pathProduto) {
            try {
                await fs.unlink(pathProduto);
                console.log('Arquivo de produtos final limpo.');
            } catch (e) { /* ignora */ }
        }
    }
}

main();
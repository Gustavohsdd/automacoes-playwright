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
    // Colunas do 'Planilha Produto.xlsx - Produtos.csv'
    PRODUTOS: {
        CODIGO: 'Código',       // Corrigido de 'Codigo' 
        FAMILIA: 'Família',     // Corrigido de 'Familia' 
        DEPARTAMENTO: 'Departamento' // Já estava correto 
    },
    // Colunas do 'Relatório de Entrada.xlsx - Relatório.csv'
    ENTRADA: {
        // 'DATA' não é uma coluna no CSV, ela será extraída do cabeçalho da nota
        CODIGO: 'Código',       // Corrigido de 'Codigo' [cite: 7598]
        DESCRICAO: 'Descrição',   // Corrigido de 'Descricao' [cite: 7598]
        QNTD: 'Qntd.',          // Corrigido de 'Quantidade' [cite: 7598]
        PRECO_UN: 'Preço Unít.', // Corrigido de 'Valor Unitario' [cite: 7598]
        DESC: 'Desc.',          // Corrigido de 'Desconto' [cite: 7598]
        TOTAL: 'Total'          // Corrigido de 'Valor Total' [cite: 7598]
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
 * Lê e parseia um arquivo XLSX (Excel) de Entrada, retornando um array de arrays (linhas).
 * Esta função é específica para o formato complexo do relatório de entrada.
 */
async function parseXLSXEntrada(filePath) {
    console.log(`Lendo arquivo XLSX de Entrada: ${filePath}`);
    try {
        const fileBuffer = await fs.readFile(filePath);
        const workbook = xlsx.read(fileBuffer);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Converte para array de arrays (header: 1), 
        // preservando linhas e tratando células vazias como nulas
        const jsonData = xlsx.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: null 
        }); 
        return jsonData;
    } catch (error) {
        console.error(`Erro ao ler o arquivo Excel de Entrada ${filePath}:`, error);
        return [];
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
 * Processa os dados dos arquivos de entrada e de produtos.
 * ESTA FUNÇÃO FOI COMPLETAMENTE REESCRITA PARA SE ADAPTAR AOS SEUS RELATÓRIOS.
 */
async function processarDados(pathEntrada, pathProduto) {
    console.log('Processando dados...');

    // --- 1. Criar Mapa de "Tradução" de Produtos ---
    // (Usando a aba 'Produtos', que é uma tabela simples)
    const colProdutos = MAPEAMENTO_COLUNAS_CSV.PRODUTOS;
    const dadosProdutos = await parseXLSX(pathProduto); // Usa o parseXLSX original
    const mapaProdutos = new Map();
    
    for (const produto of dadosProdutos) {
        // 
        const codigo = produto[colProdutos.CODIGO]; 
        if (codigo) {
            mapaProdutos.set(codigo.toString().trim(), {
                // 
                familia: produto[colProdutos.FAMILIA],
                // 
                departamento: produto[colProdutos.DEPARTAMENTO]
            });
        }
    }
    console.log(`Mapa de ${mapaProdutos.size} produtos criado.`);

    // --- 2. Processar Relatório de Entrada (Formato Complexo) ---
    const colEntrada = MAPEAMENTO_COLUNAS_CSV.ENTRADA;
    // Usa a nova função para ler o relatório como linhas/colunas
    const linhasEntrada = await parseXLSXEntrada(pathEntrada); 
    
    const dadosFinais = [];
    let dataEntradaAtual = null;
    let indicesCabecalho = null;

    // Cabeçalhos-alvo da planilha Google (conforme sua descrição)
    const TARGET_HEADERS = {
        DATA_ENTRADA: 'Data Entrada',
        CODIGO: 'Código',
        DESCRICAO: 'Descrição',
        FAMILIA: 'Família',
        DEPARTAMENTO: 'Departamento',
        QNTD: 'Qntd.',
        PRECO_UN: 'Preço Unít.',
        DESC: 'Desc.',
        TOTAL: 'Total'
    };

    // Expressão regular para encontrar a data de entrada no cabeçalho da nota
    // Ex: "Entrada: 01/10/2025" [cite: 7596, 7611]
    const regexDataEntrada = /Entrada:\s*(\d{2}\/\d{2}\/\d{4})/;

    for (const linha of linhasEntrada) {
        // Ignora linhas vazias ou sem dados
        if (!linha || linha.length === 0 || linha.every(cell => cell === null)) {
            continue;
        }

        // Converte a linha inteira para string para facilitar a busca de padrões
        const linhaString = linha.join(',');

        // --- A. Procurar Cabeçalho da Nota (para achar a Data) ---
        const matchData = linhaString.match(regexDataEntrada);
        if (matchData && matchData[1]) {
            dataEntradaAtual = matchData[1]; // Ex: "01/10/2025" [cite: 7596, 7611]
            
            // Reseta o cabeçalho do produto, pois estamos em uma nova nota
            indicesCabecalho = null; 
            console.log(`Nota encontrada. Data de Entrada: ${dataEntradaAtual}`);
            continue;
        }

        // --- B. Procurar Cabeçalho dos Produtos ---
        // (Procura pela linha que define as colunas dos produtos)
        // [cite: 7598, 7614]
        if (linha[0] === colEntrada.CODIGO && linha[1] === colEntrada.DESCRICAO) {
            // Encontramos o cabeçalho. Mapeia os índices (posição) das colunas.
            indicesCabecalho = {
                codigo: linha.indexOf(colEntrada.CODIGO),
                descricao: linha.indexOf(colEntrada.DESCRICAO),
                qntd: linha.indexOf(colEntrada.QNTD),
                precoUn: linha.indexOf(colEntrada.PRECO_UN),
                desc: linha.indexOf(colEntrada.DESC),
                total: linha.indexOf(colEntrada.TOTAL)
            };
            
            // Verifica se todas as colunas necessárias foram encontradas
            if (Object.values(indicesCabecalho).some(v => v === -1)) {
                console.warn('Alerta: O cabeçalho de produtos foi encontrado, mas algumas colunas do MAPEAMENTO_COLUNAS_CSV não correspondem.');
                console.warn('Cabeçalho no arquivo:', linha);
                console.warn('Índices encontrados:', indicesCabecalho);
                indicesCabecalho = null; // Invalida para não processar errado
            } else {
                console.log('Cabeçalho de produtos encontrado. Iniciando leitura dos itens...');
            }
            continue;
        }

        // --- C. Processar Linhas de Produto ---
        // Só processa se já tivermos uma data e um cabeçalho de produtos válido
        if (dataEntradaAtual && indicesCabecalho) {
            
            const codigoProduto = linha[indicesCabecalho.codigo];
            
            // Se a primeira coluna não for um código válido (ex: linha vazia, "Total:", etc.),
            // assume que os produtos desta nota acabaram.
            if (!codigoProduto || (typeof codigoProduto !== 'number' && !/^\d+$/.test(codigoProduto.toString()))) {
                // Reseta o cabeçalho para parar de ler produtos até a próxima nota
                indicesCabecalho = null;
                continue;
            }

            const codigoStr = codigoProduto.toString().trim();
            // Busca no mapa de produtos; se não achar, usa 'N/A'
            const produtoInfo = mapaProdutos.get(codigoStr) || { familia: 'N/A', departamento: 'N/A' };
            
            // Monta o objeto final com os nomes exatos das colunas da planilha Google
            const novaLinha = {
                [TARGET_HEADERS.DATA_ENTRADA]: dataEntradaAtual,
                [TARGET_HEADERS.CODIGO]: codigoStr,
                [TARGET_HEADERS.DESCRICAO]: linha[indicesCabecalho.descricao],
                [TARGET_HEADERS.FAMILIA]: produtoInfo.familia,
                [TARGET_HEADERS.DEPARTAMENTO]: produtoInfo.departamento,
                [TARGET_HEADERS.QNTD]: linha[indicesCabecalho.qntd],
                [TARGET_HEADERS.PRECO_UN]: linha[indicesCabecalho.precoUn],
                [TARGET_HEADERS.DESC]: linha[indicesCabecalho.desc],
                [TARGET_HEADERS.TOTAL]: linha[indicesCabecalho.total]
            };

            dadosFinais.push(novaLinha);
        }
    }

    console.log(`Processamento concluído. ${dadosFinais.length} linhas de produtos extraídas.`);
    return dadosFinais;
}

/**
 * Adiciona os dados processados na planilha Google Sheets.
 * ESTA FUNÇÃO FOI CORRIGIDA.
 */
async function atualizarPlanilha(sheet, dados) {
    if (dados.length === 0) {
        console.log('Nenhum dado novo para adicionar à planilha.');
        return;
    }
    
    console.log(`Adicionando ${dados.length} linhas à planilha...`);
    
    // CORREÇÃO:
    // A função 'processarDados' já retorna um array de objetos 
    // com os cabeçalhos corretos (ex: { 'Data Entrada': '...', 'Código': '...' }).
    // A biblioteca 'google-spreadsheet' aceita esse formato diretamente.
    // A conversão anterior (map) estava quebrando os dados.
    await sheet.addRows(dados);
    
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
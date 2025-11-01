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
const DEFAULT_START_DATE = '2022-03-01'; // Data (yyyy-MM-dd) para usar se a planilha estiver vazia

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
 * [FUNÇÃO AUXILIAR - ADICIONADA FORA DE PROCESSAR DADOS, NO BLOCO 4]
 * Normaliza o código do produto.
 * Garante que "1234", 1234, "1234.0", 1234.0, " 1234 " virem "1234".
 */
function normalizarCodigo(codigo) {
    if (codigo === null || typeof codigo === 'undefined') {
        return null;
    }
    // Converte para string, remove espaços, e pega a parte antes de um ponto ou vírgula.
    // Isso trata tanto 1234.0 quanto 1234,0
    const codigoStr = codigo.toString().trim();
    const parteInteira = codigoStr.split('.')[0].split(',')[0];
    
    // Se a string original for vazia ou se tornar vazia, retorna nulo
    if (parteInteira.length === 0) {
        return null;
    }
    return parteInteira;
}


/**
 * Processa os dados dos arquivos de entrada e de produtos.
 * * [VERSÃO 3 - CORREÇÃO DE ZEROS À ESQUERDA]
 * * Esta versão corrige o problema de "00002" (Entrada) vs "2" (Produtos).
 * * Ela faz isso convertendo AMBOS os códigos para número e depois para string,
 * * garantindo que a chave de busca seja idêntica.
 */
async function processarDados(pathEntrada, pathProduto) {
    console.log('Processando dados...');

    // --- Lógica de Normalização (para garantir que "00002" e "2" sejam iguais) ---
    // Esta função "limpa" os códigos antes de compará-los
    const normalizarCodigo = (codigo) => {
        if (codigo === null || typeof codigo === 'undefined') {
            return null; // Ignora valores nulos/undefined
        }
        
        // Converte para string, remove espaços
        let codigoStr = codigo.toString().trim();
        
        // Remove a parte decimal (ex: "1234.0" -> "1234")
        // Isso também trata "1234,0"
        codigoStr = codigoStr.split('.')[0].split(',')[0];

        // Se ficou vazio (ex: era ".0"), retorna nulo
        if (codigoStr.length === 0) {
            return null;
        }

        // *** A CORREÇÃO PRINCIPAL ESTÁ AQUI ***
        // 1. Tenta converter para número.
        //    Isso transforma "00002" em 2 (número).
        //    Isso transforma "2" em 2 (número).
        const codigoNum = Number(codigoStr);

        // 2. Verifica se a conversão falhou (ex: o código era "ABC")
        if (isNaN(codigoNum)) {
            // Se não for um número (ex: código de barras com letras),
            // apenas retorna a string original limpa.
            // NOTA: Se seus códigos NUNCA tiverem letras, você pode 
            // até remover este 'if' e o 'console.warn'.
            console.warn(`Código não-numérico encontrado: "${codigoStr}". Usando como está.`);
            return codigoStr; // Retorna o código "ABC" como ele é
        }

        // 3. Converte de volta para string.
        //    Isso transforma 2 (número) em "2" (texto).
        // Agora, "00002" e "2" SÃO IGUAIS (ambos viram "2").
        return codigoNum.toString();
    };
    // --- Fim da Lógica de Normalização ---


    // --- 1. Criar o Dicionário (Mapa) com a Planilha Produto ---
    // (O "banco de dados" do PROCV)
    console.log('Criando dicionário de produtos (PROCV)...');
    const colProdutos = MAPEAMENTO_COLUNAS_CSV.PRODUTOS;
    const dadosProdutos = await parseXLSX(pathProduto);
    const mapaProdutos = new Map();
    
    for (const produto of dadosProdutos) {
        // [APLICA A CORREÇÃO]
        const codigoOriginal = produto[colProdutos.CODIGO]; 
        const codigoNorm = normalizarCodigo(codigoOriginal); // Normaliza (ex: "2" -> "2")

        if (codigoNorm) { 
            // Armazena no mapa. Chave: "2", Valor: { familia: "...", depto: "..." }
            mapaProdutos.set(codigoNorm, {
                familia: produto[colProdutos.FAMILIA],
                departamento: produto[colProdutos.DEPARTAMENTO]
            });
        }
    }
    console.log(`Dicionário criado com ${mapaProdutos.size} produtos.`);
    if (mapaProdutos.size === 0) {
        console.warn('ALERTA: O dicionário de produtos está vazio. Verifique o nome da coluna "Código" no seu MAPEAMENTO_COLUNAS_CSV.PRODUTOS e no arquivo Excel.');
    }


    // --- 2. Processar Relatório de Entrada (onde o PROCV será aplicado) ---
    const colEntrada = MAPEAMENTO_COLUNAS_CSV.ENTRADA;
    const linhasEntrada = await parseXLSXEntrada(pathEntrada); 
    
    const dadosFinais = [];
    let dataEntradaAtual = null;
    let indicesCabecalho = null; // Armazena a posição das colunas

    // Cabeçalhos-alvo da planilha Google
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

    const regexDataEntrada = /Entrada:\s*(\d{2}\/\d{2}\/\d{4})/;

    for (const linha of linhasEntrada) {
        if (!linha || linha.length === 0 || linha.every(cell => cell === null)) {
            continue; // Ignora linha vazia
        }

        const linhaString = linha.join(',');

        // --- A. Achar a Data da Nota ---
        const matchData = linhaString.match(regexDataEntrada);
        if (matchData && matchData[1]) {
            dataEntradaAtual = matchData[1];
            indicesCabecalho = null; // Reseta o cabeçalho, pois é uma nova nota
            continue;
        }

        // --- B. Achar o Cabeçalho dos Produtos na Nota ---
        // (Verifica se a linha é a de cabeçalho, ex: "Código", "Descrição", ...)
        const primeiraCelula = linha[0] ? String(linha[0]).trim() : "";
        const segundaCelula = linha[1] ? String(linha[1]).trim() : "";

        if (primeiraCelula === colEntrada.CODIGO && segundaCelula === colEntrada.DESCRICAO) {
            // Encontramos o cabeçalho! Mapear os índices.
            indicesCabecalho = {
                codigo: linha.indexOf(colEntrada.CODIGO),
                descricao: linha.indexOf(colEntrada.DESCRICAO),
                qntd: linha.indexOf(colEntrada.QNTD),
                precoUn: linha.indexOf(colEntrada.PRECO_UN),
                desc: linha.indexOf(colEntrada.DESC),
                total: linha.indexOf(colEntrada.TOTAL)
            };
            
            if (Object.values(indicesCabecalho).some(v => v === -1)) {
                console.warn('Alerta: O cabeçalho de produtos foi encontrado, mas algumas colunas do MAPEAMENTO_COLUNAS_CSV.ENTRADA não correspondem.');
                console.warn('Cabeçalho no arquivo:', linha);
                console.warn('MAPEAMENTO:', colEntrada);
                indicesCabecalho = null; // Invalida para não processar errado
            } else {
                console.log('Cabeçalho de produtos encontrado. Lendo itens...');
            }
            continue;
        }

        // --- C. Processar Linhas de Produto (Aplicar o PROCV) ---
        // Só executa se já tivermos encontrado uma Data (A) e um Cabeçalho (B)
        if (dataEntradaAtual && indicesCabecalho) {
            
            const codigoProdutoOriginal = linha[indicesCabecalho.codigo];
            
            // [APLICA A CORREÇÃO]
            const codigoNorm = normalizarCodigo(codigoProdutoOriginal); // Normaliza (ex: "00002" -> "2")

            // Se o código não for válido (linha em branco, linha de "Total:", etc.),
            // significa que os produtos desta nota acabaram.
            if (!codigoNorm) { 
                indicesCabecalho = null; // Para de ler produtos até a próxima nota
                continue;
            }

            // --- O PROCV (VLOOKUP) ACONTECE AQUI ---
            // Procura o 'codigoNorm' (ex: "2") no 'mapaProdutos'.
            // Se não achar, usa o valor padrão (N/A).
            const produtoInfo = mapaProdutos.get(codigoNorm) || { familia: 'N/A', departamento: 'N/A' };
            
            // [LOG DE AJUDA] Se ainda falhar, isso vai te dizer o porquê
            if (produtoInfo.familia === 'N/A') { 
                console.log(`PROCV FALHOU: Código Original: "${codigoProdutoOriginal}", Normalizado: "${codigoNorm}". Não encontrado no mapa.`); 
            }

            // --- Monta a linha final ---
            const novaLinha = {
                [TARGET_HEADERS.DATA_ENTRADA]: dataEntradaAtual,
                [TARGET_HEADERS.CODIGO]: codigoNorm, // Salva o código normalizado
                [TARGET_HEADERS.DESCRICAO]: linha[indicesCabecalho.descricao],
                [TARGET_HEADERS.FAMILIA]: produtoInfo.familia, // Resultado do PROCV
                [TARGET_HEADERS.DEPARTAMENTO]: produtoInfo.departamento, // Resultado do PROCV
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
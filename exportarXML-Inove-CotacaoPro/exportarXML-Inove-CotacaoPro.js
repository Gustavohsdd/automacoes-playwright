// ===================================================================================
// BLOCO 0: IMPORTAÇÕES E CONFIGURAÇÕES INICIAIS
// ===================================================================================
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const AdmZip = require('adm-zip');
const readline = require('readline');
require('dotenv').config();

// Função para pedir informações ao usuário, agora com validação e cores
function askQuestion(query, validationRegex) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    const yellow = '\x1b[33m'; // Cor amarela para a pergunta
    const red = '\x1b[31m';    // Cor vermelha para o erro
    const reset = '\x1b[0m';   // Reseta a cor para o padrão

    return new Promise(resolve => {
        const ask = () => {
            rl.question(`${yellow}${query}${reset}`, ans => {
                // Se uma validação (regex) foi fornecida e o input não passar, pede de novo
                if (validationRegex && !validationRegex.test(ans)) {
                    console.log(`${red}--> Formato inválido. Por favor, use o formato solicitado.${reset}`);
                    ask();
                } else {
                    rl.close();
                    resolve(ans);
                }
            });
        };
        ask();
    });
}


// ===================================================================================
// FUNÇÃO PRINCIPAL DA AUTOMAÇÃO
// ===================================================================================
(async () => {
    const browser = await chromium.launch({
        headless: process.env.HEADLESS === 'true'
    });
    const context = await browser.newContext();
    const page = await context.newPage();

    // Variáveis para guardar os caminhos dos arquivos para a limpeza no final
    let downloadPath = '';
    let extractDir = '';

    try {
        // ===========================================================================
        // BLOCO 1: ACESSO E LOGIN
        // ===========================================================================
        console.log('Acessando a página de login...');
        await page.goto(process.env.URL_LOGIN);

        console.log('Realizando login...');
        await page.getByRole('textbox', { name: 'Codigo' }).fill(process.env.USUARIO);
        await page.getByRole('textbox', { name: 'Senha' }).fill(process.env.SENHA);
        await page.getByLabel('Empresa *').selectOption(process.env.EMPRESA_VALUE);
        await page.getByRole('button', { name: 'Entrar' }).click();
        await page.waitForURL(process.env.URL_HOME + '/**');
        console.log('Login realizado com sucesso!');


        // ===========================================================================
        // BLOCO 2: NAVEGAÇÃO ATÉ A PÁGINA DE NOTAS FISCAIS
        // ===========================================================================
        console.log('Navegando para a tela de importação de NF-e...');
        await page.getByRole('button', { name: 'Compra' }).click();
        await page.getByRole('paragraph').filter({ hasText: 'Importação NF-e' }).click();
        await page.getByRole('button', { name: 'Dados de Pesquisa Informações' }).click();


        // ===========================================================================
        // BLOCO 3: PERGUNTAR E INSERir DATAS (COM VALIDAÇÃO E CORES)
        // ===========================================================================
        console.log('Aguardando entrada de datas...');
        const dateFormatRegex = /^\d{2}\/\d{2}\/\d{4}$/; // Valida o formato DD/MM/AAAA

        const dataInicial = await askQuestion('Digite a data inicial (formato DD/MM/AAAA): ', dateFormatRegex);
        const dataFinal = await askQuestion('Digite a data final (formato DD/MM/AAAA): ', dateFormatRegex);

        console.log(`Pesquisando do dia ${dataInicial} até ${dataFinal}...`);

        const dataInicialNumeros = dataInicial.replace(/\//g, '');
        const dataFinalNumeros = dataFinal.replace(/\//g, '');

        const campoDataInicial = page.locator('ia-input-list').filter({ hasText: 'Inicial' }).locator('#DataEmissao');
        await campoDataInicial.click();
        await campoDataInicial.type(dataInicialNumeros, { delay: 100 });

        const campoDataFinal = page.locator('ia-input-list').filter({ hasText: 'Final' }).locator('#DataEmissao2');
        await campoDataFinal.click();
        await campoDataFinal.type(dataFinalNumeros, { delay: 100 });


        // ===========================================================================
        // BLOCO 4: PESQUISA E DOWNLOAD DO ARQUIVO ZIP
        // ===========================================================================
        console.log('Iniciando a pesquisa e aguardando o download...');
        const downloadPromise = page.waitForEvent('download');
        await page.getByRole('button', { name: 'Pesquisar' }).click();
        console.log('Aguardando o carregamento dos dados...');
        await page.waitForTimeout(5000);
        await page.getByRole('button', { name: 'Exportar XML' }).click();
        const download = await downloadPromise;

        const downloadDir = path.join(__dirname, 'downloads');
        if (!fs.existsSync(downloadDir)) {
            fs.mkdirSync(downloadDir);
        }
        // Atribui o caminho à variável para ser usada na limpeza
        downloadPath = path.join(downloadDir, download.suggestedFilename());
        await download.saveAs(downloadPath);
        console.log(`Arquivo ZIP salvo em: ${downloadPath}`);


        // ===========================================================================
        // BLOCO 5: DESCOMPACTAR O ARQUIVO ZIP
        // ===========================================================================
        console.log('Descompactando arquivos XML...');
        // Atribui o caminho à variável para ser usada na limpeza
        extractDir = path.join(__dirname, 'extracted_xmls');
        
        if (fs.existsSync(extractDir)) {
            fs.rmSync(extractDir, { recursive: true, force: true });
        }
        fs.mkdirSync(extractDir);

        const zip = new AdmZip(downloadPath);
        zip.extractAllTo(extractDir, true);
        console.log(`Arquivos extraídos para: ${extractDir}`);


        // ===========================================================================
        // BLOCO 6: ACESSAR O SISTEMA DE UPLOAD
        // ===========================================================================
        console.log('Acessando o sistema de conciliação de NF-e...');
        const uploadUrl = 'https://script.google.com/macros/s/AKfycbxnryJxK8kxhul0UXXsOJpCGkukuomrItq16DngLw1u0IIyp7liFlyFaOXRYCzhJTY/exec?view=conciliacaonf';
        await page.goto(uploadUrl);
        console.log('Aguardando o carregamento completo da página de upload...');
        await page.waitForLoadState('networkidle', { timeout: 60000 });


        // ===========================================================================
        // BLOCO 7: FAZER UPLOAD DOS ARQUIVOS XML
        // ===========================================================================
        console.log('Iniciando o processo de upload...');
        const fileChooserPromise = page.waitForEvent('filechooser');
        const frame1 = await page.frameLocator('iframe[title="Conciliação de NF-e"]');
        const frame2 = await frame1.frameLocator('iframe[title="Conciliação de NF-e"]');
        await frame2.getByRole('button', { name: ' Fazer Upload' }).click();
        const fileChooser = await fileChooserPromise;

        const xmlFiles = fs.readdirSync(extractDir)
            .filter(file => file.toLowerCase().endsWith('.xml'))
            .map(file => path.join(extractDir, file));

        if (xmlFiles.length > 0) {
            await fileChooser.setFiles(xmlFiles);
            console.log(`${xmlFiles.length} arquivos XML foram selecionados para upload.`);
            console.log('A automação iniciou o upload. Monitore a janela do navegador para ver o processamento finalizar.');
        } else {
            console.log('Nenhum arquivo XML encontrado na pasta de extração.');
        }
        
        await page.waitForTimeout(15000); // Tempo para o upload iniciar e processar


        // ===========================================================================
        // BLOCO 8: LIMPEZA DOS ARQUIVOS (NOVO)
        // ===========================================================================
        console.log('Operação finalizada com sucesso. Realizando limpeza...');
        
        if (downloadPath && fs.existsSync(downloadPath)) {
            fs.unlinkSync(downloadPath);
            console.log(`--> Arquivo ZIP temporário deletado.`);
        }
        if (extractDir && fs.existsSync(extractDir)) {
            fs.rmSync(extractDir, { recursive: true, force: true });
            console.log(`--> Pasta de XMLs temporária deletada.`);
        }

    } catch (error) {
        console.error('Ocorreu um erro durante a automação:', error);
    } finally {
        await browser.close();
        console.log('Automação finalizada.');
    }
})();
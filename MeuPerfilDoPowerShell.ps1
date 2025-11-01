# ===== Atalhos de automações do Gustavo =====

# Função: ENTRADA NF-e
function inove-nfe {
    param([string[]]$Args)
    Push-Location 'C:\Users\gusta\Documents\AutomacoesPro\INOVE-EntradaNF'
    try {
        & node 'entrada-nfe-automatizada.js' @Args
    } finally { Pop-Location }
}
Set-Alias nfe inove-nfe  # atalho curto

# Função: ENTRADA NF-e com XML
function inove-nfe-xml {
    param([string[]]$Args)
    Push-Location 'C:\Users\gusta\Documents\AutomacoesPro\INOVE-EntradaNFcomXML'
    try {
        & node 'entrada-nfe-com-xml.js' @Args
    } finally { Pop-Location }
}
Set-Alias nfexml inove-nfe-xml  # atalho curto

# Função: PERDAS (Sheets)
function inove-perdas {
    param([string[]]$Args)
    Push-Location 'C:\Users\gusta\Documents\AutomacoesPro\INOVEeSHEETS-Perdas'
    try {
        & node 'perdas-from-sheets' @Args
    } finally { Pop-Location }
}
Set-Alias perdas inove-perdas  # atalho curto

# Função: EXPORTAR XML (NOVO)
function inove-exportarxml {
    param([string[]]$Args)
    Push-Location 'C:\Users\gusta\Documents\AutomacoesPro\exportarXML-Inove-CotacaoPro'
    try {
        & node 'exportarXML-Inove-CotacaoPro.js' @Args
    } finally { Pop-Location }
}
Set-Alias exportxml inove-exportarxml # atalho curto

# Função: RELATÓRIO PRODUTOS POR HORÁRIO (NOVO)
function inove-relatorio-horario {
    param([string[]]$Args)
    Push-Location 'C:\Users\gusta\Documents\AutomacoesPro\Relatorio-Produtos-por-Horario-INOVE-SHEETS'
    try {
        & node 'RelatorioProdutosporHorario.js' @Args
    } finally { Pop-Location }
}
Set-Alias relatorio inove-relatorio-horario # atalho curto


# ---- Menu (digite: rpa) ----
function automacoes {
    param([string]$acao)

    # Mapa de automações atualizado
    $map = [ordered]@{
        '1' = @{ Nome='Entrada NF-e no INOVE';                   Run={ inove-nfe } }
        '2' = @{ Nome='Entrada NF-e com XML';                    Run={ inove-nfe-xml } }
        '3' = @{ Nome='Perdas do Sheets para o INOVE';           Run={ inove-perdas } }
        '4' = @{ Nome='Exportar XMLs do INOVE para o CotaçãoPro'; Run={ inove-exportarxml } }
        '5' = @{ Nome='Produto por horário (data automática)';   Run={ inove-relatorio-horario } } # NOVO ITEM
    }

    # Lógica para execução direta (ex: rpa 1)
    if ($acao) {
        $hit = $map.Keys | Where-Object { $_ -eq $acao -or $map[$_].Nome -like "*$acao*" }
        if ($hit.Count -eq 1) { & $map[$hit].Run; return }
    }

    # --- Lógica de exibição do menu ATUALIZADA ---
    Write-Host "`n====== Tarefas ======" -ForegroundColor Cyan
    foreach ($k in $map.Keys) {
        # Adiciona o cabeçalho "Relatórios" antes de imprimir o item 5
        if ($k -eq '5') {
            Write-Host "`n===== Relatórios =====" -ForegroundColor Cyan
        }
        # Imprime o item do menu
        Write-Host (" [{0}] {1}" -f $k,$map[$k].Nome) -ForegroundColor Cyan
    }
    # --- Fim da lógica de exibição ---

    $escolha = Read-Host "Digite o número ou parte do nome"
    $match = $map.Keys | Where-Object { $_ -eq $escolha -or $map[$_].Nome -like "*$escolha*" }
    
    if ($match.Count -eq 1) { 
        & $map[$match].Run 
    }
    else { 
        Write-Host "Opção inválida." -ForegroundColor Red 
    }
}
Set-Alias rpa automacoes

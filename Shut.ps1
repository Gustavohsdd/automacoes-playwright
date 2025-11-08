### SCRIPT DE CONTROLE PARA DVR INTELBRAS (MHDX 3132) - VERSÃO 3 (MODIFICADO) ###

# --- Configure suas informações aqui ---
$ipDVR = "10.0.0.91" 
$portaHTTP = "80" 
$usuario = "admin2" 

# !!! ATENÇÃO !!!
# COLOQUE SUA SENHA ABAIXO, ENTRE AS ASPAS.
# ESTE SCRIPT NÃO VAI MAIS PERGUNTAR POR ELA.
$senha = "12345@12345" 
# ----------------------------------------

# 1. Define a ação diretamente (sem perguntar)
$acao = "shutdown" # "shutdown" para desligar, "reboot" para reiniciar

# 2. Verifica se a senha foi alterada
if ($senha -eq "!SUA_SENHA_AQUI!") {
    Write-Error "ERRO: Você não definiu a variável '$senha' no script."
    Write-Error "Por favor, edite o arquivo .ps1 e coloque sua senha."
    return # Para o script
}

Write-Host "Iniciando comando direto para DESLIGAR o DVR ($ipDVR)..." -ForegroundColor Yellow

# 3. Converte a senha de texto puro para um formato seguro (SecureString)
#    Necessário para criar o objeto PSCredential
$securePassword = ConvertTo-SecureString -String $senha -AsPlainText -Force

# 4. Cria o objeto de credencial manualmente
$cred = New-Object System.Management.Automation.PSCredential($usuario, $securePassword)

# 5. Monta a URL da API da Intelbras/Dahua
$url = "http://$ipDVR`:$portaHTTP/cgi-bin/magicBox.cgi?action=$acao"

# 6. Envia o comando para o DVR
try {
    # Usamos Invoke-RestMethod que lida muito bem com autenticação
    Write-Host "Enviando comando '$acao' para $url..."
    Invoke-RestMethod -Uri $url -Method Get -Credential $cred -AllowUnencryptedAuthentication
    
    Write-Host "Comando '$acao' enviado com sucesso!" -ForegroundColor Green
    Write-Host "O DVR deve iniciar o processo de desligamento."
}
catch {
    Write-Error "FALHA AO ENVIAR O COMANDO: $_"
    Write-Error "Verifique o IP ($ipDVR), usuário ($usuario) e se a senha no script está correta."
}

Write-Host "Script finalizado."
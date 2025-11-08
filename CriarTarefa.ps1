# --- CONFIGURE AQUI ---

# 1. Defina o nome da sua tarefa (será usado para criar e deletar)
$taskName = "Iniciar Shut"

# 2. Defina o caminho COMPLETO para o seu script que desliga o DVR
$scriptPath = "C:\Users\gusta\Documents\AutomacoesPro\Shut.ps1"

# 3. Defina o horário de execução (Formato 24h: "HH:mm")
#    A tarefa será agendada para ESTE horário, NO DIA DE HOJE.
$horaDeExecucao = "23:30"

# --- FIM DA CONFIGURAÇÃO ---

# 4. Constrói o gatilho (trigger)
#    (Get-Date).Date pega a data de hoje (à meia-noite)
#    [timespan]::Parse() converte a string de hora em um objeto de tempo
#    O resultado é um DateTime exato: Hoje, às 23:30 (por exemplo)
$runTime = (Get-Date).Date + [timespan]::Parse($horaDeExecucao)

# 5. Define o gatilho como "-Once" (uma vez) e "-At" (no horário calculado)
$trigger = New-ScheduledTaskTrigger -Once -At $runTime

# 6. Define a ação: executar o PowerShell, que por sua vez executa seu script
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File `"$scriptPath`""

# 7. Define quem vai executar a tarefa (SYSTEM é o mais robusto)
$principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -RunLevel Highest

# 8. Define configurações
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

# 9. Registra (cria) a tarefa no Agendador de Tarefas do Windows
Write-Host "Registrando a tarefa agendada '$taskName'..."

try {
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
    Write-Host "Sucesso! A tarefa '$taskName' foi criada." -ForegroundColor Green
    Write-Host "Ela será executada UMA ÚNICA VEZ hoje às $horaDeExecucao."
}
catch {
    Write-Error "FALHA ao criar a tarefa: $_"
}

> **IMPORTANTE:**
> Você deve executar este script *antes* do horário definido em `$horaDeExecucao`.
> Se você rodar este script às 23:40, mas o horário estiver para 23:30, a tarefa será criada para um horário que *já passou* e não será executada.
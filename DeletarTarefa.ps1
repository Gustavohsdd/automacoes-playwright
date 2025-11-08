# --- CONFIGURE AQUI ---

# 1. Defina o nome EXATO da tarefa que você criou no script anterior
$taskName = "Iniciar Shut"

# --- FIM DA CONFIGURAÇÃO ---

Write-Host "Tentando excluir a tarefa agendada '$taskName'..."

try {
    # -Confirm:$false faz com que ele não peça confirmação
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
    Write-Host "Tarefa '$taskName' excluída com sucesso." -ForegroundColor Green
}
catch {
    Write-Error "FALHA ao excluir a tarefa: $_"
    Write-Error "Verifique se o nome '$taskName' está correto e se você executou este script como Administrador."
}
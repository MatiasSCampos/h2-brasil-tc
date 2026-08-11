# =============================================================================
# Agenda a execução semanal automática no Agendador de Tarefas do Windows:
# roda o coletor, prepara os dados da API e publica no GitHub (que redeploya
# a API pública no Render sozinho) -- tudo em uma tacada só, com log em arquivo.
#
# Como usar:
#   1. Abra o PowerShell NESTA pasta (a mesma do coletor_h2_brasil.py,
#      atualizar_e_publicar.ps1 e preparar_dados_api.py)
#   2. Rode:  .\agendar_tarefa.ps1
#      (se der erro de permissão de execução de script, rode antes:
#       Set-ExecutionPolicy -Scope CurrentUser RemoteSigned)
#   3. Pronto -- a tarefa "ColetorH2Brasil" passa a rodar toda semana.
#
# Onde ver se funcionou:
#   dados_h2\logs\execucao_AAAA-MM-DD_HH-mm-ss.log  (um arquivo por execução)
#
# Para rodar manualmente agora (teste imediato, sem esperar a agenda):
#   Start-ScheduledTask -TaskName "ColetorH2Brasil"
#
# Para remover a tarefa depois:
#   Unregister-ScheduledTask -TaskName "ColetorH2Brasil" -Confirm:$false
# =============================================================================

$NomeTarefa = "ColetorH2Brasil"
$PastaProjeto = $PSScriptRoot
$CaminhoScriptPublicacao = Join-Path $PastaProjeto "atualizar_e_publicar.ps1"
$PastaLogs = Join-Path $PastaProjeto "dados_h2\logs"
$CaminhoPowerShell = (Get-Command powershell).Source

if (-not (Test-Path $CaminhoScriptPublicacao)) {
    Write-Error "Não encontrei atualizar_e_publicar.ps1 em $PastaProjeto. Rode este .ps1 na mesma pasta dos outros scripts."
    exit 1
}

New-Item -ItemType Directory -Force -Path $PastaLogs | Out-Null

# Script .cmd "envelope" que roda o atualizar_e_publicar.ps1 e redireciona a
# saída para um log com data/hora no nome.
$CaminhoEnvelope = Join-Path $PastaProjeto "_rodar_coletor.cmd"
$ConteudoEnvelope = @"
@echo off
setlocal
set TIMESTAMP=%date:~-4%-%date:~3,2%-%date:~0,2%_%time:~0,2%-%time:~3,2%-%time:~6,2%
set TIMESTAMP=%TIMESTAMP: =0%
cd /d "$PastaProjeto"
"$CaminhoPowerShell" -ExecutionPolicy Bypass -File "$CaminhoScriptPublicacao" > "$PastaLogs\execucao_%TIMESTAMP%.log" 2>&1
endlocal
"@
Set-Content -Path $CaminhoEnvelope -Value $ConteudoEnvelope -Encoding ASCII

# Roda toda segunda-feira às 07:00. Ajuste -DaysOfWeek / -At como preferir.
$Gatilho = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 7:00am

$Acao = New-ScheduledTaskAction `
    -Execute $CaminhoEnvelope `
    -WorkingDirectory $PastaProjeto

$Configuracoes = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -DontStopOnIdleEnd `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask `
    -TaskName $NomeTarefa `
    -Trigger $Gatilho `
    -Action $Acao `
    -Settings $Configuracoes `
    -Description "Executa semanalmente o coletor de H2, prepara os dados e publica no GitHub/Render. Log em dados_h2\logs\." `
    -Force

Write-Host ""
Write-Host "Tarefa '$NomeTarefa' agendada com sucesso: toda segunda-feira as 07:00." -ForegroundColor Green
Write-Host "Executa: coleta -> prepara dados da API -> git push (redeploy automatico no Render)"
Write-Host "Logs de cada execucao: $PastaLogs\"
Write-Host ""
Write-Host "Para rodar manualmente agora (teste):"
Write-Host "  Start-ScheduledTask -TaskName '$NomeTarefa'"
Write-Host ""
Write-Host "Para ver o log mais recente logo apos um teste manual:"
Write-Host "  Get-ChildItem '$PastaLogs' | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content"
Write-Host ""
Write-Host "Para remover a tarefa:"
Write-Host "  Unregister-ScheduledTask -TaskName '$NomeTarefa' -Confirm:`$false"

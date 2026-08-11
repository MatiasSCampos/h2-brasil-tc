# =============================================================================
# Roda a coleta completa, prepara os dados enxutos para a API e publica no
# GitHub -- o push para o repositório dispara o redeploy automático no
# Render, então a API pública fica com os dados desta semana.
#
# Pré-requisito (só uma vez): o repositório já precisa estar configurado com
# 'git remote add origin <url-do-seu-repo>' e você já ter feito login
# (git config user.name / user.email, e autenticação configurada).
# =============================================================================

$PastaProjeto = $PSScriptRoot
Set-Location $PastaProjeto

Write-Host "== 1/4: Rodando o coletor de dados ==" -ForegroundColor Cyan
python coletor_h2_brasil.py
if ($LASTEXITCODE -ne 0) {
    Write-Error "coletor_h2_brasil.py terminou com erro (codigo $LASTEXITCODE). Abortando antes de publicar."
    exit 1
}

Write-Host "== 2/4: Preparando dados enxutos para a API ==" -ForegroundColor Cyan
python preparar_dados_api.py
if ($LASTEXITCODE -ne 0) {
    Write-Error "preparar_dados_api.py terminou com erro. Abortando antes de publicar."
    exit 1
}

Write-Host "== 3/4: Verificando se há mudança para publicar ==" -ForegroundColor Cyan
git add dados_api
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "  Nenhuma mudança nos dados desde a última publicação -- nada para enviar." -ForegroundColor Yellow
    exit 0
}

Write-Host "== 4/4: Publicando no GitHub (dispara redeploy automático no Render) ==" -ForegroundColor Cyan
$dataHora = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "Atualizacao automatica de dados - $dataHora"
git push

Write-Host ""
Write-Host "Publicado. O Render deve redeployar sozinho em alguns minutos." -ForegroundColor Green
Write-Host "Acompanhe em: https://dashboard.render.com"

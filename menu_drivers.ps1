# ================================================================================
# SCRIPT: Menu Principal - Sistema de Instalacao de Impressoras Samsung
# VERSAO: 2.1
# ================================================================================

if (-not ([Security.Principal.WindowsPrincipal]
    [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {

    Write-Host "`n[ERRO] Execute como ADMINISTRADOR" -ForegroundColor Red
    Read-Host "ENTER para sair"
    exit
}

Set-ExecutionPolicy Bypass -Scope Process -Force

$Config = @{
    Usuario     = "vitorarrudaa"
    Repositorio = "arquivos"
    Branch      = "main"
}
$Config.BaseUrl = "https://raw.githubusercontent.com/$($Config.Usuario)/$($Config.Repositorio)/$($Config.Branch)"

$Paths = @{
    Raiz  = "$env:USERPROFILE\Downloads\Suporte_Tech3"
    CSV   = "$env:USERPROFILE\Downloads\Suporte_Tech3\dados_impressoras.csv"
    Motor = "$env:USERPROFILE\Downloads\Suporte_Tech3\instalar_universal.ps1"
}

if (-not (Test-Path $Paths.Raiz)) {
    New-Item $Paths.Raiz -ItemType Directory -Force | Out-Null
}

function Sync-GitHubFiles {
    Write-Host "`n[INFO] Sincronizando arquivos..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest "$($Config.BaseUrl)/dados_impressoras.csv" -OutFile $Paths.CSV -ErrorAction Stop
        Invoke-WebRequest "$($Config.BaseUrl)/instalar_universal.ps1" -OutFile $Paths.Motor -ErrorAction Stop
    }
    catch {
        Write-Host "[ERRO] Falha ao baixar arquivos" -ForegroundColor Red
        Read-Host "ENTER para sair"
        exit
    }
}

do {
    Clear-Host
    Write-Host "=== MENU IMPRESSORAS SAMSUNG ===" -ForegroundColor Cyan
    Write-Host "1 - Sincronizar arquivos"
    Write-Host "2 - Instalar impressora"
    Write-Host "3 - Apagar arquivos locais"
    Write-Host "0 - Sair"

    $op = Read-Host "Opcao"

    switch ($op) {
        1 { Sync-GitHubFiles }
        2 { powershell -ExecutionPolicy Bypass -File $Paths.Motor }
        3 {
            Remove-Item $Paths.Raiz -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "Arquivos removidos."
            Pause
        }
    }
} while ($op -ne 0)

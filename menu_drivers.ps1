# ================================================================================
# MENU DE SELECAO DE IMPRESSORAS (Compativel com v4.0)
# ================================================================================

# Verificacao de privilegios administrativos
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "ATENCAO: Este script precisa ser executado como ADMINISTRADOR!" -ForegroundColor Red
    Write-Host "Por favor, clique com botao direito e selecione 'Executar como Administrador'" -ForegroundColor Yellow
    Read-Host "Pressione ENTER para sair"
    exit
}

# Configura o caminho base (mesma pasta do script)
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$csvPath = Join-Path $scriptPath "dados_impressoras.csv"
$universalScript = Join-Path $scriptPath "instalar_universal.ps1"

if (-not (Test-Path $csvPath)) {
    Write-Host "ERRO: Arquivo dados_impressoras.csv nao encontrado!" -ForegroundColor Red
    Read-Host "Pressione ENTER para sair"
    exit
}

# Le o arquivo CSV
try {
    $conteudo = Get-Content $csvPath -Encoding UTF8
} catch {
    $conteudo = Get-Content $csvPath # Fallback se falhar UTF8
}

# Remove cabecalho
$linhas = $conteudo | Select-Object -Skip 1 | Where-Object { $_ -ne "" }

# Processa os dados para um array de objetos
$listaImpressoras = @()

foreach ($linha in $linhas) {
    $dados = $linha.Split(';')
    
    # Valida se a linha tem colunas suficientes (Agora sao 11 colunas)
    if ($dados.Count -ge 11) {
        $obj = [PSCustomObject]@{
            Modelo         = $dados[0]
            UrlPrint       = $dados[1]
            TemScan        = $dados[2]
            UrlScan        = $dados[3]
            UrlEPM         = $dados[4]  # NOVA COLUNA
            UrlEDC         = $dados[5]  # NOVA COLUNA
            FiltroDriver   = $dados[6]
            InstalarPrint  = $dados[7]
            InstalarScan   = $dados[8]
            InstalarEPM    = $dados[9]
            InstalarEDC    = $dados[10]
        }
        $listaImpressoras += $obj
    }
}

# Loop do Menu
do {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      INSTALADOR DE DRIVERS SAMSUNG       " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
    
    for ($i = 0; $i -lt $listaImpressoras.Count; $i++) {
        Write-Host "[$($i+1)] $($listaImpressoras[$i].Modelo)"
    }
    Write-Host "[S] Sair"
    Write-Host ""
    
    $escolha = Read-Host "Selecione o numero da impressora"
    
    if ($escolha -eq "S" -or $escolha -eq "s") {
        exit
    }
    
    if ($escolha -match "^\d+$" -and [int]$escolha -le $listaImpressoras.Count -and [int]$escolha -gt 0) {
        $selecionada = $listaImpressoras[[int]$escolha - 1]
        
        Write-Host ""
        Write-Host "Iniciando instalacao para: $($selecionada.Modelo)" -ForegroundColor Yellow
        Write-Host "Carregando motor de instalacao..." -ForegroundColor Gray
        
        # Chama o script universal passando os parametros, incluindo as novas URLs
        # Nota: Aspas simples sao usadas para envolver strings para evitar erro com espacos
        & $universalScript `
            -modelo "$($selecionada.Modelo)" `
            -urlPrint "$($selecionada.UrlPrint)" `
            -temScan "$($selecionada.TemScan)" `
            -urlScan "$($selecionada.UrlScan)" `
            -urlEPM "$($selecionada.UrlEPM)" `
            -urlEDC "$($selecionada.UrlEDC)" `
            -filtroDriverWindows "$($selecionada.FiltroDriver)" `
            -instalarPrint ([bool]::Parse($selecionada.InstalarPrint)) `
            -instalarScan ([bool]::Parse($selecionada.InstalarScan)) `
            -instalarEPM ([bool]::Parse($selecionada.InstalarEPM)) `
            -instalarEDC ([bool]::Parse($selecionada.InstalarEDC))
            
        Write-Host ""
        Read-Host "Pressione ENTER para voltar ao menu"
    }
    else {
        Write-Host "Opcao invalida!" -ForegroundColor Red
        Start-Sleep -Seconds 1
    }
    
} while ($true)

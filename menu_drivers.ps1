# ================================================================================
# SCRIPT: Menu Principal - Sistema de Instalacao de Impressoras
# VERSAO: 3.0
# DESCRICAO: Menu interativo para instalacao de drivers de impressoras
# ================================================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "`n[ERRO] Este script requer privilegios de ADMINISTRADOR" -ForegroundColor Red
    Write-Host "Abra o PowerShell como Administrador e execute novamente.`n" -ForegroundColor Yellow
    Read-Host "Pressione ENTER para sair"
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
    Raiz    = "$env:USERPROFILE\Downloads\Suporte_Tech3"
    CSV     = "$env:USERPROFILE\Downloads\Suporte_Tech3\dados_impressoras.csv"
    Motor   = "$env:USERPROFILE\Downloads\Suporte_Tech3\instalar_universal.ps1"
    Drivers = "$env:USERPROFILE\Downloads\Instalacao_Impressoras"
}

if (-not (Test-Path $Paths.Raiz)) {
    New-Item $Paths.Raiz -ItemType Directory -Force | Out-Null
}

function Write-Mensagem {
    param(
        [Parameter(Mandatory=$true)][string]$texto,
        [ValidateSet("Info","Sucesso","Aviso","Erro","Titulo")][string]$tipo = "Info"
    )

    $cor = switch ($tipo) {
        "Info"    { "Gray" }
        "Sucesso" { "Green" }
        "Aviso"   { "Yellow" }
        "Erro"    { "Red" }
        "Titulo"  { "Cyan" }
    }

    $prefixo = switch ($tipo) {
        "Sucesso" { "[OK]" }
        "Aviso"   { "[AVISO]" }
        "Erro"    { "[ERRO]" }
        default    { "[INFO]" }
    }

    Write-Host "$prefixo $texto" -ForegroundColor $cor
}

function Sync-GitHubFiles {
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "  SINCRONIZANDO ARQUIVOS..." -ForegroundColor Cyan
    Write-Host "================================================`n" -ForegroundColor Cyan

    try {
        $wc = New-Object System.Net.WebClient

        Write-Host "Baixando lista de impressoras..." -ForegroundColor Gray
        $wc.DownloadFile("$($Config.BaseUrl)/dados_impressoras.csv", $Paths.CSV)

        Write-Host "Baixando motor de instalacao..." -ForegroundColor Gray
        $wc.DownloadFile("$($Config.BaseUrl)/instalar_universal.ps1", $Paths.Motor)

        Write-Mensagem "Arquivos sincronizados com sucesso!" "Sucesso"
        Write-Host ""
        return $true
    }
    catch {
        Write-Mensagem "Falha ao sincronizar arquivos do GitHub" "Erro"
        Write-Host "  Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Verifique sua conexao com a internet e o repositorio.`n" -ForegroundColor Yellow
        Read-Host "Pressione ENTER para sair"
        return $false
    }
}

function Get-PrinterData {
    try {
        $dados = Import-Csv -Path $Paths.CSV -Delimiter "," -ErrorAction Stop
        if ($dados.Count -eq 0) {
            throw "CSV esta vazio"
        }
        return $dados
    }
    catch {
        Write-Mensagem "Falha ao carregar dados das impressoras" "Erro"
        Write-Host "  Detalhes: $($_.Exception.Message)`n" -ForegroundColor Red
        Read-Host "Pressione ENTER para sair"
        exit
    }
}

function Show-ModelMenu {
    param($listaModelos)

    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "   TECH3 - INSTALACAO DE IMPRESSORAS" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan

    $fabricantes = $listaModelos | Select-Object -ExpandProperty Fabricante -Unique

    foreach ($fabricante in $fabricantes) {
        Write-Host ""
        Write-Host "  --- $fabricante ---" -ForegroundColor Magenta
        $modelosFabricante = $listaModelos | Where-Object { $_.Fabricante -eq $fabricante }

        foreach ($item in $modelosFabricante) {
            $temScan = if ($item.TemScan -eq "S") { "(Impressao + Scan)" } else { "(Apenas Impressao)" }
            Write-Host "  $($item.ID)) $($item.Modelo) $temScan" -ForegroundColor White
        }
    }

    Write-Host ""
    Write-Host "  Q) Sair" -ForegroundColor Gray
    Write-Host ""
}

function Show-ConfigMenu {
    param($modelo)

    Clear-Host
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "  MODELO SELECIONADO: $($modelo.Modelo)" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  1) Instalacao Completa (Todos os componentes)" -ForegroundColor White
    Write-Host "  2) Instalacao Personalizada (Escolher componentes)" -ForegroundColor White
    Write-Host ""
    Write-Host "  V) Voltar ao menu anterior" -ForegroundColor Gray
    Write-Host ""
}

function Show-CustomMenu {
    param($modelo)

    Clear-Host
    Write-Host "================================================" -ForegroundColor Yellow
    Write-Host "  INSTALACAO PERSONALIZADA: $($modelo.Modelo)" -ForegroundColor Yellow
    Write-Host "================================================" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1) Driver de Impressao" -ForegroundColor White

    if ($modelo.TemScan -eq "S") {
        Write-Host "  2) Driver de Digitalizacao (Scan)" -ForegroundColor White
        Write-Host "  3) Easy Document Creator (EDC)" -ForegroundColor White
        Write-Host "  4) Easy Printer Manager (EPM)" -ForegroundColor White
    } else {
        Write-Host "  2) Easy Printer Manager (EPM)" -ForegroundColor White
    }

    Write-Host ""
    Write-Host "  V) Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Invoke-Installation {
    param([hashtable]$parametros)

    try {
        & $Paths.Motor @parametros
    }
    catch {
        Write-Mensagem "Falha ao executar instalacao" "Erro"
        Write-Host "  Detalhes: $($_.Exception.Message)`n" -ForegroundColor Red
        Read-Host "Pressione ENTER para continuar"
    }
}

if (-not (Sync-GitHubFiles)) {
    exit
}

$listaImpressoras = Get-PrinterData

do {
    Show-ModelMenu -listaModelos $listaImpressoras
    $escolhaModelo = Read-Host "Escolha o ID do modelo ou (Q) para sair"

    if ($escolhaModelo -ieq "Q") {
        Write-Host ""
        do {
            $limparDrivers = Read-Host "Deseja apagar os drivers baixados? [S/N]"

            if ($limparDrivers -ieq "S") {
                Remove-Item $Paths.Raiz -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item $Paths.Drivers -Recurse -Force -ErrorAction SilentlyContinue
                Write-Mensagem "Todos os arquivos foram removidos!" "Sucesso"
                Write-Host ""
                break
            }
            elseif ($limparDrivers -ieq "N") {
                Remove-Item $Paths.Raiz -Recurse -Force -ErrorAction SilentlyContinue
                break
            }
            else {
                Write-Mensagem "Digite apenas S ou N" "Aviso"
            }
        } while ($true)
        break
    }

    $modeloSelecionado = $listaImpressoras | Where-Object { $_.ID -eq $escolhaModelo }

    if (-not $modeloSelecionado) {
        Write-Mensagem "Opcao invalida! Tente novamente." "Aviso"
        Start-Sleep -Seconds 2
        continue
    }

    $continuarNoModelo = $true

    while ($continuarNoModelo) {
        Show-ConfigMenu -modelo $modeloSelecionado
        $tipoInstalacao = Read-Host "Escolha uma opcao"

        if ($tipoInstalacao -ieq "V") {
            $continuarNoModelo = $false
            break
        }

        $params = @{
            modelo              = $modeloSelecionado.Modelo
            fabricante          = $modeloSelecionado.Fabricante
            urlPrint            = $modeloSelecionado.UrlPrint
            temScan             = $modeloSelecionado.TemScan
            urlScan             = $modeloSelecionado.UrlScan
            filtroDriverWindows = $modeloSelecionado.FiltroDriver
            instalarPrint       = $false
            instalarScan        = $false
            instalarEPM         = $false
            instalarEDC         = $false
        }

        if ($tipoInstalacao -eq "1") {
            $params.instalarPrint = $true
            $params.instalarEPM = $true

            if ($modeloSelecionado.TemScan -eq "S") {
                $params.instalarScan = $true
                $params.instalarEDC = $true
            }

            Invoke-Installation -parametros $params
            $continuarNoModelo = $false
        }
        elseif ($tipoInstalacao -eq "2") {
            $sairPersonalizado = $false

            while (-not $sairPersonalizado) {
                Show-CustomMenu -modelo $modeloSelecionado
                $componenteEscolhido = Read-Host "Escolha o componente"

                if ($componenteEscolhido -ieq "V") {
                    $sairPersonalizado = $true
                    break
                }

                $params.instalarPrint = $false
                $params.instalarScan = $false
                $params.instalarEDC = $false
                $params.instalarEPM = $false

                if ($modeloSelecionado.TemScan -eq "S") {
                    switch ($componenteEscolhido) {
                        "1" { $params.instalarPrint = $true }
                        "2" { $params.instalarScan = $true }
                        "3" { $params.instalarEDC = $true }
                        "4" { $params.instalarEPM = $true }
                        default {
                            Write-Mensagem "Opcao invalida!" "Aviso"
                            Start-Sleep -Seconds 2
                            continue
                        }
                    }
                } else {
                    switch ($componenteEscolhido) {
                        "1" { $params.instalarPrint = $true }
                        "2" { $params.instalarEPM = $true }
                        default {
                            Write-Mensagem "Opcao invalida!" "Aviso"
                            Start-Sleep -Seconds 2
                            continue
                        }
                    }
                }

                Invoke-Installation -parametros $params
            }
        }
        else {
            Write-Mensagem "Opcao invalida! Tente novamente." "Aviso"
            Start-Sleep -Seconds 2
        }
    }

} while ($true)

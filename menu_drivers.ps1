# ================================================================================
# SCRIPT: Menu Principal - Sistema de Instalacao de Impressoras
# VERSAO: 3.0
# DESCRICAO: Menu interativo para instalacao de drivers de impressoras
# ================================================================================

# --- VERIFICACAO DE PRIVILEGIOS ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host ""
    Write-Host "[ERRO] Este script requer privilegios de ADMINISTRADOR" -ForegroundColor Red
    Write-Host "Abra o PowerShell como Administrador e execute novamente." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Pressione ENTER para sair"
    exit
}

Set-ExecutionPolicy Bypass -Scope Process -Force

# --- CONFIGURACAO DO REPOSITORIO GITHUB ---
$Config = @{
    Usuario     = "vitorarrudaa"
    Repositorio = "arquivos"
    Branch      = "main"
}
$Config.BaseUrl = "https://raw.githubusercontent.com/$($Config.Usuario)/$($Config.Repositorio)/$($Config.Branch)"

# --- DIRETORIOS LOCAIS ---
$Paths = @{
    Raiz  = "$env:USERPROFILE\Downloads\Suporte_Tech3"
    CSV   = "$env:USERPROFILE\Downloads\Suporte_Tech3\dados_impressoras.csv"
    Motor = "$env:USERPROFILE\Downloads\Suporte_Tech3\instalar_universal.ps1"
}

if (-not (Test-Path $Paths.Raiz)) {
    New-Item $Paths.Raiz -ItemType Directory -Force | Out-Null
}

# ================================================================================
# FUNCOES VISUAIS
# ================================================================================

$LARGURA = 50

function Write-Borda {
    Write-Host ("=" * $LARGURA) -ForegroundColor DarkGray
}

function Write-Separador {
    Write-Host ("-" * $LARGURA) -ForegroundColor DarkGray
}

function Write-SeparadorFabricante {
    param([string]$fabricante)
    $prefixo = "-- $fabricante "
    $resto = "-" * ($LARGURA - $prefixo.Length)
    Write-Host "$prefixo$resto" -ForegroundColor Yellow
}

function Write-Cabecalho {
    param([string]$titulo, [string]$subtitulo = "")
    Write-Host ""
    Write-Borda
    if ($subtitulo -ne "") {
        Write-Host "  $titulo > $subtitulo" -ForegroundColor Cyan
    } else {
        Write-Host "  $titulo" -ForegroundColor Cyan
    }
    Write-Borda
    Write-Host ""
}

function Write-Cabecalho-Principal {
    param([string]$titulo)
    Write-Host ""
    Write-Borda
    $espacos = [math]::Max(0, [math]::Floor(($LARGURA - $titulo.Length) / 2))
    Write-Host (" " * $espacos + $titulo) -ForegroundColor Magenta
    Write-Borda
    Write-Host ""
}

# ================================================================================
# FUNCOES DE SISTEMA
# ================================================================================

function Sync-GitHubFiles {
    Write-Host "  Sincronizando arquivos com GitHub..." -ForegroundColor Cyan

    try {
        Invoke-WebRequest -Uri "$($Config.BaseUrl)/dados_impressoras.csv" `
                         -OutFile $Paths.CSV `
                         -ErrorAction Stop `
                         -UseBasicParsing

        Invoke-WebRequest -Uri "$($Config.BaseUrl)/instalar_universal.ps1" `
                         -OutFile $Paths.Motor `
                         -ErrorAction Stop `
                         -UseBasicParsing

        Write-Host "  [OK] Arquivos sincronizados com sucesso!" -ForegroundColor Green
        Write-Host ""
        return $true
    }
    catch {
        Write-Host ""
        Write-Host "  [ERRO] Falha ao sincronizar arquivos do GitHub" -ForegroundColor Red
        Write-Host "  Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Verifique sua conexao com a internet e o repositorio." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "  Pressione ENTER para sair"
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
        Write-Host ""
        Write-Host "  [ERRO] Falha ao carregar dados das impressoras" -ForegroundColor Red
        Write-Host "  Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Read-Host "  Pressione ENTER para sair"
        exit
    }
}

# ================================================================================
# FUNCOES DE MENU
# ================================================================================

function Show-ModelMenu {
    param($listaModelos)

    Clear-Host
    Write-Cabecalho-Principal "TECH3 - INSTALADOR DE IMPRESSORAS"

    # Agrupar modelos por fabricante
    $fabricantes = $listaModelos | Select-Object -ExpandProperty Fabricante -Unique

    foreach ($fab in $fabricantes) {
        Write-SeparadorFabricante $fab
        $modelos = $listaModelos | Where-Object { $_.Fabricante -eq $fab }

        foreach ($item in $modelos) {
            $capacidade = if ($item.TemScan -eq "S") { "Impressao + Scan" } else { "Impressao" }
            $nome = "  $($item.ID)) $($item.Modelo)"
            $espacos = " " * ([math]::Max(1, 34 - $nome.Length))
            Write-Host "$nome$espacos$capacidade" -ForegroundColor White
        }

        Write-Host ""
    }

    Write-Separador
    Write-Host "  Q) Sair" -ForegroundColor Gray
    Write-Host ""
}

function Show-ConfigMenu {
    param($modelo)

    Clear-Host
    Write-Cabecalho "$($modelo.Fabricante) $($modelo.Modelo)"

    Write-Host "  1) Instalacao Completa" -ForegroundColor White
    Write-Host "  2) Instalacao Personalizada" -ForegroundColor White
    Write-Host ""
    Write-Host "  V) Voltar" -ForegroundColor Gray
    Write-Host ""
}

function Show-CustomMenu {
    param($modelo)

    Clear-Host
    Write-Cabecalho "$($modelo.Fabricante) $($modelo.Modelo)" "Instalacao Personalizada"

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
        Write-Host ""
        Write-Host "  [ERRO] Falha ao executar instalacao" -ForegroundColor Red
        Write-Host "  Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Read-Host "  Pressione ENTER para continuar"
    }
}

# ================================================================================
# INICIO
# ================================================================================

Clear-Host
Write-Borda
Write-Host ("  TECH3 - INSTALADOR DE IMPRESSORAS") -ForegroundColor Magenta
Write-Borda
Write-Host ""

if (-not (Sync-GitHubFiles)) {
    exit
}

$listaImpressoras = Get-PrinterData

# ================================================================================
# LOOP PRINCIPAL DO MENU
# ================================================================================

do {
    Show-ModelMenu -listaModelos $listaImpressoras
    $escolhaModelo = Read-Host "  Escolha o numero do modelo ou (Q) para sair"

    if ($escolhaModelo -eq "Q" -or $escolhaModelo -eq "q") {
        Write-Host ""

        do {
            $limparDrivers = Read-Host "  Deseja apagar os drivers baixados? (S/N)"

            if ($limparDrivers -eq "S" -or $limparDrivers -eq "s") {
                Remove-Item "$env:USERPROFILE\Downloads\Suporte_Tech3" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item "$env:USERPROFILE\Downloads\Instalacao_Samsung" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host ""
                Write-Host "  [OK] Todos os arquivos foram removidos!" -ForegroundColor Green
                Write-Host ""
                break
            }
            elseif ($limparDrivers -eq "N" -or $limparDrivers -eq "n") {
                Remove-Item "$env:USERPROFILE\Downloads\Suporte_Tech3" -Recurse -Force -ErrorAction SilentlyContinue
                break
            }
            else {
                Write-Host "  [AVISO] Digite apenas S ou N" -ForegroundColor Yellow
            }
        } while ($true)

        break
    }

    $modeloSelecionado = $listaImpressoras | Where-Object { $_.ID -eq $escolhaModelo }

    if (-not $modeloSelecionado) {
        Write-Host ""
        Write-Host "  [AVISO] Opcao invalida! Tente novamente." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        continue
    }

    $continuarNoModelo = $true

    while ($continuarNoModelo) {
        Show-ConfigMenu -modelo $modeloSelecionado
        $tipoInstalacao = Read-Host "  Escolha uma opcao"

        if ($tipoInstalacao -eq "V" -or $tipoInstalacao -eq "v") {
            $continuarNoModelo = $false
            break
        }

        $params = @{
            modelo              = $modeloSelecionado.Modelo
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
            $params.instalarEPM   = $true

            if ($modeloSelecionado.TemScan -eq "S") {
                $params.instalarScan = $true
                $params.instalarEDC  = $true
            }

            Invoke-Installation -parametros $params
            $continuarNoModelo = $false
        }
        elseif ($tipoInstalacao -eq "2") {
            $sairPersonalizado = $false

            while (-not $sairPersonalizado) {
                Show-CustomMenu -modelo $modeloSelecionado
                $componenteEscolhido = Read-Host "  Escolha um componente"

                if ($componenteEscolhido -eq "V" -or $componenteEscolhido -eq "v") {
                    $sairPersonalizado = $true
                    break
                }

                $params.instalarPrint = $false
                $params.instalarScan  = $false
                $params.instalarEDC   = $false
                $params.instalarEPM   = $false

                if ($modeloSelecionado.TemScan -eq "S") {
                    switch ($componenteEscolhido) {
                        "1" { $params.instalarPrint = $true }
                        "2" { $params.instalarScan  = $true }
                        "3" { $params.instalarEDC   = $true }
                        "4" { $params.instalarEPM   = $true }
                        default {
                            Write-Host ""
                            Write-Host "  [AVISO] Opcao invalida!" -ForegroundColor Yellow
                            Start-Sleep -Seconds 2
                            continue
                        }
                    }
                } else {
                    switch ($componenteEscolhido) {
                        "1" { $params.instalarPrint = $true }
                        "2" { $params.instalarEPM   = $true }
                        default {
                            Write-Host ""
                            Write-Host "  [AVISO] Opcao invalida!" -ForegroundColor Yellow
                            Start-Sleep -Seconds 2
                            continue
                        }
                    }
                }

                Invoke-Installation -parametros $params
            }
        }
        else {
            Write-Host ""
            Write-Host "  [AVISO] Opcao invalida! Tente novamente." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }
    }

} while ($true)

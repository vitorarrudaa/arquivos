# 1. Identifica automaticamente a pasta onde o script esta sendo executado
$caminhoBase = $PSScriptRoot

# Definicao dos nomes dos arquivos
$driverImp     = Join-Path $caminhoBase "Driver_M4070_Print.exe"
$driverScan    = Join-Path $caminhoBase "Driver_M4070_Scan.exe"
$easyCreator   = Join-Path $caminhoBase "EasyDocumentCreator.exe"
$easyManager   = Join-Path $caminhoBase "EasyPrinterManager.exe"

# --- FUNÇÃO DE VERIFICAÇÃO DE INSTALAÇÃO (Somente para Utilitários) ---
function Test-JaInstalado {
    param([string]$nomePrograma)
    $chaves = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $resultado = Get-ItemProperty $chaves -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$nomePrograma*" }
    return [bool]$resultado
}

# 2. Solicita os dados ao usuario
$novoNome  = Read-Host "Digite o NOME desejado para a impressora"
$printerIP = Read-Host "Digite o endereco IP da impressora"

# Função padrão para execução
function Executar-Instalador {
    param([string]$caminho, [string]$nomeExibicao)
    if (Test-Path $caminho) {
        Write-Host "-> Instalando $nomeExibicao..." -ForegroundColor Cyan
        Start-Process -FilePath $caminho -ArgumentList "/S" -Wait -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    } else {
        Write-Host "-> Erro: Arquivo $nomeExibicao nao encontrado em $caminho" -ForegroundColor Red
    }
}

# --- [1/6] DRIVER DE IMPRESSÃO ---
Write-Host "`n[1/6] Iniciando Driver de Impressao..." -ForegroundColor White
Executar-Instalador $driverImp "Driver de Impressao"

# --- [2/6] AJUSTE DE NOME E PORTA ---
Write-Host "[2/6] Configurando Nome e IP..." -ForegroundColor Yellow
$impressoraPadrao = Get-Printer | Where-Object {$_.Name -like "*Samsung M337x 387x 407x*"} | Select-Object -First 1

if ($impressoraPadrao) {
    if (-not (Get-PrinterPort -Name $printerIP -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $printerIP -PrinterHostAddress $printerIP
    }
    Set-Printer -Name $impressoraPadrao.Name -PortName $printerIP
    Rename-Printer -Name $impressoraPadrao.Name -NewName $novoNome
    Write-Host "Configuracao de IP e Nome aplicada." -ForegroundColor Green
}

# --- [3/6] DRIVER DE SCAN ---
Write-Host "[3/6] Iniciando Driver de Digitalizacao..." -ForegroundColor White
Executar-Instalador $driverScan "Driver de Scan"

# --- [4/6] EASY DOCUMENT CREATOR ---
Write-Host "[4/6] Verificando Easy Document Creator..." -ForegroundColor White
if (-not (Test-JaInstalado "Easy Document Creator")) {
    Executar-Instalador $easyCreator "Easy Document Creator"
} else {
    Write-Host "-> Easy Document Creator ja detectado no sistema. Pulando..." -ForegroundColor Green
}

# --- [5/6] EASY PRINTER MANAGER ---
Write-Host "[5/6] Verificando Easy Printer Manager..." -ForegroundColor White
if (-not (Test-JaInstalado "Easy Printer Manager")) {
    Executar-Instalador $easyManager "Easy Printer Manager"
} else {
    Write-Host "-> Easy Printer Manager ja detectado no sistema. Pulando..." -ForegroundColor Green
}

# --- [6/6] LIMPEZA ---
Write-Host "[6/6] Verificando duplicatas e finalizando..." -ForegroundColor Yellow
Get-Printer | Where-Object {$_.Name -like "*Samsung M337x 387x 407x* (Copiar*"} | Remove-Printer

Write-Host "`n=================================================" -ForegroundColor White
Write-Host "   INSTALACAO CONCLUIDA COM SUCESSO!" -ForegroundColor Green
Write-Host "   Impressora: $novoNome | IP: $printerIP"
Write-Host "=================================================" -ForegroundColor White

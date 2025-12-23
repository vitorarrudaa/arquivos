# 1. Identifica automaticamente a pasta onde o script está sendo executado
$caminhoBase = $PSScriptRoot

# Definição dos nomes dos arquivos (Devem estar na mesma pasta do script)
$driverImp     = Join-Path $caminhoBase "Driver_Impressao.exe"
$driverScan    = Join-Path $caminhoBase "Driver_Digitalizacao.exe"
$easyCreator   = Join-Path $caminhoBase "Easy_Document_Creator.exe"
$easyManager   = Join-Path $caminhoBase "Easy_Printer_Manager.exe"

# 2. Solicita os dados ao usuário
$novoNome  = Read-Host "Digite o NOME desejado para a impressora"
$printerIP = Read-Host "Digite o endereço IP da impressora"

Write-Host "`n[1/6] Instalando Driver de Impressão..." -ForegroundColor Cyan
# Executa o instalador que está na mesma pasta
Start-Process -FilePath $driverImp -ArgumentList "/S" -Wait
Start-Sleep -Seconds 3

# --- AJUSTE DE NOME E PORTA ---
$impressoraPadrao = Get-Printer | Where-Object {$_.Name -like "*Samsung M337x 387x 407x*"} | Select-Object -First 1

if ($impressoraPadrao) {
    Write-Host "[2/6] Configurando Nome e IP..." -ForegroundColor Yellow
    if (-not (Get-PrinterPort -Name $printerIP -ErrorAction SilentlyContinue)) {
        Add-PrinterPort -Name $printerIP -PrinterHostAddress $printerIP
    }
    Set-Printer -Name $impressoraPadrao.Name -PortName $printerIP
    Rename-Printer -Name $impressoraPadrao.Name -NewName $novoNome
}

# --- INSTALAÇÃO DOS DEMAIS COMPONENTES (MESMA PASTA) ---
Write-Host "[3/6] Instalando Driver de Digitalização..." -ForegroundColor Cyan
Start-Process -FilePath $driverScan -ArgumentList "/S" -Wait

Write-Host "[4/6] Instalando Easy Document Creator..." -ForegroundColor Cyan
Start-Process -FilePath $easyCreator -ArgumentList "/S" -Wait

Write-Host "[5/6] Instalando Easy Printer Manager..." -ForegroundColor Cyan
Start-Process -FilePath $easyManager -ArgumentList "/S" -Wait

# --- LIMPEZA ---
Write-Host "[6/6] Verificando duplicatas..." -ForegroundColor Yellow
Get-Printer | Where-Object {$_.Name -like "*Samsung M337x 387x 407x* (Copiar*"} | Remove-Printer

Write-Host "`n=================================================" -ForegroundColor White
Write-Host "  INSTALAÇÃO CONCLUÍDA COM SUCESSO!" -ForegroundColor Green
Write-Host "  Local dos drivers: $caminhoBase"
Write-Host "  Impressora: $novoNome | IP: $printerIP"
Write-Host "=================================================" -ForegroundColor White

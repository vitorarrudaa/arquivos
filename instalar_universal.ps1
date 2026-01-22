# ================================================================================
# MOTOR UNIVERSAL – SAMSUNG
# ================================================================================
param (
    [string]$modelo,
    [string]$urlPrint,
    [string]$temScan,
    [string]$urlScan,
    [string]$filtroDriverWindows,
    [bool]$instalarPrint,
    [bool]$instalarScan,
    [bool]$instalarEPM,
    [bool]$instalarEDC
)

$Config = @{
    UrlEPM      = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
    UrlEDC      = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
    CaminhoTemp = "$env:USERPROFILE\Downloads\Instalacao_Samsung"
    TempoEspera = 10
}

if (-not (Test-Path $Config.CaminhoTemp)) {
    New-Item $Config.CaminhoTemp -ItemType Directory -Force | Out-Null
}

function Write-Mensagem {
    param ($Texto, $Tipo = "Info")
    $cor = @{ Info="Gray"; Sucesso="Green"; Aviso="Yellow"; Erro="Red" }[$Tipo]
    Write-Host "[$Tipo] $Texto" -ForegroundColor $cor
}

# ===============================
# ETAPA 1 – ENTRADA DE DADOS
# ===============================
do {
    $nomeImpressora = Read-Host "Nome da impressora"
} while ([string]::IsNullOrWhiteSpace($nomeImpressora))

do {
    $enderecoIP = Read-Host "IP da impressora"

    # FIX CRITICO – regex corrigido
    $ipValido = $enderecoIP -match '^\d{1,3}(\.\d{1,3}){3}$'
    if (-not $ipValido) {
        Write-Mensagem "IP invalido" "Erro"
    }
} while (-not $ipValido)

# ===============================
# ETAPA 2 – CONFLITOS
# ===============================
$impExistente = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
$portaExistente = Get-PrinterPort | Where-Object { $_.PrinterHostAddress -eq $enderecoIP }

if ($impExistente -or $portaExistente) {
    Write-Mensagem "Conflito detectado:" "Aviso"
    if ($impExistente) { Write-Host "- Impressora existente: $($impExistente.Name)" }
    if ($portaExistente) { Write-Host "- IP já utilizado: $enderecoIP" }

    $acao = Read-Host "Remover e continuar? (S/N)"
    if ($acao -ne "S") { exit }

    if ($impExistente) { Remove-Printer -Name $impExistente.Name -Confirm:$false }
    if ($portaExistente) { Remove-PrinterPort -Name $portaExistente.Name -Confirm:$false }
}

# ===============================
# ETAPA 3 – DOWNLOAD DRIVER
# ===============================
$arquivoDriver = Join-Path $Config.CaminhoTemp "driver.exe"

if (-not (Test-Path $arquivoDriver) -or (Get-Item $arquivoDriver).Length -lt 1MB) {
    Write-Mensagem "Baixando driver..."
    Invoke-WebRequest $urlPrint -OutFile $arquivoDriver
}

# ===============================
# ETAPA 4 – INSTALACAO
# ===============================
Start-Process $arquivoDriver -ArgumentList "/s" -Wait

$porta = "IP_$enderecoIP"
Add-PrinterPort -Name $porta -PrinterHostAddress $enderecoIP
Add-Printer -Name $nomeImpressora -DriverName $filtroDriverWindows -PortName $porta

Write-Mensagem "Instalacao concluida com sucesso" "Sucesso"

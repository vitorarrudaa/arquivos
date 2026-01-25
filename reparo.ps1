# ================================================================================
# SCRIPT: Reparo de Impressoras Samsung
# VERSAO: 3.0 (Reformulado - Comandos Testados)
# DESCRICAO: Remove drivers e impressoras para reparo do sistema
# ================================================================================

param (
    [Parameter(Mandatory=$false)][string]$csvPath = "$env:USERPROFILE\Downloads\Suporte_Tech3\dados_impressoras.csv"
)

# --- VERIFICACAO DE PRIVILEGIOS ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "`n[ERRO] Este script requer privilegios de ADMINISTRADOR" -ForegroundColor Red
    Write-Host "Abra o PowerShell como Administrador e execute novamente.`n" -ForegroundColor Yellow
    Read-Host "Pressione ENTER para sair"
    exit
}

Set-ExecutionPolicy Bypass -Scope Process -Force

# ================================================================================
# FUNCOES AUXILIARES
# ================================================================================

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
        default   { "" }
    }
    
    Write-Host "$prefixo $texto" -ForegroundColor $cor
}

function Read-OpcaoValidada {
    param(
        [Parameter(Mandatory=$true)][string]$prompt,
        [Parameter(Mandatory=$true)][string[]]$opcoesValidas
    )
    
    do {
        $entrada = Read-Host $prompt
        if ($opcoesValidas -contains $entrada) {
            return $entrada
        }
        Write-Mensagem "Opcao invalida! Tente novamente." "Aviso"
    } while ($true)
}

function Get-PrinterData {
    if (-not (Test-Path $csvPath)) {
        Write-Mensagem "Arquivo CSV nao encontrado: $csvPath" "Erro"
        return $null
    }
    
    try {
        $dados = Import-Csv -Path $csvPath -Delimiter "," -ErrorAction Stop
        return $dados
    }
    catch {
        Write-Mensagem "Falha ao carregar dados do CSV: $($_.Exception.Message)" "Erro"
        return $null
    }
}

function Get-ImpressorasUsandoDriver {
    param([string]$nomeDriver)
    
    # Busca exata pelo nome do driver
    $impressoras = Get-Printer -ErrorAction SilentlyContinue | 
                   Where-Object { $_.DriverName -eq $nomeDriver }
    
    # Se não encontrar nada, tenta busca parcial (casos de nomes com variações)
    if ($impressoras.Count -eq 0) {
        $impressoras = Get-Printer -ErrorAction SilentlyContinue | 
                       Where-Object { $_.DriverName -like "*$nomeDriver*" }
    }
    
    return $impressoras
}

function Get-DriversOrfaos {
    param([array]$modelosCSV)
    
    $todosDriversSamsung = Get-PrinterDriver -ErrorAction SilentlyContinue | 
                          Where-Object { $_.Name -like "*Samsung*" }
    
    $driversCSV = @()
    foreach ($modelo in $modelosCSV) {
        $driversCSV += $modelo.FiltroDriver
    }
    
    $orfaos = @()
    foreach ($driver in $todosDriversSamsung) {
        $ehOrfao = $true
        foreach ($driverCSV in $driversCSV) {
            if ($driver.Name -eq $driverCSV) {
                $ehOrfao = $false
                break
            }
        }
        
        if ($ehOrfao) {
            $orfaos += $driver.Name
        }
    }
    
    return $orfaos
}

function Remove-ImpressorasEDriverEspecifico {
    param(
        [string]$nomeDriver
    )
    
    Write-Host "`n>> Analisando sistema..." -ForegroundColor Cyan
    
    # Busca impressoras que usam este driver
    $impressorasUsando = Get-Printer -ErrorAction SilentlyContinue | 
                         Where-Object { $_.DriverName -eq $nomeDriver }
    
    $totalImpressoras = if ($impressorasUsando) { $impressorasUsando.Count } else { 0 }
    
    Write-Host "   [INFO] Impressoras encontradas: $totalImpressoras" -ForegroundColor Gray
    
    if ($totalImpressoras -gt 0) {
        Write-Host "`n   Impressoras que serao removidas:" -ForegroundColor Yellow
        foreach ($imp in $impressorasUsando) {
            Write-Host "     - $($imp.Name) [$($imp.PortName)]" -ForegroundColor Gray
        }
    }
    
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "ATENCAO! ESSA OPERACAO VAI REMOVER:" -ForegroundColor Yellow
    Write-Host "   - $totalImpressoras IMPRESSORA(S)" -ForegroundColor Yellow
    Write-Host "   - 1 DRIVER(S)" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Yellow
    
    $confirmar = Read-OpcaoValidada "DESEJA PROSSEGUIR? [S/N]" @("S","s","N","n")
    
    if ($confirmar -eq "N" -or $confirmar -eq "n") {
        Write-Mensagem "`nOperacao cancelada" "Info"
        return $false
    }
    
    # PASSO 1: Remove impressoras
    if ($totalImpressoras -gt 0) {
        Write-Host "`n>> Removendo impressoras..." -ForegroundColor Cyan
        
        foreach ($impressora in $impressorasUsando) {
            Write-Host "   Removendo: $($impressora.Name)..." -ForegroundColor Gray
            try {
                Remove-Printer -Name $impressora.Name -Confirm:$false -ErrorAction Stop
                Write-Host "   [OK] Removida" -ForegroundColor Green
            }
            catch {
                Write-Host "   [ERRO] $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        Start-Sleep -Seconds 2
    }
    
    # PASSO 2: Limpa spooler
    Write-Host "`n>> Limpando spooler..." -ForegroundColor Cyan
    
    try {
        Stop-Service Spooler -Force -ErrorAction Stop
        Write-Host "   [OK] Spooler parado" -ForegroundColor Green
        
        Start-Sleep -Seconds 2
        
        Remove-Item "$env:SystemRoot\System32\Spool\Printers\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "   [OK] Cache limpo" -ForegroundColor Green
        
        Start-Service Spooler -ErrorAction Stop
        Write-Host "   [OK] Spooler reiniciado" -ForegroundColor Green
        
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Host "   [ERRO] Falha: $($_.Exception.Message)" -ForegroundColor Red
        
        try {
            Start-Service Spooler -ErrorAction SilentlyContinue
        }
        catch {}
    }
    
    # PASSO 3: Remove o driver
    Write-Host "`n>> Removendo driver..." -ForegroundColor Cyan
    
    try {
        Remove-PrinterDriver -Name $nomeDriver -ErrorAction Stop
        Write-Host "   [OK] Driver removido" -ForegroundColor Green
    }
    catch {
        Write-Host "   [ERRO] $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "REMOCAO CONCLUIDA COM SUCESSO!" -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Green
    
    return $true
}

function Show-MenuReparo {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "   TECH3 - REPARO DE IMPRESSORAS" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  1) Remover driver de impressora" -ForegroundColor White
    Write-Host "  2) Remover driver especifico" -ForegroundColor White
    Write-Host ""
    Write-Host "  V) Voltar ao menu principal" -ForegroundColor Gray
    Write-Host ""
}

function Invoke-RemoverDriverDeImpressora {
    param([array]$modelosCSV)
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   REMOVER DRIVER DE IMPRESSORA" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    foreach ($modelo in $modelosCSV) {
        $temScan = if ($modelo.TemScan -eq "S") { "(Impressao + Scan)" } else { "(Apenas Impressao)" }
        Write-Host "  $($modelo.ID)) $($modelo.Modelo) $temScan" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "  V) Voltar" -ForegroundColor Gray
    Write-Host ""
    
    $escolha = Read-Host "Escolha o modelo"
    
    if ($escolha -eq "V" -or $escolha -eq "v") {
        return
    }
    
    $modeloSelecionado = $modelosCSV | Where-Object { $_.ID -eq $escolha }
    
    if (-not $modeloSelecionado) {
        Write-Mensagem "`nOpcao invalida!" "Erro"
        Start-Sleep -Seconds 2
        return
    }
    
    $driverModelo = $modeloSelecionado.FiltroDriver
    $impressorasEncontradas = Get-ImpressorasUsandoDriver -nomeDriver $driverModelo
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   $($modeloSelecionado.Modelo.ToUpper()) - ANALISE DO SISTEMA" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    Write-Host "Driver do modelo: $driverModelo`n" -ForegroundColor White
    
    if ($impressorasEncontradas.Count -gt 0) {
        Write-Host "Impressoras encontradas usando este driver:" -ForegroundColor White
        foreach ($imp in $impressorasEncontradas) {
            Write-Host "  - $($imp.Name) ($($imp.PortName))" -ForegroundColor Gray
        }
    } else {
        Write-Host "Nenhuma impressora encontrada usando este driver." -ForegroundColor Yellow
        Read-Host "`nPressione ENTER para continuar"
        return
    }
    
    $sucesso = Remove-ImpressorasEDriverEspecifico -nomeDriver $driverModelo
    
    if ($sucesso) {
        $reinstalar = Read-OpcaoValidada "`nDeseja reinstalar agora? [S/N]" @("S","s","N","n")
        
        if ($reinstalar -eq "S" -or $reinstalar -eq "s") {
            $scriptInstalador = Join-Path (Split-Path $csvPath) "instalar_universal.ps1"
            
            if (Test-Path $scriptInstalador) {
                Write-Host "`nIniciando reinstalacao...`n" -ForegroundColor Cyan
                Start-Sleep -Seconds 2
                
                $params = @{
                    modelo = $modeloSelecionado.Modelo
                    urlPrint = $modeloSelecionado.UrlPrint
                    temScan = $modeloSelecionado.TemScan
                    urlScan = $modeloSelecionado.UrlScan
                    filtroDriverWindows = $modeloSelecionado.FiltroDriver
                    instalarPrint = $true
                    instalarScan = $false
                    instalarEPM = $false
                    instalarEDC = $false
                }
                
                & $scriptInstalador @params
            }
            else {
                Write-Mensagem "Script de instalacao nao encontrado" "Erro"
            }
        }
    }
    
    Read-Host "`nPressione ENTER para continuar"
}

function Invoke-RemoverDriverEspecifico {
    param([array]$modelosCSV)
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   REMOVER DRIVER ESPECIFICO" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Cria mapa de drivers
    $mapaDrivers = @{}
    $contador = 1
    
    # Bloco 1: Drivers cadastrados no CSV
    Write-Host "DRIVERS CADASTRADOS (do CSV):" -ForegroundColor Green
    
    foreach ($modelo in $modelosCSV) {
        $driver = Get-PrinterDriver -Name $modelo.FiltroDriver -ErrorAction SilentlyContinue
        
        if ($driver) {
            Write-Host "  $contador) $($modelo.FiltroDriver) ($($modelo.Modelo))" -ForegroundColor Cyan
            $mapaDrivers[$contador.ToString()] = @{
                Nome = $modelo.FiltroDriver
                Modelo = $modelo.Modelo
                Cadastrado = $true
                ModeloObj = $modelo
            }
            $contador++
        }
    }
    
    Write-Host ""
    
    # Bloco 2: Drivers órfãos
    $orfaos = Get-DriversOrfaos -modelosCSV $modelosCSV
    
    if ($orfaos.Count -gt 0) {
        Write-Host "OUTROS DRIVERS (nao cadastrados):" -ForegroundColor Yellow
        
        foreach ($orfao in $orfaos) {
            Write-Host "  $contador) $orfao" -ForegroundColor Cyan
            $mapaDrivers[$contador.ToString()] = @{
                Nome = $orfao
                Modelo = $null
                Cadastrado = $false
                ModeloObj = $null
            }
            $contador++
        }
        
        Write-Host ""
    }
    
    if ($mapaDrivers.Count -eq 0) {
        Write-Mensagem "Nenhum driver Samsung encontrado no sistema" "Aviso"
        Read-Host "`nPressione ENTER para continuar"
        return
    }
    
    Write-Host "  V) Voltar" -ForegroundColor Gray
    Write-Host ""
    
    $escolha = Read-Host "Escolha o driver"
    
    if ($escolha -eq "V" -or $escolha -eq "v") {
        return
    }
    
    if (-not $mapaDrivers.ContainsKey($escolha)) {
        Write-Mensagem "`nOpcao invalida!" "Erro"
        Start-Sleep -Seconds 2
        return
    }
    
    $driverSelecionado = $mapaDrivers[$escolha]
    $impressorasUsando = Get-ImpressorasUsandoDriver -nomeDriver $driverSelecionado.Nome
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   DRIVER SELECIONADO" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    Write-Host "Driver: $($driverSelecionado.Nome)" -ForegroundColor White
    
    if ($driverSelecionado.Cadastrado) {
        Write-Host "Modelo: $($driverSelecionado.Modelo)`n" -ForegroundColor White
    } else {
        Write-Host "Status: Driver nao cadastrado no CSV`n" -ForegroundColor Yellow
    }
    
    if ($impressorasUsando.Count -gt 0) {
        Write-Host "Impressoras usando este driver:" -ForegroundColor White
        foreach ($imp in $impressorasUsando) {
            Write-Host "  - $($imp.Name) ($($imp.PortName))" -ForegroundColor Gray
        }
    } else {
        Write-Host "Nenhuma impressora usando este driver." -ForegroundColor Yellow
        Write-Host ""
        
        $removerSemImpressora = Read-OpcaoValidada "Deseja remover o driver mesmo assim? [S/N]" @("S","s","N","n")
        
        if ($removerSemImpressora -eq "N" -or $removerSemImpressora -eq "n") {
            return
        }
    }
    
    $sucesso = Remove-ImpressorasEDriverEspecifico -nomeDriver $driverSelecionado.Nome
    
    if ($sucesso -and $driverSelecionado.Cadastrado) {
        $reinstalar = Read-OpcaoValidada "`nDeseja reinstalar agora? [S/N]" @("S","s","N","n")
        
        if ($reinstalar -eq "S" -or $reinstalar -eq "s") {
            $scriptInstalador = Join-Path (Split-Path $csvPath) "instalar_universal.ps1"
            
            if (Test-Path $scriptInstalador) {
                Write-Host "`nIniciando reinstalacao...`n" -ForegroundColor Cyan
                Start-Sleep -Seconds 2
                
                $params = @{
                    modelo = $driverSelecionado.ModeloObj.Modelo
                    urlPrint = $driverSelecionado.ModeloObj.UrlPrint
                    temScan = $driverSelecionado.ModeloObj.TemScan
                    urlScan = $driverSelecionado.ModeloObj.UrlScan
                    filtroDriverWindows = $driverSelecionado.ModeloObj.FiltroDriver
                    instalarPrint = $true
                    instalarScan = $false
                    instalarEPM = $false
                    instalarEDC = $false
                }
                
                & $scriptInstalador @params
            }
            else {
                Write-Mensagem "Script de instalacao nao encontrado" "Erro"
            }
        }
    }
    
    Read-Host "`nPressione ENTER para continuar"
}

# ================================================================================
# INICIO DO SCRIPT
# ================================================================================

$dadosCSV = Get-PrinterData

if (-not $dadosCSV) {
    Read-Host "Pressione ENTER para sair"
    exit
}

do {
    Show-MenuReparo
    $opcao = Read-Host "Escolha uma opcao"
    
    switch ($opcao) {
        "1" {
            Invoke-RemoverDriverDeImpressora -modelosCSV $dadosCSV
        }
        "2" {
            Invoke-RemoverDriverEspecifico -modelosCSV $dadosCSV
        }
        { $_ -eq "V" -or $_ -eq "v" } {
            Write-Host "`nVoltando ao menu principal...`n" -ForegroundColor Gray
            Start-Sleep -Seconds 1
            break
        }
        default {
            Write-Mensagem "`nOpcao invalida! Tente novamente." "Aviso"
            Start-Sleep -Seconds 2
        }
    }
    
    if ($opcao -eq "V" -or $opcao -eq "v") {
        break
    }
    
} while ($true)

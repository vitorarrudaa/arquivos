# ================================================================================
# SCRIPT: Reparo de Impressoras Samsung
# VERSAO: 1.0
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
    
    $impressoras = Get-Printer -ErrorAction SilentlyContinue | 
                   Where-Object { $_.DriverName -eq $nomeDriver }
    
    return $impressoras
}

function Get-DriversOrfaos {
    param([array]$modelosCSV)
    
    # Lista todos drivers Samsung instalados
    $todosDriversSamsung = Get-PrinterDriver -ErrorAction SilentlyContinue | 
                          Where-Object { $_.Name -like "*Samsung*" }
    
    # Extrai drivers do CSV
    $driversCSV = @()
    foreach ($modelo in $modelosCSV) {
        $driversCSV += $modelo.FiltroDriver
    }
    
    # Identifica órfãos
    $orfaos = @()
    foreach ($driver in $todosDriversSamsung) {
        $ehOrfao = $true
        foreach ($driverCSV in $driversCSV) {
            if ($driver.Name -eq $driverCSV -or $driver.Name -like "*$driverCSV*") {
                $ehOrfao = $false
                break
            }
        }
        
        if ($ehOrfao) {
            $impressorasUsando = Get-ImpressorasUsandoDriver -nomeDriver $driver.Name
            $orfaos += [PSCustomObject]@{
                Nome = $driver.Name
                Impressoras = $impressorasUsando
                Quantidade = $impressorasUsando.Count
            }
        }
    }
    
    return $orfaos
}

function Remove-DriverCompleto {
    param([string]$nomeDriver)
    
    Write-Host "  Removendo: $nomeDriver" -ForegroundColor Gray
    
    # Passo 1: Remove do sistema de impressão
    try {
        Remove-PrinterDriver -Name $nomeDriver -ErrorAction Stop
        Write-Host "    ✓ Removido do sistema de impressao" -ForegroundColor Green
    }
    catch {
        Write-Mensagem "    Falha ao remover do sistema: $($_.Exception.Message)" "Aviso"
    }
    
    # Passo 2: Remoção profunda via pnputil
    try {
        $drivers = pnputil /enum-drivers | Out-String
        
        # Procura o .inf relacionado ao driver
        $linhas = $drivers -split "`n"
        $infEncontrado = $null
        
        for ($i = 0; $i -lt $linhas.Count; $i++) {
            if ($linhas[$i] -match "Published Name\s*:\s*(oem\d+\.inf)") {
                $infAtual = $matches[1]
                
                # Verifica se nas próximas linhas aparece o nome do driver
                for ($j = $i; $j -lt [Math]::Min($i + 10, $linhas.Count); $j++) {
                    if ($linhas[$j] -match [regex]::Escape($nomeDriver)) {
                        $infEncontrado = $infAtual
                        break
                    }
                }
            }
            
            if ($infEncontrado) { break }
        }
        
        if ($infEncontrado) {
            $resultado = pnputil /delete-driver $infEncontrado /uninstall /force 2>&1
            Write-Host "    ✓ Removido do DriverStore ($infEncontrado)" -ForegroundColor Green
        }
        else {
            Write-Host "    ⚠ Arquivo .inf nao encontrado (remocao parcial)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "    ⚠ Erro na remocao profunda: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Stop-SpoolerCompleto {
    Write-Host "`nParando servico de impressao..." -ForegroundColor Gray
    
    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        # Limpa cache do spooler
        Remove-Item "C:\Windows\System32\spool\PRINTERS\*" -Force -ErrorAction SilentlyContinue
        
        Write-Mensagem "Spooler parado e cache limpo" "Sucesso"
        return $true
    }
    catch {
        Write-Mensagem "Falha ao parar spooler: $($_.Exception.Message)" "Erro"
        return $false
    }
}

function Start-SpoolerCompleto {
    Write-Host "`nReiniciando servico de impressao..." -ForegroundColor Gray
    
    try {
        Start-Service -Name Spooler -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Mensagem "Spooler reiniciado com sucesso" "Sucesso"
        return $true
    }
    catch {
        Write-Mensagem "Falha ao reiniciar spooler: $($_.Exception.Message)" "Erro"
        return $false
    }
}

function Remove-ImpressorasEDrivers {
    param(
        [array]$impressoras,
        [array]$drivers
    )
    
    $totalImpressoras = $impressoras.Count
    $totalDrivers = $drivers.Count
    
    Write-Host "`n========================================" -ForegroundColor Yellow
    Write-Host "⚠️  ATENCAO! ESSA OPERACAO VAI REMOVER:" -ForegroundColor Yellow
    Write-Host "   - $totalImpressoras IMPRESSORA(S)" -ForegroundColor Yellow
    Write-Host "   - $totalDrivers DRIVER(S)" -ForegroundColor Yellow
    Write-Host "========================================`n" -ForegroundColor Yellow
    
    $confirmar = Read-OpcaoValidada "DESEJA PROSSEGUIR? [S/N]" @("S","s","N","n")
    
    if ($confirmar -eq "N" -or $confirmar -eq "n") {
        Write-Mensagem "`nOperacao cancelada pelo usuario" "Info"
        return $false
    }
    
    # Para o spooler
    if (-not (Stop-SpoolerCompleto)) {
        return $false
    }
    
    Write-Host "`n--- REMOVENDO IMPRESSORAS ---" -ForegroundColor Cyan
    
    foreach ($impressora in $impressoras) {
        try {
            Remove-Printer -Name $impressora.Name -Confirm:$false -ErrorAction Stop
            Write-Host "  ✓ $($impressora.Name)" -ForegroundColor Green
        }
        catch {
            Write-Host "  ✗ Falha ao remover: $($impressora.Name)" -ForegroundColor Red
        }
    }
    
    Start-Sleep -Seconds 2
    
    Write-Host "`n--- REMOVENDO DRIVERS ---" -ForegroundColor Cyan
    
    foreach ($driver in $drivers) {
        Remove-DriverCompleto -nomeDriver $driver
    }
    
    # Reinicia o spooler
    Start-SpoolerCompleto | Out-Null
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Mensagem "REMOCAO CONCLUIDA COM SUCESSO!" "Sucesso"
    Write-Host "========================================`n" -ForegroundColor Green
    
    return $true
}

function Show-MenuReparo {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "   TECH3 - REPARO DE IMPRESSORAS" -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "  1) Remover drivers por modelo (CSV)" -ForegroundColor White
    Write-Host "  2) Remover impressora especifica" -ForegroundColor White
    Write-Host ""
    Write-Host "  V) Voltar ao menu principal" -ForegroundColor Gray
    Write-Host ""
}

function Invoke-RemoverPorModelo {
    param([array]$modelosCSV)
    
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   REMOVER DRIVERS POR MODELO" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Lista modelos
    foreach ($modelo in $modelosCSV) {
        $temScan = if ($modelo.TemScan -eq "S") { "(Print + Scan)" } else { "(Apenas Print)" }
        Write-Host "  $($modelo.ID)) $($modelo.Modelo) $temScan" -ForegroundColor White
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
    
    # Identifica drivers do modelo
    $driversModelo = @($modeloSelecionado.FiltroDriver)
    
    # Busca impressoras usando esses drivers
    $impressorasEncontradas = @()
    
    foreach ($driver in $driversModelo) {
        $imps = Get-ImpressorasUsandoDriver -nomeDriver $driver
        $impressorasEncontradas += $imps
    }
    
    # Busca drivers órfãos
    $orfaos = Get-DriversOrfaos -modelosCSV $modelosCSV
    
    # Exibe resumo
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "  DRIVERS ENCONTRADOS NO SISTEMA" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    Write-Host "📋 DRIVERS CADASTRADOS (CSV):" -ForegroundColor Green
    foreach ($driver in $driversModelo) {
        $qtd = ($impressorasEncontradas | Where-Object { $_.DriverName -eq $driver }).Count
        Write-Host "  ✓ $driver" -ForegroundColor White
        Write-Host "    - Impressoras: $qtd" -ForegroundColor Gray
        
        if ($qtd -gt 0) {
            $imps = $impressorasEncontradas | Where-Object { $_.DriverName -eq $driver }
            foreach ($imp in $imps) {
                Write-Host "      • $($imp.Name) ($($imp.PortName))" -ForegroundColor Gray
            }
        }
    }
    
    if ($orfaos.Count -gt 0) {
        Write-Host "`n⚠️  DRIVERS ORFAOS (nao estao no CSV):" -ForegroundColor Yellow
        foreach ($orfao in $orfaos) {
            Write-Host "  • $($orfao.Nome)" -ForegroundColor White
            Write-Host "    - Impressoras: $($orfao.Quantidade)" -ForegroundColor Gray
            
            if ($orfao.Quantidade -gt 0) {
                foreach ($imp in $orfao.Impressoras) {
                    Write-Host "      • $($imp.Name) ($($imp.PortName))" -ForegroundColor Gray
                }
            }
        }
        
        Write-Host ""
        $removerOrfaos = Read-OpcaoValidada "Deseja remover os drivers ORFAOS tambem? [S/N]" @("S","s","N","n")
        
        if ($removerOrfaos -eq "S" -or $removerOrfaos -eq "s") {
            foreach ($orfao in $orfaos) {
                $driversModelo += $orfao.Nome
                $impressorasEncontradas += $orfao.Impressoras
            }
        }
    }
    
    if ($impressorasEncontradas.Count -eq 0 -and $driversModelo.Count -eq 0) {
        Write-Mensagem "`nNenhum driver ou impressora encontrada para este modelo" "Aviso"
        Start-Sleep -Seconds 2
        return
    }
    
    # Remove
    $sucesso = Remove-ImpressorasEDrivers -impressoras $impressorasEncontradas -drivers $driversModelo
    
    if ($sucesso) {
        $reinstalar = Read-OpcaoValidada "`nDeseja reinstalar agora? [S/N]" @("S","s","N","n")
        
        if ($reinstalar -eq "S" -or $reinstalar -eq "s") {
            # Chama o instalador universal
            $scriptInstalador = Join-Path (Split-Path $csvPath) "instalar_universal.ps1"
            
            if (Test-Path $scriptInstalador) {
                Write-Host "`nIniciando processo de reinstalacao...`n" -ForegroundColor Cyan
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
                Write-Mensagem "Script de instalacao nao encontrado: $scriptInstalador" "Erro"
            }
        }
    }
    
    Read-Host "`nPressione ENTER para continuar"
}

function Invoke-RemoverEspecifica {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "   REMOVER IMPRESSORA ESPECIFICA" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # Lista todas impressoras
    $todasImpressoras = Get-Printer -ErrorAction SilentlyContinue
    
    if ($todasImpressoras.Count -eq 0) {
        Write-Mensagem "Nenhuma impressora instalada no sistema" "Aviso"
        Read-Host "`nPressione ENTER para continuar"
        return
    }
    
    $indice = 1
    $mapa = @{}
    
    foreach ($imp in $todasImpressoras) {
        Write-Host "  $indice) $($imp.Name) - Driver: $($imp.DriverName)" -ForegroundColor White
        Write-Host "     IP/Porta: $($imp.PortName)" -ForegroundColor Gray
        $mapa[$indice.ToString()] = $imp
        $indice++
    }
    
    Write-Host ""
    Write-Host "  V) Voltar" -ForegroundColor Gray
    Write-Host ""
    
    $escolha = Read-Host "Escolha a impressora"
    
    if ($escolha -eq "V" -or $escolha -eq "v") {
        return
    }
    
    if (-not $mapa.ContainsKey($escolha)) {
        Write-Mensagem "`nOpcao invalida!" "Erro"
        Start-Sleep -Seconds 2
        return
    }
    
    $impressoraSelecionada = $mapa[$escolha]
    $driverUsado = $impressoraSelecionada.DriverName
    
    # Verifica se outras impressoras usam o mesmo driver
    $outrasImpressoras = $todasImpressoras | 
                         Where-Object { $_.DriverName -eq $driverUsado -and $_.Name -ne $impressoraSelecionada.Name }
    
    $impressorasRemover = @($impressoraSelecionada)
    $impressorasRemover += $outrasImpressoras
    
    if ($outrasImpressoras.Count -gt 0) {
        Write-Host "`n========================================" -ForegroundColor Yellow
        Write-Host "⚠️  ATENCAO!" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "Outras impressoras usam o mesmo driver:" -ForegroundColor Yellow
        
        foreach ($outra in $outrasImpressoras) {
            Write-Host "  • $($outra.Name) ($($outra.PortName))" -ForegroundColor White
        }
        
        Write-Host "`nREMOVER ESSA IMPRESSORA VAI APAGAR O DRIVER" -ForegroundColor Yellow
        Write-Host "USADO POR $($outrasImpressoras.Count) IMPRESSORA(S) ADICIONAL(IS)." -ForegroundColor Yellow
        Write-Host "TODAS ELAS SERAO REMOVIDAS!`n" -ForegroundColor Yellow
    }
    
    # Remove impressoras e driver
    $sucesso = Remove-ImpressorasEDrivers -impressoras $impressorasRemover -drivers @($driverUsado)
    
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
            Invoke-RemoverPorModelo -modelosCSV $dadosCSV
        }
        "2" {
            Invoke-RemoverEspecifica
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
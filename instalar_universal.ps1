# ================================================================================
# SCRIPT: Motor de Instalacao Universal - Impressoras Samsung
# VERSAO: 3.0 (Refatorado)
# DESCRICAO: Instalador modular para drivers Samsung SPL e UPD
# ================================================================================

param (
    [Parameter(Mandatory=$true)][string]$modelo,
    [Parameter(Mandatory=$true)][string]$urlPrint,
    [Parameter(Mandatory=$true)][string]$temScan,
    [Parameter(Mandatory=$false)][string]$urlScan = "",
    [Parameter(Mandatory=$true)][string]$filtroDriverWindows,
    [Parameter(Mandatory=$true)][bool]$instalarPrint,
    [Parameter(Mandatory=$true)][bool]$instalarScan,
    [Parameter(Mandatory=$true)][bool]$instalarEPM,
    [Parameter(Mandatory=$true)][bool]$instalarEDC
)

# --- CONFIGURACAO GLOBAL ---
$Global:Config = @{
    UrlEPM      = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
    UrlEDC      = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
    CaminhoTemp = "$env:USERPROFILE\Downloads\Instalacao_Samsung"
    TempoEspera = 10
}

$Global:TipoDriver = if ($modelo -match "M4080|CLX-6260") { "UPD" } else { "SPL" }

if (-not (Test-Path $Global:Config.CaminhoTemp)) {
    New-Item $Global:Config.CaminhoTemp -ItemType Directory -Force | Out-Null
}

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

function Test-ProgramaInstalado {
    param([string]$nomePrograma)
    
    $chaves = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $programas = Get-ItemProperty $chaves -ErrorAction SilentlyContinue | 
                 Where-Object { $_.DisplayName -like "*$nomePrograma*" }
    
    return [bool]$programas
}

function Get-ArquivoLocal {
    param([string]$url, [string]$nomeDestino)
    
    $caminhoCompleto = Join-Path $Global:Config.CaminhoTemp $nomeDestino
    
    if (Test-Path $caminhoCompleto) {
        Write-Mensagem "Arquivo ja existe localmente. Reutilizando..." "Info"
        return $caminhoCompleto
    }
    
    try {
        Write-Mensagem "Baixando: $nomeDestino..." "Info"
        Invoke-WebRequest -Uri $url -OutFile $caminhoCompleto -ErrorAction Stop -UseBasicParsing
        Write-Mensagem "Download concluido!" "Sucesso"
        return $caminhoCompleto
    }
    catch {
        Write-Mensagem "Falha ao baixar arquivo: $($_.Exception.Message)" "Erro"
        Read-Host "Pressione ENTER para continuar"
        return $null
    }
}

function New-PortaIP {
    param([Parameter(Mandatory=$true)][string]$enderecoIP)
    
    if (-not (Get-PrinterPort $enderecoIP -ErrorAction SilentlyContinue)) {
        try {
            Add-PrinterPort -Name $enderecoIP -PrinterHostAddress $enderecoIP -ErrorAction Stop
            Write-Mensagem "Porta IP criada: $enderecoIP" "Info"
            return $true
        } catch {
            Write-Mensagem "Falha ao criar porta IP: $($_.Exception.Message)" "Erro"
            return $false
        }
    } else {
        Write-Mensagem "Porta IP $enderecoIP ja existe" "Info"
        return $true
    }
}

function Test-RedeImpressora {
    param([Parameter(Mandatory=$true)][string]$enderecoIP)
    
    Write-Mensagem "Verificando conectividade com a impressora..." "Info"
    
    # Obter IP do computador
    $meuIP = (Get-NetIPAddress -AddressFamily IPv4 | 
              Where-Object { $_.IPAddress -notlike "127.*" -and $_.PrefixOrigin -ne "WellKnown" } | 
              Select-Object -First 1).IPAddress
    
    if (-not $meuIP) {
        Write-Mensagem "Nao foi possivel detectar IP do computador" "Aviso"
        return $true  # Continua mesmo assim
    }
    
    # Comparar subnet (primeiros 3 octetos)
    $minhaRede = ($meuIP -split '\.')[0..2] -join '.'
    $redeImpressora = ($enderecoIP -split '\.')[0..2] -join '.'
    
    if ($minhaRede -ne $redeImpressora) {
        Write-Mensagem "IP em rede diferente detectado!" "Aviso"
        Write-Host "  Seu computador: $meuIP"
        Write-Host "  IP digitado: $enderecoIP"
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim?" @("S","s","N","n")
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    # Teste de ping (mesma rede)
    Write-Mensagem "Testando ping..." "Info"
    $pingOk = Test-Connection -ComputerName $enderecoIP -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if (-not $pingOk) {
        Write-Mensagem "Impressora nao responde ao ping!" "Aviso"
        Write-Host "  Verifique se esta ligada e conectada a rede"
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim?" @("S","s","N","n")
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    Write-Mensagem "Impressora responde ao ping!" "Sucesso"
    return $true
}

function Test-DriverExistente {
    param([Parameter(Mandatory=$true)][string]$filtroDriver)
    
    Write-Mensagem "Verificando drivers instalados..." "Info"
    
    # Buscar driver nativo (sem PCL/PS)
    $driverNativo = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*$filtroDriver*" -and
        $_.Name -notlike "*PCL*" -and
        $_.Name -notlike "* PS"
    } | Select-Object -First 1
    
    if ($driverNativo) {
        Write-Mensagem "Driver nativo '$($driverNativo.Name)' ja instalado!" "Sucesso"
        return @{ Encontrado = $true; Driver = $driverNativo.Name; Tipo = "Nativo" }
    }
    
    # Verificar se existe PCL/PS
    $driverVariacao = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*$filtroDriver*"
    } | Select-Object -First 1
    
    if ($driverVariacao) {
        Write-Mensagem "Encontrado '$($driverVariacao.Name)' (nao e nativo)" "Aviso"
        return @{ Encontrado = $true; Driver = $driverVariacao.Name; Tipo = "Variacao" }
    }
    
    Write-Mensagem "Driver nao encontrado no sistema" "Info"
    return @{ Encontrado = $false; Driver = $null; Tipo = $null }
}

function Install-DriverSPL {
    param(
        [string]$urlDriver,
        [string]$nomeModelo,
        [string]$filtroDriver,
        [string]$nomeImpressora,
        [string]$enderecoIP
    )
    
    # Verificar se driver nativo ja existe
    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    
    if ($statusDriver.Encontrado -and $statusDriver.Tipo -eq "Nativo") {
        # Driver nativo existe - usar direto
        Write-Mensagem "Pulando instalacao (usando driver existente)" "Info"
        
        New-PortaIP -enderecoIP $enderecoIP | Out-Null
        
        try {
            Add-Printer -Name $nomeImpressora -DriverName $statusDriver.Driver -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Impressora configurada com driver existente!" "Sucesso"
            return $true
        } catch {
            Write-Mensagem "Falha ao criar impressora: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    
    # Precisa instalar driver (nao existe ou so tem variacao PCL/PS)
    if ($statusDriver.Tipo -eq "Variacao") {
        Write-Mensagem "Instalando versao nativa preferencial..." "Info"
    } else {
        Write-Mensagem "Instalando driver..." "Info"
    }
    
    $nomeArquivo = "driver_print_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    
    if (-not $arquivoDriver) { return $false }
    
    Write-Mensagem "Executando instalador (aguarde)..." "Info"
    Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
    Start-Sleep -Seconds $Global:Config.TempoEspera
    
    New-PortaIP -enderecoIP $enderecoIP | Out-Null
    
    # Detectar e configurar fila criada pelo instalador
    Write-Mensagem "Configurando fila de impressao..." "Info"
    
    # Buscar fila nativa primeiro
    $filaEspecifica = Get-Printer -ErrorAction SilentlyContinue | 
                     Where-Object {
                         ($_.Name -like "*$filtroDriver*" -or $_.DriverName -like "*$filtroDriver*") -and
                         $_.DriverName -notlike "*PCL*" -and 
                         $_.DriverName -notlike "* PS"
                     } | Select-Object -First 1
    
    if (-not $filaEspecifica) {
        $filaEspecifica = Get-Printer -ErrorAction SilentlyContinue | 
                         Where-Object {
                             $_.Name -like "*$filtroDriver*" -or 
                             $_.DriverName -like "*$filtroDriver*"
                         } | Select-Object -First 1
    }
    
    try {
        if ($filaEspecifica) {
            # Trocar para nativo se for PCL/PS
            if ($filaEspecifica.DriverName -like "*PCL*" -or $filaEspecifica.DriverName -like "* PS") {
                $driverNativo = Get-PrinterDriver -ErrorAction SilentlyContinue | 
                               Where-Object { 
                                   $_.Name -like "*$filtroDriver*" -and 
                                   $_.Name -notlike "*PCL*" -and 
                                   $_.Name -notlike "* PS"
                               } | Select-Object -First 1
                
                if ($driverNativo) {
                    Set-Printer -Name $filaEspecifica.Name -DriverName $driverNativo.Name -PortName $enderecoIP -ErrorAction Stop
                } else {
                    Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop
                    Write-Mensagem "Driver nativo nao disponivel (mantendo $($filaEspecifica.DriverName))" "Aviso"
                }
            } else {
                Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop
            }
            
            Rename-Printer -Name $filaEspecifica.Name -NewName $nomeImpressora -ErrorAction Stop
            Write-Mensagem "Impressora configurada com sucesso!" "Sucesso"
            return $true
        } else {
            # Criar manualmente
            Add-Printer -Name $nomeImpressora -DriverName $filtroDriver -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Fila criada manualmente!" "Sucesso"
            return $true
        }
    }
    catch {
        Write-Mensagem "Falha ao configurar: $($_.Exception.Message)" "Erro"
        
        Write-Host "`nDrivers Samsung disponiveis:"
        Get-PrinterDriver -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -like "*Samsung*" } | 
            ForEach-Object { Write-Host "  - $($_.Name)" }
        
        Read-Host "`nPressione ENTER para continuar"
        return $false
    }
}

function Install-DriverUPD {
    param(
        [string]$urlDriver,
        [string]$nomeModelo,
        [string]$filtroDriver,
        [string]$nomeImpressora,
        [string]$enderecoIP
    )
    
    Write-Mensagem "Instalando driver UPD (Universal Print Driver)..." "Info"
    
    # TODO: Implementar logica especifica para UPD
    # Por enquanto usa mesma logica que SPL
    # Precisa testar comportamento real da M4080/CLX-6260
    
    return Install-DriverSPL -urlDriver $urlDriver `
                            -nomeModelo $nomeModelo `
                            -filtroDriver $filtroDriver `
                            -nomeImpressora $nomeImpressora `
                            -enderecoIP $enderecoIP
}

function Remove-FilaDuplicada {
    param([string]$nomeConfigurado, [string]$filtroDriver)
    
    Write-Mensagem "Verificando filas duplicadas..." "Info"
    
    $todasImpressoras = Get-Printer -ErrorAction SilentlyContinue
    
    $duplicatas = $todasImpressoras | Where-Object {
        $_.Name -ne $nomeConfigurado -and
        (
            $_.Name -eq $filtroDriver -or
            $_.Name -match "^$([regex]::Escape($filtroDriver))( PS| PCL[0-9].*)?( \((C[oó]pia|Copiar|Copy) \d+\))?$" -or
            ($_.Name -like "*Samsung Universal Print Driver*" -and $_.DriverName -like "*Samsung Universal*")
        ) -and
        $_.Name -notlike "*Fax*"
    }
    
    if ($duplicatas) {
        foreach ($fila in $duplicatas) {
            try {
                Remove-Printer -Name $fila.Name -Confirm:$false -ErrorAction Stop
                Write-Mensagem "Removida: $($fila.Name)" "Info"
            }
            catch {
                Write-Mensagem "Falha ao remover: $($fila.Name)" "Erro"
            }
        }
    } else {
        Write-Mensagem "Nenhuma duplicata encontrada" "Sucesso"
    }
}

function Show-ResumoInstalacao {
    param(
        [string]$modelo,
        [string]$nomeImpressora,
        [string]$enderecoIP,
        [string]$driver,
        [string[]]$componentes
    )
    
    Write-Host "`n========================================"
    Write-Host "       RESUMO DA INSTALACAO"
    Write-Host "========================================"
    Write-Host "Modelo:      $modelo"
    Write-Host "Nome:        $nomeImpressora"
    Write-Host "IP:          $enderecoIP"
    Write-Host "Driver:      $driver"
    Write-Host "Componentes: $($componentes -join ', ')"
    Write-Host "Status:      " -NoNewline
    Write-Host "Concluido com sucesso!" -ForegroundColor Green
    Write-Host "========================================`n"
}

# ================================================================================
# INICIO DO PROCESSAMENTO
# ================================================================================

$componentesInstalados = @()
if ($instalarPrint) { $componentesInstalados += "Print" }
if ($instalarScan) { $componentesInstalados += "Scan" }
if ($instalarEPM) { $componentesInstalados += "EPM" }
if ($instalarEDC) { $componentesInstalados += "EDC" }

$totalEtapas = $componentesInstalados.Count
$etapaAtual = 1

Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host "  INSTALACAO: $modelo" -ForegroundColor Cyan
Write-Host "  Tipo: $($Global:TipoDriver) | Etapas: $totalEtapas" -ForegroundColor Cyan
Write-Host "================================================`n" -ForegroundColor Cyan

# ================================================================================
# ETAPA 1: DRIVER DE IMPRESSAO
# ================================================================================

$driverInstalado = $null

if ($instalarPrint) {
    Write-Host "[$etapaAtual/$totalEtapas] DRIVER DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host ""
    
    # Coletar nome
    do {
        $nomeImpressora = Read-Host "Nome da impressora"
        $impressoraExistente = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
        
        if ($impressoraExistente) {
            Write-Mensagem "Ja existe impressora com nome '$nomeImpressora'" "Aviso"
            Write-Host "  IP atual: $($impressoraExistente.PortName)"
            Write-Host "  Driver:   $($impressoraExistente.DriverName)"
            
            $opcao = Read-OpcaoValidada "`n[1] Digitar outro nome [2] Cancelar" @("1","2")
            if ($opcao -eq "2") {
                Write-Mensagem "Instalacao cancelada" "Info"
                return
            }
        }
    } while ($impressoraExistente)
    
    # Coletar e validar IP
    do {
        $enderecoIP = Read-Host "Endereco IP"
        
        if ($enderecoIP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
            Write-Mensagem "IP invalido! Use formato XXX.XXX.XXX.XXX" "Erro"
            continue
        }
        
        if (-not (Test-RedeImpressora -enderecoIP $enderecoIP)) {
            Write-Mensagem "Instalacao cancelada pelo usuario" "Info"
            return
        }
        
        break
    } while ($true)
    
    # Instalar driver baseado no tipo
    $sucesso = if ($Global:TipoDriver -eq "UPD") {
        Install-DriverUPD -urlDriver $urlPrint `
                         -nomeModelo $modelo `
                         -filtroDriver $filtroDriverWindows `
                         -nomeImpressora $nomeImpressora `
                         -enderecoIP $enderecoIP
    } else {
        Install-DriverSPL -urlDriver $urlPrint `
                         -nomeModelo $modelo `
                         -filtroDriver $filtroDriverWindows `
                         -nomeImpressora $nomeImpressora `
                         -enderecoIP $enderecoIP
    }
    
    if ($sucesso) {
        $impressoraFinal = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
        $driverInstalado = $impressoraFinal.DriverName
        
        Start-Sleep -Seconds 2
        Remove-FilaDuplicada -nomeConfigurado $nomeImpressora -filtroDriver $filtroDriverWindows
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# ETAPA 2: DRIVER DE DIGITALIZACAO
# ================================================================================

if ($instalarScan) {
    Write-Host "[$etapaAtual/$totalEtapas] DRIVER DE DIGITALIZACAO" -ForegroundColor Yellow
    Write-Host ""
    
    if ([string]::IsNullOrWhiteSpace($urlScan)) {
        Write-Mensagem "URL de scan nao disponivel" "Aviso"
    } else {
        $nomeArquivoScan = "driver_scan_" + ($modelo -replace '\s+', '_') + ".exe"
        $arquivoScan = Get-ArquivoLocal -url $urlScan -nomeDestino $nomeArquivoScan
        
        if ($arquivoScan) {
            Write-Mensagem "Instalando driver de scan..." "Info"
            Start-Process $arquivoScan -ArgumentList "/S" -Wait -NoNewWindow
            Write-Mensagem "Driver de scan instalado!" "Sucesso"
        }
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# ETAPA 3: EASY PRINTER MANAGER
# ================================================================================

if ($instalarEPM) {
    Write-Host "[$etapaAtual/$totalEtapas] EASY PRINTER MANAGER" -ForegroundColor Yellow
    Write-Host ""
    
    if (Test-ProgramaInstalado "Easy Printer Manager") {
        Write-Mensagem "Ja instalado no sistema" "Info"
    } else {
        $arquivoEPM = Get-ArquivoLocal -url $Global:Config.UrlEPM -nomeDestino "EPM_Universal.exe"
        
        if ($arquivoEPM) {
            Write-Mensagem "Instalando (timeout: 60s)..." "Info"
            
            $processo = Start-Process $arquivoEPM -ArgumentList "/S" -PassThru -NoNewWindow
            $tempoLimite = 60
            $tempoDecorrido = 0
            
            while (-not $processo.HasExited -and $tempoDecorrido -lt $tempoLimite) {
                Start-Sleep -Seconds 2
                $tempoDecorrido += 2
                
                if (Test-ProgramaInstalado "Easy Printer Manager") {
                    Write-Mensagem "Instalado com sucesso!" "Sucesso"
                    if (-not $processo.HasExited) {
                        Stop-Process -Id $processo.Id -Force -ErrorAction SilentlyContinue
                    }
                    break
                }
            }
            
            if (-not (Test-ProgramaInstalado "Easy Printer Manager")) {
                Write-Mensagem "Instalacao pode nao ter sido concluida" "Aviso"
            }
        }
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# ETAPA 4: EASY DOCUMENT CREATOR
# ================================================================================

if ($instalarEDC) {
    Write-Host "[$etapaAtual/$totalEtapas] EASY DOCUMENT CREATOR" -ForegroundColor Yellow
    Write-Host ""
    
    if (Test-ProgramaInstalado "Easy Document Creator") {
        Write-Mensagem "Ja instalado no sistema" "Info"
    } else {
        $arquivoEDC = Get-ArquivoLocal -url $Global:Config.UrlEDC -nomeDestino "EDC_Universal.exe"
        
        if ($arquivoEDC) {
            Write-Mensagem "Instalando (timeout: 60s)..." "Info"
            
            $processo = Start-Process $arquivoEDC -ArgumentList "/S" -PassThru -NoNewWindow
            $tempoLimite = 60
            $tempoDecorrido = 0
            
            while (-not $processo.HasExited -and $tempoDecorrido -lt $tempoLimite) {
                Start-Sleep -Seconds 2
                $tempoDecorrido += 2
                
                if (Test-ProgramaInstalado "Easy Document Creator") {
                    Write-Mensagem "Instalado com sucesso!" "Sucesso"
                    if (-not $processo.HasExited) {
                        Stop-Process -Id $processo.Id -Force -ErrorAction SilentlyContinue
                    }
                    break
                }
            }
            
            if (-not (Test-ProgramaInstalado "Easy Document Creator")) {
                Write-Mensagem "Instalacao pode nao ter sido concluida" "Aviso"
            }
        }
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# FINALIZACAO
# ================================================================================

if ($instalarPrint) {
    Show-ResumoInstalacao -modelo $modelo `
                         -nomeImpressora $nomeImpressora `
                         -enderecoIP $enderecoIP `
                         -driver $driverInstalado `
                         -componentes $componentesInstalados
    
    $imprimirTeste = Read-OpcaoValidada "Deseja imprimir uma pagina de teste?" @("S","s","N","n")
    
    if ($imprimirTeste -eq "S" -or $imprimirTeste -eq "s") {
        try {
            Start-Process -FilePath "rundll32.exe" `
                         -ArgumentList "printui.dll,PrintUIEntry /k /n `"$nomeImpressora`"" `
                         -NoNewWindow -Wait
            Write-Mensagem "Pagina de teste enviada!" "Sucesso"
        } catch {
            Write-Mensagem "Falha ao enviar pagina de teste" "Erro"
        }
    }
}

Write-Host ""
Start-Sleep -Seconds 2

# ================================================================================
# SCRIPT: Motor de Instalacao Universal - Impressoras Samsung
# VERSAO: 2.1 (Atualizado)
# DESCRICAO: Instalador universal para drivers Samsung
# ================================================================================

param (
    [Parameter(Mandatory=$true)]
    [string]$modelo,
    
    [Parameter(Mandatory=$true)]
    [string]$urlPrint,
    
    [Parameter(Mandatory=$true)]
    [string]$temScan,
    
    [Parameter(Mandatory=$false)]
    [string]$urlScan = "",
    
    [Parameter(Mandatory=$true)]
    [string]$filtroDriverWindows,
    
    [Parameter(Mandatory=$true)]
    [bool]$instalarPrint,
    
    [Parameter(Mandatory=$true)]
    [bool]$instalarScan,
    
    [Parameter(Mandatory=$true)]
    [bool]$instalarEPM,
    
    [Parameter(Mandatory=$true)]
    [bool]$instalarEDC
)

# --- CONFIGURACAO GLOBAL ---
$Global:Config = @{
    UrlEPM        = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
    UrlEDC        = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
    CaminhoTemp   = "$env:USERPROFILE\Downloads\Instalacao_Samsung"
    TempoEspera   = 10
}

# Criar pasta temporaria
if (-not (Test-Path $Global:Config.CaminhoTemp)) {
    New-Item $Global:Config.CaminhoTemp -ItemType Directory -Force | Out-Null
}

# --- FUNCOES AUXILIARES ---

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
    param(
        [string]$url,
        [string]$nomeDestino
    )
    
    $caminhoCompleto = Join-Path $Global:Config.CaminhoTemp $nomeDestino
    
    if (Test-Path $caminhoCompleto) {
        Write-Host "  -> Arquivo ja existe localmente. Reutilizando..." -ForegroundColor Cyan
        return $caminhoCompleto
    }
    
    try {
        Write-Host "  -> Baixando: $nomeDestino..." -ForegroundColor Gray
        Invoke-WebRequest -Uri $url -OutFile $caminhoCompleto -ErrorAction Stop -UseBasicParsing
        Write-Host "  -> Download concluido!" -ForegroundColor Green
        return $caminhoCompleto
    }
    catch {
        Write-Host "  -> [ERRO] Falha ao baixar arquivo" -ForegroundColor Red
        Write-Host "  -> Detalhes: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host "Pressione ENTER para continuar"
        return $null
    }
}

function Remove-FilaDuplicada {
    param(
        [string]$nomeConfigurado,
        [string]$filtroDriver
    )
    
    Write-Host "  -> Verificando filas duplicadas..." -ForegroundColor Gray
    
    $todasImpressoras = Get-Printer -ErrorAction SilentlyContinue
    
    # Remover filas que atendam QUALQUER uma destas condicoes:
    # 1. Nome EXATO do filtroDriver (ex: "Samsung CLP-680 Series")
    # 2. Nome do filtroDriver com variacao (ex: "Samsung CLP-680 Series (Copia 1)")
    # 3. Drivers Universal Samsung que sobraram
    $duplicatas = $todasImpressoras | Where-Object {
        $_.Name -ne $nomeConfigurado -and
        (
            # Nome exato do FiltroDriver
            $_.Name -eq $filtroDriver -or
            # FiltroDriver com variacoes (Copia 1, Copia 2, etc)
            $_.Name -match "^$([regex]::Escape($filtroDriver))( PS| PCL[0-9].*)?( \((C[oó]pia|Copiar|Copy) \d+\))?$" -or
            # Drivers Universal sobrando
            ($_.Name -like "*Samsung Universal Print Driver*" -and $_.DriverName -like "*Samsung Universal*")
        ) -and
        $_.Name -notlike "*Fax*"
    }
    
    if ($duplicatas) {
        Write-Host "  -> Removendo filas duplicadas:" -ForegroundColor Yellow
        foreach ($fila in $duplicatas) {
            try {
                Remove-Printer -Name $fila.Name -Confirm:$false -ErrorAction Stop
                Write-Host "    [OK] Removida: $($fila.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Host "    [ERRO] Falha ao remover: $($fila.Name)" -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "  -> Nenhuma duplicata encontrada!" -ForegroundColor Green
    }
}

# --- CALCULO DE ETAPAS ---
$etapas = @()
if ($instalarPrint) { $etapas += "PRINT" }
if ($instalarScan)  { $etapas += "SCAN" }
if ($instalarEPM)   { $etapas += "EPM" }
if ($instalarEDC)   { $etapas += "EDC" }

$totalEtapas = $etapas.Count
$etapaAtual = 1

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "  PROCESSANDO $totalEtapas ETAPA(S) - $modelo" -ForegroundColor Cyan
Write-Host "========================================================`n" -ForegroundColor Cyan

# ================================================================================
# ETAPA 1: DRIVER DE IMPRESSAO
# ================================================================================

if ($instalarPrint) {
    Write-Host "[$etapaAtual/$totalEtapas] === DRIVER DE IMPRESSAO ===" -ForegroundColor Yellow
    Write-Host ""
    
    # Coletar e validar nome da impressora
    $nomeImpressora = Read-Host "  -> Nome da impressora"
    
    $impressoraExistente = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
    
    if ($impressoraExistente) {
        Write-Host "`n  Ja existe uma impressora com o nome '$nomeImpressora'!" -ForegroundColor Yellow
        Write-Host "`n  Detalhes da impressora existente:" -ForegroundColor Cyan
        Write-Host "    IP atual: $($impressoraExistente.PortName)" -ForegroundColor Gray
        Write-Host "    Driver: $($impressoraExistente.DriverName)" -ForegroundColor Gray
        Write-Host "`n  1) Digitar outro nome"
        Write-Host "  2) Cancelar instalacao"
        $opcao = Read-Host "`n  Escolha"
        
        if ($opcao -eq "1") {
            $nomeImpressora = Read-Host "`n  -> Novo nome da impressora"
        } else {
            Write-Host "`n  [INFO] Instalacao cancelada pelo usuario`n" -ForegroundColor Cyan
            $etapaAtual++
            Write-Host ""
            return
        }
    }

$enderecoIP = Read-Host "  -> Endereco IP"
    
    # Validar IP
    if ($enderecoIP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
        Write-Host "`n  [AVISO] IP invalido! Usando padrao 192.168.1.100" -ForegroundColor Yellow
        $enderecoIP = "192.168.1.100"
    }
    
    # Download do driver
    $nomeArquivo = "driver_print_" + ($modelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlPrint -nomeDestino $nomeArquivo
    
    if (-not $arquivoDriver) {
        Write-Host "  [ERRO] Nao foi possivel obter o arquivo. Pulando etapa...`n" -ForegroundColor Red
        $etapaAtual++
    }
    else {
        # Instalar driver
        Write-Host "  -> Instalando driver... (Aguarde)" -ForegroundColor Gray
        Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
        Start-Sleep -Seconds $Global:Config.TempoEspera
        
        # Criar porta IP
        if (-not (Get-PrinterPort $enderecoIP -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $enderecoIP -PrinterHostAddress $enderecoIP -ErrorAction SilentlyContinue
            Write-Host "  -> Porta IP criada: $enderecoIP" -ForegroundColor Gray
        }
        
        Write-Host "  -> Configurando fila de impressao..." -ForegroundColor Gray
        
        # Buscar fila criada pelo instalador
        # PRIORIDADE: Driver nativo (sem PCL/PS) > PCL > PS > Universal
        $filaEspecifica = Get-Printer -ErrorAction SilentlyContinue | 
                         Where-Object {
                             ($_.Name -like "*$filtroDriverWindows*" -or $_.DriverName -like "*$filtroDriverWindows*") -and
                             $_.DriverName -notlike "*PCL*" -and 
                             $_.DriverName -notlike "* PS"
                         } | Select-Object -First 1
        
        # Se nao encontrou driver nativo, buscar qualquer um com o filtro
        if (-not $filaEspecifica) {
            $filaEspecifica = Get-Printer -ErrorAction SilentlyContinue | 
                             Where-Object {
                                 $_.Name -like "*$filtroDriverWindows*" -or 
                                 $_.DriverName -like "*$filtroDriverWindows*"
                             } | Select-Object -First 1
        }
        
        $filaUniversal = $null
        if (-not $filaEspecifica) {
            $filaUniversal = Get-Printer -ErrorAction SilentlyContinue | 
                            Where-Object {
                                $_.DriverName -like "*Samsung Universal*" -or 
                                $_.Name -like "*Samsung Universal*"
                            } | Select-Object -First 1
        }
        
        # Configurar impressora
        try {
            if ($filaEspecifica) {
                # Verificar se o driver atual e PCL/PS e tentar trocar para nativo
                if ($filaEspecifica.DriverName -like "*PCL*" -or $filaEspecifica.DriverName -like "* PS") {
                    Write-Host "  -> Detectado driver PCL/PS. Buscando driver nativo..." -ForegroundColor Yellow
                    
                    # Buscar o driver nativo (sem PCL/PS no nome)
                    $driverNativo = Get-PrinterDriver -ErrorAction SilentlyContinue | 
                                   Where-Object { 
                                       $_.Name -like "*$filtroDriverWindows*" -and 
                                       $_.Name -notlike "*PCL*" -and 
                                       $_.Name -notlike "* PS"
                                   } | Select-Object -First 1
                    
                    if ($driverNativo) {
                        Set-Printer -Name $filaEspecifica.Name -DriverName $driverNativo.Name -PortName $enderecoIP -ErrorAction Stop
                        Write-Host "  -> Trocado para driver nativo: $($driverNativo.Name)" -ForegroundColor Green
                    } else {
                        # Se nao encontrar driver nativo, manter o atual
                        Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop
                        Write-Host "  [AVISO] Driver nativo nao encontrado. Mantendo: $($filaEspecifica.DriverName)" -ForegroundColor Yellow
                    }
                } else {
                    # Driver ja e nativo, apenas configurar IP
                    Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop
                }
                
                Rename-Printer -Name $filaEspecifica.Name -NewName $nomeImpressora -ErrorAction Stop
                Write-Host "  [OK] Impressora configurada com sucesso!" -ForegroundColor Green
            }
            elseif ($filaUniversal) {
                # Caso M4080: Fila Universal criada, tentar trocar driver
                try {
                    Set-Printer -Name $filaUniversal.Name -DriverName $filtroDriverWindows -PortName $enderecoIP -ErrorAction Stop
                    Rename-Printer -Name $filaUniversal.Name -NewName $nomeImpressora -ErrorAction Stop
                    Write-Host "  [OK] Fila Universal convertida para driver especifico!" -ForegroundColor Green
                }
                catch {
                    # Se falhar troca de driver, manter Universal
                    Write-Host "  [AVISO] Driver especifico nao disponivel. Usando Universal..." -ForegroundColor Yellow
                    Set-Printer -Name $filaUniversal.Name -PortName $enderecoIP -ErrorAction Stop
                    Rename-Printer -Name $filaUniversal.Name -NewName $nomeImpressora -ErrorAction Stop
                    Write-Host "  [OK] Impressora configurada com driver Universal!" -ForegroundColor Green
                }
            }
            else {
                # Nenhuma fila encontrada: criar manualmente
                Add-Printer -Name $nomeImpressora -DriverName $filtroDriverWindows -PortName $enderecoIP -ErrorAction Stop
                Write-Host "  [OK] Fila criada manualmente!" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  [ERRO] Falha ao configurar impressora" -ForegroundColor Red
            Write-Host "  -> Detalhes: $($_.Exception.Message)" -ForegroundColor Red
            
            # Listar drivers disponiveis
            Write-Host "`n  -> Drivers Samsung disponiveis no sistema:" -ForegroundColor Yellow
            Get-PrinterDriver -ErrorAction SilentlyContinue | 
                Where-Object { $_.Name -like "*Samsung*" } | 
                ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor Cyan }
            
            Read-Host "`n  Pressione ENTER para continuar"
        }
        
        # Limpar duplicatas (AGORA PASSA O FILTRO DO DRIVER)
        Start-Sleep -Seconds 2
        Remove-FilaDuplicada -nomeConfigurado $nomeImpressora -filtroDriver $filtroDriverWindows
        
        $etapaAtual++
        Write-Host ""
    }
}

# ================================================================================
# ETAPA 2: DRIVER DE DIGITALIZACAO
# ================================================================================

if ($instalarScan) {
    Write-Host "[$etapaAtual/$totalEtapas] === DRIVER DE DIGITALIZACAO ===" -ForegroundColor Yellow
    Write-Host ""
    
    if ([string]::IsNullOrWhiteSpace($urlScan)) {
        Write-Host "  [AVISO] URL de scan nao disponivel. Pulando...`n" -ForegroundColor Yellow
    }
    else {
        $nomeArquivoScan = "driver_scan_" + ($modelo -replace '\s+', '_') + ".exe"
        $arquivoScan = Get-ArquivoLocal -url $urlScan -nomeDestino $nomeArquivoScan
        
        if ($arquivoScan) {
            Write-Host "  -> Instalando driver de scan..." -ForegroundColor Gray
            Start-Process $arquivoScan -ArgumentList "/S" -Wait -NoNewWindow
            Write-Host "  [OK] Driver de scan instalado!" -ForegroundColor Green
        }
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# ETAPA 3: EASY PRINTER MANAGER
# ================================================================================

if ($instalarEPM) {
    Write-Host "[$etapaAtual/$totalEtapas] === EASY PRINTER MANAGER ===" -ForegroundColor Yellow
    Write-Host ""
    
    if (Test-ProgramaInstalado "Easy Printer Manager") {
        Write-Host "  -> Ja instalado no sistema!" -ForegroundColor Cyan
    }
    else {
        $arquivoEPM = Get-ArquivoLocal -url $Global:Config.UrlEPM -nomeDestino "EPM_Universal.exe"
        
        if ($arquivoEPM) {
            Write-Host "  -> Instalando Easy Printer Manager..." -ForegroundColor Gray
            Write-Host "  -> Aguardando instalacao (maximo 60 segundos)..." -ForegroundColor Gray
            
            # Iniciar processo em background
            $processo = Start-Process $arquivoEPM -ArgumentList "/S" -PassThru -NoNewWindow
            
            # Aguardar com timeout de 60 segundos
            $tempoLimite = 60
            $tempoDecorrido = 0
            
            while (-not $processo.HasExited -and $tempoDecorrido -lt $tempoLimite) {
                Start-Sleep -Seconds 2
                $tempoDecorrido += 2
                
                # Verificar se ja foi instalado mesmo com processo ativo
                if (Test-ProgramaInstalado "Easy Printer Manager") {
                    Write-Host "  [OK] Easy Printer Manager instalado com sucesso!" -ForegroundColor Green
                    
                    # Finalizar processo se ainda estiver ativo
                    if (-not $processo.HasExited) {
                        Write-Host "  -> Finalizando processo de instalacao..." -ForegroundColor Gray
                        Stop-Process -Id $processo.Id -Force -ErrorAction SilentlyContinue
                    }
                    break
                }
            }
            
            # Verificacao final
            if (-not (Test-ProgramaInstalado "Easy Printer Manager")) {
                Write-Host "  [AVISO] Instalacao pode nao ter sido concluida. Verifique manualmente." -ForegroundColor Yellow
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
    Write-Host "[$etapaAtual/$totalEtapas] === EASY DOCUMENT CREATOR ===" -ForegroundColor Yellow
    Write-Host ""
    
    if (Test-ProgramaInstalado "Easy Document Creator") {
        Write-Host "  -> Ja instalado no sistema!" -ForegroundColor Cyan
    }
    else {
        $arquivoEDC = Get-ArquivoLocal -url $Global:Config.UrlEDC -nomeDestino "EDC_Universal.exe"
        
        if ($arquivoEDC) {
            Write-Host "  -> Instalando Easy Document Creator..." -ForegroundColor Gray
            Write-Host "  -> Aguardando instalacao (maximo 60 segundos)..." -ForegroundColor Gray
            
            # Iniciar processo em background
            $processo = Start-Process $arquivoEDC -ArgumentList "/S" -PassThru -NoNewWindow
            
            # Aguardar com timeout de 60 segundos
            $tempoLimite = 60
            $tempoDecorrido = 0
            
            while (-not $processo.HasExited -and $tempoDecorrido -lt $tempoLimite) {
                Start-Sleep -Seconds 2
                $tempoDecorrido += 2
                
                # Verificar se ja foi instalado mesmo com processo ativo
                if (Test-ProgramaInstalado "Easy Document Creator") {
                    Write-Host "  [OK] Easy Document Creator instalado com sucesso!" -ForegroundColor Green
                    
                    # Finalizar processo se ainda estiver ativo
                    if (-not $processo.HasExited) {
                        Write-Host "  -> Finalizando processo de instalacao..." -ForegroundColor Gray
                        Stop-Process -Id $processo.Id -Force -ErrorAction SilentlyContinue
                    }
                    break
                }
            }
            
            # Verificacao final
            if (-not (Test-ProgramaInstalado "Easy Document Creator")) {
                Write-Host "  [AVISO] Instalacao pode nao ter sido concluida. Verifique manualmente." -ForegroundColor Yellow
            }
        }
    }
    
    $etapaAtual++
    Write-Host ""
}

# ================================================================================
# FINALIZACAO
# ================================================================================

Write-Host "========================================================" -ForegroundColor Green
Write-Host "  PROCESSO CONCLUIDO COM SUCESSO!" -ForegroundColor Green
Write-Host "========================================================" -ForegroundColor Green
Write-Host ""

Start-Sleep -Seconds 3







# ================================================================================
# SCRIPT: Motor de Instalacao Universal - Impressoras Samsung
# VERSAO: 4.0 (Nativo - Sem 7zip - Opcao Substituir)
# ================================================================================

param (
    [Parameter(Mandatory=$true)][string]$modelo,
    [Parameter(Mandatory=$true)][string]$urlPrint,
    [Parameter(Mandatory=$true)][string]$temScan,
    [Parameter(Mandatory=$false)][string]$urlScan = "",
    # NOVOS PARAMETROS PARA RECEBER DO ARQUIVO .SVC
    [Parameter(Mandatory=$false)][string]$urlEPM = "", 
    [Parameter(Mandatory=$false)][string]$urlEDC = "",
    
    [Parameter(Mandatory=$true)][string]$filtroDriverWindows,
    [Parameter(Mandatory=$true)][bool]$instalarPrint,
    [Parameter(Mandatory=$true)][bool]$instalarScan,
    [Parameter(Mandatory=$true)][bool]$instalarEPM,
    [Parameter(Mandatory=$true)][bool]$instalarEDC
)

# --- CONFIGURACAO GLOBAL ---
$Global:Config = @{
    # URLs padrao caso nao sejam passadas (Fallback)
    UrlEPM_Fallback = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
    UrlEDC_Fallback = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
    CaminhoTemp     = "$env:USERPROFILE\Downloads\Instalacao_Samsung"
    TempoEspera     = 5
}

# Define URL final baseada no parametro ou no fallback
$Global:UrlFinalEPM = if ([string]::IsNullOrWhiteSpace($urlEPM)) { $Global:Config.UrlEPM_Fallback } else { $urlEPM }
$Global:UrlFinalEDC = if ([string]::IsNullOrWhiteSpace($urlEDC)) { $Global:Config.UrlEDC_Fallback } else { $urlEDC }

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
        # Validacao simples de tamanho (evita arquivos corrompidos de 0kb)
        $tamanho = (Get-Item $caminhoCompleto).Length
        if ($tamanho -gt 1024) {
            Write-Mensagem "Reutilizando arquivo local: $nomeDestino" "Info"
            return $caminhoCompleto
        } else {
            Remove-Item $caminhoCompleto -Force
        }
    }
    
    try {
        Write-Host "Baixando: $nomeDestino..." -ForegroundColor Gray
        # Força protocolo de segurança para evitar erro em downloads HTTPS antigos
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
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

function Expand-ArquivoNativo {
    param(
        [string]$ArquivoOrigem,
        [string]$PastaDestino
    )
    
    # Metodo nativo do Windows (Renomear para .zip e usar Expand-Archive)
    # Funciona para a maioria dos wrappers Samsung/HP modernos
    
    try {
        if (Test-Path $PastaDestino) { Remove-Item $PastaDestino -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -Path $PastaDestino -ItemType Directory -Force | Out-Null
        
        # Copia como .zip temporario
        $zipTemp = Join-Path $Global:Config.CaminhoTemp ("temp_extract_" + [guid]::NewGuid().ToString() + ".zip")
        Copy-Item $ArquivoOrigem $zipTemp -Force
        
        Write-Host "Extraindo arquivos nativamente..." -ForegroundColor Gray
        Expand-Archive -Path $zipTemp -DestinationPath $PastaDestino -Force -ErrorAction Stop
        
        Remove-Item $zipTemp -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        Write-Mensagem "Tentativa de extracao nativa falhou. Tentando metodo alternativo..." "Aviso"
        # Fallback: Tentar executar o .exe com parametros comuns de extracao silenciosa
        try {
            Start-Process $ArquivoOrigem -ArgumentList "/extract_all:`"$PastaDestino`" /q" -Wait -NoNewWindow -ErrorAction Stop
            if ((Get-ChildItem $PastaDestino).Count -gt 0) { return $true }
        } catch {}
        
        return $false
    }
}

function New-PortaIP {
    param([Parameter(Mandatory=$true)][string]$enderecoIP)
    
    if (-not (Get-PrinterPort $enderecoIP -ErrorAction SilentlyContinue)) {
        try {
            Add-PrinterPort -Name $enderecoIP -PrinterHostAddress $enderecoIP -ErrorAction Stop
            return $true
        } catch {
            Write-Mensagem "Falha ao criar porta IP: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    return $true
}

function Test-RedeImpressora {
    param([Parameter(Mandatory=$true)][string]$enderecoIP)
    
    $meuIP = (Get-NetIPAddress -AddressFamily IPv4 | 
              Where-Object { $_.IPAddress -notlike "127.*" -and $_.PrefixOrigin -ne "WellKnown" } | 
              Select-Object -First 1).IPAddress
    
    if (-not $meuIP) { return $true }
    
    $minhaRede = ($meuIP -split '\.')[0..2] -join '.'
    $redeImpressora = ($enderecoIP -split '\.')[0..2] -join '.'
    
    # Mantida verificacao estrita conforme solicitado
    if ($minhaRede -ne $redeImpressora) {
        Write-Host ""
        Write-Mensagem "IP em rede diferente detectado!" "Aviso"
        Write-Host "  Seu computador: $meuIP"
        Write-Host "  IP digitado:    $enderecoIP"
        Write-Host "  Para redes domesticas, os 3 primeiros numeros devem ser iguais."
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim? [S/N]" @("S","s","N","n")
        Write-Host ""
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    $pingOk = Test-Connection -ComputerName $enderecoIP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $pingOk) {
        Write-Host ""
        Write-Mensagem "Impressora nao responde ao ping!" "Aviso"
        Write-Host "  Verifique se esta ligada e conectada a rede."
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim? [S/N]" @("S","s","N","n")
        Write-Host ""
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    return $true
}

function Remove-ImpressoraCompleta {
    param(
        [Parameter(Mandatory=$true)][string]$NomeOuIP,
        [Parameter(Mandatory=$true)][string]$TipoBusca # "Nome" ou "IP"
    )
    
    Write-Host "Removendo impressora antiga..." -ForegroundColor Gray
    
    try {
        $impressora = $null
        if ($TipoBusca -eq "Nome") {
            $impressora = Get-Printer -Name $NomeOuIP -ErrorAction SilentlyContinue
        } else {
            $impressora = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $NomeOuIP } | Select-Object -First 1
        }
        
        if ($impressora) {
            Remove-Printer -Name $impressora.Name -ErrorAction Stop
            Start-Sleep -Seconds 3 # Tempo para o spooler liberar
            Write-Mensagem "Fila de impressao removida." "Sucesso"
        }
        
        # Se for remocao por IP, tenta remover a porta tambem para garantir limpeza
        if ($TipoBusca -eq "IP" -or ($impressora -and $impressora.PortName -match '\d+\.\d+\.\d+\.\d+')) {
            $porta = if ($TipoBusca -eq "IP") { $NomeOuIP } else { $impressora.PortName }
            
            # So remove a porta se nao houver outra impressora usando ela
            $outras = Get-Printer | Where-Object { $_.PortName -eq $porta }
            if (-not $outras) {
                Remove-PrinterPort -Name $porta -ErrorAction SilentlyContinue
                Write-Host "Porta TCP/IP liberada." -ForegroundColor Gray
            }
        }
        return $true
    }
    catch {
        Write-Mensagem "Erro ao remover: $($_.Exception.Message)" "Erro"
        return $false
    }
}

function Test-DriverExistente {
    param([Parameter(Mandatory=$true)][string]$filtroDriver)
    
    $driver = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*$filtroDriver*" -and $_.Name -notlike "*PCL*" -and $_.Name -notlike "* PS"
    } | Select-Object -First 1
    
    if ($driver) { return @{ Encontrado = $true; Driver = $driver.Name } }
    return @{ Encontrado = $false; Driver = $null }
}

# --- LOGICA DE INSTALACAO (SIMPLIFICADA) ---
function Install-GenericDriver {
    param($urlDriver, $nomeModelo, $filtroDriver, $nomeImpressora, $enderecoIP, $isUPD)

    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    
    if ($statusDriver.Encontrado) {
        Write-Mensagem "Driver '$($statusDriver.Driver)' encontrado no sistema." "Info"
        New-PortaIP -enderecoIP $enderecoIP | Out-Null
        try {
            Add-Printer -Name $nomeImpressora -DriverName $statusDriver.Driver -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Impressora instalada com sucesso!" "Sucesso"
            return $true
        } catch {
            Write-Mensagem "Erro ao adicionar impressora: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    
    # Se nao achou driver, baixa e instala
    $nomeArquivo = "driver_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    if (-not $arquivoDriver) { return $false }
    
    $pastaExtracao = Join-Path $Global:Config.CaminhoTemp ("EXT_" + $nomeModelo.Replace(" ","_"))
    
    if ($isUPD) {
        # Para UPD: Extrai e usa PNPUtil (mais garantido)
        if (Expand-ArquivoNativo -ArquivoOrigem $arquivoDriver -PastaDestino $pastaExtracao) {
            $inf = Get-ChildItem -Path $pastaExtracao -Filter "*.inf" -Recurse | Where-Object { $_.FullName -notlike "*autorun*" } | Select-Object -First 1
            if ($inf) {
                Write-Host "Instalando driver via Repositorio (PNPUtil)..." -ForegroundColor Gray
                pnputil.exe /add-driver "$($inf.FullName)" /install | Out-Null
            }
            Remove-Item $pastaExtracao -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            # Falha na extracao, tenta instalacao padrao
             Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
        }
    } else {
        # Para SPL comum: Executa instalador silencioso
        Write-Host "Executando instalador do driver..." -ForegroundColor Gray
        Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
    }
    
    Start-Sleep -Seconds 5
    New-PortaIP -enderecoIP $enderecoIP | Out-Null
    
    # Tenta adicionar a impressora novamente apos instalacao do driver
    # Busca o nome exato do driver que acabou de ser instalado
    $driverRecemInstalado = Get-PrinterDriver | Where-Object { $_.Name -like "*$filtroDriver*" } | Select-Object -First 1
    
    if ($driverRecemInstalado) {
        try {
            Add-Printer -Name $nomeImpressora -DriverName $driverRecemInstalado.Name -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Impressora instalada com sucesso!" "Sucesso"
            return $true
        } catch {
            Write-Mensagem "Erro final: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    
    # Se chegou aqui, driver nao apareceu ou falhou
    Write-Mensagem "O driver foi instalado, mas nao foi possivel criar a fila automaticamente." "Aviso"
    Write-Host "Tente adicionar manualmente selecionando o driver Samsung."
    return $false
}

function Show-ResumoInstalacao {
    param($modelo, $nomeImpressora, $enderecoIP, $componentes)
    Write-Host "`n========================================"
    Write-Host "       RESUMO DA INSTALACAO"
    Write-Host "========================================"
    Write-Host "Modelo:      $modelo"
    Write-Host "Nome:        $nomeImpressora"
    Write-Host "IP:          $enderecoIP"
    Write-Host "Componentes: $($componentes -join ', ')"
    Write-Host "Status:      " -NoNewline
    Write-Host "CONCLUIDO" -ForegroundColor Green
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
Write-Host "================================================`n"

# ================================================================================
# ETAPA 1: DRIVER DE IMPRESSAO (COM LOGICA DE SUBSTITUICAO)
# ================================================================================

$instalacaoSucesso = $false
$nomeImpressora = ""
$enderecoIP = ""

if ($instalarPrint) {
    Write-Host "[$etapaAtual/$totalEtapas] CONFIGURACAO DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host ""
    
    # --- LOOP PARA NOME ---
    do {
        $nomeImpressora = Read-Host "- Digite o NOME desejado para a impressora"
        
        if ([string]::IsNullOrWhiteSpace($nomeImpressora)) { continue }
        
        $existeNome = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
        
        if ($existeNome) {
            Write-Host ""
            Write-Mensagem "Ja existe uma impressora com este nome!" "Aviso"
            Write-Host "  Nome: $($existeNome.Name)"
            Write-Host "  IP:   $($existeNome.PortName)"
            Write-Host ""
            Write-Host "[1] Digitar outro nome"
            Write-Host "[2] SUBSTITUIR (Remover a antiga e usar este nome)" -ForegroundColor Yellow
            Write-Host "[3] Cancelar instalacao"
            
            $opcao = Read-OpcaoValidada "Escolha uma opcao [1-3]:" @("1","2","3")
            
            if ($opcao -eq "3") { Write-Mensagem "Cancelado pelo usuario." "Info"; return }
            if ($opcao -eq "2") {
                if (Remove-ImpressoraCompleta -NomeOuIP $nomeImpressora -TipoBusca "Nome") {
                    break # Sai do loop e usa o nome
                } else {
                    Write-Host "Nao foi possivel substituir. Tente outro nome."
                }
            }
            # Se opcao 1, o loop roda de novo
        } else {
            break # Nome livre
        }
    } while ($true)
    
    Write-Host ""
    
    # --- LOOP PARA IP ---
    do {
        $enderecoIP = Read-Host "- Digite o ENDERECO IP da impressora"
        
        # Regex corrigido e rigoroso
        if ($enderecoIP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
            Write-Mensagem "IP invalido! Use formato XXX.XXX.XXX.XXX" "Erro"
            continue
        }
        
        $existeIP = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $enderecoIP } | Select-Object -First 1
        
        if ($existeIP) {
            Write-Host ""
            Write-Mensagem "Ja existe uma impressora usando este IP!" "Aviso"
            Write-Host "  Impressora atual: $($existeIP.Name)"
            Write-Host "  IP Ocupado:       $($existeIP.PortName)"
            Write-Host ""
            Write-Host "[1] Digitar outro IP"
            Write-Host "[2] SUBSTITUIR (Remover impressora antiga e usar este IP)" -ForegroundColor Yellow
            Write-Host "[3] Cancelar instalacao"
            
            $opcao = Read-OpcaoValidada "Escolha uma opcao [1-3]:" @("1","2","3")
            
            if ($opcao -eq "3") { Write-Mensagem "Cancelado pelo usuario." "Info"; return }
            if ($opcao -eq "2") {
                if (Remove-ImpressoraCompleta -NomeOuIP $enderecoIP -TipoBusca "IP") {
                    break # Sai do loop e usa o IP
                }
            }
            continue
        }
        
        # Validação de rede domestica (3 octetos)
        if (-not (Test-RedeImpressora -enderecoIP $enderecoIP)) {
            Write-Mensagem "Instalacao cancelada na verificacao de rede." "Info"
            return
        }
        
        break
    } while ($true)
    
    # Instala usando a nova funcao otimizada sem 7zip
    $instalacaoSucesso = Install-GenericDriver `
        -urlDriver $urlPrint `
        -nomeModelo $modelo `
        -filtroDriver $filtroDriverWindows `
        -nomeImpressora $nomeImpressora `
        -enderecoIP $enderecoIP `
        -isUPD ($Global:TipoDriver -eq "UPD")

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
        Write-Mensagem "URL de scan nao fornecida." "Aviso"
    } else {
        # Validacao simples se ja existe algo do modelo instalado
        $programasScan = Test-ProgramaInstalado -nomePrograma $modelo
        
        if ($programasScan -and (Test-ProgramaInstalado -nomePrograma "Scan")) {
            Write-Mensagem "Software de Scan ja detectado." "Sucesso"
        } else {
            $nomeArquivoScan = "driver_scan_" + ($modelo -replace '\s+', '_') + ".exe"
            $arquivoScan = Get-ArquivoLocal -url $urlScan -nomeDestino $nomeArquivoScan
            
            if ($arquivoScan) {
                Write-Host "Instalando driver de scan..." -ForegroundColor Gray
                Start-Process $arquivoScan -ArgumentList "/S" -Wait -NoNewWindow
                Write-Mensagem "Driver de scan instalado!" "Sucesso"
            }
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
    
    if (Test-ProgramaInstalado "Easy Printer Manager") {
        Write-Host "Ja instalado no sistema." -ForegroundColor Gray
    } else {
        $arquivoEPM = Get-ArquivoLocal -url $Global:UrlFinalEPM -nomeDestino "EPM_Universal.exe"
        if ($arquivoEPM) {
            Write-Host "Iniciando instalacao (Aguarde)..." -ForegroundColor Gray
            Start-Process $arquivoEPM -ArgumentList "/S" -Wait -NoNewWindow
            Write-Mensagem "Instalacao finalizada." "Sucesso"
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
    
    if (Test-ProgramaInstalado "Easy Document Creator") {
        Write-Host "Ja instalado no sistema." -ForegroundColor Gray
    } else {
        $arquivoEDC = Get-ArquivoLocal -url $Global:UrlFinalEDC -nomeDestino "EDC_Universal.exe"
        if ($arquivoEDC) {
            Write-Host "Iniciando instalacao (Aguarde)..." -ForegroundColor Gray
            Start-Process $arquivoEDC -ArgumentList "/S" -Wait -NoNewWindow
            Write-Mensagem "Instalacao finalizada." "Sucesso"
        }
    }
    Write-Host ""
}

# ================================================================================
# FINALIZACAO
# ================================================================================

if ($instalarPrint -and $instalacaoSucesso) {
    Show-ResumoInstalacao -modelo $modelo `
                         -nomeImpressora $nomeImpressora `
                         -enderecoIP $enderecoIP `
                         -componentes $componentesInstalados
    
    $imprimirTeste = Read-OpcaoValidada "Deseja imprimir uma pagina de teste? [S/N]" @("S","s","N","n")
    
    if ($imprimirTeste -eq "S" -or $imprimirTeste -eq "s") {
        try {
            Start-Process "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /k /n `"$nomeImpressora`"" -NoNewWindow
            Write-Mensagem "Pagina de teste enviada!" "Sucesso"
        } catch {
            Write-Mensagem "Erro ao enviar teste." "Erro"
        }
    }
}
elseif ($instalarPrint -and -not $instalacaoSucesso) {
    Write-Host "`n[FALHA] A instalacao da impressora nao foi concluida corretamente." -ForegroundColor Red
}

Write-Host ""
Start-Sleep -Seconds 3

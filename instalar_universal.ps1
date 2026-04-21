# ================================================================================
# SCRIPT: Motor de Instalacao Universal - Impressoras Samsung
# VERSAO: 3.4 (Corrigido - Sem Duplicacoes)
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
        Write-Mensagem "Reutilizando arquivo local: $nomeDestino" "Info"
        return $caminhoCompleto
    }
    
    try {
        Write-Host "Baixando: $nomeDestino..." -ForegroundColor Gray
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
    
    if (-not $meuIP) {
        return $true
    }
    
    $minhaRede = ($meuIP -split '\.')[0..2] -join '.'
    $redeImpressora = ($enderecoIP -split '\.')[0..2] -join '.'
    
    if ($minhaRede -ne $redeImpressora) {
        Write-Host ""
        Write-Mensagem "IP em rede diferente detectado!" "Aviso"
        Write-Host "  Seu computador: $meuIP"
        Write-Host "  IP digitado: $enderecoIP"
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim? [S/N]" @("S","s","N","n")
        Write-Host ""
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    $pingOk = Test-Connection -ComputerName $enderecoIP -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    if (-not $pingOk) {
        Write-Host ""
        Write-Mensagem "Impressora nao responde ao ping!" "Aviso"
        Write-Host "  Verifique se esta ligada e conectada a rede"
        
        $continuar = Read-OpcaoValidada "`nContinuar mesmo assim? [S/N]" @("S","s","N","n")
        Write-Host ""
        return ($continuar -eq "S" -or $continuar -eq "s")
    }
    
    return $true
}

function Test-DriverExistente {
    param([Parameter(Mandatory=$true)][string]$filtroDriver)
    
    $driverNativo = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*$filtroDriver*" -and
        $_.Name -notlike "*PCL*" -and
        $_.Name -notlike "* PS"
    } | Select-Object -First 1
    
    if ($driverNativo) {
        return @{ Encontrado = $true; Driver = $driverNativo.Name; Tipo = "Nativo" }
    }
    
    $driverVariacao = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*$filtroDriver*"
    } | Select-Object -First 1
    
    if ($driverVariacao) {
        return @{ Encontrado = $true; Driver = $driverVariacao.Name; Tipo = "Variacao" }
    }
    
    return @{ Encontrado = $false; Driver = $null; Tipo = $null }
}

function Test-DriverScanExistente {
    param([Parameter(Mandatory=$true)][string]$nomeModelo)
    
    $filtroScan = $nomeModelo -replace '\s+', '.*'
    
    $programas = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", 
                                  "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like "*$nomeModelo*" -and $_.DisplayName -like "*Scan*" }
    
    return [bool]$programas
}

function Install-DriverSPL {
    param(
        [string]$urlDriver,
        [string]$nomeModelo,
        [string]$filtroDriver,
        [string]$nomeImpressora,
        [string]$enderecoIP
    )
    
    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    
    if ($statusDriver.Encontrado -and $statusDriver.Tipo -eq "Nativo") {
        Write-Host ""
        Write-Mensagem "Driver '$($statusDriver.Driver)' ja presente no sistema" "Sucesso"
        Write-Host "Configurando impressora..." -ForegroundColor Gray
        
        New-PortaIP -enderecoIP $enderecoIP | Out-Null
        
        try {
            Add-Printer -Name $nomeImpressora -DriverName $statusDriver.Driver -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Impressora configurada com sucesso!" "Sucesso"
            return $true
        } catch {
            Write-Mensagem "Falha ao criar impressora: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    
    Write-Host ""
    $nomeArquivo = "driver_print_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    
    if (-not $arquivoDriver) { return $false }
    
    Write-Host "Instalando driver (aguarde)..." -ForegroundColor Gray
    Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
    Start-Sleep -Seconds $Global:Config.TempoEspera
    
    New-PortaIP -enderecoIP $enderecoIP | Out-Null
    
    Write-Host "Configurando fila de impressao..." -ForegroundColor Gray
    
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
                }
            } else {
                Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop
            }
            
            Rename-Printer -Name $filaEspecifica.Name -NewName $nomeImpressora -ErrorAction Stop
            Write-Mensagem "Impressora configurada com sucesso!" "Sucesso"
            return $true
        } else {
            Add-Printer -Name $nomeImpressora -DriverName $filtroDriver -PortName $enderecoIP -ErrorAction Stop
            Write-Mensagem "Impressora configurada com sucesso!" "Sucesso"
            return $true
        }
    }
    catch {
        Write-Mensagem "Falha ao configurar impressora: $($_.Exception.Message)" "Erro"
        
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

    Write-Host ""
    Show-Titulo -titulo "DEBUG DRIVER UPD" -cor "Yellow"
    Write-Mensagem "Modelo: $nomeModelo" "Info"
    Write-Mensagem "Filtro esperado: $filtroDriver" "Info"
    Write-Mensagem "Nome da fila: $nomeImpressora" "Info"
    Write-Mensagem "IP informado: $enderecoIP" "Info"

    Write-Mensagem "ETAPA 1/10 - Verificando se o driver ja existe no sistema..." "Info"
    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    if ($statusDriver.Encontrado) {
        Write-Mensagem "Driver encontrado previamente: $($statusDriver.Driver) | Tipo: $($statusDriver.Tipo)" "Aviso"
    }
    else {
        Write-Mensagem "Nenhum driver atual corresponde ao filtro informado." "Info"
    }

    if ($statusDriver.Encontrado -and $statusDriver.Tipo -eq "Nativo") {
        Write-Mensagem "Driver nativo ja presente no sistema. Pulando instalacao do pacote." "Sucesso"
        Write-Host "Configurando impressora..." -ForegroundColor Gray
        return (Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -driverPreferencial $statusDriver.Driver -aceitaUniversal $true)
    }

    Write-Mensagem "ETAPA 2/10 - Baixando pacote do driver UPD..." "Info"
    $nomeArquivo = "driver_UPD_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    if (-not $arquivoDriver) {
        Write-Mensagem "Falha no download do pacote UPD." "Erro"
        return $false
    }

    try {
        $infoArquivoDriver = Get-Item $arquivoDriver -ErrorAction Stop
        Write-Mensagem "Pacote local: $($infoArquivoDriver.FullName)" "Sucesso"
        Write-Mensagem "Tamanho do pacote: $(Format-TamanhoArquivo -bytes $infoArquivoDriver.Length)" "Info"
    }
    catch {
        Write-Mensagem "Nao foi possivel obter detalhes do pacote baixado." "Aviso"
    }

    $nomePastaExtracao = "driver_UPD_" + ($nomeModelo -replace '\s+', '_')
    $pastaExtracao = Join-Path $Global:Config.CaminhoTemp $nomePastaExtracao

    Write-Mensagem "ETAPA 3/10 - Preparando pasta de extracao..." "Info"
    if (Test-Path $pastaExtracao) {
        Write-Mensagem "Removendo extracao anterior: $pastaExtracao" "Info"
        Remove-Item $pastaExtracao -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -Path $pastaExtracao -ItemType Directory -Force | Out-Null
    Write-Mensagem "Pasta pronta: $pastaExtracao" "Sucesso"

    Write-Mensagem "ETAPA 4/10 - Verificando 7-Zip..." "Info"
    $controle7Zip = Ensure-7Zip
    $caminho7Zip = $controle7Zip.Caminho
    $instaladoPeloScript = $controle7Zip.InstaladoPeloScript
    if (-not $caminho7Zip) {
        Write-Mensagem "7-Zip nao foi encontrado ou instalado." "Erro"
        return $false
    }
    Write-Mensagem "7-Zip em uso: $caminho7Zip" "Sucesso"
    if ($instaladoPeloScript) {
        Write-Mensagem "7-Zip foi instalado temporariamente pelo script." "Aviso"
    }
    else {
        Write-Mensagem "7-Zip ja existia no sistema." "Info"
    }

    Write-Mensagem "ETAPA 5/10 - Extraindo pacote UPD com 7-Zip..." "Info"
    $saidaExtracao = & $caminho7Zip x "$arquivoDriver" "-o$pastaExtracao" -y 2>&1
    $codigoExtracao = $LASTEXITCODE
    if ($saidaExtracao) {
        Write-Host ($saidaExtracao | Out-String) -ForegroundColor DarkGray
    }
    Write-Mensagem "Codigo de saida da extracao: $codigoExtracao" "Info"
    Start-Sleep -Seconds 2

    $itensExtraidos = Get-ChildItem -Path $pastaExtracao -Recurse -Force -ErrorAction SilentlyContinue
    if (-not $itensExtraidos -or $itensExtraidos.Count -eq 0) {
        Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
        Write-Mensagem "Falha na extracao do pacote UPD. Pasta de extracao vazia." "Erro"
        return $false
    }
    Write-Mensagem "Extracao concluida com $($itensExtraidos.Count) itens." "Sucesso"

    $primeirosArquivos = $itensExtraidos | Select-Object -First 15
    if ($primeirosArquivos) {
        Write-Host "Primeiros itens extraidos:" -ForegroundColor DarkGray
        foreach ($item in $primeirosArquivos) {
            Write-Host "  - $($item.FullName)" -ForegroundColor DarkGray
        }
    }

    Write-Mensagem "ETAPA 6/10 - Localizando arquivos INF..." "Info"
    $arquivosInf = Get-ChildItem -Path $pastaExtracao -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -notlike "*autorun*" -and $_.Name -notlike "*setup*" }

    if (-not $arquivosInf) {
        Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
        Write-Mensagem "Nenhum arquivo INF foi encontrado na extracao." "Erro"
        return $false
    }

    Write-Mensagem "Quantidade de INFs encontrados: $($arquivosInf.Count)" "Sucesso"
    foreach ($inf in ($arquivosInf | Select-Object -First 20)) {
        Write-Host "  INF: $($inf.FullName)" -ForegroundColor DarkGray
    }

    Write-Mensagem "ETAPA 7/10 - Procurando o filtro '$filtroDriver' dentro dos INFs..." "Info"
    $infCompativeis = foreach ($inf in $arquivosInf) {
        try {
            $conteudo = Get-Content -Path $inf.FullName -Raw -ErrorAction Stop
            if ($conteudo -match [regex]::Escape($filtroDriver)) {
                [pscustomobject]@{
                    Arquivo = $inf
                    Caminho = $inf.FullName
                    PCL = ($inf.FullName -match '(?i)(\\|/|_|-)PCL(\\|/|_|-|$)' -or $inf.DirectoryName -match '(?i)\\PCL($|\\)')
                    PS  = ($inf.FullName -match '(?i)(\\|/|_|-)PS(\\|/|_|-|$)' -or $inf.DirectoryName -match '(?i)\\PS($|\\)')
                }
            }
        }
        catch {
            Write-Host "  Falha ao ler INF: $($inf.FullName)" -ForegroundColor DarkYellow
        }
    }

    if (-not $infCompativeis) {
        Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
        Write-Mensagem "Nenhum INF continha o filtro '$filtroDriver'." "Erro"
        return $false
    }

    Write-Mensagem "INFs compativeis encontrados: $($infCompativeis.Count)" "Sucesso"
    foreach ($item in $infCompativeis) {
        $tipoInf = if ($item.PCL) { 'PCL' } elseif ($item.PS) { 'PS' } else { 'NEUTRO' }
        Write-Host "  Compativel [$tipoInf]: $($item.Caminho)" -ForegroundColor DarkGray
    }

    $infEscolhido = $infCompativeis |
        Sort-Object @{Expression = {
            if ($_.PCL) { 0 }
            elseif (-not $_.PS) { 1 }
            else { 2 }
        }} |
        Select-Object -First 1

    if (-not $infEscolhido) {
        Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
        Write-Mensagem "Nao foi possivel escolher um INF valido." "Erro"
        return $false
    }

    $infEspecifico = $infEscolhido.Arquivo
    Write-Mensagem "INF escolhido: $($infEspecifico.FullName)" "Sucesso"
    if ($infEscolhido.PCL) {
        Write-Mensagem "Prioridade aplicada: PCL" "Info"
    }
    elseif ($infEscolhido.PS) {
        Write-Mensagem "Prioridade aplicada: PS" "Aviso"
    }
    else {
        Write-Mensagem "Prioridade aplicada: INF neutro" "Info"
    }

    Write-Mensagem "ETAPA 8/10 - Adicionando pacote ao Driver Store com pnputil..." "Info"
    $saidaPnputil = & pnputil.exe /add-driver "$($infEspecifico.FullName)" /install 2>&1
    $codigoPnputil = $LASTEXITCODE
    if ($saidaPnputil) {
        Write-Host ($saidaPnputil | Out-String) -ForegroundColor DarkGray
    }
    Write-Mensagem "Codigo de saida do pnputil: $codigoPnputil" "Info"

    Start-Sleep -Seconds $Global:Config.TempoEspera

    Write-Mensagem "ETAPA 9/10 - Verificando se o driver apareceu no sistema..." "Info"
    $driversSistema = Get-PrinterDriver -ErrorAction SilentlyContinue
    $driverExato = $driversSistema | Where-Object { $_.Name -eq $filtroDriver } | Select-Object -First 1
    $driversParciais = $driversSistema | Where-Object { $_.Name -like "*$filtroDriver*" }

    if ($driverExato) {
        Write-Mensagem "Driver encontrado por nome exato: $($driverExato.Name)" "Sucesso"
        $driverPreferencial = $driverExato.Name
    }
    elseif ($driversParciais) {
        Write-Mensagem "Driver nao apareceu por nome exato, mas houve correspondencia parcial." "Aviso"
        foreach ($drv in $driversParciais) {
            Write-Host "  Driver parcial: $($drv.Name)" -ForegroundColor DarkGray
        }
        $driverPreferencial = ($driversParciais | Select-Object -First 1).Name
        Write-Mensagem "Driver selecionado para teste: $driverPreferencial" "Aviso"
    }
    else {
        Write-Mensagem "Nenhum driver de impressao apareceu apos o pnputil." "Erro"
        Write-Host "Drivers Samsung existentes no sistema:" -ForegroundColor DarkGray
        $driversSistema | Where-Object { $_.Name -like "*Samsung*" -or $_.Name -like "*$fabricante*" } | ForEach-Object {
            Write-Host "  - $($_.Name)" -ForegroundColor DarkGray
        }
        Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
        return $false
    }

    Write-Mensagem "ETAPA 10/10 - Configurando fila de impressao com o driver localizado..." "Info"
    $sucesso = Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -driverPreferencial $driverPreferencial -aceitaUniversal $true

    Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript

    if ($sucesso) {
        Write-Mensagem "Fila configurada com sucesso usando '$driverPreferencial'." "Sucesso"
    }
    else {
        Write-Mensagem "Falha ao configurar a fila apos localizar o driver." "Erro"
    }

    return $sucesso
}

function Remove-ImpressoraExistente {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Nome","IP")]
        [string]$tipoBusca,
        
        [Parameter(Mandatory=$true)]
        [string]$valor
    )
    
    try {
        if ($tipoBusca -eq "Nome") {
            $impressora = Get-Printer -Name $valor -ErrorAction Stop
        } else {
            $impressora = Get-Printer -ErrorAction Stop | 
                         Where-Object { $_.PortName -eq $valor } | 
                         Select-Object -First 1
        }
        
        if ($impressora) {
            Write-Host "Removendo impressora: $($impressora.Name)..." -ForegroundColor Gray
            Remove-Printer -Name $impressora.Name -Confirm:$false -ErrorAction Stop
            Write-Mensagem "Impressora removida com sucesso!" "Sucesso"
            Start-Sleep -Seconds 1
            return $true
        }
        
        return $false
    }
    catch {
        Write-Mensagem "Erro ao remover impressora: $($_.Exception.Message)" "Erro"
        return $false
    }
}

function Remove-FilaDuplicada {
    param([string]$nomeConfigurado, [string]$filtroDriver)
    
    $todasImpressoras = Get-Printer -ErrorAction SilentlyContinue
    
    $duplicatas = $todasImpressoras | Where-Object {
        $_.Name -ne $nomeConfigurado -and
        (
            $_.Name -eq $filtroDriver -or
            $_.Name -match "^$([regex]::Escape($filtroDriver))( PS| PCL[0-9].*)?( \((Copia|Copiar|Copy) \d+\))?$" -or
            ($_.Name -like "*Samsung Universal Print Driver*" -and $_.DriverName -like "*Samsung Universal*")
        ) -and
        $_.Name -notlike "*Fax*"
    }
    
    if ($duplicatas) {
        Write-Host "Removendo filas duplicadas..." -ForegroundColor Gray
        foreach ($fila in $duplicatas) {
            try {
                Remove-Printer -Name $fila.Name -Confirm:$false -ErrorAction Stop
            } catch { }
        }
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
Write-Host "================================================`n" -ForegroundColor Cyan

# ================================================================================
# ETAPA 1: DRIVER DE IMPRESSAO
# ================================================================================

$driverInstalado = $null
$nomeImpressora = ""
$enderecoIP = ""
$instalacaoSucesso = $false

if ($instalarPrint) {
    Write-Host "[$etapaAtual/$totalEtapas] DRIVER DE IMPRESSAO" -ForegroundColor Yellow
    Write-Host ""
    
    do {
        $nomeImpressora = Read-Host "- Nome da impressora"
        Write-Host ""
        
        $impressoraExistente = Get-Printer -Name $nomeImpressora -ErrorAction SilentlyContinue
        
        if ($impressoraExistente) {
            Write-Mensagem "Ja existe impressora com nome '$nomeImpressora'" "Aviso"
            Write-Host "  IP atual: $($impressoraExistente.PortName)"
            Write-Host "  Driver:   $($impressoraExistente.DriverName)`n"
            
            $opcao = Read-OpcaoValidada "[1] Digitar outro nome  [2] Apagar e prosseguir  [3] Cancelar" @("1","2","3")
            Write-Host ""
            
            if ($opcao -eq "3") {
                Write-Mensagem "Instalacao cancelada" "Info"
                return
            }
            elseif ($opcao -eq "2") {
                if (Remove-ImpressoraExistente -tipoBusca "Nome" -valor $nomeImpressora) {
                    break
                } else {
                    Write-Host "Tente novamente com outro nome.`n" -ForegroundColor Yellow
                }
            }
        } else {
            break
        }
    } while ($true)
    
    do {
        $enderecoIP = Read-Host "- Endereco IP"
        Write-Host ""
        
        if ($enderecoIP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
            Write-Mensagem "IP invalido! Use formato XXX.XXX.XXX.XXX" "Erro"
            Write-Host ""
            continue
        }
        
        $impressoraMesmoIP = Get-Printer -ErrorAction SilentlyContinue | 
                            Where-Object { $_.PortName -eq $enderecoIP } | 
                            Select-Object -First 1
        
        if ($impressoraMesmoIP) {
            Write-Mensagem "Uma impressora com esse mesmo IP foi detectada no sistema!" "Aviso"
            Write-Host "  Nome:   $($impressoraMesmoIP.Name)"
            Write-Host "  Driver: $($impressoraMesmoIP.DriverName)`n"
            
            $opcao = Read-OpcaoValidada "[1] Digitar outro IP  [2] Apagar e prosseguir  [3] Cancelar" @("1","2","3")
            Write-Host ""
            
            if ($opcao -eq "3") {
                Write-Mensagem "Instalacao cancelada" "Info"
                return
            }
            elseif ($opcao -eq "2") {
                if (Remove-ImpressoraExistente -tipoBusca "IP" -valor $enderecoIP) {
                    break
                } else {
                    Write-Host "Tente novamente com outro IP.`n" -ForegroundColor Yellow
                }
            }
            continue
        }
        
        if (-not (Test-RedeImpressora -enderecoIP $enderecoIP)) {
            Write-Mensagem "Instalacao cancelada pelo usuario" "Info"
            return
        }
        
        break
    } while ($true)
    
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
        
        if ($impressoraFinal) {
            $driverInstalado = $impressoraFinal.DriverName
            $instalacaoSucesso = $true
            
            Start-Sleep -Seconds 2
            Remove-FilaDuplicada -nomeConfigurado $nomeImpressora -filtroDriver $filtroDriverWindows
        } else {
            Write-Mensagem "Impressora nao foi encontrada apos instalacao" "Erro"
        }
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
        if (Test-DriverScanExistente -nomeModelo $modelo) {
            Write-Mensagem "Driver de scan ja presente no sistema" "Sucesso"
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
    Write-Host ""
    
    if (Test-ProgramaInstalado "Easy Printer Manager") {
        Write-Host "Ja instalado no sistema" -ForegroundColor Gray
    } else {
        $arquivoEPM = Get-ArquivoLocal -url $Global:Config.UrlEPM -nomeDestino "EPM_Universal.exe"
        
        if ($arquivoEPM) {
            Write-Host "Instalando (timeout: 60s)..." -ForegroundColor Gray
            
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
        Write-Host "Ja instalado no sistema" -ForegroundColor Gray
    } else {
        $arquivoEDC = Get-ArquivoLocal -url $Global:Config.UrlEDC -nomeDestino "EDC_Universal.exe"
        
        if ($arquivoEDC) {
            Write-Host "Instalando (timeout: 60s)..." -ForegroundColor Gray
            
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
    
    Write-Host ""
}

# ================================================================================
# FINALIZACAO
# ================================================================================

if ($instalarPrint -and $instalacaoSucesso) {
    Show-ResumoInstalacao -modelo $modelo `
                         -nomeImpressora $nomeImpressora `
                         -enderecoIP $enderecoIP `
                         -driver $driverInstalado `
                         -componentes $componentesInstalados
    
    $imprimirTeste = Read-OpcaoValidada "Deseja imprimir uma pagina de teste? [S/N]" @("S","s","N","n")
    
    if ($imprimirTeste -eq "S" -or $imprimirTeste -eq "s") {
        try {
            Start-Process -FilePath "rundll32.exe" `
                         -ArgumentList "printui.dll,PrintUIEntry /k /n `"$nomeImpressora`"" `
                         -NoNewWindow -Wait
            Write-Host ""
            Write-Mensagem "Pagina de teste enviada!" "Sucesso"
        } catch {
            Write-Mensagem "Falha ao enviar pagina de teste" "Erro"
        }
    }
}
elseif ($instalarPrint -and -not $instalacaoSucesso) {
    Write-Host "`n========================================"
    Write-Host "     FALHA NA INSTALACAO"
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "A instalacao da impressora nao foi concluida."
    Write-Host "Verifique os erros acima."
    Write-Host "========================================`n"
}

Write-Host ""
Start-Sleep -Seconds 2


# ================================================================================
# SCRIPT: Motor de Instalacao Universal - Impressoras
# VERSAO: 4.0
# DESCRICAO: Instalador modular para drivers de impressoras SPL e UPD
# ================================================================================

param (
    [Parameter(Mandatory=$true)][string]$modelo,
    [Parameter(Mandatory=$false)][string]$fabricante = "",
    [Parameter(Mandatory=$true)][string]$urlPrint,
    [Parameter(Mandatory=$true)][string]$temScan,
    [Parameter(Mandatory=$false)][string]$urlScan = "",
    [Parameter(Mandatory=$true)][string]$filtroDriverWindows,
    [Parameter(Mandatory=$true)][bool]$instalarPrint,
    [Parameter(Mandatory=$true)][bool]$instalarScan,
    [Parameter(Mandatory=$true)][bool]$instalarEPM,
    [Parameter(Mandatory=$true)][bool]$instalarEDC
)

$Global:Config = @{
    UrlEPM        = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
    UrlEDC        = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
    Url7Zip       = "https://github.com/ip7z/7zip/releases/download/26.00/7z2600-x64.exe"
    CaminhoTemp   = "$env:USERPROFILE\Downloads\Instalacao_Impressoras"
    TempoEspera   = 10
    TempoPing     = 2
    LarguraUI     = 48
}

$Global:TipoDriver = if ($modelo -match "M4080|CLX[ -]?6260") { "UPD" } else { "SPL" }

if (-not (Test-Path $Global:Config.CaminhoTemp)) {
    New-Item $Global:Config.CaminhoTemp -ItemType Directory -Force | Out-Null
}

function Write-Linha {
    param([string]$cor = "Gray")
    Write-Host ("=" * $Global:Config.LarguraUI) -ForegroundColor $cor
}

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
        default    { "[INFO]" }
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

function Show-Titulo {
    param([string]$titulo,[string]$cor = "Cyan")
    Write-Host ""
    Write-Linha -cor $cor
    Write-Host ("  " + $titulo) -ForegroundColor $cor
    Write-Linha -cor $cor
    Write-Host ""
}

function Test-ProgramaInstalado {
    param([string]$nomePrograma)
    $chaves = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $programas = Get-ItemProperty $chaves -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$nomePrograma*" }
    return [bool]$programas
}

function Get-MeuIPv4 {
    $ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
           Where-Object {
               $_.IPAddress -notlike "127.*" -and
               $_.PrefixOrigin -ne "WellKnown" -and
               $_.ValidLifetime -gt 0
           }
    return ($ips | Select-Object -First 1).IPAddress
}

function Get-TamanhoArquivoRemoto {
    param([string]$url)
    try {
        $request = [System.Net.HttpWebRequest]::Create($url)
        $request.Method = "HEAD"
        $response = $request.GetResponse()
        $length = $response.ContentLength
        $response.Close()
        return $length
    }
    catch {
        return -1
    }
}

function Format-TamanhoArquivo {
    param([double]$bytes)
    if ($bytes -lt 1KB) { return "{0:N0} B" -f $bytes }
    if ($bytes -lt 1MB) { return "{0:N1} KB" -f ($bytes / 1KB) }
    if ($bytes -lt 1GB) { return "{0:N1} MB" -f ($bytes / 1MB) }
    return "{0:N2} GB" -f ($bytes / 1GB)
}

function Get-ArquivoLocal {
    param([string]$url, [string]$nomeDestino)
    $caminhoCompleto = Join-Path $Global:Config.CaminhoTemp $nomeDestino
    if (Test-Path $caminhoCompleto) {
        Write-Mensagem "Reutilizando arquivo local: $nomeDestino" "Info"
        return $caminhoCompleto
    }
    try {
        $tamanhoTotal = Get-TamanhoArquivoRemoto -url $url
        $wc = New-Object System.Net.WebClient
        $downloadConcluido = $false
        $erroDownload = $null
        Write-Host "Baixando: $nomeDestino" -ForegroundColor Gray
        $eventoProgresso = Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -Action {
            $script:percentualAtual = $Event.SourceEventArgs.ProgressPercentage
            $script:bytesRecebidos = $Event.SourceEventArgs.BytesReceived
            $script:bytesTotais = $Event.SourceEventArgs.TotalBytesToReceive
        }
        $eventoFim = Register-ObjectEvent -InputObject $wc -EventName DownloadFileCompleted -Action {
            $script:downloadConcluido = $true
            $script:erroDownload = $Event.SourceEventArgs.Error
        }
        $script:percentualAtual = 0
        $script:bytesRecebidos = 0
        $script:bytesTotais = $tamanhoTotal
        $script:downloadConcluido = $false
        $script:erroDownload = $null
        $wc.DownloadFileAsync($url, $caminhoCompleto)
        while (-not $script:downloadConcluido) {
            $larguraBarra = 24
            $percentual = [math]::Max(0, [math]::Min(100, $script:percentualAtual))
            $preenchido = [math]::Floor(($percentual / 100) * $larguraBarra)
            $barra = ("=" * $preenchido).PadRight($larguraBarra, ' ')
            $recebidoFormatado = Format-TamanhoArquivo -bytes $script:bytesRecebidos
            $totalFormatado = if ($script:bytesTotais -gt 0) { Format-TamanhoArquivo -bytes $script:bytesTotais } else { "?" }
            Write-Host (("`r[{0}] {1,3}% ({2} / {3})") -f $barra, $percentual, $recebidoFormatado, $totalFormatado) -NoNewline -ForegroundColor Cyan
            Start-Sleep -Milliseconds 250
        }
        Write-Host ""
        Unregister-Event -SourceIdentifier $eventoProgresso.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $eventoFim.Name -ErrorAction SilentlyContinue
        Remove-Job -Id $eventoProgresso.Id -Force -ErrorAction SilentlyContinue
        Remove-Job -Id $eventoFim.Id -Force -ErrorAction SilentlyContinue
        $wc.Dispose()
        if ($script:erroDownload) { throw $script:erroDownload }
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
        }
        catch {
            Write-Mensagem "Falha ao criar porta IP: $($_.Exception.Message)" "Erro"
            return $false
        }
    }
    return $true
}

function Test-RedeImpressora {
    param([Parameter(Mandatory=$true)][string]$enderecoIP)
    $meuIP = Get-MeuIPv4
    if (-not $meuIP) { return $true }
    $pingOk = $timeoutMs = $Global:Config.TempoPing * 1000
    $pingOk = Test-Connection -ComputerName $enderecoIP -Count 1 -Quiet -Timeout 
    $timeoutMs -ErrorAction SilentlyContinue
    if ($pingOk) { return $true }
    $minhaRede = ($meuIP -split '\.')[0..2] -join '.'
    $redeImpressora = ($enderecoIP -split '\.')[0..2] -join '.'
    Write-Host ""
    if ($minhaRede -eq $redeImpressora) {
        Write-Mensagem "Impressora nao responde ao ping!" "Aviso"
        Write-Host "  Verifique se esta ligada e conectada a rede"
    }
    else {
        Write-Mensagem "IP em rede diferente detectado!" "Aviso"
        Write-Host "  Seu computador: $meuIP"
        Write-Host "  IP digitado:    $enderecoIP"
    }
    $continuar = Read-OpcaoValidada "`nContinuar mesmo assim? [S/N]" @("S","s","N","n")
    Write-Host ""
    return ($continuar -ieq "S")
}

function Test-DriverExistente {
    param([Parameter(Mandatory=$true)][string]$filtroDriver)
    $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$filtroDriver*" }
    $driverNativo = $drivers | Where-Object {
        $_.Name -notlike "*PCL*" -and $_.Name -notlike "* PS" -and $_.Name -notlike "*Universal Print Driver*"
    } | Select-Object -First 1
    if ($driverNativo) { return @{ Encontrado = $true; Driver = $driverNativo.Name; Tipo = "Nativo" } }
    $driverVariacao = $drivers | Select-Object -First 1
    if ($driverVariacao) { return @{ Encontrado = $true; Driver = $driverVariacao.Name; Tipo = "Variacao" } }
    return @{ Encontrado = $false; Driver = $null; Tipo = $null }
}

function Test-DriverScanExistente {
    param([Parameter(Mandatory=$true)][string]$nomeModelo)
    $programas = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -like "*$nomeModelo*" -and $_.DisplayName -like "*Scan*" }
    return [bool]$programas
}

function Get-Caminho7Zip {
    $candidatos = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    foreach ($caminho in $candidatos) {
        if ($caminho -and (Test-Path $caminho)) { return $caminho }
    }
    $comando = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($comando) { return $comando.Source }
    return $null
}

function Ensure-7Zip {
    $caminho7Zip = Get-Caminho7Zip
    if ($caminho7Zip) { return @{ Caminho = $caminho7Zip; InstaladoPeloScript = $false } }
    Show-Titulo -titulo "INSTALANDO 7-ZIP" -cor "Yellow"
    $instalador = Get-ArquivoLocal -url $Global:Config.Url7Zip -nomeDestino "7zip_instalador.exe"
    if (-not $instalador) { return @{ Caminho = $null; InstaladoPeloScript = $false } }
    Write-Host "Instalando 7-Zip..." -ForegroundColor Gray
    Start-Process $instalador -ArgumentList "/S" -Wait -NoNewWindow
    Start-Sleep -Seconds 2
    $caminho7Zip = Get-Caminho7Zip
    return @{ Caminho = $caminho7Zip; InstaladoPeloScript = [bool]$caminho7Zip }
}

function Remove-7ZipIfNeeded {
    param([bool]$instaladoPeloScript)
    if (-not $instaladoPeloScript) { return }
    try {
        $chaves = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $app = Get-ItemProperty $chaves -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*7-Zip*" } | Select-Object -First 1
        if ($app -and $app.UninstallString) {
            Write-Host "Removendo 7-Zip instalado temporariamente..." -ForegroundColor Gray
            if ($app.UninstallString -match "msiexec") {
                Start-Process "cmd.exe" -ArgumentList "/c $($app.UninstallString) /qn" -Wait -NoNewWindow
            }
            else {
                Start-Process "cmd.exe" -ArgumentList "/c `"$($app.UninstallString) /S`"" -Wait -NoNewWindow
            }
        }
    }
    catch {
        Write-Mensagem "Nao foi possivel remover o 7-Zip temporario" "Aviso"
    }
}

function Set-FilaImpressora {
    param(
        [string]$nomeImpressora,
        [string]$enderecoIP,
        [string]$filtroDriver,
        [string]$driverPreferencial = "",
        [bool]$aceitaUniversal = $true
    )
    New-PortaIP -enderecoIP $enderecoIP | Out-Null
    Write-Host "Configurando fila de impressao..." -ForegroundColor Gray
    $filas = Get-Printer -ErrorAction SilentlyContinue
    $filaEspecifica = $filas | Where-Object {
        ($_.Name -like "*$filtroDriver*" -or $_.DriverName -like "*$filtroDriver*") -and
        $_.DriverName -notlike "*PCL*" -and $_.DriverName -notlike "* PS" -and $_.DriverName -notlike "*Universal Print Driver*"
    } | Select-Object -First 1
    $filaVariacao = if (-not $filaEspecifica) {
        $filas | Where-Object { $_.Name -like "*$filtroDriver*" -or $_.DriverName -like "*$filtroDriver*" } | Select-Object -First 1
    }
    $filaUniversal = $filas | Where-Object {
        $_.Name -like "*Samsung Universal Print Driver*" -or $_.DriverName -like "*Samsung Universal Print Driver*"
    } | Select-Object -First 1
    try {
        if ($filaEspecifica) {
            if ($driverPreferencial) { Set-Printer -Name $filaEspecifica.Name -DriverName $driverPreferencial -PortName $enderecoIP -ErrorAction Stop }
            else { Set-Printer -Name $filaEspecifica.Name -PortName $enderecoIP -ErrorAction Stop }
            if ($filaEspecifica.Name -ne $nomeImpressora) { Rename-Printer -Name $filaEspecifica.Name -NewName $nomeImpressora -ErrorAction Stop }
            return $true
        }
        elseif ($filaVariacao) {
            if ($driverPreferencial) { Set-Printer -Name $filaVariacao.Name -DriverName $driverPreferencial -PortName $enderecoIP -ErrorAction Stop }
            else { Set-Printer -Name $filaVariacao.Name -PortName $enderecoIP -ErrorAction Stop }
            if ($filaVariacao.Name -ne $nomeImpressora) { Rename-Printer -Name $filaVariacao.Name -NewName $nomeImpressora -ErrorAction Stop }
            return $true
        }
        elseif ($aceitaUniversal -and $filaUniversal) {
            Set-Printer -Name $filaUniversal.Name -PortName $enderecoIP -ErrorAction Stop
            if ($filaUniversal.Name -ne $nomeImpressora) { Rename-Printer -Name $filaUniversal.Name -NewName $nomeImpressora -ErrorAction Stop }
            return $true
        }
        elseif ($driverPreferencial) {
            Add-Printer -Name $nomeImpressora -DriverName $driverPreferencial -PortName $enderecoIP -ErrorAction Stop
            return $true
        }
        else {
            Add-Printer -Name $nomeImpressora -DriverName $filtroDriver -PortName $enderecoIP -ErrorAction Stop
            return $true
        }
    }
    catch {
        Write-Mensagem "Falha ao configurar impressora: $($_.Exception.Message)" "Erro"
        Write-Host "`nDrivers disponiveis no sistema:"
        Get-PrinterDriver -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*$fabricante*" -or $_.Name -like "*Samsung*" } |
            ForEach-Object { Write-Host "  - $($_.Name)" }
        Read-Host "`nPressione ENTER para continuar"
        return $false
    }
}

function Install-ComponenteExe {
    param(
        [string]$nomeComponente,
        [string]$nomePrograma,
        [string]$url,
        [string]$nomeArquivoDestino,
        [int]$timeout = 60
    )
    if (Test-ProgramaInstalado $nomePrograma) {
        Write-Mensagem "$nomeComponente ja presente no sistema" "Sucesso"
        return $true
    }
    $arquivo = Get-ArquivoLocal -url $url -nomeDestino $nomeArquivoDestino
    if (-not $arquivo) { return $false }
    Write-Host "Instalando $nomeComponente (timeout: ${timeout}s)..." -ForegroundColor Gray
    $processo = Start-Process $arquivo -ArgumentList "/S" -PassThru -NoNewWindow
    $tempoDecorrido = 0
    while (-not $processo.HasExited -and $tempoDecorrido -lt $timeout) {
        Start-Sleep -Seconds 2
        $tempoDecorrido += 2
        if (Test-ProgramaInstalado $nomePrograma) {
            Write-Mensagem "$nomeComponente instalado com sucesso!" "Sucesso"
            if (-not $processo.HasExited) { Stop-Process -Id $processo.Id -Force -ErrorAction SilentlyContinue }
            return $true
        }
    }
    if (Test-ProgramaInstalado $nomePrograma) {
        Write-Mensagem "$nomeComponente instalado com sucesso!" "Sucesso"
        return $true
    }
    Write-Mensagem "Instalacao de $nomeComponente pode nao ter sido concluida" "Aviso"
    return $false
}

function Install-DriverSPL {
    param([string]$urlDriver,[string]$nomeModelo,[string]$filtroDriver,[string]$nomeImpressora,[string]$enderecoIP)
    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    if ($statusDriver.Encontrado -and $statusDriver.Tipo -eq "Nativo") {
        Write-Host ""
        Write-Mensagem "Driver '$($statusDriver.Driver)' ja presente no sistema" "Sucesso"
        Write-Host "Configurando impressora..." -ForegroundColor Gray
        return (Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -driverPreferencial $statusDriver.Driver -aceitaUniversal $false)
    }
    Write-Host ""
    $nomeArquivo = "driver_print_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    if (-not $arquivoDriver) { return $false }
    Write-Host "Instalando driver (aguarde)..." -ForegroundColor Gray
    Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
    Start-Sleep -Seconds $Global:Config.TempoEspera
    $sucesso = Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -aceitaUniversal $false
    if ($sucesso) { Write-Mensagem "Impressora configurada com sucesso!" "Sucesso" }
    return $sucesso
}

function Install-DriverUPD {
    param([string]$urlDriver,[string]$nomeModelo,[string]$filtroDriver,[string]$nomeImpressora,[string]$enderecoIP)
    $statusDriver = Test-DriverExistente -filtroDriver $filtroDriver
    if ($statusDriver.Encontrado -and $statusDriver.Tipo -eq "Nativo") {
        Write-Host ""
        Write-Mensagem "Driver '$($statusDriver.Driver)' ja presente no sistema" "Sucesso"
        Write-Host "Configurando impressora..." -ForegroundColor Gray
        return (Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -driverPreferencial $statusDriver.Driver -aceitaUniversal $true)
    }
    Write-Host ""
    $nomeArquivo = "driver_UPD_" + ($nomeModelo -replace '\s+', '_') + ".exe"
    $arquivoDriver = Get-ArquivoLocal -url $urlDriver -nomeDestino $nomeArquivo
    if (-not $arquivoDriver) { return $false }
    $pastaExtracao = Join-Path $Global:Config.CaminhoTemp (("UPD_Extract_") + [guid]::NewGuid().ToString().Substring(0,8))
    New-Item -Path $pastaExtracao -ItemType Directory -Force | Out-Null
    $controle7Zip = Ensure-7Zip
    $caminho7Zip = $controle7Zip.Caminho
    $instaladoPeloScript = $controle7Zip.InstaladoPeloScript
    if ($caminho7Zip) {
        Write-Host "Extraindo pacote de drivers com 7-Zip..." -ForegroundColor Gray
        & $caminho7Zip x "$arquivoDriver" "-o$pastaExtracao" -y | Out-Null
    }
    else {
        Write-Mensagem "7-Zip nao foi encontrado ou instalado. Usando metodo padrao." "Aviso"
        Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
        Start-Sleep -Seconds $Global:Config.TempoEspera
    }
    $driverEspecifico = $null
    if ($caminho7Zip) {
        $arquivosInf = Get-ChildItem -Path $pastaExtracao -Filter "*.inf" -Recurse -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -notlike "*autorun*" -and $_.Name -notlike "*setup*" }
        $infEspecifico = $arquivosInf | Where-Object {
            $_.FullName -match [regex]::Escape(($filtroDriver -replace 'Samsung ', '')) -or $_.FullName -match "M408|6260|UPD"
        } | Select-Object -First 1
        if (-not $infEspecifico) { $infEspecifico = $arquivosInf | Select-Object -First 1 }
        if ($infEspecifico) {
            Write-Host "Instalando driver via pnputil..." -ForegroundColor Gray
            & pnputil.exe /add-driver "$($infEspecifico.FullName)" /install 2>&1 | Out-Null
            Start-Sleep -Seconds $Global:Config.TempoEspera
            $driverEspecifico = Get-PrinterDriver -ErrorAction SilentlyContinue |
                               Where-Object {
                                   $_.Name -like "*$filtroDriver*" -and $_.Name -notlike "*PCL*" -and $_.Name -notlike "* PS" -and $_.Name -notlike "*Universal Print Driver*"
                               } | Select-Object -First 1
        }
        else {
            Write-Mensagem "Nenhum arquivo INF valido foi encontrado. Usando instalacao padrao." "Aviso"
            Start-Process $arquivoDriver -ArgumentList "/S" -Wait -NoNewWindow
            Start-Sleep -Seconds $Global:Config.TempoEspera
        }
    }
    Remove-Item $pastaExtracao -Recurse -Force -ErrorAction SilentlyContinue
    Remove-7ZipIfNeeded -instaladoPeloScript $instaladoPeloScript
    if (-not $driverEspecifico) {
        $driverEspecifico = Get-PrinterDriver -ErrorAction SilentlyContinue |
                           Where-Object {
                               $_.Name -like "*$filtroDriver*" -and $_.Name -notlike "*PCL*" -and $_.Name -notlike "* PS" -and $_.Name -notlike "*Universal Print Driver*"
                           } | Select-Object -First 1
    }
    $driverPreferencial = if ($driverEspecifico) { $driverEspecifico.Name } else { "" }
    $sucesso = Set-FilaImpressora -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -filtroDriver $filtroDriver -driverPreferencial $driverPreferencial -aceitaUniversal $true
    if ($sucesso) {
        if ($driverEspecifico) { Write-Mensagem "Impressora configurada com driver especifico!" "Sucesso" }
        else { Write-Mensagem "Impressora configurada com driver universal" "Aviso" }
    }
    return $sucesso
}

function Remove-ImpressoraExistente {
    param([Parameter(Mandatory=$true)][ValidateSet("Nome","IP")][string]$tipoBusca,[Parameter(Mandatory=$true)][string]$valor)
    try {
        if ($tipoBusca -eq "Nome") { $impressora = Get-Printer -Name $valor -ErrorAction Stop }
        else { $impressora = Get-Printer -ErrorAction Stop | Where-Object { $_.PortName -eq $valor } | Select-Object -First 1 }
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
        $_.Name -ne $nomeConfigurado -and (
            $_.Name -eq $filtroDriver -or
            $_.Name -match "^$([regex]::Escape($filtroDriver))( PS| PCL[0-9].*)?( \((Copia|Copiar|Copy) \d+\))?$" -or
            ($_.Name -like "*Samsung Universal Print Driver*" -and $_.DriverName -like "*Samsung Universal*")
        ) -and $_.Name -notlike "*Fax*"
    }
    if ($duplicatas) {
        Write-Host "Removendo filas duplicadas..." -ForegroundColor Gray
        foreach ($fila in $duplicatas) {
            try { Remove-Printer -Name $fila.Name -Confirm:$false -ErrorAction Stop }
            catch { }
        }
    }
}

function Show-ResumoInstalacao {
    param([string]$modelo,[string]$nomeImpressora,[string]$enderecoIP,[string]$driver,[string[]]$componentes)
    Write-Host ""
    Write-Linha -cor "Cyan"
    Write-Host "       RESUMO DA INSTALACAO" -ForegroundColor Cyan
    Write-Linha -cor "Cyan"
    Write-Host "Modelo:      $modelo"
    if (-not [string]::IsNullOrWhiteSpace($fabricante)) { Write-Host "Fabricante:  $fabricante" }
    Write-Host "Nome:        $nomeImpressora"
    Write-Host "IP:          $enderecoIP"
    Write-Host "Driver:      $driver"
    Write-Host "Componentes: $($componentes -join ', ')"
    Write-Host "Status:      " -NoNewline
    Write-Host "Concluido com sucesso!" -ForegroundColor Green
    Write-Linha -cor "Cyan"
    Write-Host ""
}

$componentesInstalados = @()
if ($instalarPrint) { $componentesInstalados += "Print" }
if ($instalarScan)  { $componentesInstalados += "Scan" }
if ($instalarEPM)   { $componentesInstalados += "EPM" }
if ($instalarEDC)   { $componentesInstalados += "EDC" }
$totalEtapas = $componentesInstalados.Count
$etapaAtual = 1
Show-Titulo -titulo "INSTALACAO: $modelo" -cor "Cyan"
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
            if ($opcao -eq "3") { Write-Mensagem "Instalacao cancelada" "Info"; return }
            elseif ($opcao -eq "2") {
                if (Remove-ImpressoraExistente -tipoBusca "Nome" -valor $nomeImpressora) { break }
                else { Write-Host "Tente novamente com outro nome.`n" -ForegroundColor Yellow }
            }
        }
        else { break }
    } while ($true)
    do {
        $enderecoIP = Read-Host "- Endereco IP"
        Write-Host ""
        if ($enderecoIP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
            Write-Mensagem "IP invalido! Use formato XXX.XXX.XXX.XXX" "Erro"
            Write-Host ""
            continue
        }
        $impressoraMesmoIP = Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.PortName -eq $enderecoIP } | Select-Object -First 1
        if ($impressoraMesmoIP) {
            Write-Mensagem "Uma impressora com esse mesmo IP foi detectada no sistema!" "Aviso"
            Write-Host "  Nome:   $($impressoraMesmoIP.Name)"
            Write-Host "  Driver: $($impressoraMesmoIP.DriverName)`n"
            $opcao = Read-OpcaoValidada "[1] Digitar outro IP  [2] Apagar e prosseguir  [3] Cancelar" @("1","2","3")
            Write-Host ""
            if ($opcao -eq "3") { Write-Mensagem "Instalacao cancelada" "Info"; return }
            elseif ($opcao -eq "2") {
                if (Remove-ImpressoraExistente -tipoBusca "IP" -valor $enderecoIP) { break }
                else { Write-Host "Tente novamente com outro IP.`n" -ForegroundColor Yellow }
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
        Install-DriverUPD -urlDriver $urlPrint -nomeModelo $modelo -filtroDriver $filtroDriverWindows -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP
    } else {
        Install-DriverSPL -urlDriver $urlPrint -nomeModelo $modelo -filtroDriver $filtroDriverWindows -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP
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
if ($instalarEPM) {
    Write-Host "[$etapaAtual/$totalEtapas] EASY PRINTER MANAGER" -ForegroundColor Yellow
    Write-Host ""
    Install-ComponenteExe -nomeComponente "Easy Printer Manager" -nomePrograma "Easy Printer Manager" -url $Global:Config.UrlEPM -nomeArquivoDestino "EPM_Universal.exe" | Out-Null
    $etapaAtual++
    Write-Host ""
}
if ($instalarEDC) {
    Write-Host "[$etapaAtual/$totalEtapas] EASY DOCUMENT CREATOR" -ForegroundColor Yellow
    Write-Host ""
    Install-ComponenteExe -nomeComponente "Easy Document Creator" -nomePrograma "Easy Document Creator" -url $Global:Config.UrlEDC -nomeArquivoDestino "EDC_Universal.exe" | Out-Null
    Write-Host ""
}
if ($instalarPrint -and $instalacaoSucesso) {
    Show-ResumoInstalacao -modelo $modelo -nomeImpressora $nomeImpressora -enderecoIP $enderecoIP -driver $driverInstalado -componentes $componentesInstalados
    $imprimirTeste = Read-OpcaoValidada "Deseja imprimir uma pagina de teste? [S/N]" @("S","s","N","n")
    if ($imprimirTeste -ieq "S") {
        try {
            Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /k /n `"$nomeImpressora`"" -NoNewWindow -Wait
            Write-Host ""
            Write-Mensagem "Pagina de teste enviada!" "Sucesso"
        }
        catch {
            Write-Mensagem "Falha ao enviar pagina de teste" "Erro"
        }
    }
}
elseif ($instalarPrint -and -not $instalacaoSucesso) {
    Write-Host ""
    Write-Linha -cor "Red"
    Write-Host "     FALHA NA INSTALACAO" -ForegroundColor Red
    Write-Linha -cor "Red"
    Write-Host "A instalacao da impressora nao foi concluida."
    Write-Host "Verifique os erros acima."
    Write-Linha -cor "Red"
    Write-Host ""
}
Write-Host ""
Start-Sleep -Seconds 2

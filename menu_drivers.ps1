# Verifica se o script esta rodando como Administrador
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Este script precisa ser executado como ADMINISTRADOR."
    Write-Host "Por favor, abra o PowerShell como Administrador antes de rodar o comando." -ForegroundColor Yellow
    Pause
    exit
}

function Mostrar-MenuPrincipal {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "     SUPORTE TECH3 - DRIVERS SAMSUNG    " -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "1) Samsung ProXpress M4020"
    Write-Host "2) Samsung ProXpress M4070"
    Write-Host "3) Samsung ProXpress M4080"
    Write-Host "q) Sair"
    Write-Host "=========================================="
}

function Baixar-Arquivo($url, $nome, $modeloPasta) {
    $diretorioBase = "$env:USERPROFILE\Downloads\$modeloPasta"
    if (-not (Test-Path $diretorioBase)) {
        New-Item -Path $diretorioBase -ItemType Directory | Out-Null
    }
    $destinoFull = Join-Path $diretorioBase $nome
    Write-Host "Baixando $nome..." -ForegroundColor Green
    try {
        Invoke-WebRequest -Uri $url -OutFile $destinoFull -ErrorAction Stop
        return $destinoFull
    } catch {
        Write-Host "Erro ao baixar $nome. Verifique a conexao." -ForegroundColor Red
        return $null
    }
}

do {
    Mostrar-MenuPrincipal
    $opcao = Read-Host "Escolha o modelo"

    switch ($opcao) {
        "2" { # Samsung M4070
            do {
                Clear-Host
                $mod = "Samsung_M4070"
                $urlPrint = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M3370FD/M337x_387x_407x_Series_WIN_SPL_PCL_V3.13.13.00.01_CDV1.38.exe"
                $urlScan  = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M3370FD/M337x_387x_407x_Series_WIN_Scanner_V3.31.39.11_CDV1.38.exe"
                $urlEPM   = "https://ftp.hp.com/pub/softlib/software13/printers/SS/Common_SW/WIN_EPM_V2.00.01.36.exe"
                $urlEDC   = "https://ftp.hp.com/pub/softlib/software13/printers/SS/SL-M5270LX/WIN_EDC_V2.02.61.exe"
                $urlInstalador = "https://raw.githubusercontent.com/suportetech3brasil/arquivos/refs/heads/main/instalar4070.ps1"

                Write-Host "--- MENU $mod ---" -ForegroundColor Yellow
                Write-Host "1) Driver de Impressao"
                Write-Host "2) Driver de Digitalizacao"
                Write-Host "3) Easy Printer Manager"
                Write-Host "4) Easy Document Creator"
                Write-Host "5) Baixar tudo e instalar"
                Write-Host "v) Voltar ao Menu Principal"
                
                $sub = Read-Host "Opcao"

                if ($sub -eq "5") {
                    Write-Host "Iniciando pacote completo..." -ForegroundColor Cyan
                    Baixar-Arquivo $urlPrint "Driver_M4070_Print.exe" $mod
                    Baixar-Arquivo $urlScan  "Driver_M4070_Scan.exe" $mod
                    Baixar-Arquivo $urlEPM   "EasyPrinterManager.exe" $mod
                    Baixar-Arquivo $urlEDC   "EasyDocumentCreator.exe" $mod
                    $scriptLocal = Baixar-Arquivo $urlInstalador "instalar4070.ps1" $mod
                    
                    if ($scriptLocal) {
                        Write-Host "Iniciando instalacao automatizada..." -ForegroundColor Yellow
                        Set-Location (Split-Path $scriptLocal)
                        & $scriptLocal
                        Write-Host "Processo concluido!" -ForegroundColor Green
                    }
                    Pause
                }
                elseif ($sub -eq "1") { Baixar-Arquivo $urlPrint "Driver_M4070_Print.exe" $mod; Pause }
                elseif ($sub -eq "2") { Baixar-Arquivo $urlScan  "Driver_M4070_Scan.exe" $mod; Pause }
                elseif ($sub -eq "3") { Baixar-Arquivo $urlEPM   "EasyPrinterManager.exe" $mod; Pause }
                elseif ($sub -eq "4") { Baixar-Arquivo $urlEDC   "EasyDocumentCreator.exe" $mod; Pause }

            } while ($sub -ne "v")
        }
    }
} while ($opcao -ne "q")


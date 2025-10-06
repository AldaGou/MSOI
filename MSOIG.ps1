#requires -Version 5.1
<#
    Office LTSC Installer GUI
    - Descarga ODT (URL evergreen + fallback)
    - Genera configuration.xml
    - Ejecuta instalación
    - Muestra estado/detección y log
#>

#region Elevación a Administrador
function Ensure-Elevated {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "Reiniciando como Administrador..."
        if ($PSCommandPath) {
            Start-Process -FilePath "powershell.exe" `
                -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
                -Verb RunAs
        } else {
            Start-Process -FilePath "powershell.exe" -Verb RunAs
        }
        exit
    }
}
Ensure-Elevated
#endregion

#region Utilidades
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
$global:LogPath = Join-Path $env:TEMP ("OfficeLTSC-GUI_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $global:LogPath -Append | Out-Null

# Forzar TLS moderno
try {
    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.SecurityProtocolType]::Tls12 -bor `
        [Net.SecurityProtocolType]::Tls13
} catch {}

function Write-Log($msg) {
    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$stamp  $msg" | Add-Content -Path $global:LogPath
}

function Test-Internet {
    try {
        $r = Invoke-WebRequest -Uri "https://www.microsoft.com" -Method Head -TimeoutSec 10
        return ($r.StatusCode -ge 200 -and $r.StatusCode -lt 400)
    } catch { return $false }
}

function Get-OfficeState {
    $result = [ordered]@{
        ClickToRun = $null
        MSI        = $false
        Arch       = $null
        Channel    = $null
        ProductIDs = @()
    }
    try {
        $c2r = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
        $result.ClickToRun = $true
        $result.Arch       = $c2r.Platform
        $result.Channel    = $c2r.UpdateChannel
        if ($c2r.ProductReleaseIds) {
            $result.ProductIDs = $c2r.ProductReleaseIds -split '\s*,\s*'
        }
    } catch {
        $result.ClickToRun = $false
    }

    # Detección muy básica de MSI (queda a criterio confirmar/mostrar)
    try {
        $msiKeys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                   "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        foreach ($k in $msiKeys) {
            Get-ChildItem $k -ErrorAction SilentlyContinue | ForEach-Object {
                $dn = (Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue).DisplayName
                if ($dn -and $dn -match 'Microsoft Office(?! Click-to-Run)|Visio|Project') {
                    $result.MSI = $true
                }
            }
        }
    } catch {}
    return $result
}

function Get-ProductMap {
    # Mapa de IDs por versión
    return @{
        "LTSC 2024" = @{
            Channel   = "PerpetualVL2024"
            Office    = "ProPlus2024Volume"
            Visio     = "VisioPro2024Volume"
            Project   = "ProjectPro2024Volume"
        }
        "LTSC 2021" = @{
            Channel   = "PerpetualVL2021"
            Office    = "ProPlus2021Volume"
            Visio     = "VisioPro2021Volume"
            Project   = "ProjectPro2021Volume"
        }
        "LTSC 2019" = @{
            Channel   = "PerpetualVL2019"
            Office    = "ProPlus2019Volume"
            Visio     = "VisioPro2019Volume"
            Project   = "ProjectPro2019Volume"
        }
    }
}

function New-OfficeConfigXml {
    param(
        [ValidateSet("32","64")] [string]$Edition = "64",
        [Parameter(Mandatory)]   [string]$Channel,
        [Parameter(Mandatory)]   [string]$ProductId,
        [Parameter(Mandatory)]   [string]$Language,
        [string[]]               $IncludeApps = @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote"),
        [switch]                 $IncludeVisio,
        [switch]                 $IncludeProject,
        [string]                 $VisioId,
        [string]                 $ProjectId,
        [switch]                 $RemoveMSI,
        [ValidateSet("None","Basic","Full")] [string]$DisplayLevel = "Full"
    )
    # Apps válidas conocidas para ExcludeApp
    $knownApps = @("Access","Excel","OneNote","Outlook","PowerPoint","Publisher","Teams","Word","OneDrive","Groove","Lync")

    $allOfficeApps = @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote")
    $exclude = $allOfficeApps | Where-Object { $_ -notin $IncludeApps }

    $xml = New-Object System.Text.StringBuilder
    [void]$xml.AppendLine('<Configuration>')
    [void]$xml.AppendLine("  <Add OfficeClientEdition=""$Edition"" Channel=""$Channel"">")
    [void]$xml.AppendLine("    <Product ID=""$ProductId"">")
    [void]$xml.AppendLine("      <Language ID=""$Language"" />")

    foreach($app in $exclude) {
        if ($knownApps -contains $app) {
            [void]$xml.AppendLine("      <ExcludeApp ID=""$app"" />")
        }
    }

    # Exclusiones adicionales típicas (opcionales)
    foreach($extra in @("Teams","OneDrive","Groove","Lync")) {
        if ($extra -notin $IncludeApps -and $knownApps -contains $extra) {
            [void]$xml.AppendLine("      <ExcludeApp ID=""$extra"" />")
        }
    }

    [void]$xml.AppendLine("    </Product>")

    if ($IncludeProject -and $ProjectId) {
        [void]$xml.AppendLine("    <Product ID=""$ProjectId"">")
        [void]$xml.AppendLine("      <Language ID=""$Language"" />")
        [void]$xml.AppendLine("    </Product>")
    }
    if ($IncludeVisio -and $VisioId) {
        [void]$xml.AppendLine("    <Product ID=""$VisioId"">")
        [void]$xml.AppendLine("      <Language ID=""$Language"" />")
        [void]$xml.AppendLine("    </Product>")
    }

    [void]$xml.AppendLine("  </Add>")
    if ($RemoveMSI) {
        [void]$xml.AppendLine("  <RemoveMSI />")
    }
    [void]$xml.AppendLine("  <Display Level=""$DisplayLevel"" AcceptEULA=""TRUE"" />")
    [void]$xml.AppendLine('</Configuration>')
    return $xml.ToString()
}

function Get-ODT {
    param(
        [Parameter(Mandatory)] [string]$ExtractTo
    )
    $temp     = $env:TEMP
    $exePath  = Join-Path $temp "OfficeDeploymentTool.exe"
    $primary  = "https://aka.ms/ODT" # evergreen
    $fallback = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"

    New-Item -ItemType Directory -Path $ExtractTo -Force | Out-Null

    Write-Log "Descargando ODT..."
    $downloaded = $false
    foreach ($url in @($primary,$fallback)) {
        try {
            Invoke-WebRequest -Uri $url -OutFile $exePath -TimeoutSec 180
            $downloaded = $true
            Write-Log "ODT descargado desde: $url"
            break
        } catch {
            Write-Log "Fallo descarga ODT desde $url : $($_.Exception.Message)"
        }
    }
    if (-not $downloaded) {
        throw "No se pudo descargar ODT. Revisa tu conexión o proxy."
    }

    Write-Log "Extrayendo ODT a $ExtractTo"
    Start-Process -FilePath $exePath -ArgumentList "/quiet /extract:`"$ExtractTo`"" -Wait
    $setup = Join-Path $ExtractTo "setup.exe"
    if (-not (Test-Path $setup)) {
        throw "setup.exe no se encontró tras la extracción."
    }
    return $setup
}

function Start-ODTInstall {
    param(
        [Parameter(Mandatory)] [string]$SetupExe,
        [Parameter(Mandatory)] [string]$ConfigXml
    )
    Write-Log "Iniciando instalación: `"$SetupExe`" /configure `"$ConfigXml`""
    $p = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$ConfigXml`"" -PassThru
    $p.WaitForExit()
    return $p.ExitCode
}
#endregion

#region UI (WinForms)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form         = New-Object Windows.Forms.Form
$form.Text    = "Office LTSC Installer"
$form.Size    = New-Object Drawing.Size(720, 640)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $false

$font = New-Object Drawing.Font("Segoe UI",10)

# Controles
$lblVersion = New-Object Windows.Forms.Label
$lblVersion.Text = "Versión:"
$lblVersion.Location = '20,20'
$lblVersion.AutoSize = $true
$lblVersion.Font = $font

$cmbVersion = New-

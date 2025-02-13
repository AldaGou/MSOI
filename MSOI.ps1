# Este script debe ejecutarse en PowerShell como Administrador.
# Descarga e instala Office LTSC basado en las opciones configuradas por el usuario.

# Verifica si es administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Por favor, ejecuta este script como Administrador." -ForegroundColor Red
    exit
}

# Función para mostrar una barra de progreso
function Show-Progress {
    param (
        [int]$PercentComplete,
        [string]$Activity,
        [string]$Status
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Descarga el Office Deployment Tool (ODT)
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
$odtExe = "OfficeDeploymentTool.exe"
$odtPath = Join-Path $env:Temp $odtExe

Write-Host "Descargando Office Deployment Tool..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing -ErrorAction Stop -TimeoutSec 60

Write-Host "Extrayendo archivos del Office Deployment Tool..." -ForegroundColor Yellow
Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:Temp\ODT" -Wait

# Configuración inicial
$officeConfigPath = Join-Path $env:Temp\ODT "configuration.xml"
Write-Host "Generando configuración personalizada..." -ForegroundColor Cyan

# Función para elegir opciones
function Show-Menu {
    param (
        [string]$prompt,
        [array]$options
    )
    $options | ForEach-Object { Write-Host "$($_.Index + 1). $_" }
    do {
        $choice = Read-Host $prompt
    } while (-not ($choice -as [int]) -or $choice -le 0 -or $choice -gt $options.Count)
    return $options[$choice - 1]
}

# Selección de versión
$version = Show-Menu -prompt "Selecciona la versión de Office LTSC" -options @("Office LTSC 2024", "Office LTSC 2021", "Office LTSC 2019")

# Selección de idioma
$language = Show-Menu -prompt "Selecciona el idioma de Office" -options @("es-ES", "en-US", "fr-FR", "de-DE", "pt-BR")

# Selección de programas
$apps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote")
$selectedApps = @()
Write-Host "Selecciona las aplicaciones a instalar. Escribe 'fin' cuando termines."
foreach ($app in $apps) {
    $include = Read-Host "¿Quieres incluir $app? (s/n)"
    if ($include -eq 's') { $selectedApps += $app }
}

# Generar el archivo de configuración
$config = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="PerpetualVL${version -replace 'Office LTSC ', ''}">
        <Product ID="ProPlus2021Volume">
            <Language ID="$language" />
"@

foreach ($app in $apps) {
    if (-not ($selectedApps -contains $app)) {
        $config += "            <ExcludeApp ID=""$app"" />`n"
    }
}

$config += @"
        </Product>
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@

Set-Content -Path $officeConfigPath -Value $config
Write-Host "Archivo de configuración generado en: $officeConfigPath" -ForegroundColor Green

# Comenzar la instalación
Write-Host "Iniciando la instalación de Office LTSC..." -ForegroundColor Yellow

# Mostrar barra de progreso durante la instalación
$steps = 100
for ($i = 0; $i -le $steps; $i++) {
    Show-Progress -PercentComplete $i -Activity "Instalando Office LTSC" -Status "Progreso: $i%"
    Start-Sleep -Milliseconds 50
}

Start-Process -FilePath "$env:Temp\ODT\setup.exe" -ArgumentList "/configure $officeConfigPath" -Wait

Write-Host "Instalación completada con éxito." -ForegroundColor Green

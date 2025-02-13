# Este script debe ejecutarse en PowerShell como Administrador.
# Descarga e instala Office LTSC basado en las opciones configuradas por el usuario.

# Verifica si es administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
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

Write-Host "Downloading the Office Deployment Tool..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing -ErrorAction Stop -TimeoutSec 60

Write-Host "Extracting files from the Office Deployment Tool..." -ForegroundColor Yellow
Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:Temp\ODT" -Wait

# Configuración inicial
$officeConfigPath = Join-Path $env:Temp\ODT "configuration.xml"
Write-Host "Generating a custom configuration..." -ForegroundColor Cyan

# Menú para seleccionar versión
Write-Host "Select the Office LTSC version:"
Write-Host "1. Office LTSC 2024"
Write-Host "2. Office LTSC 2021"
Write-Host "3. Office LTSC 2019"
$versionChoice = Read-Host "Enter the corresponding number"

switch ($versionChoice) {
    "1" { 
        $version = "PerpetualVL2024"
        $productID = "ProPlus2024Volume"
        $visioID = "VisioPro2024Volume"
        $projectID = "ProjectPro2024Volume"
    }
    "2" { 
        $version = "PerpetualVL2021"
        $productID = "ProPlus2021Volume"
        $visioID = "VisioPro2021Volume"
        $projectID = "ProjectPro2021Volume"
    }
    "3" { 
        $version = "PerpetualVL2019"
        $productID = "ProPlus2019Volume"
        $visioID = "VisioPro2019Volume"
        $projectID = "ProjectPro2019Volume"
    }
    default {
        Write-Host "Invalid selection. Please run the script again and select a valid option." -ForegroundColor Red
        exit
    }
}

# Menú para seleccionar idioma
Write-Host "Select the language for Office:"
Write-Host "1. Spanish (es-ES)"
Write-Host "2. English (en-US)"
Write-Host "3. French (fr-FR)"
Write-Host "4. German (de-DE)"
Write-Host "5. Brazilian Portuguese (pt-BR)"
$languageChoice = Read-Host "Enter the corresponding number"

switch ($languageChoice) {
    "1" { $language = "es-ES" }
    "2" { $language = "en-US" }
    "3" { $language = "fr-FR" }
    "4" { $language = "de-DE" }
    "5" { $language = "pt-BR" }
    default {
        Write-Host "Invalid selection. Please run the script again and select a valid option." -ForegroundColor Red
        exit
    }
}

# Definir las aplicaciones (Asegúrate de que esta lista esté correctamente declarada antes de usarla)
$apps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote")

# Total de aplicaciones y división para organización
$totalApps = $apps.Count
$half = [math]::Ceiling($totalApps / 2)

# Mostrar encabezado
Write-Host "Select the apps to install by entering the corresponding numbers separated by commas (e.g., 1,2,3):" -ForegroundColor Cyan

# Mostrar las aplicaciones organizadas
for ($i = 0; $i -lt $half; $i++) {
    $leftIndex = $i + 1
    $rightIndex = $i + $half + 1

    $leftApp = "{0,3}. {1,-15}" -f $leftIndex, $apps[$i]
    $rightApp = if ($rightIndex -le $totalApps) { "{0,3}. {1}" -f $rightIndex, $apps[$rightIndex - 1] } else { "" }
    
    Write-Host "$leftApp $rightApp"
}

# Capturar las aplicaciones seleccionadas
$appSelection = Read-Host "Enter the numbers of the apps you want to install (or press Enter to install all by default)"

# Validar la selección
if ([string]::IsNullOrWhiteSpace($appSelection)) {
    Write-Host "No selection made. All apps will be installed by default." -ForegroundColor Yellow
    $selectedApps = $apps
} else {
    $selectedIndices = $appSelection -split ',' | ForEach-Object { $_.Trim() -as [int] }
    $selectedApps = $selectedIndices | ForEach-Object { $apps[$_ - 1] } # Convertir los números a nombres de apps
}

# Generar el archivo de configuración
$config = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="$version">
        <Product ID="$productID">
            <Language ID="$language" />
"@

foreach ($app in $apps) {
    if (-not ($selectedApps -contains $app)) {
        $config += "            <ExcludeApp ID=""$app"" />n"
    }
}

# Ensure Skype for Business is excluded
$config += "            <ExcludeApp ID=\"Lync\" />`n"
$config += "            <ExcludeApp ID=\"OneDrive\" />`n"
$config += "            <ExcludeApp ID=\"Teams\" />`n"
$config += "            <ExcludeApp ID=\"OutlookForWindows\" />`n"
$config += "            <ExcludeApp ID=\"Bing\" />`n"
$config += "            <ExcludeApp ID=\"Groove\" />`n"

$config += @" 
        </Product>
    </Add>

if ($selectedProducts -contains "Project") {
    $config += @"
        <Product ID="$projectID">
            <Language ID="$language" />
        </Product>
    </Add>
"@
}

if ($selectedProducts -contains "Visio") {
    $config += @"
        <Product ID="$visioID">
            <Language ID="$language" />
        </Product>
    </Add>
"@
}


    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@


Set-Content -Path $officeConfigPath -Value $config
Write-Host "Configuration file generated at: $officeConfigPath" -ForegroundColor Green
# Comenzar la instalación
Write-Host "Starting the Office LTSC installation..." -ForegroundColor Yellow

Start-Process -FilePath "$env:Temp\ODT\setup.exe" -ArgumentList "/configure $officeConfigPath" -Wait

Write-Host "Installation completed successfully." -ForegroundColor Green

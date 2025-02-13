# This script must be run in PowerShell as Administrator.
# Downloads and installs Office LTSC based on user-configured options.

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Red
    exit
}

# Function to display a progress bar
function Show-Progress {
    param (
        [int]$PercentComplete,
        [string]$Activity,
        [string]$Status
    )
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
}

# Download the Office Deployment Tool (ODT)
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
$odtExe = "OfficeDeploymentTool.exe"
$odtPath = Join-Path $env:Temp $odtExe

Write-Host "Downloading Office Deployment Tool..." -ForegroundColor Yellow
Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing -ErrorAction Stop -TimeoutSec 60

Write-Host "Extracting Office Deployment Tool files..." -ForegroundColor Yellow
Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:Temp\ODT" -Wait

# Initial configuration
$officeConfigPath = Join-Path $env:Temp\ODT "configuration.xml"
Write-Host "Generating custom configuration..." -ForegroundColor Cyan

# Version selection menu
Write-Host "Select the version of Office LTSC:" -ForegroundColor Cyan
Write-Host "1. Office LTSC 2024"
Write-Host "2. Office LTSC 2021"
Write-Host "3. Office LTSC 2019"
$versionChoice = Read-Host "Enter the corresponding number"

switch ($versionChoice) {
    "1" { 
        $version = "PerpetualVL2024"
        $productID = "ProPlus2024Volume"
    }
    "2" { 
        $version = "PerpetualVL2021"
        $productID = "ProPlus2021Volume"
    }
    "3" { 
        $version = "PerpetualVL2019"
        $productID = "ProPlus2019Volume"
    }
    default {
        Write-Host "Invalid selection. Please run the script again and choose a valid option." -ForegroundColor Red
        exit
    }
}

# Language selection menu
Write-Host "Select the language for Office:" -ForegroundColor Cyan
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
        Write-Host "Invalid selection. Please run the script again and choose a valid option." -ForegroundColor Red
        exit
    }
}

# Program selection menu
$apps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote", "Skype", "OneDrive")
Write-Host "Select the applications to install by entering the corresponding numbers separated by commas (e.g., 1,2,3)." -ForegroundColor Cyan

for ($i = 0; $i -lt $apps.Count; $i++) {
    Write-Host "$($i + 1). $($apps[$i])"
}

$appInput = Read-Host "Enter the numbers of the applications you want to install (or press Enter to install all by default)"

if ([string]::IsNullOrWhiteSpace($appInput)) {
    Write-Host "No applications selected. All will be installed by default." -ForegroundColor Yellow
    $selectedApps = $apps
} else {
    $selectedIndexes = $appInput -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ - 1 }
    $selectedApps = @()
    
    foreach ($index in $selectedIndexes) {
        if ($index -ge 0 -and $index -lt $apps.Count) {
            $selectedApps += $apps[$index]
        } else {
            Write-Host "Number out of range: $($index + 1). Ignoring this value." -ForegroundColor Red
        }
    }

    if ($selectedApps.Count -eq 0) {
        Write-Host "No valid applications selected. All will be installed by default." -ForegroundColor Yellow
        $selectedApps = $apps
    } else {
        Write-Host "The following applications will be installed: $($selectedApps -join ', ')" -ForegroundColor Cyan
    }
}

# Generate the configuration file
$config = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="$version">
        <Product ID="$productID">
            <Language ID="$language" />
"@

foreach ($app in $apps) {
    if (-not ($selectedApps -contains $app)) {
        $config += "            <ExcludeApp ID=\"$app\" />`n"
    }
}

# Ensure Skype for Business is excluded
$config += "            <ExcludeApp ID=\"Lync\" />`n"

$config += @"
        </Product>
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@

Set-Content -Path $officeConfigPath -Value $config
Write-Host "Configuration file generated at: $officeConfigPath" -ForegroundColor Green

# Start the installation
Write-Host "Starting the installation of Office LTSC..." -ForegroundColor Yellow

Start-Process -FilePath "$env:Temp\ODT\setup.exe" -ArgumentList "/configure $officeConfigPath" -Wait

Write-Host "Installation completed successfully." -ForegroundColor Green

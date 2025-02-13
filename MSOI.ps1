# Carpeta temporal para trabajar
$TempFolder = "$env:TEMP\OfficeSetup"

# URL oficial para descargar la herramienta de implementación de Office
$ODT_URL = "https://download.microsoft.com/download/2/6/E/26E3AEDE-10B1-4B6C-B3C1-9DB2B2E99328/OfficeDeploymentTool.exe"

# Paso 1: Crear carpeta temporal
if (-Not (Test-Path -Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
}

# Paso 2: Descargar Office Deployment Tool (ODT)
Write-Host "Descargando la herramienta de implementación de Office..."
$ODT_File = "$TempFolder\OfficeDeploymentTool.exe"
Invoke-WebRequest -Uri $ODT_URL -OutFile $ODT_File -ErrorAction Stop

if (-Not (Test-Path -Path $ODT_File)) {
    Write-Host "Error al descargar Office Deployment Tool. Verifica tu conexión a Internet." -ForegroundColor Red
    Pause
    Exit
}

# Paso 3: Extraer ODT
Write-Host "Extrayendo la herramienta de implementación de Office..."
Start-Process -FilePath $ODT_File -ArgumentList "/quiet /extract:$TempFolder" -NoNewWindow -Wait

# Paso 4: Solicitar opciones al usuario
Write-Host "=================================================="
Write-Host "Selecciona la version de Office que deseas instalar:" -ForegroundColor Cyan
Write-Host "1. Office 2024 LTSC ProPlus (Volume)"
Write-Host "2. Office 2021 LTSC ProPlus (Volume)"
Write-Host "3. Office 2019 LTSC ProPlus (Volume)"
$version = Read-Host "Ingresa el numero de tu opcion"

switch ($version) {
    "1" { $ProductID = "ProPlus2024Volume" }
    "2" { $ProductID = "ProPlus2021Volume" }
    "3" { $ProductID = "ProPlus2019Volume" }
    default {
        Write-Host "Opcion invalida. Saliendo..." -ForegroundColor Red
        Exit
    }
}

Write-Host "=================================================="
Write-Host "Selecciona los programas que deseas instalar (separados por comas):" -ForegroundColor Cyan
Write-Host "Word, Excel, PowerPoint, Outlook, Access, Publisher, Teams, OneDrive"
Write-Host "Nota: Deja en blanco para instalar todos."
$apps = Read-Host "Ingresa los programas"

Write-Host "=================================================="
Write-Host "Selecciona el idioma de instalación:" -ForegroundColor Cyan
Write-Host "1. Español"
Write-Host "2. Inglés"
$idioma = Read-Host "Ingresa el numero de tu opcion"

switch ($idioma) {
    "1" { $LanguageID = "es-es" }
    "2" { $LanguageID = "en-us" }
    default {
        Write-Host "Opcion invalida. Saliendo..." -ForegroundColor Red
        Exit
    }
}

# Paso 5: Crear archivo de configuración XML
$ConfigFile = "$TempFolder\configuration.xml"
Write-Host "Creando archivo de configuración..."

$xmlContent = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
        <Product ID="$ProductID">
"@

if ($apps -ne "") {
    $excludedApps = @("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "Teams", "OneDrive") | Where-Object { $apps -notmatch $_ }
    foreach ($app in $excludedApps) {
        $xmlContent += "            <ExcludeApp ID=\"$app\" />`n"
    }
}

$xmlContent += @"
            <Language ID="$LanguageID" />
        </Product>
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
    <Logging Name="install.log" Path="$TempFolder" Level="Standard" />
</Configuration>
"@

Set-Content -Path $ConfigFile -Value $xmlContent

# Paso 6: Iniciar la instalación de Office
Write-Host "Iniciando instalación de Office..."
Start-Process -FilePath "$TempFolder\setup.exe" -ArgumentList "/configure $ConfigFile" -NoNewWindow -Wait

# Paso 7: Limpiar archivos temporales
Write-Host "Eliminando archivos temporales..."
Remove-Item -Path $TempFolder -Recurse -Force

Write-Host "Instalación completada. ¡Disfruta tu nueva versión de Office!" -ForegroundColor Green
Pause

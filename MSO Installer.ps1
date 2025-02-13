# Habilitar depuración y detener en errores
$ErrorActionPreference = "Stop"
Set-PSDebug -Trace 1

# Permitir la ejecución del script sin restricciones en la sesión actual
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Configuración inicial
$TempFolder = "$env:TEMP\MSOInstaller"
if (-not (Test-Path -Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder | Out-Null
    Write-Host "Carpeta temporal creada: $TempFolder"
} else {
    Write-Host "Carpeta temporal ya existe: $TempFolder"
}

# Descargar Office Deployment Tool si no existe
$ODTPath = Join-Path -Path $TempFolder -ChildPath "OfficeDeploymentTool.exe"
if (-not (Test-Path -Path $ODTPath)) {
    Write-Host "Descargando Office Deployment Tool..."
    $ODTUrl = "https://go.microsoft.com/fwlink/p/?linkid=2085641"
    Invoke-WebRequest -Uri $ODTUrl -OutFile $ODTPath
    Write-Host "Descarga completa: $ODTPath"
} else {
    Write-Host "El archivo Office Deployment Tool ya existe: $ODTPath"
}

# Extraer el contenido del instalador
Write-Host "Extrayendo archivos de Office Deployment Tool..."
Start-Process -FilePath $ODTPath -ArgumentList "/quiet /extract:$TempFolder" -NoNewWindow -Wait
Write-Host "Extracción completa en: $TempFolder"

# Configurar el ProductID según la versión deseada
$Version = 2  # Cambia este valor a 1, 2 o 3 según la versión que necesites
switch ($Version) {
    1 { $ProductID = "ProPlus2024Volume" }
    2 { $ProductID = "ProPlus2021Volume" }
    3 { $ProductID = "ProPlus2019Volume" }
    default { Write-Host "Versión no válida. Usa 1, 2 o 3."; exit }
}
Write-Host "Configurando instalación con el ProductID: $ProductID"

# Crear el archivo de configuración XML
$ConfigXMLPath = Join-Path -Path $TempFolder -ChildPath "Configuration.xml"
$ConfigXMLContent = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
        <Product ID="$ProductID">
            <Language ID="en-us" />
        </Product>
    </Add>
    <Display Level="None" AcceptEULA="True" />
    <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@
Set-Content -Path $ConfigXMLPath -Value $ConfigXMLContent
Write-Host "Archivo de configuración creado en: $ConfigXMLPath"

# Ejecutar la instalación de Office
$SetupExePath = Join-Path -Path $TempFolder -ChildPath "setup.exe"
if (Test-Path -Path $SetupExePath) {
    Write-Host "Iniciando instalación de Office..."
    Start-Process -FilePath $SetupExePath -ArgumentList "/configure $ConfigXMLPath" -NoNewWindow -Wait
    Write-Host "Instalación completada."
} else {
    Write-Host "Error: No se encontró setup.exe en $TempFolder. Verifica la extracción."
    exit
}

Write-Host "El proceso ha finalizado correctamente."
Pause

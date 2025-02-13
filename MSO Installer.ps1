# Habilitar manejo de errores estrictos
$ErrorActionPreference = "Stop"

# Función para manejar errores
function Mostrar-Error {
    param ($ErrorMessage)
    Write-Host "`nERROR: $ErrorMessage" -ForegroundColor Red
    Read-Host "Presiona Enter para cerrar el script"
    Exit 1
}

# Ajustar tamaño de la ventana de PowerShell
function Ajustar-TamanoVentana {
    $Width = 80
    $Height = 30
    $host.ui.RawUI.WindowSize = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList $Width, $Height
    $host.ui.RawUI.BufferSize = New-Object -TypeName System.Management.Automation.Host.Size -ArgumentList $Width, 300
}

Ajustar-TamanoVentana

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
$ODTPath = Join-Path -Path $TempFolder -ChildPath "officedeploymenttool_18227-20162.exe"
if (-not (Test-Path -Path $ODTPath)) {
    Write-Host "Descargando Office Deployment Tool..."
    $ODTUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
    Invoke-WebRequest -Uri $ODTUrl -OutFile $ODTPath -UseBasicParsing -ProgressAction Show
    Write-Host "Descarga completa: $ODTPath"
} else {
    Write-Host "El archivo Office Deployment Tool ya existe: $ODTPath"
}

# Extraer el contenido del instalador
Write-Host "Extrayendo archivos de Office Deployment Tool..."
Start-Process -FilePath $ODTPath -ArgumentList "/quiet /extract:$TempFolder" -NoNewWindow -Wait
Write-Host "Extracción completa en: $TempFolder"

# Menú principal
function Mostrar-Menu {
    cls
    Write-Host "=========================================="
    Write-Host "    Instalador de Office - Menú Principal"
    Write-Host "=========================================="
    Write-Host "[1] Seleccionar versión de Office"
    Write-Host "[2] Configurar idioma de instalación"
    Write-Host "[3] Iniciar instalación de Office"
    Write-Host "[0] Salir"
    Write-Host "=========================================="
    return Read-Host "Elige una opción"
}

# Variables globales
$ProductID = $null
$Channel = $null
$Language = $null

# Lógica del menú
while ($true) {
    $Opcion = Mostrar-Menu
    switch ($Opcion) {
        "1" {
            Write-Host "`nSelecciona la versión de Office:"
            Write-Host "[1] Office 2024"
            Write-Host "[2] Office 2021"
            Write-Host "[3] Office 2019"
            $Version = Read-Host "Elige una opción (1, 2 o 3)"
            switch ($Version) {
                1 {
                    $ProductID = "ProPlus2024Volume"
                    $Channel = "PerpetualVL2024"
                }
                2 {
                    $ProductID = "ProPlus2021Volume"
                    $Channel = "PerpetualVL2021"
                }
                3 {
                    $ProductID = "ProPlus2019Volume"
                    $Channel = "PerpetualVL2019"
                }
                default {
                    Mostrar-Error "Versión no válida. Usa 1, 2 o 3."
                }
            }
            Write-Host "Configuración establecida: ProductID=$ProductID, Canal=$Channel" -ForegroundColor Cyan
        }
        "2" {
            $Language = Read-Host "Ingresa el idioma (ejemplo: en-us para inglés, es-es para español)"
            if (-not $Language) {
                Mostrar-Error "No ingresaste un idioma válido."
            } else {
                Write-Host "Idioma configurado: $Language" -ForegroundColor Cyan
            }
        }
        "3" {
            if (-not $ProductID -or -not $Channel -or -not $Language) {
                Mostrar-Error "Debes configurar la versión y el idioma antes de iniciar la instalación."
            }

            # Crear el archivo de configuración XML
            $ConfigXMLPath = Join-Path -Path $TempFolder -ChildPath "Configuration.xml"
            $ConfigXMLContent = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="$Channel">
        <Product ID="$ProductID">
            <Language ID="$Language" />
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
                Mostrar-Error "No se encontró setup.exe en $TempFolder. Verifica la extracción."
            }
        }
        "0" {
            Write-Host "Saliendo del instalador..." -ForegroundColor Green
            break
        }
        default {
            Write-Host "Opción no válida. Intenta de nuevo." -ForegroundColor Red
        }
    }
}

Write-Host "El proceso ha finalizado correctamente."
Pause

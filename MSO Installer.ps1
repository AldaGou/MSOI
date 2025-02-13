# Configuración inicial
$ErrorActionPreference = "Stop"

function MostrarMenu {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "            Office Installer       " -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "[1] Instalar Office 2024 Volume"
    Write-Host "[2] Instalar Office 2021 Volume"
    Write-Host "[3] Instalar Office 2019 Volume"
    Write-Host "[4] Cambiar idioma de instalación"
    Write-Host "[5] Salir"
    Write-Host "===================================" -ForegroundColor Cyan
}

function ElegirOpcion {
    param (
        [string]$Mensaje
    )
    Write-Host "`n$Mensaje"
    return Read-Host "Elige una opción"
}

function CambiarIdioma {
    Write-Host "`nIdiomas disponibles:" -ForegroundColor Yellow
    Write-Host "[1] Español (es-es)"
    Write-Host "[2] Inglés (en-us)"
    Write-Host "[3] Francés (fr-fr)"
    $opcionIdioma = Read-Host "Selecciona el idioma (1-3)"
    switch ($opcionIdioma) {
        1 { return "es-es" }
        2 { return "en-us" }
        3 { return "fr-fr" }
        default {
            Write-Host "Opción no válida, se usará Inglés por defecto."
            return "en-us"
        }
    }
}

# Variables de configuración
$Idioma = "en-us"

do {
    MostrarMenu
    $opcion = ElegirOpcion "Selecciona una opción (1-5)"
    switch ($opcion) {
        1 {
            Write-Host "Iniciando instalación de Office 2024 Volume con idioma $Idioma..."
            # Aquí iría la lógica de instalación para Office 2024 Volume
            Pause
        }
        2 {
            Write-Host "Iniciando instalación de Office 2021 Volume con idioma $Idioma..."
            # Aquí iría la lógica de instalación para Office 2021 Volume
            Pause
        }
        3 {
            Write-Host "Iniciando instalación de Office 2019 Volume con idioma $Idioma..."
            # Aquí iría la lógica de instalación para Office 2019 Volume
            Pause
        }
        4 {
            $Idioma = CambiarIdioma
            Write-Host "Idioma cambiado a $Idioma."
            Pause
        }
        5 {
            Write-Host "Saliendo del instalador. ¡Hasta pronto!" -ForegroundColor Green
            break
        }
        default {
            Write-Host "Opción no válida. Intenta de nuevo." -ForegroundColor Red
        }
    }
} while ($opcion -ne 5)

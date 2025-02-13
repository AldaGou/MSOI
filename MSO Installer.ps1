# Configuración inicial
$ErrorActionPreference = "Stop"

function MostrarMenuVersion {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "        Office Installer           " -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "[1] Office 2024 Volume"
    Write-Host "[2] Office 2021 Volume"
    Write-Host "[3] Office 2019 Volume"
    Write-Host "[4] Salir"
    Write-Host "===================================" -ForegroundColor Cyan
}

function ElegirVersion {
    param (
        [string]$Mensaje
    )
    Write-Host "`n$Mensaje"
    return Read-Host "Elige una versión (1-4)"
}

function MostrarMenuIdioma {
    Clear-Host
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "        Selección de Idioma        " -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Cyan
    Write-Host "[1] Español (es-es)"
    Write-Host "[2] Inglés (en-us)"
    Write-Host "[3] Francés (fr-fr)"
    Write-Host "[4] Volver al menú principal"
    Write-Host "===================================" -ForegroundColor Cyan
}

function ElegirIdioma {
    param (
        [string]$Mensaje
    )
    Write-Host "`n$Mensaje"
    return Read-Host "Selecciona el idioma (1-4)"
}

function ConfirmarInstalacion {
    Write-Host "`nHas seleccionado:" -ForegroundColor Yellow
    Write-Host "Versión de Office: $SeleccionVersion"
    Write-Host "Idioma: $SeleccionIdioma"
    Write-Host "===================================" -ForegroundColor Cyan
    $confirmar = Read-Host "¿Deseas iniciar la instalación? (S/N)"
    return $confirmar
}

# Variables para selección
$SeleccionVersion = ""
$SeleccionIdioma = ""

do {
    MostrarMenuVersion
    $opcionVersion = ElegirVersion "Selecciona la versión de Office que deseas instalar"
    switch ($opcionVersion) {
        1 { $SeleccionVersion = "Office 2024 Volume" }
        2 { $SeleccionVersion = "Office 2021 Volume" }
        3 { $SeleccionVersion = "Office 2019 Volume" }
        4 { break }
        default {
            Write-Host "Opción no válida. Intenta de nuevo." -ForegroundColor Red
            Pause
            continue
        }
    }

    if ($opcionVersion -ne 4) {
        do {
            MostrarMenuIdioma
            $opcionIdioma = ElegirIdioma "Selecciona el idioma para la instalación"
            switch ($opcionIdioma) {
                1 { $SeleccionIdioma = "es-es" }
                2 { $SeleccionIdioma = "en-us" }
                3 { $SeleccionIdioma = "fr-fr" }
                4 { break }
                default {
                    Write-Host "Opción no válida. Intenta de nuevo." -ForegroundColor Red
                    Pause
                    continue
                }
            }

            if ($opcionIdioma -ne 4) {
                $confirmar = ConfirmarInstalacion
                if ($confirmar -match "^[sS]$") {
                    Write-Host "Iniciando instalación de $SeleccionVersion en idioma $SeleccionIdioma..." -ForegroundColor Green
                    # Aquí puedes añadir la lógica de instalación
                    Pause
                    break
                } else {
                    Write-Host "Instalación cancelada. Regresando al menú principal." -ForegroundColor Red
                    Pause
                    break
                }
            }
        } while ($opcionIdioma -ne 4)
    }
} while ($opcionVersion -ne 4)

Write-Host "Saliendo del instalador. ¡Hasta pronto!" -ForegroundColor Green

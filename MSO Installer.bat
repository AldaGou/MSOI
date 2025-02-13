@echo off
:: Carpeta temporal para trabajar
set TempFolder=%TEMP%\OfficeSetup

:: URL oficial para descargar la herramienta de implementación de Office
set ODT_URL=https://download.microsoft.com/download/2/6/E/26E3AEDE-10B1-4B6C-B3C1-9DB2B2E99328/OfficeDeploymentTool.exe

:: Paso 1: Crear carpeta temporal
if not exist "%TempFolder%" mkdir "%TempFolder%"

:: Paso 2: Descargar Office Deployment Tool (ODT)
echo Descargando la herramienta de implementación de Office...
bitsadmin /transfer "ODTDownload" %ODT_URL% "%TempFolder%\OfficeDeploymentTool.exe"

if not exist "%TempFolder%\OfficeDeploymentTool.exe" (
    echo Error al descargar Office Deployment Tool. Verifica tu conexión a Internet.
    pause
    exit /b
)

:: Paso 3: Extraer ODT
echo Extrayendo la herramienta de implementación de Office...
"%TempFolder%\OfficeDeploymentTool.exe" /quiet /extract:%TempFolder%

:: Paso 4: Solicitar opciones al usuario
cls
echo ==================================================
echo Selecciona la version de Office que deseas instalar:
echo 1. Office 2024 LTSC ProPlus (Volume)
echo 2. Office 2021 LTSC ProPlus (Volume)
echo 3. Office 2019 LTSC ProPlus (Volume)
set /p version="Ingresa el numero de tu opcion: "

if "%version%"=="1" set ProductID=ProPlus2024Volume
if "%version%"=="2" set ProductID=ProPlus2021Volume
if "%version%"=="3" set ProductID=ProPlus2019Volume

cls
echo ==================================================
echo Selecciona los programas que deseas instalar (separados por comas):
echo Word, Excel, PowerPoint, Outlook, Access, Publisher, Teams, OneDrive
echo Nota: Deja en blanco para instalar todos.
set /p apps="Ingresa los programas: "

cls
echo ==================================================
echo Selecciona el idioma de instalación:
echo 1. Español
echo 2. Inglés
set /p idioma="Ingresa el numero de tu opcion: "

if "%idioma%"=="1" set LanguageID=es-es
if "%idioma%"=="2" set LanguageID=en-us

:: Paso 5: Crear archivo de configuración XML
set ConfigFile=%TempFolder%\configuration.xml

echo Creando archivo de configuración...
(
    echo ^<Configuration^>
    echo     ^<Add OfficeClientEdition="64" Channel="PerpetualVL2021" ^>
    echo         ^<Product ID="%ProductID%"^>
    if not "%apps%"=="" (
        echo             ^<ExcludeApp ID="Groove" /^>
        for %%A in (Word Excel PowerPoint Outlook Access Publisher Teams OneDrive) do (
            echo %apps% | find /i "%%A" >nul || echo             ^<ExcludeApp ID="%%A" /^>
        )
    )
    echo             ^<Language ID="%LanguageID%" /^>
    echo         ^</Product^>
    echo     ^</Add^>
    echo     ^<Display Level="Full" AcceptEULA="TRUE" /^>
    echo     ^<Logging Name="install.log" Path="%TempFolder%" Level="Standard" /^>
    echo ^</Configuration^>
) > "%ConfigFile%"

:: Paso 6: Iniciar la instalación de Office
echo Iniciando instalación de Office...
"%TempFolder%\setup.exe" /configure "%ConfigFile%"

:: Paso 7: Limpiar archivos temporales
echo Eliminando archivos temporales...
rd /s /q "%TempFolder%"

echo Instalación completada. ¡Disfruta tu nueva versión de Office!
pause

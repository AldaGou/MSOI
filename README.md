# MSOI - Microsoft Office Installer

Herramienta para descargar e instalar Microsoft Office LTSC/2019/2016/2013 desde PowerShell.

## Versiones disponibles

- Office LTSC Professional Plus 2024
- Office LTSC Professional Plus 2021
- Office Professional Plus 2019
- Office Professional Plus 2016
- Office Professional Plus 2013

## Uso

### Interfaz gráfica (GUI)

```powershell
irm https://aldagou.github.io/MSOI/MSOIGUI.ps1 | iex
```

### Línea de comandos (CLI)

```powershell
irm https://aldagou.github.io/MSOI/MSOICLI.ps1 | iex
```

## Requisitos

- Windows 10 / 11
- PowerShell 5.0 o superior
- Ejecutar como **Administrador**
- .NET Framework 4.5+

## Características

- Selección de versión de Office
- Arquitectura 64/32 bits
- Selección de idioma
- Incluir Project y Visio
- Selección individual de aplicaciones (Word, Excel, PowerPoint, etc.)
- Modos: Descargar e Instalar, Sólo Descargar, Instalar desde caché
- Las aplicaciones no deseadas (Bing, Groove, Lync, OneDrive, Teams) se excluyen automáticamente

## Notas

La herramienta descarga automáticamente el **Office Deployment Tool** de Microsoft y genera el archivo `configuration.xml` según tus selecciones.

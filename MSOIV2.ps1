<#
.SYNOPSIS
    MSOI - Microsoft Office Installation Tool (GUI)
.DESCRIPTION
    Interfaz gráfica para descargar e instalar Microsoft Office LTSC.
    Soporta múltiples ediciones, arquitecturas, idiomas y personalización.
.NOTES
    Requiere: Administrador, PowerShell 5.0+, .NET Framework 4.5+
    Uso: irm https://aldagou.github.io/MSOI/MSOI.ps1 | iex
#>

#Requires -RunAsAdministrator

# ---- CARGAR ENSAMBLADOS ----
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- DPI AWARENESS (resolve rendering en alta resolución) ----
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class MSOI_DpiHelper {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int awareness);
}
"@
try { [MSOI_DpiHelper]::SetProcessDpiAwareness(1) } catch { try { [MSOI_DpiHelper]::SetProcessDPIAware() } catch {} }

# ---- VARIABLES GLOBALES ----
$script:logFile  = Join-Path $env:Temp "MSOI_Install.log"
$script:odtTemp  = Join-Path $env:Temp "ODT"
$script:odtExe   = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtReady = $false

$script:odtUrls  = @(
    "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
)

# ---- FUNCIONES AUXILIARES ----
function Write-Log {
    param([string]$Message, [string]$Color = "Gray")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $script:logFile -Value "[$timestamp] $Message" -ErrorAction SilentlyContinue
    Write-Host $Message -ForegroundColor $Color
}

function Get-LanguageCode {
    param([string]$DisplayText)
    if ($DisplayText -match '\(([^)]+)\)') { return $matches[1] }
    return "es-ES"
}

# ---- FORMULARIO DE PREPARACIÓN (ODT) ----
function Show-PreparationDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text              = "MSOI - Preparación"
    $form.Size              = New-Object System.Drawing.Size(460, 190)
    $form.AutoScaleMode     = "Dpi"
    $form.FormBorderStyle   = "FixedDialog"
    $form.ControlBox        = $false
    $form.StartPosition     = "CenterScreen"
    $form.BackColor         = "White"
    $form.TopMost           = $true

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text          = "Preparando Office Deployment Tool..."
    $lblTitle.Location      = New-Object System.Drawing.Point(25, 30)
    $lblTitle.Size          = New-Object System.Drawing.Size(410, 25)
    $lblTitle.Font          = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblTitle)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Name         = "lblStatus"
    $lblStatus.Text         = "Iniciando..."
    $lblStatus.Location     = New-Object System.Drawing.Point(25, 60)
    $lblStatus.Size         = New-Object System.Drawing.Size(410, 20)
    $lblStatus.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($lblStatus)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Name       = "progressBar"
    $progressBar.Location   = New-Object System.Drawing.Point(25, 95)
    $progressBar.Size       = New-Object System.Drawing.Size(410, 25)
    $progressBar.Style      = "Marquee"
    $progressBar.MarqueeAnimationSpeed = 30
    $form.Controls.Add($progressBar)

    $form.Show()
    $form.Refresh()

    function Update-Status($text) {
        $form.Controls["lblStatus"].Text = $text
        $form.Refresh()
        Start-Sleep -Milliseconds 100
    }

    try {
        Update-Status "Limpiando archivos temporales anteriores..."
        if (Test-Path $script:odtTemp) {
            Remove-Item -Path $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $script:odtExe) {
            Remove-Item -Path $script:odtExe -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $script:odtTemp -Force | Out-Null

        $downloaded = $false
        foreach ($url in $script:odtUrls) {
            Update-Status "Descargando Office Deployment Tool..."
            try {
                $wc = New-Object System.Net.WebClient
                $wc.DownloadFile($url, $script:odtExe)
                $downloaded = $true
                break
            } catch {
                Update-Status "Error, intentando con servidor alternativo..."
                Start-Sleep -Milliseconds 500
            }
        }

        if (-not $downloaded) {
            throw "No se pudo descargar el ODT desde ninguna fuente."
        }

        Update-Status "Extrayendo Office Deployment Tool..."
        Start-Sleep -Milliseconds 300
        $proc = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$($script:odtTemp)`"" -Wait -PassThru
        if ($proc.ExitCode -ne 0) {
            throw "Extracción falló (ExitCode: $($proc.ExitCode))."
        }

        $setupExe = Join-Path $script:odtTemp "setup.exe"
        if (-not (Test-Path $setupExe)) {
            throw "No se encontró setup.exe después de extraer."
        }

        $script:odtReady = $true
        Update-Status "Listo."
        Start-Sleep -Milliseconds 400
    } catch {
        Update-Status "ERROR: $_"
        Write-Log "Error en preparación: $_" "Red"
        Start-Sleep -Milliseconds 800
        [System.Windows.Forms.MessageBox]::Show(
            "Error al preparar el Office Deployment Tool.`n`n$_`n`nDescárgalo manualmente desde:`nhttps://www.microsoft.com/en-us/download/details.aspx?id=49117",
            "MSOI - Error",
            "OK",
            "Error"
        )
    }

    $form.Close()
}

# ---- FORMULARIO PRINCIPAL ----
function Show-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text              = "MSOI - Instalación de Microsoft Office"
    $form.Size              = New-Object System.Drawing.Size(720, 670)
    $form.MinimumSize       = New-Object System.Drawing.Size(720, 670)
    $form.AutoScaleMode     = "Dpi"
    $form.StartPosition     = "CenterScreen"
    $form.BackColor         = "White"
    $form.Font              = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
    $form.FormBorderStyle   = "FixedSingle"
    $form.MaximizeBox       = $false

    # =============== TÍTULO ===============
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text          = "MSOI - Instalación de Microsoft Office"
    $lblTitle.Location      = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size          = New-Object System.Drawing.Size(680, 30)
    $lblTitle.Font          = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor     = "#1a5276"
    $form.Controls.Add($lblTitle)

    $lblSub = New-Object System.Windows.Forms.Label
    $lblSub.Text            = "Selecciona las opciones de instalación"
    $lblSub.Location        = New-Object System.Drawing.Point(20, 48)
    $lblSub.Size            = New-Object System.Drawing.Size(680, 20)
    $lblSub.ForeColor       = "#666666"
    $form.Controls.Add($lblSub)

    # =============== VERSIÓN ===============
    $grpVersion = New-Object System.Windows.Forms.GroupBox
    $grpVersion.Text        = "Versión de Office"
    $grpVersion.Location    = New-Object System.Drawing.Point(15, 75)
    $grpVersion.Size        = New-Object System.Drawing.Size(685, 55)
    $grpVersion.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpVersion)

    $cmbVersion = New-Object System.Windows.Forms.ComboBox
    $cmbVersion.Name        = "cmbVersion"
    $cmbVersion.Location    = New-Object System.Drawing.Point(15, 22)
    $cmbVersion.Size        = New-Object System.Drawing.Size(650, 25)
    $cmbVersion.DropDownStyle = "DropDownList"
    $cmbVersion.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbVersion.Items.AddRange(@(
        "Office LTSC Professional Plus 2024",
        "Office LTSC Standard 2024",
        "Office LTSC Professional Plus 2021",
        "Office LTSC Standard 2021",
        "Office Professional Plus 2019",
        "Office Standard 2019"
    ))
    $cmbVersion.SelectedIndex = 0
    $grpVersion.Controls.Add($cmbVersion)

    # =============== ARQUITECTURA + IDIOMA ===============
    $grpArchLang = New-Object System.Windows.Forms.GroupBox
    $grpArchLang.Text       = "Arquitectura e Idioma"
    $grpArchLang.Location   = New-Object System.Drawing.Point(15, 140)
    $grpArchLang.Size       = New-Object System.Drawing.Size(685, 75)
    $grpArchLang.Font       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpArchLang)

    $rbArch64 = New-Object System.Windows.Forms.RadioButton
    $rbArch64.Name          = "rbArch64"
    $rbArch64.Text          = "64 bits (recomendado)"
    $rbArch64.Location      = New-Object System.Drawing.Point(15, 22)
    $rbArch64.Size          = New-Object System.Drawing.Size(160, 22)
    $rbArch64.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
    $rbArch64.Checked       = $true
    $grpArchLang.Controls.Add($rbArch64)

    $rbArch32 = New-Object System.Windows.Forms.RadioButton
    $rbArch32.Name          = "rbArch32"
    $rbArch32.Text          = "32 bits"
    $rbArch32.Location      = New-Object System.Drawing.Point(15, 45)
    $rbArch32.Size          = New-Object System.Drawing.Size(160, 22)
    $rbArch32.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
    $grpArchLang.Controls.Add($rbArch32)

    $lblLang = New-Object System.Windows.Forms.Label
    $lblLang.Text           = "Idioma:"
    $lblLang.Location       = New-Object System.Drawing.Point(240, 24)
    $lblLang.Size           = New-Object System.Drawing.Size(60, 20)
    $lblLang.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpArchLang.Controls.Add($lblLang)

    $cmbLang = New-Object System.Windows.Forms.ComboBox
    $cmbLang.Name           = "cmbLang"
    $cmbLang.Location       = New-Object System.Drawing.Point(300, 22)
    $cmbLang.Size           = New-Object System.Drawing.Size(365, 25)
    $cmbLang.DropDownStyle  = "DropDownList"
    $cmbLang.Font           = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbLang.Items.AddRange(@(
        "Español (es-ES)",
        "English (en-US)",
        "Français (fr-FR)",
        "Deutsch (de-DE)",
        "Português (pt-BR)",
        "Italiano (it-IT)",
        "Nederlands (nl-NL)",
        "Polski (pl-PL)",
        "Русский (ru-RU)",
        "日本語 (ja-JP)"
    ))
    $cmbLang.SelectedIndex = 0
    $grpArchLang.Controls.Add($cmbLang)

    # =============== PRODUCTOS ADICIONALES ===============
    $grpExtra = New-Object System.Windows.Forms.GroupBox
    $grpExtra.Text          = "Productos adicionales"
    $grpExtra.Location      = New-Object System.Drawing.Point(15, 225)
    $grpExtra.Size          = New-Object System.Drawing.Size(685, 70)
    $grpExtra.Font          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpExtra)

    $chkProject = New-Object System.Windows.Forms.CheckBox
    $chkProject.Name        = "chkProject"
    $chkProject.Text        = "Incluir Project"
    $chkProject.Location    = New-Object System.Drawing.Point(15, 22)
    $chkProject.Size        = New-Object System.Drawing.Size(130, 22)
    $chkProject.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpExtra.Controls.Add($chkProject)

    $cmbProjectEd = New-Object System.Windows.Forms.ComboBox
    $cmbProjectEd.Name      = "cmbProjectEd"
    $cmbProjectEd.Location  = New-Object System.Drawing.Point(150, 22)
    $cmbProjectEd.Size      = New-Object System.Drawing.Size(150, 25)
    $cmbProjectEd.DropDownStyle = "DropDownList"
    $cmbProjectEd.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbProjectEd.Items.AddRange(@("Professional", "Standard"))
    $cmbProjectEd.SelectedIndex = 0
    $cmbProjectEd.Enabled   = $false
    $grpExtra.Controls.Add($cmbProjectEd)

    $chkProject.Add_CheckedChanged({
        $cmbProjectEd.Enabled = $chkProject.Checked
    })

    $chkVisio = New-Object System.Windows.Forms.CheckBox
    $chkVisio.Name          = "chkVisio"
    $chkVisio.Text          = "Incluir Visio"
    $chkVisio.Location      = New-Object System.Drawing.Point(15, 44)
    $chkVisio.Size          = New-Object System.Drawing.Size(130, 22)
    $chkVisio.Font          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpExtra.Controls.Add($chkVisio)

    $cmbVisioEd = New-Object System.Windows.Forms.ComboBox
    $cmbVisioEd.Name        = "cmbVisioEd"
    $cmbVisioEd.Location    = New-Object System.Drawing.Point(150, 44)
    $cmbVisioEd.Size        = New-Object System.Drawing.Size(150, 25)
    $cmbVisioEd.DropDownStyle = "DropDownList"
    $cmbVisioEd.Font        = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbVisioEd.Items.AddRange(@("Professional", "Standard"))
    $cmbVisioEd.SelectedIndex = 0
    $cmbVisioEd.Enabled     = $false
    $grpExtra.Controls.Add($cmbVisioEd)

    $chkVisio.Add_CheckedChanged({
        $cmbVisioEd.Enabled = $chkVisio.Checked
    })

    # =============== APLICACIONES ===============
    $grpApps = New-Object System.Windows.Forms.GroupBox
    $grpApps.Text           = "Aplicaciones a instalar"
    $grpApps.Location       = New-Object System.Drawing.Point(15, 305)
    $grpApps.Size           = New-Object System.Drawing.Size(685, 115)
    $grpApps.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpApps)

    $chkSelectAll = New-Object System.Windows.Forms.CheckBox
    $chkSelectAll.Name      = "chkSelectAll"
    $chkSelectAll.Text      = "Seleccionar todas"
    $chkSelectAll.Location  = New-Object System.Drawing.Point(15, 22)
    $chkSelectAll.Size      = New-Object System.Drawing.Size(140, 22)
    $chkSelectAll.Checked   = $true
    $chkSelectAll.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpApps.Controls.Add($chkSelectAll)

    $appDefs = @(
        @{ID="Word";       X=15;  Y=48; Checked=$true}
        @{ID="Excel";      X=175; Y=48; Checked=$true}
        @{ID="PowerPoint"; X=335; Y=48; Checked=$true}
        @{ID="Outlook";    X=495; Y=48; Checked=$true}
        @{ID="Access";     X=15;  Y=72; Checked=$true}
        @{ID="Publisher";  X=175; Y=72; Checked=$true}
        @{ID="OneNote";    X=335; Y=72; Checked=$true}
        @{ID="SkypeForBusiness"; X=495; Y=72; Checked=$false}
    )

    $appCheckboxes = New-Object System.Collections.ArrayList
    foreach ($app in $appDefs) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Name           = "chk_$($app.ID)"
        $chk.Text           = $app.ID
        $chk.Location       = New-Object System.Drawing.Point($app.X, $app.Y)
        $chk.Size           = New-Object System.Drawing.Size(150, 22)
        $chk.Checked        = $app.Checked
        $chk.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
        $grpApps.Controls.Add($chk)
        [void]$appCheckboxes.Add($chk)

        $chk.Add_CheckedChanged({
            $allChecked = $true
            foreach ($cb in $appCheckboxes) {
                if (-not $cb.Checked) { $allChecked = $false; break }
            }
            $chkSelectAll.Checked = $allChecked
        })
    }

    $chkSelectAll.Add_CheckedChanged({
        foreach ($cb in $appCheckboxes) {
            $cb.Checked = $chkSelectAll.Checked
        }
    })

    # =============== MODO DE INSTALACIÓN ===============
    $grpMode = New-Object System.Windows.Forms.GroupBox
    $grpMode.Text           = "Modo de instalación"
    $grpMode.Location       = New-Object System.Drawing.Point(15, 430)
    $grpMode.Size           = New-Object System.Drawing.Size(685, 75)
    $grpMode.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($grpMode)

    $rbMode1 = New-Object System.Windows.Forms.RadioButton
    $rbMode1.Name           = "rbMode1"
    $rbMode1.Text           = "Descargar e instalar (recomendado)"
    $rbMode1.Location       = New-Object System.Drawing.Point(15, 22)
    $rbMode1.Size           = New-Object System.Drawing.Size(300, 22)
    $rbMode1.Checked        = $true
    $rbMode1.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpMode.Controls.Add($rbMode1)

    $rbMode2 = New-Object System.Windows.Forms.RadioButton
    $rbMode2.Name           = "rbMode2"
    $rbMode2.Text           = "Solo descargar (sin instalar)"
    $rbMode2.Location       = New-Object System.Drawing.Point(15, 46)
    $rbMode2.Size           = New-Object System.Drawing.Size(300, 22)
    $rbMode2.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpMode.Controls.Add($rbMode2)

    $rbMode3 = New-Object System.Windows.Forms.RadioButton
    $rbMode3.Name           = "rbMode3"
    $rbMode3.Text           = "Instalar desde descarga previa"
    $rbMode3.Location       = New-Object System.Drawing.Point(350, 22)
    $rbMode3.Size           = New-Object System.Drawing.Size(300, 22)
    $rbMode3.Font           = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $grpMode.Controls.Add($rbMode3)

    # =============== BARRA DE ESTADO ===============
    $statusBar = New-Object System.Windows.Forms.Label
    $statusBar.Name         = "statusBar"
    $statusBar.Text         = "Estado: Listo para instalar"
    $statusBar.Location     = New-Object System.Drawing.Point(15, 520)
    $statusBar.Size         = New-Object System.Drawing.Size(685, 28)
    $statusBar.Font         = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $statusBar.ForeColor    = "#555555"
    $statusBar.BorderStyle  = "FixedSingle"
    $statusBar.TextAlign    = "MiddleLeft"
    $form.Controls.Add($statusBar)

    # =============== BOTONES ===============
    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Name        = "btnInstall"
    $btnInstall.Text        = "Instalar Office"
    $btnInstall.Location    = New-Object System.Drawing.Point(15, 565)
    $btnInstall.Size        = New-Object System.Drawing.Size(170, 40)
    $btnInstall.Font        = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnInstall.ForeColor   = [System.Drawing.Color]::White
    $btnInstall.BackColor   = [System.Drawing.Color]::FromArgb(46, 134, 193)
    $btnInstall.FlatStyle   = "Flat"
    $btnInstall.FlatAppearance.BorderSize = 0
    $btnInstall.Cursor      = "Hand"
    $form.Controls.Add($btnInstall)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Name         = "btnCancel"
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(195, 565)
    $btnCancel.Size         = New-Object System.Drawing.Size(100, 40)
    $btnCancel.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnCancel.Cursor       = "Hand"
    $form.Controls.Add($btnCancel)

    # =============== MAPA DE VERSIONES ===============
    $versionMap = @(
        @{Channel="PerpetualVL2024"; ProductID="ProPlus2024Volume"; VisioPro="VisioPro2024Volume"; VisioStd="VisioStd2024Volume"; ProjectPro="ProjectPro2024Volume"; ProjectStd="ProjectStd2024Volume"}
        @{Channel="PerpetualVL2024"; ProductID="Standard2024Volume"; VisioPro="VisioPro2024Volume"; VisioStd="VisioStd2024Volume"; ProjectPro="ProjectPro2024Volume"; ProjectStd="ProjectStd2024Volume"}
        @{Channel="PerpetualVL2021"; ProductID="ProPlus2021Volume"; VisioPro="VisioPro2021Volume"; VisioStd="VisioStd2021Volume"; ProjectPro="ProjectPro2021Volume"; ProjectStd="ProjectStd2021Volume"}
        @{Channel="PerpetualVL2021"; ProductID="Standard2021Volume"; VisioPro="VisioPro2021Volume"; VisioStd="VisioStd2021Volume"; ProjectPro="ProjectPro2021Volume"; ProjectStd="ProjectStd2021Volume"}
        @{Channel="PerpetualVL2019"; ProductID="ProPlus2019Volume"; VisioPro="VisioPro2019Volume"; VisioStd="VisioStd2019Volume"; ProjectPro="ProjectPro2019Volume"; ProjectStd="ProjectStd2019Volume"}
        @{Channel="PerpetualVL2019"; ProductID="Standard2019Volume"; VisioPro="VisioPro2019Volume"; VisioStd="VisioStd2019Volume"; ProjectPro="ProjectPro2019Volume"; ProjectStd="ProjectStd2019Volume"}
    )

    # =============== EVENTOS ===============
    $btnCancel.Add_Click({ $form.Close() })

    $btnInstall.Add_Click({
        $btnInstall.Enabled = $false
        $btnCancel.Enabled  = $false
        $form.Cursor        = "WaitCursor"
        $statusBar.Text     = "Estado: Preparando configuración..."
        $form.Refresh()

        try {
            $vinfo          = $versionMap[$cmbVersion.SelectedIndex]
            $arch           = if ($rbArch64.Checked) { "64" } else { "32" }
            $langCode       = Get-LanguageCode $cmbLang.SelectedItem.ToString()
            $includeProject = $chkProject.Checked
            $projectEd      = if ($includeProject) { if ($cmbProjectEd.SelectedItem -eq "Standard") { "Std" } else { "Pro" } } else { $null }
            $includeVisio   = $chkVisio.Checked
            $visioEd        = if ($includeVisio) { if ($cmbVisioEd.SelectedItem -eq "Standard") { "Std" } else { "Pro" } } else { $null }

            $downloadMode = "1"
            if ($rbMode2.Checked) { $downloadMode = "2" }
            if ($rbMode3.Checked) { $downloadMode = "3" }

            $selectedApps = @()
            foreach ($cb in $appCheckboxes) {
                if ($cb.Checked) { $selectedApps += $cb.Text }
            }

            $statusBar.Text = "Estado: Generando configuración XML..."
            $form.Refresh()

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('<Configuration>')
            [void]$sb.AppendLine("    <Add OfficeClientEdition=`"$arch`" Channel=`"$($vinfo.Channel)`">")
            [void]$sb.AppendLine("        <Product ID=`"$($vinfo.ProductID)`">")
            [void]$sb.AppendLine("            <Language ID=`"$langCode`" />")

            $allAppIDs = @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")
            foreach ($app in $allAppIDs) {
                if ($selectedApps -notcontains $app) {
                    [void]$sb.AppendLine("            <ExcludeApp ID=`"$app`" />")
                }
            }

            foreach ($ex in @("Bing","Groove","Lync","OneDrive","Teams")) {
                [void]$sb.AppendLine("            <ExcludeApp ID=`"$ex`" />")
            }

            [void]$sb.AppendLine("        </Product>")

            if ($includeProject) {
                $projID = if ($projectEd -eq "Std") { $vinfo.ProjectStd } else { $vinfo.ProjectPro }
                [void]$sb.AppendLine("        <Product ID=`"$projID`">")
                [void]$sb.AppendLine("            <Language ID=`"$langCode`" />")
                [void]$sb.AppendLine("        </Product>")
            }

            if ($includeVisio) {
                $visID = if ($visioEd -eq "Std") { $vinfo.VisioStd } else { $vinfo.VisioPro }
                [void]$sb.AppendLine("        <Product ID=`"$visID`">")
                [void]$sb.AppendLine("            <Language ID=`"$langCode`" />")
                [void]$sb.AppendLine("        </Product>")
            }

            [void]$sb.AppendLine("    </Add>")

            if ($downloadMode -eq "2") {
                [void]$sb.AppendLine('    <Display Level="None" AcceptEULA="TRUE" />')
                [void]$sb.AppendLine("    <Download Path=`"$script:odtTemp`" />")
            } else {
                if ($downloadMode -eq "3") {
                    $srcPath = Join-Path $script:odtTemp "Office"
                    if (-not (Test-Path $srcPath)) {
                        throw "No se encontró la carpeta 'Office' en $script:odtTemp. Usa 'Solo descargar' primero."
                    }
                }
                [void]$sb.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />')
            }

            [void]$sb.AppendLine('</Configuration>')

            $configXml  = $sb.ToString()
            $configPath = Join-Path $script:odtTemp "configuration.xml"
            Set-Content -Path $configPath -Value $configXml -Encoding UTF8
            Write-Log "Configuración generada: $configPath" "Green"

            if ($downloadMode -eq "2") {
                $statusBar.Text = "Estado: Descargando Office (sin instalar)..."
            } else {
                $statusBar.Text = "Estado: Instalando Office..."
            }
            $form.Refresh()

            $setupExe = Join-Path $script:odtTemp "setup.exe"
            $arg = if ($downloadMode -eq "2") { "/download" } else { "/configure" }
            $proc = Start-Process -FilePath $setupExe -ArgumentList "$arg `"$configPath`"" -Wait -PassThru

            if ($proc.ExitCode -eq 0) {
                $statusBar.Text   = "Estado: Operación completada exitosamente."
                $statusBar.ForeColor = "Green"
                Write-Log "Operación completada exitosamente." "Green"
                [System.Windows.Forms.MessageBox]::Show(
                    "La operación se completó exitosamente.",
                    "MSOI - Éxito",
                    "OK",
                    "Information"
                )
            } else {
                $statusBar.Text   = "Estado: Error (código: $($proc.ExitCode)). Revisa los logs."
                $statusBar.ForeColor = "Red"
                Write-Log "Operación falló con código: $($proc.ExitCode)." "Red"
                [System.Windows.Forms.MessageBox]::Show(
                    "La operación falló (código: $($proc.ExitCode)).`n`nLogs: $script:odtTemp",
                    "MSOI - Error",
                    "OK",
                    "Error"
                )
            }
        } catch {
            $statusBar.Text   = "Estado: Error - $_"
            $statusBar.ForeColor = "Red"
            Write-Log "Error: $_" "Red"
            [System.Windows.Forms.MessageBox]::Show(
                "Error: $_",
                "MSOI - Error",
                "OK",
                "Error"
            )
        } finally {
            $btnInstall.Enabled = $true
            $btnCancel.Enabled  = $true
            $form.Cursor        = "Default"
            $form.Refresh()
        }
    })

    # =============== MOSTRAR ===============
    [void]$form.ShowDialog()
}

# ====================================================
# MAIN
# ====================================================
Write-Log "=== MSOI GUI v2.0 iniciado ===" "Cyan"
Write-Log "Sistema: $((Get-CimInstance Win32_OperatingSystem).Caption)" "Gray"

Show-PreparationDialog

if (-not $script:odtReady) {
    Write-Log "ODT no disponible. Abortando." "Red"
    [System.Windows.Forms.MessageBox]::Show(
        "No se pudo preparar el Office Deployment Tool. El script no puede continuar.",
        "MSOI - Error crítico",
        "OK",
        "Error"
    )
    exit 1
}

$null = Show-MainForm

Write-Log "=== MSOI GUI finalizado ===" "Cyan"

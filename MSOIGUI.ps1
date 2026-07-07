<#
.SYNOPSIS
    MSOI - Microsoft Office Installation Tool (GUI)
.DESCRIPTION
    Graphical tool to download and install Microsoft Office LTSC.
    Supports multiple editions, architectures, languages and app selection.
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+
    Usage: irm https://aldagou.github.io/MSOI/MSOIGUI.ps1 | iex
#>

#Requires -RunAsAdministrator

# ---- LOAD ASSEMBLIES ----
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- DPI AWARENESS (using here-string to avoid escaping issues) ----
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class MsoiDpi {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int a);
}
"@
try { [MsoiDpi]::SetProcessDpiAwareness(1) } catch { try { [MsoiDpi]::SetProcessDPIAware() } catch {} }

# ---- CONSTANTS ----
$script:logFile = Join-Path $env:Temp "MSOI_Install.log"
$script:odtTemp = Join-Path $env:Temp "ODT"
$script:odtExe  = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtUrls = @("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
$script:odtReady = $false
$script:is64Bit = [Environment]::Is64BitOperatingSystem

# Margins & sizes (defined inside Show-MainForm)
# ---- HELPERS ----
function Write-Log {
    param([string]$M, [string]$C = "Gray")
    Add-Content -Path $script:logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" -ErrorAction SilentlyContinue
    Write-Host $M -ForegroundColor $C
}

function Get-LangCode {
    param([string]$T)
    if ($T -match '\(([^)]+)\)') { return $matches[1] }
    return "en-US"
}

# ---- PREPARATION DIALOG ----
function Show-PrepDialog {
    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "MSOI - Preparation"
    $f.Size = New-Object System.Drawing.Size(460, 185)
    $f.FormBorderStyle = "FixedDialog"
    $f.ControlBox = $false
    $f.StartPosition = "CenterScreen"
    $f.BackColor = "White"
    $f.TopMost = $true
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $m = 20
    $w = 420

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Preparing Office Deployment Tool..."
    $l1.Location = New-Object System.Drawing.Point($m, 28)
    $l1.Size = New-Object System.Drawing.Size($w, 24)
    $l1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l1)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Name = "s"
    $l2.Text = "Starting..."
    $l2.Location = New-Object System.Drawing.Point($m, 56)
    $l2.Size = New-Object System.Drawing.Size($w, 18)
    $f.Controls.Add($l2)

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Name = "pb"
    $pb.Location = New-Object System.Drawing.Point($m, 85)
    $pb.Size = New-Object System.Drawing.Size($w, 26)
    $pb.Style = "Marquee"
    $f.Controls.Add($pb)

    $f.Show()
    $f.Refresh()

    function Set-Status($t) { $f.Controls["s"].Text = $t; $f.Refresh(); Start-Sleep -Milliseconds 80 }

    try {
        Set-Status "Cleaning up previous files..."
        if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $script:odtTemp -Force | Out-Null

        $ok = $false
        foreach ($u in $script:odtUrls) {
            Set-Status "Downloading Office Deployment Tool..."
            try { (New-Object System.Net.WebClient).DownloadFile($u, $script:odtExe); $ok = $true; break } catch { Set-Status "Retrying alternate source..." }
        }
        if (-not $ok) { throw "Could not download ODT from any source." }

        Set-Status "Extracting Office Deployment Tool..."
        Start-Sleep -Milliseconds 200
        $p = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$script:odtTemp`"" -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "Extraction failed (ExitCode: $($p.ExitCode))." }
        if (-not (Test-Path (Join-Path $script:odtTemp "setup.exe"))) { throw "setup.exe not found after extraction." }

        $script:odtReady = $true
        Set-Status "Done."
        Start-Sleep -Milliseconds 300
    } catch {
        Set-Status "ERROR: $_"
        Write-Log "Prep error: $_" "Red"
        Start-Sleep -Milliseconds 600
        [System.Windows.Forms.MessageBox]::Show("Failed to prepare the Office Deployment Tool.`n`n$_`n`nDownload manually from:`nhttps://www.microsoft.com/en-us/download/details.aspx?id=49117", "MSOI - Error", "OK", "Error")
    }
    $f.Close()
}

# ---- MAIN FORM ----
function Show-MainForm {
    $M = 20
    $FW = 800
    $GW = 760
    $y = 0

    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "MSOI - Microsoft Office Installer"
    $f.ClientSize = New-Object System.Drawing.Size($FW, 660)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedSingle"
    $f.MaximizeBox = $false
    $f.BackColor = "White"
    try { $f.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch {}
    try { $f.DoubleBuffered = $true } catch {}

    # ===== TITLE =====
    $y += 20
    $t = New-Object System.Windows.Forms.Label
    $t.Text = "MSOI - Microsoft Office Installer"
    $t.Location = New-Object System.Drawing.Point($M, $y)
    $t.Size = New-Object System.Drawing.Size($GW, 36)
    $t.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $t.TextAlign = "MiddleCenter"
    $f.Controls.Add($t)

    # ===== VERSION =====
    $y += 46
    $gv = New-Object System.Windows.Forms.GroupBox
    $gv.Text = " Office Version "
    $gv.Location = New-Object System.Drawing.Point($M, $y)
    $gv.Size = New-Object System.Drawing.Size($GW, 56)
    $f.Controls.Add($gv)

    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Name = "cv"
    $cv.Location = New-Object System.Drawing.Point(10, 20)
    $cv.Size = New-Object System.Drawing.Size(($GW - 24), 26)
    $cv.DropDownStyle = "DropDownList"
    $cv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cv.Items.AddRange(@("Office LTSC Professional Plus 2024", "Office LTSC Professional Plus 2021", "Office Professional Plus 2019", "Office Professional Plus 2016", "Office Professional Plus 2013"))
    $cv.SelectedIndex = 0
    $gv.Controls.Add($cv)

    # ===== ARCHITECTURE + LANGUAGE =====
    $y += 66
    $gl = New-Object System.Windows.Forms.GroupBox
    $gl.Text = " Architecture & Language "
    $gl.Location = New-Object System.Drawing.Point($M, $y)
    $gl.Size = New-Object System.Drawing.Size($GW, 80)
    $gl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($gl)

    # Radio buttons
    $ra1 = New-Object System.Windows.Forms.RadioButton
    $ra1.Name = "ra1"
    if ($script:is64Bit) { $ra1.Text = "64-bit (recommended)"; $ra1.Checked = $true } else { $ra1.Text = "64-bit (unavailable)"; $ra1.Enabled = $false }
    $ra1.Location = New-Object System.Drawing.Point(14, 24)
    $ra1.Size = New-Object System.Drawing.Size(160, 24)
    $ra1.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gl.Controls.Add($ra1)

    $ra2 = New-Object System.Windows.Forms.RadioButton
    $ra2.Name = "ra2"
    if ($script:is64Bit) { $ra2.Text = "32-bit" } else { $ra2.Text = "32-bit (recommended)"; $ra2.Checked = $true }
    $ra2.Location = New-Object System.Drawing.Point(14, 50)
    $ra2.Size = New-Object System.Drawing.Size(160, 24)
    $ra2.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gl.Controls.Add($ra2)

    # Language
    $ll = New-Object System.Windows.Forms.Label
    $ll.Text = "Language:"
    $ll.Location = New-Object System.Drawing.Point(260, 26)
    $ll.Size = New-Object System.Drawing.Size(80, 24)
    $ll.TextAlign = "MiddleLeft"
    $ll.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gl.Controls.Add($ll)

    $cl = New-Object System.Windows.Forms.ComboBox
    $cl.Name = "cl"
    $cl.Location = New-Object System.Drawing.Point(340, 24)
    $cl.Size = New-Object System.Drawing.Size(($GW - 360), 26)
    $cl.DropDownStyle = "DropDownList"
    $cl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cl.Items.AddRange(@("English (en-US)", "Spanish (es-ES)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)", "Italian (it-IT)", "Dutch (nl-NL)", "Polish (pl-PL)", "Russian (ru-RU)", "Japanese (ja-JP)"))
    $cl.SelectedIndex = 0
    $gl.Controls.Add($cl)

    # ===== ADDITIONAL PRODUCTS =====
    $y += 90
    $gp = New-Object System.Windows.Forms.GroupBox
    $gp.Text = " Additional Products "
    $gp.Location = New-Object System.Drawing.Point($M, $y)
    $gp.Size = New-Object System.Drawing.Size($GW, 76)
    $gp.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($gp)

    $cp1 = New-Object System.Windows.Forms.CheckBox
    $cp1.Name = "cp1"; $cp1.Text = "Include Project"; $cp1.Location = New-Object System.Drawing.Point(14, 22); $cp1.Size = New-Object System.Drawing.Size(130, 26); $cp1.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gp.Controls.Add($cp1)

    $cc1 = New-Object System.Windows.Forms.ComboBox
    $cc1.Name = "cc1"; $cc1.Location = New-Object System.Drawing.Point(160, 22); $cc1.Size = New-Object System.Drawing.Size(140, 26); $cc1.DropDownStyle = "DropDownList"; $cc1.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cc1.Items.AddRange(@("Professional", "Standard")); $cc1.SelectedIndex = 0; $cc1.Enabled = $false
    $gp.Controls.Add($cc1)
    $cp1.Add_CheckedChanged({ $cc1.Enabled = $cp1.Checked })

    $cp2 = New-Object System.Windows.Forms.CheckBox
    $cp2.Name = "cp2"; $cp2.Text = "Include Visio"; $cp2.Location = New-Object System.Drawing.Point(14, 46); $cp2.Size = New-Object System.Drawing.Size(130, 26); $cp2.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gp.Controls.Add($cp2)

    $cc2 = New-Object System.Windows.Forms.ComboBox
    $cc2.Name = "cc2"; $cc2.Location = New-Object System.Drawing.Point(160, 46); $cc2.Size = New-Object System.Drawing.Size(140, 26); $cc2.DropDownStyle = "DropDownList"; $cc2.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cc2.Items.AddRange(@("Professional", "Standard")); $cc2.SelectedIndex = 0; $cc2.Enabled = $false
    $gp.Controls.Add($cc2)
    $cp2.Add_CheckedChanged({ $cc2.Enabled = $cp2.Checked })

    # ===== APPLICATIONS =====
    $y += 86
    $ga = New-Object System.Windows.Forms.GroupBox
    $ga.Text = " Applications to Install "
    $ga.Location = New-Object System.Drawing.Point($M, $y)
    $ga.Size = New-Object System.Drawing.Size($GW, 120)
    $ga.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($ga)

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Name = "csa"; $csa.Text = "Select All"; $csa.Location = New-Object System.Drawing.Point(14, 22); $csa.Size = New-Object System.Drawing.Size(110, 24); $csa.Checked = $true; $csa.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $ga.Controls.Add($csa)

    $ac = New-Object System.Collections.ArrayList
    $spc = [Math]::Floor(($GW - 40) / 4)
    $ad = @(@{I="Word";X=14;Y=50;C=$true},@{I="Excel";X=14+$spc;Y=50;C=$true},@{I="PowerPoint";X=14+$spc*2;Y=50;C=$true},@{I="Outlook";X=14+$spc*3;Y=50;C=$true},@{I="Access";X=14;Y=78;C=$true},@{I="Publisher";X=14+$spc;Y=78;C=$true},@{I="OneNote";X=14+$spc*2;Y=78;C=$true},@{I="SkypeForBusiness";X=14+$spc*3;Y=78;C=$false})
    foreach ($a in $ad) {
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Name = "c_$($a.I)"; $c.Text = $a.I; $c.Location = New-Object System.Drawing.Point($a.X, $a.Y); $c.Size = New-Object System.Drawing.Size(($spc - 10), 24); $c.Checked = $a.C; $c.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $ga.Controls.Add($c); [void]$ac.Add($c)
        $c.Add_CheckedChanged({ if (-not $updatingAll) { $ca = $true; foreach ($b in $ac) { if (-not $b.Checked) { $ca = $false; break } }; $updatingAll = $true; $csa.Checked = $ca; $updatingAll = $false } })
    }
    $updatingAll = $false
    $csa.Add_CheckedChanged({ if (-not $updatingAll) { $updatingAll = $true; foreach ($b in $ac) { $b.Checked = $csa.Checked }; $updatingAll = $false } })

    # ===== INSTALLATION MODE =====
    $y += 130
    $gm = New-Object System.Windows.Forms.GroupBox
    $gm.Text = " Installation Mode "
    $gm.Location = New-Object System.Drawing.Point($M, $y)
    $gm.Size = New-Object System.Drawing.Size($GW, 80)
    $gm.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($gm)

    $gap = [Math]::Floor($GW / 3)
    $rm1 = New-Object System.Windows.Forms.RadioButton
    $rm1.Name = "rm1"; $rm1.Text = "Download and Install (recommended)"; $rm1.Location = New-Object System.Drawing.Point(14, 22); $rm1.Size = New-Object System.Drawing.Size(($gap - 20), 28); $rm1.Checked = $true; $rm1.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gm.Controls.Add($rm1)

    $rm2 = New-Object System.Windows.Forms.RadioButton
    $rm2.Name = "rm2"; $rm2.Text = "Download Only"; $rm2.Location = New-Object System.Drawing.Point(($gap + 10), 22); $rm2.Size = New-Object System.Drawing.Size(($gap - 20), 28); $rm2.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gm.Controls.Add($rm2)

    $rm3 = New-Object System.Windows.Forms.RadioButton
    $rm3.Name = "rm3"; $rm3.Text = "Install from Cache"; $rm3.Location = New-Object System.Drawing.Point(($gap * 2 + 10), 22); $rm3.Size = New-Object System.Drawing.Size(($gap - 20), 28); $rm3.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $gm.Controls.Add($rm3)

    # ===== STATUS BAR =====
    $y += 90
    $sb = New-Object System.Windows.Forms.Label
    $sb.Name = "sb"
    $sb.Text = "Status: Ready"
    $sb.Location = New-Object System.Drawing.Point($M, $y)
    $sb.Size = New-Object System.Drawing.Size($GW, 30)
    $sb.BorderStyle = "FixedSingle"
    $sb.TextAlign = "MiddleCenter"
    $sb.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($sb)

    # ===== BUTTONS =====
    $y += 40
    $btnW = 170
    $btnGap = 20
    $btnStart = [Math]::Floor(($FW - $btnW * 2 - $btnGap) / 2)
    $bi = New-Object System.Windows.Forms.Button
    $bi.Name = "bi"; $bi.Text = "Install Office"; $bi.Location = New-Object System.Drawing.Point($btnStart, $y); $bi.Size = New-Object System.Drawing.Size($btnW, 42)
    $bi.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $bi.FlatStyle = "Flat"
    $f.Controls.Add($bi)

    $bc = New-Object System.Windows.Forms.Button
    $bc.Name = "bc"; $bc.Text = "Cancel"; $bc.Location = New-Object System.Drawing.Point(($btnStart + $btnW + $btnGap), $y); $bc.Size = New-Object System.Drawing.Size($btnW, 42)
    $bc.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $bc.FlatStyle = "Flat"
    $f.Controls.Add($bc)

    # ===== VERSION MAP =====
    $vm = @(@{C="PerpetualVL2024";P="ProPlus2024Volume";VP="VisioPro2024Volume";VS="VisioStd2024Volume";PP="ProjectPro2024Volume";PS="ProjectStd2024Volume"},@{C="PerpetualVL2021";P="ProPlus2021Volume";VP="VisioPro2021Volume";VS="VisioStd2021Volume";PP="ProjectPro2021Volume";PS="ProjectStd2021Volume"},@{C="PerpetualVL2019";P="ProPlus2019Volume";VP="VisioPro2019Volume";VS="VisioStd2019Volume";PP="ProjectPro2019Volume";PS="ProjectStd2019Volume"},@{C="PerpetualVL2016";P="ProPlus2016Volume";VP="VisioPro2016Volume";VS="VisioStd2016Volume";PP="ProjectPro2016Volume";PS="ProjectStd2016Volume"},@{C="PerpetualVL2013";P="ProPlus2013Volume";VP="VisioPro2013Volume";VS="VisioStd2013Volume";PP="ProjectPro2013Volume";PS="ProjectStd2013Volume"})

    # ===== EVENTS =====
    $bc.Add_Click({ $f.Close() })

    $bi.Add_Click({
        $bi.Enabled = $false; $bc.Enabled = $false; $f.Cursor = "WaitCursor"
        $sb.Text = "Preparing configuration..."; $f.Refresh()
        try {
            $vi = $vm[$cv.SelectedIndex]
            $arch = if ($ra1.Checked) { "64" } else { "32" }
            $lang = Get-LangCode $cl.SelectedItem.ToString()
            $incP = $cp1.Checked; $pe = if ($incP) { if ($cc1.SelectedItem -eq "Standard") { "Std" } else { "Pro" } } else { $null }
            $incV = $cp2.Checked; $ve = if ($incV) { if ($cc2.SelectedItem -eq "Standard") { "Std" } else { "Pro" } } else { $null }
            $dm = "1"; if ($rm2.Checked) { $dm = "2" }; if ($rm3.Checked) { $dm = "3" }

            $sa = @(); foreach ($b in $ac) { if ($b.Checked) { $sa += $b.Text } }

            $sb.Text = "Generating XML configuration..."; $f.Refresh()

            $x = New-Object System.Text.StringBuilder
            [void]$x.AppendLine('<Configuration>')
            [void]$x.AppendLine("    <Add OfficeClientEdition=`"$arch`" Channel=`"$($vi.C)`">")
            [void]$x.AppendLine("        <Product ID=`"$($vi.P)`">")
            [void]$x.AppendLine("            <Language ID=`"$lang`" />")
            foreach ($app in @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")) { if ($sa -notcontains $app) { [void]$x.AppendLine("            <ExcludeApp ID=`"$app`" />") } }
            foreach ($ex in @("Bing","Groove","Lync","OneDrive","Teams")) { [void]$x.AppendLine("            <ExcludeApp ID=`"$ex`" />") }
            [void]$x.AppendLine("        </Product>")
            if ($incP) { $pid = if ($pe -eq "Std") { $vi.PS } else { $vi.PP }; [void]$x.AppendLine("        <Product ID=`"$pid`"><Language ID=`"$lang`" /></Product>") }
            if ($incV) { $vid = if ($ve -eq "Std") { $vi.VS } else { $vi.VP }; [void]$x.AppendLine("        <Product ID=`"$vid`"><Language ID=`"$lang`" /></Product>") }
            [void]$x.AppendLine("    </Add>")
            if ($dm -eq "2") { [void]$x.AppendLine('    <Display Level="None" AcceptEULA="TRUE" />'); [void]$x.AppendLine("    <Download Path=`"$script:odtTemp`" />") }
            elseif ($dm -eq "3") { $sp = Join-Path $script:odtTemp "Office"; if (-not (Test-Path $sp)) { throw "Cache folder not found. Use Download Only first." }; [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />') }
            else { [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />') }
            [void]$x.AppendLine('</Configuration>')

            $cp = Join-Path $script:odtTemp "configuration.xml"
            Set-Content -Path $cp -Value $x.ToString() -Encoding UTF8
            Write-Log "Config saved: $cp" "Green"

            $se = Join-Path $script:odtTemp "setup.exe"
            $arg = if ($dm -eq "2") { "/download" } else { "/configure" }
            $sb.Text = if ($dm -eq "2") { "Downloading Office..." } else { "Installing Office..." }
            $f.Refresh()

            $proc = Start-Process -FilePath $se -ArgumentList "$arg `"$cp`"" -Wait -PassThru
            if ($proc.ExitCode -eq 0) {
                $sb.Text = "Done."; $sb.ForeColor = "Green"
                Write-Log "Success." "Green"
                [System.Windows.Forms.MessageBox]::Show("Operation completed successfully.", "MSOI - Success", "OK", "Information")
            } else {
                $sb.Text = "Error (code: $($proc.ExitCode))."; $sb.ForeColor = "Red"
                Write-Log "Failed. Exit code: $($proc.ExitCode)" "Red"
                [System.Windows.Forms.MessageBox]::Show("Operation failed (code: $($proc.ExitCode)).`n`nLogs: $script:odtTemp", "MSOI - Error", "OK", "Error")
            }
        } catch {
            $sb.Text = "Error: $_"; $sb.ForeColor = "Red"
            Write-Log "Error: $_" "Red"
            [System.Windows.Forms.MessageBox]::Show("$_", "MSOI - Error", "OK", "Error")
        } finally {
            $bi.Enabled = $true; $bc.Enabled = $true; $f.Cursor = "Default"; $f.Refresh()
        }
    })

    [void]$f.ShowDialog()
}

# ====================================================
Write-Log "=== MSOI started ===" "Cyan"
Write-Log "System: $((Get-CimInstance Win32_OperatingSystem).Caption)" "Gray"
Write-Log "OS: $(if ($script:is64Bit) { '64' } else { '32' })-bit" "Gray"

Show-PrepDialog

if (-not $script:odtReady) {
    [System.Windows.Forms.MessageBox]::Show("Office Deployment Tool could not be prepared.", "MSOI - Error", "OK", "Error")
    exit 1
}

Show-MainForm
Write-Log "=== MSOI finished ===" "Cyan"

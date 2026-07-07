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
$script:iconPng = if ($PSScriptRoot) { Join-Path $PSScriptRoot "icono.png" } else { "" }
$script:iconUrl = "https://aldagou.github.io/MSOI/icono.png"

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
    $f.Size = New-Object System.Drawing.Size(460, 200)
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
    $l2.Location = New-Object System.Drawing.Point($m, 58)
    $l2.Size = New-Object System.Drawing.Size($w, 20)
    $f.Controls.Add($l2)

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Name = "pb"
    $pb.Location = New-Object System.Drawing.Point($m, 90)
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
    $M = 24
    $FW = 800
    $GW = $FW - 2 * $M
    $y = 0

    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "MSOIGUI - AldaGou"
    $f.ClientSize = New-Object System.Drawing.Size($FW, 500)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedSingle"
    $f.MaximizeBox = $false
    $f.BackColor = "White"
    try {
        if ($script:iconPng -and (Test-Path $script:iconPng)) { $bmp = [Drawing.Bitmap]::FromFile($script:iconPng) }
        else { $wc = New-Object Net.WebClient; $data = $wc.DownloadData($script:iconUrl); $ms = [IO.MemoryStream]::new($data, $false); $fs = $ms; $bmp = [Drawing.Bitmap]::FromStream($ms) }
        $f.Icon = [Drawing.Icon]::FromHandle($bmp.GetHicon())
    } catch {
        $iconPaths = @("$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE", "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE", "$env:ProgramFiles\Microsoft Office 15\root\Office15\WINWORD.EXE", "${env:ProgramFiles(x86)}\Microsoft Office 15\root\Office15\WINWORD.EXE")
        foreach ($p in $iconPaths) { if (Test-Path $p) { try { $f.Icon = [Drawing.Icon]::ExtractAssociatedIcon($p); break } catch {} } }
        if (-not $f.Icon) { try { $f.Icon = [Drawing.Icon]::ExtractAssociatedIcon("$env:SystemRoot\explorer.exe") } catch {} }
    } finally { if ($bmp) { $bmp.Dispose() }; if ($ms) { $ms.Dispose() } }
    try { $f.DoubleBuffered = $true } catch {}

    # ===== TITLE =====
    $y += 20
    $t = New-Object System.Windows.Forms.Label
    $t.Text = "MSOIGUI - AldaGou"
    $t.Location = New-Object System.Drawing.Point($M, $y)
    $t.Size = New-Object System.Drawing.Size($GW, 32)
    $t.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $t.TextAlign = "MiddleCenter"
    $f.Controls.Add($t)

    # ===== VERSION =====
    $y += 50
    $gv = New-Object System.Windows.Forms.GroupBox
    $gv.Text = " Office Version "
    $gv.Location = New-Object System.Drawing.Point($M, $y)
    $gv.Size = New-Object System.Drawing.Size($GW, 72)
    $gv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($gv)

    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Name = "cv"
    $cv.Location = New-Object System.Drawing.Point(16, 30)
    $cv.Size = New-Object System.Drawing.Size(($GW - 32), 26)
    $cv.DropDownStyle = "DropDownList"
    $cv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cv.Items.AddRange(@("Office LTSC Professional Plus 2024", "Office LTSC Professional Plus 2021", "Office Professional Plus 2019", "Office Professional Plus 2016", "Office Professional Plus 2013"))
    $cv.SelectedIndex = 0
    $gv.Controls.Add($cv)

    # ===== LANGUAGE =====
    $y += 82
    $gl = New-Object System.Windows.Forms.GroupBox
    $gl.Text = " Language "
    $gl.Location = New-Object System.Drawing.Point($M, $y)
    $gl.Size = New-Object System.Drawing.Size($GW, 72)
    $gl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($gl)

    $cl = New-Object System.Windows.Forms.ComboBox
    $cl.Name = "cl"
    $cl.Location = New-Object System.Drawing.Point(16, 30)
    $cl.Size = New-Object System.Drawing.Size(($GW - 32), 26)
    $cl.DropDownStyle = "DropDownList"
    $cl.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cl.Items.AddRange(@("English (en-US)", "Spanish (es-ES)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)", "Italian (it-IT)", "Dutch (nl-NL)", "Polish (pl-PL)", "Russian (ru-RU)", "Japanese (ja-JP)"))
    $cl.SelectedIndex = 0
    $gl.Controls.Add($cl)

    # ===== APPLICATIONS =====
    $y += 82
    $ga = New-Object System.Windows.Forms.GroupBox
    $ga.Text = " Applications "
    $ga.Location = New-Object System.Drawing.Point($M, $y)
    $ga.Size = New-Object System.Drawing.Size($GW, 148)
    $ga.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($ga)

    $gap = 10; $pd = 16; $ac = New-Object System.Collections.ArrayList
    $spc = [Math]::Floor(($GW - 2 * $pd - 3 * $gap) / 4) + $gap
    $bw = $spc - $gap

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Name = "csa"; $csa.Text = "Select All";     $csa.Location = New-Object System.Drawing.Point($pd, 30); $csa.Size = New-Object System.Drawing.Size(110, 24); $csa.Checked = $true; $csa.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $ga.Controls.Add($csa)

    $ad = @(@{I="Word";D="Word";C=$true},@{I="Excel";D="Excel";C=$true},@{I="PowerPoint";D="PowerPoint";C=$true},@{I="Outlook";D="Outlook";C=$false},@{I="Access";D="Access";C=$false},@{I="Publisher";D="Publisher";C=$false},@{I="OneNote";D="OneNote";C=$false},@{I="SkypeForBusiness";D="Skype for Business";C=$false},@{I="Project";D="Project";C=$false},@{I="Visio";D="Visio";C=$false})
    for ($i = 0; $i -lt $ad.Count; $i++) {
        $a = $ad[$i]; $col = $i % 4; $row = [Math]::Floor($i / 4)
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Name = "c_$($a.I)"; $c.Text = $a.D; $c.Tag = $a.I
        $c.Location = New-Object System.Drawing.Point(($pd + $col * $spc), (58 + $row * 28))
        $c.Size = New-Object System.Drawing.Size($bw, 24)
        $c.Checked = $a.C; $c.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $ga.Controls.Add($c); [void]$ac.Add($c)
        $c.Add_CheckedChanged({ if (-not $updatingAll) { $ca = $true; foreach ($b in $ac) { if (-not $b.Checked) { $ca = $false; break } }; $updatingAll = $true; $csa.Checked = $ca; $updatingAll = $false } })
    }
    $updatingAll = $false
    $csa.Add_CheckedChanged({ if (-not $updatingAll) { $updatingAll = $true; foreach ($b in $ac) { $b.Checked = $csa.Checked }; $updatingAll = $false } })

    # ===== STATUS BAR =====
    $y += 158
    $sb = New-Object System.Windows.Forms.Label
    $sb.Name = "sb"
    $sb.Text = "Status: Ready"
    $sb.Location = New-Object System.Drawing.Point($M, $y)
    $sb.Size = New-Object System.Drawing.Size($GW, 28)
    $sb.BorderStyle = "FixedSingle"
    $sb.TextAlign = "MiddleCenter"
    $sb.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $f.Controls.Add($sb)

    # ===== BUTTONS =====
    $y += 44
    $btnW = 170
    $btnGap = 20
    $btnStart = [Math]::Floor(($FW - $btnW * 2 - $btnGap) / 2)
    $bi = New-Object System.Windows.Forms.Button
    $bi.Name = "bi"; $bi.Text = "Install Office"; $bi.Location = New-Object System.Drawing.Point($btnStart, $y); $bi.Size = New-Object System.Drawing.Size($btnW, 42)
    $bi.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $bi.FlatStyle = "Standard"
    $f.Controls.Add($bi)

    $bc = New-Object System.Windows.Forms.Button
    $bc.Name = "bc"; $bc.Text = "Cancel"; $bc.Location = New-Object System.Drawing.Point(($btnStart + $btnW + $btnGap), $y); $bc.Size = New-Object System.Drawing.Size($btnW, 42)
    $bc.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $bc.FlatStyle = "Standard"
    $f.Controls.Add($bc)

    # ===== BYLINE =====
    $by = New-Object System.Windows.Forms.Label
    $by.Text = "By: AldaGou"
    $by.Location = New-Object System.Drawing.Point(($FW - 110), ($y + 42))
    $by.Size = New-Object System.Drawing.Size(90, 18)
    $by.TextAlign = "MiddleRight"
    $by.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $by.ForeColor = "DarkGray"
    $f.Controls.Add($by)

    # ===== VERSION MAP =====
    $vm = @(@{C="PerpetualVL2024";P="ProPlus2024Volume";V="VisioPro2024Volume";J="ProjectPro2024Volume"},@{C="PerpetualVL2021";P="ProPlus2021Volume";V="VisioPro2021Volume";J="ProjectPro2021Volume"},@{C="PerpetualVL2019";P="ProPlus2019Volume";V="VisioPro2019Volume";J="ProjectPro2019Volume"},@{C="PerpetualVL2016";P="ProPlus2016Volume";V="VisioPro2016Volume";J="ProjectPro2016Volume"},@{C="PerpetualVL2013";P="ProPlus2013Volume";V="VisioPro2013Volume";J="ProjectPro2013Volume"})

    # ===== EVENTS =====
    $bc.Add_Click({ $f.Close() })

    $bi.Add_Click({
        $bi.Enabled = $false; $bc.Enabled = $false; $f.Cursor = "WaitCursor"
        $sb.Text = "Preparing configuration..."; $f.Refresh()
        try {
            $vi = $vm[$cv.SelectedIndex]
            $arch = if ($script:is64Bit) { "64" } else { "32" }
            $lang = Get-LangCode $cl.SelectedItem.ToString()
            $incP = ($ac | Where-Object { $_.Tag -eq "Project" }).Checked
            $incV = ($ac | Where-Object { $_.Tag -eq "Visio" }).Checked
            $sa = @(); foreach ($b in $ac) { if ($b.Checked) { $sa += $b.Tag } }

            $sb.Text = "Generating XML configuration..."; $f.Refresh()

            $x = New-Object System.Text.StringBuilder
            [void]$x.AppendLine('<Configuration>')
            [void]$x.AppendLine("    <Add OfficeClientEdition=`"$arch`" Channel=`"$($vi.C)`">")
            [void]$x.AppendLine("        <Product ID=`"$($vi.P)`">")
            [void]$x.AppendLine("            <Language ID=`"$lang`" />")
            foreach ($app in @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")) { if ($sa -notcontains $app) { [void]$x.AppendLine("            <ExcludeApp ID=`"$app`" />") } }
            foreach ($ex in @("Bing","Groove","Lync","OneDrive","Teams")) { [void]$x.AppendLine("            <ExcludeApp ID=`"$ex`" />") }
            [void]$x.AppendLine("        </Product>")
            if ($incP) { [void]$x.AppendLine("        <Product ID=`"$($vi.J)`"><Language ID=`"$lang`" /></Product>") }
            if ($incV) { [void]$x.AppendLine("        <Product ID=`"$($vi.V)`"><Language ID=`"$lang`" /></Product>") }
            [void]$x.AppendLine("    </Add>")
            [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />')
            [void]$x.AppendLine('</Configuration>')

            $cp = Join-Path $script:odtTemp "configuration.xml"
            Set-Content -Path $cp -Value $x.ToString() -Encoding UTF8
            Write-Log "Config saved: $cp" "Green"

            $se = Join-Path $script:odtTemp "setup.exe"
            $sb.Text = "Downloading and installing Office..."; $f.Refresh()

            $proc = Start-Process -FilePath $se -ArgumentList "/configure `"$cp`"" -Wait -PassThru
            if ($proc.ExitCode -eq 0) {
                $sb.Text = "Cleaning up..."; $f.Refresh()
                if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
                if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
                $sb.Text = "Done."; $sb.ForeColor = "Green"
                Write-Log "Success." "Green"
                [System.Windows.Forms.MessageBox]::Show("Office installed successfully.", "MSOI - Success", "OK", "Information")
            } else {
                $sb.Text = "Error (code: $($proc.ExitCode))."; $sb.ForeColor = "Red"
                Write-Log "Failed. Exit code: $($proc.ExitCode)" "Red"
                [System.Windows.Forms.MessageBox]::Show("Installation failed (code: $($proc.ExitCode)).", "MSOI - Error", "OK", "Error")
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

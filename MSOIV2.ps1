<#
.SYNOPSIS
    MSOI - Microsoft Office Installation Tool (GUI)
.DESCRIPTION
    Graphical tool to download and install Microsoft Office LTSC.
    Supports multiple editions, architectures, languages and app selection.
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+
    Usage: irm https://aldagou.github.io/MSOI/MSOI.ps1 | iex
#>

#Requires -RunAsAdministrator

# ---- LOAD ASSEMBLIES ----
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---- DPI AWARENESS (prevent blurry text on high-DPI displays) ----
Add-Type -TypeDefinition "using System; using System.Runtime.InteropServices; public class _Dpi { [DllImport(\"user32.dll\")] public static extern bool SetProcessDPIAware(); [DllImport(\"shcore.dll\")] public static extern int SetProcessDpiAwareness(int a); }"
try { [void][_Dpi]::SetProcessDpiAwareness(1) } catch { try { [void][_Dpi]::SetProcessDPIAware() } catch {} }

# ---- CONSTANTS ----
$script:logFile = Join-Path $env:Temp "MSOI_Install.log"
$script:odtTemp = Join-Path $env:Temp "ODT"
$script:odtExe  = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtUrls = @("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
$script:odtReady = $false
$script:is64Bit = [Environment]::Is64BitOperatingSystem

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
    $f.Size = New-Object System.Drawing.Size(440, 170)
    $f.FormBorderStyle = "FixedDialog"
    $f.ControlBox = $false
    $f.StartPosition = "CenterScreen"
    $f.TopMost = $true

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "Preparing Office Deployment Tool..."
    $l1.Location = New-Object System.Drawing.Point(20, 25)
    $l1.Size = New-Object System.Drawing.Size(400, 22)
    $l1.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l1)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Name = "s"
    $l2.Text = "Starting..."
    $l2.Location = New-Object System.Drawing.Point(20, 52)
    $l2.Size = New-Object System.Drawing.Size(400, 18)
    $f.Controls.Add($l2)

    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Name = "pb"
    $pb.Location = New-Object System.Drawing.Point(20, 80)
    $pb.Size = New-Object System.Drawing.Size(400, 25)
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
            try { (New-Object System.Net.WebClient).DownloadFile($u, $script:odtExe); $ok = $true; break } catch { Set-Status "Retrying..." }
        }
        if (-not $ok) { throw "Could not download ODT from any source." }

        Set-Status "Extracting Office Deployment Tool..."
        Start-Sleep -Milliseconds 200
        $p = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$($script:odtTemp)`"" -Wait -PassThru
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
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.Form
    $f.AutoScaleMode = "Dpi"
    $f.Text = "MSOI - Microsoft Office Installer"
    $f.Size = New-Object System.Drawing.Size(700, 620)
    $f.StartPosition = "CenterScreen"
    $f.FormBorderStyle = "FixedSingle"
    $f.MaximizeBox = $false
    try { $f.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) } catch {}

    # Title
    $t = New-Object System.Windows.Forms.Label
    $t.Text = "MSOI - Microsoft Office Installer"
    $t.Location = New-Object System.Drawing.Point(18, 15)
    $t.Size = New-Object System.Drawing.Size(664, 28)
    $t.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($t)

    # ===== VERSION =====
    $gv = New-Object System.Windows.Forms.GroupBox
    $gv.Text = "Office Version"
    $gv.Location = New-Object System.Drawing.Point(12, 52)
    $gv.Size = New-Object System.Drawing.Size(676, 50)
    $f.Controls.Add($gv)

    $cv = New-Object System.Windows.Forms.ComboBox
    $cv.Name = "cv"
    $cv.Location = New-Object System.Drawing.Point(12, 20)
    $cv.Size = New-Object System.Drawing.Size(650, 23)
    $cv.DropDownStyle = "DropDownList"
    $cv.Items.AddRange(@("Office LTSC Professional Plus 2024", "Office LTSC Standard 2024", "Office LTSC Professional Plus 2021", "Office LTSC Standard 2021", "Office Professional Plus 2019", "Office Standard 2019"))
    $cv.SelectedIndex = 0
    $gv.Controls.Add($cv)

    # ===== ARCHITECTURE & LANGUAGE =====
    $ga = New-Object System.Windows.Forms.GroupBox
    $ga.Text = "Architecture & Language"
    $ga.Location = New-Object System.Drawing.Point(12, 110)
    $ga.Size = New-Object System.Drawing.Size(676, 65)
    $f.Controls.Add($ga)

    if ($script:is64Bit) {
        $ra1 = New-Object System.Windows.Forms.RadioButton
        $ra1.Name = "ra1"; $ra1.Text = "64-bit (recommended)"; $ra1.Location = New-Object System.Drawing.Point(12, 20); $ra1.Size = New-Object System.Drawing.Size(170, 20); $ra1.Checked = $true
        $ga.Controls.Add($ra1)

        $ra2 = New-Object System.Windows.Forms.RadioButton
        $ra2.Name = "ra2"; $ra2.Text = "32-bit"; $ra2.Location = New-Object System.Drawing.Point(12, 40); $ra2.Size = New-Object System.Drawing.Size(170, 20)
        $ga.Controls.Add($ra2)
    } else {
        $ra1 = New-Object System.Windows.Forms.RadioButton
        $ra1.Name = "ra1"; $ra1.Text = "64-bit (not available)"; $ra1.Location = New-Object System.Drawing.Point(12, 20); $ra1.Size = New-Object System.Drawing.Size(170, 20); $ra1.Enabled = $false
        $ga.Controls.Add($ra1)

        $ra2 = New-Object System.Windows.Forms.RadioButton
        $ra2.Name = "ra2"; $ra2.Text = "32-bit (recommended)"; $ra2.Location = New-Object System.Drawing.Point(12, 40); $ra2.Size = New-Object System.Drawing.Size(170, 20); $ra2.Checked = $true
        $ga.Controls.Add($ra2)
    }

    $ll = New-Object System.Windows.Forms.Label
    $ll.Text = "Language:"
    $ll.Location = New-Object System.Drawing.Point(230, 22)
    $ll.Size = New-Object System.Drawing.Size(65, 20)
    $ga.Controls.Add($ll)

    $cl = New-Object System.Windows.Forms.ComboBox
    $cl.Name = "cl"
    $cl.Location = New-Object System.Drawing.Point(295, 20)
    $cl.Size = New-Object System.Drawing.Size(365, 23)
    $cl.DropDownStyle = "DropDownList"
    $cl.Items.AddRange(@("English (en-US)", "Spanish (es-ES)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)", "Italian (it-IT)", "Dutch (nl-NL)", "Polish (pl-PL)", "Russian (ru-RU)", "Japanese (ja-JP)"))
    $cl.SelectedIndex = 0
    $ga.Controls.Add($cl)

    # ===== ADDITIONAL PRODUCTS =====
    $gp = New-Object System.Windows.Forms.GroupBox
    $gp.Text = "Additional Products"
    $gp.Location = New-Object System.Drawing.Point(12, 183)
    $gp.Size = New-Object System.Drawing.Size(676, 65)
    $f.Controls.Add($gp)

    $cp1 = New-Object System.Windows.Forms.CheckBox
    $cp1.Name = "cp1"; $cp1.Text = "Include Project"; $cp1.Location = New-Object System.Drawing.Point(12, 20); $cp1.Size = New-Object System.Drawing.Size(120, 22)
    $gp.Controls.Add($cp1)

    $cc1 = New-Object System.Windows.Forms.ComboBox
    $cc1.Name = "cc1"; $cc1.Location = New-Object System.Drawing.Point(138, 20); $cc1.Size = New-Object System.Drawing.Size(130, 23); $cc1.DropDownStyle = "DropDownList"
    $cc1.Items.AddRange(@("Professional", "Standard")); $cc1.SelectedIndex = 0; $cc1.Enabled = $false
    $gp.Controls.Add($cc1)

    $cp1.Add_CheckedChanged({ $cc1.Enabled = $cp1.Checked })

    $cp2 = New-Object System.Windows.Forms.CheckBox
    $cp2.Name = "cp2"; $cp2.Text = "Include Visio"; $cp2.Location = New-Object System.Drawing.Point(12, 42); $cp2.Size = New-Object System.Drawing.Size(120, 22)
    $gp.Controls.Add($cp2)

    $cc2 = New-Object System.Windows.Forms.ComboBox
    $cc2.Name = "cc2"; $cc2.Location = New-Object System.Drawing.Point(138, 42); $cc2.Size = New-Object System.Drawing.Size(130, 23); $cc2.DropDownStyle = "DropDownList"
    $cc2.Items.AddRange(@("Professional", "Standard")); $cc2.SelectedIndex = 0; $cc2.Enabled = $false
    $gp.Controls.Add($cc2)

    $cp2.Add_CheckedChanged({ $cc2.Enabled = $cp2.Checked })

    # ===== APPLICATIONS =====
    $ga2 = New-Object System.Windows.Forms.GroupBox
    $ga2.Text = "Applications to Install"
    $ga2.Location = New-Object System.Drawing.Point(12, 256)
    $ga2.Size = New-Object System.Drawing.Size(676, 105)
    $f.Controls.Add($ga2)

    $csa = New-Object System.Windows.Forms.CheckBox
    $csa.Name = "csa"; $csa.Text = "Select All"; $csa.Location = New-Object System.Drawing.Point(12, 20); $csa.Size = New-Object System.Drawing.Size(100, 20); $csa.Checked = $true
    $ga2.Controls.Add($csa)

    $ac = New-Object System.Collections.ArrayList
    $ad = @(@{I="Word";X=12;Y=44;C=$true},@{I="Excel";X=175;Y=44;C=$true},@{I="PowerPoint";X=338;Y=44;C=$true},@{I="Outlook";X=500;Y=44;C=$true},@{I="Access";X=12;Y=68;C=$true},@{I="Publisher";X=175;Y=68;C=$true},@{I="OneNote";X=338;Y=68;C=$true},@{I="SkypeForBusiness";X=500;Y=68;C=$false})
    foreach ($a in $ad) {
        $c = New-Object System.Windows.Forms.CheckBox
        $c.Name = "c_$($a.I)"; $c.Text = $a.I; $c.Location = New-Object System.Drawing.Point($a.X, $a.Y); $c.Size = New-Object System.Drawing.Size(140, 20); $c.Checked = $a.C
        $ga2.Controls.Add($c); [void]$ac.Add($c)
        $c.Add_CheckedChanged({ $ca = $true; foreach ($b in $ac) { if (-not $b.Checked) { $ca = $false; break } }; $csa.Checked = $ca })
    }
    $csa.Add_CheckedChanged({ foreach ($b in $ac) { $b.Checked = $csa.Checked } })

    # ===== INSTALLATION MODE =====
    $gm = New-Object System.Windows.Forms.GroupBox
    $gm.Text = "Installation Mode"
    $gm.Location = New-Object System.Drawing.Point(12, 369)
    $gm.Size = New-Object System.Drawing.Size(676, 65)
    $f.Controls.Add($gm)

    $rm1 = New-Object System.Windows.Forms.RadioButton
    $rm1.Name = "rm1"; $rm1.Text = "Download and Install (recommended)"; $rm1.Location = New-Object System.Drawing.Point(12, 20); $rm1.Size = New-Object System.Drawing.Size(300, 20); $rm1.Checked = $true
    $gm.Controls.Add($rm1)

    $rm2 = New-Object System.Windows.Forms.RadioButton
    $rm2.Name = "rm2"; $rm2.Text = "Download Only"; $rm2.Location = New-Object System.Drawing.Point(12, 42); $rm2.Size = New-Object System.Drawing.Size(300, 20)
    $gm.Controls.Add($rm2)

    $rm3 = New-Object System.Windows.Forms.RadioButton
    $rm3.Name = "rm3"; $rm3.Text = "Install from Cache"; $rm3.Location = New-Object System.Drawing.Point(340, 20); $rm3.Size = New-Object System.Drawing.Size(310, 20)
    $gm.Controls.Add($rm3)

    # ===== STATUS =====
    $sb = New-Object System.Windows.Forms.Label
    $sb.Name = "sb"
    $sb.Text = "Status: Ready"
    $sb.Location = New-Object System.Drawing.Point(12, 445)
    $sb.Size = New-Object System.Drawing.Size(676, 26)
    $sb.BorderStyle = "FixedSingle"
    $f.Controls.Add($sb)

    # ===== BUTTONS =====
    $bi = New-Object System.Windows.Forms.Button
    $bi.Name = "bi"; $bi.Text = "Install Office"; $bi.Location = New-Object System.Drawing.Point(12, 485); $bi.Size = New-Object System.Drawing.Size(160, 38)
    $bi.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($bi)

    $bc = New-Object System.Windows.Forms.Button
    $bc.Name = "bc"; $bc.Text = "Cancel"; $bc.Location = New-Object System.Drawing.Point(180, 485); $bc.Size = New-Object System.Drawing.Size(90, 38)
    $f.Controls.Add($bc)

    # ===== VERSION MAP =====
    $vm = @(@{C="PerpetualVL2024";P="ProPlus2024Volume";VP="VisioPro2024Volume";VS="VisioStd2024Volume";PP="ProjectPro2024Volume";PS="ProjectStd2024Volume"},@{C="PerpetualVL2024";P="Standard2024Volume";VP="VisioPro2024Volume";VS="VisioStd2024Volume";PP="ProjectPro2024Volume";PS="ProjectStd2024Volume"},@{C="PerpetualVL2021";P="ProPlus2021Volume";VP="VisioPro2021Volume";VS="VisioStd2021Volume";PP="ProjectPro2021Volume";PS="ProjectStd2021Volume"},@{C="PerpetualVL2021";P="Standard2021Volume";VP="VisioPro2021Volume";VS="VisioStd2021Volume";PP="ProjectPro2021Volume";PS="ProjectStd2021Volume"},@{C="PerpetualVL2019";P="ProPlus2019Volume";VP="VisioPro2019Volume";VS="VisioStd2019Volume";PP="ProjectPro2019Volume";PS="ProjectStd2019Volume"},@{C="PerpetualVL2019";P="Standard2019Volume";VP="VisioPro2019Volume";VS="VisioStd2019Volume";PP="ProjectPro2019Volume";PS="ProjectStd2019Volume"})

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
            [void]$x.AppendLine("<Configuration>")
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
            [void]$x.AppendLine("</Configuration>")

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
Write-Log "=== MSOI GUI started ===" "Cyan"
Write-Log "System: $((Get-CimInstance Win32_OperatingSystem).Caption)" "Gray"
Write-Log "OS: $(if ($script:is64Bit) { '64' } else { '32' })-bit" "Gray"

Show-PrepDialog

if (-not $script:odtReady) {
    [System.Windows.Forms.MessageBox]::Show("Office Deployment Tool could not be prepared.", "MSOI - Error", "OK", "Error")
    exit 1
}

Show-MainForm
Write-Log "=== MSOI GUI finished ===" "Cyan"

<#
.SYNOPSIS
    MSOICLI - Microsoft Office Installer (Command Line)
.DESCRIPTION
    Command-line tool to download and install Microsoft Office LTSC.
    Supports multiple editions, architectures, languages and app selection.
.NOTES
    Requirements: Administrator, PowerShell 5.0+, .NET Framework 4.5+
    Usage: irm https://aldagou.github.io/MSOI/MSOICLI.ps1 | iex
#>

# ---- ADMIN CHECK ----
$script:isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $script:isAdmin) { Write-Host "  [!] This script requires Administrator privileges." -ForegroundColor Red; Read-Host "Press Enter to exit"; exit 1 }

# ---- CONSTANTS ----
$script:logFile = Join-Path $env:Temp "MSOICLI_Install.log"
$script:odtTemp = Join-Path $env:Temp "ODT"
$script:odtExe  = Join-Path $env:Temp "OfficeDeploymentTool.exe"
$script:odtUrls = @("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
$script:is64Bit = [Environment]::Is64BitOperatingSystem
$script:startAt = Get-Date

# ---- COLOR HELPERS ----
function Color {
    param([string]$T, [string]$C = "White", [string]$B = "")
    if ($B) { Write-Host $T -ForegroundColor $C -BackgroundColor $B -NoNewline }
    else    { Write-Host $T -ForegroundColor $C -NoNewline }
}

function Line  { param([string]$C = "DarkGray")    Color "  $("-" * 72)" $C; Write-Host "" }
function Hr    { param([string]$C = "DarkGray")    Color "  $("-" * 72)" $C; Write-Host "" }

# ---- LOGGING ----
function Write-Log {
    param([string]$M)
    Add-Content -Path $script:logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $M" -ErrorAction SilentlyContinue
}

# ---- BANNER ----
function Show-Banner {
    Clear-Host
    Write-Host ""
    $w = 66
    $hl = "+" + ("=" * $w) + "+"
    $fmt = "{0,-" + $w + "}"
    Color "  $hl" "Cyan"; Write-Host ""
    Color "  |" "Cyan"; Color ($fmt -f "         MSOI - Microsoft Office Installer          ") "White"; Color "|" "Cyan"; Write-Host ""
    Color "  |" "Cyan"; Color ($fmt -f "              Command Line Edition                   ") "DarkGray"; Color "|" "Cyan"; Write-Host ""
    Color "  $hl" "Cyan"; Write-Host ""
    Write-Host ""
}

# ---- PROGRESS ----
function Show-ProgressBar {
    param([int]$Pct, [string]$Label = "Progress", [int]$W = 50)
    $f = [Math]::Floor($W * $Pct / 100)
    $e = $W - $f
    $bar = ("#" * $f) + ("." * $e)
    Color "  $Label " "DarkGray"; Color "[$bar] " "Cyan"; Color "$Pct%" "Yellow"
    if ($Pct -eq 100) { Write-Host "" } else { Write-Host "`r" -NoNewline }
}

# ---- ODT PREPARATION ----
function Step-PrepareODT {
    Color "  >> " "Yellow"; Color "Preparing Office Deployment Tool..." "White"; Write-Host ""
    Write-Log "Starting ODT preparation"
    Write-Host ""

    if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Path $script:odtTemp -Force | Out-Null

    $ok = $false
    foreach ($u in $script:odtUrls) {
        Color "    * Downloading ODT..." "DarkGray"
        Write-Progress -Activity "Downloading Office Deployment Tool" -Status "Connecting..." -PercentComplete -1
        try {
            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($u, $script:odtExe)
            $ok = $true
            Write-Progress -Activity "Downloading Office Deployment Tool" -Completed
            Color "  [OK]" "Green"; Write-Host ""
            break
        } catch {
            Write-Progress -Activity "Downloading Office Deployment Tool" -Completed
            Color "  [FAILED]" "Red"; Write-Host "    $_" -ForegroundColor DarkGray
        }
    }
    if (-not $ok) { throw "Could not download ODT from any source." }

    Color "    * Extracting ODT..." "DarkGray"
    Write-Progress -Activity "Extracting Office Deployment Tool" -Status "Please wait..." -PercentComplete -1
    $p = Start-Process -FilePath $script:odtExe -ArgumentList "/quiet /extract:`"$script:odtTemp`"" -Wait -PassThru
    Write-Progress -Activity "Extracting Office Deployment Tool" -Completed
    if ($p.ExitCode -ne 0) { throw "Extraction failed (code: $($p.ExitCode))." }
    if (-not (Test-Path (Join-Path $script:odtTemp "setup.exe"))) { throw "setup.exe not found after extraction." }
    Color "  [OK]" "Green"; Write-Host ""

    Write-Host ""
    Color "  >> " "Yellow"; Color "ODT ready" "Green"; Write-Host ""
    Write-Host ""
}

# ---- VERSION SELECTION ----
function Step-Version {
    $versions = @(
        "Office LTSC Professional Plus 2024",
        "Office LTSC Professional Plus 2021",
        "Office Professional Plus 2019",
        "Office Professional Plus 2016",
        "Office Professional Plus 2013"
    )
    $vm = @(
        @{C="PerpetualVL2024";P="ProPlus2024Volume";V="VisioPro2024Volume";J="ProjectPro2024Volume"}
        @{C="PerpetualVL2021";P="ProPlus2021Volume";V="VisioPro2021Volume";J="ProjectPro2021Volume"}
        @{C="PerpetualVL2019";P="ProPlus2019Volume";V="VisioPro2019Volume";J="ProjectPro2019Volume"}
        @{C="PerpetualVL2016";P="ProPlus2016Volume";V="VisioPro2016Volume";J="ProjectPro2016Volume"}
        @{C="PerpetualVL2013";P="ProPlus2013Volume";V="VisioPro2013Volume";J="ProjectPro2013Volume"}
    )

    Color "  >> " "Yellow"; Color "Select Office Version" "White"; Write-Host ""
    Hr
    for ($i = 0; $i -lt $versions.Count; $i++) {
        Color "    $($i+1). " "DarkGray"; Color "$($versions[$i])" "White"; Write-Host ""
    }
    Hr
    $choice = Read-Host "  Enter choice (1-$($versions.Count)) [1]"
    if ([string]::IsNullOrWhiteSpace($choice)) { $choice = "1" }
    if ($choice -notmatch '^[1-5]$') { Color "    Invalid, using default (1)" "Red"; $choice = "1"; Write-Host "" }
    $idx = [int]$choice - 1
    Write-Host ""

    return @{
        Label   = $versions[$idx]
        Index   = $idx
        Channel = $vm[$idx].C
        PID     = $vm[$idx].P
        VisioID  = $vm[$idx].V
        ProjID   = $vm[$idx].J
    }
}

# ---- LANGUAGE ----
function Step-Lang {
    $langs = @(
        "English (en-US)", "Spanish (es-ES)", "French (fr-FR)", "German (de-DE)",
        "Brazilian Portuguese (pt-BR)", "Italian (it-IT)", "Dutch (nl-NL)",
        "Polish (pl-PL)", "Russian (ru-RU)", "Japanese (ja-JP)"
    )

    Color "  >> " "Yellow"; Color "Select Language" "White"; Write-Host ""
    Hr
    for ($i = 0; $i -lt $langs.Count; $i++) {
        $name = ($langs[$i] -split ' \(')[0]
        $code = ($langs[$i] -split '\(')[1] -replace '\)',''
        Color "    $($i+1). " "DarkGray"; Color "$name " "White"; Color "($code)" "DarkGray"; Write-Host ""
    }
    Hr
    $c = Read-Host "  Enter choice (1-$($langs.Count)) [1]"
    if ([string]::IsNullOrWhiteSpace($c)) { $c = "1" }
    if ($c -notmatch '^[1-9]$|^10$') { Color "    Invalid, using English" "Red"; $c = "1"; Write-Host "" }
    Write-Host ""

    $idx = [int]$c - 1
    $code = if ($langs[$idx] -match '\(([^)]+)\)') { $matches[1] } else { "en-US" }
    return @{ Label = $langs[$idx]; Code = $code }
}

# ---- ADDITIONAL PRODUCTS ----
function Step-Products {
    Color "  >> " "Yellow"; Color "Additional Products" "White"; Write-Host ""
    Hr

    $proj = Read-Host "  Include Microsoft Project? (Y/N) [N]"
    $projInclude = $proj -eq "Y" -or $proj -eq "y"

    $vis = Read-Host "  Include Microsoft Visio? (Y/N) [N]"
    $visInclude = $vis -eq "Y" -or $vis -eq "y"

    Write-Host ""
    return @{
        Project = $projInclude
        Visio   = $visInclude
    }
}

# ---- APPLICATIONS ----
function Step-Apps {
    $all = @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")

    Color "  >> " "Yellow"; Color "Applications to Install" "White"; Write-Host ""
    Hr
    Color "    All applications are selected by default." "DarkGray"; Write-Host ""
    Color "    Enter numbers to EXCLUDE (separated by commas)." "DarkGray"; Write-Host ""
    Write-Host ""
    for ($i = 0; $i -lt $all.Count; $i++) {
        $mark = if ($all[$i] -eq "SkypeForBusiness") { " [excluded by default]" } else { "" }
        Color "    $($i+1). " "DarkGray"; Color "$($all[$i])" "White"
        if ($mark) { Color "$mark" "DarkGray" }
        Write-Host ""
    }
    Write-Host ""
    Color "    Example: " "DarkGray"; Color "4,8" "Yellow"; Color " to skip Outlook and Skype" "DarkGray"; Write-Host ""
    Hr

    $excl = Read-Host "  Numbers to exclude [Enter for all]"
    $selected = @($all)  # start with all

    if (-not [string]::IsNullOrWhiteSpace($excl)) {
        $nums = $excl -split ',' | ForEach-Object { $_.Trim() }
        foreach ($n in $nums) {
            $idx = $n -as [int]
            if ($idx -ge 1 -and $idx -le $all.Count) {
                $selected = $selected | Where-Object { $_ -ne $all[$idx - 1] }
            }
        }
    }
    Write-Host ""
    return $selected
}

# ---- SUMMARY ----
function Show-Summary {
    param(
        $Version, $Arch, $Lang, $Products, $Apps
    )

    $appsStr = if ($Apps.Count -gt 5) {
        ($Apps[0..4] -join ', ') + ", +$($Apps.Count-5) more"
    } else {
        $Apps -join ', '
    }

    Write-Host ""
    $sw = 66
    $hl = "+" + ("=" * $sw) + "+"
    $fmt = "{0,-" + $sw + "}"
    $paddedTitle = $fmt -f "INSTALLATION SUMMARY"
    Color "  $hl" "Cyan"; Write-Host ""
    Color "  |$paddedTitle|" "Cyan"; Write-Host ""
    Color "  $hl" "Cyan"; Write-Host ""

    $projStr = if ($Products.Project) { "Yes" } else { "No" }
    $visStr  = if ($Products.Visio)   { "Yes" } else { "No" }
    $rows = @(
        @("Office Version", $Version.Label),
        @("Architecture",   "$Arch-bit"),
        @("Language",       $Lang.Label),
        @("Project",        $projStr),
        @("Visio",          $visStr),
        @("Applications",   $appsStr)
    )

    $labelW = 17
    $valueW = $sw - $labelW - 3
    $lfmt = "{0,-" + $labelW + "}"
    $vfmt = "{0,-" + $valueW + "}"
    foreach ($r in $rows) {
        $labelPart = $lfmt -f $r[0]
        $valuePart = $vfmt -f $r[1]
        Color "  |" "Cyan"
        Color (" " + $labelPart + ": " + $valuePart) "White"
        Color "|" "Cyan"; Write-Host ""
    }
    Color "  $hl" "Cyan"; Write-Host ""
}

# ---- XML GENERATION ----
function Build-Config {
    param($Version, $Arch, $Lang, $Products, $Apps)

    $x = New-Object System.Text.StringBuilder
    [void]$x.AppendLine('<Configuration>')
    [void]$x.AppendLine("    <Add OfficeClientEdition=`"$Arch`" Channel=`"$($Version.Channel)`">")
    [void]$x.AppendLine("        <Product ID=`"$($Version.PID)`">")
    [void]$x.AppendLine("            <Language ID=`"$($Lang.Code)`" />")

    foreach ($app in @("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote","SkypeForBusiness")) {
        if ($Apps -notcontains $app) { [void]$x.AppendLine("            <ExcludeApp ID=`"$app`" />") }
    }
    foreach ($ex in @("Bing","Groove","Lync","OneDrive","Teams")) { [void]$x.AppendLine("            <ExcludeApp ID=`"$ex`" />") }
    [void]$x.AppendLine("        </Product>")

    if ($Products.Project) { [void]$x.AppendLine("        <Product ID=`"$($Version.ProjID)`"><Language ID=`"$($Lang.Code)`" /></Product>") }
    if ($Products.Visio)   { [void]$x.AppendLine("        <Product ID=`"$($Version.VisioID)`"><Language ID=`"$($Lang.Code)`" /></Product>") }

    [void]$x.AppendLine("    </Add>")
    [void]$x.AppendLine('    <Display Level="Full" AcceptEULA="TRUE" />')
    [void]$x.AppendLine('</Configuration>')
    return $x.ToString()
}

# ---- INSTALLATION ----
function Step-Install {
    param($ConfigXml)

    $cp = Join-Path $script:odtTemp "configuration.xml"
    Set-Content -Path $cp -Value $ConfigXml -Encoding UTF8
    Write-Log "Config saved: $cp"

    $se = Join-Path $script:odtTemp "setup.exe"

    Color "  >> " "Yellow"; Color "Downloading and installing Office..." "White"; Write-Host ""
    Write-Host ""

    Write-Progress -Activity "Installing" -Status "Running ODT setup.exe..." -PercentComplete -1

    $proc = Start-Process -FilePath $se -ArgumentList "/configure `"$cp`"" -Wait -PassThru -NoNewWindow:$false
    Write-Progress -Activity "Installing" -Completed

    Write-Host ""
    if ($proc.ExitCode -eq 0) {
        Color "  >> " "Yellow"; Color "Cleaning up temporary files..." "White"; Write-Host ""
        if (Test-Path $script:odtTemp) { Remove-Item $script:odtTemp -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $script:odtExe)  { Remove-Item $script:odtExe -Force -ErrorAction SilentlyContinue }
        Color "  >> " "Yellow"; Color "Done" "Green"; Write-Host ""
        Color "  >> " "Yellow"; Color "Office installed successfully." "Green"; Write-Host ""
        Write-Log "Success"
    } else {
        Color "  >> " "Yellow"; Color "Failed (code: $($proc.ExitCode))" "Red"; Write-Host ""
        Color "      Check logs at: $script:odtTemp" "DarkGray"; Write-Host ""
        Write-Log "Failed with code $($proc.ExitCode)"
    }

    $elapsed = [math]::Round(((Get-Date) - $script:startAt).TotalSeconds, 1)
    Write-Host ""
    Color "  >> " "Yellow"; Color "Total time: " "White"; Color "$elapsed seconds" "DarkGray"; Write-Host ""
}

# ---- PAUSE ----
function Pause-Message {
    param([string]$M = "Press Enter to exit...")
    Write-Host ""
    Color "  $M" "DarkGray"
    $null = Read-Host
}

# ====================================================
# MAIN
# ====================================================
$ErrorActionPreference = "Stop"

Show-Banner

# Detect system
Color "  >> " "Yellow"; Color "System" "White"; Write-Host ""
Hr
Color "    OS: " "DarkGray"
if ($script:is64Bit) { Color "64-bit" "Green" } else { Color "32-bit" "Yellow" }
Write-Host ""

$psVer = $PSVersionTable.PSVersion.ToString()
Color "    PowerShell: " "DarkGray"; Color $psVer "White"; Write-Host ""
Write-Host ""
Write-Log "MSOCI CLI started. OS: $(if ($script:is64Bit) {'64'} else {'32'})-bit, PS: $psVer"

# Step 1: Prepare ODT
try { Step-PrepareODT } catch {
    Color "  [!] " "Red"; Color "Failed to prepare ODT: $_" "Red"; Write-Host ""
    Color "      Download manually: https://www.microsoft.com/en-us/download/details.aspx?id=49117" "DarkGray"; Write-Host ""
    Pause-Message; exit 1
}

# Step 2: Version
$version = Step-Version

# Step 3: Language
$lang = Step-Lang

# Step 4: Products
$products = Step-Products

# Step 5: Apps
$apps = Step-Apps

# Summary
$arch = if ($script:is64Bit) { "64" } else { "32" }
Show-Summary -Version $version -Arch $arch -Lang $lang -Products $products -Apps $apps

# Confirm
$confirm = Read-Host "  Proceed? (Y/N) [Y]"
if ($confirm -eq "N" -or $confirm -eq "n") {
    Color "  [-] Cancelled by user." "Yellow"; Write-Host ""
    Pause-Message; exit 0
}

Write-Host ""

# Install
$xml = Build-Config -Version $version -Arch $arch -Lang $lang -Products $products -Apps $apps
try {
    Step-Install -ConfigXml $xml
} catch {
    Color "  [!] " "Red"; Color "Installation error: $_" "Red"; Write-Host ""
    Write-Log "Fatal error: $_"
}

Pause-Message

#requires -Version 5.1
<#
  Office LTSC Installer - WPF (inspirado en WinUtil)
  - GUI en XAML
  - Autoelevación
  - Descarga ODT (aka.ms/ODT) con fallback
  - Genera configuration.xml y lanza instalación
#>

#region Elevación
function Ensure-Admin {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) {
    if ($PSCommandPath) {
      Start-Process -FilePath "powershell.exe" -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    } else {
      Start-Process -FilePath "powershell.exe" -Verb RunAs
    }
    exit
  }
}
Ensure-Admin
#endregion

#region Utilidades
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'Stop'
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls13
} catch { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 }

$Log = Join-Path $env:TEMP ("OfficeLTSC_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
function Log($m){ ("{0} {1}" -f (Get-Date -f "yyyy-MM-dd HH:mm:ss"), $m) | Add-Content $Log }

function Get-ProductMap {
  @{
    "LTSC 2024" = @{ Channel="PerpetualVL2024"; Office="ProPlus2024Volume"; Visio="VisioPro2024Volume"; Project="ProjectPro2024Volume" }
    "LTSC 2021" = @{ Channel="PerpetualVL2021"; Office="ProPlus2021Volume"; Visio="VisioPro2021Volume"; Project="ProjectPro2021Volume" }
    "LTSC 2019" = @{ Channel="PerpetualVL2019"; Office="ProPlus2019Volume"; Visio="VisioPro2019Volume"; Project="ProjectPro2019Volume" }
  }
}

function New-OfficeConfigXml {
  param(
    [ValidateSet("32","64")]$Edition="64",
    [Parameter(Mandatory)]$Channel,
    [Parameter(Mandatory)]$ProductId,
    [Parameter(Mandatory)]$Language,
    [string[]]$IncludeApps=@("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote"),
    [switch]$IncludeVisio,[switch]$IncludeProject,[string]$VisioId,[string]$ProjectId,
    [switch]$RemoveMSI,[ValidateSet("None","Basic","Full")]$DisplayLevel="Full"
  )
  $known=@("Access","Excel","OneNote","Outlook","PowerPoint","Publisher","Teams","Word","OneDrive","Groove","Lync")
  $all=@("Word","Excel","PowerPoint","Outlook","Access","Publisher","OneNote")
  $exclude = $all | Where-Object { $_ -notin $IncludeApps }

  $sb = [System.Text.StringBuilder]::new()
  $null=$sb.AppendLine('<Configuration>')
  $null=$sb.AppendLine("  <Add OfficeClientEdition=""$Edition"" Channel=""$Channel"">")
  $null=$sb.AppendLine("    <Product ID=""$ProductId"">")
  $null=$sb.AppendLine("      <Language ID=""$Language"" />")
  foreach($a in $exclude){ if($known -contains $a){ $null=$sb.AppendLine("      <ExcludeApp ID=""$a"" />") } }
  foreach($extra in @("Teams","OneDrive","Groove","Lync")){
    if($extra -notin $IncludeApps -and $known -contains $extra){
      $null=$sb.AppendLine("      <ExcludeApp ID=""$extra"" />")
    }
  }
  $null=$sb.AppendLine("    </Product>")
  if($IncludeProject -and $ProjectId){
    $null=$sb.AppendLine("    <Product ID=""$ProjectId""><Language ID=""$Language"" /></Product>")
  }
  if($IncludeVisio -and $VisioId){
    $null=$sb.AppendLine("    <Product ID=""$VisioId""><Language ID=""$Language"" /></Product>")
  }
  $null=$sb.AppendLine("  </Add>")
  if($RemoveMSI){ $null=$sb.AppendLine("  <RemoveMSI />") }
  $null=$sb.AppendLine("  <Display Level=""$DisplayLevel"" AcceptEULA=""TRUE"" />")
  $null=$sb.AppendLine('</Configuration>')
  $sb.ToString()
}

function Get-ODT {
  param([Parameter(Mandatory)][string]$ExtractTo)
  New-Item -ItemType Directory -Force -Path $ExtractTo | Out-Null
  $exe = Join-Path $env:TEMP "OfficeDeploymentTool.exe"
  $urls = @("https://aka.ms/ODT",
            "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe")
  foreach($u in $urls){
    try{ Invoke-WebRequest $u -OutFile $exe -TimeoutSec 180; Log "ODT descargado: $u"; break } catch { Log "Fallo ODT $u -> $($_.Exception.Message)" }
  }
  if(-not (Test-Path $exe)){ throw "No se pudo descargar el ODT." }
  Start-Process -FilePath $exe -ArgumentList "/quiet /extract:`"$ExtractTo`"" -Wait
  $setup = Join-Path $ExtractTo "setup.exe"
  if(-not (Test-Path $setup)){ throw "setup.exe no apareció tras extraer ODT." }
  $setup
}

function Start-ODTInstall { param($SetupExe,$ConfigXml)
  Log "Instalación: `"$SetupExe`" /configure `"$ConfigXml`""
  $p = Start-Process -FilePath $SetupExe -ArgumentList "/configure `"$ConfigXml`"" -PassThru
  $p.WaitForExit()
  $p.ExitCode
}
#endregion

#region UI XAML
Add-Type -AssemblyName PresentationCore,PresentationFramework

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office LTSC Installer (WPF)" Height="560" Width="760" WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <StackPanel Orientation="Horizontal" Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,8">
      <TextBlock Text="Versión:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <ComboBox x:Name="CmbVersion" Width="160" Margin="0,0,16,0">
        <ComboBoxItem Content="LTSC 2024" IsSelected="True"/>
        <ComboBoxItem Content="LTSC 2021"/>
        <ComboBoxItem Content="LTSC 2019"/>
      </ComboBox>
      <TextBlock Text="Idioma:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <ComboBox x:Name="CmbLang" Width="180">
        <ComboBoxItem Content="es-ES" IsSelected="True"/>
        <ComboBoxItem Content="es-MX"/>
        <ComboBoxItem Content="en-US"/>
        <ComboBoxItem Content="fr-FR"/>
        <ComboBoxItem Content="de-DE"/>
        <ComboBoxItem Content="pt-BR"/>
      </ComboBox>
      <StackPanel Orientation="Horizontal" Margin="16,0,0,0">
        <RadioButton x:Name="Rad64" Content="x64" IsChecked="True" Margin="0,0,8,0"/>
        <RadioButton x:Name="Rad32" Content="x86"/>
      </StackPanel>
    </StackPanel>

    <GroupBox Header="Aplicaciones" Grid.Row="1" Grid.Column="0" Margin="0,0,8,8">
      <WrapPanel Margin="8">
        <CheckBox x:Name="ChkWord" Content="Word" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkExcel" Content="Excel" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkPowerPoint" Content="PowerPoint" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkOutlook" Content="Outlook" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkAccess" Content="Access" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkPublisher" Content="Publisher" IsChecked="True" Margin="0,0,12,6"/>
        <CheckBox x:Name="ChkOneNote" Content="OneNote" IsChecked="True" Margin="0,0,12,6"/>
      </WrapPanel>
    </GroupBox>

    <GroupBox Header="Opciones" Grid.Row="1" Grid.Column="1" Margin="8,0,0,8">
      <StackPanel Margin="8">
        <CheckBox x:Name="ChkTeams" Content="Excluir Teams" IsChecked="True"/>
        <CheckBox x:Name="ChkOneDrive" Content="Excluir OneDrive" IsChecked="True"/>
        <CheckBox x:Name="ChkProject" Content="Incluir Project"/>
        <CheckBox x:Name="ChkVisio" Content="Incluir Visio"/>
        <CheckBox x:Name="ChkRemoveMSI" Content="Quitar Office MSI (<RemoveMSI/>)" IsChecked="True"/>
      </StackPanel>
    </GroupBox>

    <DockPanel Grid.Row="2" Grid.ColumnSpan="2" LastChildFill="True">
      <TextBlock Text="Carpeta ODT:" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBox x:Name="TxtExtract" Width="520" Text="" Margin="0,0,8,0"/>
      <Button x:Name="BtnBrowse" Content="Examinar…" Width="90"/>
    </DockPanel>

    <StackPanel Grid.Row="3" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,8" >
      <Button x:Name="BtnDownload" Content="1) Descargar ODT" Width="150" Margin="0,0,8,0"/>
      <Button x:Name="BtnGenXml"  Content="2) Generar XML"  Width="150" Margin="0,0,8,0"/>
      <Button x:Name="BtnInstall"  Content="3) Instalar"     Width="120"/>
    </StackPanel>

    <StackPanel Grid.Row="4" Grid.ColumnSpan="2">
      <ProgressBar x:Name="Bar" Height="16" Minimum="0" Maximum="100" Value="0"/>
      <TextBlock x:Name="Lbl" Text="Listo." Margin="0,6,0,0"/>
    </StackPanel>
  </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Referencias a controles
$CmbVersion = $window.FindName("CmbVersion")
$CmbLang    = $window.FindName("CmbLang")
$Rad64      = $window.FindName("Rad64")
$Rad32      = $window.FindName("Rad32")
$ChkWord,$ChkExcel,$ChkPowerPoint,$ChkOutlook,$ChkAccess,$ChkPublisher,$ChkOneNote = `
  $window.FindName("ChkWord"),$window.FindName("ChkExcel"),$window.FindName("ChkPowerPoint"),$window.FindName("ChkOutlook"),$window.FindName("ChkAccess"),$window.FindName("ChkPublisher"),$window.FindName("ChkOneNote")
$ChkTeams   = $window.FindName("ChkTeams")
$ChkOneDrive= $window.FindName("ChkOneDrive")
$ChkProject = $window.FindName("ChkProject")
$ChkVisio   = $window.FindName("ChkVisio")
$ChkRemoveMSI = $window.FindName("ChkRemoveMSI")
$TxtExtract = $window.FindName("TxtExtract")
$BtnBrowse  = $window.FindName("BtnBrowse")
$BtnDownload= $window.FindName("BtnDownload")
$BtnGenXml  = $window.FindName("BtnGenXml")
$BtnInstall = $window.FindName("BtnInstall")
$Bar        = $window.FindName("Bar")
$Lbl        = $window.FindName("Lbl")

# Estado
$global:SetupExe = $null
$global:ConfigXml = $null
$TxtExtract.Text = (Join-Path $env:TEMP "ODT")

function Get-Selection {
  $version = ($CmbVersion.SelectedItem.Content).ToString()
  $lang    = ($CmbLang.SelectedItem.Content).ToString()
  $edition = if($Rad32.IsChecked){ "32" } else { "64" }
  $apps = @()
  foreach($pair in @(
      @($ChkWord,"Word"),@($ChkExcel,"Excel"),@($ChkPowerPoint,"PowerPoint"),
      @($ChkOutlook,"Outlook"),@($ChkAccess,"Access"),
      @($ChkPublisher,"Publisher"),@($ChkOneNote,"OneNote")
  )){
    if($pair[0].IsChecked){ $apps += $pair[1] }
  }
  [pscustomobject]@{
    Version = $version; Language=$lang; Edition=$edition; Apps=$apps
    ExcludeTeams=$ChkTeams.IsChecked; ExcludeOneDrive=$ChkOneDrive.IsChecked
    IncludeProject=$ChkProject.IsChecked; IncludeVisio=$ChkVisio.IsChecked
    RemoveMSI=$ChkRemoveMSI.IsChecked; ExtractTo=$TxtExtract.Text
  }
}

# Eventos
$BtnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.SelectedPath = $TxtExtract.Text
  if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){ $TxtExtract.Text = $dlg.SelectedPath }
})

$BtnDownload.Add_Click({
  try{
    $window.IsEnabled = $false; $Bar.IsIndeterminate=$true; $Lbl.Text="Descargando/Extrayendo ODT…"
    $global:SetupExe = Get-ODT -ExtractTo $TxtExtract.Text
    $Lbl.Text = "ODT listo: $global:SetupExe"; $Bar.IsIndeterminate=$false; $Bar.Value=100
  } catch {
    $Bar.IsIndeterminate=$false; $Bar.Value=0; $Lbl.Text="Error ODT: $($_.Exception.Message)"
  } finally { $window.IsEnabled = $true }
})

$BtnGenXml.Add_Click({
  try{
    $window.IsEnabled = $false; $Bar.IsIndeterminate=$true; $Lbl.Text="Generando configuration.xml…"
    $s = Get-Selection
    $map = Get-ProductMap
    $m = $map[$s.Version]
    $apps = $s.Apps
    if($s.ExcludeTeams){ $apps = $apps | Where-Object { $_ -ne "Teams" } }
    if($s.ExcludeOneDrive){ $apps = $apps | Where-Object { $_ -ne "OneDrive" } }

    $xml = New-OfficeConfigXml -Edition $s.Edition -Channel $m.Channel -ProductId $m.Office -Language $s.Language `
          -IncludeApps $apps -IncludeVisio:([bool]$s.IncludeVisio) -IncludeProject:([bool]$s.IncludeProject) `
          -VisioId $m.Visio -ProjectId $m.Project -RemoveMSI:([bool]$s.RemoveMSI) -DisplayLevel "Full"
    $global:ConfigXml = Join-Path $s.ExtractTo "configuration.xml"
    $xml | Set-Content -Path $global:ConfigXml -Encoding UTF8
    $Lbl.Text = "XML generado en: $global:ConfigXml"; $Bar.IsIndeterminate=$false; $Bar.Value=100
  } catch {
    $Bar.IsIndeterminate=$false; $Bar.Value=0; $Lbl.Text="Error XML: $($_.Exception.Message)"
  } finally { $window.IsEnabled = $true }
})

$BtnInstall.Add_Click({
  try{
    if(-not $global:SetupExe){ throw "Falta setup.exe (ODT). Pulsa 'Descargar ODT'." }
    if(-not $global:ConfigXml){ throw "Falta configuration.xml. Pulsa 'Generar XML'." }
    $window.IsEnabled = $false; $Bar.IsIndeterminate=$true; $Lbl.Text="Instalando Office…"
    $code = Start-ODTInstall -SetupExe $global:SetupExe -ConfigXml $global:ConfigXml
    $Lbl.Text = "Instalación finalizada. Código: $code"; $Bar.IsIndeterminate=$false; $Bar.Value=100
  } catch {
    $Bar.IsIndeterminate=$false; $Bar.Value=0; $Lbl.Text="Error instalación: $($_.Exception.Message)"
  } finally { $window.IsEnabled = $true }
})

# Muestra ventana
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$window.ShowDialog() | Out-Null
Write-Host "Log: $Log"

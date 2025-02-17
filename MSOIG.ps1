# Este script crea una interfaz gráfica para configurar e instalar Office LTSC usando PowerShell.

Add-Type -AssemblyName PresentationFramework

# Crear la ventana principal
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Office LTSC Installer" Height="450" Width="600" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <!-- Selección de versión -->
        <GroupBox Header="Select Office LTSC Version" Margin="0,10,0,0">
            <StackPanel>
                <RadioButton Name="Version2024" Content="Office LTSC 2024" GroupName="Version" IsChecked="True" />
                <RadioButton Name="Version2021" Content="Office LTSC 2021" GroupName="Version" />
                <RadioButton Name="Version2019" Content="Office LTSC 2019" GroupName="Version" />
            </StackPanel>
        </GroupBox>

        <!-- Selección de idioma -->
        <GroupBox Header="Select Language" Margin="0,100,0,0">
            <ComboBox Name="LanguageDropdown" SelectedIndex="0">
                <ComboBoxItem Content="Spanish (es-ES)" />
                <ComboBoxItem Content="English (en-US)" />
                <ComboBoxItem Content="French (fr-FR)" />
                <ComboBoxItem Content="German (de-DE)" />
                <ComboBoxItem Content="Brazilian Portuguese (pt-BR)" />
            </ComboBox>
        </GroupBox>

        <!-- Selección de productos -->
        <GroupBox Header="Select Additional Products" Margin="0,200,0,0">
            <StackPanel>
                <CheckBox Name="IncludeProject" Content="Include Project" />
                <CheckBox Name="IncludeVisio" Content="Include Visio" />
            </StackPanel>
        </GroupBox>

        <!-- Selección de aplicaciones -->
        <GroupBox Header="Select Applications" Grid.Row="1" Margin="0,10,0,0">
            <StackPanel>
                <CheckBox Name="AppWord" Content="Word" IsChecked="True" />
                <CheckBox Name="AppExcel" Content="Excel" IsChecked="True" />
                <CheckBox Name="AppPowerPoint" Content="PowerPoint" IsChecked="True" />
                <CheckBox Name="AppOutlook" Content="Outlook" IsChecked="True" />
                <CheckBox Name="AppAccess" Content="Access" />
                <CheckBox Name="AppPublisher" Content="Publisher" />
                <CheckBox Name="AppOneNote" Content="OneNote" IsChecked="True" />
            </StackPanel>
        </GroupBox>

        <!-- Botón de instalación -->
        <Button Name="InstallButton" Content="Install" Grid.Row="2" Margin="0,10,0,0" HorizontalAlignment="Center" Width="100" />

    </Grid>
</Window>
"@

# Cargar el XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Asignar eventos a los controles
$InstallButton = $window.FindName("InstallButton")
$LanguageDropdown = $window.FindName("LanguageDropdown")

$InstallButton.Add_Click({
    # Obtener la versión seleccionada
    $version = if ($window.Version2024.IsChecked) {
        "PerpetualVL2024"
    } elseif ($window.Version2021.IsChecked) {
        "PerpetualVL2021"
    } elseif ($window.Version2019.IsChecked) {
        "PerpetualVL2019"
    }

    # Obtener el idioma seleccionado
    $language = $LanguageDropdown.SelectedItem.Content -split " \(" | Select-Object -First 1

    # Verificar productos adicionales
    $includeProject = $window.IncludeProject.IsChecked
    $includeVisio = $window.IncludeVisio.IsChecked

    # Verificar aplicaciones seleccionadas
    $selectedApps = @()
    if ($window.AppWord.IsChecked) { $selectedApps += "Word" }
    if ($window.AppExcel.IsChecked) { $selectedApps += "Excel" }
    if ($window.AppPowerPoint.IsChecked) { $selectedApps += "PowerPoint" }
    if ($window.AppOutlook.IsChecked) { $selectedApps += "Outlook" }
    if ($window.AppAccess.IsChecked) { $selectedApps += "Access" }
    if ($window.AppPublisher.IsChecked) { $selectedApps += "Publisher" }
    if ($window.AppOneNote.IsChecked) { $selectedApps += "OneNote" }

    # Generar archivo de configuración
    $config = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="$version">
        <Product ID="$productID">
            <Language ID="$language" />
"@

foreach ($app in $apps) {
    if (-not ($selectedApps -contains $app)) {
        $config += "            <ExcludeApp ID=`"$app`" />`r`n"
    }
}

$config += @"
            <ExcludeApp ID="Bing" />
            <ExcludeApp ID="Groove" />
            <ExcludeApp ID="Lync" />
            <ExcludeApp ID="OneDrive" />
            <ExcludeApp ID="OutlookForWindows" />
            <ExcludeApp ID="Teams" />
        </Product>
"@

if ($includeProject -eq "Y") {
    $config += @"
        <Product ID="$projectID">
            <Language ID="$language" />
        </Product>
"@
}

if ($includeVisio -eq "Y") {
    $config += @"
        <Product ID="$visioID">
            <Language ID="$language" />
        </Product>
"@
}

$config += @"
    </Add>
    <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@


    # Guardar archivo de configuración
    $configPath = Join-Path $env:Temp "configuration.xml"
    Set-Content -Path $configPath -Value $config

    [System.Windows.MessageBox]::Show("Configuration file generated at: $configPath", "Success")

    # Aquí puedes añadir la lógica para iniciar la instalación
})

# Mostrar la ventana
$window.ShowDialog()

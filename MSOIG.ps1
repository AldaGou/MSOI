# Este script debe ejecutarse en PowerShell como Administrador.
# Descarga e instala Office LTSC basado en las opciones configuradas por el usuario.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Verifica si es administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    [System.Windows.Forms.MessageBox]::Show("Por favor, ejecuta este script como Administrador.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Crear formulario
$form = New-Object System.Windows.Forms.Form
$form.Text = "Instalador de Office LTSC"
$form.Size = New-Object System.Drawing.Size(400, 600)
$form.StartPosition = "CenterScreen"

# Crear controles
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "Selecciona la versión de Office LTSC:"
$labelVersion.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelVersion)

$comboVersion = New-Object System.Windows.Forms.ComboBox
$comboVersion.Items.AddRange(@("Office LTSC 2024", "Office LTSC 2021", "Office LTSC 2019"))
$comboVersion.Location = New-Object System.Drawing.Point(10, 50)
$form.Controls.Add($comboVersion)

$labelLanguage = New-Object System.Windows.Forms.Label
$labelLanguage.Text = "Selecciona el idioma de Office:"
$labelLanguage.Location = New-Object System.Drawing.Point(10, 90)
$form.Controls.Add($labelLanguage)

$comboLanguage = New-Object System.Windows.Forms.ComboBox
$comboLanguage.Items.AddRange(@("Spanish (es-ES)", "English (en-US)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)"))
$comboLanguage.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($comboLanguage)

$checkProject = New-Object System.Windows.Forms.CheckBox
$checkProject.Text = "Incluir Project"
$checkProject.Location = New-Object System.Drawing.Point(10, 160)
$form.Controls.Add($checkProject)

$checkVisio = New-Object System.Windows.Forms.CheckBox
$checkVisio.Text = "Incluir Visio"
$checkVisio.Location = New-Object System.Drawing.Point(10, 190)
$form.Controls.Add($checkVisio)

$labelApps = New-Object System.Windows.Forms.Label
$labelApps.Text = "Selecciona las aplicaciones a instalar:"
$labelApps.Location = New-Object System.Drawing.Point(10, 230)
$form.Controls.Add($labelApps)

$listBoxApps = New-Object System.Windows.Forms.CheckedListBox
$listBoxApps.Items.AddRange(@("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote"))
$listBoxApps.Location = New-Object System.Drawing.Point(10, 260)
$listBoxApps.Size = New-Object System.Drawing.Size(200, 100)
$form.Controls.Add($listBoxApps)

$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Text = "Instalar"
$buttonInstall.Location = New-Object System.Drawing.Point(10, 380)
$form.Controls.Add($buttonInstall)

# Función para manejar el clic del botón
$buttonInstall.Add_Click({
    $version = $comboVersion.SelectedItem
    $language = $comboLanguage.SelectedItem
    $includeProject = $checkProject.Checked
    $includeVisio = $checkVisio.Checked
    $selectedApps = @()
    foreach ($item in $listBoxApps.CheckedItems) {
        $selectedApps += $item
    }

    if (-not $version -or -not $language) {
        [System.Windows.Forms.MessageBox]::Show("Por favor, selecciona la versión y el idioma de Office.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Aquí puedes añadir la lógica para descargar y configurar Office LTSC basado en las selecciones del usuario.
    # Variables definidas según la versión seleccionada
    switch ($version) {
        "Office LTSC 2024" { 
            $version = "PerpetualVL2024"
            $productID = "ProPlus2024Volume"
            $visioID = "VisioPro2024Volume"
            $projectID = "ProjectPro2024Volume"
        }
        "Office LTSC 2021" { 
            $version = "PerpetualVL2021"
            $productID = "ProPlus2021Volume"
            $visioID = "VisioPro2021Volume"
            $projectID = "ProjectPro2021Volume"
        }
        "Office LTSC 2019" { 
            $version = "PerpetualVL2019"
            $productID = "ProPlus2019Volume"
            $visioID = "VisioPro2019Volume"
            $projectID = "ProjectPro2019Volume"
        }
    }

    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18227-20162.exe"
    $odtExe = "OfficeDeploymentTool.exe"
    $odtPath = Join-Path $env:Temp $odtExe
    
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtPath -UseBasicParsing -ErrorAction Stop -TimeoutSec 60
    Start-Process -FilePath $odtPath -ArgumentList "/quiet /extract:$env:Temp\ODT" -Wait

    $officeConfigPath = Join-Path $env:Temp\ODT "configuration.xml"
    
    $config = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="$version">
        <Product ID="$productID">
            <Language ID="$language" />
"@
    
    foreach ($app in $listBoxApps.Items) {
        if (-not ($selectedApps -contains $app)) {
            $config += "            <ExcludeApp ID=""$app"" />`n"
        }
    }
    
    $config += @"
        </Product>
"@
    
    if ($includeProject) {
        $config += @"
        <Product ID="$projectID">
            <Language ID="$language" />
        </Product>
"@
    }
    
    if ($includeVisio) {
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
    
    Set-Content -Path $officeConfigPath -Value $config
    
    Start-Process -FilePath "$env:Temp\ODT\setup.exe" -ArgumentList "/configure $officeConfigPath" -Wait
    [System.Windows.Forms.MessageBox]::Show("Instalación completada con éxito.", "Información", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Mostrar formulario
$form.ShowDialog()

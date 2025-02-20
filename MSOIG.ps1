# Este script agrega una interfaz gráfica (GUI) para configurar e instalar Office LTSC
# Ejecutar como Administrador en PowerShell

Add-Type -AssemblyName PresentationFramework

function Show-GUI {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Office LTSC Installer"
    $form.Size = New-Object System.Drawing.Size(400, 500)
    $form.StartPosition = "CenterScreen"

    # Etiqueta para selección de versión
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Select Office LTSC Version:"
    $versionLabel.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($versionLabel)

    # Dropdown para versión
    $versionComboBox = New-Object System.Windows.Forms.ComboBox
    $versionComboBox.Items.AddRange(@("Office LTSC 2024", "Office LTSC 2021", "Office LTSC 2019"))
    $versionComboBox.Location = New-Object System.Drawing.Point(10, 50)
    $versionComboBox.Width = 350
    $form.Controls.Add($versionComboBox)

    # Etiqueta para idioma
    $languageLabel = New-Object System.Windows.Forms.Label
    $languageLabel.Text = "Select Language:"
    $languageLabel.Location = New-Object System.Drawing.Point(10, 100)
    $form.Controls.Add($languageLabel)

    # Dropdown para idioma
    $languageComboBox = New-Object System.Windows.Forms.ComboBox
    $languageComboBox.Items.AddRange(@("Spanish (es-ES)", "English (en-US)", "French (fr-FR)", "German (de-DE)", "Brazilian Portuguese (pt-BR)"))
    $languageComboBox.Location = New-Object System.Drawing.Point(10, 130)
    $languageComboBox.Width = 350
    $form.Controls.Add($languageComboBox)

    # Checkboxes para Project y Visio
    $projectCheckbox = New-Object System.Windows.Forms.CheckBox
    $projectCheckbox.Text = "Include Project"
    $projectCheckbox.Location = New-Object System.Drawing.Point(10, 180)
    $form.Controls.Add($projectCheckbox)

    $visioCheckbox = New-Object System.Windows.Forms.CheckBox
    $visioCheckbox.Text = "Include Visio"
    $visioCheckbox.Location = New-Object System.Drawing.Point(10, 210)
    $form.Controls.Add($visioCheckbox)

    # Etiqueta para seleccionar apps
    $appsLabel = New-Object System.Windows.Forms.Label
    $appsLabel.Text = "Select Apps to Install:"
    $appsLabel.Location = New-Object System.Drawing.Point(10, 260)
    $form.Controls.Add($appsLabel)

    # ListBox para seleccionar apps
    $appsListBox = New-Object System.Windows.Forms.CheckedListBox
    $appsListBox.Items.AddRange(@("Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote"))
    $appsListBox.Location = New-Object System.Drawing.Point(10, 290)
    $appsListBox.Size = New-Object System.Drawing.Size(350, 100)
    $form.Controls.Add($appsListBox)

    # Botón para iniciar
    $startButton = New-Object System.Windows.Forms.Button
    $startButton.Text = "Start Installation"
    $startButton.Location = New-Object System.Drawing.Point(10, 420)
    $startButton.Width = 350
    $form.Controls.Add($startButton)

    # Acción del botón
    $startButton.Add_Click({
        $versionChoice = $versionComboBox.SelectedItem
        $languageChoice = $languageComboBox.SelectedItem
        $includeProject = $projectCheckbox.Checked
        $includeVisio = $visioCheckbox.Checked
        $selectedApps = $appsListBox.CheckedItems

        if (-not $versionChoice -or -not $languageChoice) {
            [System.Windows.Forms.MessageBox]::Show("Please select a version and language.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Mapeo de las selecciones
        $versionMap = @{ "Office LTSC 2024" = "PerpetualVL2024"; "Office LTSC 2021" = "PerpetualVL2021"; "Office LTSC 2019" = "PerpetualVL2019" }
        $productMap = @{ "Office LTSC 2024" = "ProPlus2024Volume"; "Office LTSC 2021" = "ProPlus2021Volume"; "Office LTSC 2019" = "ProPlus2019Volume" }
        $visioMap = @{ "Office LTSC 2024" = "VisioPro2024Volume"; "Office LTSC 2021" = "VisioPro2021Volume"; "Office LTSC 2019" = "VisioPro2019Volume" }
        $projectMap = @{ "Office LTSC 2024" = "ProjectPro2024Volume"; "Office LTSC 2021" = "ProjectPro2021Volume"; "Office LTSC 2019" = "ProjectPro2019Volume" }

        $version = $versionMap[$versionChoice]
        $productID = $productMap[$versionChoice]
        $visioID = $visioMap[$versionChoice]
        $projectID = $projectMap[$versionChoice]
        $language = ($languageChoice -split ' ')[1]

        # Generar archivo de configuración
        $config = "<Configuration>\n    <Add OfficeClientEdition=\"64\" Channel=\"$version\">\n        <Product ID=\"$productID\">\n            <Language ID=\"$language\" />\n"
        foreach ($app in "Word", "Excel", "PowerPoint", "Outlook", "Access", "Publisher", "OneNote") {
            if (-not $selectedApps.Contains($app)) {
                $config += "            <ExcludeApp ID=\"$app\" />\n"
            }
        }
        $config += "        </Product>\n"
        if ($includeProject) {
            $config += "        <Product ID=\"$projectID\">\n            <Language ID=\"$language\" />\n        </Product>\n"
        }
        if ($includeVisio) {
            $config += "        <Product ID=\"$visioID\">\n            <Language ID=\"$language\" />\n        </Product>\n"
        }
        $config += "    </Add>\n    <Display Level=\"Full\" AcceptEULA=\"TRUE\" />\n</Configuration>"

        $configPath = Join-Path $env:Temp "configuration.xml"
        Set-Content -Path $configPath -Value $config

        # Ejecutar instalación
        Start-Process -FilePath "$env:Temp\ODT\setup.exe" -ArgumentList "/configure $configPath" -Wait
        [System.Windows.Forms.MessageBox]::Show("Installation completed successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })

    $form.ShowDialog()
}

# Ejecutar GUI
Show-GUI

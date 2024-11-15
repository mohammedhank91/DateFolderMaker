Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Creation Tool"
$form.Size = New-Object System.Drawing.Size(500,450)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
# Load the icon file and set it to the form
$iconPath = "create_folders.ico"
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

# Create labels and text boxes for start date and end date
$labelStartDate = New-Object System.Windows.Forms.Label
$labelStartDate.Text = "Start Date (dd-MM-yy):"
$labelStartDate.Location = New-Object System.Drawing.Point(10,20)
$labelStartDate.Size = New-Object System.Drawing.Size(150,20)
$form.Controls.Add($labelStartDate)

# Create a DateTimePicker for Start Date
$dateTimePickerStart = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dateTimePickerStart.CustomFormat = "dd-MM-yy"
$dateTimePickerStart.Location = New-Object System.Drawing.Point(170,20)
$dateTimePickerStart.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($dateTimePickerStart)

$labelEndDate = New-Object System.Windows.Forms.Label
$labelEndDate.Text = "End Date (dd-MM-yy):"
$labelEndDate.Location = New-Object System.Drawing.Point(10,60)
$labelEndDate.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($labelEndDate)

# Create a DateTimePicker for End Date
$dateTimePickerEnd = New-Object System.Windows.Forms.DateTimePicker
$dateTimePickerEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dateTimePickerEnd.CustomFormat = "dd-MM-yy"
$dateTimePickerEnd.Location = New-Object System.Drawing.Point(170,60)
$dateTimePickerEnd.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($dateTimePickerEnd)

# Create labels and text boxes for subfolder names
$labelSubFolder1 = New-Object System.Windows.Forms.Label
$labelSubFolder1.Text = "Subfolder 1 Name:"
$labelSubFolder1.Location = New-Object System.Drawing.Point(10,100)
$labelSubFolder1.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($labelSubFolder1)

$textBoxSubFolder1 = New-Object System.Windows.Forms.TextBox
$textBoxSubFolder1.Location = New-Object System.Drawing.Point(170,100)
$textBoxSubFolder1.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxSubFolder1)

$labelSubFolder2 = New-Object System.Windows.Forms.Label
$labelSubFolder2.Text = "Subfolder 2 Name:"
$labelSubFolder2.Location = New-Object System.Drawing.Point(10,140)
$labelSubFolder2.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($labelSubFolder2)

$textBoxSubFolder2 = New-Object System.Windows.Forms.TextBox
$textBoxSubFolder2.Location = New-Object System.Drawing.Point(170,140)
$textBoxSubFolder2.Size = New-Object System.Drawing.Size(200,15)
$form.Controls.Add($textBoxSubFolder2)

# Create a button to select output directory
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse"
$buttonBrowse.Location = New-Object System.Drawing.Point(380,215)
$buttonBrowse.Size = New-Object System.Drawing.Size(75,25)
$form.Controls.Add($buttonBrowse)
$buttonBrowse.BackColor = [System.Drawing.Color]::FromArgb(60, 179, 113) # MediumSeaGreen
$buttonBrowse.ForeColor = [System.Drawing.Color]::White
$buttonBrowse.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(144, 238, 144) # LightGreen
$buttonBrowse.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(34, 139, 34) # ForestGreen


# Create a label to display the selected directory
$labelDirectory = New-Object System.Windows.Forms.Label
$labelDirectory.Text = "Output Directory:"
$labelDirectory.Location = New-Object System.Drawing.Point(10,220)
$labelDirectory.Size = New-Object System.Drawing.Size(130,20)
$form.Controls.Add($labelDirectory)

$textBoxDirectory = New-Object System.Windows.Forms.TextBox
$textBoxDirectory.Location = New-Object System.Drawing.Point(170,220)
$textBoxDirectory.Size = New-Object System.Drawing.Size(200,20)
$form.Controls.Add($textBoxDirectory)


# Add event to browse button
$buttonBrowse.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxDirectory.Text = $folderBrowser.SelectedPath
    }
})

# Button to clear the entered data
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "Clear Data"
$buttonClear.Location = New-Object System.Drawing.Point(380,340)
$buttonClear.Size = New-Object System.Drawing.Size(75,25)
$form.Controls.Add($buttonClear)

# Quit button
$buttonQuit = New-Object System.Windows.Forms.Button
$buttonQuit.Text = "Quit"
$buttonQuit.Location = New-Object System.Drawing.Point(380, 380) 
$buttonQuit.Size = New-Object System.Drawing.Size(75,25)
$buttonQuit.BackColor = [System.Drawing.Color]::Red  # Set background color to red
$buttonQuit.ForeColor = [System.Drawing.Color]::White  # Set text color to white for contrast
$form.Controls.Add($buttonQuit)

# Add Click event for Quit button to close the form
$buttonQuit.Add_Click({
    [System.Windows.Forms.Application]::Exit()
})

# Event to clear the entered data
$buttonClear.Add_Click({
    # Clear the DateTimePickers
    $dateTimePickerStart.Value = [System.DateTime]::Now
    $dateTimePickerEnd.Value = [System.DateTime]::Now

    # Clear the TextBoxes
    $textBoxSubFolder1.Clear()
    $textBoxSubFolder2.Clear()
    $textBoxDirectory.Clear()

})


# Create a progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10,270)
$progressBar.Size = New-Object System.Drawing.Size(450,30)
$form.Controls.Add($progressBar)

# Create a button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Create Folders"
$button.Location = New-Object System.Drawing.Point(150, 310)
$button.Size = New-Object System.Drawing.Size(150, 50)
$button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) # SteelBlue
$button.ForeColor = [System.Drawing.Color]::Azure
$button.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$button.FlatAppearance.BorderSize = 0
$button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 149, 237) # CornflowerBlue
$button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(65, 105, 225) # RoyalBlue

# Button to save template
$buttonSaveTemplate = New-Object System.Windows.Forms.Button
$buttonSaveTemplate.Text = "Save Template"
$buttonSaveTemplate.Location = New-Object System.Drawing.Point(130, 380)
$buttonSaveTemplate.Size = New-Object System.Drawing.Size(90,25)
$form.Controls.Add($buttonSaveTemplate)

# Event to save template data to JSON file
$buttonSaveTemplate.Add_Click({
    # Create SaveFileDialog to choose where to save the file
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON Files (*.json)|*.json"
    $saveFileDialog.DefaultExt = "json"
    $saveFileDialog.AddExtension = $true

    # Show the dialog and check if the user clicked Save
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $output_dir = $saveFileDialog.FileName  # Get the selected file path

        # Create the template data
        $template = @{
            StartDateFormat = $dateTimePickerStart.Value.ToString("dd-MM-yyyy")
            EndDateFormat = $dateTimePickerEnd.Value.ToString("dd-MM-yyyy")
            Subfolder1 = $textBoxSubFolder1.Text
            Subfolder2 = $textBoxSubFolder2.Text
            OutputDirectory = $output_dir
        }

        # Convert the template to JSON format
        $json = $template | ConvertTo-Json

        try {
            # Save JSON data to the selected file path
            $json | Out-File -FilePath $output_dir -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Template saved successfully!", "Template Saved")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save the template. Check your permissions.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})



# Button to load template
$buttonLoadTemplate = New-Object System.Windows.Forms.Button
$buttonLoadTemplate.Text = "Load Template"
$buttonLoadTemplate.Location = New-Object System.Drawing.Point(230, 380)
$buttonLoadTemplate.Size = New-Object System.Drawing.Size(90,25)
$form.Controls.Add($buttonLoadTemplate)

# Event to load template data from JSON file
$buttonLoadTemplate.Add_Click({
    $templatePath = [System.Windows.Forms.OpenFileDialog]::new()
    $templatePath.Filter = "JSON Files|*.json"
    if ($templatePath.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $template = Get-Content -Path $templatePath.FileName | ConvertFrom-Json
        
        # Parse the date strings to DateTime objects with the correct format
        $dateTimePickerStart.Value = [datetime]::ParseExact($template.StartDateFormat, "dd-MM-yyyy", $null)
        $dateTimePickerEnd.Value = [datetime]::ParseExact($template.EndDateFormat, "dd-MM-yyyy", $null)
        
        # Populate the other fields
        $textBoxSubFolder1.Text = $template.Subfolder1
        $textBoxSubFolder2.Text = $template.Subfolder2
        
        # Extract the directory from OutputDirectory (remove the filename)
        $outputDirectoryPath = [System.IO.Path]::GetDirectoryName($template.OutputDirectory)
        $textBoxDirectory.Text = $outputDirectoryPath
        
    }
})



# Update the validation and folder creation code to use the DateTimePicker values
function Validate-Inputs {
    if ([string]::IsNullOrEmpty($textBoxSubFolder1.Text) -or [string]::IsNullOrEmpty($textBoxSubFolder2.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Subfolder names cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}

$checkboxLogging = New-Object System.Windows.Forms.CheckBox
$checkboxLogging.Location = New-Object System.Drawing.Point(170,250)
$checkboxLogging.Size = New-Object System.Drawing.Size(20,20)
$form.Controls.Add($checkboxLogging)

# Add a label and checkbox for logging
$labelLogging = New-Object System.Windows.Forms.Label
$labelLogging.Text = "Enable Logging:"
$labelLogging.Location = New-Object System.Drawing.Point(10,250)
$labelLogging.Size = New-Object System.Drawing.Size(140,40)
$form.Controls.Add($labelLogging)



# Add a Help button
$buttonHelp = New-Object System.Windows.Forms.Button
$buttonHelp.Text = "Help"
$buttonHelp.Location = New-Object System.Drawing.Point(380,310)
$buttonHelp.Size = New-Object System.Drawing.Size(75,25)
$form.Controls.Add($buttonHelp)
$buttonHelp.BackColor = [System.Drawing.Color]::FromArgb(100, 149, 237) # CornflowerBlue
$buttonHelp.ForeColor = [System.Drawing.Color]::White
$buttonHelp.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(135, 206, 250) # LightSkyBlue
$buttonHelp.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) # SteelBlue

# Add a Preview button
$buttonPreview = New-Object System.Windows.Forms.Button
$buttonPreview.Text = "Preview"
$buttonPreview.Location = New-Object System.Drawing.Point(10,310)
$buttonPreview.Size = New-Object System.Drawing.Size(75,40)
$form.Controls.Add($buttonPreview)
$buttonPreview.BackColor = [System.Drawing.Color]::FromArgb(255, 165, 0) # Orange
$buttonPreview.ForeColor = [System.Drawing.Color]::White
$buttonPreview.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 215, 0) # Gold
$buttonPreview.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(255, 140, 0) # DarkOrange


# Add a ComboBox for language selection
$comboBoxLanguage = New-Object System.Windows.Forms.ComboBox
$comboBoxLanguage.Items.AddRange(@("English", "French"))
$comboBoxLanguage.SelectedIndex = 0
$comboBoxLanguage.Location = New-Object System.Drawing.Point(380,20)
$comboBoxLanguage.Size = New-Object System.Drawing.Size(80,20)
$form.Controls.Add($comboBoxLanguage)

# Add event to handle language change
$comboBoxLanguage.Add_SelectedIndexChanged({
    switch ($comboBoxLanguage.SelectedItem) {
        "English" {
            $labelStartDate.Text = "Start Date (dd-MM-yy):"
            $labelEndDate.Text = "End Date (dd-MM-yy):"
            $labelSubFolder1.Text = "Subfolder 1 Name:"
            $labelSubFolder2.Text = "Subfolder 2 Name:"
            $labelDirectory.Text = "Output Directory:"
            $labelLogging.Text = "Enable Logging:"
            $button.Text = "Create Folders"
            $buttonBrowse.Text = "Browse"
            $buttonHelp.Text = "Help"
            $buttonPreview.Text = "Preview"
        }
        "French" {
            $labelStartDate.Text = "Date de debut (jj-MM-aa):"
            $labelEndDate.Text = "Date de fin (jj-MM-aa):"
            $labelSubFolder1.Text = "Nom du sous-dossier 1:"
            $labelSubFolder2.Text = "Nom du sous-dossier 2:"
            $labelDirectory.Text = "Repertoire de sortie:"
            $labelLogging.Text = "Activer la journalisation:"
            $button.Text = "Creer des dossiers"
            $buttonBrowse.Text = "Parcourir"
            $buttonHelp.Text = "Aide"
            $buttonPreview.Text = "Apercu"
        }
    }
})

# Update the Preview button event
$buttonPreview.Add_Click({
    if (-not (Validate-Inputs)) {
        return
    }

    $start_date = $dateTimePickerStart.Value
    $end_date = $dateTimePickerEnd.Value
    $subfolder1 = $textBoxSubFolder1.Text
    $subfolder2 = $textBoxSubFolder2.Text

    # Generate preview of the folder structure in a list format
    $dateList = @()
    $current_date = $start_date
    while ($current_date -le $end_date) {
        $date_str = $current_date.ToString("dd-MM-yy")
        $dateList += "📅 Date: $date_str`n  ├── $subfolder1`n  └── $subfolder2`n"
        $current_date = $current_date.AddDays(1)
    }

    # Group dates into columns for better display
    $columns = 2  # Number of columns to display
    $rows = [math]::Ceiling($dateList.Count / $columns)
    $preview = "Folder Structure Preview:`n"
    $preview += "=============================`n`n"
    for ($i = 0; $i -lt $rows; $i++) {
        for ($j = 0; $j -lt $columns; $j++) {
            $index = $i + ($j * $rows)
            if ($index -lt $dateList.Count) {
                $preview += $dateList[$index].PadRight(30)  # Adjust padding if needed
            }
        }
        $preview += "`n"
    }

    # Create a new form for the preview
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "Preview"
    $previewForm.Size = New-Object System.Drawing.Size(500, 400)

    # Create a TextBox to display the preview text
    $previewBox = New-Object System.Windows.Forms.TextBox
    $previewBox.Multiline = $true
    $previewBox.ReadOnly = $true
    $previewBox.ScrollBars = "Vertical"
    $previewBox.Size = New-Object System.Drawing.Size(480, 360)
    $previewBox.Text = $preview

    # Add the TextBox to the form
    $previewForm.Controls.Add($previewBox)

    # Show the preview form
    $previewForm.ShowDialog()
})




# Create a checkbox for Dark Mode
$checkboxDarkMode = New-Object System.Windows.Forms.CheckBox
$checkboxDarkMode.Text = "Dark Mode"
$checkboxDarkMode.Location = New-Object System.Drawing.Point(380, 60)
$checkboxDarkMode.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($checkboxDarkMode)

# Add event to handle Dark Mode toggle
$checkboxDarkMode.Add_CheckedChanged({
    if ($checkboxDarkMode.Checked) {
        $form.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $form.ForeColor = [System.Drawing.Color]::White
        foreach ($control in $form.Controls) {
            $control.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
            $control.ForeColor = [System.Drawing.Color]::White
            
            # Retain button colors in dark mode
            if ($control -is [System.Windows.Forms.Button]) {
                $control.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) # SteelBlue
                $control.ForeColor = [System.Drawing.Color]::White
                $control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 149, 237) # CornflowerBlue
                $control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(65, 105, 225) # RoyalBlue
            }
        }
    } else {
        $form.BackColor = [System.Drawing.Color]::White
        $form.ForeColor = [System.Drawing.Color]::Black
        foreach ($control in $form.Controls) {
            $control.BackColor = [System.Drawing.Color]::White
            $control.ForeColor = [System.Drawing.Color]::Black
            
            # Retain button colors in white mode
            if ($control -is [System.Windows.Forms.Button]) {
                $control.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180) # SteelBlue
                $control.ForeColor = [System.Drawing.Color]::White
                $control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 149, 237) # CornflowerBlue
                $control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(65, 105, 225) # RoyalBlue
            }
        }
    }
})


# Add a ComboBox for date format selection
$comboBoxDateFormat = New-Object System.Windows.Forms.ComboBox
$comboBoxDateFormat.Items.AddRange(@("dd-MM-yy", "MM-dd-yy", "yyyy-MM-dd"))
$comboBoxDateFormat.SelectedIndex = 0
$comboBoxDateFormat.Location = New-Object System.Drawing.Point(380, 100)
$comboBoxDateFormat.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($comboBoxDateFormat)

# Update the date format used in DateTimePickers
$comboBoxDateFormat.Add_SelectedIndexChanged({
    # This will trigger when the user selects a different date format
    $selectedDateFormat = $comboBoxDateFormat.SelectedItem

    # Update the DateTimePickers' CustomFormat to reflect the selected format
    $dateTimePickerStart.CustomFormat = $selectedDateFormat
    $dateTimePickerEnd.CustomFormat = $selectedDateFormat
})



# Update the validation and folder creation code to use the DateTimePicker values
function Validate-Inputs {
    if ([string]::IsNullOrEmpty($textBoxSubFolder1.Text) -or [string]::IsNullOrEmpty($textBoxSubFolder2.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Subfolder names cannot be empty.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}


# Add click event to Help button
$buttonHelp.Add_Click({
    $helpText = @"
Instructions:

1. Enter the start and end dates in dd-MM-yy format.
2. Enter names for the two Sub-folders.
3. Select the output directory using the Browse button.
4. Check 'Enable Logging' to save a log file.
5. Click 'Create Folders' to generate the folder structure.
"@
    [System.Windows.Forms.MessageBox]::Show($helpText, "Help", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})

# Quit button
$buttonQuit = New-Object System.Windows.Forms.Button
$buttonQuit.Text = "Quit"
$buttonQuit.Location = New-Object System.Drawing.Point(380,340)
$buttonQuit.Size = New-Object System.Drawing.Size(75,25)
$buttonQuit.BackColor = [System.Drawing.Color]::Red  # Set background color to red
$buttonQuit.ForeColor = [System.Drawing.Color]::White  # Set text color to white for contrast
$form.Controls.Add($buttonQuit)

# Add Click event for Quit button to close the form
$buttonQuit.Add_Click({
    $form.Close()
})

# Create ToolTip object
$toolTip = New-Object System.Windows.Forms.ToolTip

# Update tooltips
$toolTip.SetToolTip($dateTimePickerStart, "Select the start date (dd-MM-yyyy format).")
$toolTip.SetToolTip($dateTimePickerEnd, "Select the end date (dd-MM-yyyy format).")
$toolTip.SetToolTip($textBoxSubFolder1, "Enter the name for the first subfolder (e.g., Documents).")
$toolTip.SetToolTip($textBoxSubFolder2, "Enter the name for the second subfolder (e.g., Images).")
$toolTip.SetToolTip($buttonBrowse, "Click to browse and select the output directory where folders will be created.")
$toolTip.SetToolTip($checkboxLogging, "Enable this option to log the folder creation process for troubleshooting.")
$toolTip.SetToolTip($button, "Click here to create the folders based on the specified start date, end date, and subfolder names.")
$toolTip.SetToolTip($buttonHelp, "Click to view a detailed guide on how to use the application.")
$toolTip.SetToolTip($buttonPreview, "Click to preview the folder structure before creating the folders.")

# Add button click event
$button.Add_Click({
    if (-not (Validate-Inputs)) {
        return
    }
    $logPath = ""
    if ($checkboxLogging.Checked) {
        $logPath = Join-Path -Path $output_dir -ChildPath "creation_log.txt"
        Add-Content -Path $logPath -Value "Folder Creation Log - $(Get-Date)"
    }
    
    try {
        # Directly use the DateTimePicker values
        $start_date = $dateTimePickerStart.Value
        $end_date = $dateTimePickerEnd.Value
        $subfolder1 = $textBoxSubFolder1.Text
        $subfolder2 = $textBoxSubFolder2.Text
        $output_dir = $textBoxDirectory.Text

        if (-not (Test-Path $output_dir)) {
            throw "Invalid output directory."
        }

        # Calculate total days for progress bar
        $total_days = ($end_date - $start_date).Days + 1
        $progressBar.Maximum = $total_days
        $progressBar.Value = 0

        # Loop through the dates
        $current_date = $start_date
        while ($current_date -le $end_date) {
            # Format the date string
            $date_str = $current_date.ToString("dd-MM-yy")
            
            # Create the main folder
            $main_folder = New-Item -Path $output_dir -Name $date_str -ItemType Directory -Force
            
            # Create subfolders
            New-Item -Path $main_folder.FullName -Name $subfolder1 -ItemType Directory -Force
            New-Item -Path $main_folder.FullName -Name $subfolder2 -ItemType Directory -Force

            # Increment the date by one day
            $current_date = $current_date.AddDays(1)
            # Update progress bar
            $progressBar.Value += 1
        }

        [System.Windows.Forms.MessageBox]::Show("Folders created successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to create folders. Check your inputs and try again.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }

    if ($checkboxLogging.Checked) {
        Add-Content -Path $logPath -Value "Created folder: $date_str"
    }
})

$form.Add_FormClosing({
    param($sender, $e)

    # Prevent any popups on form close by directly disposing of the form
    $form.DialogResult = [System.Windows.Forms.DialogResult]::None
    $e.Cancel = $false
    [System.Windows.Forms.Application]::Exit()  # Ensure application exits fully
})

# Set the form background color
$form.BackColor = [System.Drawing.Color]::White


# Add the button to the form
$form.Controls.Add($button)

# Show the form
$form.ShowDialog()

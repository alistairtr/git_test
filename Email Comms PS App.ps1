# Load the necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users.Actions
Import-Module SQLServer

#User Capture
$Global:currentuser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$Data = Invoke-SQLCmd -ServerInstance "LPADM-77MSSV3"  -Database "Incidents" -TrustServerCertificate  -Query "SELECT * FROM dbo.Incident WHERE NOT isOpen = 'Closed'"

# Email Body Generation Func
Function GenerateEmailBody ($templateType) {
  $headerColor = $null
  switch ($templateType) {
      "Service Outage" {
          $headerColor = "#d32f2f" # Red
      }
      "Service Degradation" {
          $headerColor = "#e0ac1e" # Yellow
      }
      "Service Update" {
          $headerColor = "#e0ac1e" # Yellow
      }
      "Service Resolution" {
          $headerColor = "#0c8016" # Green
      }
  }
  
  $emailBody = @"
<!DOCTYPE html>
<html lang="en">
<head>
<title>$templateType Notification</title>
</head>
<body style="font-family: Arial, sans-serif; background-color: #ffffff; color: #000000; margin: 0; padding: 0;">
<div style="max-width: 900px; margin: auto; border: 1px solid #cccccc; border-radius: 5px; overflow: hidden;">
<div style="background-color: $headerColor; color: white; padding: 15px; text-align: center; font-size: 20px;">
  $templateType - $($textboxes[1].Text)
</div>
<div style="padding: 25px; font-size: 16px;">
  <p><strong>Line of business:</strong> $($textboxes[2].Text)</p>
  <p><strong>Impact:</strong> $($textboxes[8].Text)</p>
  <p><strong>Status:</strong> $($Status.SelectedItem)</p>
  <p><strong>Actions:</strong> $($textboxes[4].Text)</p>
  <p><strong>Next update due:</strong> $($textboxes[5].Text)</p>
  <p><strong>Incident reference:</strong> $($textboxes[0].Text)</p>
  <p><strong>Incident start:</strong> $($textboxes[6].Text)</p>
  <p><strong>Promoted to a High Priority incident:</strong> $($textboxes[7].Text)</p>
</div>
<div style="background-color: #ffffff; padding: 10px; text-align: center;">
  <img src="https://i.imgur.com/DlGDQug.png" style="max-width:75%; height: auto;">
</div>
</div>
</body>
</html>
"@
  return $emailBody
}




function RefreshData {
    # Fetch updated data from the database
    $Global:Data = Invoke-SQLCmd -ServerInstance "LPADM-77MSSV3" -Database "Incidents" -TrustServerCertificate -Query "SELECT * FROM dbo.Incident WHERE NOT isOpen = 'Closed'"
    
    # Clear existing items in the ComboBox
    $comboBox.Items.Clear()
    
    # Repopulate the presets and ComboBox with the updated data
    $Global:presets = foreach ($row in $Global:Data) {
        [PSCustomObject]@{
            IssueNo = $row.IssueNumber
            IssueService = $row.IssueService
            Impact = $row.Impact
            DivisionsImpacted = $row.Divisions
            Company = $row.Company
            Status = $row.Status
            Actions = $row.Actions
            NextUpdate = $row.NextUpdate
            IncidentStart = $row.IncidentStart
            PriorityUpdateTime = $row.PriorityUpdateTime
            NotifType = $row.NotifType
        }
    }
    
    $Global:presets | ForEach-Object { 
        $_
        $comboBox.Items.Add($_)
    }
    $comboBox.DisplayMember = 'IssueService'
}

Connect-MgGraph -ClientId "CLIENT ID HERE" -TenantId "TENANT ID HERE" -CertificateThumbprint "THUMBPRINT HERE" -NoWelcome

# Initialize the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IT Service Delivery Alert System'
$form.Size = New-Object System.Drawing.Size(1450,950)
$form.StartPosition = 'CenterScreen'#

# Define the theme colors
$redColor = [System.Drawing.Color]::FromArgb(211, 47, 47) # This color is similar to the red in the logo
$whiteColor = [System.Drawing.Color]::FromArgb(217, 217, 217)

# Define presets for ComboBox selection
$presets = @()

foreach ($row in $Data) {
    $presets += @{
        IssueNo = $row.IssueNumber
        IssueService = $row.IssueService
        Impact = $row.Impact
        DivisionsImpacted = $row.Divisions
        NotifType = $row.NotifType
        Status = $row.Status
        Company = $row.Company
        Actions = $row.Actions
        NextUpdate = $row.NextUpdate
        IncidentStart = $row.IncidentStart
        PriorityUpdateTime = $row.PriorityUpdateTime
    }
}

# Create the ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 10)
$comboBox.Size = New-Object System.Drawing.Size(200, 20)

$NotifTypeCombo = New-Object System.Windows.Forms.ComboBox
$NotifTypeCombo.Location = New-Object System.Drawing.Point(220, 10)
$NotifTypeCombo.Size = New-Object System.Drawing.Size(200, 20)

$Status = New-Object System.Windows.Forms.ComboBox
$Status.Location = New-Object System.Drawing.Point(440, 10)
$Status.Size = New-Object System.Drawing.Size(200, 20)

$IsOpenCombo = New-Object System.Windows.Forms.ComboBox
$IsOpenCombo.Location = New-Object System.Drawing.Point(660, 10)
$IsOpenCombo.Size = New-Object System.Drawing.Size(200, 20)

$IsOpenValues = @('Open', 'Closed')
foreach ($type in $IsOpenValues) {
    $IsOpenCombo.Items.Add($type)
}

# Create a WebBrowser control
$webBrowser = New-Object System.Windows.Forms.WebBrowser
$webBrowser.Location = New-Object System.Drawing.Point(360,(65 + ($i * 30)))
$webBrowser.Size = New-Object System.Drawing.Size(1000, 750)

$notifTypes = @('Service Outage', 'Service Degradation', 'Service Update', 'Service Resolution')
foreach ($type in $notifTypes) {
    $NotifTypeCombo.Items.Add($type)
}

# Define the possible status options based on NotifType
$statusOptions = @{
    'Service Outage' = @('Investigating', 'Analysing', 'Monitoring')
    'Service Degradation' = @('Investigating', 'Analysing', 'Monitoring')
    'Service Update' = @('Investigating', 'Analysing', 'Monitoring')
    'Service Resolution' = @('Monitoring', 'Resolved')
}

# Event handler for when the selection in NotifTypeCombo changes
$NotifTypeCombo.Add_SelectedIndexChanged({
    # Clear existing items in the Status ComboBox
    $Status.Items.Clear()

    # Retrieve the selected NotifType
    $selectedNotifType = $NotifTypeCombo.SelectedItem.ToString()

    # Based on the selected NotifType, add the relevant items to the Status ComboBox
    $statusOptions[$selectedNotifType] | ForEach-Object{
        $Status.Items.Add($_)
    }

    # Optionally, you can set the default selected item for Status here
    if ($Status.Items.Count -gt 0) {
        $Status.SelectedIndex = 0
    }
})

# Add presets to the ComboBox and display the IssueService property
$presets | ForEach-Object { 
    $item = [pscustomobject]$_
    $comboBox.Items.Add($item)
}
$comboBox.DisplayMember = 'IssueService'

# Create labels and textboxes based on the presets
$labelsText = @('Issue Number', 'Issue Service', 'Divisions Impacted', 'Company', 'Actions', 'Next Update', 'Incident Start', 'Incident Priority Status Time', 'Impact')
$labels = @()
$textboxes = @()

for ($i = 0; $i -lt $labelsText.Length; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelsText[$i]
    $label.Location = New-Object System.Drawing.Point(10,(65 + ($i * 30)))
    $label.Size = New-Object System.Drawing.Size(100, 30)
    $form.Controls.Add($label)
    $labels += $label

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(120,(65 + ($i * 30)))
    $textbox.Size = New-Object System.Drawing.Size(200, 100)
    $textbox.ReadOnly = $false
    $form.Controls.Add($textbox)
    $textboxes += $textbox
}

# Event handler for changing the selection in the ComboBox
$comboBox.Add_SelectedIndexChanged({
    $selectedItem = $comboBox.SelectedItem
    $NotifTypeCombo.SelectedItem = $combobox.SelectedItem.NotifType 
    $Status.SelectedItem = $combobox.SelectedItem.Status 
    $IsOpenCombo.SelectedIndex = 0
    $textboxes[0].Text = $selectedItem.IssueNo
    $textboxes[1].Text = $selectedItem.IssueService
    $textboxes[2].Text = $selectedItem.DivisionsImpacted
    $textboxes[3].Text = $selectedItem.Company
    $textboxes[4].Text = $selectedItem.Actions
    $textboxes[5].Text = $selectedItem.NextUpdate
    $textboxes[6].Text = $selectedItem.IncidentStart
    $textboxes[7].Text = $selectedItem.PriorityUpdateTime
    $textboxes[8].Text = $selectedItem.Impact
    $textboxes[8].Multiline = $true
    $textboxes[8].ScrollBars = 'Vertical'
})

# Create Update Button
$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Location = New-Object System.Drawing.Point(120, (150 + ($labelsText.Length * 35)))
$updateButton.Size = New-Object System.Drawing.Size(200, 30)
$updateButton.Text = 'Update Incident'
$updateButton.Add_Click({ 
    $query = @"
    UPDATE dbo.Incident
    SET IssueNumber = '$($textboxes[0].Text.Replace("'", "''"))',
        Impact = '$($textboxes[8].Text.Replace("'", "''"))',
        Status = '$($Status.SelectedItem.Replace("'", "''"))',
        Divisions = '$($textboxes[2].Text.Replace("'", "''"))',
        Company = '$($textboxes[3].Text.Replace("'", "''"))',
        Actions = '$($textboxes[4].Text.Replace("'", "''"))',
        NextUpdate = '$($textboxes[5].Text.Replace("'", "''"))',
        IncidentStart = '$($textboxes[6].Text.Replace("'", "''"))',
        PriorityUpdateTime = '$($textboxes[7].Text.Replace("'", "''"))',
        NotifType = '$($NotifTypeCombo.SelectedItem.Replace("'", "''"))',
        IssueService = '$($textboxes[1].Text.Replace("'", "''"))',
        IsOpen = '$($IsOpenCombo.SelectedItem.Replace("'", "''"))',
        LastModifyByUser = '$($Global:currentuser.Replace("'", "''"))'
    WHERE IssueNumber = '$($textboxes[0].Text.Replace("'", "''"))';
"@  
    Invoke-SQLCmd -ServerInstance "LPADM-77MSSV3"  -Database "Incidents" -TrustServerCertificate  -Query $query

    RefreshData
    $comboBox.SelectedIndex = 0
})

# Create the SendEmail button
$sendemailbut = New-Object System.Windows.Forms.Button
$sendemailbut.Location = New-Object System.Drawing.Point(120, (150 + ($labelsText.Length * 30)))
$sendemailbut.Size = New-Object System.Drawing.Size(200, 30)
$sendemailbut.Text = 'Send Email'
$sendemailbut.Add_Click({
    # Body Generation

    $emailSendBody = $NotifTypeCombo.SelectedItem.ToString()

    $emailBody = GenerateEmailBody $emailSendBody
    # Email payload construction
    $params = @{
        message = @{
            subject = $NotifTypeCombo.SelectedItem.ToString() + " - " +$textboxes[1].Text.ToString()
            body = @{
                contentType = "HTML"
                content = $emailBody
            }
            BccRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "alistair.trout@bca.com"
                    }
                },
                @{
                    EmailAddress = @{
                        Address = "d9a73b40.bca.com@emea.teams.ms"
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = "ITServiceAlerts@bca.com"
                }
            }
        }
        saveToSentItems = $false
    } | ConvertTo-Json -Depth 10  # Use -Depth parameter to ensure deep conversion
        
    Send-MgUserMail -UserId "RECIPIENT OBJECT ID" -BodyParameter $params
    $query = @"
    UPDATE dbo.Incident
    SET IssueNumber = '$($textboxes[0].Text.Replace("'", "''"))',
        Impact = '$($textboxes[8].Text.Replace("'", "''"))',
        Status = '$($Status.SelectedItem.Replace("'", "''"))',
        Divisions = '$($textboxes[2].Text.Replace("'", "''"))',
        Company = '$($textboxes[3].Text.Replace("'", "''"))',
        Actions = '$($textboxes[4].Text.Replace("'", "''"))',
        NextUpdate = '$($textboxes[5].Text.Replace("'", "''"))',
        IncidentStart = '$($textboxes[6].Text.Replace("'", "''"))',
        PriorityUpdateTime = '$($textboxes[7].Text.Replace("'", "''"))',
        NotifType = '$($NotifTypeCombo.SelectedItem.Replace("'", "''"))',
        IssueService = '$($textboxes[1].Text.Replace("'", "''"))',
        IsOpen = '$($IsOpenCombo.SelectedItem.Replace("'", "''"))',
        LastModifyByUser = '$($Global:currentuser.Replace("'", "''"))'
    WHERE IssueNumber = '$($textboxes[0].Text.Replace("'", "''"))';
"@
    

    Invoke-SQLCmd -ServerInstance "LPADM-77MSSV3"  -Database "Incidents" -TrustServerCertificate  -Query $query
    RefreshData
})

$previewButton = New-Object System.Windows.Forms.Button
$previewButton.Location = New-Object System.Drawing.Point(120, (150 + ($labelsText.Length * 40)))
$previewButton.Size = New-Object System.Drawing.Size(200, 30)
$previewButton.Text = 'Preview Email'
$previewButton.Add_Click({
    $emailTemplateType = $NotifTypeCombo.SelectedItem.ToString()

    $emailpreview = GenerateEmailBody $emailTemplateType


    $webBrowser.DocumentText = $emailpreview
    
})


# Create the 'Clear' button
$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Location = New-Object System.Drawing.Point(880, 10) # Adjust the location as needed
$clearButton.Size = New-Object System.Drawing.Size(100, 30) # Adjust the size as needed
$clearButton.Text = 'New Incident'

# Define the 'Clear' button click event handler
$clearButton.Add_Click({
    ShowNewEntryForm
})

# CAG Branding
$imageBox = New-Object System.Windows.Forms.PictureBox
$imageBox.Location = New-Object System.Drawing.Point(50, 600)
$imageBox.Size = New-Object System.Drawing.Size(240, 140)
$imageBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$imageBox.Image = [System.Drawing.Image]::FromFile("C:\Users\AlistairTrout\OneDrive - British Car Auctions - Europe\Desktop\Email Application\caglogo.png")

$imageTextLabel = New-Object System.Windows.Forms.Label
$imageTextLabel.Text = "Property of CAG - Do not use without authorisation to do so."
$imageTextLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$imageTextLabel.AutoSize = $false
$imageTextLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$imageTextLabel.Width = $imageBox.Width
$imageTextLabel.Height = 30
$imageTextLabel.Location = New-Object System.Drawing.Point(50, ($imageBox.Bottom + 5))


# Add the ComboBox to the form
$form.Controls.Add($comboBox)
$form.Controls.Add($sendemailbut)
$form.Controls.Add($updateButton)
$form.Controls.Add($NotifTypeCombo)
$form.Controls.Add($IsOpenCombo)
$form.Controls.Add($Status)
$form.Controls.Add($webBrowser)
$form.Controls.Add($previewButton)
$form.Controls.Add($clearButton)
$form.Controls.Add($imageBox)
$form.Controls.Add($imageTextLabel)

# Set the default ComboBox selection
$comboBox.SelectedIndex = 0
$IsOpenCombo.SelectedIndex = 0
$NotifTypeCombo.SelectedItem = $combobox.SelectedItem.NotifType 
$Status.SelectedItem = $combobox.SelectedItem.Status

function ShowNewEntryForm {
    $newEntryForm = New-Object System.Windows.Forms.Form
    $newEntryForm.Text = 'Add New Incident'
    $newEntryForm.Size = New-Object System.Drawing.Size(800, 650) # Adjust the size as needed
    $newEntryForm.StartPosition = 'CenterScreen'

    $textboxstore = @()

    # Define labels
    $labelsText = @('Issue Number', 'Issue Service', 'Divisions Impacted', 'Company', 'Actions', 'Next Update', 'Incident Start', 'Incident Priority Status Time')
    $labelLocations = @(10, 10), @(10, 40), @(10, 70), @(10, 100), @(10, 130), @(10, 160), @(10, 190), @(10, 220), @(10, 250)
    $textBoxLocations = @(120, 10), @(120, 40), @(120, 70), @(120, 100), @(120, 130), @(120, 160), @(120, 190), @(120, 220), @(120, 250)
    $textBoxSizes = New-Object System.Drawing.Size 260, 30

    for ($i = 0; $i -lt $labelsText.Length; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labelsText[$i]
        $label.Location = New-Object System.Drawing.Point $labelLocations[$i]
        $label.Size = New-Object System.Drawing.Size 100, 30
        $newEntryForm.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point $textBoxLocations[$i]
        $textBox.Size = $textBoxSizes
        $textBox.Multiline = $i -eq 999 # Set multiline where needed
        $textBox.ScrollBars = 'Vertical' # Add scrollbars to multiline textboxes
        $newEntryForm.Controls.Add($textBox)

        $textboxstore += $textbox
    }

    $impactLabel = New-Object System.Windows.Forms.Label
    $impactLabel.Text = "Impact"
    $impactLabel.Location = New-Object System.Drawing.Point 10, 280
    $impactLabel.Size = New-Object System.Drawing.Size 100, 30

    $impactTextBox = New-Object System.Windows.Forms.TextBox
    $impactTextBox.Location = New-Object System.Drawing.Point 120, 280
    $impactTextBox.Size = New-Object System.Drawing.Size 400, 150
    $impactTextBox.Multiline = $true
    $impactTextBox.ScrollBars = 'Vertical' # Add scrollbars to multiline textboxes

    $newEntryForm.Controls.Add($impactLabel)
    $newEntryForm.Controls.Add($impactTextBox)

    $newNotifyLabel = New-Object System.Windows.Forms.Label
    $newNotifyLabel.Text = "Notification Type"
    $newNotifyLabel.Location = New-Object System.Drawing.Point 420, 10
    $newNotifyLabel.Size = New-Object System.Drawing.Size 100, 30

    $newNotifType = New-Object System.Windows.Forms.ComboBox
    $newNotifType.Location = New-Object System.Drawing.Point(530, 10)
    $newNotifType.Size = New-Object System.Drawing.Size(200, 20)

    $newStatusLabel = New-Object System.Windows.Forms.Label
    $newStatusLabel.Text = "Status Type"
    $newStatusLabel.Location = New-Object System.Drawing.Point 420, 50
    $newStatusLabel.Size = New-Object System.Drawing.Size 100, 30

    $newStatus = New-Object System.Windows.Forms.ComboBox
    $newStatus.Location = New-Object System.Drawing.Point(530, 50)
    $newStatus.Size = New-Object System.Drawing.Size(200, 20)

    $newEntryForm.Controls.Add($newStatus)
    $newEntryForm.Controls.Add($newNotifyLabel)
    $newEntryForm.Controls.Add($newStatusLabel)
    $newEntryForm.Controls.Add($newNotifType)

    $newEntryForm.BackColor = $whiteColor

    $newnotifTypes = @('Service Outage', 'Service Degradation', 'Service Update', 'Service Resolution')
    foreach ($type in $newnotifTypes) {
        $newNotifType.Items.Add($type)
    }

    # Define the possible status options based on NotifType
    $newStatusOps = @{
        'Service Outage' = @('Investigating', 'Analysing', 'Monitoring')
        'Service Degradation' = @('Investigating', 'Analysing', 'Monitoring')
        'Service Update' = @('Investigating', 'Analysing', 'Monitoring')
        'Service Resolution' = @('Monitoring', 'Resolved')
    }

        # Event handler for when the selection in NotifTypeCombo changes
    $newNotifType.Add_SelectedIndexChanged({
        # Clear existing items in the Status ComboBox
        $newStatus.Items.Clear()

        # Retrieve the selected NotifType
        $selectedNotifType = $newNotifType.SelectedItem.ToString()

        # Based on the selected NotifType, add the relevant items to the Status ComboBox
        $newStatusOps[$selectedNotifType] | ForEach-Object{
            $newStatus.Items.Add($_)
        }

        # Optionally, you can set the default selected item for Status here
        if ($newStatus.Items.Count -gt 0) {
            $newStatus.SelectedIndex = 0
        }
    })


    # Add a button to insert the new entry into the database
    $insertButton = New-Object System.Windows.Forms.Button
    $insertButton.Location = New-Object System.Drawing.Point(120, 500) # Adjust as needed
    $insertButton.Size = New-Object System.Drawing.Size(180, 30) # Adjust as needed
    $insertButton.Text = 'Insert New Incident'
    $insertButton.BackColor = $redColor
    $insertButton.ForeColor = $whiteColor
    $insertButton.Add_Click({
        $query = @"
INSERT INTO dbo.Incident (IssueNumber, IssueService, Divisions, Company, Actions, NextUpdate, IncidentStart, PriorityUpdateTime, Impact, Status, NotifType, IsOpen, CreatedByUser) 
VALUES ('$($textboxstore[0].Text)', '$($textboxstore[1].Text)', '$($textboxstore[2].Text)', '$($textboxstore[3].Text)', '$($textboxstore[4].Text)', '$($textboxstore[5].Text)', '$($textboxstore[6].Text)', '$($textboxstore[7].Text)', '$($impactTextBox.Text)', '$($newStatus.SelectedItem)', '$($newNotifType.SelectedItem)', 'Open', '$($Global:currentuser)')
"@
        Invoke-SQLCmd -ServerInstance "LPADM-77MSSV3" -Database "Incidents" -TrustServerCertificate -Query $query
        
        # Close the new entry form
        $newEntryForm.Close()
        RefreshData
    })
    $newEntryForm.Controls.Add($insertButton)

    # Show the new entry form
    $newEntryForm.ShowDialog()
}



# Set the form's background color
$form.BackColor = $whiteColor

# Example for setting a button's color with the theme
$sendemailbut.BackColor = $redColor
$sendemailbut.ForeColor = $whiteColor

$previewButton.BackColor = $redColor
$previewButton.ForeColor = $whiteColor

$updateButton.BackColor = $redColor
$updateButton.ForeColor = $whiteColor

$clearButton.BackColor = $redColor
$clearButton.ForeColor = $whiteColor



# Show the form
$form.ShowDialog()

# ------------------------------------------------------------
# github:  https://github.com/pranjulknit
# Email       : pranjul.19641@knit.ac.in
# Description : PowerShell GUI to manage SCOM 2019 Maintenance Mode used hash for optimization
# Target Org  : HCLTech
# Version     : 2.0
# Last Updated: 13 June 2025
# "And Miles to go before I Sleep"
# ------------------------------------------------------------


# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import SCOM Module
Import-Module OperationsManager

# Connect to SCOM
New-SCOMManagementGroupConnection -ComputerName "PMMANT4000ZAVJ.mmpci.net"
$scomclass = Get-SCOMClass -Name Microsoft.Windows.Computer

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "SCOM Maintenance Mode GUI"
$form.Size = '550, 640'
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ====== Server File Picker ======
$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Text = "Select Server List (.txt):"
$labelFile.Location = '20,20'
$labelFile.Size = '300,20'
$form.Controls.Add($labelFile)

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Location = '20,45'
$textBoxFile.Size = '360,24'
$textBoxFile.ReadOnly = $true
$form.Controls.Add($textBoxFile)

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse"
$buttonBrowse.Location = '390,44'
$buttonBrowse.Size = '100,25'
$buttonBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Text files (*.txt)|*.txt"
    if ($dialog.ShowDialog() -eq 'OK') {
        $textBoxFile.Text = $dialog.FileName
    }
})
$form.Controls.Add($buttonBrowse)

# ====== Comment ======
$labelComment = New-Object System.Windows.Forms.Label
$labelComment.Text = "Comment:"
$labelComment.Location = '20,80'
$form.Controls.Add($labelComment)

$textBoxComment = New-Object System.Windows.Forms.TextBox
$textBoxComment.Location = '20,105'
$textBoxComment.Size = '470,24'
$textBoxComment.Text = "Maintenance via GUI"
$form.Controls.Add($textBoxComment)

# ====== Action Dropdown ======
$labelAction = New-Object System.Windows.Forms.Label
$labelAction.Text = "Select Action:"
$labelAction.Location = '20,140'
$form.Controls.Add($labelAction)

$comboBoxAction = New-Object System.Windows.Forms.ComboBox
$comboBoxAction.Location = '20,165'
$comboBoxAction.Size = '200,25'
$comboBoxAction.Items.AddRange(@("Start Maintenance Mode", "Stop Maintenance Mode", "Update Maintenance Mode"))
$comboBoxAction.SelectedIndex = 0
$form.Controls.Add($comboBoxAction)

# ====== Duration Dropdown ======
$labelDuration = New-Object System.Windows.Forms.Label
$labelDuration.Text = "Duration (select or enter in minutes):"
$labelDuration.Location = '250,140'
$form.Controls.Add($labelDuration)

$comboBoxDuration = New-Object System.Windows.Forms.ComboBox
$comboBoxDuration.Location = '250,165'
$comboBoxDuration.Size = '240,25'
$comboBoxDuration.Items.AddRange(@("60", "120", "180", "240", "1440", "10080", "43200"))
$comboBoxDuration.Text = "60"
$form.Controls.Add($comboBoxDuration)

$labelNote = New-Object System.Windows.Forms.Label
$labelNote.Text = "(60 = 1h, 120 = 2h, ..., 43200 = 1 month)"
$labelNote.Location = '250,195'
$labelNote.Size = '250,15'
$labelNote.ForeColor = 'Gray'
$form.Controls.Add($labelNote)

# ====== Status Label ======
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Status: Waiting"
$labelStatus.Location = '20,230'
$labelStatus.Size = '500,20'
$form.Controls.Add($labelStatus)

# ====== Scrollable Result Box ======
$textBoxResult = New-Object System.Windows.Forms.TextBox
$textBoxResult.Multiline = $true
$textBoxResult.ScrollBars = "Vertical"
$textBoxResult.Location = '20,260'
$textBoxResult.Size = '500,260'
$textBoxResult.ReadOnly = $true
$form.Controls.Add($textBoxResult)

# ====== Execute Button ======
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Execute"
$buttonExecute.Location = '20,540'
$buttonExecute.Size = '100,30'
$form.Controls.Add($buttonExecute)

# ====== Duration Toggle ======
$comboBoxAction.Add_SelectedIndexChanged({
    $comboBoxDuration.Enabled = ($comboBoxAction.SelectedItem -ne "Stop Maintenance Mode")
})

# ====== Main Action ======
$buttonExecute.Add_Click({
    $buttonExecute.Enabled = $false
    $textBoxResult.Clear()
    $labelStatus.Text = "Running..."

    if (-not $textBoxFile.Text) {
        $labelStatus.Text = "Error: No server file selected"
        $buttonExecute.Enabled = $true
        return
    }

    $servers = Get-Content $textBoxFile.Text | ForEach-Object { $_.Trim().ToLower() }
    $comment = $textBoxComment.Text
    $action = $comboBoxAction.SelectedItem
    $durationMinutes = if ($comboBoxDuration.Enabled) { [int]$comboBoxDuration.Text } else { 0 }
    $now = Get-Date
    $reason = [Microsoft.EnterpriseManagement.Monitoring.MaintenanceModeReason]::PlannedOther

    $success = 0
    $fail = 0
    $failedList = @()

    # Optimized Hash Lookup
    $instanceMap = @{}
    $allInstances = Get-SCOMClassInstance -Class $scomclass
    foreach ($inst in $allInstances) {
        $instName = $inst.DisplayName.Trim().ToLower()
        if (-not $instanceMap.ContainsKey($instName)) {
            $instanceMap[$instName] = $inst
        }
    }

    foreach ($server in $servers) {
        $labelStatus.Text = "Processing: $server"
        $form.Refresh()

        if ($instanceMap.ContainsKey($server)) {
            $instance = $instanceMap[$server]
            try {
                $mm = Get-SCOMMaintenanceMode -Instance $instance

                if ($action -eq "Stop Maintenance Mode") {
                    if ($mm) {
                        $instance.StopMaintenanceMode($now.ToUniversalTime(), "Recursive")
                        $textBoxResult.AppendText("$server - Stopped MM`r`n")
                        $success++
                    } else {
                        $textBoxResult.AppendText("$server - Not in MM`r`n")
                        $fail++
                        $failedList += "$server - Not in MM"
                    }
                }
                elseif ($action -eq "Start Maintenance Mode") {
                    if (-not $mm) {
                        $endTime = $now.AddMinutes($durationMinutes)
                        $instance.ScheduleMaintenanceMode($now.ToUniversalTime(), $endTime.ToUniversalTime(), $reason, $comment, "Recursive")
                        $textBoxResult.AppendText("$server - Started MM for $durationMinutes mins`r`n")
                        $success++
                    } else {
                        $textBoxResult.AppendText("$server - Already in MM`r`n")
                        $fail++
                        $failedList += "$server - Already in MM"
                    }
                }
                elseif ($action -eq "Update Maintenance Mode") {
                    if ($mm) {
                        $newEnd = $now.AddMinutes($durationMinutes)
                        $instance.UpdateMaintenanceMode($newEnd.ToUniversalTime(), $reason, $comment, "Recursive")
                        $textBoxResult.AppendText("$server - Extended MM by $durationMinutes mins`r`n")
                        $success++
                    } else {
                        $textBoxResult.AppendText("$server - Not in MM`r`n")
                        $fail++
                        $failedList += "$server - Not in MM"
                    }
                }
            } catch {
                $textBoxResult.AppendText("$server - ERROR: $_`r`n")
                $fail++
                $failedList += "$server - Exception: $_"
            }
        } else {
            $textBoxResult.AppendText("$server - Not found in SCOM`r`n")
            $fail++
            $failedList += "$server - Not found"
        }
    }

    # Write failed list to file
    if ($failedList.Count -gt 0) {
        $folderPath = Split-Path $textBoxFile.Text
        $failedFile = Join-Path $folderPath "FailedServers.txt"
        $failedList | Out-File -FilePath $failedFile -Encoding utf8
    }

    $labelStatus.Text = "Completed - Success: $success, Fail: $fail"
    $buttonExecute.Enabled = $true
})

# Show GUI
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
$form.ShowDialog()



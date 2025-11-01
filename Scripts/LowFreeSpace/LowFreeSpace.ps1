Param(
    [Parameter(Mandatory= $false)]
    [System.Management.Automation.PSCredential]$ADM_Credential
)

. (Join-Path $PSScriptRoot "..\..\Modules\Get-Session.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Get-DiskSpaceDetails.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Export-DiskReport.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Get-TopItems.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Clear-SystemCache.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Compress-IISLogs.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Test-DiskAvailability.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Test-ReportFileCreation.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Test-ServerAvailability.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Write-Log.ps1")
. (Join-Path $PSScriptRoot "..\..\Modules\Write-WindowsEventLog.ps1")

<#
# Import modules
$modulesToImport = @(
    "$PSScriptRoot\..\..\Modules\Get-Session.ps1",
    "$PSScriptRoot\..\..\Modules\Get-DiskSpaceDetails.ps1",
    "$PSScriptRoot\..\..\Modules\Export-DiskReport.ps1",
    "$PSScriptRoot\..\..\Modules\Get-TopItems.ps1",
    "$PSScriptRoot\..\..\Modules\Clear-SystemCache.ps1",
    "$PSScriptRoot\..\..\Modules\Compress-IISLogs.ps1",
    "$PSScriptRoot\..\..\Modules\Test-DiskAvailability.ps1",
    "$PSScriptRoot\..\..\Modules\Test-ReportFileCreation.ps1",
    "$PSScriptRoot\..\..\Modules\Test-ServerAvailability.ps1",
    "$PSScriptRoot\..\..\Modules\Write-Log.ps1",
    "$PSScriptRoot\..\..\Modules\Write-WindowsEventLog.ps1"
)

foreach ($modulePath in $modulesToImport) {
    try {
        # Read the module content
        $moduleContent = Get-Content -Path $modulePath -Raw

        # Use Invoke-Expression to execute the module content
        Invoke-Expression -Command $moduleContent
        Write-Host "Successfully imported module $moduleName in remote session" -ForegroundColor Green
    } catch {
        Write-Host "Error importing module $modulePath : $_" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error importing module $([System.IO.Path]::GetFileNameWithoutExtension($modulePath)) : $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        exit 1
    }
}#>

<#
# Temporary workaround for testing
if (-not $ADM_Credential) {
    $userName = "user1"
    $password = ConvertTo-SecureString "Leduc123" -AsPlainText -Force
    $ADM_Credential = New-Object System.Management.Automation.PSCredential($userName, $password)
}#>

# Get current user
$CurrentUser = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Update-StatusLabel {
    param(
        [Parameter(Mandatory=$true)]
        [string]$text
    )
    
    $statusLabel.Text = $text
    $statusLabel_width = $statusLabel.PreferredWidth
    $label_x = ($main_form.ClientSize.Width - $statusLabel_width) / 2
    $statusLabel.Location = New-Object System.Drawing.Point($label_x, $statusLabel.Location.Y)
    $statusLabel.Refresh()

}
function Remove-Session {
    <#
    .SYNOPSIS
        Removes the PowerShell session and cleans up resources.
    .DESCRIPTION
        This function closes the PowerShell session if it exists and is still open.
        It also disposes of the main form to free up resources.
    .PARAMETER session
        The PowerShell session to remove.
    .EXAMPLE
        Remove-Session -session $session
        This will close the PowerShell session and dispose of the main form if it exists.
    #>
    try {
        # Check if session exists and is still open before removing it
        if ($session -and $session.State -eq "Open") {
            Remove-PSSession -Session $session
        }
        else {
            Write-Log "No session to close or session already closed" "Info"
        }
    } catch {
        $errorDetails = "Exception: $($_.Exception.GetType().FullName)`nMessage: $($_.Exception.Message)`nStackTrace: $($_.ScriptStackTrace)"
        Write-Log "Error closing session: $errorDetails" "Error"
    }

    # Optionally, clean up form resources to free memory
    if ($main_form) {
        $main_form.Dispose()
        Write-Log "Form disposed and cleaned up"
    }
}

# Get all video controller objects
$screens = Get-WmiObject -Class Win32_VideoController

# Initialize scale factors
$scaleX = 1
$scaleY = 1

# Set design resolution
$designWidth = 1920
$designHeight = 1080

<#
# Loop through all video controllers
foreach ($screen in $screens) {
    $screenWidth = $screen.CurrentHorizontalResolution
    $screenHeight = $screen.CurrentVerticalResolution
    if ($screenWidth -and $screenHeight) {
        $scaleX = $screenWidth / $designWidth
        $scaleY = $screenHeight / $designHeight
    }
}#>

# Vertical padding between objects
$verticalPadding = 7 * $scaleY

# Create ToolTip object
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000  # Time the tooltip remains visible (in milliseconds)
$toolTip.InitialDelay = 500   # Time before the tooltip appears (in milliseconds)
$toolTip.ReshowDelay = 500    # Time before tooltip reappears if mouse moves away and back
$toolTip.ShowAlways = $true   # Show tooltip even if the form is not active

# Server Name Label
$labelServerName = New-Object System.Windows.Forms.Label
$labelServerName.Location = New-Object System.Drawing.Point([Math]::Round(20 * $scaleX), [Math]::Round(20 * $scaleY))
$labelServerName.Size = New-Object System.Drawing.Size([Math]::Round(120 * $scaleX), [Math]::Round(30 * $scaleY))
$labelServerName.Text = "Server Name:"
$labelServerName.Font = New-Object System.Drawing.Font("Arial", [Math]::Round(11 * $scaleY))
$toolTip.SetToolTip($labelServerName, "Enter the hostname or IP address of the remote server to analyze or clean.")

# Disk Name TextBox
$textBoxServerName = New-Object System.Windows.Forms.TextBox
$textBoxServerName.Location = New-Object System.Drawing.Point(($labelServerName.Location.X + $labelServerName.Width), $labelServerName.Location.Y)
$textBoxServerName.Size = New-Object System.Drawing.Size([Math]::Round(250 * $scaleX), $labelServerName.Height)
$textBoxServerName.Font = $labelServerName.Font
$textBoxServerName.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        # Select all text in the ComboBox
        $textBoxServerName.SelectAll()
        $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Copy selected text to clipboard
        if ($textBoxServerName.SelectedText) {
            [System.Windows.Forms.Clipboard]::SetText($textBoxServerName.SelectedText)
        } else {
            [System.Windows.Forms.Clipboard]::SetText($textBoxServerName.Text)
        }
        $e.SuppressKeyPress = $true
    }
})

$mainFormWidth = [Math]::Round(($textBoxServerName.Location.X + $textBoxServerName.Width + 40 * $scaleX))

# Disk Name Label
$diskLabel = New-Object System.Windows.Forms.Label
$diskLabel.Location = New-Object System.Drawing.Point($labelServerName.Location.X, ($labelServerName.Location.Y + $labelServerName.Height + $verticalPadding))
$diskLabel.Size = $labelServerName.Size
$diskLabel.Text = "Drive Letter:"
$diskLabel.Font = $labelServerName.Font
$toolTip.SetToolTip($diskLabel, "Enter the drive letter to process (e.g., C or C: or C:\).")

# Disk Name TextBox
$diskTextBox = New-Object System.Windows.Forms.TextBox
$diskTextBox.Location = New-Object System.Drawing.Point($textBoxServerName.Location.X, $diskLabel.Location.Y)
$diskTextBox.Size = $textBoxServerName.Size
$diskTextBox.Font = $labelServerName.Font
$diskTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        # Select all text in the ComboBox
        $diskTextBox.SelectAll()
        $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Copy selected text to clipboard
        if ($diskTextBox.SelectedText) {
            [System.Windows.Forms.Clipboard]::SetText($diskTextBox.SelectedText)
        } else {
            [System.Windows.Forms.Clipboard]::SetText($diskTextBox.Text)
        }
        $e.SuppressKeyPress = $true
    }
})

# Ticket number Label
$ticketNumberLabel = New-Object System.Windows.Forms.Label
$ticketNumberLabel.Location = New-Object System.Drawing.Point($labelServerName.Location.X, ($diskLabel.Location.Y + $diskLabel.Height + $verticalPadding))
$ticketNumberLabel.Size = $labelServerName.Size
$ticketNumberLabel.Text = "Ticket Number:"
$ticketNumberLabel.Font = $labelServerName.Font
$toolTip.SetToolTip($ticketNumberLabel, "Enter the ticket number associated with this operation.")

# Ticket number TextBox
$ticketNumberTextBox = New-Object System.Windows.Forms.TextBox
$ticketNumberTextBox.Location = New-Object System.Drawing.Point($textBoxServerName.Location.X, $ticketNumberLabel.Location.Y)
$ticketNumberTextBox.Size = $textBoxServerName.Size
$ticketNumberTextBox.Font = $labelServerName.Font
$ticketNumberTextBox.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        # Select all text in the ComboBox
        $ticketNumberTextBox.SelectAll()
        $e.SuppressKeyPress = $true
    }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) {
        # Copy selected text to clipboard
        if ($ticketNumberTextBox.SelectedText) {
            [System.Windows.Forms.Clipboard]::SetText($ticketNumberTextBox.SelectedText)
        } else {
            [System.Windows.Forms.Clipboard]::SetText($ticketNumberTextBox.Text)
        }
        $e.SuppressKeyPress = $true
    }
})

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Size = New-Object System.Drawing.Size([Math]::Round(80 * $scaleX), [Math]::Round(30 * $scaleY))
$okButton.Text = "OK"
$okButton.Add_Click({
    try {
        # Normalize disk name input
        $rawDiskName = $diskTextBox.Text.Trim()
        $diskName = $rawDiskName -replace '[:\\]', ''
        $diskName = $diskName.ToUpper()
        $serverName = $textBoxServerName.Text.Trim()
        $ticketNumber = $ticketNumberTextBox.Text

        if ([string]::IsNullOrEmpty($diskName) -or [string]::IsNullOrEmpty($serverName) -or [string]::IsNullOrEmpty($ticketNumber)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter server name, disk name and ticket number.", 
                "Warning", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        Update-StatusLabel -text "Testing server availability..."
        $result = Test-ServerAvailability -serverName $serverName
        if (-not $result.RemotingAvailable) {
            [System.Windows.Forms.MessageBox]::Show(
                "Server '$serverName' is not available for remoting. Details: $($result.ErrorDetails)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        Update-StatusLabel -text "Creating session with server..."
        $session = Get-Session -serverName $serverName
        if ($null -eq $session) {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to create a session with server '$serverName'.", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        Update-StatusLabel -text "Checking disk $diskName availability..."
        if (-not (Test-DiskAvailability -session $session -diskName $diskName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Disk '$diskName' is not available on server '$serverName'.", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        if (-not (Test-ReportFileCreation)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Cannot proceed - local log file creation failed", 
                "Error", 
                [System.Windows.Forms.MessageBoxButtons]::OK, 
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        try {
            if ($diskName -eq "C") {
                Update-StatusLabel -text "Cleaning C disk. Please wait..."
                $Before = Get-DiskSpaceDetails -session $session -diskName $diskName

                Update-StatusLabel -text "Cleaning system cache..."
                $clearSystemCache = Clear-SystemCache -session $session -Verbose *>&1 | ForEach-Object {
                    "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss'): $_"
                } | Out-String

                Update-StatusLabel -text "Compressing IIS logs..."
                $clearIISLogs = Compress-IISLogs -session $session -Verbose *>&1 | ForEach-Object {
                    "$(Get-Date -Format 'dd-MM-yyyy HH:mm:ss'): $_"
                } | Out-String

                $After = Get-DiskSpaceDetails -session $session -diskName $diskName
                $freePercentageDisk = $After.FreePercentage
                $topRoot = $null
                $topUsers = $null

                if ($After.FreePercentage -lt 10) {
                    Update-StatusLabel -text "Free space still low. Identifying top items..."
                    $topRoot = Get-TopItems -session $session -path "$($diskName):\" -exclude @("Windows", "Program Files", "Program Files (x86)", "ProgramData","Users") -topN 10
                    $topUsers = Get-TopItems -session $session -path "$($diskName):\Users" -topN 10
                }

                [System.Windows.Forms.MessageBox]::Show(
                    "Drive $($diskName). Free space is $($freePercentageDisk)%.`nPlease check report for details.", 
                    "Information", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                # Export disk report
                $reportPath = Export-DiskReport -serverName $serverName -diskName $diskName `
                    -diskInfo $After -beforeDiskInfo $Before `
                    -systemCacheLog $clearSystemCache `
                    -iisLogCleanupLog $clearIISLogs `
                    -topUsers $topUsers -topRoot $topRoot
                
                # Write Windows Event Log Entry on the remote server
                $eventMessage = "User: $CurrentUser`n" + "Ticket Number: $ticketNumber`n" + "Message: C drive cleanup performed. Free space is now $($freePercentageDisk)%.`n"
            } else {
                Update-StatusLabel -text "Getting disk information and top items..."
                $diskInfo = Get-DiskSpaceDetails -session $session -diskName $diskName
                $topItems = Get-TopItems -session $session -path "$($diskName):\" -topN 10

                $freePercentageDisk = $diskInfo.FreePercentage

                [System.Windows.Forms.MessageBox]::Show(
                    "Drive $($diskName). Free space is $($freePercentageDisk)%.`nPlease check report for details.", 
                    "Information", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                $reportPath = Export-DiskReport -serverName $serverName -diskName $diskName `
                    -diskInfo $diskInfo -topItems $topItems           

                # Write Windows Event Log Entry on the remote server
                $eventMessage = "User: $CurrentUser`n" + "Ticket Number: $ticketNumber`n" + "Message: Disk $($diskName) analysis performed. Free space is now $($freePercentageDisk)%.`n"
            }

            # Check if report was successfully created
            if ($reportPath) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Disk report exported to $reportPath", 
                    "Success", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                Update-StatusLabel -text "Low free space analysis completed."
                Start-Process -FilePath $reportPath -ErrorAction SilentlyContinue
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to export disk report.", 
                    "Error", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }

            Write-WindowsEventLog -LogName "Application" -Source "DiskAnalysisScript" `
                    -EventID 1002 -EntryType "Information" `
                    -Message $eventMessage -Session $session
            
            Update-StatusLabel -text "Disk analysis completed successfully."
            $main_form.Close()
            Remove-Session
        } catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error")
            Write-Log "Error during disk analysis: $_" "Error"
            Remove-Session
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log "Error in OK button click event: $_" "Error"
        Remove-Session
    }
})

# Exit Button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Size = $okButton.Size
$cancelButton.Text = "Cancel"
$cancelButton.BackColor = [System.Drawing.Color]::LightCoral
$cancelButton.Add_Click({
    $main_form.Close()
    Remove-Session
}
)

# Calculate horizontal positions for centered alignment
$buttonWidth = $okButton.Size.Width
$spaceBetween = 25
$totalWidth = ($buttonWidth * 2) + $spaceBetween
$startX = ($mainFormWidth - $totalWidth) / 2

# Position buttons
$okButton.Location = New-Object System.Drawing.Point($startX, ($ticketNumberLabel.Location.Y + $ticketNumberLabel.Height + $verticalPadding))
$cancelButton.Location = New-Object System.Drawing.Point(($startX + $buttonWidth + $spaceBetween), $okButton.Location.Y)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.AutoSize = $true  # Important:  Let the label size itself to the text
$statusLabel_width = $statusLabel.PreferredWidth # get the actual width of the label based on the text
$label_x = ($main_form.ClientSize.Width - $statusLabel_width) / 2  # Center horizontally
$label_y = $cancelButton.Location.Y + $cancelButton.Height + $verticalPadding
$statusLabel.Location = New-Object System.Drawing.Point($label_x, ($label_y + 10))  # Add some vertical padding

# Main Form Length Calculation
$mainFormLength = [Math]::Round($statusLabel.Location.Y + $statusLabel.Height + $verticalPadding + 50*$scaleY)

# Main form setup
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = "Low Free Space - $CurrentUser"
$main_form.Size = New-Object System.Drawing.Size($mainFormWidth, $mainFormLength) #430x270 pixels
$main_form.StartPosition = "CenterScreen"
$main_form.FormBorderStyle = 'FixedSingle'  # Or 'FixedDialog'
$main_form.MaximizeBox = $false
$main_form.TopMost = $false  # Keep form on top
$main_form.KeyPreview = $true  # Important: This allows the form to receive key events before controls
$main_form.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
        $cancelButton.PerformClick()
    }
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $okButton.PerformClick()
    }
})

# Add components to form
$main_form.Controls.Add($labelServerName)
$main_form.Controls.Add($textBoxServerName)
$main_form.Controls.Add($diskLabel)
$main_form.Controls.Add($diskTextBox)
$main_form.Controls.Add($ticketNumberLabel)
$main_form.Controls.Add($ticketNumberTextBox)
$main_form.Controls.Add($okButton)
$main_form.Controls.Add($cancelButton)
$main_form.Controls.Add($statusLabel)

# Show form
if ($null -eq $env:UNIT_TEST) {
    $main_form.ShowDialog()
}

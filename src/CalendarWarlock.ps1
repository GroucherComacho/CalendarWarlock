#Requires -Version 5.1

<#
.SYNOPSIS
    CalendarWarlock - Exchange Online Bulk Calendar Permissions Manager
.DESCRIPTION
    A GUI tool for managing Exchange Online calendar permissions in bulk by job title.
    Features:
    - Grant a user access to all calendars of users with a specific job title
    - Grant all users with a specific job title access to a user's calendar
    - Simple, secure, and easy-to-use interface
.NOTES
    Author: CalendarWarlock
    Requires: PowerShell 5.1+, ExchangeOnlineManagement module, Microsoft.Graph.Users module
.EXAMPLE
    .\CalendarWarlock.ps1
#>

#region Script Configuration
$ErrorActionPreference = "Stop"
$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogPath = Join-Path $script:ScriptPath "Logs"
$script:IsConnected = $false
#endregion

#region Load Required Assemblies and Modules
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Import custom modules
try {
    Import-Module (Join-Path $script:ScriptPath "Modules\ExchangeOperations.psm1") -Force
    Import-Module (Join-Path $script:ScriptPath "Modules\AzureADOperations.psm1") -Force
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load required modules. Please ensure the Modules folder exists.`n`nError: $($_.Exception.Message)",
        "CalendarWarlock - Module Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Check for required PowerShell modules
$requiredModules = @("ExchangeOnlineManagement", "Microsoft.Graph.Users")
$missingModules = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    $message = "The following required PowerShell modules are missing:`n`n"
    $message += ($missingModules -join "`n")
    $message += "`n`nPlease install them using:`nInstall-Module -Name <ModuleName> -Scope CurrentUser"

    [System.Windows.Forms.MessageBox]::Show(
        $message,
        "CalendarWarlock - Missing Modules",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}
#endregion

#region Logging Functions
function Initialize-Logging {
    if (-not (Test-Path $script:LogPath)) {
        New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
    }
    $script:LogFile = Join-Path $script:LogPath "CalendarWarlock_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }

    # Also update the status in the GUI if available
    if ($script:StatusLabel) {
        $script:StatusLabel.Text = $Message
        [System.Windows.Forms.Application]::DoEvents()
    }
}
#endregion

#region GUI Helper Functions
function Update-ProgressBar {
    param(
        [int]$Value,
        [int]$Maximum = 100
    )

    if ($script:ProgressBar) {
        $script:ProgressBar.Maximum = $Maximum
        $script:ProgressBar.Value = [Math]::Min($Value, $Maximum)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Update-ResultsLog {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )

    if ($script:ResultsTextBox) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        $prefix = switch ($Type) {
            "Success" { "[OK]" }
            "Error" { "[ERROR]" }
            "Warning" { "[WARN]" }
            default { "[INFO]" }
        }
        $script:ResultsTextBox.AppendText("$timestamp $prefix $Message`r`n")
        $script:ResultsTextBox.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Clear-ResultsLog {
    if ($script:ResultsTextBox) {
        $script:ResultsTextBox.Clear()
    }
}

function Set-UIEnabled {
    param([bool]$Enabled)

    $controls = @(
        $script:ConnectButton,
        $script:OrganizationTextBox,
        $script:JobTitleComboBox,
        $script:DepartmentComboBox,
        $script:OfficeComboBox,
        $script:RefreshTitlesButton,
        $script:TargetUserTextBox,
        $script:SearchUserButton,
        $script:GetPermissionsButton,
        $script:PermissionComboBox,
        $script:GrantToUserButton,
        $script:GrantToTitleButton,
        $script:RemoveFromUserButton,
        $script:RemoveFromTitleButton
    )

    foreach ($control in $controls) {
        if ($control) {
            $control.Enabled = $Enabled
        }
    }
}
#endregion

#region Core Business Logic
function Connect-Services {
    param([string]$Organization)

    Clear-ResultsLog
    Update-ResultsLog "Connecting to Exchange Online and Microsoft Graph..." "Info"
    Write-Log "Initiating connection to $Organization" "INFO"

    Set-UIEnabled -Enabled $false
    $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee

    try {
        # Connect to Exchange Online
        Update-ResultsLog "Connecting to Exchange Online..." "Info"
        $exoResult = Connect-ExchangeOnlineSession -Organization $Organization

        if (-not $exoResult.Success) {
            throw "Exchange Online: $($exoResult.Message)"
        }
        Update-ResultsLog $exoResult.Message "Success"

        # Connect to Microsoft Graph
        Update-ResultsLog "Connecting to Microsoft Graph..." "Info"
        $graphResult = Connect-GraphSession

        if (-not $graphResult.Success) {
            throw "Microsoft Graph: $($graphResult.Message)"
        }
        Update-ResultsLog $graphResult.Message "Success"

        $script:IsConnected = $true
        $script:ConnectButton.Text = "Disconnect"
        $script:StatusLabel.Text = "Connected to $Organization"

        Update-ResultsLog "Successfully connected to all services!" "Success"
        Write-Log "Successfully connected to $Organization" "SUCCESS"

        # Load job titles, departments, and offices
        Refresh-AllSelections
    }
    catch {
        Update-ResultsLog "Connection failed: $($_.Exception.Message)" "Error"
        Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        $script:IsConnected = $false

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect: $($_.Exception.Message)",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $script:ProgressBar.Value = 0
        Set-UIEnabled -Enabled $true
    }
}

function Disconnect-Services {
    Update-ResultsLog "Disconnecting from services..." "Info"
    Write-Log "Disconnecting from services" "INFO"

    try {
        Disconnect-ExchangeOnlineSession | Out-Null
        Disconnect-GraphSession | Out-Null

        $script:IsConnected = $false
        $script:ConnectButton.Text = "Connect"
        $script:StatusLabel.Text = "Disconnected"
        $script:JobTitleComboBox.Items.Clear()
        $script:DepartmentComboBox.Items.Clear()
        $script:OfficeComboBox.Items.Clear()

        Update-ResultsLog "Disconnected from all services" "Success"
        Write-Log "Disconnected from all services" "SUCCESS"
    }
    catch {
        Update-ResultsLog "Error during disconnect: $($_.Exception.Message)" "Warning"
    }
}

function Refresh-JobTitles {
    Update-ResultsLog "Loading job titles from Azure AD..." "Info"
    $script:JobTitleComboBox.Items.Clear()

    try {
        $result = Get-AllJobTitles

        if ($result.Success -and $result.JobTitles.Count -gt 0) {
            foreach ($title in $result.JobTitles) {
                $script:JobTitleComboBox.Items.Add($title) | Out-Null
            }
            Update-ResultsLog "Loaded $($result.Count) job titles" "Success"

            if ($script:JobTitleComboBox.Items.Count -gt 0) {
                $script:JobTitleComboBox.SelectedIndex = 0
            }
        }
        else {
            Update-ResultsLog "No job titles found or error: $($result.Message)" "Warning"
        }
    }
    catch {
        Update-ResultsLog "Failed to load job titles: $($_.Exception.Message)" "Error"
    }
}

function Refresh-Departments {
    Update-ResultsLog "Loading departments from Azure AD..." "Info"
    $script:DepartmentComboBox.Items.Clear()

    try {
        $result = Get-AllDepartments

        if ($result.Success -and $result.Departments.Count -gt 0) {
            foreach ($dept in $result.Departments) {
                $script:DepartmentComboBox.Items.Add($dept) | Out-Null
            }
            Update-ResultsLog "Loaded $($result.Count) departments" "Success"

            if ($script:DepartmentComboBox.Items.Count -gt 0) {
                $script:DepartmentComboBox.SelectedIndex = 0
            }
        }
        else {
            Update-ResultsLog "No departments found or error: $($result.Message)" "Warning"
        }
    }
    catch {
        Update-ResultsLog "Failed to load departments: $($_.Exception.Message)" "Error"
    }
}

function Refresh-Offices {
    Update-ResultsLog "Loading office locations from Azure AD..." "Info"
    $script:OfficeComboBox.Items.Clear()

    try {
        $result = Get-AllOffices

        if ($result.Success -and $result.Offices.Count -gt 0) {
            foreach ($office in $result.Offices) {
                $script:OfficeComboBox.Items.Add($office) | Out-Null
            }
            Update-ResultsLog "Loaded $($result.Count) office locations" "Success"

            if ($script:OfficeComboBox.Items.Count -gt 0) {
                $script:OfficeComboBox.SelectedIndex = 0
            }
        }
        else {
            Update-ResultsLog "No office locations found or error: $($result.Message)" "Warning"
        }
    }
    catch {
        Update-ResultsLog "Failed to load office locations: $($_.Exception.Message)" "Error"
    }
}

function Refresh-AllSelections {
    Refresh-JobTitles
    Refresh-Departments
    Refresh-Offices
}

function Search-TargetUser {
    $searchTerm = $script:TargetUserTextBox.Text.Trim()

    if ([string]::IsNullOrEmpty($searchTerm)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a user email or name to search.",
            "Search User",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    Update-ResultsLog "Searching for user: $searchTerm" "Info"

    try {
        # First try exact match
        $result = Get-UserByEmail -Email $searchTerm

        if ($result.Success) {
            $user = $result.User
            $script:TargetUserTextBox.Text = $user.Email
            $script:TargetUserTextBox.Tag = $user
            Update-ResultsLog "Found: $($user.DisplayName) ($($user.Email)) - $($user.JobTitle)" "Success"
        }
        else {
            # Try partial search
            $searchResult = Search-Users -SearchTerm $searchTerm -MaxResults 10

            if ($searchResult.Success -and $searchResult.Count -gt 0) {
                # Show selection dialog
                $selectedUser = Show-UserSelectionDialog -Users $searchResult.Users

                if ($selectedUser) {
                    $script:TargetUserTextBox.Text = $selectedUser.Email
                    $script:TargetUserTextBox.Tag = $selectedUser
                    Update-ResultsLog "Selected: $($selectedUser.DisplayName) ($($selectedUser.Email))" "Success"
                }
            }
            else {
                Update-ResultsLog "No users found matching '$searchTerm'" "Warning"
                [System.Windows.Forms.MessageBox]::Show(
                    "No users found matching '$searchTerm'",
                    "User Not Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    }
    catch {
        Update-ResultsLog "Search failed: $($_.Exception.Message)" "Error"
    }
}

function Get-TargetUserPermissions {
    $targetUser = $script:TargetUserTextBox.Text.Trim()

    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a user email to get permissions for.",
            "Get Permissions",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to Exchange Online first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Getting calendar permissions for: $targetUser" "Info"

    try {
        $result = Get-CalendarPermissions -CalendarOwner $targetUser

        if ($result.Success) {
            Update-ResultsLog "-----------------------------------" "Info"
            Update-ResultsLog "Calendar Permissions for $targetUser" "Info"
            Update-ResultsLog "-----------------------------------" "Info"

            if ($result.Permissions.Count -eq 0) {
                Update-ResultsLog "No permissions found" "Warning"
            }
            else {
                foreach ($perm in $result.Permissions) {
                    $userName = $perm.User.DisplayName
                    if (-not $userName) { $userName = $perm.User.ToString() }
                    $accessRights = $perm.AccessRights -join ", "
                    Update-ResultsLog "$userName : $accessRights" "Success"
                }
                Update-ResultsLog "-----------------------------------" "Info"
                Update-ResultsLog "Total: $($result.Permissions.Count) permission(s)" "Info"
            }
        }
        else {
            Update-ResultsLog "Failed to get permissions: $($result.Message)" "Error"
        }
    }
    catch {
        Update-ResultsLog "Error getting permissions: $($_.Exception.Message)" "Error"
    }
}

function Show-UserSelectionDialog {
    param([array]$Users)

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Select User"
    $dialog.Size = New-Object System.Drawing.Size(500, 350)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.Size = New-Object System.Drawing.Size(465, 250)
    $listView.View = "Details"
    $listView.FullRowSelect = $true
    $listView.GridLines = $true

    $listView.Columns.Add("Name", 150) | Out-Null
    $listView.Columns.Add("Email", 180) | Out-Null
    $listView.Columns.Add("Job Title", 120) | Out-Null

    foreach ($user in $Users) {
        $item = New-Object System.Windows.Forms.ListViewItem($user.DisplayName)
        $item.SubItems.Add($user.Email) | Out-Null
        $item.SubItems.Add($user.JobTitle) | Out-Null
        $item.Tag = $user
        $listView.Items.Add($item) | Out-Null
    }

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Text = "Select"
    $selectButton.Location = New-Object System.Drawing.Point(300, 270)
    $selectButton.Size = New-Object System.Drawing.Size(80, 30)
    $selectButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(395, 270)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dialog.Controls.AddRange(@($listView, $selectButton, $cancelButton))
    $dialog.AcceptButton = $selectButton
    $dialog.CancelButton = $cancelButton

    $script:SelectedUser = $null
    $listView.Add_DoubleClick({
        if ($listView.SelectedItems.Count -gt 0) {
            $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dialog.Close()
        }
    })

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($listView.SelectedItems.Count -gt 0) {
            return $listView.SelectedItems[0].Tag
        }
    }

    return $null
}

function Grant-BulkPermissionsToUser {
    <#
    .SYNOPSIS
        Grants a single user access to all calendars of users matching the selected method
    #>

    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to Exchange Online first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $targetUser = $script:TargetUserTextBox.Text.Trim()
    $permission = $script:PermissionComboBox.SelectedItem

    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a target user email.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $methodResult = Get-UsersForSelectedMethod
    if (-not $methodResult.Success -and -not $methodResult.Method) {
        [System.Windows.Forms.MessageBox]::Show(
            $methodResult.Message,
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Confirmation dialog
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "This will grant '$targetUser' $permission access to the calendars of ALL users with $($methodResult.Method) '$($methodResult.Value)'.`n`nAre you sure you want to continue?",
        "Confirm Bulk Permission Grant",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Starting bulk permission grant..." "Info"
    Update-ResultsLog "Target User: $targetUser" "Info"
    Update-ResultsLog "$($methodResult.Method): $($methodResult.Value)" "Info"
    Update-ResultsLog "Permission: $permission" "Info"
    Write-Log "Grant-BulkPermissionsToUser: User=$targetUser, $($methodResult.Method)=$($methodResult.Value), Permission=$permission" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        Update-ResultsLog "Fetching users with $($methodResult.Method) '$($methodResult.Value)'..." "Info"
        $usersResult = Get-UsersForSelectedMethod

        if (-not $usersResult.Success) {
            throw $usersResult.Message
        }

        $users = $usersResult.Users
        Update-ResultsLog "Found $($users.Count) users" "Info"

        if ($users.Count -eq 0) {
            Update-ResultsLog "No users found." "Warning"
            return
        }

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $users.Count

        for ($i = 0; $i -lt $users.Count; $i++) {
            $user = $users[$i]

            # Skip if granting to themselves
            if ($user.Email -eq $targetUser -or $user.UserPrincipalName -eq $targetUser) {
                Update-ResultsLog "Skipping $($user.DisplayName) (same as target user)" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
                continue
            }

            Update-ResultsLog "Granting access to $($user.DisplayName)'s calendar..." "Info"

            $result = Grant-CalendarPermission -CalendarOwner $user.Email -Trustee $targetUser -AccessRights $permission

            if ($result.Success) {
                Update-ResultsLog "$($result.Action): $($user.DisplayName) ($($user.Email))" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $($user.DisplayName) - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk operation completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk grant completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk permission grant completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $($_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
        Update-ProgressBar -Value 0
    }
}

function Grant-BulkPermissionsToTitle {
    <#
    .SYNOPSIS
        Grants all users matching the selected method access to a single user's calendar
    #>

    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to Exchange Online first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $calendarOwner = $script:TargetUserTextBox.Text.Trim()
    $permission = $script:PermissionComboBox.SelectedItem

    if ([string]::IsNullOrEmpty($calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter the calendar owner's email.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $methodResult = Get-UsersForSelectedMethod
    if (-not $methodResult.Success -and -not $methodResult.Method) {
        [System.Windows.Forms.MessageBox]::Show(
            $methodResult.Message,
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Confirmation dialog
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "This will grant ALL users with $($methodResult.Method) '$($methodResult.Value)' $permission access to $calendarOwner's calendar.`n`nAre you sure you want to continue?",
        "Confirm Bulk Permission Grant",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Starting bulk permission grant..." "Info"
    Update-ResultsLog "Calendar Owner: $calendarOwner" "Info"
    Update-ResultsLog "$($methodResult.Method): $($methodResult.Value)" "Info"
    Update-ResultsLog "Permission: $permission" "Info"
    Write-Log "Grant-BulkPermissionsToTitle: Owner=$calendarOwner, $($methodResult.Method)=$($methodResult.Value), Permission=$permission" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        Update-ResultsLog "Fetching users with $($methodResult.Method) '$($methodResult.Value)'..." "Info"
        $usersResult = Get-UsersForSelectedMethod

        if (-not $usersResult.Success) {
            throw $usersResult.Message
        }

        $users = $usersResult.Users
        Update-ResultsLog "Found $($users.Count) users" "Info"

        if ($users.Count -eq 0) {
            Update-ResultsLog "No users found." "Warning"
            return
        }

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $users.Count

        for ($i = 0; $i -lt $users.Count; $i++) {
            $user = $users[$i]

            # Skip if granting to themselves
            if ($user.Email -eq $calendarOwner -or $user.UserPrincipalName -eq $calendarOwner) {
                Update-ResultsLog "Skipping $($user.DisplayName) (calendar owner)" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
                continue
            }

            Update-ResultsLog "Granting $($user.DisplayName) access to calendar..." "Info"

            $result = Grant-CalendarPermission -CalendarOwner $calendarOwner -Trustee $user.Email -AccessRights $permission

            if ($result.Success) {
                Update-ResultsLog "$($result.Action): $($user.DisplayName) ($($user.Email))" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $($user.DisplayName) - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk operation completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk grant completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk permission grant completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $($_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
        Update-ProgressBar -Value 0
    }
}

function Get-UsersForSelectedMethod {
    <#
    .SYNOPSIS
        Gets users based on the currently selected method (Job Title, Department, or Office)
    .DESCRIPTION
        Returns users filtered by whichever method dropdown has a selection
    #>

    $jobTitle = $script:JobTitleComboBox.Text.Trim()
    $department = $script:DepartmentComboBox.Text.Trim()
    $office = $script:OfficeComboBox.Text.Trim()

    # Determine which method to use - prioritize in order: Job Title, Department, Office
    if (-not [string]::IsNullOrEmpty($jobTitle)) {
        $result = Get-UsersByJobTitle -JobTitle $jobTitle
        return @{
            Success = $result.Success
            Users = $result.Users
            Count = $result.Count
            Method = "Job Title"
            Value = $jobTitle
            Message = $result.Message
        }
    }
    elseif (-not [string]::IsNullOrEmpty($department)) {
        $result = Get-UsersByDepartment -Department $department
        return @{
            Success = $result.Success
            Users = $result.Users
            Count = $result.Count
            Method = "Department"
            Value = $department
            Message = $result.Message
        }
    }
    elseif (-not [string]::IsNullOrEmpty($office)) {
        $result = Get-UsersByOffice -Office $office
        return @{
            Success = $result.Success
            Users = $result.Users
            Count = $result.Count
            Method = "Office"
            Value = $office
            Message = $result.Message
        }
    }
    else {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Method = $null
            Value = $null
            Message = "Please select a Job Title, Department, or Office"
        }
    }
}

function Remove-BulkPermissionsFromUser {
    <#
    .SYNOPSIS
        Removes a single user's access from all calendars of users matching the selected method
    #>

    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to Exchange Online first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $targetUser = $script:TargetUserTextBox.Text.Trim()

    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a target user email.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $methodResult = Get-UsersForSelectedMethod
    if (-not $methodResult.Success -and -not $methodResult.Method) {
        [System.Windows.Forms.MessageBox]::Show(
            $methodResult.Message,
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Confirmation dialog
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "This will REMOVE '$targetUser' access from the calendars of ALL users with $($methodResult.Method) '$($methodResult.Value)'.`n`nAre you sure you want to continue?",
        "Confirm Bulk Permission Removal",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Starting bulk permission removal..." "Info"
    Update-ResultsLog "Target User: $targetUser" "Info"
    Update-ResultsLog "$($methodResult.Method): $($methodResult.Value)" "Info"
    Write-Log "Remove-BulkPermissionsFromUser: User=$targetUser, $($methodResult.Method)=$($methodResult.Value)" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        Update-ResultsLog "Fetching users with $($methodResult.Method) '$($methodResult.Value)'..." "Info"
        $usersResult = Get-UsersForSelectedMethod

        if (-not $usersResult.Success) {
            throw $usersResult.Message
        }

        $users = $usersResult.Users
        Update-ResultsLog "Found $($users.Count) users" "Info"

        if ($users.Count -eq 0) {
            Update-ResultsLog "No users found." "Warning"
            return
        }

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $users.Count

        for ($i = 0; $i -lt $users.Count; $i++) {
            $user = $users[$i]

            # Skip if removing from themselves
            if ($user.Email -eq $targetUser -or $user.UserPrincipalName -eq $targetUser) {
                Update-ResultsLog "Skipping $($user.DisplayName) (same as target user)" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
                continue
            }

            Update-ResultsLog "Removing access from $($user.DisplayName)'s calendar..." "Info"

            $result = Remove-CalendarPermission -CalendarOwner $user.Email -Trustee $targetUser

            if ($result.Success) {
                Update-ResultsLog "Removed: $($user.DisplayName) ($($user.Email))" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $($user.DisplayName) - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk removal completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk removal completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk permission removal completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $($_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
        Update-ProgressBar -Value 0
    }
}

function Remove-BulkPermissionsFromTitle {
    <#
    .SYNOPSIS
        Removes all users matching the selected method from a single user's calendar
    #>

    if (-not $script:IsConnected) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please connect to Exchange Online first.",
            "Not Connected",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $calendarOwner = $script:TargetUserTextBox.Text.Trim()

    if ([string]::IsNullOrEmpty($calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter the calendar owner's email.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $methodResult = Get-UsersForSelectedMethod
    if (-not $methodResult.Success -and -not $methodResult.Method) {
        [System.Windows.Forms.MessageBox]::Show(
            $methodResult.Message,
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Confirmation dialog
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "This will REMOVE access for ALL users with $($methodResult.Method) '$($methodResult.Value)' from $calendarOwner's calendar.`n`nAre you sure you want to continue?",
        "Confirm Bulk Permission Removal",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Starting bulk permission removal..." "Info"
    Update-ResultsLog "Calendar Owner: $calendarOwner" "Info"
    Update-ResultsLog "$($methodResult.Method): $($methodResult.Value)" "Info"
    Write-Log "Remove-BulkPermissionsFromTitle: Owner=$calendarOwner, $($methodResult.Method)=$($methodResult.Value)" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        Update-ResultsLog "Fetching users with $($methodResult.Method) '$($methodResult.Value)'..." "Info"
        $usersResult = Get-UsersForSelectedMethod

        if (-not $usersResult.Success) {
            throw $usersResult.Message
        }

        $users = $usersResult.Users
        Update-ResultsLog "Found $($users.Count) users" "Info"

        if ($users.Count -eq 0) {
            Update-ResultsLog "No users found." "Warning"
            return
        }

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $users.Count

        for ($i = 0; $i -lt $users.Count; $i++) {
            $user = $users[$i]

            # Skip if removing from themselves
            if ($user.Email -eq $calendarOwner -or $user.UserPrincipalName -eq $calendarOwner) {
                Update-ResultsLog "Skipping $($user.DisplayName) (calendar owner)" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
                continue
            }

            Update-ResultsLog "Removing $($user.DisplayName)'s access from calendar..." "Info"

            $result = Remove-CalendarPermission -CalendarOwner $calendarOwner -Trustee $user.Email

            if ($result.Success) {
                Update-ResultsLog "Removed: $($user.DisplayName) ($($user.Email))" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $($user.DisplayName) - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $users.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk removal completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk removal completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk permission removal completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $($_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
        Update-ProgressBar -Value 0
    }
}
#endregion

#region Build Main Form
function Build-MainForm {
    # Main Form
    $script:MainForm = New-Object System.Windows.Forms.Form
    $script:MainForm.Text = "CalendarWarlock - Bulk Calendar Permissions Manager"
    $script:MainForm.Size = New-Object System.Drawing.Size(700, 820)
    $script:MainForm.StartPosition = "CenterScreen"
    $script:MainForm.FormBorderStyle = "FixedSingle"
    $script:MainForm.MaximizeBox = $false
    $script:MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Header Panel
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size = New-Object System.Drawing.Size(700, 75)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "CalendarWarlock"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::White
    $titleLabel.Location = New-Object System.Drawing.Point(15, 8)
    $titleLabel.AutoSize = $true

    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Exchange Online Bulk Calendar Permissions Manager"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::White
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 42)
    $subtitleLabel.AutoSize = $true

    $headerPanel.Controls.AddRange(@($titleLabel, $subtitleLabel))

    # Connection Group
    $connectionGroup = New-Object System.Windows.Forms.GroupBox
    $connectionGroup.Text = "Connection"
    $connectionGroup.Location = New-Object System.Drawing.Point(15, 85)
    $connectionGroup.Size = New-Object System.Drawing.Size(655, 70)

    $orgLabel = New-Object System.Windows.Forms.Label
    $orgLabel.Text = "Organization:"
    $orgLabel.Location = New-Object System.Drawing.Point(15, 30)
    $orgLabel.AutoSize = $true

    $script:OrganizationTextBox = New-Object System.Windows.Forms.TextBox
    $script:OrganizationTextBox.Location = New-Object System.Drawing.Point(100, 27)
    $script:OrganizationTextBox.Size = New-Object System.Drawing.Size(350, 23)
    try { $script:OrganizationTextBox.PlaceholderText = "contoso.onmicrosoft.com" } catch {}

    $script:ConnectButton = New-Object System.Windows.Forms.Button
    $script:ConnectButton.Text = "Connect"
    $script:ConnectButton.Location = New-Object System.Drawing.Point(470, 25)
    $script:ConnectButton.Size = New-Object System.Drawing.Size(100, 28)
    $script:ConnectButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $script:ConnectButton.ForeColor = [System.Drawing.Color]::White
    $script:ConnectButton.FlatStyle = "Flat"
    $script:ConnectButton.Add_Click({
        if ($script:IsConnected) {
            Disconnect-Services
        }
        else {
            $org = $script:OrganizationTextBox.Text.Trim()
            if ([string]::IsNullOrEmpty($org)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please enter your organization domain (e.g., contoso.onmicrosoft.com)",
                    "Organization Required",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
            Connect-Services -Organization $org
        }
    })

    $connectionGroup.Controls.AddRange(@($orgLabel, $script:OrganizationTextBox, $script:ConnectButton))

    # Method Selection Group
    $methodGroup = New-Object System.Windows.Forms.GroupBox
    $methodGroup.Text = "Method Selection"
    $methodGroup.Location = New-Object System.Drawing.Point(15, 165)
    $methodGroup.Size = New-Object System.Drawing.Size(655, 130)

    $jobTitleLabel = New-Object System.Windows.Forms.Label
    $jobTitleLabel.Text = "Job Title:"
    $jobTitleLabel.Location = New-Object System.Drawing.Point(15, 28)
    $jobTitleLabel.AutoSize = $true

    $script:JobTitleComboBox = New-Object System.Windows.Forms.ComboBox
    $script:JobTitleComboBox.Location = New-Object System.Drawing.Point(100, 25)
    $script:JobTitleComboBox.Size = New-Object System.Drawing.Size(350, 23)
    $script:JobTitleComboBox.DropDownStyle = "DropDown"
    $script:JobTitleComboBox.AutoCompleteMode = "SuggestAppend"
    $script:JobTitleComboBox.AutoCompleteSource = "ListItems"

    $departmentLabel = New-Object System.Windows.Forms.Label
    $departmentLabel.Text = "Department:"
    $departmentLabel.Location = New-Object System.Drawing.Point(15, 60)
    $departmentLabel.AutoSize = $true

    $script:DepartmentComboBox = New-Object System.Windows.Forms.ComboBox
    $script:DepartmentComboBox.Location = New-Object System.Drawing.Point(100, 57)
    $script:DepartmentComboBox.Size = New-Object System.Drawing.Size(350, 23)
    $script:DepartmentComboBox.DropDownStyle = "DropDown"
    $script:DepartmentComboBox.AutoCompleteMode = "SuggestAppend"
    $script:DepartmentComboBox.AutoCompleteSource = "ListItems"

    $officeLabel = New-Object System.Windows.Forms.Label
    $officeLabel.Text = "Office:"
    $officeLabel.Location = New-Object System.Drawing.Point(15, 92)
    $officeLabel.AutoSize = $true

    $script:OfficeComboBox = New-Object System.Windows.Forms.ComboBox
    $script:OfficeComboBox.Location = New-Object System.Drawing.Point(100, 89)
    $script:OfficeComboBox.Size = New-Object System.Drawing.Size(350, 23)
    $script:OfficeComboBox.DropDownStyle = "DropDown"
    $script:OfficeComboBox.AutoCompleteMode = "SuggestAppend"
    $script:OfficeComboBox.AutoCompleteSource = "ListItems"

    $script:RefreshTitlesButton = New-Object System.Windows.Forms.Button
    $script:RefreshTitlesButton.Text = "Refresh"
    $script:RefreshTitlesButton.Location = New-Object System.Drawing.Point(470, 55)
    $script:RefreshTitlesButton.Size = New-Object System.Drawing.Size(100, 28)
    $script:RefreshTitlesButton.Add_Click({ Refresh-AllSelections })

    $methodGroup.Controls.AddRange(@($jobTitleLabel, $script:JobTitleComboBox, $departmentLabel, $script:DepartmentComboBox, $officeLabel, $script:OfficeComboBox, $script:RefreshTitlesButton))

    # Target User Group
    $targetUserGroup = New-Object System.Windows.Forms.GroupBox
    $targetUserGroup.Text = "Target User"
    $targetUserGroup.Location = New-Object System.Drawing.Point(15, 305)
    $targetUserGroup.Size = New-Object System.Drawing.Size(655, 70)

    $targetUserLabel = New-Object System.Windows.Forms.Label
    $targetUserLabel.Text = "User Email:"
    $targetUserLabel.Location = New-Object System.Drawing.Point(15, 30)
    $targetUserLabel.AutoSize = $true

    $script:TargetUserTextBox = New-Object System.Windows.Forms.TextBox
    $script:TargetUserTextBox.Location = New-Object System.Drawing.Point(100, 27)
    $script:TargetUserTextBox.Size = New-Object System.Drawing.Size(280, 23)
    try { $script:TargetUserTextBox.PlaceholderText = "user@contoso.com" } catch {}

    $script:SearchUserButton = New-Object System.Windows.Forms.Button
    $script:SearchUserButton.Text = "Search"
    $script:SearchUserButton.Location = New-Object System.Drawing.Point(395, 25)
    $script:SearchUserButton.Size = New-Object System.Drawing.Size(80, 28)
    $script:SearchUserButton.Add_Click({ Search-TargetUser })

    $script:GetPermissionsButton = New-Object System.Windows.Forms.Button
    $script:GetPermissionsButton.Text = "Get Permissions"
    $script:GetPermissionsButton.Location = New-Object System.Drawing.Point(485, 25)
    $script:GetPermissionsButton.Size = New-Object System.Drawing.Size(110, 28)
    $script:GetPermissionsButton.Add_Click({ Get-TargetUserPermissions })

    $targetUserGroup.Controls.AddRange(@($targetUserLabel, $script:TargetUserTextBox, $script:SearchUserButton, $script:GetPermissionsButton))

    # Permission Level Group
    $permissionGroup = New-Object System.Windows.Forms.GroupBox
    $permissionGroup.Text = "Permission Level"
    $permissionGroup.Location = New-Object System.Drawing.Point(15, 385)
    $permissionGroup.Size = New-Object System.Drawing.Size(655, 70)

    $permissionLabel = New-Object System.Windows.Forms.Label
    $permissionLabel.Text = "Access Level:"
    $permissionLabel.Location = New-Object System.Drawing.Point(15, 30)
    $permissionLabel.AutoSize = $true

    $script:PermissionComboBox = New-Object System.Windows.Forms.ComboBox
    $script:PermissionComboBox.Location = New-Object System.Drawing.Point(100, 27)
    $script:PermissionComboBox.Size = New-Object System.Drawing.Size(200, 23)
    $script:PermissionComboBox.DropDownStyle = "DropDownList"

    # Add permission levels
    $permissionLevels = Get-CalendarPermissionLevels
    foreach ($level in $permissionLevels) {
        $script:PermissionComboBox.Items.Add($level.Name) | Out-Null
    }
    $script:PermissionComboBox.SelectedIndex = 6  # Default to Reviewer

    $script:PermissionDescLabel = New-Object System.Windows.Forms.Label
    $script:PermissionDescLabel.Location = New-Object System.Drawing.Point(320, 30)
    $script:PermissionDescLabel.Size = New-Object System.Drawing.Size(320, 20)
    $script:PermissionDescLabel.ForeColor = [System.Drawing.Color]::Gray

    $script:PermissionComboBox.Add_SelectedIndexChanged({
        $selectedLevel = $script:PermissionComboBox.SelectedItem
        $levels = Get-CalendarPermissionLevels
        $desc = ($levels | Where-Object { $_.Name -eq $selectedLevel }).Description
        $script:PermissionDescLabel.Text = $desc
    })

    # Trigger initial description update
    $script:PermissionDescLabel.Text = ($permissionLevels | Where-Object { $_.Name -eq "Reviewer" }).Description

    $permissionGroup.Controls.AddRange(@($permissionLabel, $script:PermissionComboBox, $script:PermissionDescLabel))

    # Actions Group
    $actionsGroup = New-Object System.Windows.Forms.GroupBox
    $actionsGroup.Text = "Bulk Actions"
    $actionsGroup.Location = New-Object System.Drawing.Point(15, 465)
    $actionsGroup.Size = New-Object System.Drawing.Size(655, 120)

    $script:GrantToUserButton = New-Object System.Windows.Forms.Button
    $script:GrantToUserButton.Text = "Grant User Access to All Calendars of Selection"
    $script:GrantToUserButton.Location = New-Object System.Drawing.Point(15, 25)
    $script:GrantToUserButton.Size = New-Object System.Drawing.Size(305, 35)
    $script:GrantToUserButton.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 0)
    $script:GrantToUserButton.ForeColor = [System.Drawing.Color]::White
    $script:GrantToUserButton.FlatStyle = "Flat"
    $script:GrantToUserButton.Add_Click({ Grant-BulkPermissionsToUser })

    $script:GrantToTitleButton = New-Object System.Windows.Forms.Button
    $script:GrantToTitleButton.Text = "Grant All of Selection Access to User's Calendar"
    $script:GrantToTitleButton.Location = New-Object System.Drawing.Point(335, 25)
    $script:GrantToTitleButton.Size = New-Object System.Drawing.Size(305, 35)
    $script:GrantToTitleButton.BackColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $script:GrantToTitleButton.ForeColor = [System.Drawing.Color]::White
    $script:GrantToTitleButton.FlatStyle = "Flat"
    $script:GrantToTitleButton.Add_Click({ Grant-BulkPermissionsToTitle })

    $script:RemoveFromUserButton = New-Object System.Windows.Forms.Button
    $script:RemoveFromUserButton.Text = "Remove User Access from All Calendars of Selection"
    $script:RemoveFromUserButton.Location = New-Object System.Drawing.Point(15, 70)
    $script:RemoveFromUserButton.Size = New-Object System.Drawing.Size(305, 35)
    $script:RemoveFromUserButton.BackColor = [System.Drawing.Color]::FromArgb(180, 50, 50)
    $script:RemoveFromUserButton.ForeColor = [System.Drawing.Color]::White
    $script:RemoveFromUserButton.FlatStyle = "Flat"
    $script:RemoveFromUserButton.Add_Click({ Remove-BulkPermissionsFromUser })

    $script:RemoveFromTitleButton = New-Object System.Windows.Forms.Button
    $script:RemoveFromTitleButton.Text = "Remove All of Selection Access from User's Calendar"
    $script:RemoveFromTitleButton.Location = New-Object System.Drawing.Point(335, 70)
    $script:RemoveFromTitleButton.Size = New-Object System.Drawing.Size(305, 35)
    $script:RemoveFromTitleButton.BackColor = [System.Drawing.Color]::FromArgb(150, 50, 50)
    $script:RemoveFromTitleButton.ForeColor = [System.Drawing.Color]::White
    $script:RemoveFromTitleButton.FlatStyle = "Flat"
    $script:RemoveFromTitleButton.Add_Click({ Remove-BulkPermissionsFromTitle })

    $actionsGroup.Controls.AddRange(@($script:GrantToUserButton, $script:GrantToTitleButton, $script:RemoveFromUserButton, $script:RemoveFromTitleButton))

    # Results Group
    $resultsGroup = New-Object System.Windows.Forms.GroupBox
    $resultsGroup.Text = "Results Log"
    $resultsGroup.Location = New-Object System.Drawing.Point(15, 595)
    $resultsGroup.Size = New-Object System.Drawing.Size(655, 140)

    $script:ResultsTextBox = New-Object System.Windows.Forms.TextBox
    $script:ResultsTextBox.Location = New-Object System.Drawing.Point(15, 25)
    $script:ResultsTextBox.Size = New-Object System.Drawing.Size(625, 100)
    $script:ResultsTextBox.Multiline = $true
    $script:ResultsTextBox.ScrollBars = "Vertical"
    $script:ResultsTextBox.ReadOnly = $true
    $script:ResultsTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $script:ResultsTextBox.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $script:ResultsTextBox.ForeColor = [System.Drawing.Color]::LightGreen

    $resultsGroup.Controls.Add($script:ResultsTextBox)

    # Status Bar
    $statusStrip = New-Object System.Windows.Forms.StatusStrip

    $script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:StatusLabel.Text = "Not connected"
    $script:StatusLabel.Spring = $true
    $script:StatusLabel.TextAlign = "MiddleLeft"

    $script:ProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
    $script:ProgressBar.Size = New-Object System.Drawing.Size(200, 16)
    $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    $statusStrip.Items.AddRange(@($script:StatusLabel, $script:ProgressBar))

    # Add all controls to main form
    $script:MainForm.Controls.AddRange(@(
        $headerPanel,
        $connectionGroup,
        $methodGroup,
        $targetUserGroup,
        $permissionGroup,
        $actionsGroup,
        $resultsGroup,
        $statusStrip
    ))

    # Form closing event
    $script:MainForm.Add_FormClosing({
        if ($script:IsConnected) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "You are still connected. Disconnect before closing?",
                "Confirm Close",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Disconnect-Services
            }
            elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                $_.Cancel = $true
            }
        }
    })
}
#endregion

#region Main Entry Point
# Initialize logging
Initialize-Logging
Write-Log "CalendarWarlock started" "INFO"

# Build and show the form
Build-MainForm
[void]$script:MainForm.ShowDialog()

# Cleanup
Write-Log "CalendarWarlock closed" "INFO"
#endregion

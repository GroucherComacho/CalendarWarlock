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
$script:CSVFilePath = $null
$script:CurrentTheme = "Dark"

# Load required assemblies for Windows Forms and Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide the console window when GUI launches
Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
if ($consolePtr -ne [IntPtr]::Zero) {
    [void][Console.Window]::ShowWindow($consolePtr, 0) # 0 = SW_HIDE
}

# Theme Configurations
$script:Themes = @{
    Dark = @{
        # Dark Theme - Deep purple/violet
        FormBackground = [System.Drawing.Color]::FromArgb(15, 12, 25)           # Deep black-purple
        CardBackground = [System.Drawing.Color]::FromArgb(25, 20, 40)           # Dark purple
        HeaderBackground = [System.Drawing.Color]::FromArgb(45, 25, 70)         # Rich purple header
        HeaderText = [System.Drawing.Color]::FromArgb(220, 200, 255)            # Lavender
        PrimaryText = [System.Drawing.Color]::FromArgb(200, 190, 230)           # Soft white-purple
        SecondaryText = [System.Drawing.Color]::FromArgb(130, 120, 170)         # Muted purple
        PrimaryButton = [System.Drawing.Color]::FromArgb(100, 60, 180)          # Violet
        SecondaryButton = [System.Drawing.Color]::FromArgb(60, 140, 160)        # Teal
        DisabledButton = [System.Drawing.Color]::FromArgb(60, 50, 80)           # Shadowed purple
        RemoveButton = [System.Drawing.Color]::FromArgb(140, 40, 70)            # Crimson
        ButtonText = [System.Drawing.Color]::FromArgb(240, 235, 255)            # White
        ButtonTextLight = [System.Drawing.Color]::FromArgb(200, 190, 230)       # Soft text
        ResultsBackground = [System.Drawing.Color]::FromArgb(18, 15, 30)        # Deep background
        ResultsText = [System.Drawing.Color]::FromArgb(120, 220, 200)           # Teal
        InputBackground = [System.Drawing.Color]::FromArgb(30, 25, 50)          # Input field dark purple
        InputText = [System.Drawing.Color]::FromArgb(200, 190, 230)             # Input text
        AccentGlow = [System.Drawing.Color]::FromArgb(180, 100, 255)            # Purple accent
        SuccessColor = [System.Drawing.Color]::FromArgb(80, 200, 120)           # Green
        WarningColor = [System.Drawing.Color]::FromArgb(255, 180, 80)           # Amber warning
        ErrorColor = [System.Drawing.Color]::FromArgb(220, 60, 90)              # Crimson error
        BorderColor = [System.Drawing.Color]::FromArgb(80, 60, 120)             # Subtle purple border
        ToggleText = "Dark"
    }
    Light = @{
        # Light Theme - Clean light theme with purple accents
        FormBackground = [System.Drawing.Color]::FromArgb(240, 238, 248)        # Soft lavender white
        CardBackground = [System.Drawing.Color]::FromArgb(252, 250, 255)        # White
        HeaderBackground = [System.Drawing.Color]::FromArgb(70, 45, 110)        # Deep amethyst
        HeaderText = [System.Drawing.Color]::FromArgb(255, 250, 255)            # Pure white
        PrimaryText = [System.Drawing.Color]::FromArgb(35, 25, 55)              # Dark purple text
        SecondaryText = [System.Drawing.Color]::FromArgb(100, 85, 130)          # Muted purple
        PrimaryButton = [System.Drawing.Color]::FromArgb(110, 70, 190)          # Vivid purple
        SecondaryButton = [System.Drawing.Color]::FromArgb(60, 150, 170)        # Teal
        DisabledButton = [System.Drawing.Color]::FromArgb(200, 195, 210)        # Light gray-purple
        RemoveButton = [System.Drawing.Color]::FromArgb(180, 60, 90)            # Rose crimson
        ButtonText = [System.Drawing.Color]::FromArgb(255, 255, 255)            # Pure white
        ButtonTextLight = [System.Drawing.Color]::FromArgb(35, 25, 55)          # Dark text for light buttons
        ResultsBackground = [System.Drawing.Color]::FromArgb(248, 246, 255)     # Soft white-purple
        ResultsText = [System.Drawing.Color]::FromArgb(50, 120, 130)            # Deep teal
        InputBackground = [System.Drawing.Color]::FromArgb(255, 255, 255)       # White input
        InputText = [System.Drawing.Color]::FromArgb(35, 25, 55)                # Dark text
        AccentGlow = [System.Drawing.Color]::FromArgb(140, 80, 200)             # Purple accent
        SuccessColor = [System.Drawing.Color]::FromArgb(60, 160, 100)           # Forest green
        WarningColor = [System.Drawing.Color]::FromArgb(200, 140, 50)           # Amber
        ErrorColor = [System.Drawing.Color]::FromArgb(180, 50, 70)              # Crimson
        BorderColor = [System.Drawing.Color]::FromArgb(180, 170, 200)           # Soft purple border
        ToggleText = "Light"
    }
    Warlock = @{
        # Warlock Theme - Deep void with green glow
        FormBackground = [System.Drawing.Color]::FromArgb(5, 5, 10)              # Deep black
        CardBackground = [System.Drawing.Color]::FromArgb(12, 15, 10)            # Dark green-black
        HeaderBackground = [System.Drawing.Color]::FromArgb(10, 30, 15)          # Deep forest
        HeaderText = [System.Drawing.Color]::FromArgb(100, 255, 140)             # Bright green
        PrimaryText = [System.Drawing.Color]::FromArgb(170, 220, 180)            # Pale green
        SecondaryText = [System.Drawing.Color]::FromArgb(90, 140, 100)           # Faded green
        PrimaryButton = [System.Drawing.Color]::FromArgb(30, 120, 60)            # Emerald
        SecondaryButton = [System.Drawing.Color]::FromArgb(120, 50, 160)         # Purple
        DisabledButton = [System.Drawing.Color]::FromArgb(30, 35, 30)            # Dark gray-green
        RemoveButton = [System.Drawing.Color]::FromArgb(160, 30, 30)             # Red
        ButtonText = [System.Drawing.Color]::FromArgb(220, 255, 230)             # Bright white-green
        ButtonTextLight = [System.Drawing.Color]::FromArgb(140, 180, 150)        # Dim green text
        ResultsBackground = [System.Drawing.Color]::FromArgb(3, 8, 5)            # Deep dark background
        ResultsText = [System.Drawing.Color]::FromArgb(50, 255, 100)             # Green terminal
        InputBackground = [System.Drawing.Color]::FromArgb(15, 20, 15)           # Dark input field
        InputText = [System.Drawing.Color]::FromArgb(170, 220, 180)              # Pale green text
        AccentGlow = [System.Drawing.Color]::FromArgb(60, 255, 120)              # Neon green accent
        SuccessColor = [System.Drawing.Color]::FromArgb(40, 220, 80)             # Green
        WarningColor = [System.Drawing.Color]::FromArgb(255, 160, 30)            # Amber
        ErrorColor = [System.Drawing.Color]::FromArgb(255, 40, 40)               # Red
        BorderColor = [System.Drawing.Color]::FromArgb(40, 100, 50)              # Green border
        ToggleText = "Warlock"
    }
}
#endregion

#region Load Required Assemblies and Modules
# Note: Assemblies already loaded above for console hiding; just enable visual styles
[System.Windows.Forms.Application]::EnableVisualStyles()

# Import custom modules with security validation
try {
    # Get absolute path to modules directory
    $modulesPath = [System.IO.Path]::GetFullPath((Join-Path $script:ScriptPath "Modules"))

    # Validate modules directory exists and is within the script directory
    if (-not (Test-Path -Path $modulesPath -PathType Container)) {
        throw "Modules directory not found: $modulesPath"
    }

    # Verify the modules path is within the script directory (prevent path traversal)
    $resolvedScriptPath = [System.IO.Path]::GetFullPath($script:ScriptPath)
    if (-not $modulesPath.StartsWith($resolvedScriptPath)) {
        throw "Security error: Modules path is outside the application directory"
    }

    # Define expected modules with their full paths
    $customModules = @(
        @{ Name = "ExchangeOperations"; Path = Join-Path $modulesPath "ExchangeOperations.psm1" },
        @{ Name = "AzureADOperations"; Path = Join-Path $modulesPath "AzureADOperations.psm1" }
    )

    foreach ($module in $customModules) {
        $modulePath = [System.IO.Path]::GetFullPath($module.Path)

        # Verify module is within the modules directory
        if (-not $modulePath.StartsWith($modulesPath)) {
            throw "Security error: Module path traversal detected for $($module.Name)"
        }

        # Verify module file exists and has correct extension
        if (-not (Test-Path -Path $modulePath -PathType Leaf)) {
            throw "Module file not found: $($module.Name)"
        }

        if ([System.IO.Path]::GetExtension($modulePath) -ne ".psm1") {
            throw "Invalid module file extension for $($module.Name)"
        }

        Import-Module $modulePath -Force
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to load required modules. Please ensure the Modules folder exists and contains valid module files.",
        "CalendarWarlock - Module Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}

# Note: Required module checks (ExchangeOnlineManagement, Microsoft.Graph.Users) are
# handled by the launcher script Start-CalendarWarlock.ps1 to avoid duplicate prompts.
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

#region Security Helper Functions
function Sanitize-CSVValue {
    <#
    .SYNOPSIS
        Sanitizes a string value to prevent CSV formula injection
    .DESCRIPTION
        Escapes values that could be interpreted as formulas by Excel/spreadsheet software.
        Prefixes with a single quote any value starting with =, +, -, @, tab, or carriage return.
    .PARAMETER Value
        The string value to sanitize
    .EXAMPLE
        Sanitize-CSVValue -Value "=cmd|'/C calc'!A0"
        Returns: "'=cmd|'/C calc'!A0"
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Value
    )

    if ([string]::IsNullOrEmpty($Value)) {
        return $Value
    }

    # Characters that can trigger formula interpretation in Excel
    $formulaTriggers = @('=', '+', '-', '@', "`t", "`r", "`n")

    foreach ($trigger in $formulaTriggers) {
        if ($Value.StartsWith($trigger)) {
            # Prefix with single quote to treat as text in Excel
            return "'" + $Value
        }
    }

    return $Value
}

function Test-ValidEmailFormat {
    <#
    .SYNOPSIS
        Validates that a string is a properly formatted email address
    .PARAMETER Email
        The email address to validate
    .EXAMPLE
        Test-ValidEmailFormat -Email "user@contoso.com"
        Returns: $true
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email
    )

    # Standard email regex pattern
    $emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

    return $Email -match $emailPattern
}

function Test-ValidAccessLevel {
    <#
    .SYNOPSIS
        Validates that an access level is one of the allowed values
    .PARAMETER AccessLevel
        The access level to validate
    .EXAMPLE
        Test-ValidAccessLevel -AccessLevel "Reviewer"
        Returns: $true
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$AccessLevel
    )

    $validAccessLevels = @(
        "Owner",
        "PublishingEditor",
        "Editor",
        "PublishingAuthor",
        "Author",
        "NonEditingAuthor",
        "Reviewer",
        "Contributor",
        "AvailabilityOnly",
        "LimitedDetails",
        "None"
    )

    return $validAccessLevels -contains $AccessLevel
}

function Sanitize-ErrorMessage {
    <#
    .SYNOPSIS
        Sanitizes error messages to prevent information disclosure
    .DESCRIPTION
        Removes potentially sensitive information like file paths, server names, etc.
    .PARAMETER ErrorMessage
        The error message to sanitize
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage
    )

    if ([string]::IsNullOrEmpty($ErrorMessage)) {
        return "An error occurred"
    }

    # Remove file paths (Windows and Unix style)
    $sanitized = $ErrorMessage -replace '[A-Za-z]:\\[^:]*\\', '[path]\'
    $sanitized = $sanitized -replace '/[^\s:]+/', '/[path]/'

    # Remove potential connection strings
    $sanitized = $sanitized -replace 'Server=[^;]+;', 'Server=[redacted];'
    $sanitized = $sanitized -replace 'Data Source=[^;]+;', 'Data Source=[redacted];'

    # Remove potential IP addresses
    $sanitized = $sanitized -replace '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b', '[IP]'

    return $sanitized
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
        $script:SingleRadio,
        $script:JobTitleRadio,
        $script:DepartmentRadio,
        $script:BulkCSVRadio,
        $script:SingleCalendarOwnerTextBox,
        $script:SingleUserTextBox,
        $script:JobTitleComboBox,
        $script:DepartmentComboBox,
        $script:RefreshTitlesButton,
        $script:BrowseCSVButton,
        $script:DownloadTemplateButton,
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

    # When enabling, respect the radio button state for controls
    if ($Enabled) {
        Update-MethodSelectionUI
    }
}

function Update-MethodSelectionUI {
    <#
    .SYNOPSIS
        Updates the UI based on the selected method radio button
    #>

    # Single mode
    $singleEnabled = $script:SingleRadio.Checked
    if ($script:SingleCalendarOwnerTextBox) { $script:SingleCalendarOwnerTextBox.Enabled = $singleEnabled }
    if ($script:SingleUserTextBox) { $script:SingleUserTextBox.Enabled = $singleEnabled }
    if ($script:SingleCalendarOwnerLabel) { $script:SingleCalendarOwnerLabel.Enabled = $singleEnabled }
    if ($script:SingleUserLabel) { $script:SingleUserLabel.Enabled = $singleEnabled }

    # Job Title mode
    $jobTitleEnabled = $script:JobTitleRadio.Checked
    if ($script:JobTitleComboBox) { $script:JobTitleComboBox.Enabled = $jobTitleEnabled }

    # Department mode
    $departmentEnabled = $script:DepartmentRadio.Checked
    if ($script:DepartmentComboBox) { $script:DepartmentComboBox.Enabled = $departmentEnabled }

    # Bulk CSV mode
    $csvEnabled = $script:BulkCSVRadio.Checked
    if ($script:BrowseCSVButton) { $script:BrowseCSVButton.Enabled = $csvEnabled }
    if ($script:DownloadTemplateButton) { $script:DownloadTemplateButton.Enabled = $csvEnabled }
    if ($script:CSVFileLabel) { $script:CSVFileLabel.Enabled = $csvEnabled }

    # Target User section visibility based on mode
    $showTargetUser = $script:JobTitleRadio.Checked -or $script:DepartmentRadio.Checked
    if ($script:TargetUserTextBox) { $script:TargetUserTextBox.Enabled = $showTargetUser }
    if ($script:SearchUserButton) { $script:SearchUserButton.Enabled = $showTargetUser }
    if ($script:GetPermissionsButton) { $script:GetPermissionsButton.Enabled = $showTargetUser }

    # Update button labels based on mode
    if ($script:GrantToUserButton -and $script:GrantToTitleButton -and $script:RemoveFromUserButton -and $script:RemoveFromTitleButton) {
        if ($script:SingleRadio.Checked) {
            $script:GrantToUserButton.Text = "Grant Permission"
            $script:GrantToTitleButton.Text = "Grant Permission"
            $script:RemoveFromUserButton.Text = "Remove Permission"
            $script:RemoveFromTitleButton.Text = "Remove Permission"
        }
        elseif ($script:BulkCSVRadio.Checked) {
            $script:GrantToUserButton.Text = "Grant Permissions from CSV"
            $script:GrantToTitleButton.Text = "Grant Permissions from CSV"
            $script:RemoveFromUserButton.Text = "Remove Permissions from CSV"
            $script:RemoveFromTitleButton.Text = "Remove Permissions from CSV"
        }
        else {
            $script:GrantToUserButton.Text = "Grant User Access to All Calendars of Selection"
            $script:GrantToTitleButton.Text = "Grant All of Selection Access to User's Calendar"
            $script:RemoveFromUserButton.Text = "Remove User Access from All Calendars of Selection"
            $script:RemoveFromTitleButton.Text = "Remove All of Selection Access from User's Calendar"
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
        Update-ResultsLog "Connection failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        $script:IsConnected = $false

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

        Update-ResultsLog "Disconnected from all services" "Success"
        Write-Log "Disconnected from all services" "SUCCESS"
    }
    catch {
        Update-ResultsLog "Error during disconnect: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Warning"
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
        Update-ResultsLog "Failed to load job titles: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
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
        Update-ResultsLog "Failed to load departments: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
    }
}

function Refresh-AllSelections {
    Refresh-JobTitles
    Refresh-Departments
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
        Update-ResultsLog "Search failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
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
        Update-ResultsLog "Error getting permissions: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
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

    # Validate email format
    if (-not (Test-ValidEmailFormat -Email $targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the target user.",
            "Invalid Email Format",
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
        if (-not $methodResult.Success) {
            throw $methodResult.Message
        }

        $users = $methodResult.Users
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
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

    # Validate email format
    if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the calendar owner.",
            "Invalid Email Format",
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
        if (-not $methodResult.Success) {
            throw $methodResult.Message
        }

        $users = $methodResult.Users
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
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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
        Gets users based on the currently selected method (Job Title or Department)
    .DESCRIPTION
        Returns users filtered by the selected radio button option
    #>

    # Determine which method to use based on radio button selection
    if ($script:JobTitleRadio.Checked) {
        $jobTitle = $script:JobTitleComboBox.Text.Trim()
        if ([string]::IsNullOrEmpty($jobTitle)) {
            return @{
                Success = $false
                Users = @()
                Count = 0
                Method = $null
                Value = $null
                Message = "Please select a Job Title"
            }
        }
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
    elseif ($script:DepartmentRadio.Checked) {
        $department = $script:DepartmentComboBox.Text.Trim()
        if ([string]::IsNullOrEmpty($department)) {
            return @{
                Success = $false
                Users = @()
                Count = 0
                Method = $null
                Value = $null
                Message = "Please select a Department"
            }
        }
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
    else {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Method = $null
            Value = $null
            Message = "Please select Job Title or Department"
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

    # Validate email format
    if (-not (Test-ValidEmailFormat -Email $targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the target user.",
            "Invalid Email Format",
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
        if (-not $methodResult.Success) {
            throw $methodResult.Message
        }

        $users = $methodResult.Users
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
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

    # Validate email format
    if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the calendar owner.",
            "Invalid Email Format",
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
        if (-not $methodResult.Success) {
            throw $methodResult.Message
        }

        $users = $methodResult.Users
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
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

function Grant-SinglePermission {
    <#
    .SYNOPSIS
        Grants a single user access to a single calendar
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

    $calendarOwner = $script:SingleCalendarOwnerTextBox.Text.Trim()
    $targetUser = $script:SingleUserTextBox.Text.Trim()
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

    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter the user email to grant access to.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if ($calendarOwner -eq $targetUser) {
        [System.Windows.Forms.MessageBox]::Show(
            "Calendar owner and target user cannot be the same.",
            "Invalid Input",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Validate email format
    if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the calendar owner.",
            "Invalid Email Format",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if (-not (Test-ValidEmailFormat -Email $targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the target user.",
            "Invalid Email Format",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Granting single calendar permission..." "Info"
    Update-ResultsLog "Calendar Owner: $calendarOwner" "Info"
    Update-ResultsLog "User: $targetUser" "Info"
    Update-ResultsLog "Permission: $permission" "Info"
    Write-Log "Grant-SinglePermission: Owner=$calendarOwner, User=$targetUser, Permission=$permission" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        $result = Grant-CalendarPermission -CalendarOwner $calendarOwner -Trustee $targetUser -AccessRights $permission

        if ($result.Success) {
            Update-ResultsLog "$($result.Action): $targetUser now has $permission access to $calendarOwner's calendar" "Success"
            Write-Log "Single permission grant completed successfully" "SUCCESS"

            [System.Windows.Forms.MessageBox]::Show(
                "$($result.Action)!`n`n$targetUser now has $permission access to $calendarOwner's calendar.",
                "Operation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        else {
            Update-ResultsLog "Failed: $($result.Message)" "Error"
            Write-Log "Single permission grant failed: $($result.Message)" "ERROR"

            [System.Windows.Forms.MessageBox]::Show(
                "Failed to grant permission: $($result.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    catch {
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
    }
}

function Remove-SinglePermission {
    <#
    .SYNOPSIS
        Removes a single user's access from a single calendar
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

    $calendarOwner = $script:SingleCalendarOwnerTextBox.Text.Trim()
    $targetUser = $script:SingleUserTextBox.Text.Trim()

    if ([string]::IsNullOrEmpty($calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter the calendar owner's email.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if ([string]::IsNullOrEmpty($targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter the user email to remove access from.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Validate email formats
    if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the calendar owner.",
            "Invalid Email Format",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    if (-not (Test-ValidEmailFormat -Email $targetUser)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid email address for the target user.",
            "Invalid Email Format",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Confirmation dialog
    $confirmResult = [System.Windows.Forms.MessageBox]::Show(
        "This will REMOVE $targetUser's access from $calendarOwner's calendar.`n`nAre you sure you want to continue?",
        "Confirm Permission Removal",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Clear-ResultsLog
    Update-ResultsLog "Removing single calendar permission..." "Info"
    Update-ResultsLog "Calendar Owner: $calendarOwner" "Info"
    Update-ResultsLog "User: $targetUser" "Info"
    Write-Log "Remove-SinglePermission: Owner=$calendarOwner, User=$targetUser" "INFO"

    Set-UIEnabled -Enabled $false

    try {
        $result = Remove-CalendarPermission -CalendarOwner $calendarOwner -Trustee $targetUser

        if ($result.Success) {
            Update-ResultsLog "Removed: $targetUser's access from $calendarOwner's calendar" "Success"
            Write-Log "Single permission removal completed successfully" "SUCCESS"

            [System.Windows.Forms.MessageBox]::Show(
                "Permission removed!`n`n$targetUser no longer has access to $calendarOwner's calendar.",
                "Operation Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        else {
            Update-ResultsLog "Failed: $($result.Message)" "Error"
            Write-Log "Single permission removal failed: $($result.Message)" "ERROR"

            [System.Windows.Forms.MessageBox]::Show(
                "Failed to remove permission: $($result.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
    catch {
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        Set-UIEnabled -Enabled $true
    }
}

function Grant-BulkCSVPermissions {
    <#
    .SYNOPSIS
        Grants calendar permissions based on a CSV file
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

    if ([string]::IsNullOrEmpty($script:CSVFilePath) -or -not (Test-Path $script:CSVFilePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a valid CSV file.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        $csvData = Import-Csv -Path $script:CSVFilePath

        # Validate CSV structure
        $requiredColumns = @("MailboxEmail", "UserEmail", "AccessLevel")
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvData[0].PSObject.Properties.Name }

        if ($missingColumns.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file is missing required columns: $($missingColumns -join ', ')`n`nRequired columns: MailboxEmail, UserEmail, AccessLevel",
                "Invalid CSV Format",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        if ($csvData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file contains no data rows.",
                "Empty CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Confirmation dialog
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "This will process $($csvData.Count) permission grant(s) from the CSV file.`n`nAre you sure you want to continue?",
            "Confirm Bulk CSV Grant",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        Clear-ResultsLog
        Update-ResultsLog "Starting bulk CSV permission grant..." "Info"
        Update-ResultsLog "Processing $($csvData.Count) entries from CSV" "Info"
        Write-Log "Grant-BulkCSVPermissions: Processing $($csvData.Count) entries from $($script:CSVFilePath)" "INFO"

        Set-UIEnabled -Enabled $false

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $csvData.Count

        for ($i = 0; $i -lt $csvData.Count; $i++) {
            $entry = $csvData[$i]
            # Trim CSV values (Sanitize-CSVValue is only for OUTPUT/export, not input processing)
            $calendarOwner = $entry.MailboxEmail.Trim()
            $targetUser = $entry.UserEmail.Trim()
            $permission = $entry.AccessLevel.Trim()

            if ([string]::IsNullOrEmpty($calendarOwner) -or [string]::IsNullOrEmpty($targetUser) -or [string]::IsNullOrEmpty($permission)) {
                Update-ResultsLog "Skipping row $($i + 1): Missing required data" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            if ($calendarOwner -eq $targetUser) {
                Update-ResultsLog "Skipping row $($i + 1): Same user for owner and target" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            # Validate email formats
            if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
                Update-ResultsLog "Skipping row $($i + 1): Invalid email format for MailboxEmail '$calendarOwner'" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            if (-not (Test-ValidEmailFormat -Email $targetUser)) {
                Update-ResultsLog "Skipping row $($i + 1): Invalid email format for UserEmail '$targetUser'" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            # Validate access level
            if (-not (Test-ValidAccessLevel -AccessLevel $permission)) {
                Update-ResultsLog "Skipping row $($i + 1): Invalid AccessLevel '$permission'. Valid values: None, AvailabilityOnly, LimitedDetails, Reviewer, Editor, Author, PublishingAuthor, PublishingEditor" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            Update-ResultsLog "Granting $targetUser $permission access to $calendarOwner's calendar..." "Info"

            $result = Grant-CalendarPermission -CalendarOwner $calendarOwner -Trustee $targetUser -AccessRights $permission

            if ($result.Success) {
                Update-ResultsLog "$($result.Action): $targetUser -> $calendarOwner ($permission)" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $targetUser -> $calendarOwner - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk CSV operation completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk CSV grant completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk CSV permission grant completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

function Remove-BulkCSVPermissions {
    <#
    .SYNOPSIS
        Removes calendar permissions based on a CSV file
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

    if ([string]::IsNullOrEmpty($script:CSVFilePath) -or -not (Test-Path $script:CSVFilePath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a valid CSV file.",
            "Missing Information",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    try {
        $csvData = Import-Csv -Path $script:CSVFilePath

        # Validate CSV structure (only need MailboxEmail and UserEmail for removal)
        $requiredColumns = @("MailboxEmail", "UserEmail")
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvData[0].PSObject.Properties.Name }

        if ($missingColumns.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file is missing required columns: $($missingColumns -join ', ')`n`nRequired columns: MailboxEmail, UserEmail",
                "Invalid CSV Format",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        if ($csvData.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file contains no data rows.",
                "Empty CSV",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Confirmation dialog
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "This will REMOVE $($csvData.Count) permission(s) based on the CSV file.`n`nAre you sure you want to continue?",
            "Confirm Bulk CSV Removal",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirmResult -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }

        Clear-ResultsLog
        Update-ResultsLog "Starting bulk CSV permission removal..." "Info"
        Update-ResultsLog "Processing $($csvData.Count) entries from CSV" "Info"
        Write-Log "Remove-BulkCSVPermissions: Processing $($csvData.Count) entries from $($script:CSVFilePath)" "INFO"

        Set-UIEnabled -Enabled $false

        $successCount = 0
        $failCount = 0
        $skipCount = 0

        Update-ProgressBar -Value 0 -Maximum $csvData.Count

        for ($i = 0; $i -lt $csvData.Count; $i++) {
            $entry = $csvData[$i]
            # Trim CSV values (Sanitize-CSVValue is only for OUTPUT/export, not input processing)
            $calendarOwner = $entry.MailboxEmail.Trim()
            $targetUser = $entry.UserEmail.Trim()

            if ([string]::IsNullOrEmpty($calendarOwner) -or [string]::IsNullOrEmpty($targetUser)) {
                Update-ResultsLog "Skipping row $($i + 1): Missing required data" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            # Validate email formats
            if (-not (Test-ValidEmailFormat -Email $calendarOwner)) {
                Update-ResultsLog "Skipping row $($i + 1): Invalid email format for MailboxEmail '$calendarOwner'" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            if (-not (Test-ValidEmailFormat -Email $targetUser)) {
                Update-ResultsLog "Skipping row $($i + 1): Invalid email format for UserEmail '$targetUser'" "Warning"
                $skipCount++
                Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
                continue
            }

            Update-ResultsLog "Removing $targetUser's access from $calendarOwner's calendar..." "Info"

            $result = Remove-CalendarPermission -CalendarOwner $calendarOwner -Trustee $targetUser

            if ($result.Success) {
                Update-ResultsLog "Removed: $targetUser from $calendarOwner" "Success"
                $successCount++
            }
            else {
                Update-ResultsLog "Failed: $targetUser from $calendarOwner - $($result.Message)" "Error"
                $failCount++
            }

            Update-ProgressBar -Value ($i + 1) -Maximum $csvData.Count
        }

        Update-ResultsLog "-----------------------------------" "Info"
        Update-ResultsLog "Bulk CSV removal completed!" "Success"
        Update-ResultsLog "Success: $successCount | Failed: $failCount | Skipped: $skipCount" "Info"
        Write-Log "Bulk CSV removal completed: Success=$successCount, Failed=$failCount, Skipped=$skipCount" "SUCCESS"

        [System.Windows.Forms.MessageBox]::Show(
            "Bulk CSV permission removal completed!`n`nSuccess: $successCount`nFailed: $failCount`nSkipped: $skipCount",
            "Operation Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Update-ResultsLog "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"
        Write-Log "Operation failed: $($_.Exception.Message)" "ERROR"

        [System.Windows.Forms.MessageBox]::Show(
            "Operation failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)",
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

#region Theme Functions
function Apply-Theme {
    $theme = $script:Themes[$script:CurrentTheme]

    # Main Form
    $script:MainForm.BackColor = $theme.FormBackground

    # Header Panel
    $script:HeaderPanel.BackColor = $theme.HeaderBackground
    $script:TitleLabel.ForeColor = $theme.HeaderText
    $script:SubtitleLabel.ForeColor = $theme.HeaderText
    $script:ThemeToggleButton.Text = $theme.ToggleText
    $script:ThemeToggleButton.BackColor = $theme.CardBackground
    $script:ThemeToggleButton.ForeColor = $theme.PrimaryText
    $script:ThemeToggleButton.FlatAppearance.BorderColor = $theme.BorderColor
    $script:ThemeToggleButton.FlatAppearance.BorderSize = 1

    # Group Boxes
    foreach ($group in @($script:ConnectionGroup, $script:MethodGroup, $script:TargetUserGroup, $script:PermissionGroup, $script:ActionsGroup, $script:ResultsGroup)) {
        $group.BackColor = $theme.CardBackground
        $group.ForeColor = $theme.PrimaryText
    }

    # Style all TextBoxes
    foreach ($textBox in @($script:OrganizationTextBox, $script:TargetUserTextBox, $script:SingleCalendarOwnerTextBox, $script:SingleUserTextBox)) {
        if ($null -ne $textBox) {
            $textBox.BackColor = $theme.InputBackground
            $textBox.ForeColor = $theme.InputText
            $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        }
    }

    # Style ComboBoxes
    foreach ($comboBox in @($script:JobTitleComboBox, $script:DepartmentComboBox, $script:PermissionComboBox)) {
        if ($null -ne $comboBox) {
            $comboBox.BackColor = $theme.InputBackground
            $comboBox.ForeColor = $theme.InputText
        }
    }

    # Primary Action Buttons
    $script:ConnectButton.BackColor = $theme.PrimaryButton
    $script:ConnectButton.ForeColor = $theme.ButtonText
    $script:ConnectButton.FlatAppearance.BorderColor = $theme.AccentGlow
    $script:ConnectButton.FlatAppearance.BorderSize = 1
    $script:ConnectButton.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(
        [Math]::Min(255, $theme.PrimaryButton.R + 25),
        [Math]::Min(255, $theme.PrimaryButton.G + 25),
        [Math]::Min(255, $theme.PrimaryButton.B + 25)
    )

    $script:GrantToUserButton.BackColor = $theme.PrimaryButton
    $script:GrantToUserButton.ForeColor = $theme.ButtonText
    $script:GrantToUserButton.FlatAppearance.BorderColor = $theme.AccentGlow
    $script:GrantToUserButton.FlatAppearance.BorderSize = 1

    $script:GrantToTitleButton.BackColor = $theme.SecondaryButton
    $script:GrantToTitleButton.ForeColor = $theme.ButtonText
    $script:GrantToTitleButton.FlatAppearance.BorderColor = $theme.AccentGlow
    $script:GrantToTitleButton.FlatAppearance.BorderSize = 1

    # Secondary/Remove Buttons
    $script:RemoveFromUserButton.BackColor = $theme.DisabledButton
    $script:RemoveFromUserButton.ForeColor = $theme.ButtonTextLight
    $script:RemoveFromUserButton.FlatAppearance.BorderColor = $theme.BorderColor
    $script:RemoveFromUserButton.FlatAppearance.BorderSize = 1

    $script:RemoveFromTitleButton.BackColor = $theme.RemoveButton
    $script:RemoveFromTitleButton.ForeColor = $theme.ButtonText
    $script:RemoveFromTitleButton.FlatAppearance.BorderColor = $theme.ErrorColor
    $script:RemoveFromTitleButton.FlatAppearance.BorderSize = 1

    # Style utility buttons
    foreach ($btn in @($script:SearchUserButton, $script:GetPermissionsButton, $script:RefreshTitlesButton, $script:BrowseCSVButton, $script:DownloadTemplateButton)) {
        if ($null -ne $btn) {
            $btn.BackColor = $theme.CardBackground
            $btn.ForeColor = $theme.PrimaryText
            $btn.FlatStyle = "Flat"
            $btn.FlatAppearance.BorderColor = $theme.BorderColor
            $btn.FlatAppearance.BorderSize = 1
        }
    }

    # Results TextBox
    $script:ResultsTextBox.BackColor = $theme.ResultsBackground
    $script:ResultsTextBox.ForeColor = $theme.ResultsText
    $script:ResultsTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    # Secondary text labels
    $script:CSVFileLabel.ForeColor = $theme.SecondaryText
    $script:PermissionDescLabel.ForeColor = $theme.SecondaryText

    # Radio buttons styling
    foreach ($radio in @($script:SingleRadio, $script:JobTitleRadio, $script:DepartmentRadio, $script:BulkCSVRadio)) {
        if ($null -ne $radio) {
            $radio.ForeColor = $theme.PrimaryText
        }
    }

    # Refresh the form
    $script:MainForm.Refresh()
}

function Toggle-Theme {
    # Cycle through: Dark -> Light -> Warlock -> Dark
    switch ($script:CurrentTheme) {
        "Dark"    { $script:CurrentTheme = "Light" }
        "Light"   { $script:CurrentTheme = "Warlock" }
        "Warlock" { $script:CurrentTheme = "Dark" }
        default   { $script:CurrentTheme = "Dark" }
    }
    Apply-Theme
}
#endregion

#region Build Main Form
function Build-MainForm {
    # Main Form
    $script:MainForm = New-Object System.Windows.Forms.Form
    $script:MainForm.Text = "CalendarWarlock - Bulk Calendar Permission Manager"
    $script:MainForm.Size = New-Object System.Drawing.Size(700, 885)
    $script:MainForm.StartPosition = "CenterScreen"
    $script:MainForm.FormBorderStyle = "FixedSingle"
    $script:MainForm.MaximizeBox = $false
    $script:MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:MainForm.BackColor = $script:Themes[$script:CurrentTheme].FormBackground

    # Set application icon
    $iconPath = Join-Path (Split-Path -Parent $script:ScriptPath) "icon.ico"
    if (Test-Path $iconPath) {
        $script:MainForm.Icon = New-Object System.Drawing.Icon($iconPath)
    }

    # Header Panel
    $script:HeaderPanel = New-Object System.Windows.Forms.Panel
    $script:HeaderPanel.Location = New-Object System.Drawing.Point(0, 0)
    $script:HeaderPanel.Size = New-Object System.Drawing.Size(700, 110)
    $script:HeaderPanel.BackColor = $script:Themes[$script:CurrentTheme].HeaderBackground

    # Logo PictureBox
    $script:LogoPictureBox = New-Object System.Windows.Forms.PictureBox
    $script:LogoPictureBox.Location = New-Object System.Drawing.Point(8, 6)
    $script:LogoPictureBox.Size = New-Object System.Drawing.Size(98, 98)
    $script:LogoPictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $logoPath = Join-Path (Split-Path -Parent $script:ScriptPath) "icon.png"
    if (Test-Path $logoPath) {
        $script:LogoPictureBox.Image = [System.Drawing.Image]::FromFile($logoPath)
    }

    $script:TitleLabel = New-Object System.Windows.Forms.Label
    $script:TitleLabel.Text = "CalendarWarlock"
    $script:TitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
    $script:TitleLabel.ForeColor = $script:Themes[$script:CurrentTheme].HeaderText
    $script:TitleLabel.Location = New-Object System.Drawing.Point(115, 22)
    $script:TitleLabel.AutoSize = $true

    $script:SubtitleLabel = New-Object System.Windows.Forms.Label
    $script:SubtitleLabel.Text = "Manage Bulk Calendar Permissions for Exchange Online"
    $script:SubtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $script:SubtitleLabel.ForeColor = $script:Themes[$script:CurrentTheme].HeaderText
    $script:SubtitleLabel.Location = New-Object System.Drawing.Point(120, 58)
    $script:SubtitleLabel.AutoSize = $true

    # Theme Toggle Button
    $script:ThemeToggleButton = New-Object System.Windows.Forms.Button
    $script:ThemeToggleButton.Text = $script:Themes[$script:CurrentTheme].ToggleText
    $script:ThemeToggleButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $script:ThemeToggleButton.Location = New-Object System.Drawing.Point(600, 38)
    $script:ThemeToggleButton.Size = New-Object System.Drawing.Size(80, 32)
    $script:ThemeToggleButton.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:ThemeToggleButton.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText
    $script:ThemeToggleButton.FlatStyle = "Flat"
    $script:ThemeToggleButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].BorderColor
    $script:ThemeToggleButton.FlatAppearance.BorderSize = 1
    $script:ThemeToggleButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:ThemeToggleButton.Add_Click({ Toggle-Theme })

    $script:HeaderPanel.Controls.AddRange(@($script:LogoPictureBox, $script:TitleLabel, $script:SubtitleLabel, $script:ThemeToggleButton))

    # Connection Group
    $script:ConnectionGroup = New-Object System.Windows.Forms.GroupBox
    $script:ConnectionGroup.Text = "Connection"
    $script:ConnectionGroup.Location = New-Object System.Drawing.Point(15, 120)
    $script:ConnectionGroup.Size = New-Object System.Drawing.Size(655, 70)
    $script:ConnectionGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:ConnectionGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

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
    $script:ConnectButton.Location = New-Object System.Drawing.Point(470, 23)
    $script:ConnectButton.Size = New-Object System.Drawing.Size(110, 32)
    $script:ConnectButton.BackColor = $script:Themes[$script:CurrentTheme].PrimaryButton
    $script:ConnectButton.ForeColor = $script:Themes[$script:CurrentTheme].ButtonText
    $script:ConnectButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:ConnectButton.FlatStyle = "Flat"
    $script:ConnectButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].AccentGlow
    $script:ConnectButton.FlatAppearance.BorderSize = 1
    $script:ConnectButton.Cursor = [System.Windows.Forms.Cursors]::Hand
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

    $script:ConnectionGroup.Controls.AddRange(@($orgLabel, $script:OrganizationTextBox, $script:ConnectButton))

    # Selection Method Group
    $script:MethodGroup = New-Object System.Windows.Forms.GroupBox
    $script:MethodGroup.Text = "Selection Method"
    $script:MethodGroup.Location = New-Object System.Drawing.Point(15, 200)
    $script:MethodGroup.Size = New-Object System.Drawing.Size(655, 160)
    $script:MethodGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:MethodGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

    # Radio buttons for method selection
    $script:SingleRadio = New-Object System.Windows.Forms.RadioButton
    $script:SingleRadio.Text = "Single"
    $script:SingleRadio.Location = New-Object System.Drawing.Point(15, 25)
    $script:SingleRadio.Size = New-Object System.Drawing.Size(80, 20)
    $script:SingleRadio.Checked = $true
    $script:SingleRadio.Add_CheckedChanged({
        Update-MethodSelectionUI
    })

    $script:JobTitleRadio = New-Object System.Windows.Forms.RadioButton
    $script:JobTitleRadio.Text = "Job Title"
    $script:JobTitleRadio.Location = New-Object System.Drawing.Point(15, 55)
    $script:JobTitleRadio.Size = New-Object System.Drawing.Size(80, 20)
    $script:JobTitleRadio.Add_CheckedChanged({
        Update-MethodSelectionUI
    })

    $script:DepartmentRadio = New-Object System.Windows.Forms.RadioButton
    $script:DepartmentRadio.Text = "Dept."
    $script:DepartmentRadio.Location = New-Object System.Drawing.Point(15, 85)
    $script:DepartmentRadio.Size = New-Object System.Drawing.Size(80, 20)
    $script:DepartmentRadio.Add_CheckedChanged({
        Update-MethodSelectionUI
    })

    $script:BulkCSVRadio = New-Object System.Windows.Forms.RadioButton
    $script:BulkCSVRadio.Text = "Bulk CSV"
    $script:BulkCSVRadio.Location = New-Object System.Drawing.Point(15, 115)
    $script:BulkCSVRadio.Size = New-Object System.Drawing.Size(80, 20)
    $script:BulkCSVRadio.Add_CheckedChanged({
        Update-MethodSelectionUI
    })

    # Single mode controls - Calendar Owner Email
    $script:SingleCalendarOwnerLabel = New-Object System.Windows.Forms.Label
    $script:SingleCalendarOwnerLabel.Text = "Calendar Owner:"
    $script:SingleCalendarOwnerLabel.Location = New-Object System.Drawing.Point(105, 25)
    $script:SingleCalendarOwnerLabel.Size = New-Object System.Drawing.Size(100, 20)

    $script:SingleCalendarOwnerTextBox = New-Object System.Windows.Forms.TextBox
    $script:SingleCalendarOwnerTextBox.Location = New-Object System.Drawing.Point(210, 23)
    $script:SingleCalendarOwnerTextBox.Size = New-Object System.Drawing.Size(240, 23)
    try { $script:SingleCalendarOwnerTextBox.PlaceholderText = "owner@contoso.com" } catch {}

    # Single mode controls - User Email
    $script:SingleUserLabel = New-Object System.Windows.Forms.Label
    $script:SingleUserLabel.Text = "User Email:"
    $script:SingleUserLabel.Location = New-Object System.Drawing.Point(470, 25)
    $script:SingleUserLabel.Size = New-Object System.Drawing.Size(70, 20)

    $script:SingleUserTextBox = New-Object System.Windows.Forms.TextBox
    $script:SingleUserTextBox.Location = New-Object System.Drawing.Point(545, 23)
    $script:SingleUserTextBox.Size = New-Object System.Drawing.Size(95, 23)
    try { $script:SingleUserTextBox.PlaceholderText = "user@contoso.com" } catch {}

    # Job Title controls
    $script:JobTitleComboBox = New-Object System.Windows.Forms.ComboBox
    $script:JobTitleComboBox.Location = New-Object System.Drawing.Point(105, 53)
    $script:JobTitleComboBox.Size = New-Object System.Drawing.Size(345, 23)
    $script:JobTitleComboBox.DropDownStyle = "DropDown"
    $script:JobTitleComboBox.AutoCompleteMode = "SuggestAppend"
    $script:JobTitleComboBox.AutoCompleteSource = "ListItems"
    $script:JobTitleComboBox.Enabled = $false

    # Department controls
    $script:DepartmentComboBox = New-Object System.Windows.Forms.ComboBox
    $script:DepartmentComboBox.Location = New-Object System.Drawing.Point(105, 83)
    $script:DepartmentComboBox.Size = New-Object System.Drawing.Size(345, 23)
    $script:DepartmentComboBox.DropDownStyle = "DropDown"
    $script:DepartmentComboBox.AutoCompleteMode = "SuggestAppend"
    $script:DepartmentComboBox.AutoCompleteSource = "ListItems"
    $script:DepartmentComboBox.Enabled = $false

    $script:RefreshTitlesButton = New-Object System.Windows.Forms.Button
    $script:RefreshTitlesButton.Text = "Refresh"
    $script:RefreshTitlesButton.Location = New-Object System.Drawing.Point(470, 68)
    $script:RefreshTitlesButton.Size = New-Object System.Drawing.Size(100, 28)
    $script:RefreshTitlesButton.Add_Click({ Refresh-AllSelections })

    # Bulk CSV controls
    $script:CSVFileLabel = New-Object System.Windows.Forms.Label
    $script:CSVFileLabel.Text = "No file selected"
    $script:CSVFileLabel.Location = New-Object System.Drawing.Point(105, 117)
    $script:CSVFileLabel.Size = New-Object System.Drawing.Size(240, 20)
    $script:CSVFileLabel.ForeColor = $script:Themes[$script:CurrentTheme].SecondaryText
    $script:CSVFileLabel.Enabled = $false

    $script:BrowseCSVButton = New-Object System.Windows.Forms.Button
    $script:BrowseCSVButton.Text = "Browse..."
    $script:BrowseCSVButton.Location = New-Object System.Drawing.Point(355, 113)
    $script:BrowseCSVButton.Size = New-Object System.Drawing.Size(80, 25)
    $script:BrowseCSVButton.Enabled = $false
    $script:BrowseCSVButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select CSV File"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:CSVFilePath = $openFileDialog.FileName
            $script:CSVFileLabel.Text = [System.IO.Path]::GetFileName($script:CSVFilePath)
            $script:CSVFileLabel.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText
        }
    })

    $script:DownloadTemplateButton = New-Object System.Windows.Forms.Button
    $script:DownloadTemplateButton.Text = "Download Template"
    $script:DownloadTemplateButton.Location = New-Object System.Drawing.Point(445, 113)
    $script:DownloadTemplateButton.Size = New-Object System.Drawing.Size(125, 25)
    $script:DownloadTemplateButton.Enabled = $false
    $script:DownloadTemplateButton.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $saveFileDialog.FileName = "bulktemplate.csv"
        $saveFileDialog.Title = "Save CSV Template"
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $templateContent = "MailboxEmail,UserEmail,AccessLevel`r`nuser1@domain.com,assistant1@domain.com,Editor`r`nuser2@domain.com,assistant2@domain.com,Reviewer"
            [System.IO.File]::WriteAllText($saveFileDialog.FileName, $templateContent)
            Update-ResultsLog "Template saved to: $($saveFileDialog.FileName)" "Success"
        }
    })

    $script:MethodGroup.Controls.AddRange(@(
        $script:SingleRadio, $script:JobTitleRadio, $script:DepartmentRadio, $script:BulkCSVRadio,
        $script:SingleCalendarOwnerLabel, $script:SingleCalendarOwnerTextBox, $script:SingleUserLabel, $script:SingleUserTextBox,
        $script:JobTitleComboBox, $script:DepartmentComboBox, $script:RefreshTitlesButton,
        $script:CSVFileLabel, $script:BrowseCSVButton, $script:DownloadTemplateButton
    ))

    # Target User Group
    $script:TargetUserGroup = New-Object System.Windows.Forms.GroupBox
    $script:TargetUserGroup.Text = "Target User"
    $script:TargetUserGroup.Location = New-Object System.Drawing.Point(15, 370)
    $script:TargetUserGroup.Size = New-Object System.Drawing.Size(655, 70)
    $script:TargetUserGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:TargetUserGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

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

    $script:TargetUserGroup.Controls.AddRange(@($targetUserLabel, $script:TargetUserTextBox, $script:SearchUserButton, $script:GetPermissionsButton))

    # Permission Level Group
    $script:PermissionGroup = New-Object System.Windows.Forms.GroupBox
    $script:PermissionGroup.Text = "Permission Level"
    $script:PermissionGroup.Location = New-Object System.Drawing.Point(15, 450)
    $script:PermissionGroup.Size = New-Object System.Drawing.Size(655, 70)
    $script:PermissionGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:PermissionGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

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
    $script:PermissionDescLabel.ForeColor = $script:Themes[$script:CurrentTheme].SecondaryText

    $script:PermissionComboBox.Add_SelectedIndexChanged({
        $selectedLevel = $script:PermissionComboBox.SelectedItem
        $levels = Get-CalendarPermissionLevels
        $desc = ($levels | Where-Object { $_.Name -eq $selectedLevel }).Description
        $script:PermissionDescLabel.Text = $desc
    })

    # Trigger initial description update
    $script:PermissionDescLabel.Text = ($permissionLevels | Where-Object { $_.Name -eq "Reviewer" }).Description

    $script:PermissionGroup.Controls.AddRange(@($permissionLabel, $script:PermissionComboBox, $script:PermissionDescLabel))

    # Actions Group
    $script:ActionsGroup = New-Object System.Windows.Forms.GroupBox
    $script:ActionsGroup.Text = "Bulk Operations"
    $script:ActionsGroup.Location = New-Object System.Drawing.Point(15, 530)
    $script:ActionsGroup.Size = New-Object System.Drawing.Size(655, 120)
    $script:ActionsGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:ActionsGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

    $script:GrantToUserButton = New-Object System.Windows.Forms.Button
    $script:GrantToUserButton.Text = "Grant User Access to Calendars"
    $script:GrantToUserButton.Location = New-Object System.Drawing.Point(15, 25)
    $script:GrantToUserButton.Size = New-Object System.Drawing.Size(305, 38)
    $script:GrantToUserButton.BackColor = $script:Themes[$script:CurrentTheme].PrimaryButton
    $script:GrantToUserButton.ForeColor = $script:Themes[$script:CurrentTheme].ButtonText
    $script:GrantToUserButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:GrantToUserButton.FlatStyle = "Flat"
    $script:GrantToUserButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].AccentGlow
    $script:GrantToUserButton.FlatAppearance.BorderSize = 1
    $script:GrantToUserButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:GrantToUserButton.Add_Click({
        if ($script:SingleRadio.Checked) {
            Grant-SinglePermission
        }
        elseif ($script:BulkCSVRadio.Checked) {
            Grant-BulkCSVPermissions
        }
        else {
            Grant-BulkPermissionsToUser
        }
    })

    $script:GrantToTitleButton = New-Object System.Windows.Forms.Button
    $script:GrantToTitleButton.Text = "Grant Selection Access to Calendar"
    $script:GrantToTitleButton.Location = New-Object System.Drawing.Point(335, 25)
    $script:GrantToTitleButton.Size = New-Object System.Drawing.Size(305, 38)
    $script:GrantToTitleButton.BackColor = $script:Themes[$script:CurrentTheme].SecondaryButton
    $script:GrantToTitleButton.ForeColor = $script:Themes[$script:CurrentTheme].ButtonText
    $script:GrantToTitleButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $script:GrantToTitleButton.FlatStyle = "Flat"
    $script:GrantToTitleButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].AccentGlow
    $script:GrantToTitleButton.FlatAppearance.BorderSize = 1
    $script:GrantToTitleButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:GrantToTitleButton.Add_Click({
        if ($script:SingleRadio.Checked) {
            Grant-SinglePermission
        }
        elseif ($script:BulkCSVRadio.Checked) {
            Grant-BulkCSVPermissions
        }
        else {
            Grant-BulkPermissionsToTitle
        }
    })

    $script:RemoveFromUserButton = New-Object System.Windows.Forms.Button
    $script:RemoveFromUserButton.Text = "Remove User Access from Calendars"
    $script:RemoveFromUserButton.Location = New-Object System.Drawing.Point(15, 72)
    $script:RemoveFromUserButton.Size = New-Object System.Drawing.Size(305, 38)
    $script:RemoveFromUserButton.BackColor = $script:Themes[$script:CurrentTheme].DisabledButton
    $script:RemoveFromUserButton.ForeColor = $script:Themes[$script:CurrentTheme].ButtonTextLight
    $script:RemoveFromUserButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:RemoveFromUserButton.FlatStyle = "Flat"
    $script:RemoveFromUserButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].BorderColor
    $script:RemoveFromUserButton.FlatAppearance.BorderSize = 1
    $script:RemoveFromUserButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:RemoveFromUserButton.Add_Click({
        if ($script:SingleRadio.Checked) {
            Remove-SinglePermission
        }
        elseif ($script:BulkCSVRadio.Checked) {
            Remove-BulkCSVPermissions
        }
        else {
            Remove-BulkPermissionsFromUser
        }
    })

    $script:RemoveFromTitleButton = New-Object System.Windows.Forms.Button
    $script:RemoveFromTitleButton.Text = "Remove Selection Access from Calendar"
    $script:RemoveFromTitleButton.Location = New-Object System.Drawing.Point(335, 72)
    $script:RemoveFromTitleButton.Size = New-Object System.Drawing.Size(305, 38)
    $script:RemoveFromTitleButton.BackColor = $script:Themes[$script:CurrentTheme].RemoveButton
    $script:RemoveFromTitleButton.ForeColor = $script:Themes[$script:CurrentTheme].ButtonText
    $script:RemoveFromTitleButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:RemoveFromTitleButton.FlatStyle = "Flat"
    $script:RemoveFromTitleButton.FlatAppearance.BorderColor = $script:Themes[$script:CurrentTheme].ErrorColor
    $script:RemoveFromTitleButton.FlatAppearance.BorderSize = 1
    $script:RemoveFromTitleButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $script:RemoveFromTitleButton.Add_Click({
        if ($script:SingleRadio.Checked) {
            Remove-SinglePermission
        }
        elseif ($script:BulkCSVRadio.Checked) {
            Remove-BulkCSVPermissions
        }
        else {
            Remove-BulkPermissionsFromTitle
        }
    })

    $script:ActionsGroup.Controls.AddRange(@($script:GrantToUserButton, $script:GrantToTitleButton, $script:RemoveFromUserButton, $script:RemoveFromTitleButton))

    # Results Group
    $script:ResultsGroup = New-Object System.Windows.Forms.GroupBox
    $script:ResultsGroup.Text = "Operation Log"
    $script:ResultsGroup.Location = New-Object System.Drawing.Point(15, 660)
    $script:ResultsGroup.Size = New-Object System.Drawing.Size(655, 140)
    $script:ResultsGroup.BackColor = $script:Themes[$script:CurrentTheme].CardBackground
    $script:ResultsGroup.ForeColor = $script:Themes[$script:CurrentTheme].PrimaryText

    $script:ResultsTextBox = New-Object System.Windows.Forms.TextBox
    $script:ResultsTextBox.Location = New-Object System.Drawing.Point(15, 25)
    $script:ResultsTextBox.Size = New-Object System.Drawing.Size(625, 100)
    $script:ResultsTextBox.Multiline = $true
    $script:ResultsTextBox.ScrollBars = "Vertical"
    $script:ResultsTextBox.ReadOnly = $true
    $script:ResultsTextBox.Font = New-Object System.Drawing.Font("Cascadia Code", 9)
    $script:ResultsTextBox.BackColor = $script:Themes[$script:CurrentTheme].ResultsBackground
    $script:ResultsTextBox.ForeColor = $script:Themes[$script:CurrentTheme].ResultsText

    $script:ResultsGroup.Controls.Add($script:ResultsTextBox)

    # Status Bar
    $statusStrip = New-Object System.Windows.Forms.StatusStrip

    $script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:StatusLabel.Text = "Disconnected - Click Connect to begin"
    $script:StatusLabel.Spring = $true
    $script:StatusLabel.TextAlign = "MiddleLeft"

    $script:ProgressBar = New-Object System.Windows.Forms.ToolStripProgressBar
    $script:ProgressBar.Size = New-Object System.Drawing.Size(200, 16)
    $script:ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    $statusStrip.Items.AddRange(@($script:StatusLabel, $script:ProgressBar))

    # Add all controls to main form
    $script:MainForm.Controls.AddRange(@(
        $script:HeaderPanel,
        $script:ConnectionGroup,
        $script:MethodGroup,
        $script:TargetUserGroup,
        $script:PermissionGroup,
        $script:ActionsGroup,
        $script:ResultsGroup,
        $statusStrip
    ))

    # Form closing event
    $script:MainForm.Add_FormClosing({
        if ($script:IsConnected) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "You are still connected. Disconnect before closing?",
                "Confirm Exit",
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

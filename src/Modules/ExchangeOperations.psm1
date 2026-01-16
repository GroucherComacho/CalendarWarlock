#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Exchange Online Calendar Operations Module for CalendarWarlock
.DESCRIPTION
    Provides functions for managing Exchange Online calendar permissions in bulk.
    Requires the ExchangeOnlineManagement module to be installed.
.NOTES
    Author: CalendarWarlock
    Requires: PowerShell 5.1+, ExchangeOnlineManagement module
#>

# Module-level variable to track connection status
$script:IsConnected = $false

function Connect-ExchangeOnlineSession {
    <#
    .SYNOPSIS
        Connects to Exchange Online using modern authentication
    .DESCRIPTION
        Establishes a connection to Exchange Online. Uses interactive modern auth
        which supports MFA and is the most secure method.
    .PARAMETER Organization
        The organization domain (e.g., contoso.onmicrosoft.com)
    .EXAMPLE
        Connect-ExchangeOnlineSession -Organization "contoso.onmicrosoft.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Organization
    )

    try {
        # Check if already connected
        $existingSession = Get-PSSession | Where-Object {
            $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"
        }

        if ($existingSession) {
            Write-Verbose "Already connected to Exchange Online"
            $script:IsConnected = $true
            return @{
                Success = $true
                Message = "Already connected to Exchange Online"
            }
        }

        # Connect using modern authentication (interactive)
        Connect-ExchangeOnline -Organization $Organization -ShowBanner:$false
        $script:IsConnected = $true

        return @{
            Success = $true
            Message = "Successfully connected to Exchange Online"
        }
    }
    catch {
        $script:IsConnected = $false
        return @{
            Success = $false
            Message = "Failed to connect: $($_.Exception.Message)"
        }
    }
}

function Disconnect-ExchangeOnlineSession {
    <#
    .SYNOPSIS
        Disconnects from Exchange Online
    #>
    [CmdletBinding()]
    param()

    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        $script:IsConnected = $false
        return @{
            Success = $true
            Message = "Disconnected from Exchange Online"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error disconnecting: $($_.Exception.Message)"
        }
    }
}

function Test-ExchangeConnection {
    <#
    .SYNOPSIS
        Tests if connected to Exchange Online
    #>
    [CmdletBinding()]
    param()

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        $script:IsConnected = $true
        return $true
    }
    catch {
        $script:IsConnected = $false
        return $false
    }
}

function Get-CalendarPermissionLevels {
    <#
    .SYNOPSIS
        Returns available calendar permission levels
    #>
    [CmdletBinding()]
    param()

    return @(
        @{ Name = "Owner"; Description = "Full control - read, create, modify, delete all items and manage permissions" }
        @{ Name = "PublishingEditor"; Description = "Create, read, modify, delete all items; create subfolders" }
        @{ Name = "Editor"; Description = "Create, read, modify, delete all items" }
        @{ Name = "PublishingAuthor"; Description = "Create, read items; modify, delete own items; create subfolders" }
        @{ Name = "Author"; Description = "Create, read items; modify, delete own items" }
        @{ Name = "NonEditingAuthor"; Description = "Create, read items; delete own items" }
        @{ Name = "Reviewer"; Description = "Read items only (full details)" }
        @{ Name = "Contributor"; Description = "Create items only (cannot read)" }
        @{ Name = "AvailabilityOnly"; Description = "View free/busy time only" }
        @{ Name = "LimitedDetails"; Description = "View free/busy time with subject and location" }
        @{ Name = "None"; Description = "No access" }
    )
}

function Grant-CalendarPermission {
    <#
    .SYNOPSIS
        Grants calendar permission to a user
    .PARAMETER CalendarOwner
        The email/UPN of the calendar owner
    .PARAMETER Trustee
        The email/UPN of the user being granted access
    .PARAMETER AccessRights
        The permission level to grant
    .EXAMPLE
        Grant-CalendarPermission -CalendarOwner "john@contoso.com" -Trustee "jane@contoso.com" -AccessRights "Reviewer"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CalendarOwner,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Trustee,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author",
                     "NonEditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly",
                     "LimitedDetails", "None")]
        [string]$AccessRights
    )

    try {
        $calendarIdentity = "${CalendarOwner}:\Calendar"

        # Check if permission already exists
        $existingPermission = Get-MailboxFolderPermission -Identity $calendarIdentity -User $Trustee -ErrorAction SilentlyContinue

        if ($existingPermission) {
            # Update existing permission
            Set-MailboxFolderPermission -Identity $calendarIdentity -User $Trustee -AccessRights $AccessRights -ErrorAction Stop
            return @{
                Success = $true
                Action = "Updated"
                Message = "Updated permission for $Trustee on $CalendarOwner's calendar to $AccessRights"
            }
        }
        else {
            # Add new permission
            Add-MailboxFolderPermission -Identity $calendarIdentity -User $Trustee -AccessRights $AccessRights -ErrorAction Stop
            return @{
                Success = $true
                Action = "Added"
                Message = "Granted $AccessRights access to $Trustee on $CalendarOwner's calendar"
            }
        }
    }
    catch {
        return @{
            Success = $false
            Action = "Failed"
            Message = "Failed to grant permission: $($_.Exception.Message)"
        }
    }
}

function Remove-CalendarPermission {
    <#
    .SYNOPSIS
        Removes calendar permission from a user
    .PARAMETER CalendarOwner
        The email/UPN of the calendar owner
    .PARAMETER Trustee
        The email/UPN of the user whose access is being removed
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CalendarOwner,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Trustee
    )

    try {
        $calendarIdentity = "${CalendarOwner}:\Calendar"
        Remove-MailboxFolderPermission -Identity $calendarIdentity -User $Trustee -Confirm:$false -ErrorAction Stop

        return @{
            Success = $true
            Message = "Removed $Trustee's access to $CalendarOwner's calendar"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to remove permission: $($_.Exception.Message)"
        }
    }
}

function Get-CalendarPermissions {
    <#
    .SYNOPSIS
        Gets current calendar permissions for a mailbox
    .PARAMETER CalendarOwner
        The email/UPN of the calendar owner
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CalendarOwner
    )

    try {
        $calendarIdentity = "${CalendarOwner}:\Calendar"
        $permissions = Get-MailboxFolderPermission -Identity $calendarIdentity -ErrorAction Stop

        return @{
            Success = $true
            Permissions = $permissions
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to get permissions: $($_.Exception.Message)"
            Permissions = @()
        }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Connect-ExchangeOnlineSession',
    'Disconnect-ExchangeOnlineSession',
    'Test-ExchangeConnection',
    'Get-CalendarPermissionLevels',
    'Grant-CalendarPermission',
    'Remove-CalendarPermission',
    'Get-CalendarPermissions'
)

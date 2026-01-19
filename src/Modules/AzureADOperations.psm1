#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Azure AD / Microsoft Graph Operations Module for CalendarWarlock
.DESCRIPTION
    Provides functions for querying Azure AD users by job title using Microsoft Graph.
    Requires the Microsoft.Graph.Users module to be installed.
.NOTES
    Author: CalendarWarlock
    Requires: PowerShell 5.1+, Microsoft.Graph.Users module
#>

# Module-level variable to track connection status
$script:IsGraphConnected = $false

function Connect-GraphSession {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using modern authentication
    .DESCRIPTION
        Establishes a connection to Microsoft Graph with required scopes for
        reading user information.
    .EXAMPLE
        Connect-GraphSession
    #>
    [CmdletBinding()]
    param()

    try {
        # Required scopes for reading user data
        $scopes = @(
            "User.Read.All",
            "Directory.Read.All"
        )

        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Write-Verbose "Already connected to Microsoft Graph"
            $script:IsGraphConnected = $true
            return @{
                Success = $true
                Message = "Already connected to Microsoft Graph as $($context.Account)"
            }
        }

        # Connect using interactive authentication
        Connect-MgGraph -Scopes $scopes -NoWelcome
        $script:IsGraphConnected = $true

        $context = Get-MgContext
        return @{
            Success = $true
            Message = "Successfully connected to Microsoft Graph as $($context.Account)"
        }
    }
    catch {
        $script:IsGraphConnected = $false
        return @{
            Success = $false
            Message = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        }
    }
}

function Disconnect-GraphSession {
    <#
    .SYNOPSIS
        Disconnects from Microsoft Graph
    #>
    [CmdletBinding()]
    param()

    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:IsGraphConnected = $false
        return @{
            Success = $true
            Message = "Disconnected from Microsoft Graph"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error disconnecting: $($_.Exception.Message)"
        }
    }
}

function Test-GraphConnection {
    <#
    .SYNOPSIS
        Tests if connected to Microsoft Graph
    #>
    [CmdletBinding()]
    param()

    try {
        $context = Get-MgContext -ErrorAction Stop
        if ($context) {
            $script:IsGraphConnected = $true
            return $true
        }
        return $false
    }
    catch {
        $script:IsGraphConnected = $false
        return $false
    }
}

function Get-UsersByJobTitle {
    <#
    .SYNOPSIS
        Gets all users with a specific job title
    .PARAMETER JobTitle
        The job title to search for (case-insensitive, exact match)
    .PARAMETER IncludeDisabled
        Include disabled/blocked user accounts
    .EXAMPLE
        Get-UsersByJobTitle -JobTitle "Software Engineer"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$JobTitle,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDisabled
    )

    try {
        # Build filter - escape single quotes in job title
        $escapedTitle = $JobTitle.Replace("'", "''")
        $filter = "jobTitle eq '$escapedTitle'"

        if (-not $IncludeDisabled) {
            $filter += " and accountEnabled eq true"
        }

        # Get users with the specified job title
        $users = Get-MgUser -Filter $filter -Property @(
            "Id",
            "DisplayName",
            "UserPrincipalName",
            "Mail",
            "JobTitle",
            "Department",
            "AccountEnabled"
        ) -All -ErrorAction Stop

        $userList = @()
        foreach ($user in $users) {
            $userList += @{
                Id = $user.Id
                DisplayName = $user.DisplayName
                Email = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                UserPrincipalName = $user.UserPrincipalName
                JobTitle = $user.JobTitle
                Department = $user.Department
                Enabled = $user.AccountEnabled
            }
        }

        return @{
            Success = $true
            Users = $userList
            Count = $userList.Count
            Message = "Found $($userList.Count) user(s) with job title '$JobTitle'"
        }
    }
    catch {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Message = "Failed to get users: $($_.Exception.Message)"
        }
    }
}

function Get-AllJobTitles {
    <#
    .SYNOPSIS
        Gets all unique job titles in the organization
    .DESCRIPTION
        Retrieves all users and extracts unique job titles for use in dropdowns/autocomplete
    .EXAMPLE
        Get-AllJobTitles
    #>
    [CmdletBinding()]
    param()

    try {
        # Get all users with job titles
        $users = Get-MgUser -Property "JobTitle" -All -ErrorAction Stop |
                 Where-Object { $_.JobTitle -and $_.JobTitle.Trim() -ne "" }

        $titles = $users | Select-Object -ExpandProperty JobTitle | Sort-Object -Unique

        return @{
            Success = $true
            JobTitles = $titles
            Count = $titles.Count
            Message = "Found $($titles.Count) unique job title(s)"
        }
    }
    catch {
        return @{
            Success = $false
            JobTitles = @()
            Count = 0
            Message = "Failed to get job titles: $($_.Exception.Message)"
        }
    }
}

function Get-AllDepartments {
    <#
    .SYNOPSIS
        Gets all unique departments in the organization
    .DESCRIPTION
        Retrieves all users and extracts unique departments for use in dropdowns/autocomplete
    .EXAMPLE
        Get-AllDepartments
    #>
    [CmdletBinding()]
    param()

    try {
        # Get all users with departments
        $users = Get-MgUser -Property "Department" -All -ErrorAction Stop |
                 Where-Object { $_.Department -and $_.Department.Trim() -ne "" }

        $departments = $users | Select-Object -ExpandProperty Department | Sort-Object -Unique

        return @{
            Success = $true
            Departments = $departments
            Count = $departments.Count
            Message = "Found $($departments.Count) unique department(s)"
        }
    }
    catch {
        return @{
            Success = $false
            Departments = @()
            Count = 0
            Message = "Failed to get departments: $($_.Exception.Message)"
        }
    }
}

function Get-AllOffices {
    <#
    .SYNOPSIS
        Gets all unique office locations in the organization
    .DESCRIPTION
        Retrieves all users and extracts unique office locations for use in dropdowns/autocomplete
    .EXAMPLE
        Get-AllOffices
    #>
    [CmdletBinding()]
    param()

    try {
        # Get all users with office locations
        $users = Get-MgUser -Property "OfficeLocation" -All -ErrorAction Stop |
                 Where-Object { $_.OfficeLocation -and $_.OfficeLocation.Trim() -ne "" }

        $offices = $users | Select-Object -ExpandProperty OfficeLocation | Sort-Object -Unique

        return @{
            Success = $true
            Offices = $offices
            Count = $offices.Count
            Message = "Found $($offices.Count) unique office location(s)"
        }
    }
    catch {
        return @{
            Success = $false
            Offices = @()
            Count = 0
            Message = "Failed to get office locations: $($_.Exception.Message)"
        }
    }
}

function Get-UsersByDepartment {
    <#
    .SYNOPSIS
        Gets all users in a specific department
    .PARAMETER Department
        The department to search for (case-insensitive, exact match)
    .PARAMETER IncludeDisabled
        Include disabled/blocked user accounts
    .EXAMPLE
        Get-UsersByDepartment -Department "Engineering"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Department,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDisabled
    )

    try {
        # Build filter - escape single quotes in department
        $escapedDepartment = $Department.Replace("'", "''")
        $filter = "department eq '$escapedDepartment'"

        if (-not $IncludeDisabled) {
            $filter += " and accountEnabled eq true"
        }

        # Get users with the specified department
        $users = Get-MgUser -Filter $filter -Property @(
            "Id",
            "DisplayName",
            "UserPrincipalName",
            "Mail",
            "JobTitle",
            "Department",
            "OfficeLocation",
            "AccountEnabled"
        ) -All -ErrorAction Stop

        $userList = @()
        foreach ($user in $users) {
            $userList += @{
                Id = $user.Id
                DisplayName = $user.DisplayName
                Email = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                UserPrincipalName = $user.UserPrincipalName
                JobTitle = $user.JobTitle
                Department = $user.Department
                OfficeLocation = $user.OfficeLocation
                Enabled = $user.AccountEnabled
            }
        }

        return @{
            Success = $true
            Users = $userList
            Count = $userList.Count
            Message = "Found $($userList.Count) user(s) in department '$Department'"
        }
    }
    catch {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Message = "Failed to get users: $($_.Exception.Message)"
        }
    }
}

function Get-UsersByOffice {
    <#
    .SYNOPSIS
        Gets all users in a specific office location
    .PARAMETER Office
        The office location to search for (case-insensitive, exact match)
    .PARAMETER IncludeDisabled
        Include disabled/blocked user accounts
    .EXAMPLE
        Get-UsersByOffice -Office "New York"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Office,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeDisabled
    )

    try {
        # Build filter - escape single quotes in office
        $escapedOffice = $Office.Replace("'", "''")
        $filter = "officeLocation eq '$escapedOffice'"

        if (-not $IncludeDisabled) {
            $filter += " and accountEnabled eq true"
        }

        # Get users with the specified office location
        $users = Get-MgUser -Filter $filter -Property @(
            "Id",
            "DisplayName",
            "UserPrincipalName",
            "Mail",
            "JobTitle",
            "Department",
            "OfficeLocation",
            "AccountEnabled"
        ) -All -ErrorAction Stop

        $userList = @()
        foreach ($user in $users) {
            $userList += @{
                Id = $user.Id
                DisplayName = $user.DisplayName
                Email = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                UserPrincipalName = $user.UserPrincipalName
                JobTitle = $user.JobTitle
                Department = $user.Department
                OfficeLocation = $user.OfficeLocation
                Enabled = $user.AccountEnabled
            }
        }

        return @{
            Success = $true
            Users = $userList
            Count = $userList.Count
            Message = "Found $($userList.Count) user(s) in office '$Office'"
        }
    }
    catch {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Message = "Failed to get users: $($_.Exception.Message)"
        }
    }
}

function Get-UserByEmail {
    <#
    .SYNOPSIS
        Gets a single user by email or UPN
    .PARAMETER Email
        The email address or UserPrincipalName
    .EXAMPLE
        Get-UserByEmail -Email "john@contoso.com"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Email
    )

    try {
        # Try to get user by UPN first, then by mail
        $user = Get-MgUser -UserId $Email -Property @(
            "Id",
            "DisplayName",
            "UserPrincipalName",
            "Mail",
            "JobTitle",
            "Department",
            "AccountEnabled"
        ) -ErrorAction SilentlyContinue

        if (-not $user) {
            # Escape single quotes to prevent OData injection
            $escapedEmail = $Email.Replace("'", "''")
            # Try searching by mail
            $user = Get-MgUser -Filter "mail eq '$escapedEmail'" -Property @(
                "Id",
                "DisplayName",
                "UserPrincipalName",
                "Mail",
                "JobTitle",
                "Department",
                "AccountEnabled"
            ) -ErrorAction Stop | Select-Object -First 1
        }

        if ($user) {
            return @{
                Success = $true
                User = @{
                    Id = $user.Id
                    DisplayName = $user.DisplayName
                    Email = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                    UserPrincipalName = $user.UserPrincipalName
                    JobTitle = $user.JobTitle
                    Department = $user.Department
                    Enabled = $user.AccountEnabled
                }
                Message = "Found user: $($user.DisplayName)"
            }
        }
        else {
            return @{
                Success = $false
                User = $null
                Message = "User not found: $Email"
            }
        }
    }
    catch {
        return @{
            Success = $false
            User = $null
            Message = "Failed to get user: $($_.Exception.Message)"
        }
    }
}

function Search-Users {
    <#
    .SYNOPSIS
        Searches for users by name or email (partial match)
    .PARAMETER SearchTerm
        The search term to match against display name or email
    .PARAMETER MaxResults
        Maximum number of results to return (default 50)
    .EXAMPLE
        Search-Users -SearchTerm "john"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchTerm,

        [Parameter(Mandatory = $false)]
        [int]$MaxResults = 50
    )

    try {
        # Escape single quotes to prevent OData injection
        $escapedSearchTerm = $SearchTerm.Replace("'", "''")

        # Use startsWith for searching (Graph API doesn't support contains easily)
        $filter = "startsWith(displayName, '$escapedSearchTerm') or startsWith(mail, '$escapedSearchTerm') or startsWith(userPrincipalName, '$escapedSearchTerm')"

        $users = Get-MgUser -Filter $filter -Property @(
            "Id",
            "DisplayName",
            "UserPrincipalName",
            "Mail",
            "JobTitle",
            "Department",
            "AccountEnabled"
        ) -Top $MaxResults -ErrorAction Stop

        $userList = @()
        foreach ($user in $users) {
            $userList += @{
                Id = $user.Id
                DisplayName = $user.DisplayName
                Email = if ($user.Mail) { $user.Mail } else { $user.UserPrincipalName }
                UserPrincipalName = $user.UserPrincipalName
                JobTitle = $user.JobTitle
                Department = $user.Department
                Enabled = $user.AccountEnabled
            }
        }

        return @{
            Success = $true
            Users = $userList
            Count = $userList.Count
            Message = "Found $($userList.Count) user(s) matching '$SearchTerm'"
        }
    }
    catch {
        return @{
            Success = $false
            Users = @()
            Count = 0
            Message = "Failed to search users: $($_.Exception.Message)"
        }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Connect-GraphSession',
    'Disconnect-GraphSession',
    'Test-GraphConnection',
    'Get-UsersByJobTitle',
    'Get-AllJobTitles',
    'Get-AllDepartments',
    'Get-AllOffices',
    'Get-UsersByDepartment',
    'Get-UsersByOffice',
    'Get-UserByEmail',
    'Search-Users'
)

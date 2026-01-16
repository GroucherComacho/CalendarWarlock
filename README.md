# CalendarWarlock

A secure, easy-to-use GUI tool for managing Exchange Online calendar permissions in bulk by job title.

## Features

- **Bulk Permission by Job Title**: Grant calendar permissions to multiple users based on their job title
- **Two Bulk Operations**:
  1. Grant a single user access to all calendars of users with a specific job title
  2. Grant all users with a specific job title access to a single user's calendar
- **Simple GUI Interface**: No command-line knowledge required
- **Modern Authentication**: Uses Microsoft's modern authentication (supports MFA)
- **Secure**: No credentials stored locally; uses interactive authentication
- **Comprehensive Logging**: All operations are logged for audit purposes

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Required PowerShell Modules:
  - `ExchangeOnlineManagement` (v3.0.0+)
  - `Microsoft.Graph.Users` (v2.0.0+)

## Installation

### 1. Install Required PowerShell Modules

Open PowerShell as Administrator and run:

```powershell
# Install Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force

# Install Microsoft Graph module
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
```

### 2. Download CalendarWarlock

Clone or download this repository to your local machine.

### 3. Unblock Scripts (if needed)

If you downloaded the files, you may need to unblock them:

```powershell
Get-ChildItem -Path "C:\Path\To\CalendarWarlock" -Recurse | Unblock-File
```

## Usage

### Starting the Application

```powershell
.\Start-CalendarWarlock.ps1
```

Or run directly:

```powershell
.\src\CalendarWarlock.ps1
```

### Connecting to Exchange Online

1. Enter your organization domain (e.g., `contoso.onmicrosoft.com`)
2. Click **Connect**
3. Sign in with your Microsoft 365 admin credentials
4. Complete MFA if prompted

### Granting Bulk Permissions

#### Option 1: Grant a User Access to All Calendars of a Job Title

Use this when you want to give one person (e.g., an executive assistant) access to view/edit the calendars of everyone with a specific job title (e.g., all "Directors").

1. Select the **Job Title** from the dropdown
2. Enter the **Target User's email** (the person who will receive access)
3. Select the desired **Permission Level**
4. Click **"Grant User Access to All Calendars of Job Title"**

#### Option 2: Grant All of a Job Title Access to a User's Calendar

Use this when you want everyone with a specific job title to have access to one person's calendar.

1. Select the **Job Title** from the dropdown
2. Enter the **Calendar Owner's email** (the person whose calendar will be shared)
3. Select the desired **Permission Level**
4. Click **"Grant All of Job Title Access to User's Calendar"**

### Permission Levels

| Level | Description |
|-------|-------------|
| Owner | Full control - read, create, modify, delete all items and manage permissions |
| PublishingEditor | Create, read, modify, delete all items; create subfolders |
| Editor | Create, read, modify, delete all items |
| PublishingAuthor | Create, read items; modify, delete own items; create subfolders |
| Author | Create, read items; modify, delete own items |
| NonEditingAuthor | Create, read items; delete own items |
| Reviewer | Read items only (full details) |
| Contributor | Create items only (cannot read) |
| AvailabilityOnly | View free/busy time only |
| LimitedDetails | View free/busy time with subject and location |
| None | No access |

## Required Permissions

The user running CalendarWarlock needs the following permissions:

### Exchange Online
- Exchange Administrator role, or
- Recipient Management role (for calendar permissions)

### Microsoft Graph (Azure AD)
- `User.Read.All` - To read user profiles and job titles
- `Directory.Read.All` - To query directory information

## Security Considerations

- **No Stored Credentials**: CalendarWarlock uses interactive modern authentication and never stores credentials
- **MFA Support**: Fully compatible with Multi-Factor Authentication
- **Audit Logging**: All operations are logged to the `Logs` folder with timestamps
- **Confirmation Prompts**: Bulk operations require explicit user confirmation
- **Session Management**: Properly disconnects sessions when closing the application

## Project Structure

```
CalendarWarlock/
├── src/
│   ├── CalendarWarlock.ps1          # Main GUI application
│   └── Modules/
│       ├── ExchangeOperations.psm1  # Exchange Online operations
│       └── AzureADOperations.psm1   # Azure AD/Graph operations
├── docs/
│   └── USAGE.md                     # Detailed usage guide
├── Start-CalendarWarlock.ps1        # Launcher script
└── README.md                        # This file
```

## Troubleshooting

### "Missing required PowerShell modules"
Install the required modules using the commands in the Installation section.

### "Failed to connect to Exchange Online"
- Verify you have the correct permissions
- Check your organization domain is correct
- Ensure your account has Exchange Admin or Recipient Management role

### "Failed to connect to Microsoft Graph"
- Consent to the required permissions when prompted
- Verify you have User.Read.All and Directory.Read.All permissions

### "No job titles found"
- Ensure users in your organization have job titles populated in Azure AD
- Check your Microsoft Graph connection is active

## License

MIT License - See LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.

# CalendarWarlock

**Manage Exchange Online calendar permissions in bulk - no PowerShell expertise required.**

CalendarWarlock is a Windows GUI application that simplifies granting and removing calendar permissions across your organization. Instead of running complex PowerShell commands, use a simple point-and-click interface to manage permissions by job title or department.

## What Can CalendarWarlock Do?

- **Grant bulk permissions by job title** - Give an executive assistant access to all Director calendars with one click
- **Grant bulk permissions by department** - Share a calendar with everyone in the Sales department
- **Remove permissions in bulk** - Revoke access just as easily as granting it
- **Process CSV files** - Handle large batches of permission changes via spreadsheet

## Quick Start

### Step 1: Install Required Components

Open PowerShell as Administrator and run:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
```

### Step 2: Launch CalendarWarlock

Double-click `CalendarWarlock.bat` or run:

```powershell
.\Start-CalendarWarlock.ps1
```

### Step 3: Connect to Your Organization

1. Enter your Microsoft 365 domain (e.g., `contoso.onmicrosoft.com`)
2. Click **Connect**
3. Sign in with your admin credentials
4. Complete MFA if prompted

### Step 4: Start Managing Permissions

Choose your operation type and follow the on-screen prompts.

## Common Use Cases

### Give an Assistant Access to Executive Calendars

*Scenario: Your CEO's assistant needs to view and edit all Director calendars*

1. Select **Job Title** mode
2. Choose "Director" from the dropdown
3. Enter the assistant's email address
4. Select **Editor** permission level
5. Click **Grant User Access to All Calendars of Job Title**

### Share a Meeting Room Calendar with a Department

*Scenario: Everyone in Engineering needs to see the main conference room calendar*

1. Select **Department** mode
2. Choose "Engineering" from the dropdown
3. Enter the conference room's email address
4. Select **Reviewer** permission level
5. Click **Grant All of Department Access to User's Calendar**

### Remove a Former Employee's Calendar Access

*Scenario: A manager left and you need to remove their access to team calendars*

1. Select **Job Title** or **Department** mode
2. Enter the former employee's email
3. Click the appropriate **Remove** button

## Permission Levels Explained

| Level | What They Can Do |
|-------|------------------|
| **Owner** | Full control - read, edit, delete everything, manage permissions |
| **Editor** | Read, create, edit, delete all items |
| **Reviewer** | Read-only access (view full details) |
| **AvailabilityOnly** | See only free/busy status |
| **LimitedDetails** | See free/busy plus meeting subject and location |
| **None** | Remove all access |

*For more permission levels, see the [full documentation](docs/USAGE.md).*

## Requirements

- **Windows 10/11** or Windows Server 2016+
- **PowerShell 5.1** or higher
- **Microsoft 365 Admin Permissions**:
  - Exchange Administrator role (or Recipient Management)
  - User.Read.All and Directory.Read.All Graph permissions

## Troubleshooting

**"Missing required PowerShell modules"**
Run the installation commands in Step 1 above.

**"Failed to connect"**
- Verify your domain is correct
- Ensure you have the required admin permissions
- Check your internet connection

**"No job titles/departments found"**
Your Azure AD user profiles may not have these fields populated. Contact your IT admin.

## Security

CalendarWarlock is designed with security in mind:

- Your credentials are never stored locally
- Full MFA support
- All operations are logged for audit purposes
- Comprehensive input validation

See [SECURITY.md](SECURITY.md) for details on security measures and testing.

## Documentation

- [Detailed Usage Guide](docs/USAGE.md) - Step-by-step instructions for all features
- [Security Information](SECURITY.md) - Security features and assessment results
- [Security Assessment](SECURITY_ASSESSMENT.md) - Technical security audit report

## License

MIT License - See [LICENSE](LICENSE) for details.

## Support

For issues or feature requests, visit the [GitHub Issues](https://github.com/GroucherComacho/CalendarWarlock/issues) page.

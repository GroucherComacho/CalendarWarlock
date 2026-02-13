# CalendarWarlock

**Bulk Exchange Online calendar permission management through a simple Windows GUI — no PowerShell expertise required.**

CalendarWarlock lets Microsoft 365 administrators grant and remove calendar permissions across their organization using a point-and-click interface. Manage permissions by job title, department, individual user, or CSV import — all without writing a single line of PowerShell.

---

## Features

- **Bulk by Job Title** — Grant or remove calendar access for all users sharing a job title
- **Bulk by Department** — Manage permissions across entire departments at once
- **Single-User Management** — Add or remove one user's access to multiple calendars
- **CSV Import** — Process large batches of permission changes from a spreadsheet
- **Permission Viewer** — See current calendar permissions for any mailbox in real time
- **Interactive Search** — Find users by name, email, department, or office location
- **Audit Logging** — Every operation is logged with timestamps for compliance
- **Theme Support** — Dark, Light, and Warlock themes
- **Session Timeout** — Automatic disconnect after 30 minutes of inactivity

---

## Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10/11 or Windows Server 2016+ |
| **PowerShell** | 5.1 or higher (included with Windows 10+) |
| **Modules** | ExchangeOnlineManagement 3.0+, Microsoft.Graph.Users 2.0+ |
| **Admin Roles** | Exchange Administrator *or* Recipient Management |
| **Graph Permissions** | `User.Read.All`, `Directory.Read.All` |

---

## Installation

### Option 1: MSI Installer (Recommended)

1. Download `CalendarWarlock.msi` from the [Releases](https://github.com/GroucherComacho/CalendarWarlock/releases) page
2. Run the installer and follow the prompts
3. Launch from the Start Menu or Desktop shortcut

### Option 2: Manual Setup

1. Clone or download this repository
2. Open PowerShell and install the required modules:

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
```

3. Launch by double-clicking `CalendarWarlock.bat` or running:

```powershell
.\Start-CalendarWarlock.ps1
```

> The launcher script automatically checks for prerequisites and will prompt you if anything is missing.

---

## Quick Start

### 1. Connect

Enter your Microsoft 365 domain (e.g., `contoso.onmicrosoft.com`), click **Connect**, and sign in through your browser. MFA is fully supported.

### 2. Choose an Operation Mode

| Mode | Use When... |
|---|---|
| **Job Title** | You want to manage permissions for all users with a specific title |
| **Department** | You want to manage permissions for an entire department |
| **CSV Import** | You have a spreadsheet of permission changes to process |

### 3. Configure and Execute

Select the target group, enter the user's email, choose a permission level, and click the action button. A confirmation dialog appears before any changes are made.

### 4. Review Results

The results panel shows a per-user breakdown of successes, failures, and skipped operations. Full details are written to the log file.

---

## Permission Levels

| Level | Can Do |
|---|---|
| **Owner** | Full control: read, create, edit, delete all items, create subfolders, manage permissions |
| **PublishingEditor** | Read, create, edit, delete all items, create subfolders |
| **Editor** | Read, create, edit, delete all items |
| **PublishingAuthor** | Read all, create items, edit/delete own items, create subfolders |
| **Author** | Read all, create items, edit/delete own items |
| **NonEditingAuthor** | Read all, create items, delete own items (no editing) |
| **Reviewer** | Read-only access with full details |
| **Contributor** | Create items only (no read access) |
| **AvailabilityOnly** | See free/busy status only |
| **LimitedDetails** | See free/busy plus subject and location |
| **None** | Remove all access |

> **Tip:** Use the least privilege needed. `Reviewer` is sufficient for read-only access; `AvailabilityOnly` is enough for scheduling.

---

## Common Use Cases

### Give an Executive Assistant Access to All Director Calendars

1. Select **Job Title** mode
2. Choose "Director" from the dropdown
3. Enter the assistant's email address
4. Select **Editor** permission level
5. Click **Grant User Access to All Calendars of Job Title**

### Share a Calendar with an Entire Department

1. Select **Department** mode
2. Choose "Engineering" from the dropdown
3. Enter the calendar owner's email
4. Select **Reviewer** permission level
5. Click **Grant All of Department Access to User's Calendar**

### Revoke a Former Employee's Access

1. Select **Job Title** or **Department** mode
2. Enter the former employee's email
3. Click the corresponding **Remove** button

### Bulk Changes via CSV

1. Prepare a CSV file using the template in `Bulk/bulktemplate.csv`:

```csv
MailboxEmail,UserEmail,AccessLevel
ceo@contoso.com,assistant@contoso.com,Editor
cfo@contoso.com,assistant@contoso.com,Reviewer
```

2. Click **Import CSV**, select your file, and confirm

> CSV files are validated before processing. Maximum file size is 10 MB. Files with more than 1,000 rows will prompt a warning about potential API throttling.

---

## Troubleshooting

| Problem | Solution |
|---|---|
| **"Missing required PowerShell modules"** | Run the `Install-Module` commands in the [Installation](#installation) section |
| **"Failed to connect"** | Verify your domain, admin permissions, and internet connection |
| **"No job titles/departments found"** | Ensure Azure AD user profiles have these fields populated |
| **Session disconnects unexpectedly** | The 30-minute idle timeout triggered. Reconnect and resume. |
| **CSV import errors** | Check that all emails are valid, access levels are spelled correctly, and the file is under 10 MB |

---

## Project Structure

```
CalendarWarlock/
├── CalendarWarlock.bat              # Entry point launcher
├── Start-CalendarWarlock.ps1        # Prerequisites checker & launcher
├── src/
│   ├── CalendarWarlock.ps1          # Main GUI application
│   └── Modules/
│       ├── AzureADOperations.psm1   # Microsoft Graph user queries
│       └── ExchangeOperations.psm1  # Exchange Online calendar operations
├── Bulk/
│   └── bulktemplate.csv             # CSV template for batch operations
├── installer/
│   ├── Product.wxs                  # WiX MSI definition
│   └── Build-Installer.bat          # Installer build script
└── docs/
    └── USAGE.md                     # Detailed usage guide
```

---

## Security

CalendarWarlock takes security seriously:

- **No credential storage** — OAuth 2.0 interactive authentication only; credentials never touch disk
- **MFA support** — Full compatibility with multi-factor authentication
- **Input validation** — Email format, permission level, domain, and CSV content validation
- **Injection prevention** — OData query escaping and module path traversal protection
- **Error sanitization** — Sensitive data (paths, IPs, connection strings) stripped from UI and logs
- **Session timeout** — Automatic disconnect after 30 minutes of inactivity
- **Audit logging** — All operations logged with restricted file permissions

The application has undergone multiple security assessments. Current risk level: **LOW**.

See [SECURITY.md](SECURITY.md) for the full security breakdown and [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md) for the detailed audit report.

---

## Documentation

| Document | Description |
|---|---|
| [Usage Guide](docs/USAGE.md) | Step-by-step instructions, permission level matrix, best practices |
| [Security Overview](SECURITY.md) | Security architecture, controls, and vulnerability history |
| [Security Assessment](SECURITY_ASSESSMENT.md) | Full penetration test report with CVSS scores |

---

## License

MIT License — see [LICENSE](LICENSE) for details.

Copyright (c) 2026 Groucher Labs

---

## Support

For issues or feature requests, visit the [GitHub Issues](https://github.com/GroucherComacho/CalendarWarlock/issues) page.

# CalendarWarlock — Security Overview

This document describes the security architecture, controls, and vulnerability history of CalendarWarlock. For the full penetration test report with CVSS scores and test scenarios, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

**Current Risk Level: LOW** (as of 2026-02-13)

---

## Table of Contents

- [Security Architecture](#security-architecture)
- [Authentication & Session Management](#authentication--session-management)
- [Input Validation & Injection Prevention](#input-validation--injection-prevention)
- [Data Protection](#data-protection)
- [Audit Logging](#audit-logging)
- [Deployment Security](#deployment-security)
- [Vulnerability History](#vulnerability-history)
- [Open / Acknowledged Items](#open--acknowledged-items)
- [Best Practices for Administrators](#best-practices-for-administrators)
- [Reporting Security Issues](#reporting-security-issues)

---

## Security Architecture

CalendarWarlock is a stateless GUI application. It stores no data locally — all user and calendar information is fetched on-demand from Microsoft 365 via OAuth 2.0-authenticated API calls.

```
┌──────────────────────┐
│   CalendarWarlock     │
│   (PowerShell GUI)    │
│                       │
│  No local data store  │
│  No credential cache  │
│  No config secrets    │
└────────┬─────────────┘
         │  OAuth 2.0 (HTTPS)
         │
    ┌────▼────────────────────────────┐
    │       Microsoft 365             │
    │  ┌─────────────┐ ┌───────────┐ │
    │  │ Graph API    │ │ Exchange  │ │
    │  │ (User data)  │ │ Online    │ │
    │  └─────────────┘ └───────────┘ │
    └─────────────────────────────────┘
```

**Key architectural properties:**

- **No secrets on disk** — OAuth tokens are managed entirely by the Microsoft PowerShell modules; CalendarWarlock never accesses or persists them
- **No local database** — All data comes from and goes to Microsoft 365
- **No network listeners** — The application makes outbound connections only
- **No background services** — Runs only when the user launches it

---

## Authentication & Session Management

### Authentication

| Property | Detail |
|---|---|
| **Method** | OAuth 2.0 interactive flow via browser |
| **MFA** | Fully supported (handled by Microsoft's auth flow) |
| **Credential storage** | None — credentials are entered in the browser, never in the application |
| **Token handling** | Managed by `ExchangeOnlineManagement` and `Microsoft.Graph` modules |
| **Required roles** | Exchange Administrator or Recipient Management |
| **Required Graph scopes** | `User.Read.All`, `Directory.Read.All` |

### Session Management

| Control | Detail |
|---|---|
| **Idle timeout** | 30 minutes — timer checks every 60 seconds |
| **Activity tracking** | Button clicks, text changes, and dropdown selections reset the idle timer |
| **Timeout behavior** | Session disconnected automatically with user notification |
| **Manual disconnect** | Available at any time via the Disconnect button |
| **Form close** | Prompts for disconnect if a session is active |
| **Connection cleanup** | Both Exchange Online and Graph sessions are terminated on disconnect |

---

## Input Validation & Injection Prevention

### Email Validation

All email inputs are validated against an RFC-compliant regex before any API call:

```
^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$
```

This validation is applied consistently across all operations: single-user grants/removals, bulk by job title, bulk by department, and CSV imports.

### Permission Level Validation

Only the 11 valid Exchange Online calendar permission levels are accepted:

`Owner` | `PublishingEditor` | `Editor` | `PublishingAuthor` | `Author` | `NonEditingAuthor` | `Reviewer` | `Contributor` | `AvailabilityOnly` | `LimitedDetails` | `None`

Enforced at two layers:
1. **Application layer** — `Test-ValidAccessLevel` function validates before any operation
2. **Module layer** — `ValidateSet` attribute on Exchange Operations functions rejects invalid values server-side

### OData Injection Prevention

All user-supplied values used in Microsoft Graph API filter queries are escaped by replacing single quotes (`'` to `''`). This escaping is applied at all six query points in `AzureADOperations.psm1`:

- `Get-UsersByJobTitle`
- `Get-UsersByDepartment`
- `Get-UsersByOffice`
- `Get-UserByEmail`
- `Search-Users`

### Organization Domain Validation

The domain input is validated against:

```
^[a-zA-Z0-9][a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$
```

Invalid domains are rejected before any connection attempt.

### CSV Import Validation

| Check | Detail |
|---|---|
| **File size** | Rejected if > 10 MB |
| **Row count** | Warning prompt if > 1,000 rows (API throttling risk) |
| **Email format** | Each row's MailboxEmail and UserEmail validated |
| **Access level** | Each row's AccessLevel checked against the valid set |
| **Formula injection** | Email validation rejects formula-prefixed values (`=`, `+`, `-`, `@`) as invalid |

### Module Path Security

PowerShell module imports use canonical path validation:

1. `[System.IO.Path]::GetFullPath()` resolves the canonical path
2. `.StartsWith()` confirms the path is within the application directory
3. Only `.psm1` file extensions are accepted

This prevents path traversal attacks from loading code outside the application.

---

## Data Protection

### Data in Transit

All communication with Microsoft 365 uses HTTPS via OAuth 2.0. The application does not manage TLS or certificates directly — this is handled by the Microsoft PowerShell modules.

### Data at Rest

CalendarWarlock stores no user data, calendar data, credentials, or tokens on disk. The only files written locally are audit logs (see below).

### Error Sanitization

The `Sanitize-ErrorMessage` function strips sensitive data from all error messages before they appear in the UI or log files:

| Data type | Pattern removed |
|---|---|
| **File paths** | Windows (`C:\...`) and Unix (`/...`) style paths |
| **Connection strings** | Strings matching connection string patterns |
| **IP addresses** | IPv4 addresses |
| **Server names** | Internal server hostname patterns |

---

## Audit Logging

### Log Location and Format

Logs are written to `src/Logs/CalendarWarlock_<timestamp>.log` with the format:

```
[2026-02-13 14:30:22] [INFO] CalendarWarlock started
[2026-02-13 14:31:15] [SUCCESS] Bulk grant completed: Success=12, Failed=0, Skipped=1
```

### Log Security Controls

| Control | Detail |
|---|---|
| **File permissions** | Logs directory has restrictive ACLs — only the current user has access |
| **Version control** | Excluded via `.gitignore` (`*.log`, `src/Logs/`) |
| **Content** | Operation types, timestamps, email addresses, and results |
| **Sanitization** | Error messages are sanitized before logging (no paths, IPs, or connection strings) |
| **Credentials** | Never logged under any circumstances |

### What Is Logged

- Application start/stop
- Connection and disconnection events
- All permission grant and removal operations (mailbox, user, permission level)
- Operation results (success/failure/skipped counts)
- CSV import events (sanitized — no file paths)
- Errors (sanitized)

---

## Deployment Security

### Execution Policy

The batch launcher uses `-ExecutionPolicy RemoteSigned`, which:
- Allows locally-created scripts to run
- Requires downloaded/remote scripts to be digitally signed
- Provides a balance between usability and security

### Installer (MSI)

- Built with WiX Toolset v3
- Per-machine installation to `C:\Program Files\CalendarWarlock\`
- Requires administrator privileges to install
- Feature selection UI allows users to choose which components to install

### Known Limitations

- **No code signing** — Scripts and the MSI installer are not digitally signed. Administrators should verify file integrity after download and restrict write access to the installation directory.

---

## Vulnerability History

CalendarWarlock has undergone four security assessments. All HIGH and MEDIUM findings have been remediated.

### Assessment Timeline

| Date | Type | Outcome |
|---|---|---|
| 2026-01-19 | Initial security audit | MEDIUM risk — 2 HIGH, 5 MEDIUM, 3 LOW found |
| 2026-01-19 | Remediation & re-assessment | LOW risk — all HIGH and MEDIUM fixed |
| 2026-02-10 | Vulnerability scan & penetration test | LOW-MEDIUM risk — 5 new MEDIUM, 4 new LOW found |
| 2026-02-13 | Full remediation | LOW risk — all new findings resolved |

### Remediated Findings

#### High Severity (2 found, 2 fixed)

| ID | Finding | Resolution |
|---|---|---|
| HIGH-001 | OData injection in `Search-Users` | Single-quote escaping applied to all Graph API filter queries |
| HIGH-002 | OData injection in `Get-UserByEmail` | Same fix — escaping applied at all six query construction points |

#### Medium Severity (10 found, 10 fixed)

| ID | Finding | Resolution |
|---|---|---|
| MEDIUM-001 | CSV formula injection | Email validation rejects formula-prefixed values |
| MEDIUM-002 | No email format validation | RFC-compliant regex validation added to all operations |
| MEDIUM-003 | Incomplete access level list | All 11 Exchange permission levels now validated |
| MEDIUM-004 | Module path traversal | Canonical path resolution + containment check + extension verification |
| MEDIUM-005 | Inconsistent email validation | Uniform validation applied across all bulk operations |
| MEDIUM-006 | `ExecutionPolicy Bypass` in launcher | Changed to `RemoteSigned` |
| MEDIUM-007 | No CSV file size limit | 10 MB hard limit + 1,000-row warning added |
| MEDIUM-008 | Unsanitized errors in log files | `Sanitize-ErrorMessage` applied to all log entries |
| MEDIUM-009 | No domain format validation | Domain regex validation added to Connect handler |
| MEDIUM-010 | No session timeout | 30-minute idle timeout implemented |

#### Low Severity (7 found, 5 fixed, 2 acknowledged)

| ID | Finding | Resolution |
|---|---|---|
| LOW-001 | Verbose error messages in UI | Error sanitization strips sensitive data |
| LOW-004 | Dead `Sanitize-CSVValue` code | Removed |
| LOW-005 | Default log file permissions | Restrictive ACLs set on Logs directory |
| LOW-006 | `DoEvents` re-entrancy risk | Replaced with targeted `Refresh()` calls |
| LOW-007 | ComboBox free-text input | Documented as intentional (OData escaping provides protection) |
| LOW-002 | Log files contain email addresses | Acknowledged — required for audit trail |
| LOW-003 | No rate limiting on bulk ops | Acknowledged — Microsoft 365 provides its own throttling |

---

## Open / Acknowledged Items

These items have been reviewed and accepted:

| ID | Severity | Item | Rationale |
|---|---|---|---|
| LOW-002 | Low | Log files contain email addresses | Required for audit compliance. Mitigated by restrictive ACLs and `.gitignore` exclusion. |
| LOW-003 | Low | No application-level rate limiting | Microsoft 365 enforces its own API throttling. A >1,000-row warning alerts users. |
| INFO-001 | Info | MSI requires admin to install | Appropriate for an administrative tool installed to Program Files. |
| INFO-002 | Info | No code signing | Recommended for future production deployments. |

---

## Best Practices for Administrators

1. **Keep modules updated** — Run `Update-Module ExchangeOnlineManagement` and `Update-Module Microsoft.Graph` regularly for security patches.

2. **Use least privilege** — Grant the minimum permission level needed. `Reviewer` for read-only, `AvailabilityOnly` for scheduling visibility.

3. **Protect the installation directory** — Since scripts are unsigned, ensure only administrators can write to the CalendarWarlock directory.

4. **Manage log files** — Logs contain email addresses and operation history. Handle them according to your organization's data retention policies.

5. **Verify CSV sources** — Only import CSV files from trusted sources. While the application validates content, reviewing files before import is good practice.

6. **Close when finished** — The 30-minute idle timeout provides protection, but explicitly disconnecting and closing the application is better.

7. **Lock your workstation** — Physical access to an active session bypasses all application-level controls.

8. **Audit regularly** — Review log files periodically to verify that permission changes align with authorized requests.

---

## Reporting Security Issues

If you discover a security vulnerability in CalendarWarlock, please report it through [GitHub Issues](https://github.com/GroucherComacho/CalendarWarlock/issues).

For the full technical penetration test report, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

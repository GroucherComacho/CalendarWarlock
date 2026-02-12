# CalendarWarlock v1.1.0 Release Notes

**Release Date:** 2026-02-12

---

## What's New

### Warlock Theme
A brand-new **Warlock theme** joins the existing Dark and Light themes. Cycle through all three with the theme toggle button. The Warlock theme features a deep void background with neon green terminal-style accents for that classic sysadmin aesthetic.

### 30-Minute Idle Session Timeout
Connected sessions now automatically disconnect after 30 minutes of inactivity. A notification informs you when a timeout occurs. Activity is tracked on button clicks and text input, and the timer resets with each interaction.

### Organization Domain Validation
The connection dialog now validates your organization domain format before attempting to connect, providing immediate feedback if the domain is malformed instead of waiting for a connection failure.

### CSV File Size Protection
Bulk CSV imports now enforce a **10 MB file size limit** to prevent memory exhaustion from oversized files. Additionally, CSV files with more than 500 rows display a warning in the confirmation dialog about potential Microsoft 365 API throttling.

---

## Security Improvements

This release includes the remediation of all findings from a comprehensive vulnerability scan and penetration test. Overall security risk level: **LOW**.

- **ExecutionPolicy hardened** - Launcher changed from `Bypass` to `RemoteSigned`, blocking untrusted remote scripts
- **Log file sanitization** - Error messages written to log files are now sanitized to remove file paths, IP addresses, and connection strings
- **Restrictive log permissions** - The Logs directory is created with ACLs restricted to the current user only
- **DoEvents removal** - All `DoEvents()` calls replaced with targeted `Refresh()` to eliminate UI re-entrancy risks

See [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md) for the full technical report.

---

## Bug Fixes

- Fixed broken theme references and redundant code across the application
- Reverted unintended UI text changes back to professional context
- Enlarged application icon (1.5x) for better visibility
- Terminal/console window now hidden on launch for cleaner UX

---

## Installer Updates

- Added proper license agreement display during installation

---

## Full Changelog

| Category | Change |
|----------|--------|
| Feature | Warlock theme with 3-way cycling (Dark / Light / Warlock) |
| Feature | 30-minute idle session auto-disconnect |
| Feature | Organization domain format validation |
| Feature | CSV file size limit (10 MB) and large batch warnings (>500 rows) |
| Security | ExecutionPolicy changed from Bypass to RemoteSigned |
| Security | Log file error messages sanitized via `Sanitize-ErrorMessage` |
| Security | Log directory ACLs restricted to current user |
| Security | `DoEvents()` replaced with `MainForm.Refresh()` |
| Fix | Broken theme color references corrected |
| Fix | Redundant/dead code cleaned up |
| Fix | Application icon enlarged for better visibility |
| Fix | Console window hidden on GUI launch |
| Installer | License agreement added to MSI installer |
| Docs | Security assessment updated with all findings remediated |
| Docs | Usage guide and security documentation added |

---

## Requirements

- **Windows 10/11** or Windows Server 2016+
- **PowerShell 5.1** or higher
- **ExchangeOnlineManagement** module v3.0.0+
- **Microsoft.Graph.Users** module v2.0.0+
- Microsoft 365 admin permissions (Exchange Administrator or Recipient Management)

---

## Upgrade Notes

- No breaking changes from v1.0.0
- The idle timeout is set to 30 minutes by default; reconnect after timeout as needed
- The CSV file size limit of 10 MB applies to both grant and remove bulk operations
- If your scripts were relying on `ExecutionPolicy Bypass` via the batch launcher, note that it now uses `RemoteSigned` - locally authored scripts will still run without issue

---

## Security Assessment Summary

| Severity | Total Found | Remediated |
|----------|-------------|------------|
| Critical | 0 | 0 |
| High | 2 | 2 |
| Medium | 10 | 10 |
| Low | 7 | 5 |

All HIGH and MEDIUM findings resolved. Remaining LOW items are acknowledged operational considerations.

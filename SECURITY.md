# CalendarWarlock Security

This document provides an overview of the security measures implemented in CalendarWarlock and the security issues that were identified and resolved during development.

## Security Overview

CalendarWarlock was designed with security as a priority. The application has undergone multiple rounds of comprehensive security testing. All identified vulnerabilities have been remediated.

**Current Security Status: LOW RISK**

## Key Security Features

### Authentication & Credentials

- **No Stored Credentials**: CalendarWarlock uses Microsoft's modern interactive authentication. Your credentials are never stored locally.
- **Multi-Factor Authentication (MFA)**: Full compatibility with MFA - authentication is handled through Microsoft's secure OAuth 2.0 flow.
- **Session Management**: Sessions are properly cleaned up when you close the application.
- **Idle Timeout**: Sessions automatically disconnect after 30 minutes of inactivity.

### Input Protection

- **Email Validation**: All email addresses are validated before processing to prevent malformed input.
- **Permission Level Validation**: Only valid Exchange Online permission levels are accepted.
- **OData Injection Prevention**: All inputs to Microsoft Graph queries are properly escaped.
- **CSV Validation**: CSV file imports validate email formats, access levels, and enforce a 10 MB file size limit.
- **Domain Validation**: Organization domains are validated before connection attempts.
- **CSV Formula Protection**: `Sanitize-CSVValue` function available for CSV output sanitization.

### Audit & Logging

- **Operation Logging**: All permission changes are logged with timestamps for audit purposes.
- **Sanitized Error Messages**: Error messages are cleaned to prevent sensitive information disclosure in both UI and log files.
- **Restrictive Log Permissions**: Log directory access is restricted to the current user only.

## Security Issues Found and Fixed

During security assessments, a total of 19 vulnerabilities were identified across two rounds of testing. All actionable findings have been remediated.

### High Severity Issues (2 Fixed)

#### OData Injection Prevention
**Issue**: User search queries could potentially be manipulated to alter database queries.

**Resolution**: All user inputs are now properly escaped before being used in Microsoft Graph API queries. Single quotes and special characters are sanitized to prevent injection attacks.

---

### Medium Severity Issues (10 Fixed)

| Finding | Fix Applied |
|---------|-------------|
| CSV Formula Injection | `Sanitize-CSVValue` function prefixes formula trigger characters |
| Email Format Validation | RFC-compliant regex validation across all operations |
| Incomplete AccessLevel List | All 11 Exchange Online permission levels validated |
| Module Path Traversal | Canonical path validation with `GetFullPath()` |
| Inconsistent Email Validation | `Test-ValidEmailFormat` applied to all bulk operations |
| ExecutionPolicy Bypass | Changed from `Bypass` to `RemoteSigned` |
| No CSV File Size Limit | 10 MB file size limit + row count warnings |
| Unsanitized Log Errors | `Sanitize-ErrorMessage` applied to log entries |
| No Domain Validation | Domain format regex validation before connection |
| No Session Timeout | 30-minute idle timeout with auto-disconnect |

---

### Low Severity Issues (5 Fixed, 2 Acknowledged)

| Finding | Status |
|---------|--------|
| Verbose Error Messages in UI | Fixed - `Sanitize-ErrorMessage` applied |
| CSV Sanitization Documentation | Fixed - Function purpose documented |
| Default Log Permissions | Fixed - Restrictive ACLs on Logs directory |
| DoEvents Re-entrancy | Fixed - Replaced with `Refresh()` calls |
| Sensitive Log Data | Acknowledged - Mitigated by ACLs and sanitization |
| No Rate Limiting | Acknowledged - M365 handles throttling |

---

## Best Practices for Users

1. **Keep Modules Updated**: Regularly update the ExchangeOnlineManagement and Microsoft.Graph PowerShell modules for security patches.

2. **Protect Log Files**: If your environment has strict data handling requirements, manage log files according to your organization's policies.

3. **Use Least Privilege**: Select the minimum permission level needed for each use case. "Reviewer" is sufficient for read-only access.

4. **Verify CSV Files**: Only import CSV files from trusted sources. The application validates email formats and access levels, but it's best to verify file contents before import.

5. **Close Application When Done**: Always close CalendarWarlock when finished to ensure sessions are properly disconnected.

6. **Lock Your Workstation**: While the application has a 30-minute idle timeout, always lock your workstation when stepping away.

7. **Verify Installation Integrity**: Ensure the installation directory has appropriate access controls to prevent script tampering.

## Reporting Security Issues

If you discover a security vulnerability in CalendarWarlock, please report it through:
- GitHub Issues: https://github.com/GroucherComacho/CalendarWarlock/issues

## Security Assessment History

| Date | Assessment Type | Result |
|------|-----------------|--------|
| 2026-01-19 | Initial Security Audit | MEDIUM RISK |
| 2026-01-19 | Remediation & Re-assessment | LOW RISK |
| 2026-02-10 | Vulnerability Scan & Penetration Test | LOW RISK (all findings remediated) |

For detailed technical security information, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

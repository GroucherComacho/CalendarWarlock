# CalendarWarlock Security

This document provides an overview of the security measures implemented in CalendarWarlock and the security issues that were identified and resolved during development.

## Security Overview

CalendarWarlock was designed with security as a priority. The application underwent comprehensive security testing and all identified vulnerabilities have been addressed.

**Current Security Status: LOW RISK**

## Key Security Features

### Authentication & Credentials

- **No Stored Credentials**: CalendarWarlock uses Microsoft's modern interactive authentication. Your credentials are never stored locally.
- **Multi-Factor Authentication (MFA)**: Full compatibility with MFA - authentication is handled through Microsoft's secure OAuth 2.0 flow.
- **Session Management**: Sessions are properly cleaned up when you close the application.

### Input Protection

- **Email Validation**: All email addresses are validated before processing to prevent malformed input.
- **Permission Level Validation**: Only valid Exchange Online permission levels are accepted.
- **CSV Sanitization**: CSV file imports are protected against formula injection attacks.

### Audit & Logging

- **Operation Logging**: All permission changes are logged with timestamps for audit purposes.
- **Sanitized Error Messages**: Error messages are cleaned to prevent sensitive information disclosure.

## Security Issues Found and Fixed

During security assessment, several vulnerabilities were identified and remediated:

### High Severity Issues (2 Fixed)

#### OData Injection Prevention
**Issue**: User search queries could potentially be manipulated to alter database queries.

**Resolution**: All user inputs are now properly escaped before being used in Microsoft Graph API queries. Single quotes and special characters are sanitized to prevent injection attacks.

**Files Fixed**:
- `src/Modules/AzureADOperations.psm1`

---

### Medium Severity Issues (5 Fixed)

#### CSV Formula Injection Protection
**Issue**: Malicious CSV files could contain formulas that execute when opened in Excel.

**Resolution**: Implemented `Sanitize-CSVValue` function that prefixes potentially dangerous values (starting with `=`, `+`, `-`, `@`, tab, or newline) with a single quote to prevent formula execution.

**Files Fixed**:
- `src/CalendarWarlock.ps1`

#### Email Format Validation
**Issue**: Invalid email formats could be passed to Exchange Online, potentially causing errors or unexpected behavior.

**Resolution**: Added `Test-ValidEmailFormat` function with RFC-compliant regex validation. Applied consistently across all bulk operations.

**Files Fixed**:
- `src/CalendarWarlock.ps1`

#### Complete Permission Level List
**Issue**: The access level validation was missing some valid permission types.

**Resolution**: Updated validation to include all 11 Exchange Online calendar permission levels: Owner, PublishingEditor, Editor, PublishingAuthor, Author, NonEditingAuthor, Reviewer, Contributor, AvailabilityOnly, LimitedDetails, and None.

**Files Fixed**:
- `src/CalendarWarlock.ps1`

#### Module Path Security
**Issue**: Relative module paths could potentially be exploited for path traversal attacks.

**Resolution**: Implemented canonical path validation to ensure modules are only loaded from within the application directory. File extensions are verified and paths outside the app folder are rejected.

**Files Fixed**:
- `src/CalendarWarlock.ps1`

#### Consistent Validation Across Operations
**Issue**: Some bulk operation functions lacked email validation that was present in others.

**Resolution**: Applied consistent email format validation to all bulk permission functions, ensuring uniform security across the application.

**Files Fixed**:
- `src/CalendarWarlock.ps1`

---

### Low Severity Issues (1 Fixed, 2 Acknowledged)

#### Error Message Sanitization (Fixed)
**Issue**: Error messages could reveal sensitive system information like file paths or IP addresses.

**Resolution**: Implemented `Sanitize-ErrorMessage` function that removes file paths, connection strings, and IP addresses from error output.

#### Log File Content (Acknowledged)
**Issue**: Log files contain email addresses and operation details.

**Mitigation**: Logs are excluded from version control via `.gitignore` and stored in a dedicated `Logs/` folder. No credentials are ever logged.

#### Rate Limiting (Acknowledged)
**Issue**: Bulk operations don't have built-in rate limiting.

**Note**: This is an operational consideration. Microsoft 365 has its own throttling mechanisms. Very large bulk operations may experience throttling.

---

## Best Practices for Users

1. **Keep Modules Updated**: Regularly update the ExchangeOnlineManagement and Microsoft.Graph PowerShell modules for security patches.

2. **Protect Log Files**: If your environment has strict data handling requirements, manage log files according to your organization's policies.

3. **Use Least Privilege**: Select the minimum permission level needed for each use case. "Reviewer" is sufficient for read-only access.

4. **Verify CSV Files**: Only import CSV files from trusted sources. The application sanitizes input, but it's best to verify file contents before import.

5. **Close Application When Done**: Always close CalendarWarlock when finished to ensure sessions are properly disconnected.

## Reporting Security Issues

If you discover a security vulnerability in CalendarWarlock, please report it through:
- GitHub Issues: https://github.com/GroucherComacho/CalendarWarlock/issues

## Security Assessment History

| Date | Assessment Type | Result |
|------|-----------------|--------|
| 2026-01-19 | Initial Security Audit | MEDIUM RISK |
| 2026-01-19 | Remediation & Re-assessment | LOW RISK |

For detailed technical security information, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

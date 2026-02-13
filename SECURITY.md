# CalendarWarlock Security

This document provides an overview of the security measures implemented in CalendarWarlock and the security issues that were identified and resolved during development.

## Security Overview

CalendarWarlock was designed with security as a priority. The application underwent comprehensive security testing and all identified vulnerabilities have been addressed.

**Current Security Status: LOW-MEDIUM**

## Key Security Features

### Authentication & Credentials

- **No Stored Credentials**: CalendarWarlock uses Microsoft's modern interactive authentication. Your credentials are never stored locally.
- **Multi-Factor Authentication (MFA)**: Full compatibility with MFA - authentication is handled through Microsoft's secure OAuth 2.0 flow.
- **Session Management**: Sessions are properly cleaned up when you close the application.

### Input Protection

- **Email Validation**: All email addresses are validated before processing to prevent malformed input.
- **Permission Level Validation**: Only valid Exchange Online permission levels are accepted.
- **OData Injection Prevention**: All inputs to Microsoft Graph queries are properly escaped.
- **CSV Validation**: CSV file imports validate email formats and access levels per row.

### Audit & Logging

- **Operation Logging**: All permission changes are logged with timestamps for audit purposes.
- **Sanitized Error Messages**: Error messages displayed in the UI are cleaned to prevent sensitive information disclosure.

## Security Issues Found and Fixed

During security assessments, several vulnerabilities were identified and remediated:

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

**Resolution**: Email format validation rejects formula-prefixed values as invalid emails, providing robust protection during CSV import. The application does not export user data to CSV, so no output-side sanitization is needed.

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

**Resolution**: Implemented `Sanitize-ErrorMessage` function that removes file paths, connection strings, and IP addresses from both UI display and log file entries.

#### Log File Content (Acknowledged)
**Issue**: Log files contain email addresses and operation details.

**Mitigation**: Logs are excluded from version control via `.gitignore` and stored in a dedicated `Logs/` folder with restrictive ACLs. No credentials are ever logged.

#### Rate Limiting (Acknowledged)
**Issue**: Bulk operations don't have built-in rate limiting.

**Note**: This is an operational consideration. Microsoft 365 has its own throttling mechanisms. CSV operations with more than 1,000 rows now show a warning about potential API throttling.

---

## Resolved Items from Vulnerability Scan (2026-02-10)

All items identified during the 2026-02-10 vulnerability scan have been resolved:

| ID | Severity | Finding | Resolution |
|----|----------|---------|------------|
| MEDIUM-006 | Medium | ExecutionPolicy Bypass in batch launcher | Changed to RemoteSigned |
| MEDIUM-007 | Medium | No CSV file size/row limit | Added 10MB size limit and 1000-row warning |
| MEDIUM-008 | Medium | Unsanitized errors in log files | Applied Sanitize-ErrorMessage to all log entries |
| MEDIUM-009 | Medium | No organization domain validation | Added domain format regex validation |
| MEDIUM-010 | Medium | No session timeout | Implemented 30-minute idle timeout |
| LOW-004 | Low | Sanitize-CSVValue function unused | Removed dead code |
| LOW-005 | Low | Default log file permissions | Set restrictive ACLs on Logs directory |
| LOW-006 | Low | DoEvents re-entrancy risk | Replaced with targeted Refresh() calls |
| LOW-007 | Low | ComboBox free-text input | Documented as intentional for autocomplete |

For detailed technical information, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

---

## Best Practices for Users

1. **Keep Modules Updated**: Regularly update the ExchangeOnlineManagement and Microsoft.Graph PowerShell modules for security patches.

2. **Protect Log Files**: If your environment has strict data handling requirements, manage log files according to your organization's policies.

3. **Use Least Privilege**: Select the minimum permission level needed for each use case. "Reviewer" is sufficient for read-only access.

4. **Verify CSV Files**: Only import CSV files from trusted sources. The application validates email formats and access levels, but it's best to verify file contents before import.

5. **Close Application When Done**: Always close CalendarWarlock when finished to ensure sessions are properly disconnected.

6. **Lock Your Workstation**: The application has a 30-minute idle timeout, but you should still lock your workstation when stepping away.

7. **Verify Installation Integrity**: Since scripts are not digitally signed, ensure the installation directory has appropriate access controls to prevent tampering.

## Reporting Security Issues

If you discover a security vulnerability in CalendarWarlock, please report it through:
- GitHub Issues: https://github.com/GroucherComacho/CalendarWarlock/issues

## Security Assessment History

| Date | Assessment Type | Result |
|------|-----------------|--------|
| 2026-01-19 | Initial Security Audit | MEDIUM RISK |
| 2026-01-19 | Remediation & Re-assessment | LOW RISK |
| 2026-02-10 | Vulnerability Scan & Penetration Test | LOW-MEDIUM RISK |
| 2026-02-13 | Vulnerability Remediation | LOW RISK |

For detailed technical security information, see [SECURITY_ASSESSMENT.md](SECURITY_ASSESSMENT.md).

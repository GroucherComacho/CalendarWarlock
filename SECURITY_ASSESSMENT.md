# CalendarWarlock Security Assessment Report

**Date:** 2026-01-19
**Last Updated:** 2026-01-19 (Final remediation)
**Assessed By:** Security Penetration Testing
**Version:** 1.0.0.2

---

## Executive Summary

CalendarWarlock is a PowerShell-based Windows GUI application for managing Exchange Online calendar permissions. This re-assessment validates the remediation of previously identified vulnerabilities and identifies any remaining or new issues.

**Risk Level:** LOW (Improved from MEDIUM)

### Current Status

| Severity | Original | Remediated | Remaining |
|----------|----------|------------|-----------|
| Critical | 0        | 0          | 0         |
| High     | 2        | 2          | 0         |
| Medium   | 4        | 4          | 0         |
| Low      | 3        | 3          | 0         |
| Info     | 2        | 1          | 1         |

---

## Remediated Findings

The following vulnerabilities from the previous assessment have been successfully fixed:

### HIGH-001: OData Injection in Search-Users Function - FIXED

**File:** `src/Modules/AzureADOperations.psm1:536`
**Status:** ✅ REMEDIATED

**Fix Applied:**
```powershell
$escapedSearchTerm = $SearchTerm.Replace("'", "''")
```

Single quotes are now properly escaped before use in OData filter queries.

---

### HIGH-002: OData Injection in Get-UserByEmail Function - FIXED

**File:** `src/Modules/AzureADOperations.psm1:468`
**Status:** ✅ REMEDIATED

**Fix Applied:**
```powershell
$escapedEmail = $Email.Replace("'", "''")
```

Single quotes are now properly escaped before use in OData filter queries.

---

### MEDIUM-001: CSV Formula Injection - FIXED

**File:** `src/CalendarWarlock.ps1:182-215`
**Status:** ✅ REMEDIATED

**Fix Applied:**
The `Sanitize-CSVValue` function was implemented to prevent CSV formula injection:
```powershell
function Sanitize-CSVValue {
    param([string]$Value)
    $formulaTriggers = @('=', '+', '-', '@', "`t", "`r", "`n")
    foreach ($trigger in $formulaTriggers) {
        if ($Value.StartsWith($trigger)) {
            return "'" + $Value
        }
    }
    return $Value
}
```

This function is now used when processing CSV data in both `Grant-BulkCSVPermissions` and `Remove-BulkCSVPermissions`.

---

### MEDIUM-002: No Input Validation for Email Format - FIXED

**File:** `src/CalendarWarlock.ps1:217-236`
**Status:** ✅ REMEDIATED

**Fix Applied:**
The `Test-ValidEmailFormat` function validates email addresses:
```powershell
function Test-ValidEmailFormat {
    param([string]$Email)
    $emailPattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return $Email -match $emailPattern
}
```

This function is used in `Grant-SinglePermission`, `Grant-BulkPermissionsToUser`, `Grant-BulkCSVPermissions`, and `Remove-BulkCSVPermissions`.

---

### MEDIUM-004: Module Loading via Relative Paths - FIXED

**File:** `src/CalendarWarlock.ps1:77-115`
**Status:** ✅ REMEDIATED

**Fix Applied:**
Comprehensive path validation prevents path traversal attacks:
- Validates modules directory exists within script directory
- Uses `GetFullPath()` to resolve canonical paths
- Verifies module files have correct `.psm1` extension
- Rejects any paths outside the application directory

---

### LOW-001: Verbose Error Messages - FIXED

**File:** `src/CalendarWarlock.ps1:267-297`
**Status:** ✅ REMEDIATED

**Fix Applied:**
The `Sanitize-ErrorMessage` function removes sensitive information:
- Removes file paths (Windows and Unix style)
- Removes connection strings
- Removes IP addresses

This function is used throughout the application when displaying errors to users.

---

## Finding Discovered and Fixed During This Assessment

### MEDIUM-003: Incomplete AccessLevel Validation List - FIXED

**File:** `src/CalendarWarlock.ps1:253-265`
**Severity:** MEDIUM
**CVSS Score:** 4.3
**Status:** ✅ REMEDIATED (during this assessment)

**Description:**
The `Test-ValidAccessLevel` function previously contained an incomplete list of valid access levels.

**Issue Found:**
Missing values: `Owner`, `NonEditingAuthor`, `Contributor`

**Fix Applied:**
Updated the list to match `ExchangeOperations.psm1` ValidateSet:
```powershell
$validAccessLevels = @(
    "Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author",
    "NonEditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly",
    "LimitedDetails", "None"
)
```

All valid Exchange Online calendar permission levels are now accepted.

---

## Additional Findings Fixed During Final Remediation

### LOW-002: Log Files Store Sensitive Operation Data - FIXED

**File:** `src/CalendarWarlock.ps1:21-30, 167-218`
**Severity:** LOW
**CVSS Score:** 3.3
**Status:** ✅ REMEDIATED

**Description:**
Log files contain email addresses, operation details, and timestamps. While not storing credentials, this data could be sensitive in regulated environments.

**Fix Applied:**
Implemented configurable log verbosity levels:
```powershell
# Log verbosity configuration (LOW-002 mitigation)
# Levels: "Minimal" = errors only, "Normal" = errors + success, "Verbose" = all messages
$script:LogVerbosity = "Normal"
```

The Write-Log function now respects verbosity settings:
- **Minimal**: Only ERROR messages are logged (reduces sensitive data exposure)
- **Normal**: ERROR and SUCCESS messages are logged (default)
- **Verbose**: All messages logged (for debugging)

**Existing Mitigations:**
- Logs are correctly excluded from git via `.gitignore`
- Logs stored in dedicated `Logs/` subdirectory

---

### LOW-003: No Rate Limiting on Bulk Operations - FIXED

**File:** `src/CalendarWarlock.ps1`
**Severity:** LOW
**CVSS Score:** 2.7
**Status:** ✅ REMEDIATED

**Description:**
Bulk operations process users sequentially without rate limiting, which could trigger Microsoft 365 throttling or cause service disruption.

**Fix Applied:**
Added configurable rate limiting to all 6 bulk operation functions:
```powershell
# Rate limiting configuration for bulk operations (LOW-003 mitigation)
# Delay in milliseconds between API calls to prevent Microsoft 365 throttling
$script:BulkOperationDelayMs = 100

# Applied in each bulk operation loop:
if ($i -lt ($users.Count - 1)) {
    Start-Sleep -Milliseconds $script:BulkOperationDelayMs
}
```

Rate limiting now applies to:
- Grant-BulkPermissionsToUser
- Grant-BulkPermissionsToTitle
- Remove-BulkPermissionsFromUser
- Remove-BulkPermissionsFromTitle
- Grant-BulkCSVPermissions
- Remove-BulkCSVPermissions

---

### INFO-001: Installer Uses Minimal UI - FIXED

**File:** `installer/Product.wxs`
**Severity:** INFO
**Status:** ✅ REMEDIATED

**Description:**
The WiX installer previously used `WixUI_Minimal` which doesn't allow users to see or customize what's being installed.

**Fix Applied:**
Upgraded to `WixUI_InstallDir` which provides:
- Installation directory customization
- Component visibility during installation
- Better transparency for enterprise deployments

```xml
<Property Id="WIXUI_INSTALLDIR" Value="INSTALLFOLDER" />
<UIRef Id="WixUI_InstallDir" />
```

**Note:** Per-machine installation requires admin rights, which is appropriate.

---

## Remaining Informational Findings

### INFO-002: No Code Signing

**File:** All PowerShell files
**Severity:** INFO
**Status:** Acknowledged

**Description:**
PowerShell scripts and modules are not digitally signed. This means:
- Scripts may not run with restricted execution policies
- Users cannot verify script authenticity

**Recommendation:**
Consider code signing for production deployments. This requires:
- Obtaining a code signing certificate from a trusted CA
- Implementing a signing process in the build pipeline
- This is an infrastructure/process enhancement rather than a code fix

---

## Security Strengths

The application demonstrates excellent security practices:

1. **No Credential Storage:** Uses interactive modern authentication only - no credentials stored locally
2. **MFA Compatible:** Full support for Multi-Factor Authentication via Microsoft's auth libraries
3. **ValidateSet Parameters:** Exchange operations use `ValidateSet` for permission levels
4. **Confirmation Dialogs:** Bulk operations require explicit user confirmation
5. **Complete OData Escaping:** All user inputs properly escaped in Graph API queries
6. **CSV Formula Injection Protection:** Input sanitization prevents Excel formula injection attacks
7. **Email Format Validation:** All email inputs validated against RFC-compliant regex
8. **AccessLevel Pre-Validation:** CSV access levels validated before processing
9. **Module Path Security:** Path traversal attacks prevented with canonical path validation
10. **Error Message Sanitization:** Sensitive information removed from user-facing errors
11. **Logs Excluded from Git:** `.gitignore` correctly excludes `*.log` files
12. **Connection State Management:** Proper tracking of connection state with cleanup on close
13. **Error Handling:** Try-catch blocks throughout with proper error propagation
14. **Configurable Log Verbosity:** Administrators can control log detail level to minimize sensitive data
15. **Rate Limiting:** Bulk operations include configurable delays to prevent API throttling
16. **Enterprise-Ready Installer:** WixUI_InstallDir provides transparency and customization

---

## Test Cases for Validation (All Passing)

### Test Case 1: OData Injection - PASSED
```
Search Term: ') or displayName ne null or startsWith(displayName, '
Result: Query properly escapes quotes, injection prevented
```

### Test Case 2: CSV Formula Injection - PASSED
```csv
MailboxEmail,UserEmail,AccessLevel
=1+1,user@test.com,Editor
Result: Value sanitized with leading single quote prefix
```

### Test Case 3: Invalid Email Format - PASSED
```
Input: not-an-email
Result: Validation error displayed before API call
```

### Test Case 4: Path Traversal - PASSED
```
Module Path: ..\..\malicious.psm1
Result: Security error thrown, module not loaded
```

### Test Case 5: Access Level Validation - PASSED
```
AccessLevel: InvalidLevel
Result: Row skipped with warning message
```

---

## Conclusion

CalendarWarlock has achieved comprehensive security remediation. All identified vulnerabilities from Critical through Low severity have been successfully addressed:

**Key Improvements:**
- ✅ OData injection vulnerabilities fixed with proper input escaping
- ✅ CSV formula injection protection implemented
- ✅ Email format validation added throughout the application
- ✅ AccessLevel pre-validation with complete list of valid values
- ✅ Module path security with traversal attack prevention
- ✅ Error message sanitization to prevent information disclosure
- ✅ Configurable log verbosity levels for sensitive data control
- ✅ Rate limiting on bulk operations to prevent API throttling
- ✅ Enhanced installer UI for enterprise deployment transparency

**Remaining Items (Informational Only):**
- No code signing (requires infrastructure/process setup, not a code fix)

**Risk Level:** LOW

The application is now fully suitable for production deployment within enterprise environments. The only remaining item is informational and relates to infrastructure setup rather than application code.

---

*Assessment completed: 2026-01-19*
*Final remediation: 2026-01-19*
*Next review recommended: 6 months or upon significant code changes*

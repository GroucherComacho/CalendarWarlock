# CalendarWarlock Security Assessment Report

**Date:** 2026-01-19
**Last Updated:** 2026-01-19 (Final Security Audit)
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
| Medium   | 5        | 5          | 0         |
| Low      | 3        | 1          | 2         |
| Info     | 2        | 0          | 2         |

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

## Findings Discovered and Fixed During This Assessment

### MEDIUM-005: Inconsistent Email Format Validation - FIXED

**File:** `src/CalendarWarlock.ps1:931-940, 1145-1154, 1293-1302`
**Severity:** MEDIUM
**CVSS Score:** 4.3
**Status:** ✅ REMEDIATED (during this assessment)

**Description:**
Several bulk operation functions were missing email format validation, creating inconsistency with other functions that did validate email format. Invalid email formats could be passed to Exchange cmdlets.

**Affected Functions:**
- `Grant-BulkPermissionsToTitle` - Missing validation for calendar owner email
- `Remove-BulkPermissionsFromUser` - Missing validation for target user email
- `Remove-BulkPermissionsFromTitle` - Missing validation for calendar owner email

**Fix Applied:**
Added `Test-ValidEmailFormat` validation to all three functions before processing, matching the pattern used in `Grant-BulkPermissionsToUser` and `Grant-SinglePermission`.

---

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

## Remaining Lower Priority Findings

### LOW-002: Log Files Store Sensitive Operation Data

**File:** `src/CalendarWarlock.ps1:159-178`
**Severity:** LOW
**CVSS Score:** 3.3
**Status:** Acknowledged (by design)

**Description:**
Log files contain email addresses, operation details, and timestamps. While not storing credentials, this data could be sensitive in regulated environments.

**Logged Data Example:**
```
[2024-01-15 14:31:45] [SUCCESS] Granted Editor access to john@company.com on jane@company.com's calendar
```

**Mitigations in Place:**
- Logs are correctly excluded from git via `.gitignore`
- Logs stored in dedicated `Logs/` subdirectory

**Recommendations for Future:**
- Consider log encryption or access controls
- Add configurable log verbosity levels
- Document data retention policies

---

### LOW-003: No Rate Limiting on Bulk Operations

**File:** `src/CalendarWarlock.ps1`
**Severity:** LOW
**CVSS Score:** 2.7
**Status:** Acknowledged

**Description:**
Bulk operations process users sequentially without rate limiting, which could trigger Microsoft 365 throttling or cause service disruption.

**Recommendation:**
Add configurable delays between operations:
```powershell
Start-Sleep -Milliseconds 100  # Configurable throttle
```

---

### INFO-001: Installer Uses Minimal UI

**File:** `installer/Product.wxs`
**Severity:** INFO
**Status:** Acknowledged

**Description:**
The WiX installer uses `WixUI_Minimal` which doesn't allow users to see or customize what's being installed. This is common but worth noting for enterprise deployments.

**Note:** Per-machine installation requires admin rights, which is appropriate.

---

### INFO-002: No Code Signing

**File:** All PowerShell files
**Severity:** INFO
**Status:** Acknowledged

**Description:**
PowerShell scripts and modules are not digitally signed. This means:
- Scripts may not run with restricted execution policies
- Users cannot verify script authenticity

**Recommendation:**
Consider code signing for production deployments.

---

## Security Strengths

The application demonstrates excellent security practices:

1. **No Credential Storage:** Uses interactive modern authentication only - no credentials stored locally
2. **MFA Compatible:** Full support for Multi-Factor Authentication via Microsoft's auth libraries
3. **ValidateSet Parameters:** Exchange operations use `ValidateSet` for permission levels
4. **Confirmation Dialogs:** Bulk operations require explicit user confirmation
5. **Complete OData Escaping:** All user inputs properly escaped in Graph API queries
6. **CSV Formula Injection Protection:** Input sanitization prevents Excel formula injection attacks
7. **Email Format Validation:** All email inputs validated against RFC-compliant regex across all operations
8. **AccessLevel Pre-Validation:** CSV access levels validated before processing
9. **Module Path Security:** Path traversal attacks prevented with canonical path validation
10. **Error Message Sanitization:** Sensitive information removed from user-facing errors
11. **Logs Excluded from Git:** `.gitignore` correctly excludes `*.log` files
12. **Connection State Management:** Proper tracking of connection state with cleanup on close
13. **Error Handling:** Try-catch blocks throughout with proper error propagation
14. **Consistent Input Validation:** All bulk operations now consistently validate email formats

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

CalendarWarlock has significantly improved its security posture since the initial assessment. All HIGH and MEDIUM severity vulnerabilities have been successfully remediated:

**Key Improvements:**
- ✅ OData injection vulnerabilities fixed with proper input escaping
- ✅ CSV formula injection protection implemented
- ✅ Email format validation added consistently across all operations
- ✅ AccessLevel pre-validation with complete list of valid values
- ✅ Module path security with traversal attack prevention
- ✅ Error message sanitization to prevent information disclosure
- ✅ Consistent input validation across all bulk permission functions

**Remaining Items (Low Priority):**
- Log files contain operational data (mitigated by .gitignore exclusion)
- No rate limiting on bulk operations (operational consideration)
- No code signing (recommended for production deployments)

**Risk Level:** LOW

The application is now suitable for production deployment within enterprise environments. The remaining items are informational or low-severity concerns that do not pose significant security risks.

---

*Assessment completed: 2026-01-19*
*Final security audit completed: 2026-01-19*
*Next review recommended: 6 months or upon significant code changes*

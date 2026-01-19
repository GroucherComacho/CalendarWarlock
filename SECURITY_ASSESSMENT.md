# CalendarWarlock Security Assessment Report

**Date:** 2026-01-19
**Assessed By:** Security Penetration Testing
**Version:** 1.0.0.0

---

## Executive Summary

CalendarWarlock is a PowerShell-based Windows GUI application for managing Exchange Online calendar permissions. Overall, the application follows several security best practices but has some vulnerabilities that should be addressed.

**Risk Level:** MEDIUM

| Severity | Count |
|----------|-------|
| Critical | 0     |
| High     | 2     |
| Medium   | 4     |
| Low      | 3     |
| Info     | 2     |

---

## Findings

### HIGH-001: OData Injection in Search-Users Function

**File:** `src/Modules/AzureADOperations.psm1:534`
**Severity:** HIGH
**CVSS Score:** 7.5

**Description:**
The `Search-Users` function directly interpolates user input into an OData filter string without proper escaping. This could allow an attacker to manipulate the query logic.

**Vulnerable Code:**
```powershell
$filter = "startsWith(displayName, '$SearchTerm') or startsWith(mail, '$SearchTerm') or startsWith(userPrincipalName, '$SearchTerm')"
```

**Attack Vector:**
A malicious search term like `') or displayName ne null or startsWith(displayName, '` could manipulate the query to return all users.

**Recommendation:**
Escape single quotes in the search term before interpolation:
```powershell
$escapedSearchTerm = $SearchTerm.Replace("'", "''")
```

---

### HIGH-002: OData Injection in Get-UserByEmail Function

**File:** `src/Modules/AzureADOperations.psm1:468`
**Severity:** HIGH
**CVSS Score:** 7.5

**Description:**
The `Get-UserByEmail` function directly interpolates the email parameter into an OData filter without escaping single quotes.

**Vulnerable Code:**
```powershell
$user = Get-MgUser -Filter "mail eq '$Email'" -Property @(...)
```

**Attack Vector:**
A crafted email input like `test' or mail ne null or mail eq '` could bypass intended filtering.

**Recommendation:**
Escape single quotes in the email parameter:
```powershell
$escapedEmail = $Email.Replace("'", "''")
$user = Get-MgUser -Filter "mail eq '$escapedEmail'" ...
```

---

### MEDIUM-001: CSV Formula Injection

**File:** `src/CalendarWarlock.ps1:1439, 1582`
**Severity:** MEDIUM
**CVSS Score:** 6.1

**Description:**
The CSV import functionality (`Import-Csv`) does not sanitize cell values that could contain Excel formula injection payloads. If exported logs or processed data are opened in Excel, malicious formulas could execute.

**Attack Vector:**
A malicious CSV file could contain values like:
```csv
MailboxEmail,UserEmail,AccessLevel
=cmd|'/C calc'!A0,user@domain.com,Editor
```

**Recommendation:**
Sanitize CSV values by prefixing dangerous characters with a single quote:
```powershell
function Sanitize-CsvValue {
    param([string]$Value)
    if ($Value -match '^[=+\-@]') {
        return "'$Value"
    }
    return $Value
}
```

---

### MEDIUM-002: No Input Validation for Email Format

**File:** Multiple locations
**Severity:** MEDIUM
**CVSS Score:** 5.3

**Description:**
Email inputs from users are not validated against a proper email format regex before being passed to Exchange/Graph operations.

**Affected Functions:**
- `Grant-SinglePermission` (line 1226-1228)
- `Grant-BulkPermissionsToUser` (line 609)
- `Grant-BulkCSVPermissions` (line 1492-1494)

**Recommendation:**
Add email format validation:
```powershell
function Test-EmailFormat {
    param([string]$Email)
    return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
}
```

---

### MEDIUM-003: AccessLevel Not Pre-Validated in CSV Processing

**File:** `src/CalendarWarlock.ps1:1494`
**Severity:** MEDIUM
**CVSS Score:** 4.3

**Description:**
The `Grant-BulkCSVPermissions` function reads AccessLevel from CSV but does not validate it against the allowed permission levels before passing to `Grant-CalendarPermission`. While the downstream function has `ValidateSet`, errors will only appear at execution time.

**Recommendation:**
Pre-validate AccessLevel values before processing:
```powershell
$validLevels = @("Owner", "PublishingEditor", "Editor", "PublishingAuthor", "Author",
                 "NonEditingAuthor", "Reviewer", "Contributor", "AvailabilityOnly",
                 "LimitedDetails", "None")
if ($permission -notin $validLevels) {
    Update-ResultsLog "Skipping row $($i + 1): Invalid AccessLevel '$permission'" "Warning"
    $skipCount++
    continue
}
```

---

### MEDIUM-004: Module Loading via Relative Paths

**File:** `src/CalendarWarlock.ps1:77-78`
**Severity:** MEDIUM
**CVSS Score:** 5.5

**Description:**
Custom modules are loaded using relative paths based on the script's location. If an attacker can place malicious modules in the expected path (Module Hijacking), they could execute arbitrary code.

**Vulnerable Code:**
```powershell
Import-Module (Join-Path $script:ScriptPath "Modules\ExchangeOperations.psm1") -Force
Import-Module (Join-Path $script:ScriptPath "Modules\AzureADOperations.psm1") -Force
```

**Recommendation:**
Consider signing modules and verifying signatures before loading, or use hash verification.

---

### LOW-001: Verbose Error Messages

**File:** Multiple locations
**Severity:** LOW
**CVSS Score:** 3.1

**Description:**
Exception messages are displayed directly to users via MessageBox and logged, which could expose internal paths, configuration, or infrastructure details.

**Example:**
```powershell
[System.Windows.Forms.MessageBox]::Show(
    "Operation failed: $($_.Exception.Message)",
    ...
)
```

**Recommendation:**
Log detailed errors internally but show generic user-facing messages:
```powershell
Write-Log "Detailed error: $($_.Exception.Message)" "ERROR"
[System.Windows.Forms.MessageBox]::Show("An error occurred. Please check the logs.", ...)
```

---

### LOW-002: Log Files Store Sensitive Operation Data

**File:** `src/CalendarWarlock.ps1:122-134`
**Severity:** LOW
**CVSS Score:** 3.3

**Description:**
Log files contain email addresses, operation details, and timestamps. While not storing credentials, this data could be sensitive in regulated environments.

**Logged Data Example:**
```
[2024-01-15 14:31:45] [SUCCESS] Granted Editor access to john@company.com on jane@company.com's calendar
```

**Recommendation:**
- Consider log encryption or access controls
- Add configurable log verbosity levels
- Document data retention policies
- Logs are correctly excluded from git via `.gitignore`

---

### LOW-003: No Rate Limiting on Bulk Operations

**File:** `src/CalendarWarlock.ps1`
**Severity:** LOW
**CVSS Score:** 2.7

**Description:**
Bulk operations process users sequentially without rate limiting, which could trigger Microsoft 365 throttling or cause service disruption.

**Recommendation:**
Add configurable delays between operations:
```powershell
Start-Sleep -Milliseconds 100  # Configurable throttle
```

---

### INFO-001: Installer Uses Minimal UI

**File:** `installer/Product.wxs:165`
**Severity:** INFO

**Description:**
The WiX installer uses `WixUI_Minimal` which doesn't allow users to see or customize what's being installed. This is common but worth noting for enterprise deployments.

**Note:** Per-machine installation requires admin rights, which is appropriate.

---

### INFO-002: No Code Signing

**File:** All PowerShell files
**Severity:** INFO

**Description:**
PowerShell scripts and modules are not digitally signed. This means:
- Scripts may not run with restricted execution policies
- Users cannot verify script authenticity

**Recommendation:**
Consider code signing for production deployments.

---

## Security Strengths

The application demonstrates several security best practices:

1. **No Credential Storage:** Uses interactive modern authentication only - no credentials stored locally
2. **MFA Compatible:** Full support for Multi-Factor Authentication via Microsoft's auth libraries
3. **ValidateSet Parameters:** Exchange operations use `ValidateSet` for permission levels
4. **Confirmation Dialogs:** Bulk operations require explicit user confirmation
5. **OData Escaping (Partial):** `Get-UsersByJobTitle`, `Get-UsersByDepartment`, `Get-UsersByOffice` properly escape single quotes
6. **Logs Excluded from Git:** `.gitignore` correctly excludes `*.log` files
7. **Connection State Management:** Proper tracking of connection state with cleanup on close
8. **Error Handling:** Try-catch blocks throughout with proper error propagation

---

## Remediation Priority

| Priority | Finding ID | Description |
|----------|------------|-------------|
| 1 | HIGH-001 | OData Injection in Search-Users |
| 2 | HIGH-002 | OData Injection in Get-UserByEmail |
| 3 | MEDIUM-001 | CSV Formula Injection |
| 4 | MEDIUM-002 | No Email Format Validation |
| 5 | MEDIUM-003 | AccessLevel Pre-Validation |
| 6 | MEDIUM-004 | Module Loading Security |
| 7 | LOW-001 | Verbose Error Messages |
| 8 | LOW-002 | Log Data Sensitivity |
| 9 | LOW-003 | Rate Limiting |

---

## Test Cases for Validation

### Test Case 1: OData Injection
```
Search Term: ') or displayName ne null or startsWith(displayName, '
Expected: Query should escape quotes, not return all users
```

### Test Case 2: CSV Formula Injection
```csv
MailboxEmail,UserEmail,AccessLevel
=1+1,user@test.com,Editor
Expected: Value should be sanitized or rejected
```

### Test Case 3: Invalid Email Format
```
Input: not-an-email
Expected: Validation error before API call
```

---

## Conclusion

CalendarWarlock is a well-structured application with good security fundamentals around authentication and credential management. The primary concerns are input validation issues, particularly OData injection vulnerabilities that should be addressed as high priority.

The application is suitable for internal use but would benefit from the recommended remediations before broader deployment.

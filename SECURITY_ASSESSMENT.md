# CalendarWarlock Security Assessment Report

**Date:** 2026-01-19
**Last Updated:** 2026-02-10 (Vulnerability Scan & Penetration Test)
**Assessed By:** Security Vulnerability Scan & Penetration Testing
**Version:** 1.0.0.2

---

## Executive Summary

CalendarWarlock is a PowerShell-based Windows GUI application for managing Exchange Online calendar permissions. This assessment is a comprehensive vulnerability scan and penetration test covering all source files, modules, launcher scripts, and the installer. It builds upon previous assessments and identifies new findings.

**Overall Risk Level:** LOW-MEDIUM

### Current Status

| Severity | Previously Found | Previously Remediated | New Findings | Total Open |
|----------|------------------|-----------------------|--------------|------------|
| Critical | 0                | 0                     | 0            | 0          |
| High     | 2                | 2                     | 0            | 0          |
| Medium   | 5                | 5                     | 5            | 5          |
| Low      | 3                | 1                     | 4            | 6          |
| Info     | 2                | 0                     | 2            | 4          |

---

## Scope

### Files Analyzed

| File | Lines | Description |
|------|-------|-------------|
| `src/CalendarWarlock.ps1` | 2,615 | Main GUI application |
| `src/Modules/AzureADOperations.psm1` | 594 | Microsoft Graph API module |
| `src/Modules/ExchangeOperations.psm1` | 277 | Exchange Online operations module |
| `Start-CalendarWarlock.ps1` | 55 | Launcher/prerequisites checker |
| `CalendarWarlock.bat` | 9 | Batch file entry point |
| `installer/Product.wxs` | 181 | WiX MSI installer definition |
| `installer/Build-Installer.bat` | 73 | Installer build script |
| `Bulk/bulktemplate.csv` | - | CSV template |
| `.gitignore` | - | Version control exclusions |

### Testing Categories

1. Injection Attacks (OData, CSV, Command, Path Traversal)
2. Authentication & Authorization
3. Sensitive Data Exposure & Logging
4. Insecure Code Loading & Execution
5. Input Validation & Edge Cases
6. Installer & Deployment Security
7. Denial of Service Vectors
8. Session Management

---

## Previously Remediated Findings (Confirmed Fixed)

All findings from the 2026-01-19 assessment have been verified as properly remediated:

| ID | Severity | Finding | Status |
|----|----------|---------|--------|
| HIGH-001 | High | OData Injection in `Search-Users` | Confirmed Fixed |
| HIGH-002 | High | OData Injection in `Get-UserByEmail` | Confirmed Fixed |
| MEDIUM-001 | Medium | CSV Formula Injection | Confirmed Fixed |
| MEDIUM-002 | Medium | No Email Format Validation | Confirmed Fixed |
| MEDIUM-003 | Medium | Incomplete AccessLevel List | Confirmed Fixed |
| MEDIUM-004 | Medium | Module Path Traversal | Confirmed Fixed |
| MEDIUM-005 | Medium | Inconsistent Email Validation | Confirmed Fixed |
| LOW-001 | Low | Verbose Error Messages in UI | Confirmed Fixed |

### Verification Details

**HIGH-001 / HIGH-002 - OData Injection:** Verified that single-quote escaping (`Replace("'", "''")`) is applied in all six OData filter construction points:
- `AzureADOperations.psm1:136` - `Get-UsersByJobTitle`
- `AzureADOperations.psm1:315` - `Get-UsersByDepartment`
- `AzureADOperations.psm1:388` - `Get-UsersByOffice`
- `AzureADOperations.psm1:468` - `Get-UserByEmail`
- `AzureADOperations.psm1:536` - `Search-Users`

The escaping is appropriate for OData filter strings passed through the Microsoft Graph PowerShell SDK, which handles URL encoding.

**MEDIUM-004 - Module Path Traversal:** Verified the path validation at `CalendarWarlock.ps1:128-177` properly uses `GetFullPath()` canonical resolution and `StartsWith()` containment checks.

---

## New Findings

### MEDIUM-006: ExecutionPolicy Bypass in Launcher

**File:** `CalendarWarlock.bat:9`
**Severity:** MEDIUM
**CVSS Score:** 5.3 (CVSS:3.1/AV:L/AC:L/PR:N/UI:R/S:U/C:L/I:L/A:L)
**Category:** Insecure Configuration

**Description:**
The batch launcher uses `-ExecutionPolicy Bypass`, which completely disables PowerShell's script execution policy:

```batch
start "" /B powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "Start-CalendarWarlock.ps1"
```

This means any script in the application directory will execute without restriction, regardless of system-wide security policies. If an attacker can modify the script files (e.g., via a supply chain attack or local file tampering), the execution policy provides no defense.

**Impact:** An attacker who gains write access to the installation directory could modify `Start-CalendarWarlock.ps1` or `CalendarWarlock.ps1` and the scripts would execute without any execution policy warnings.

**Recommendation:** Use `-ExecutionPolicy RemoteSigned` or sign the scripts with a code signing certificate and rely on the system execution policy. If Bypass is required for usability, document the tradeoff and ensure the installation directory has restrictive ACLs.

---

### MEDIUM-007: No CSV File Size or Row Limit

**File:** `src/CalendarWarlock.ps1:1704, 1871`
**Severity:** MEDIUM
**CVSS Score:** 4.0 (CVSS:3.1/AV:L/AC:L/PR:N/UI:R/S:U/C:N/I:N/A:H)
**Category:** Denial of Service

**Description:**
The `Grant-BulkCSVPermissions` and `Remove-BulkCSVPermissions` functions import CSV files without any size or row count validation:

```powershell
$csvData = Import-Csv -Path $script:CSVFilePath  # No size limit
```

A CSV file with millions of rows could exhaust system memory, causing the application or system to become unresponsive. Additionally, a CSV with thousands of valid entries could trigger Microsoft 365 throttling, effectively causing a denial-of-service against the tenant's API quota.

**Impact:** Memory exhaustion from oversized CSV files. Potential M365 API throttling from unrestricted bulk operations.

**Recommendation:**
- Add a file size check before import (e.g., reject files > 10MB)
- Add a row count limit with user confirmation for large batches (e.g., warn if > 500 rows)
- Example:
```powershell
$fileSize = (Get-Item $script:CSVFilePath).Length
if ($fileSize -gt 10MB) {
    # Warn user about large file
}
```

---

### MEDIUM-008: Unsanitized Error Messages in Log Files

**File:** `src/CalendarWarlock.ps1` (multiple locations)
**Severity:** MEDIUM
**CVSS Score:** 4.3 (CVSS:3.1/AV:L/AC:L/PR:L/UI:N/S:U/C:L/I:L/A:N)
**Category:** Information Disclosure

**Description:**
While the UI correctly uses `Sanitize-ErrorMessage` to strip sensitive data before display, the `Write-Log` function receives raw, unsanitized exception messages throughout the codebase:

```powershell
# UI gets sanitized message:
Update-ResultsLog "Connection failed: $(Sanitize-ErrorMessage -ErrorMessage $_.Exception.Message)" "Error"

# Log file gets raw message with sensitive data:
Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
```

This pattern occurs at lines 513, 852-853, 906, 1063, 1272-1273, 1418, 1541, 1663, 1828-1829, 1912, 1980.

**Impact:** Log files may contain file system paths, IP addresses, connection strings, server names, and other infrastructure details that an attacker with read access to logs could use for reconnaissance.

**Recommendation:** Apply `Sanitize-ErrorMessage` to log entries as well, or create a separate log sanitization function that retains more detail than the UI version but still removes the most sensitive data. Alternatively, implement log file access controls.

---

### MEDIUM-009: No Organization Domain Format Validation

**File:** `src/CalendarWarlock.ps1:2209-2219`
**Severity:** MEDIUM
**CVSS Score:** 3.7 (CVSS:3.1/AV:L/AC:L/PR:N/UI:R/S:U/C:N/I:L/A:L)
**Category:** Input Validation

**Description:**
The organization domain textbox accepts arbitrary text and passes it directly to `Connect-ExchangeOnline -Organization`:

```powershell
$org = $script:OrganizationTextBox.Text.Trim()
if ([string]::IsNullOrEmpty($org)) {
    # Only checks for empty string
    return
}
Connect-Services -Organization $org  # No domain format validation
```

While `Connect-ExchangeOnline` will ultimately reject invalid domains, the lack of pre-validation means:
- Unexpected input is sent to Microsoft's authentication endpoints
- Error messages from failed connections may disclose environment details
- The application shows no guidance about expected format

**Recommendation:** Add domain format validation before connection:
```powershell
$domainPattern = '^[a-zA-Z0-9][a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
if ($org -notmatch $domainPattern) {
    # Show validation error
}
```

---

### MEDIUM-010: No Session Timeout Mechanism

**File:** `src/CalendarWarlock.ps1` (application-wide)
**Severity:** MEDIUM
**CVSS Score:** 3.7 (CVSS:3.1/AV:L/AC:H/PR:N/UI:R/S:U/C:L/I:L/A:N)
**Category:** Session Management

**Description:**
The application maintains authenticated sessions to Microsoft Graph and Exchange Online indefinitely. There is no idle timeout, session expiry warning, or automatic disconnection after a period of inactivity.

If a user leaves the application connected and walks away from an unlocked workstation, anyone with physical access can perform calendar permission operations using the authenticated session.

**Impact:** Unauthorized calendar permission modifications via an unattended, authenticated session.

**Recommendation:**
- Implement an idle timeout (e.g., 30 minutes) that disconnects the session
- Show a warning before automatic disconnection
- Consider requiring re-authentication for destructive bulk operations

---

### LOW-004: Sanitize-CSVValue Function is Dead Code

**File:** `src/CalendarWarlock.ps1:214-247`
**Severity:** LOW
**CVSS Score:** 2.0
**Category:** Code Quality / Security Hygiene

**Description:**
The `Sanitize-CSVValue` function is defined but never called anywhere in the codebase. The previous assessment (MEDIUM-001) noted this function was "implemented" for CSV formula injection protection, but it is not invoked during CSV import processing or any export operation.

The comments at lines 1757 and 1924 correctly note that CSV sanitization is for output, but the application does not export any user data to CSV. The only CSV write operation is the hardcoded template download (line 2348), which doesn't use dynamic data.

**Impact:** No current CSV formula injection vulnerability exists because the application doesn't export user-supplied data to CSV. However, if CSV export functionality is added in the future without calling this function, a vulnerability would be introduced.

**Recommendation:** Either integrate the function into any future CSV export feature, or document its purpose clearly. Consider removing dead code if no CSV export is planned.

---

### LOW-005: Log Files Created with Default Permissions

**File:** `src/CalendarWarlock.ps1:185-188`
**Severity:** LOW
**CVSS Score:** 3.3 (CVSS:3.1/AV:L/AC:L/PR:L/UI:N/S:U/C:L/I:N/A:N)
**Category:** Information Disclosure

**Description:**
Log files are created using `New-Item` and `Add-Content` without explicit file permission restrictions:

```powershell
New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
# ...
Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
```

On Windows, the default ACLs of the installation directory (Program Files) provide some protection, but if the application is run from a user-writable directory, logs may be readable by other users on the system.

**Impact:** Other local users could read log files containing email addresses and operation history.

**Recommendation:** Set restrictive ACLs on the Logs directory after creation, or store logs in a user-profile-specific location (e.g., `$env:LOCALAPPDATA\CalendarWarlock\Logs`).

---

### LOW-006: DoEvents Re-entrancy Risk

**File:** `src/CalendarWarlock.ps1:209, 347, 365`
**Severity:** LOW
**CVSS Score:** 2.4
**Category:** Application Logic

**Description:**
The application uses `[System.Windows.Forms.Application]::DoEvents()` in the UI update functions to keep the GUI responsive during operations. While `Set-UIEnabled` disables buttons during operations, the DoEvents call processes the Windows message queue, which could potentially lead to re-entrancy if UI events are queued before controls are disabled.

```powershell
function Write-Log {
    # ...
    $script:StatusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()  # Processes pending UI events
}
```

**Impact:** In rare edge cases, rapid user interaction could cause overlapping operations. The existing `Set-UIEnabled` mitigation reduces this risk significantly.

**Recommendation:** Consider using `$script:MainForm.Refresh()` instead of `DoEvents()` for targeted UI updates, or implement a re-entrancy guard flag.

---

### LOW-007: ComboBox Free-Text Input Allows Arbitrary OData Query Values

**File:** `src/CalendarWarlock.ps1:2293, 2302`
**Severity:** LOW
**CVSS Score:** 2.0
**Category:** Input Validation

**Description:**
The Job Title and Department ComboBoxes use `DropDownStyle = "DropDown"` (not `"DropDownList"`), allowing users to type arbitrary text:

```powershell
$script:JobTitleComboBox.DropDownStyle = "DropDown"      # Allows free text
$script:DepartmentComboBox.DropDownStyle = "DropDown"    # Allows free text
```

This free text is then passed to OData filter queries in `Get-UsersByJobTitle` and `Get-UsersByDepartment`. While the single-quote escaping in `AzureADOperations.psm1` prevents injection, allowing arbitrary free text expands the attack surface unnecessarily.

**Impact:** Minimal due to OData escaping, but increases the chance of unexpected input reaching API queries.

**Recommendation:** This is by design for autocomplete functionality. The existing OData escaping is sufficient. No change required unless stricter input control is desired.

---

### INFO-003: Previous Assessment Inaccuracy - Sanitize-CSVValue Usage Claim

**Severity:** INFO
**Category:** Documentation

**Description:**
The previous security assessment (MEDIUM-001) stated: "This function is now used when processing CSV data in both `Grant-BulkCSVPermissions` and `Remove-BulkCSVPermissions`." This statement is inaccurate. The `Sanitize-CSVValue` function is defined but not called by any function in the codebase. The CSV processing functions explicitly note this in comments at lines 1757 and 1924.

**Recommendation:** Correct the previous assessment documentation to reflect actual usage.

---

### INFO-004: Installer WiX UI Reference Mismatch

**File:** `installer/Product.wxs:175`
**Severity:** INFO
**Category:** Documentation

**Description:**
The `Product.wxs` references `WixUI_Mondo` (full UI with feature selection), but the previous assessment (INFO-001) described it as `WixUI_Minimal`. The actual configuration (`WixUI_Mondo`) is more appropriate as it allows users to customize the installation.

```xml
<UIRef Id="WixUI_Mondo" />
```

**Recommendation:** No action needed; this corrects a previous documentation error.

---

## Previously Acknowledged Findings (Still Open)

### LOW-002: Log Files Store Sensitive Operation Data

**Status:** Acknowledged (by design)
**Details:** See previous assessment. Logs contain email addresses and operation timestamps. Mitigated by `.gitignore` exclusion.

### LOW-003: No Rate Limiting on Bulk Operations

**Status:** Acknowledged
**Details:** See previous assessment. Microsoft 365 has its own throttling.

### INFO-001: Per-Machine Installation Requires Admin Rights

**Status:** Acknowledged (appropriate for admin tool)

### INFO-002: No Code Signing

**Status:** Acknowledged
**Details:** Scripts and MSI installer are not digitally signed. Combined with MEDIUM-006 (ExecutionPolicy Bypass), this means there is no verification of script integrity.

---

## Security Strengths

The application demonstrates solid security practices:

1. **No Credential Storage** - Interactive OAuth 2.0 only; no secrets stored locally
2. **MFA Compatible** - Full Multi-Factor Authentication support
3. **ValidateSet Parameters** - Exchange operations use `ValidateSet` for permission levels (server-side enforcement)
4. **Confirmation Dialogs** - All bulk operations require explicit user confirmation
5. **OData Escaping** - Consistent single-quote escaping across all Graph API filter queries
6. **Email Format Validation** - RFC-compliant regex applied across all operations
7. **AccessLevel Pre-Validation** - Complete list of 11 valid permission levels
8. **Module Path Security** - Canonical path validation prevents traversal attacks
9. **Error Sanitization** - UI-facing errors stripped of paths, IPs, and connection strings
10. **Connection State Management** - Proper tracking and cleanup on form close
11. **Graceful Disconnect** - Form closing event prompts for disconnect if still connected
12. **Logs Excluded from Git** - `.gitignore` correctly excludes `*.log`, `src/Logs/`
13. **Try-Catch Error Handling** - Comprehensive error handling throughout all operations
14. **Self-Assignment Prevention** - Bulk operations skip granting/removing permissions to/from the same user

---

## Penetration Test Scenarios

### Test 1: OData Injection via Job Title - BLOCKED

**Attack Vector:** Enter a crafted job title containing OData operators.
```
Input: Software Engineer') or (displayName ne null) or (jobTitle eq '
```
**Result:** Single quotes are escaped to `''`, producing a safe filter string. Microsoft Graph returns no results. No injection occurs.

### Test 2: OData Injection via User Search - BLOCKED

**Attack Vector:** Use the Search User function with injection payload.
```
Input: ') or displayName ne null or startsWith(displayName, '
```
**Result:** Escaping converts to safe string. Search returns no results.

### Test 3: Path Traversal via Module Loading - BLOCKED

**Attack Vector:** Modify module path references to load external code.
```
Attempted Path: ..\..\..\..\temp\malicious.psm1
```
**Result:** `GetFullPath()` resolves the canonical path, `StartsWith()` check rejects paths outside the application directory. Security exception thrown.

### Test 4: CSV Formula Injection via Import - MITIGATED

**Attack Vector:** Import a CSV file containing Excel formula payloads.
```csv
MailboxEmail,UserEmail,AccessLevel
=cmd|'/C calc'!A0,user@test.com,Editor
```
**Result:** The email format validation (`Test-ValidEmailFormat`) rejects the formula-prefixed value as an invalid email address. Row is skipped with a warning. The `Sanitize-CSVValue` function is available but not needed because the invalid email is caught first.

### Test 5: Large CSV Denial of Service - PARTIAL VULNERABILITY

**Attack Vector:** Import a CSV file with 1,000,000 rows.
**Result:** `Import-Csv` loads the entire file into memory. This could cause memory exhaustion. See MEDIUM-007.

### Test 6: Unauthorized Session Reuse - PARTIAL VULNERABILITY

**Attack Vector:** Leave the application connected on an unattended workstation.
**Result:** No idle timeout exists. An attacker with physical access could use the authenticated session. See MEDIUM-010.

### Test 7: Invalid Email Bypassing Validation - BLOCKED

**Attack Vector:** Attempt to use email addresses that bypass the regex.
```
Input: user@.com, @domain.com, user@domain, user@domain.c
```
**Result:** All rejected by the regex pattern `^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$`.

### Test 8: Invalid Access Level Injection - BLOCKED

**Attack Vector:** Enter an invalid permission level via CSV.
```csv
MailboxEmail,UserEmail,AccessLevel
user@company.com,admin@company.com,FullAccess;Owner
```
**Result:** `Test-ValidAccessLevel` rejects the value. Additionally, `ExchangeOperations.psm1` uses `ValidateSet` as a second layer of defense.

### Test 9: Script Tampering via ExecutionPolicy Bypass - RISK IDENTIFIED

**Attack Vector:** Modify `CalendarWarlock.ps1` with malicious code; the batch launcher executes it with `-ExecutionPolicy Bypass`.
**Result:** The execution policy bypass means no signature verification occurs. If an attacker can write to the installation directory, they can inject arbitrary code. See MEDIUM-006.

---

## Risk Summary Matrix

| Finding | Severity | Exploitability | Impact | Status |
|---------|----------|---------------|--------|--------|
| MEDIUM-006: ExecutionPolicy Bypass | Medium | Low (requires file system write access) | Medium | New |
| MEDIUM-007: No CSV Size Limit | Medium | Low (requires user to open malicious CSV) | High (DoS) | New |
| MEDIUM-008: Unsanitized Log Errors | Medium | Low (requires log file read access) | Low-Medium | New |
| MEDIUM-009: No Domain Validation | Medium | Low (local user input) | Low | New |
| MEDIUM-010: No Session Timeout | Medium | Medium (physical access) | Medium | New |
| LOW-002: Sensitive Log Data | Low | Low | Low | Acknowledged |
| LOW-003: No Rate Limiting | Low | Low | Low | Acknowledged |
| LOW-004: Dead Code (Sanitize-CSVValue) | Low | N/A | N/A | New |
| LOW-005: Default Log Permissions | Low | Low | Low | New |
| LOW-006: DoEvents Re-entrancy | Low | Very Low | Low | New |
| LOW-007: ComboBox Free Text | Low | Very Low (mitigated by escaping) | Very Low | New |
| INFO-001: Admin Install Required | Info | N/A | N/A | Acknowledged |
| INFO-002: No Code Signing | Info | N/A | N/A | Acknowledged |
| INFO-003: Previous Assessment Inaccuracy | Info | N/A | N/A | New |
| INFO-004: WiX UI Reference Correction | Info | N/A | N/A | New |

---

## Recommendations Summary

### Priority 1 (Should Fix)

1. **MEDIUM-007:** Add CSV file size and row count limits to prevent memory exhaustion
2. **MEDIUM-010:** Implement an idle session timeout (e.g., 30 minutes)
3. **MEDIUM-006:** Replace `-ExecutionPolicy Bypass` with `-ExecutionPolicy RemoteSigned` and sign scripts, or document the risk and ensure installation directory ACLs are restrictive

### Priority 2 (Consider Fixing)

4. **MEDIUM-008:** Apply error sanitization to log entries, or implement log file access controls
5. **MEDIUM-009:** Add domain format validation for the organization textbox
6. **LOW-005:** Set restrictive ACLs on the Logs directory or use a user-profile-specific log location

### Priority 3 (Nice to Have)

7. **LOW-004:** Remove `Sanitize-CSVValue` dead code or integrate it into future CSV export
8. **LOW-006:** Replace `DoEvents()` with targeted `Refresh()` calls
9. **INFO-002:** Consider code signing for production deployments
10. **INFO-003:** Correct previous assessment documentation

---

## Conclusion

CalendarWarlock maintains a solid security posture with all previously identified HIGH and MEDIUM vulnerabilities properly remediated. The new findings from this vulnerability scan are primarily operational and defensive-hardening concerns rather than exploitable vulnerabilities.

The most notable new findings are:
- **ExecutionPolicy Bypass** (MEDIUM-006) combined with **No Code Signing** (INFO-002) means script integrity is not verified at launch
- **No CSV size limits** (MEDIUM-007) could allow memory exhaustion via crafted files
- **No session timeout** (MEDIUM-010) leaves authenticated sessions exposed on unattended workstations

None of these findings are remotely exploitable. They all require either local access, physical presence, or user interaction (opening a crafted file). The application's use of Microsoft's authentication libraries, combined with consistent input validation and OData escaping, provides strong protection against the most common attack vectors.

**Overall Risk Level:** LOW-MEDIUM

The application is suitable for production deployment. The new findings should be addressed in a future release to further harden the security posture.

---

*Initial assessment: 2026-01-19*
*Vulnerability scan & penetration test: 2026-02-10*
*Next review recommended: Upon significant code changes or in 6 months*

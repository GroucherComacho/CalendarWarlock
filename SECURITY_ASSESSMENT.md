# CalendarWarlock Security Assessment Report

**Date:** 2026-01-19
**Last Updated:** 2026-02-13 (All Vulnerability Findings Remediated)
**Assessed By:** Security Vulnerability Scan & Penetration Testing
**Version:** 1.0.0.2

---

## Executive Summary

CalendarWarlock is a PowerShell-based Windows GUI application for managing Exchange Online calendar permissions. This assessment is a comprehensive vulnerability scan and penetration test covering all source files, modules, launcher scripts, and the installer. It builds upon previous assessments and identifies new findings.

**Overall Risk Level:** LOW

### Current Status

| Severity | Previously Found | Previously Remediated | New Findings | Newly Remediated | Total Open |
|----------|------------------|-----------------------|--------------|------------------|------------|
| Critical | 0                | 0                     | 0            | 0                | 0          |
| High     | 2                | 2                     | 0            | 0                | 0          |
| Medium   | 5                | 5                     | 5            | 5                | 0          |
| Low      | 3                | 1                     | 4            | 4                | 2          |
| Info     | 2                | 0                     | 2            | 1                | 3          |

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
**Status:** REMEDIATED (2026-02-13)

**Description:**
The batch launcher previously used `-ExecutionPolicy Bypass`, which completely disabled PowerShell's script execution policy.

**Resolution:** Changed to `-ExecutionPolicy RemoteSigned`, which allows locally-created scripts to run while requiring remote scripts to be signed.

---

### MEDIUM-007: No CSV File Size or Row Limit

**File:** `src/CalendarWarlock.ps1:1704, 1871`
**Severity:** MEDIUM
**CVSS Score:** 4.0 (CVSS:3.1/AV:L/AC:L/PR:N/UI:R/S:U/C:N/I:N/A:H)
**Category:** Denial of Service
**Status:** REMEDIATED (2026-02-13)

**Description:**
The `Grant-BulkCSVPermissions` and `Remove-BulkCSVPermissions` functions previously imported CSV files without any size or row count validation.

**Resolution:**
- Added file size validation (rejects files > 10 MB) before `Import-Csv` in both functions
- Added row count warning (prompts user if > 1,000 rows) to alert about potential API throttling

---

### MEDIUM-008: Unsanitized Error Messages in Log Files

**File:** `src/CalendarWarlock.ps1` (multiple locations)
**Severity:** MEDIUM
**CVSS Score:** 4.3 (CVSS:3.1/AV:L/AC:L/PR:L/UI:N/S:U/C:L/I:L/A:N)
**Category:** Information Disclosure
**Status:** REMEDIATED (2026-02-13)

**Description:**
The `Write-Log` function previously received raw, unsanitized exception messages.

**Resolution:**
- Applied `Sanitize-ErrorMessage` to all `Write-Log` error entries throughout the codebase
- Removed file paths from informational log messages (e.g., CSV file path references)

---

### MEDIUM-009: No Organization Domain Format Validation

**File:** `src/CalendarWarlock.ps1:2209-2219`
**Severity:** MEDIUM
**CVSS Score:** 3.7 (CVSS:3.1/AV:L/AC:L/PR:N/UI:R/S:U/C:N/I:L/A:L)
**Category:** Input Validation
**Status:** REMEDIATED (2026-02-13)

**Description:**
The organization domain textbox previously accepted arbitrary text without format validation.

**Resolution:** Added domain format validation using regex pattern `^[a-zA-Z0-9][a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$` in the Connect button click handler. Invalid domains are rejected with a descriptive error message before any connection attempt.

---

### MEDIUM-010: No Session Timeout Mechanism

**File:** `src/CalendarWarlock.ps1` (application-wide)
**Severity:** MEDIUM
**CVSS Score:** 3.7 (CVSS:3.1/AV:L/AC:H/PR:N/UI:R/S:U/C:L/I:L/A:N)
**Category:** Session Management
**Status:** REMEDIATED (2026-02-13)

**Description:**
The application previously maintained authenticated sessions indefinitely with no idle timeout.

**Resolution:**
- Implemented a 30-minute idle session timeout using a Windows Forms Timer
- The timer checks every 60 seconds if idle time has exceeded the threshold
- User activity (button clicks, text changes, dropdown selections) resets the idle timer
- On timeout, the session is automatically disconnected with a notification to the user

---

### LOW-004: Sanitize-CSVValue Function is Dead Code

**File:** `src/CalendarWarlock.ps1:214-247`
**Severity:** LOW
**CVSS Score:** 2.0
**Category:** Code Quality / Security Hygiene
**Status:** REMEDIATED (2026-02-13)

**Description:**
The `Sanitize-CSVValue` function was defined but never called anywhere in the codebase.

**Resolution:** Removed the dead `Sanitize-CSVValue` function and associated comments referencing it. The application does not export user data to CSV, so the function was unnecessary.

---

### LOW-005: Log Files Created with Default Permissions

**File:** `src/CalendarWarlock.ps1:185-188`
**Severity:** LOW
**CVSS Score:** 3.3 (CVSS:3.1/AV:L/AC:L/PR:L/UI:N/S:U/C:L/I:N/A:N)
**Category:** Information Disclosure
**Status:** REMEDIATED (2026-02-13)

**Description:**
Log files were previously created with default permissions.

**Resolution:** Added restrictive ACL configuration when creating the Logs directory. Inheritance is disabled and only the current user is granted FullControl access. Fails gracefully if ACL setting is not possible.

---

### LOW-006: DoEvents Re-entrancy Risk

**File:** `src/CalendarWarlock.ps1:209, 347, 365`
**Severity:** LOW
**CVSS Score:** 2.4
**Category:** Application Logic
**Status:** REMEDIATED (2026-02-13)

**Description:**
The application previously used `[System.Windows.Forms.Application]::DoEvents()` which could lead to re-entrancy issues.

**Resolution:** Replaced all three `DoEvents()` calls with targeted `$script:MainForm.Refresh()` calls in `Write-Log`, `Update-ProgressBar`, and `Update-ResultsLog` functions.

---

### LOW-007: ComboBox Free-Text Input Allows Arbitrary OData Query Values

**File:** `src/CalendarWarlock.ps1:2293, 2302`
**Severity:** LOW
**CVSS Score:** 2.0
**Category:** Input Validation
**Status:** DOCUMENTED AS INTENTIONAL (2026-02-13)

**Description:**
The Job Title and Department ComboBoxes use `DropDownStyle = "DropDown"` for autocomplete/type-ahead functionality. OData injection is mitigated by single-quote escaping in `AzureADOperations.psm1`.

**Resolution:** Added code comments documenting this as an intentional design choice for usability. The existing OData escaping provides sufficient protection.

---

### INFO-003: Previous Assessment Inaccuracy - Sanitize-CSVValue Usage Claim

**Severity:** INFO
**Category:** Documentation
**Status:** REMEDIATED (2026-02-13)

**Description:**
The previous security assessment (MEDIUM-001) inaccurately stated that `Sanitize-CSVValue` was in use.

**Resolution:** The `Sanitize-CSVValue` function has been removed entirely (see LOW-004). SECURITY.md documentation has been corrected to accurately reflect that CSV formula injection protection comes from email format validation, not from the removed function.

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
| MEDIUM-006: ExecutionPolicy Bypass | Medium | Low | Medium | Remediated |
| MEDIUM-007: No CSV Size Limit | Medium | Low | High (DoS) | Remediated |
| MEDIUM-008: Unsanitized Log Errors | Medium | Low | Low-Medium | Remediated |
| MEDIUM-009: No Domain Validation | Medium | Low | Low | Remediated |
| MEDIUM-010: No Session Timeout | Medium | Medium | Medium | Remediated |
| LOW-002: Sensitive Log Data | Low | Low | Low | Acknowledged |
| LOW-003: No Rate Limiting | Low | Low | Low | Acknowledged |
| LOW-004: Dead Code (Sanitize-CSVValue) | Low | N/A | N/A | Remediated |
| LOW-005: Default Log Permissions | Low | Low | Low | Remediated |
| LOW-006: DoEvents Re-entrancy | Low | Very Low | Low | Remediated |
| LOW-007: ComboBox Free Text | Low | Very Low | Very Low | Documented |
| INFO-001: Admin Install Required | Info | N/A | N/A | Acknowledged |
| INFO-002: No Code Signing | Info | N/A | N/A | Acknowledged |
| INFO-003: Previous Assessment Inaccuracy | Info | N/A | N/A | Remediated |
| INFO-004: WiX UI Reference Correction | Info | N/A | N/A | Acknowledged |

---

## Recommendations Summary

### All Previously Open Findings - REMEDIATED (2026-02-13)

All Priority 1, 2, and 3 recommendations have been addressed:

1. **MEDIUM-006:** Changed to `-ExecutionPolicy RemoteSigned`
2. **MEDIUM-007:** Added 10MB file size limit and 1000-row warning
3. **MEDIUM-008:** Applied `Sanitize-ErrorMessage` to all log entries
4. **MEDIUM-009:** Added domain format validation regex
5. **MEDIUM-010:** Implemented 30-minute idle session timeout
6. **LOW-004:** Removed dead `Sanitize-CSVValue` code
7. **LOW-005:** Set restrictive ACLs on Logs directory
8. **LOW-006:** Replaced `DoEvents()` with targeted `Refresh()` calls
9. **LOW-007:** Documented ComboBox free-text as intentional design
10. **INFO-003:** Corrected previous assessment documentation

### Remaining Acknowledged Items (No Action Required)

- **LOW-002:** Sensitive log data (by design, mitigated by ACLs)
- **LOW-003:** No rate limiting (relies on M365 throttling)
- **INFO-001:** Per-machine installation (appropriate for admin tool)
- **INFO-002:** No code signing (consider for future production deployments)

---

## Conclusion

CalendarWarlock maintains a strong security posture with all identified HIGH, MEDIUM, and actionable LOW vulnerabilities remediated. As of 2026-02-13, all findings from both the initial assessment and the 2026-02-10 vulnerability scan have been addressed.

Key remediation highlights:
- **ExecutionPolicy** changed from Bypass to RemoteSigned (MEDIUM-006)
- **CSV file size and row limits** prevent memory exhaustion and API throttling (MEDIUM-007)
- **Log sanitization** applied consistently to both UI and log files (MEDIUM-008)
- **Domain format validation** added for organization input (MEDIUM-009)
- **30-minute idle timeout** protects against unattended session abuse (MEDIUM-010)
- **Dead code removed**, **DoEvents re-entrancy fixed**, **log ACLs set** (LOW-004/005/006)

The application's use of Microsoft's authentication libraries, combined with consistent input validation, OData escaping, and the new defensive hardening measures, provides strong protection against common attack vectors.

**Overall Risk Level:** LOW

The application is suitable for production deployment.

---

*Initial assessment: 2026-01-19*
*Vulnerability scan & penetration test: 2026-02-10*
*All findings remediated: 2026-02-13*
*Next review recommended: Upon significant code changes or in 6 months*

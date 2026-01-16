# CalendarWarlock - Detailed Usage Guide

## Overview

CalendarWarlock is designed to simplify bulk calendar permission management in Exchange Online. This guide covers detailed usage scenarios and best practices.

## Common Use Cases

### Use Case 1: Executive Assistant Access

**Scenario**: An executive assistant needs to manage calendars for all Directors in the company.

**Steps**:
1. Connect to Exchange Online with your organization domain
2. Select "Director" from the Job Title dropdown
3. Enter the executive assistant's email address
4. Select "Editor" permission level (allows full calendar management)
5. Click "Grant User Access to All Calendars of Job Title"

**Result**: The executive assistant can now view and edit all Director calendars.

### Use Case 2: Team Calendar Visibility

**Scenario**: All members of the Sales team need to see their manager's calendar.

**Steps**:
1. Connect to Exchange Online
2. Select "Sales Representative" from the Job Title dropdown
3. Enter the sales manager's email address
4. Select "Reviewer" permission level (read-only access)
5. Click "Grant All of Job Title Access to User's Calendar"

**Result**: All Sales Representatives can now view their manager's calendar.

### Use Case 3: Department-Wide Free/Busy Access

**Scenario**: All engineers should be able to see when other engineers are available.

**Steps**:
1. Connect to Exchange Online
2. For each engineer who wants their calendar shared:
   - Enter their email as the calendar owner
   - Select "Software Engineer" as the job title
   - Select "AvailabilityOnly" permission level
   - Click "Grant All of Job Title Access to User's Calendar"

**Result**: Engineers can see free/busy status for scheduling meetings.

## Understanding Permission Levels

### Full Access Permissions

| Permission | Create | Read | Edit | Delete | Subfolders |
|------------|--------|------|------|--------|------------|
| Owner | All | All | All | All | Yes + Permissions |
| PublishingEditor | All | All | All | All | Yes |
| Editor | All | All | All | All | No |

### Limited Access Permissions

| Permission | Create | Read | Edit Own | Delete Own | Subfolders |
|------------|--------|------|----------|------------|------------|
| PublishingAuthor | Yes | All | Yes | Yes | Yes |
| Author | Yes | All | Yes | Yes | No |
| NonEditingAuthor | Yes | All | No | Yes | No |

### Read-Only Permissions

| Permission | Description |
|------------|-------------|
| Reviewer | Can read all calendar items with full details |
| AvailabilityOnly | Can only see free/busy status (no details) |
| LimitedDetails | Can see free/busy plus subject and location |

### Special Permissions

| Permission | Description |
|------------|-------------|
| Contributor | Can create items but cannot read any items |
| None | Removes all access |

## Best Practices

### Security

1. **Use Least Privilege**: Grant only the minimum permission level needed
   - Use "Reviewer" instead of "Editor" if users only need to view
   - Use "AvailabilityOnly" if users only need to check availability

2. **Regular Audits**: Review the logs periodically to ensure permissions are appropriate

3. **Confirmation**: Always verify the job title and user before executing bulk operations

### Operational

1. **Test First**: Before running bulk operations on large groups, test with a small subset

2. **Off-Hours Execution**: Run large bulk operations during off-peak hours

3. **Document Changes**: Keep records of what permissions were granted and why

4. **Review Job Titles**: Ensure job titles in Azure AD are accurate and standardized

## Logging

CalendarWarlock creates detailed logs in the `src/Logs` folder:

```
Logs/
└── CalendarWarlock_20240115_143022.log
```

Log entries include:
- Timestamp
- Operation type (INFO, SUCCESS, ERROR, WARNING)
- Detailed message

Example log entries:
```
[2024-01-15 14:30:22] [INFO] CalendarWarlock started
[2024-01-15 14:30:45] [INFO] Initiating connection to contoso.onmicrosoft.com
[2024-01-15 14:30:52] [SUCCESS] Successfully connected to contoso.onmicrosoft.com
[2024-01-15 14:31:15] [INFO] Grant-BulkPermissionsToUser: User=assistant@contoso.com, JobTitle=Director, Permission=Editor
[2024-01-15 14:31:45] [SUCCESS] Bulk grant completed: Success=12, Failed=0, Skipped=1
```

## Error Handling

### Common Errors

**"User not found"**
- Verify the email address is correct
- Check if the user exists in Azure AD
- Ensure the user has a mailbox

**"Access denied"**
- Verify you have Exchange Admin or Recipient Management role
- Check your Microsoft Graph permissions

**"Calendar not found"**
- The user may not have a calendar folder
- The mailbox might be disabled or inactive

### Recovery

If an operation partially fails:
1. Review the results log to see which users succeeded/failed
2. Address the failing users individually
3. Re-run the operation for remaining users if needed

## Advanced Topics

### Removing Bulk Permissions

Currently, CalendarWarlock focuses on granting permissions. To remove permissions in bulk, you can use PowerShell directly:

```powershell
# Get users by job title
$users = Get-MgUser -Filter "jobTitle eq 'Director'" -All

# Remove permissions
foreach ($user in $users) {
    Remove-MailboxFolderPermission -Identity "$($user.Mail):\Calendar" -User "assistant@contoso.com" -Confirm:$false
}
```

### Viewing Current Permissions

To view existing permissions on a calendar:

```powershell
Get-MailboxFolderPermission -Identity "user@contoso.com:\Calendar"
```

### Handling Localized Calendar Folder Names

In non-English environments, the calendar folder might have a different name. CalendarWarlock uses the default "Calendar" folder name. For localized environments, you may need to modify the `ExchangeOperations.psm1` module to detect the correct folder name:

```powershell
# Get the actual calendar folder name
$calendarFolder = Get-MailboxFolderStatistics -Identity $user -FolderScope Calendar | Select-Object -First 1
$calendarIdentity = "$user`:$($calendarFolder.Name)"
```

## Support

For issues or feature requests, please:
1. Check the troubleshooting section in README.md
2. Review the logs for detailed error messages
3. Submit an issue on the project repository

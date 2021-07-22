---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphUser

## SYNOPSIS
Gets information from the MS-Graph API about the a user (current user by default)

## SYNTAX

### None (Default)
```
Get-GraphUser [[-UserID] <Object>] [-Current] [<CommonParameters>]
```

### Calendars
```
Get-GraphUser [[-UserID] <Object>] [-Calendars] [-Current] [<CommonParameters>]
```

### DirectReports
```
Get-GraphUser [[-UserID] <Object>] [-DirectReports] [-Current] [<CommonParameters>]
```

### Drive
```
Get-GraphUser [[-UserID] <Object>] [-Drive] [-Current] [<CommonParameters>]
```

### LicenseDetails
```
Get-GraphUser [[-UserID] <Object>] [-LicenseDetails] [-Current] [<CommonParameters>]
```

### MailboxSettings
```
Get-GraphUser [[-UserID] <Object>] [-MailboxSettings] [-Current] [<CommonParameters>]
```

### OutlookCategories
```
Get-GraphUser [[-UserID] <Object>] [-OutlookCategories] [-Current] [<CommonParameters>]
```

### Manager
```
Get-GraphUser [[-UserID] <Object>] [-Manager] [-Current] [<CommonParameters>]
```

### Teams
```
Get-GraphUser [[-UserID] <Object>] [-Teams] [-Current] [<CommonParameters>]
```

### Groups
```
Get-GraphUser [[-UserID] <Object>] [-Groups] [-SecurityGroups] [-Current] [<CommonParameters>]
```

### SecurityGroups
```
Get-GraphUser [[-UserID] <Object>] [-SecurityGroups] [-Current] [<CommonParameters>]
```

### MemberOf
```
Get-GraphUser [[-UserID] <Object>] [-MemberOf] [-Current] [<CommonParameters>]
```

### TransitiveMemberOf
```
Get-GraphUser [[-UserID] <Object>] [-TransitiveMemberOf] [-Current] [<CommonParameters>]
```

### Notebooks
```
Get-GraphUser [[-UserID] <Object>] [-Notebooks] [-Current] [<CommonParameters>]
```

### Photo
```
Get-GraphUser [[-UserID] <Object>] [-Photo] [-Current] [<CommonParameters>]
```

### PlannerTasks
```
Get-GraphUser [[-UserID] <Object>] [-PlannerTasks] [-Current] [<CommonParameters>]
```

### PlannerPlans
```
Get-GraphUser [[-UserID] <Object>] [-Plans] [-Current] [<CommonParameters>]
```

### Presence
```
Get-GraphUser [[-UserID] <Object>] [-Presence] [-Current] [<CommonParameters>]
```

### Site
```
Get-GraphUser [[-UserID] <Object>] [-Site] [-Current] [<CommonParameters>]
```

### ToDoLists
```
Get-GraphUser [[-UserID] <Object>] [-ToDoLists] [-Current] [<CommonParameters>]
```

### Select
```
Get-GraphUser [[-UserID] <Object>] -Select <String[]> [-Current] [<CommonParameters>]
```

## DESCRIPTION
Queries https://graph.microsoft.com/v1.0/me or https://graph.microsoft.com/v1.0/name@domain
or https://graph.microsoft.com/v1.0/\<\<guid\>\> for information about a user.
Getting a user returns a default set of properties only (businessPhones, displayName, givenName,
id, jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName).
Use -select to get the other properties.
Most options need consent to use the Directory.Read.All or Directory.AccessAsUser.All scopes.
Some options will also work with user.read; and the following need consent which is task specific
Calendars needs Calendars.Read, OutLookCategries needs MailboxSettings.Read, PlannerTasks needs
Group.Read.All, Drive needs Files.Read (or better), Notebooks needs either Notes.Create or
Notes.Read (or better).

## EXAMPLES

### EXAMPLE 1
```
Get-GraphUser -MemberOf | ft displayname, description, mail, id
Shows the name description, email address and internal ID for the groups this user is a direct member of
```

### EXAMPLE 2
```
(get-graphuser -Drive).root.children.name
Gets the user's one drive. The drive object has a .root property which is represents its
root-directory, and this has a .children property which is a collection of the objects
in the root directory. So this command shows the names of files and folders in the root directory. To just see sub folders it is possible to use
get-graphuser -Drive | Get-GraphDrive -subfolders
```

## PARAMETERS

### -UserID
UserID as a guid or User Principal name.
If not specified, it will assume "Current user" if other paraneters are given, or "All users" otherwise.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: id

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -Calendars
Get the user's Calendar(s)

```yaml
Type: SwitchParameter
Parameter Sets: Calendars
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -DirectReports
Select people who have the user as their manager

```yaml
Type: SwitchParameter
Parameter Sets: DirectReports
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Drive
Get the user's one drive

```yaml
Type: SwitchParameter
Parameter Sets: Drive
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -LicenseDetails
Get user's license Details

```yaml
Type: SwitchParameter
Parameter Sets: LicenseDetails
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -MailboxSettings
Get the user's Mailbox Settings

```yaml
Type: SwitchParameter
Parameter Sets: MailboxSettings
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutlookCategories
Get the users Outlook-categories (by default, 6 color names)

```yaml
Type: SwitchParameter
Parameter Sets: OutlookCategories
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Manager
Get the user's manager

```yaml
Type: SwitchParameter
Parameter Sets: Manager
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Teams
Get the user's teams

```yaml
Type: SwitchParameter
Parameter Sets: Teams
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Groups
Get the user's Groups

```yaml
Type: SwitchParameter
Parameter Sets: Groups
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -SecurityGroups
{{ Fill SecurityGroups Description }}

```yaml
Type: SwitchParameter
Parameter Sets: Groups
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

```yaml
Type: SwitchParameter
Parameter Sets: SecurityGroups
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -MemberOf
Get the Directory-Roles and Groups the user belongs to; -Groups or -Teams only return one type of object.

```yaml
Type: SwitchParameter
Parameter Sets: MemberOf
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -TransitiveMemberOf
Get the Directory-Roles and Groups the user belongs to; -Groups or -Teams only return one type of object.

```yaml
Type: SwitchParameter
Parameter Sets: TransitiveMemberOf
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Notebooks
Get the user's Notebook(s)

```yaml
Type: SwitchParameter
Parameter Sets: Notebooks
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Photo
Get the user's photo

```yaml
Type: SwitchParameter
Parameter Sets: Photo
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -PlannerTasks
Get the user's assigned tasks in planner.

```yaml
Type: SwitchParameter
Parameter Sets: PlannerTasks
Aliases: AssignedTasks

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Plans
Get the plans owned by the user in planner.

```yaml
Type: SwitchParameter
Parameter Sets: PlannerPlans
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presence
Get the users presence in Teams

```yaml
Type: SwitchParameter
Parameter Sets: Presence
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Site
Get the user's MySite in SharePoint

```yaml
Type: SwitchParameter
Parameter Sets: Site
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ToDoLists
Get the user's To-do lists

```yaml
Type: SwitchParameter
Parameter Sets: ToDoLists
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
specifies which properties of the user object should be returned Additional options are available when selecting individual users
The API documents list deviceEnrollmentLimit, deviceManagementTroubleshootingEvents , mailboxSettings which cause errors

```yaml
Type: String[]
Parameter Sets: Select
Aliases:

Required: True
Position: Named
Default value: $Script:DefaultUserProperties
Accept pipeline input: False
Accept wildcard characters: False
```

### -Current
Used to explicitly say "Current user" and will over-ride UserID if one is given.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser
## NOTES

## RELATED LINKS

---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/datetimecolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphGroup

## SYNOPSIS
Adds a new group/team

## SYNTAX

### None (Default)
```
New-GraphGroup [-Name] <String> [-Description <String>] [-MailNickName <String>] [-Visibility <String>]
 [-Members <Object>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Security
```
New-GraphGroup [-Name] <String> [-AsSecurity] [-AsAssignableToRole] [-Description <String>]
 [-MailNickName <String>] [-Visibility <String>] [-Members <Object>] [-Owners <Object>] [-Force] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

### Team
```
New-GraphGroup [-Name] <String> [-AsTeam] [-Description <String>] [-MailNickName <String>]
 [-Visibility <String>] [-Members <Object>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Owners
```
New-GraphGroup [-Name] <String> [-Description <String>] [-MailNickName <String>] [-Visibility <String>]
 [-Members <Object>] [-Owners <Object>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Every team is also a group, but not every group is team enabled.
This Command has an alias of New-GraphTeam so you call it as team or group
By default it creates the group as a team UNLESS you specify -NoTeam.
A non-Teams enabled group can be teams enabled with Set-GraphGroup -EnableTeam
Creating and modifying groups requires consent to use the Group.ReadWrite.All scope

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Name
The name of the Group / Team

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsSecurity
Unless specified, groups will be mail enabled "unfied" / Microsoft365 groups
The Graph API doesn't allow mail-enabled & security-enabled,  or mail-disabled & unified
Only unified groups can be made into teams.
Unified groups can only contain users,
Security groups can contain other security principals

```yaml
Type: SwitchParameter
Parameter Sets: Security
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsAssignableToRole
If specified allows Azure AD roles can be assigned to the group.
This forces visibility to be private, and can't be changed.

```yaml
Type: SwitchParameter
Parameter Sets: Security
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AsTeam
New-GraphGroup only enables teams functonality if -AsTeam is specified.
Calling as New-GraphTeam defaults AsTeam to true

```yaml
Type: SwitchParameter
Parameter Sets: Team
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Description
A description for the group

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MailNickName
The group/team's mail nickname

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Visibility
The visibility of the group, Public by default, it can be 'private' or 'hidden membership'

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Public
Accept pipeline input: False
Accept wildcard characters: False
```

### -Members
Ordinary Members of the group - assumed to be users, given by their User Principal Name or ID or as objects

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Owners
Owners of the group - assumed to be users, given by their User Principal Name or ID or as objects

```yaml
Type: Object
Parameter Sets: Security, Owners
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
if specified group will be added without prompting

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

### -WhatIf
Shows what would happen if the cmdlet runs.
The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm
Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphGroup
## NOTES

## RELATED LINKS

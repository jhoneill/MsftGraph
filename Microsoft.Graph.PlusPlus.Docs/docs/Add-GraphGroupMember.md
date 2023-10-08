---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphGroupMember

## SYNOPSIS
Adds a user (or group) to a group/team as either a member or owner.

## SYNTAX

```
Add-GraphGroupMember [-Group] <Object> [-Member] <Object> [-AsOwner] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Because the group may be a team the this command has alias of Add-GraphTeamMember.
it requires consent to use the Group.ReadWrite.All, Directory.ReadWrite.All, or
Directory.AccessAsUser.All scope.

## EXAMPLES

### EXAMPLE 1
```
$newGroup = New-GraphGroup -Name Test101
>Get-GraphUserList -Filter "Department eq 'Accounts'" | Add-GraphGroupMember -Group $newGroup
Creates a new group; then gets a list of users and adds them to the group.
```

### EXAMPLE 2
```
Add-GraphTeamMember -Team $Newteam -Member alex@contoso.com -AsOwner
Adds an owner to a team, using aliases for both the command and the group parameter
```

## PARAMETERS

### -Group
The group / team either as an ID or a group/team object with an IDn

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Member
The user or nested-group to add, either as a UPN or ID or as a object with an ID

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -AsOwner
If specified the user will be added as an owner, otherwise they will be a standard member

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

### -Force
If specified group member will be added without prompting

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

## NOTES

## RELATED LINKS

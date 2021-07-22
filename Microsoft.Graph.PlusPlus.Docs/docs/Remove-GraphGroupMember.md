---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Remove-GraphGroupMember

## SYNOPSIS
Removes a user (or group) from a group/team

## SYNTAX

```
Remove-GraphGroupMember [-Group] <Object> [-Member] <Object> [-FromOwners] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Because the group may be a team the command has an alias of Remove-GraphTeamMember.
It requires consent to use the Group.ReadWrite.All, Directory.ReadWrite.All, or
Directory.AccessAsUser.All scope.

## EXAMPLES

### EXAMPLE 1
```
Remove-GraphGroupMember -Group $g -FromOwners -Member alex@contoso.com -Force
Removes a user from the owners of a group without prompting for confirmation.
```

### EXAMPLE 2
```
Get-GraphUserList -Filter "Department eq 'Accounts'" | Remove-GraphGroupMember -Group $g
Gets a list of users and removes them from from a group.
```

## PARAMETERS

### -Group
The ID of the Group / team

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
A group object with an ID field, or a user object, user ID or UPN

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

### -FromOwners
If specified the member will be removed from the owners rather than members

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
If specified the member will be removed without prompting for confirmation

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

---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphDefaultGroup

## SYNOPSIS
Sets the default paramater for group or team in most functions which take one.

## SYNTAX

```
Set-GraphDefaultGroup [-Group] <Object> [<CommonParameters>]
```

## DESCRIPTION
Takes a group as a parameter or via the pipeline.
If a string is passed it will try to get a matching group from Get-GraphGroup,
a string may be a wildcard for a group name - provided that it only finds one matching group.
If the group has been provisioned as a team then it will be the default for commands which take a -Team parameter.
The primary purpose is to avoid specifying a Group/Team when working with messages, calendar / planner / team channels,
but working with the group itself or its membership it is safer not to default the selection, so no defaults
are set for for Set-Team, Set-Group, Get- Remove-Group Remove-GroupMember or Add-GroupMember or Import and Export

## EXAMPLES

### EXAMPLE 1
```
Set-GraphDefaultGroup Accounts
>  Get-GraphChannel
Display Name description
----------- -----------
General      The Accounts Department
Mccaw        For anything about project Mccaw
```

The first command sets the default group - because "Accounts" has been provisioned as a team,
it becomes the default team for Get-GraphChannel

## PARAMETERS

### -Group
The group to set as the default for other commands

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/calculatedcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphChannel

## SYNOPSIS
Adds a channel to a team

## SYNTAX

```
New-GraphChannel [-Team] <Object> [-Name] <String[]> [-Description <String>] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
This requires the Group.ReadWrite.All scope.

## EXAMPLES

### EXAMPLE 1
```
$newChannel  = New-GraphChannel -Team $newTeam -Name $newProjectName -Description "For anything about project $newProjectName"
$newTeam holds the result of creating a team with New-GraphTeam...
$newProjectName holds the name of a project the team will be working on.
This command creates a new channel in Teams, and stores the result in a variable
which can then be used to post messages to the channel, or add tabs to it.
```

## PARAMETERS

### -Team
The team where the channel will be added, either as an ID or a team object

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Display name for the new channel

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: DisplayName

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Description
Description for the new channel

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

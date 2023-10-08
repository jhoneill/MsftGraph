---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphPlannerTab

## SYNOPSIS
Adds a planner tab to a team-channel for a pre-existing plan

## SYNTAX

```
Add-GraphPlannerTab [-Plan] <Object> [-Channel] <Object> [-Team <Object>] [-TabLabel <Object>]
 [-Force <Object>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
which requires consent to use the Group.ReadWrite.All scope.

## EXAMPLES

### EXAMPLE 1
```
>$channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
>$plan   = Get-GraphTeam -ByName accounts  -Plans | where title -Like "year end*"
>Add-GraphPlannerTab -Plan $plan -Channel $channel -TabLabel "Planner"
The first line gets the 'year-end' channel for the accounts team
The second gets a plan with tile which matches 'year end'
and the third creates a tab labelled 'Planner' in the channel for that plan.
```

## PARAMETERS

### -Plan
An ID or Plan object for a plan within the team

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

### -Channel
An ID or Channel object for a channel (which may contain the team ID)

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Team
A team ID, or a team object, if not specified as part of the channel

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

### -TabLabel
The label for the tab.

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

### -Force
If Specified the tab will be added without confirming

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

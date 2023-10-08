---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphSharePointTab

## SYNOPSIS
Adds a planner tab to a team-channel for sharepoint deurl

## SYNTAX

```
Add-GraphSharePointTab [-WebUrl] <Object> [-TabLabel] <Object> [[-Template] <Object>] [-Channel] <Object>
 [-Team <Object>] [-Force <Object>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
which requires consent to use the Group.ReadWrite.All scope.

## EXAMPLES

### EXAMPLE 1
```

```

## PARAMETERS

### -WebUrl
An ID or Plan object for a plan within the team

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -TabLabel
The label for the tab by default the displayname for of the list

```yaml
Type: Object
Parameter Sets: (All)
Aliases: DisplayName

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Template
The label for the tab.
Either a genericList (default) or a documentLibrary

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: GenericList
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Channel
An ID or Channel object for a channel (which may contain the team ID)

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 4
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

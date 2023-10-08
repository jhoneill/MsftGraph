---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphOneNoteTab

## SYNOPSIS
Adds a tab in a Teams channel for a OneNote section or Notebook

## SYNTAX

```
Add-GraphOneNoteTab [-Notebook] <Object> [-Channel] <Object> [-Team <Object>] [-TabLabel <Object>]
 [-Force <Object>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
which requires consent to use the Group.ReadWrite.All scope.
The Notebook Parameter has an alias of 'Section' and will accept either
a OneNote Notebook object (or its 'Self' URI - which requires the tab name to be
set explicitly) or a Section object.
If the notebook is specified it opens at the
first section.

## EXAMPLES

### EXAMPLE 1
```
> $section = Get-GraphTeam -ByName accounts -Notebooks | Select-Object -ExpandProperty sections  | where displayname -like "FY-19*"
> $channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
> Add-GraphOneNoteTab  $section $channel -TabLabel "FY-19 Notes"
```

The first command gets the Notebook for the Accounts team and finds the "FY-19 Year End" section
The second command gets the channels for the same team and finds the "Year end" channel
The Third command creates a tab in the channel named 'FY-19 Notes' which opens the team notebook
at its 'FY-19 Year End' section.

## PARAMETERS

### -Notebook
The Notebook or Section to associate with the tab

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Section

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Channel
An ID or Channel object which may contain the team ID; the tab will be created in this channel

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
A team ID, or a team object if the team can't be found from the the channel

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
The label for the tab, if left blank the name of the Notebook or Section will be sued

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
If Specified the tab will be added without pausing for confirmation, this is the default unless $ConfirmPreference has been set.

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

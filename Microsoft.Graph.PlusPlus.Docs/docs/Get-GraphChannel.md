---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphChannel

## SYNOPSIS
Gets details of a channel, or its Tabs or messages shown in Teams

## SYNTAX

### None (Default)
```
Get-GraphChannel [-Team <Object>] [<CommonParameters>]
```

### CHFiles
```
Get-GraphChannel [-Team <Object>] [-Channel] <Object> [-Files] [<CommonParameters>]
```

### CHFolder
```
Get-GraphChannel [-Team <Object>] [-Channel] <Object> [-Folder] [<CommonParameters>]
```

### CHTabs
```
Get-GraphChannel [-Team <Object>] [-Channel] <Object> [-Tabs] [<CommonParameters>]
```

### CHMsgs
```
Get-GraphChannel [-Team <Object>] [-Channel] <Object> [-Messages] [-Top <Object>] [<CommonParameters>]
```

### NoCHTabs
```
Get-GraphChannel [-Team <Object>] [-Tabs] [<CommonParameters>]
```

### NoCHFolder
```
Get-GraphChannel [-Team <Object>] [-Folder] [<CommonParameters>]
```

### NoCHFiles
```
Get-GraphChannel [-Team <Object>] [-Files] [<CommonParameters>]
```

### NoCHMsgs
```
Get-GraphChannel [-Team <Object>] [-Messages] [-Top <Object>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphTeam  consultants -ChannelName general | Get-GraphChannel -Tabs
Gets channels for the team(s) with a name beginning 'Consultants' and selects channel(s)
with a name beginning "general"; then gets the tabs shown in Teams for this channel
```

### EXAMPLE 2
```
Set-GraphDefaultGroup 'Consultants'
> ...
> Get-GraphChannel 'General' -Messages
If the default group is set to a suitable team, it is possible to tab complete the channel name
and ther is no need specify the team
```

### EXAMPLE 3
```
Get-GraphChannel -Team accounts -channel general -Messages
This specifies a non-default team, and gets messages from the teams general channel.
```

### EXAMPLE 4
```
Get-GraphChannel -Team $t
Gets the basic channel information for team.
```

## PARAMETERS

### -Team
The ID of the team if it is not in the channel object.
If not specified the current users teams are tried

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

### -Channel
The channel either as a name, an ID or as a channel object (which may contain the team as a property)

```yaml
Type: Object
Parameter Sets: CHFiles, CHFolder, CHTabs, CHMsgs
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Tabs
If specified gets the channel's Tabs

```yaml
Type: SwitchParameter
Parameter Sets: CHTabs, NoCHTabs
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Folder
If specified gets the channel's Tabs

```yaml
Type: SwitchParameter
Parameter Sets: CHFolder, NoCHFolder
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Files
{{ Fill Files Description }}

```yaml
Type: SwitchParameter
Parameter Sets: CHFiles, NoCHFiles
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Messages
if Specified uses the beta api to get the channel's messages.

```yaml
Type: SwitchParameter
Parameter Sets: CHMsgs, NoCHMsgs
Aliases: Msgs

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
If specified, returns the top n messages, otherwise the command will attempt to get all messages.
The server may return more than the specified number.

```yaml
Type: Object
Parameter Sets: CHMsgs, NoCHMsgs
Aliases:

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

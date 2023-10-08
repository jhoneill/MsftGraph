---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphTeam

## SYNOPSIS
Updates the settings for a team

## SYNTAX

```
Set-GraphTeam [[-Team] <Object>] [-AllowMemberAddRemoveApps] [-AllowMemberCreateUpdateRemoveConnectors]
 [-AllowMemberCreateUpdateRemoveTabs] [-AllowMemberCreateUpdateChannels] [-AllowMemberDeleteChannels]
 [-AllowGuestCreateUpdateChannels] [-AllowGuestDeleteChannels] [-AllowUserEditMessages]
 [-AllowUserDeleteMessages] [-AllowOwnerDeleteMessages] [-AllowTeamMentions] [-AllowChannelMentions]
 [-AllowGiphy] [-GiphyContentRating <String>] [-AllowStickersAndMemes] [-AllowCustomMemes] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use the  Group.ReadWrite.All scope

## EXAMPLES

### EXAMPLE 1
```
Get-GraphTeam accounts* | Set-GraphTeam -AllowGiphy:$false
Gets a the team(s) with a name that begins with accounts, and turns off Giphy content
Note the use of -SwitchName:$false.
```

## PARAMETERS

### -Team
The team to update either as an ID or a team object with and ID.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -AllowMemberAddRemoveApps
Allow members to add or remove apps

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

### -AllowMemberCreateUpdateRemoveConnectors
Allow members to create update or remove connectors

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

### -AllowMemberCreateUpdateRemoveTabs
Allow members to create update or remove Tabs

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

### -AllowMemberCreateUpdateChannels
Allow members to create or update Channels

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

### -AllowMemberDeleteChannels
Allow members to delete Channels

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

### -AllowGuestCreateUpdateChannels
Allow guests to create or update Channels

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

### -AllowGuestDeleteChannels
Allow guests to delete Channels

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

### -AllowUserEditMessages
Allow members to edit their own messages

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

### -AllowUserDeleteMessages
Allow members to delete their own messages

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

### -AllowOwnerDeleteMessages
Allow owners to delete mssages

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

### -AllowTeamMentions
Allow mentions of teams in messages

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

### -AllowChannelMentions
Allow mentions of channels in messages

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

### -AllowGiphy
Allow giphy graphics

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

### -GiphyContentRating
Rating for giphy graphics; either moderate or strict

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

### -AllowStickersAndMemes
Allow stickers and memes

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

### -AllowCustomMemes
Allow Custom memes

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

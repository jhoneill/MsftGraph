---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphChannelReply

## SYNOPSIS
Returns replies to messages in Teams channels

## SYNTAX

```
Get-GraphChannelReply [-Message] <Object> [-Team <Object>] [-Channel <Object>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Access to channel messages is currently in the BETA API
It is possible to start a new thread, but not to reply to the thread.

## EXAMPLES

### EXAMPLE 1
```
Get-GraphChannel $General -Messages | Get-GraphChannelReply -PassThru
The GraphAPI does not return replies when requesting messages
from a channel in Teams. By piping the messages to Get-GraphChannelReply
it is possible to get the replies; and if -Passthru is specified
the messages will returned, followed by their replies.
So if $General is a channel object, the first message and the its first
reply might be output like this.
```

From          Created          Isreply Deleted Importance Content
----          -------          ------- ------- ---------- -------
James O'Neill 17/02/2019 11:42 False   False   normal     Project Firebird now has its own channel.
James O'Neill 17/02/2019 13:06 True    False   normal     And the channel has its own planner

## PARAMETERS

### -Message
The Message to reply to as an ID or a message object containing an ID (and possibly the team and channel ID)

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Team
If the message or channel parameters don't include the team ID, the team either as an ID or an object containing the ID

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
If Message does not contain the channel, the channel either as an ID or an object containing an ID and possibly the team ID

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

### -PassThru
If specified returns the message, followed by its replies.
(Otherwise , only the replies are returned)

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphGroupConversation

## SYNOPSIS
Gets details of group converstation from outlook, or its threads.

## SYNTAX

### OneConversation
```
Get-GraphGroupConversation [-Conversation] <Object> [[-Group] <Object>] [-Threads] [<CommonParameters>]
```

### InTeam
```
Get-GraphGroupConversation [[-Group] <Object>] [[-Topic] <Object>] [-Threads] [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use the Group.Read.All scope

## EXAMPLES

### EXAMPLE 1
```
Get-GraphGroupList -Name consult | Get-GraphGroup -Conversations | Get-GraphGroupConversation -Threads
Gets group(s) matching the name "consult*" , finds their conversations and for each one gets the threads in the conversation
Note, unless you are dealing with conversations which have multiple threads, it is easier to do Get-GraphGroup -Threads
```

## PARAMETERS

### -Conversation
The Conversation, either as an ID or an object.

```yaml
Type: Object
Parameter Sets: OneConversation
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Group
The group where the conversation is found,it is not part of can't be found from the conversation object

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Topic
When selecting the Conversations for a group narrows the list by the name of the topic

```yaml
Type: Object
Parameter Sets: InTeam
Aliases:

Required: False
Position: 4
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -Threads
If specified selects the conversation's threads, otherwise an object representing the conversation itself is returned.

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

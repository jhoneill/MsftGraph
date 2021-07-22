---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphGroupThread

## SYNOPSIS
Gets a thread in a Group conversation in outlook, or its posts

## SYNTAX

### SingleThread
```
Get-GraphGroupThread [-Thread] <Object> [[-Group] <Object>] [-Posts] [<CommonParameters>]
```

### GroupThreads
```
Get-GraphGroupThread [[-Group] <Object>] [-Topic <Object>] [-Posts] [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use the Group.Read.All scope

## EXAMPLES

### EXAMPLE 1
```
Get-GraphUser -Teams  | Get-GraphGroup -Threads | Get-GraphGroupThread -Posts |
     ft -a -Wrap  @{n="from";e={$_.from.emailaddress.name}},CreatedDateTime,Topic,@{n="Body";e={$_.body.content}}
Gets a users teams, for each one gets their threads, and for each thread gets the outlook posts
Displays the result as a table showing from, message date, thread topic and message body
Note this uses Get-GraphGroup as an alias for Get-GraphTeams
```

## PARAMETERS

### -Thread
The group thread, either as an ID or as a thread object (which may have the team/group as property)

```yaml
Type: Object
Parameter Sets: SingleThread
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Group
The group holding the thread (s), if thread is either not passed or is just the ID of a thread.

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
When selecting the threads for a group narrows the list by the name of the topic

```yaml
Type: Object
Parameter Sets: GroupThreads
Aliases:

Required: False
Position: Named
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -Posts
If specified, returns the posts in the thread

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

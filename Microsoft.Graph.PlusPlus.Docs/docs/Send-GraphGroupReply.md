---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Send-GraphGroupReply

## SYNOPSIS
Replies to a group's post in outlook.

## SYNTAX

```
Send-GraphGroupReply [-Post] <Object> [-Thread <Object>] [-Group <Object>] -Content <String>
 [-ContentType <String>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
$thread.posts[0] | Send-GraphGroupReply -content '<b><font color="green">Success!!</font> Go team!</b>' -ContentType HTML
One of the examples for Add-GraphGroupThread left the result of a creating a new thread in $thread
This takes the only post in the new thread and creates a reply to it with the content in HTML format.
```

### EXAMPLE 2
```
Set-GraphDefaultGroup 'Consultants'
> ...
> $post = Get-GraphGroupThread -Topic  "Today's tests..."  -Posts | select -last 1
>Send-GraphGroupReply $post -Content "Please join a celebration of the successful test at 4PM"
This example finds threads for the consultants group, Isolates the one with the topic of
"Today's Tests..." and finds the last post in the thread. It then posts a reply with the content as plain text.
This example stores the Post between the two commands but they could be piped together as in the previous example
```

## PARAMETERS

### -Post
The Post being replied to, either as an ID or a post object containing an ID which may identify the thread and group

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

### -Thread
The thread containing the post (if not embedded in the post itself), as an ID or object, which may identify the group

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

### -Group
The group containing the thread (if not embedded in the Post or thread) as an ID or object

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
The Message body - text by default, specify -contentType if using HTML

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentType
The type of content, text by default or HTML

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Text
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
if Specified the message will be created without prompting.

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

[Add-GraphGroupThread]()


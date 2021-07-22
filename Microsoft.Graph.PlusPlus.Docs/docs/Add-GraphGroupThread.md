---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphGroupThread

## SYNOPSIS
Starts a new thread in a group in outlook.

## SYNTAX

```
Add-GraphGroupThread [-Group] <Object> [-ThreadTopic] <Object> [-Content] <String> [-ContentType <String>]
 [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use the Group.ReadWrite.All scope

## EXAMPLES

### EXAMPLE 1
```
>$G = Get-GraphGroup  consultants
>Add-GraphGroupThread -Group $G -Subject "Running tests.." -Content "We will be running a full test pass this afternoon"
Gets a group by name and creates a new thread with a message using a plain text body.
```

### EXAMPLE 2
```
$thread = Add-GraphGroupThread -passthru -Group $G -Subject "Ruuning tests.." -ContentType HTML -Content "<b><i>Drum-Roll...</i>A full test pass is running... Watch this space</i>"
Uses the group from the previous example, and creates a thread with an HTML body, and keeps a reference to it.
```

## PARAMETERS

### -Group
The group where the thread will be added

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThreadTopic
The subject line for the thread

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Subject

Required: True
Position: 2
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
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentType
The content type, (Text by default) or HTML

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
if Specified the message will be created without prompting; this is the default, unless $confirm preference has been changed

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

### -PassThru
if Specified an object containing the Thread ID will be returned

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

[Send-GraphGroupReply]()


---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-gb/graph/onenote-update-page
schema: 2.0.0
---

# Update-GraphOneNotePage

## SYNOPSIS
Update a OneNote page

## SYNTAX

```
Update-GraphOneNotePage [-Page] <Object> [[-Action] <String>] [-Content] <String> [[-Position] <String>]
 [[-Target] <String>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This command makes PATCH requests to https://graph.microsoft.com/v1.0
    /users/{id}/onenote/sections/{id}/pages/{id}/content
or /groups/{id}/onenote/sections/{id}/pages/{id}/content
or  /sites/{id}/onenote/sections/{id}/pages/{id}/content
which requires consent to use the Notes.ReadWrite  scope or better.
To understand the use of Target, action & Postion and what needs to
be in content for different scenarios, read the MSFT page at the link ...

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Page
A graph URI pointing to the page, or a page object where the .self property is a graph URI...

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

### -Action
The action to perform on the target element.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Append
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
A string of well-formed HTML to add to the page, and any image or file binary data.

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

### -Position
The location to add the supplied content, relative to the target element.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Target
The element to update.
Must be the #\<data-id\> or the generated \<id\> of the element, or the body or title keyword.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: Body
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
If specified, the page is updated without prompting.

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

[https://docs.microsoft.com/en-gb/graph/onenote-update-page](https://docs.microsoft.com/en-gb/graph/onenote-update-page)


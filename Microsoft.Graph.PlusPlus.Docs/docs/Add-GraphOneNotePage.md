---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphOneNotePage

## SYNOPSIS
Adds a page (in HTML format) to an existing OneNote Section

## SYNTAX

```
Add-GraphOneNotePage [-Section] <Object> [-HTMLPage] <Object> [[-ContentType] <Object>] [-Force] [-PassThru]
 [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This posts to https://graph.microsoft.com/v1.0
    /users/{id}/onenote/sections/{id}/pages
or /groups/{id}/onenote/sections/{id}/pages
or  /sites/{id}/onenote/sections/{id}/pages
which requires consent to use the Notes.Create or Notes.ReadWrite scope or better.
To recognise the title the page needs to be in HTML with a head tag like this
\<html\>
    \<head\>
        \<title\>A page\</title\>
        \<meta name="created" content="2015-07-22T09:00:00-08:00" /\>
    \</head\>
    \<body\>
        \<p\>Here's Some text\</p\>
    \</body\>
\</html\>

## EXAMPLES

### EXAMPLE 1
```
Add-GraphOneNotePage -Section $section -HTMLPage '<html><head><title>Test Page</Title></head><body><p>Sample Paragraph</p></body></html>'
With $Section already defined this adds a simple page, with a title and a short body.
```

## PARAMETERS

### -Section
The section either as a URL or or as section object, which contains a self URL or a pages URL  this can be set by Set-GraphOneNoteHome

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HTMLPage
The content of the page formatted as HTML

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ContentType
By default this is "text/html" - but if the content is multipart use "multipart/form-data; boundary={MARKER}"

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Text/html
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
If specified, the command will run without asking for confirmation; this is the default unless Confirm Preference has been set

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
Normally the page is added 'silently'.
If passthru is specified, an object describing the new page will be returned.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: PT

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

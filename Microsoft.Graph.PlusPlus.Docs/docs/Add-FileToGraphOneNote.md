---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-FileToGraphOneNote

## SYNOPSIS
Adds a file to a new OneNote page

## SYNTAX

```
Add-FileToGraphOneNote [-Path] <Object> [[-Title] <String>] [[-Section] <Object>] [[-PreContent] <String[]>]
 [[-PostContent] <String[]>] [[-MimeType] <String>] [-PassThru] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
Adds a file to a new one page.
If the file is an image, the it will be rendered on the page
Other files will be embedded.
OneNote can render some types (e.g.
PDF)
This builds very simple HTML, which can be updated later.
For more sophistaced pages use Add-GraphOneNotePage - with -HTMLPage as a byte array and
specify a contentType of "multipart/form-data; boundary={MARKER}"

## EXAMPLES

### EXAMPLE 1
```
>Add-FileToGraphOneNote -Path .\Modules\MsftGraph\Examples\upload.jpg -Title "Demo" -Section $notebook.sections[0] `
          -PreContent "<h1>QR Code for the GIT repo</h1>" -PostContent "<b>Share and Enjoy</b>" -PassThru
```

$Notebook holds a notebook object with one or more section(s).
The command adds a page in the first section,
titles it "Demo", and puts upload.jpg on it with formatted text before and after the image.

## PARAMETERS

### -Path
The file to upload to OneNote

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

### -Title
Title for the page.
If not specified the file name will be used.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Section to post to -  this can be set by Set-GraphOneNoteHome

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreContent
Specifies text to add before the embedded object.
By default, there is no text in that position.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PostContent
Specifies text to add after the embedded object.
By default, there is no text in that position.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MimeType
A recognized mime type for the embedded file.
on Windows the command will try to determine this from the file extension.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Normally the page containing the file is added 'silently'.
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

### -Force
If specified the command will not pause for conformation, this is the default unless $ConfirmPreference is modified,

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

### A file to be sent to OneNote
## OUTPUTS

## NOTES

## RELATED LINKS

[Add-GraphOneNotePage.]()


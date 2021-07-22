---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Out-GraphOneNote

## SYNOPSIS
Output to a new OneNote page

## SYNTAX

### Page (Default)
```
Out-GraphOneNote [-InputObject <PSObject>] [[-Property] <String[]>] [-Section <Object>] [[-Body] <String[]>]
 [[-Head] <String[]>] [[-Title] <String>] [-As <String>] [-ExcludeProperty <String[]>] [-PreContent <String[]>]
 [-PostContent <String[]>] [-PassThru] [-Show] [<CommonParameters>]
```

### Fragment
```
Out-GraphOneNote [-InputObject <PSObject>] [[-Property] <String[]>] [-Section <Object>] [-As <String>]
 [-Fragment] [-ExcludeProperty <String[]>] [-PreContent <String[]>] [-PostContent <String[]>] [-PassThru]
 [-Show] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Generates a page
```

### EXAMPLE 2
```
start ( Get-process  | Out-GraphOneNote -Title "Processes @ $(get-date)" -property Name,Handles,NPM,PM,VM,WS -passthru ).links.oneNoteWebUrl.href
Generates a page in the default section (using the environment variable DefaultOneNoteSection) and opens it in a web browser.
```

## PARAMETERS

### -InputObject
Specifies the objects to be represented in HTML.

```yaml
Type: PSObject
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Property
Includes the specified properties of the objects in the output

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: @('*')
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
The section where the content will be created: to this can be set by Set-GraphOneNoteHome

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

### -Body
Specifies the text to add after the opening \<BODY\> tag.
By default, there is no text in that position.

```yaml
Type: String[]
Parameter Sets: Page
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Head
Specifies the content of the \<HEAD\> tag.
The default is "\<title\>HTML TABLE\</title\>". 
If you use the Head parameter, the Title parameter is ignored.

```yaml
Type: String[]
Parameter Sets: Page
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Specifies a title for the Page.

```yaml
Type: String
Parameter Sets: Page
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -As
Determines whether the object is formatted as a table or a list.
Valid values are TABLE and LIST.
The default value is TABLE.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Table
Accept pipeline input: False
Accept wildcard characters: False
```

### -Fragment
Generates only an HTML table.
The HTML, HEAD, TITLE, and BODY tags are omitted.

```yaml
Type: SwitchParameter
Parameter Sets: Fragment
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty
{{ Fill ExcludeProperty Description }}

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreContent
Specifies text to add before the opening \<TABLE\> tag.
By default, there is no text in that position.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PostContent
Specifies text to add after the closing \</TABLE\> tag.
By default, there is no text in that position.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
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

### -Show
If Specified opens the newly created page

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

### You can pipe any .NET object to Out-GraphOneNote
## OUTPUTS

## NOTES

## RELATED LINKS

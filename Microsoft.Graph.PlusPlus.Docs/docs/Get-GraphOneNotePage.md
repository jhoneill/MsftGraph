---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphOneNotePage

## SYNOPSIS
Gets a OneNote page's metadata or content

## SYNTAX

### None (Default)
```
Get-GraphOneNotePage [-Notebook <Object>] [-Section <Object>] [<CommonParameters>]
```

### PagePreview
```
Get-GraphOneNotePage [-Page] <Object> [-Notebook <Object>] [-Section <Object>] [-PreviewText]
 [-SavePath <Object>] [<CommonParameters>]
```

### PageContentWithIDs
```
Get-GraphOneNotePage [-Page] <Object> [-Notebook <Object>] [-Section <Object>] [-ContentWithIDs]
 [-SavePath <Object>] [<CommonParameters>]
```

### PageContent
```
Get-GraphOneNotePage [-Page] <Object> [-Notebook <Object>] [-Section <Object>] [-Content] [-ContentWithIDs]
 [-SavePath <Object>] [<CommonParameters>]
```

### Page
```
Get-GraphOneNotePage [-Page] <Object> [-Notebook <Object>] [-Section <Object>] [<CommonParameters>]
```

## DESCRIPTION
This command interogates  https://graph.microsoft.com/v1.0
    /users/{id}/onenote/notebooks/{id}/sections/{id}/pages
or /groups/{id}/onenote/notebooks/{id}/sections/{id}/pages
or  /sites/{id}/onenote/notebooks/{id}/sections/{id}/pages
which requires consent to use the  Notes.Read scope or better.
It can get either the page metadata, the page content, or
the page content marked up with IDs to update the page.

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
Parameter Sets: PagePreview, PageContentWithIDs, PageContent, Page
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Notebook
A graph URI pointing to a notebook, or a notebook object.
this can be set by Set-GraphOneNoteHome

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

### -Section
A graph URI pointing to a section, or a Section object  this can be set by Set-GraphOneNoteHome

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

### -Content
If specified returns the contents of the page.
Ignored if ContentWithIDs is specified

```yaml
Type: SwitchParameter
Parameter Sets: PageContent
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -ContentWithIDs
If specified returs the contents with guids for each section where content can be inserted.

```yaml
Type: SwitchParameter
Parameter Sets: PageContentWithIDs
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

```yaml
Type: SwitchParameter
Parameter Sets: PageContent
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -PreviewText
If specified returs a text preview of the page

```yaml
Type: SwitchParameter
Parameter Sets: PagePreview
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -SavePath
If specified writes the preview or content to a file

```yaml
Type: Object
Parameter Sets: PagePreview, PageContentWithIDs, PageContent
Aliases:

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

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenotePage
## NOTES

## RELATED LINKS

---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphOneNoteSection

## SYNOPSIS
Gets details of  sections in OneNote notebooks or their pages

## SYNTAX

```
Get-GraphOneNoteSection [-Section] <Object> [-Notebook <Object>] [-AllPages] [-PageTitle <String>]
 [<CommonParameters>]
```

## DESCRIPTION
This command interogates  https://graph.microsoft.com/v1.0
    /users/{id}/onenote/notebooks/{id}/sections
or /groups/{id}/onenote/notebooks/{id}/sections
or  /sites/{id}/onenote/notebooks/{id}/sections
which requires consent to use the Notes.Create or Notes.Read scope or better.
If given a Notebook parameter it returns the sections in the notebook.
If given a section parameter it either returns details of the section, or
if the -Pages or -Name Parameters are given returns pages from the section

## EXAMPLES

### EXAMPLE 1
```
$notebook = Get-GraphTeam  consultants -Notebooks
>$notebook.sections[0]  | Get-GraphOneNoteSection  -PageTitle change
The first line gets the Notebooks object for the 'consultants' team. This object
has a 'sections' collection. The second line uses pipes a member of this collection as the
into Get-GraphOneNoteSection to return the pages in the first section, with the title begining "change".
```

### EXAMPLE 2
```
Get-GraphOneNoteSection private -notebook $notebook -allpages
In this example the notebook used in the first example is passed as a notebook is piped into command to get a section, by contrast with the previous section
```

### EXAMPLE 3
```
Get-GraphOneNoteSection -Section $section -Pages -Name "test" | Remove-GraphOneNotePage -Force
>Gets all pages with names that begin 'Test...' and removes
$section may be the a section object (from the Sections collection of a notebook object, or
form Get-GraphOneNotebook -Sections ) or the URL for a section.
```

## PARAMETERS

### -Section
A graph URI pointing to the section, or a section object where the .self property is a graph URI or a section name...

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

### -Notebook
The notebook to query for section(s) if sections is empty or contains a name

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

### -AllPages
If specified, returns the pages in the section(s).

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

### -PageTitle
If specified filters pages or Sections to those with names beginning ...

```yaml
Type: String
Parameter Sets: (All)
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
### Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenoteSection
## NOTES

## RELATED LINKS

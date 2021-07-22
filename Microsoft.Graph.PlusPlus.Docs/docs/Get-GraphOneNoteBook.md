---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphOneNoteBook

## SYNOPSIS
Gets notebook objects or sections of notebooks

## SYNTAX

```
Get-GraphOneNoteBook [[-Notebook] <Object>] [[-InputObject] <Object>] [-AllSections] [[-SectionName] <String>]
 [<CommonParameters>]
```

## DESCRIPTION
If run with no parameters it will return the current user's personal notebooks.
If run with just a -Notebook parameter it will return that notebook (which might belong to a group)
If run with -Notebook and -Sections it will return the sections in that notebook,
And if run with just -Sections it will return all the sections in the user's personal notebooks.

## EXAMPLES

### EXAMPLE 1
```
Get-GraphOneNoteBook   team
Looks for a workbook with a displayname begining "team" in the users workbooks. the search is case insensitive.
```

### EXAMPLE 2
```
Get-GraphOneNoteBook  -SectionName Powershell
Finds a "PowerShell" secion in any of the users workbooks. Again the search is case insensitive
```

### EXAMPLE 3
```
Get-GraphTeam 'Consultants' -Notebooks | Set-GraphHomeNotebook
>Get-GraphOneNoteBook -AllSections
The first command changes the default notebook and selects different sections from the the previous command
```

## PARAMETERS

### -Notebook
A graph URI pointing to the notebook, or a notebook object where the .self property is a graph URI...

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
{{ Fill InputObject Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -AllSections
If specified returns the sections of the notebook.

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

### -SectionName
if specified filters the returned objects by to those with names begining with ...

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
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

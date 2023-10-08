---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/numbercolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphOneNoteSection

## SYNOPSIS
Adds a section to a OneNote notebook

## SYNTAX

```
New-GraphOneNoteSection [-Notebook] <Object> [-SectionName] <Object> [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
This command Posts to  https://graph.microsoft.com/v1.0
    /users/{id}/onenote/notebooks/{id}/sections
or /groups/{id}/onenote/notebooks/{id}/sections
or  /sites/{id}/onenote/notebooks/{id}/sections
which requires consent to use the Notes.Create or Notes.ReadWrite scope or better.

## EXAMPLES

### EXAMPLE 1
```
>$notebook = Get-GraphTeam -ByName accounts -Notebooks
>$section = New-GraphOneNoteSection -Notebook $notebook -SectionName "FY-19 Year End"
>Add-GraphOneNotePage -Section $section -HTMLPage '<html><head><title>Welcome</Title></head><body><p>This section is ready for you to add your pages.</p></body></html>'
```

The first command gets the team notebook for the account team; the second adds a section to it
and the third adds a welcome page to the new section.

## PARAMETERS

### -Notebook
A graph URI pointing to the notebook, or a notebook object, this can be set by Set-GraphOneNoteHome

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

### -SectionName
Name for the new section.

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

### Returns an object representing the new section
## NOTES

## RELATED LINKS

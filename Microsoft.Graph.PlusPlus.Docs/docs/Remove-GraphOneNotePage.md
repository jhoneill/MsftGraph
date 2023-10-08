---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Remove-GraphOneNotePage

## SYNOPSIS
Removes a OneNote page

## SYNTAX

```
Remove-GraphOneNotePage [-Page] <Object> [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This command makes DELETE requests to https://graph.microsoft.com/v1.0
     /users/{id}/onenote/sections/{id}/pages/{id}
 or /groups/{id}/onenote/sections/{id}/pages/{id}
 or  /sites/{id}/onenote/sections/{id}/pages/{id}
 which requires consent to use the Notes.ReadWrite scope or better.

## EXAMPLES

### EXAMPLE 1
```
Get-GraphUser -Teams -Name Consultants | Get-GraphTeam  -Notebooks |
   Get-GraphOneNoteBook -Sections -Name General | Get-GraphOneNoteSection -Pages -Name process | Remove-GraphOneNotePage
finds a team named "consultants" which has the current user as a member, finds its notebook, finds a section named General
within this sectioned finds page names that begin "process..." and removes them
```

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

### -Force
If specified, the page is deleted without prompting.

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

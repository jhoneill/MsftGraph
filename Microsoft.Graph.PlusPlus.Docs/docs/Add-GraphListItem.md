---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Add-GraphListItem

## SYNOPSIS
Adds an item to a SharePoint List

## SYNTAX

```
Add-GraphListItem [-List] <Object> [-Fields] <Hashtable> [-Site <Object>] [-Passthru] [-Force] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This posts a new item to https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items
which requires consent to use the Sites.ReadWrite.All scope
Posting to a list is quite basic - it is a set of Name-ValuePairs and
FIELD NAMES ARE CASE SENSITIVE.
If you get a 400 error from the server the
first thing to check is the names of the fields.
It does not appear to be possible to
post certain types of field - lookup and Person/Group being the major issues.
The command will try to post what it is given, but it makes no attempt at validating it!

## EXAMPLES

### EXAMPLE 1
```
>$myteamsite = Get-GraphTeam -Site |select -first 1
>$problemslist = $myteamsite.lists.where({$_.name -like "problem*"})
>Add-GraphListItem  -List $problemslist -Fields @{Title='Demo Item';IssueStatus='Active';Priority='(2) Normal';}
```

The first command gets a team site which has a list named "Problem reports"
The second line gets that list
The third creates a list item with Title, IssueStatus and Priority fields.

## PARAMETERS

### -List
The list to add to; this can be an ID, or list object with an ID, and a site ID

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

### -Fields
The item property values in a hash table as @{col1=$value1; col2='Value2'; col3=33}

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Site
If the list parameter does not contain a .SiteID property allows the site to specified as an ID or object

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

### -Passthru
If specified the new item will be returned, otherwise it is created silently.

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
If specified the item will be added without prompting for confirmation (this is the default unless confirm preference is changed)

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

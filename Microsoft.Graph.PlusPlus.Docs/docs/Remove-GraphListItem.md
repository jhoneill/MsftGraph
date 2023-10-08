---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# Remove-GraphListItem

## SYNOPSIS
Deletes an item from a SharePoint List

## SYNTAX

```
Remove-GraphListItem [-Item] <Object> [-List <Object>] [-Site <Object>] [-Force] [-WhatIf] [-Confirm]
 [<CommonParameters>]
```

## DESCRIPTION
This Deletes an item at https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items{id}
which requires consent to use the Sites.ReadWrite.All scope

## EXAMPLES

### EXAMPLE 1
```
>$problemitems = get-graphlist $problemslist -Items -Property title,issuestatus,AssignedToLookupID,priority
>$problemitems[4] | Remove-GraphListItem
```

The first line gets the items from a list , and the second line removes the fifth one

## PARAMETERS

### -Item
The item to remove; this can be an ID or an object with an ID, and a list and site ID as well

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

### -List
If the item does not contain the list, the list to delete from an ID, or list object with an ID, and a site ID

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

### -Site
If there is no site id in the item or list parameter allows the site to specified as an ID or object

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

### -Force
If specified the item will be deleted without prompting for confirmation (prompting is the default)

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

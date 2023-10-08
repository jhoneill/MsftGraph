---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphListItem

## SYNOPSIS
Updates an item in a SharePoint List

## SYNTAX

```
Set-GraphListItem [-Item] <Object> [-Fields] <Hashtable> [-List <Object>] [-Site <Object>] [-Force] [-WhatIf]
 [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
This Patches an existing item at https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items{id}/Fields
which requires consent to use the Sites.ReadWrite.All scope
Caveats in Add-GraphListItem apply to Set-GraphListItem.

## EXAMPLES

### EXAMPLE 1
```
>$problemitems = get-graphlist $problemslist -Items -Property title,issuestatus,AssignedToLookupID,priority
>$problemitems[2] | Set-GraphListItem -Fields @{Priority='(2) Normal'}
```

The first line gets the items from a list , and the second updates the Priority field of the third one

## PARAMETERS

### -Item
The item to update; this can be an ID or an object with an ID, and a list and site ID as well

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
If specified the item will be updated without prompting for confirmation

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

[Add-GraphListItem]()


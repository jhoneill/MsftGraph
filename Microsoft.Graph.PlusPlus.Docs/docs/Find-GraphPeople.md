---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Find-GraphPeople

## SYNOPSIS
Searches people in your inbox / contacts / directory

## SYNTAX

### Default (Default)
```
Find-GraphPeople [-Topic] <Object> [-First <Int32>] [<CommonParameters>]
```

### Fuzzy
```
Find-GraphPeople -SearchTerm <Object> [-First <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Requires consent to use either the People.Read or the People.Read.All scope

## EXAMPLES

### EXAMPLE 1
```
Find-GraphPeople -Topic timesheet -First 6
Returns the top 6 results for people you have discussed timesheets with.
```

## PARAMETERS

### -Topic
Text to use in a 'Topic' Search.
Topics are not pre-defined, but inferred using machine learning based on your conversation history (!)

```yaml
Type: Object
Parameter Sets: Default
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -SearchTerm
Text to use in a search on name and email address

```yaml
Type: Object
Parameter Sets: Fuzzy
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -First
Number of results to return (10 by default)

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 10
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

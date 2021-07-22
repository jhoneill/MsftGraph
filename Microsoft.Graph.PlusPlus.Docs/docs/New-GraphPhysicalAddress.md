---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version: https://docs.microsoft.com/en-us/graph/api/resources/personorgroupcolumn?view=graph-rest-1.0
schema: 2.0.0
---

# New-GraphPhysicalAddress

## SYNOPSIS
Builds a street / postal / physical address to use in the contact commands

## SYNTAX

```
New-GraphPhysicalAddress [[-Street] <String>] [[-City] <String>] [[-State] <String>] [[-PostalCode] <String>]
 [[-CountryOrRegion] <String>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
$fabrikamAddress = New-GraphPhysicalAddress  "123 Some Street" Seattle WA 98121 "United States"
Creates an address - if the -Street, City,  State, Postalcode country are not explictly
specified they will be assigned in that order. Quotes are desireable but only necessary
when a value contains spaces.
It can then be used like this. Set-GraphContact $pavel -BusinessAddress $fabrikamAddress
```

## PARAMETERS

### -Street
Street address.
This can contain carriage returns for a district, e.g.
"101 London Road\`r\`nBotley"

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -City
City, or town as people outside the US tend to call it

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -State
State, Province, County, the administrative level below country

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

### -PostalCode
Postal code.
Even it parses as a number, as with US ZIP codes, it will be converted to a string

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CountryOrRegion
Usually a country but could be some other geographical entity

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
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

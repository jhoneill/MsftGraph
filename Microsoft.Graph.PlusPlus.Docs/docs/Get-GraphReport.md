---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphReport

## SYNOPSIS
Gets reports from MS Graph

## SYNTAX

```
Get-GraphReport [-Report] <Object> [[-Date] <DateTime>] [[-Period] <Object>] [[-Path] <Object>]
 [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphReport -Report MailboxUsageDetail | ft "Display Name",  "Storage Used (Byte)"
Displays mailbox storage used by users - note that
fields have 'friendly' names which need to be wrapped in quotes
```

## PARAMETERS

### -Report
The report to Fetch

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

### -Date
Date for the report - this should be a date in the past 30 days.
If specified, -Period is ignored.
Reports ending in Count, Storage or pages don't support date filtering

```yaml
Type: DateTime
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Period
The range of time for the report in the form "Dn" where n is the number of days.
The default is D7, except for Office365Activation activation reports

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
If specified the data will be written in CSV format to the path provided, otherwise it will be output to the pipeline

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
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

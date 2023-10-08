---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphSignInLog

## SYNOPSIS
Gets the audit log -requires a priviledged account

## SYNTAX

```
Get-GraphSignInLog [[-top] <Object>] [<CommonParameters>]
```

## DESCRIPTION
This command calls https://graph.microsoft.com/beta/auditLogs/signIns
which requires consent to use the AuditLog.Read.All Scope this can only be granted to Azure AD apps.

## EXAMPLES

### EXAMPLE 1
```
>Get-GraphSignInLog |
>  select Date,UserPrincipalName,appDisplayName,ipAddress,clientAppUsed,browser,device,city,lat,long |
>    Export-Excel -Path .\signin.xlsx -AutoSize -IncludePivotTable -PivotTableName Signins -PivotRows appdisplayName -PivotColumns browser -PivotData @{date='Count'} -show
```

Gets the sign-in Log and exports it Excel, creating a PivotTable

## PARAMETERS

### -top
{{ Fill top Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: 200
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphSignIn
## NOTES

## RELATED LINKS

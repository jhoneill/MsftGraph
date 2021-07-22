---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphOrganization

## SYNOPSIS
Gets a summary of organization information from MSGraph

## SYNTAX

```
Get-GraphOrganization [[-Organization] <Object>] [[-Proxy] <Uri>] [[-ProxyCredential] <PSCredential>]
 [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

## DESCRIPTION
Can use msonline\Get-MsolCompanyInformation instead
This needs consent to use either the User.Read or the Directory.Read.All scope

## EXAMPLES

### EXAMPLE 1
```
(Get-GraphOrganization).verifiedDomains
Displays a list of domains in the current subscription
```

## PARAMETERS

### -Organization
{{ Fill Organization Description }}

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

### -Proxy
The URI for the proxy server to use

```yaml
Type: Uri
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProxyCredential
Credentials for a proxy server to use for the remote call

```yaml
Type: PSCredential
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProxyUseDefaultCredentials
Use the default credentials for the proxy

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphOrganization
## NOTES

## RELATED LINKS

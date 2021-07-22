---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Connect-Graph

## SYNOPSIS
Starts a session with Microsoft Graph

## SYNTAX

### UserParameterSet (Default)
```
Connect-Graph [[-Scopes] <String[]>] [-ForceRefresh] [-Quiet] [<CommonParameters>]
```

### AccessTokenParameterSet
```
Connect-Graph [-AccessToken] <String> [-ForceRefresh] [-Quiet] [<CommonParameters>]
```

## DESCRIPTION
This commands is a wrapper for Connect-MgGraph it extends the authentication methods available
and caches information needed by other commands.

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Scopes
An array of delegated permissions to consent to.

```yaml
Type: String[]
Parameter Sets: UserParameterSet
Aliases:

Required: False
Position: 2
Default value: $Script:DefaultGraphScopes
Accept pipeline input: False
Accept wildcard characters: False
```

### -AccessToken
Specifies a bearer token for Microsoft Graph service.
Access tokens do timeout and you'll have to handle their refresh.

```yaml
Type: String
Parameter Sets: AccessTokenParameterSet
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForceRefresh
Forces the command to get a new access token silently.

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

### -Quiet
Suppress the welcome messages

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

## NOTES

## RELATED LINKS

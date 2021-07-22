---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Set-GraphOptions

## SYNOPSIS
Sets defaults and the tenant client ID & Client Secret used when logging on without a web dialog

## SYNTAX

```
Set-GraphOptions [[-TenantID] <Object>] [[-ClientID] <Object>] [[-ClientSecret] <Object>]
 [[-DefaultScopes] <Object>] [[-RefreshToken] <Object>] [[-DefaultUserProperties] <String[]>]
 [[-DefaultUsageLocation] <String>] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -TenantID
Your Tennant ID

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

### -ClientID
Client ID if not using the SDK default of 14d82eec-204b-4c2f-b7e8-296a70dab67e.
Must be known to your tennant

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClientSecret
Secret set for the client ID in your $TenantID

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Client_Secret,

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultScopes
Default Scopes to request

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

### -RefreshToken
Allows a saved Refresh Token (e.g.
from Show-GraphSession) to be added to the session.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 5
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultUserProperties
Changes the dafault properties returned by Get-GraphUser and Get-GraphUserList

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 6
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DefaultUsageLocation
Changes the default two letter (ISO  3166) country code - for new users so they can be assigned licenses. 
Examples include: 'US', 'JP', and 'GB'

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
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

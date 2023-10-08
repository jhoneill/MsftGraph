---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-AccessToken

## SYNOPSIS
Requests a token for a resource path, used by connect graph but available to other tools.

## SYNTAX

```
Get-AccessToken [[-Resoure] <String>] [-GrantType] <String> [[-BodyParts] <Hashtable>] [<CommonParameters>]
```

## DESCRIPTION
An access token is obtained form "https://login.microsoft.com/\<\<tenant-ID\>\>/oauth2/token"
By specifying the ID, and secret of a client app known to in that tenant,
different modes of granting a token to access a resource (logging on) are possible:
Extra fields passed in BodyParts    grant_type
* Username and password            'Password'
* Refresh_token                    'Referesh_token'
* None (logon as the app itself)   'client_credentials'
The Set-GraphOptions command sets the tenant ID, a client ID and a client secret
for the session. 
By default, when the module loads it looks at $env:GraphSettingsPath or for
Microsoft.Graph.PlusPlus.settings.ps1  in the module folder, and executes it to set these values)
Get-AccessToken relies on these if they are not set Connect-Graph removes the parameters
which support non-intereactive logons and calling it seperately will fail

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -Resoure
{{ Fill Resoure Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 1
Default value: Https://graph.microsoft.com
Accept pipeline input: False
Accept wildcard characters: False
```

### -GrantType
{{ Fill GrantType Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BodyParts
{{ Fill BodyParts Description }}

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: @{}
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS

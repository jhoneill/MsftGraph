---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphUserList

## SYNOPSIS
Returns a list of Azure active directory users for the current tennant.

## SYNTAX

### None (Default)
```
Get-GraphUserList [-Select <String[]>] [-Top <Object>] [-ExpandProperty <String>] [-Proxy <Uri>]
 [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

### FilterByName
```
Get-GraphUserList [-Name] <String[]> [-Select <String[]>] [-Top <Object>] [-ExpandProperty <String>]
 [-Proxy <Uri>] [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

### Sorted
```
Get-GraphUserList [-Select <String[]>] [-Top <Object>] -Sort <String> [-ExpandProperty <String>] [-Proxy <Uri>]
 [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

### FilterByString
```
Get-GraphUserList [-Select <String[]>] [-Top <Object>] -Filter <String> [-ExpandProperty <String>]
 [-Proxy <Uri>] [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

### FilterToMembers
```
Get-GraphUserList [-Select <String[]>] [-Top <Object>] [-MembersOnly] [-ExpandProperty <String>] [-Proxy <Uri>]
 [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

### FilterToGuests
```
Get-GraphUserList [-Select <String[]>] [-Top <Object>] [-GuestsOnly] [-ExpandProperty <String>] [-Proxy <Uri>]
 [-ProxyCredential <PSCredential>] [-ProxyUseDefaultCredentials] [<CommonParameters>]
```

## DESCRIPTION
{{ Fill in the Description }}

## EXAMPLES

### EXAMPLE 1
```
Get-GraphUserList -filter "Department eq 'Accounts'"
Gets the list with a custom filter this is typically fieldname eq 'value' for equals or
startswith(fieldname,'value') clauses can be joined with and / or.
```

## PARAMETERS

### -Name
If specified searches for users whose first name, surname, displayname, mail address or UPN start with that name.

```yaml
Type: String[]
Parameter Sets: FilterByName
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Select
Names of the fields to return for each user.Note that some properties - aboutMe, Birthday etc, are only available when getting a single user, not a list.
The  API defaults to :  businessPhones, displayName, givenName, id, jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName
The module adds to this set - the exactlist can be set with Set-GraphOption -DefaultUserProperties

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Property

Required: False
Position: Named
Default value: $Script:DefaultUserProperties
Accept pipeline input: False
Accept wildcard characters: False
```

### -Top
The default is to get all

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

### -Sort
Order by clause for the query - most fields result in an error and it can't be combined with some other query values.

```yaml
Type: String
Parameter Sets: Sorted
Aliases: OrderBy

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
Filter clause for the query for example "startswith(displayname,'Bob') or startswith(displayname,'Robert')"

```yaml
Type: String
Parameter Sets: FilterByString
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MembersOnly
Adds a filter clause "userType eq 'Member'"

```yaml
Type: SwitchParameter
Parameter Sets: FilterToMembers
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -GuestsOnly
Adds a filter clause "userType eq 'Guest'"

```yaml
Type: SwitchParameter
Parameter Sets: FilterToGuests
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExpandProperty
{{ Fill ExpandProperty Description }}

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Manager
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
Position: Named
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
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ProxyUseDefaultCredentials
Use the default credentials for the proxygit

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

### Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser
## NOTES

## RELATED LINKS

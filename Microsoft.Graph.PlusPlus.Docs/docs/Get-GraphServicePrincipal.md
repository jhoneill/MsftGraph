---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphServicePrincipal

## SYNOPSIS
Returns information about Service Principals

## SYNTAX

### List1 (Default)
```
Get-GraphServicePrincipal [-Property <String[]>] [-Filter <String>] [<CommonParameters>]
```

### Get2
```
Get-GraphServicePrincipal [-ServicePrincipalId] <Object> [-Property <String[]>] [<CommonParameters>]
```

### FilteredScopes
```
Get-GraphServicePrincipal [-ServicePrincipalId] <Object> [-Property <String[]>] -ScopeFilter <String>
 [<CommonParameters>]
```

### AllScopes
```
Get-GraphServicePrincipal [-ServicePrincipalId] <Object> [-Property <String[]>] [-ExpandScopes]
 [<CommonParameters>]
```

### FilteredRoles
```
Get-GraphServicePrincipal [-ServicePrincipalId] <Object> [-Property <String[]>] -AppRoleFilter <String>
 [<CommonParameters>]
```

### AllRoles
```
Get-GraphServicePrincipal [-ServicePrincipalId] <Object> [-Property <String[]>] [-ExpandAppRoles]
 [<CommonParameters>]
```

### List5
```
Get-GraphServicePrincipal [-AppId <String>] [-Property <String[]>] [<CommonParameters>]
```

### List2
```
Get-GraphServicePrincipal [-ManagedIdentity] [-Property <String[]>] [<CommonParameters>]
```

### List3
```
Get-GraphServicePrincipal [-Application] [-Property <String[]>] [<CommonParameters>]
```

### List4
```
Get-GraphServicePrincipal [-O365ServicePrincipals] [-Property <String[]>] [<CommonParameters>]
```

## DESCRIPTION
A replacement for the SDK's Get-MgServicePrincipal
That has orderby which doesn't work - it's in the Docs but the API errors if you try
It doesn't have find by name, or select Application or Managed IDs

## EXAMPLES

### EXAMPLE 1
```
Get-GraphServicePrincipal "Microsoft graph*"
```

Id                                   DisplayName                      AppId                                SignInAudience
--                                   -----------                      -----                                --------------
25b13fbf-2f44-457a-9e68-d3414fc97915 Microsoft Graph                  00000003-0000-0000-c000-000000000000 AzureADMultipleOrgs
4e71d88a-0a46-4274-85b8-82ad86877010 Microsoft Graph Change Tracking  0bf30f3b-4a52-48df-9a82-234910c4a086 AzureADMultipleOrgs
...

Run with a name the command returns service principals with matching names.

### EXAMPLE 2
```
Get-GraphServicePrincipal 25b13fbf-2f44-457a-9e68-d3414fc97915 -ExpandAppRoles
```

Value                         DisplayName                Enabled Id
-----                         -----------                ------- --
AccessReview.Read.All         Read all access reviews    True    d07a8cc0-3d51-4b77-b3b0-32704d1f69fa
AccessReview.ReadWrite.All    Manage all access reviews  True    ef5f7d5c-338f-44b0-86c3-351f46c8bb5f
...
In this example GUID for Microsoft Graph was used from the previous example, and the command has listed the roles available to applications

## PARAMETERS

### -ServicePrincipalId
The GUID(s) for ServicePrincipal(s).
Or SP objects.
If a name is given instead, the command will try to resolve matching Service principals

```yaml
Type: Object
Parameter Sets: Get2, FilteredScopes, AllScopes, FilteredRoles, AllRoles
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -AppId
{{ Fill AppId Description }}

```yaml
Type: String
Parameter Sets: List5
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ManagedIdentity
Produces a list filtered to only managed identities

```yaml
Type: SwitchParameter
Parameter Sets: List2
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Application
Produces a list filtered to only applications

```yaml
Type: SwitchParameter
Parameter Sets: List3
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -O365ServicePrincipals
Produces a convenience list of office 365 security principals

```yaml
Type: SwitchParameter
Parameter Sets: List4
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property
Select properties to be returned

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: Select

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Filter
Filters items by property values

```yaml
Type: String
Parameter Sets: List1
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExpandAppRoles
Returns the list of application roles to those the role name, displayname or ID match the parameter value.
Wildcards are supported

```yaml
Type: SwitchParameter
Parameter Sets: AllRoles
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppRoleFilter
Filters the list of application roles available within a SP

```yaml
Type: String
Parameter Sets: FilteredRoles
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExpandScopes
Returns the list of (user) oauth scopes available within a SP

```yaml
Type: SwitchParameter
Parameter Sets: AllScopes
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ScopeFilter
Filters the list of oauth scopes to those where the scope name, displayname or ID match the parameter value.
Wildcards are supported

```yaml
Type: String
Parameter Sets: FilteredScopes
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.Graph.PowerShell.Models.IMicrosoftGraphAppRole
### Microsoft.Graph.PowerShell.Models.IMicrosoftGraphPermissionScope
### Microsoft.Graph.PowerShell.Models.IMicrosoftGraphServicePrincipal
## NOTES

## RELATED LINKS

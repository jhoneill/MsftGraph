---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphList

## SYNOPSIS
Gets sharepoint list objects or their items

## SYNTAX

### ListID (Default)
```
Get-GraphList [[-List] <Object>] [-Site <Object>] [<CommonParameters>]
```

### ListIDColumns
```
Get-GraphList [-List] <Object> [-Site <Object>] [-ColumnList] [<CommonParameters>]
```

### ListItems
```
Get-GraphList [-List] <Object> [-Site <Object>] [-Items] [-Property <String[]>] [<CommonParameters>]
```

### ListofLists
```
Get-GraphList [-Site <Object>] [-HIDden] [<CommonParameters>]
```

## DESCRIPTION
This interogates https://graph.microsoft.com/v1.0/sites/{id}/lists{id}
which requires consent to use the Sites.Read.All scope or better.
This does not suppor the use of a filter parameter so any "where"
operation has to be done on the returned data.

## EXAMPLES

### EXAMPLE 1
```
>$myTeamSite = Get-GraphTeam -Site | select -first 1
>$problemsList = $myteamsite.lists | where name -like problem*
>
> Get-GraphList  -list $problemslist -ColumnList
```

The first command gets the current users group(s) and returns their site(s).
For this example we select the first site.
The sites returned by Get-GraphGroup /
Get-GraphTeam have a .lists property and second command selects the list we want
The third line shows calling Get-GraphList using the ID for both Site and List
and  getting the columns in the list.
The next example shows an easier way to provide the information; and in fact
there is already a .columns property of $problemsList which has the column information

### EXAMPLE 2
```
Get-graphlist $problemsList -Items
This uses $problemsList from the previous example. Get-GraphGroup (aka Get-GraphTeam)
gets the Site, it gets the sites lists, and adds the site ID as a property, so
$Problemslist has propeties for the list ID and the site ID. So this exmaple uses a
shorter form of just providing the list and returns the items in their raw state
```

### EXAMPLE 3
```
Get-graphlist $problemsList -Items -Property title, issuestatus, AssignedToLookupID, priority
This builds on the previous example. Specifying -Property causes Get-GraphList to
return the Item(s) Fields collection(s) and sets the default fields to be displayed.
By default if an object has 4 visbible properties or fewer PowerShell displays it
as a table, if it has more than 4 a list is used, this can be managed with
$FormatEnumerationLimit. In this case 4 properties are show in a table view.
However 'Person or Group' fields, like AssignedTo return a lookupID.
This comes from the hidden list 'Users' and the next example shows how to get
information from this list. (The Get-GraphSiteUserList provides a shortcut for geting
this Information)
```

### EXAMPLE 4
```
>Get-GraphList -Site $myteamSite -Hidden  | where name -eq 'users' |
    Get-Graphlist -Items -Property id,ContentType,Title,Name
```

This uses the $myTeamSite variable from the first example.
If neither Items, nor ColumnList is specified, Get-GraphList returns list objects,
(the same result as using Get-GraphSite -Lists) so the first command gets lists
in the team site including hidden ones - which aren't included in the .lists
property of the site, and users IS hidden.
The where command isolates that list,
and it is piped into a second Get-GraphList command, which gets its items
and displays the properties of interest

### EXAMPLE 5
```
>$mydocuments = Get-GraphUser -Site | Get-GraphSite -lists | where name -eq documents
>Get-GraphList $shareddocsList -items | Select -expand driveItem |
      Copy-FromGraphFolder -Destination C:\temp
```

This command works with a users "MySite" - the first command gets the user's
site, gets its lists and selects the one named "Documents"
The second gets the items in this list; when a list object has an associated drive,
items returned by Get-GraphList -items will have a .DriveItem property.
Driveitems can be piped into  Copy-FromGraphFolder .

## PARAMETERS

### -List
The list either as an ID or as a list object (which may contain the site.)

```yaml
Type: Object
Parameter Sets: ListID
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

```yaml
Type: Object
Parameter Sets: ListIDColumns, ListItems
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Site
Specifies a site, if omitted "root" will be assumed - the root site of the user's tennant.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Root
Accept pipeline input: False
Accept wildcard characters: False
```

### -HIDden
If specified returns hidden lists (like 'Users')

```yaml
Type: SwitchParameter
Parameter Sets: ListofLists
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Items
If specified returns the list's items

```yaml
Type: SwitchParameter
Parameter Sets: ListItems
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnList
If specified returns the columns in the list

```yaml
Type: SwitchParameter
Parameter Sets: ListIDColumns
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property
if specified returned items will be expanded and the default display fields will be set

```yaml
Type: String[]
Parameter Sets: ListItems
Aliases: Fields

Required: False
Position: Named
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

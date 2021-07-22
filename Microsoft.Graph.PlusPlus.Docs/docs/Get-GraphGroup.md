---
external help file: Microsoft.Graph.PlusPlus-help.xml
Module Name: Microsoft.Graph.PlusPlus
online version:
schema: 2.0.0
---

# Get-GraphGroup

## SYNOPSIS
Gets information about a Group and any associated Office 365 Team

## SYNTAX

### None (Default)
```
Get-GraphGroup [[-ID] <Object>] [-Mine] [<CommonParameters>]
```

### Apps
```
Get-GraphGroup [[-ID] <Object>] [-Apps] [-AppName <String>] [-Mine] [<CommonParameters>]
```

### Calendar
```
Get-GraphGroup [[-ID] <Object>] [-Calendar] [-Mine] [<CommonParameters>]
```

### Channels
```
Get-GraphGroup [[-ID] <Object>] [-Channels] [-ChannelName <String>] [-Mine] [<CommonParameters>]
```

### Conversations
```
Get-GraphGroup [[-ID] <Object>] [-Conversations] [-Mine] [<CommonParameters>]
```

### Drive
```
Get-GraphGroup [[-ID] <Object>] [-Drive] [-Mine] [<CommonParameters>]
```

### Members
```
Get-GraphGroup [[-ID] <Object>] [-Members] [-Mine] [<CommonParameters>]
```

### TransitiveMembers
```
Get-GraphGroup [[-ID] <Object>] [-TransitiveMembers] [-Mine] [<CommonParameters>]
```

### Memberof
```
Get-GraphGroup [[-ID] <Object>] [-MemberOf] [-Mine] [<CommonParameters>]
```

### TransitiveMemberof
```
Get-GraphGroup [[-ID] <Object>] [-TransitiveMemberOf] [-Mine] [<CommonParameters>]
```

### Owners
```
Get-GraphGroup [[-ID] <Object>] [-Owners] [-Mine] [<CommonParameters>]
```

### Notebooks
```
Get-GraphGroup [[-ID] <Object>] [-Notebooks] [-Mine] [<CommonParameters>]
```

### Planners
```
Get-GraphGroup [[-ID] <Object>] [-Plans] [-Mine] [<CommonParameters>]
```

### Threads
```
Get-GraphGroup [[-ID] <Object>] [-Threads] [-Mine] [<CommonParameters>]
```

### Site
```
Get-GraphGroup [[-ID] <Object>] [-Site] [-Mine] [<CommonParameters>]
```

### SelectFields
```
Get-GraphGroup [[-ID] <Object>] -Select <String[]> [-Mine] [<CommonParameters>]
```

### BareGroups
```
Get-GraphGroup [[-ID] <Object>] [-Mine] [-NoTeamInfo] [<CommonParameters>]
```

## DESCRIPTION
Takes a Group/Team ID or object as a parameter and gets information about it.
Apps, Calendar, Channels, Drive, Members or Planners can be requested.
Depending on which aspect of the group are queried, may need access to the following
Scopes Group.Read.All, Files.Read, Sites.Read.All, Notes.Create, Notes.Read,

## EXAMPLES

### EXAMPLE 1
```
Get-GraphUser -teams | Get-GraphTeam -Plans | select -last 1 | Get-GraphPlan -FullTasks  | ft PlanTitle,Bucketname,Title,DueDateTime,PercentComplete,Assignees
 Gets the current user's Teams, and gets the plans for each;
 Note that because we are refering to "Teams" the command is called using its alias of Get-GraphTeam.
 The last plan is selected and details of the plan are fetched, showing the result as a table.
```

### EXAMPLE 2
```
(Get-GraphGroup -Site).lists | where name -match document
If no Group/Team is provided the command gets those associated with the current user;
it this case it returns their associated site(s).
Site objects include a lists property, which holds a collection of lists
This command will fiter the lists down to those where name matches "document",
giving the "Shared Documents" list for each team
```

### EXAMPLE 3
```
Get-GraphGroup -Drive  | Get-GraphDrive -Subfolders | Select  name, weburl, id,@{n="drive";e={$_.parentReference.driveId}}
As with the previous example gets this command gets Groups/Teams for current user,
in this case the command returns their associated drive(s)
It is possible to refer to the drive's 'root' property, and the root's 'childre'n property
which contains files and folder objects, and filter to objects with a folder property but
for ease of reading this  pipeline passes the drive to Get-GraphDrive to get subfolders.
It then returns the  name, WebURl and the item ID and Drive ID needed to access each folder.
```

### EXAMPLE 4
```
Get-GraphGroup 'Consultants' -Drive  | Set-GraphHomeDrive
Sets the drive for the consultants group to be the default graph drive for the PowerShell session.
```

### EXAMPLE 5
```
Get-GraphGroup -Notebooks | select -ExpandProperty sections | where "Displayname" -eq "General_Notes"
Again gets Groups/Teams for the current user and returns their associated notebooks(s)
Notebook objects include a Sections property, which holds a collection of OneNote sections in the notebook;
This command gets returns any section in a team notebook which has the name "General_Notes"
```

### EXAMPLE 6
```
Get-GraphTeam -threads | where LastDeliveredDateTime -gt ([datetime]::Now.AddDays(-7))
Gets the teams conversation threads which have been updated in the last 7 days.
```

## PARAMETERS

### -ID
The name of a team.
One more Team IDs or team objects containing and ID.
If omitted the current user's teams will be used.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Team, Group

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Apps
If specified returns the teams Apps

```yaml
Type: SwitchParameter
Parameter Sets: Apps
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Calendar
If specified gets the team's Calendar (a team only has one)

```yaml
Type: SwitchParameter
Parameter Sets: Calendar
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Channels
If specified gets the team's channels

```yaml
Type: SwitchParameter
Parameter Sets: Channels
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Conversations
If Specified, retrun team's conversations (usually better to use threads)

```yaml
Type: SwitchParameter
Parameter Sets: Conversations
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Drive
If specified gets the Team's OneDrive to see contents of the root of the drive you can refer to the drives .root.children property

```yaml
Type: SwitchParameter
Parameter Sets: Drive
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Members
If specified returns the members of the team

```yaml
Type: SwitchParameter
Parameter Sets: Members
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -TransitiveMembers
If specified returns the transitive members of the team

```yaml
Type: SwitchParameter
Parameter Sets: TransitiveMembers
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -MemberOf
If specified returns the groups this group is directly a member of

```yaml
Type: SwitchParameter
Parameter Sets: Memberof
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -TransitiveMemberOf
If specified returns the groups this group is nested into transitively

```yaml
Type: SwitchParameter
Parameter Sets: TransitiveMemberof
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Owners
If specified returns the Owners of the team

```yaml
Type: SwitchParameter
Parameter Sets: Owners
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Notebooks
If specified returns the team's notebook(s)

```yaml
Type: SwitchParameter
Parameter Sets: Notebooks
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Plans
if Specified, returns the teams Planners.

```yaml
Type: SwitchParameter
Parameter Sets: Planners
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Threads
If Specified, retrun team's threads

```yaml
Type: SwitchParameter
Parameter Sets: Threads
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Site
if Specified, returns the teams site.

```yaml
Type: SwitchParameter
Parameter Sets: Site
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -AppName
limits searches for appsby name.

```yaml
Type: String
Parameter Sets: Apps
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChannelName
limits searches for channels by name.
Other items can't be filtered by name ... 
perhaps notebooks can but a group only has one.

```yaml
Type: String
Parameter Sets: Channels
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Select
Field(s) to select: ID and displayname are always included
The following are available when getting a single group:

```yaml
Type: String[]
Parameter Sets: SelectFields
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Mine
{{ Fill Mine Description }}

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

### -NoTeamInfo
{{ Fill NoTeamInfo Description }}

```yaml
Type: SwitchParameter
Parameter Sets: BareGroups
Aliases:

Required: True
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

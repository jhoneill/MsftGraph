Get-Module | ForEach-Object {get-command -Module $_ -CommandType Cmdlet,Function} | Measure-Object | ForEach-Object count
Import-Module Microsoft.Graph.PlusPlus
Connect-Graph

Get-Module | ForEach-Object {get-command -Module $_ -CommandType Cmdlet,Function} | Measure-Object | ForEach-Object count

#Select some users and put them in a new team. For my demo I pre-set value of department => group membership
$GroupName        = 'Presenters'
$newProjectName   = "Vienna"
Get-GraphUserList -Filter "Department eq '$GroupName'" -OutVariable users  | Format-Table Organization
New-GraphTeam -Name $GroupName  -Description "The $GroupName Department" -Visibility public -Members $users -OutVariable newTeam

Get-GraphTeam $newTeam -Drive -OutVariable teamdrive |  Set-GraphHomeDrive ; $teamDrive
#later we will add a tab in teams for this drive

#special folder tab completes
Get-GraphDrive -SpecialFolder Documents
Get-GraphDrive /

#Send a local file to onedrive, and open it -use it for exporting in a moment.
Get-ChildItem $env:temp\Test*.xlsx -OutVariable files

#Destination tab completes - use General for preference
$files  |  Copy-ToGraphFolder  -OutVariable item  -Destination 'root:/General'

Start-Process $item.webUrl

#Leave the window open to see export happen -  itempath  tab completes - use the file from the previous -
Get-GraphUserList -MembersOnly | Select-Object Organization | Export-GraphWorkSheet -SheetName sheet1 -ItemPath 'root:/General/test.xlsx' -Show

#groups upgraded to teams have channels for the teams App
Get-GraphTeam $newTeam -Channels -OutVariable teamFirstChannel

$null = New-GraphChannelMessage -Channel $teamFirstChannel -Content "Please keep posts in 'General' to admin and questions about using the group. Use the wiki or OneNote for shared notes."

#create a New channel - with its own notebook section and a planner with 3 buckets & an initial task. Make them tabs in teams.
$newChannel     = New-GraphChannel  -Team $newTeam -Name      $newProjectName -Description "For anything about project $newProjectName"
#The next commnd will fail - want to make a point about that!
$newTeamplan    = New-GraphTeamPlan -Team $newTeam -PlanName  $newProjectName
#The point to make: when you create a team yoy aren't added as a member and that stops you creating the planner so add (current user is in globalVar) and go again
Add-GraphTeamMember -Group $Newteam -Member $GraphUser
$newTeamplan    = New-GraphTeamPlan -Team $newTeam -PlanName  $newProjectName
Add-GraphPlanBucket -Plan    $NewTeamplan -Name 'Backlog', 'To-Do','Not Doing'
Add-GraphPlanTask   -Plan    $newTeamplan -Title "Project $newProjectName Objectives" -Bucket "To-Do" -DueDate ([datetime]::Today.AddDays(7)) -AssignTo $users[-1].Mail

#Add Planner and one note to teams.
Add-GraphPlannerTab -TabLabel 'Planner' -Channel $newChannel  -Plan $NewTeamplan | Out-Null

#Groups have a calendar - add a meeting and invite members
$pattern   = New-GraphRecurrence -Type weekly -DaysOfWeek wednesday -NumberOfOccurrences 52
$attendees = ((Get-GraphTeam -Team $newTeam -Members) + (Get-GraphTeam -Team $newTeam -Owners ) )| New-GraphAttendee -AttendeeType optional
Add-GraphEvent -Team $newTeam -Subject "Midweek team lunch" -Attendees $attendees -Start ([datetime]::Today.AddHours(12)) -End ([datetime]::Today.AddHours(12)) -Recurrence $Pattern

Get-GraphTeam $newTeam -Notebooks -OutVariable teamnotebook
New-GraphOneNoteSection -Notebook $teamNotebook -SectionName $newProjectName -OutVariable NewSection
$NewSection | Set-GraphOneNoteHome
Add-GraphOneNotePage -HTMLPage "<html><head><title>Project $newProjectName</Title></head><body><p>A default home for your notes.</p></body></html>"
Add-GraphOneNoteTab -TabLabel 'Project Notebook' -Channel $newChannel  -Notebook $Newsection

$teamDrive | Add-GraphSharePointTab -TabLabel "Team Drive" -Channel $NewChannel

Get-GraphTeam $newTeam -Site -OutVariable Site
$cols  = 'AssignedTo', 'IssueStatus', 'TaskDueDate', 'V3Comments' | ForEach-Object {Get-GraphSiteColumn -Raw -name $_}
$cols += Get-GraphSiteColumn -Raw -Name 'priority' -ColumnGroup 'Core Task and Issue Columns'
$newlist = New-GraphList -Name "$newProjectName Issue Tracking" -Columns $cols  -Site $site -Template genericList

Add-GraphListItem  -List $newlist -Fields @{Title='Demo Item';IssueStatus='Active';Priority='(2) Normal';}

$newlist | Add-GraphSharePointTab -Channel $NewChannel

New-GraphChannelMessage    -Channel $teamFirstChannel -Content "A new channel has been added for Project $newProjectName with its own planner, one note section and issues list on the team site. Take a look "

Start-Process $newlist.webUrl
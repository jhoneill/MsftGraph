#select some users and put them in a new team. 
$GroupName = 'Accounts'
$newProjectName = "Mccaw" 

$users    = Get-GraphUserList -Filter "Department eq '$GroupName'" 
$users
$newTeam  = New-GraphTeam -Name $GroupName  -Description "The $GroupName Department" -Visibility public -Members $users
$newTeam 
#Teams have a drive, a calendar a notebook and a default channel - let's see them, we'll use them all ...   
$teamDrive        = Get-GraphTeam $newTeam -Drive
$teamDrive
$teamCalendar     = Get-GraphTeam $newTeam -Calendar
$teamCalendar
$teamNotebook     = Get-GraphTeam $newTeam -Notebooks
$teamNotebook
$teamFirstChannel = Get-GraphTeam $newTeam -Channels
$teamFirstChannel


#Groups have a drive - add some files to it  
Get-GraphDrive -Drive $teamdrive -FolderPath /
Get-GraphDrive -Drive $teamdrive -SpecialFolder Documents
Get-GraphDrive -Drive $teamdrive -FolderPath /

dir *.xlsx |  Copy-ToGraphFolder -Drive $teamdrive -Destination 'root:/Documents'

Get-GraphDrive -Drive $teamdrive -SpecialFolder Documents
start $teamdrive.webUrl

#Groups have a calendar - add a meeting and invite members
$Pattern   = New-RecurrencePattern -Weekly -Days Wednesday -Occurrences 52 
$attendees = ((Get-GraphTeam -Team $newTeam -Members) + (Get-GraphTeam -Team $newTeam -Owners ) )| New-EventAttendee -AttendeeType optional
Add-GraphEvent -Calendar $teamCalendar  -Subject "Midweek team lunch" -Attendees $attendees -Start ([datetime]::Today.AddHours(12)) -End ([datetime]::Today.AddHours(12)) -Recurrence $Pattern


#Groups have a note book - add a section and a page. 
$firstChannelSection = New-GraphOneNoteSection -Notebook $teamNotebook -SectionName $teamFirstChannel.displayName 
$firstChannelSection 
Add-GraphOneNotePage -Section $firstChannelSection -HTMLPage '<html><head><title>$($teamFirstChannel.displayName) Section</Title></head><body><p>A default home for your notes.</p></body></html>'

#Groups start with one channel - add a wiki, and general section of the notebook to it. 
Add-GraphWikiTab       -Channel $teamFirstChannel -TabLabel Wiki
Add-GraphOneNoteTab    -Channel $teamFirstChannel -Notebook $firstChannelSection -TabLabel Notes
Add-GraphChannelThread -Channel $teamFirstChannel -Content "Please keep posts in 'General' to admin and questions about using the group. Use the wiki or OneNote for shared notes." 

#New channel - add a notebook section and a planner ,with 3 buckets and an initial task

$Newsection     = New-GraphOneNoteSection -Notebook $teamNotebook -SectionName $newProjectName 
Add-GraphOneNotePage -Section $Newsection -HTMLPage "<html><head><title>Project $newProjectName</Title></head><body><p>A default home for your notes.</p></body></html>"

$newChannel     = New-GraphChannel  -Team $newTeam -Name      $newProjectName -Description "For anything about project $newProjectName" 
$newTeamplan    = New-GraphTeamPlan -Team $newTeam -PlanName  $newProjectName
Add-GraphTeamMember -Group $Newteam -Member j@mobulaconsulting.com 

$newTeamplan    = New-GraphTeamPlan -Team $newTeam -PlanName  $newProjectName
Add-GraphOneNoteTab -Channel $newChannel  -Notebook $Newsection -TabLabel 'Project Notebook'
Add-GraphPlannerTab -Channel $newChannel  -Plan $NewTeamplan    -TabLabel "Planner" 
Add-GraphPlanBucket -Plan    $NewTeamplan -Name 'Backlog', 'To-Do','Not Doing' 
Add-GraphPlanTask   -Plan    $newTeamplan -Title "Project Objectives" -Bucket "To-Do" -DueDate ([datetime]::Today.AddDays(7)) -AssignTo jacob@mobulaconsulting.com 

$cols    = 'AssignedTo', 'IssueStatus',  'TaskDueDate',   'V3Comments' | ForEach-Object {Get-GraphSiteColumn -name $_}
$cols   += Get-GraphSiteColumn -Name 'priority' -ColumnGroup 'Core Task and Issue Columns'
$newlist = New-GraphList -Name "$newProjectName Issue Tracking" -Columns $cols  -Site $site -Template genericList

Add-GraphListItem  -List $newlist -Fields @{Title='Demo Item';IssueStatus='Active';Priority='(2) Normal';}

Add-GraphChannelThread -Channel $teamFirstChannel -Content "A new channel has been added for Project $newProjectName with its own planner, one note section and issues list on the team site. Take a look " 
 
Start $newlist.webUrl




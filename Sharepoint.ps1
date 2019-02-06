function Get-GraphSite                {
    <#
      .Synopsis
        Gets details of a sharepoint site, or its lists, drives or subsites
      .Description
        This interogates https://graph.microsoft.com/v1.0/sites/{id}
        which requires consent to use the Sites.Read.All scope or better.
        If no ID is provided it queries the Root site.
        Depending on the parameters given it will return subsites, lists
        detials of a single list, OneDrive Drives and on Note Notebooks.,
        it
      .Example
        >Get-GraphTeam -site | Get-GraphSite -Lists -Hidden
        Gets the site(s) for the current user's team(s) and gets lists
        from the site(s) including hidden ones.
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #Specifies a site, if omitted "root" will be assumed - the root site of the user's tennant.
        [Parameter( ValueFromPipeline=$true,Position=0)]
        $Site = "root",
        #If specified returns the lists in the site.
        [Parameter(Mandatory=$true, ParameterSetName="Lists")]
        [Switch]$Lists,
        #If specified returns the system lists which are hidden by default
        [Parameter(ParameterSetName="Lists")]
        [Switch]$Hidden,
        #if Specified returns the details of one list
        [Parameter(Mandatory=$true, ParameterSetName="SingleList")]
        [String]$listID,
        #If Specified returns notebooks in the s
        [Parameter(Mandatory=$true, ParameterSetName="Notebooks")]
        [Switch]$Notebooks,
        #If Specified returns the drives in the site.
        [Parameter(Mandatory=$true, ParameterSetName="Drives")]
        [Switch]$Drives,
        #If Specified returns the sub-sites within the site, if the user has suitable permissions.
        [Parameter(Mandatory=$true, ParameterSetName="SubSites")]
        [Switch]$SubSites # Needs higher permissions
    )
    begin {
        Connect-MSGraph
        $webParams = @{Method = 'Get'
                    Headers = $Script:DefaultHeader
        }
    }
    process {
        if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
        foreach ($s in $site) {

            if ($s.id)     {$siteID    = $s.id}     else {$siteid = $s}
            if ($s.weburl) {$ParentURL = $s.weburl} else {
                $ParentURL = (Invoke-RestMethod @webParams -Uri "https://graph.microsoft.com/v1.0/sites/$siteID").webUrl
            }
            if     ($ListID)     { (Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/sites/$SiteID/lists/$ListID/items?expand=fields").value }
            elseif ($Lists)      {

                              $webParams['uri'] =  "https://graph.microsoft.com/v1.0/sites/$SiteID/lists?expand=columns,contenttypes,drive"
                              if ($Hidden) {
                                $webParams['uri'] += '&$select=system,createdDateTime,description,eTag,id,lastModifiedDateTime,name,webUrl,displayName,createdBy,lastModifiedBy,list'
                              }
                              $l = (Invoke-RestMethod @webParams).value
                              foreach ($list in $l) {
                                $list.pstypeNames.add('GraphList')
                                if ($list.drive) {$list.drive.pstypeNames.add('GraphDrive')}
                                Add-Member -InputObject $list -MemberType NoteProperty -Name siteID -Value $siteID
                                if ($ParentURL) {
                                    Add-Member -InputObject $list -MemberType NoteProperty -Name ParentUrl -Value $ParentURL
                                }
                              }
                              return $l
            }
            elseif ($Drives)     {
                              $d = (Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/sites/$SiteID/drives").value
                              foreach ($drive in $d) {$drive.pstypeNames.add("GraphDrive")}
                              return $d
            }
            elseif ($SubSites)   {
                              $SubSitelist = (Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/sites/$SiteID/sites" ).value
                              foreach ($subsite in $SubSitelist) {
                                $subsite.pstypeNames.add("GraphSite")
                                Add-Member -InputObject $subsite -MemberType NoteProperty -Name siteID -Value $siteID
                                if ($ParentURL) {
                                    Add-Member -InputObject $subsite -MemberType NoteProperty -Name ParentUrl -Value $ParentURL
                                }
                              }
                              retrun $SubSitelist
            }
            elseif ($Notebooks)  {
                if ($siteID -eq "root") {
                    $siteID  =     (Invoke-RestMethod @webParams -Uri  "https://graph.microsoft.com/v1.0/sites/root").id
                }
                $books = (Invoke-RestMethod @webParams -uri   "https://graph.microsoft.com/v1.0/sites/$siteID/onenote/notebooks?`$expand=sections").value
                foreach ($b in $books) {
                    foreach ($s in $b.sections) {$s.pstypeNames.add("GraphOneNoteSection")}
                    $b.pstypeNames.add("GraphOneNoteBook")
                }
                return $books
            }
            else            {   $site = Invoke-RestMethod @webParams -Uri ("https://graph.microsoft.com/v1.0/sites/$siteID" + '?expand=drives,lists,sites')
                                $site.lists  | Add-Member -MemberType NoteProperty   -Name SiteID   -Value  $site.id
                                $site.lists  | Add-Member -MemberType ScriptProperty -Name Template -Value {$this.list.template}
                                $site.lists  | ForEach-Object {$_.pstypeNames.add("GraphList")}
                                $site.drives | ForEach-Object {$_.pstypeNames.add("GraphDrive")}
                                $site.sites  | ForEach-Object {$_.pstypeNames.add("GraphSite")}
                                $site.pstypeNames.add("GraphSite")
                                return $site
            }
        }
    }
    #https://graph.microsoft.com/v1.0/sites?search=contoso
    #https://graph.microsoft.com/v1.0/sites/root/columns
    #https://graph.microsoft.com/v1.0/sites/root/contentTypes
}

function Get-GraphList                {
    <#
      .Synopsis
        Gets sharepoint list objects or their items
      .Description
        This interogates https://graph.microsoft.com/v1.0/sites/{id}/lists{id}
        which requires consent to use the Sites.Read.All scope or better.
        This does not suppor the use of a filter parameter so any "where"
        operation has to be done on the returned data.
      .Example
        >
        >$myTeamSite = Get-GraphTeam -Site | select -first 1
        >$problemsList = $myteamsite.lists | where name -like problem*
        >
        > Get-GraphList -site $myTeamSite.id -list $problemslist.id -ColumnList

        The first command gets the current users group(s) and returns their site(s).
        For this example we select the first site. The sites returned by Get-GraphGroup /
        Get-GraphTeam have a .lists property and second command selects the list we want
        The third line shows calling Get-GraphList using the ID for both Site and List
        and  getting the columns in the list.
        The next example shows an easier way to provide the information; and in fact
        there is already a .columns property of $problemsList which has the column information
      .Example
        > Get-graphlist $problemsList -Items
        This uses $problemsList from the previous example. Get-GraphGroup (aka Get-GraphTeam)
        gets the Site, it gets the sites lists, and adds the site ID as a property, so
        $Problemslist has propeties for the list ID and the site ID. So this exmaple uses a
        shorter form of just providing the list and returns the items in their raw state
      .Example
        > Get-graphlist $problemsList -Items -Property title, issuestatus, AssignedToLookupID, priority
        This builds on the previous example. Specifying -Property causes Get-GraphList to
        return the Item(s) Fields collection(s) and sets the default fields to be displayed.
        By default if an object has 4 visbible properties or fewer PowerShell displays it
        as a table, if it has more than 4 a list is used, this can be managed with
        $FormatEnumerationLimit. In this case 4 properties are show in a table view.
        However 'Person or Group' fields, like AssignedTo return a lookupID.
        This comes from the hidden list 'Users' and the next example shows how to get
        information from this list. (The Get-GraphSiteUserList provides a shortcut for geting
        this Information)
      .Example
        >
        >Get-GraphList -Site $myteamSite -Hidden  | where name -eq 'users' |
            Get-Graphlist -Items -Property id,ContentType,Title,Name

        This uses the $myTeamSite variable from the first example.
        If neither Items, nor ColumnList is specified, Get-GraphList returns list objects,
        (the same result as using Get-GraphSite -Lists) so the first command gets lists
        in the team site including hidden ones - which aren't included in the .lists
        property of the site, and users IS hidden. The where command isolates that list,
        and it is piped into a second Get-GraphList command, which gets its items
        and displays the properties of interest
      .Example
        >
        >$mydocuments = Get-GraphUser -Site | Get-GraphSite -lists | where name -eq documents
        >Get-GraphList $shareddocsList -items | Select -expand driveItem |
              Copy-FromGraphFolder -Destination C:\temp

        This command works with a users "MySite" - the first command gets the user's
        site, gets its lists and selects the one named "Documents"
        The second gets the items in this list; when a list object has an associated drive,
        items returned by Get-GraphList -items will have a .DriveItem property.
        Driveitems can be piped into  Copy-FromGraphFolder .
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    param(
        #The list either as an ID or as a list object (which may contain the site.)
        [parameter(ValueFromPipeline=$true, ParameterSetName="ListID",Position=0)]
        [parameter(ValueFromPipeline=$true, ParameterSetName="ListItems",Position=0)]
        [parameter(ValueFromPipeline=$true, ParameterSetName="ListIDColumns",Position=0)]
        $List ,
        #Specifies a site, if omitted "root" will be assumed - the root site of the user's tennant.
        $Site = "root",
        #If specified returns hidden lists (like 'Users')
        [Parameter(ParameterSetName="ListofLists",Mandatory=$true)]
        [Switch]$Hidden,
        #If specified returns the list's items
        [Parameter(ParameterSetName="ListItems",Mandatory=$true)]
        [Switch]$Items,
        #If specified returns the columns in the list
        [parameter(ParameterSetName="ListIDColumns", Mandatory=$true)]
        [Switch]$ColumnList,
        #if specified returned items will be expanded and the default display fields will be set
        [Parameter(ParameterSetName="ListItems")]
        [Alias('Fields')]
        [String[]]$Property

    )

    if     ($Site.id)     {$siteID = $Site.ID}
    elseif ($List.siteID) {$siteID = $List.siteID}
    else                  {$siteID = $Site}  #Site has a default, so won't be empty
    if     ($List.id)     {$listid = $List.ID}
    elseif ($List)        {$listid = $List    #Don't set listID if List is empty (it has no default)
                           if (-not $PSBoundParameters.ContainsKey('Site')) { #If we got a list ID and no site it's probably not the root!
                               Write-Warning -Message 'Assuming root site. If a 404 "not found" error occurs specify the site explicitly.'
                           }
    }

    Connect-MSGraph
    $webParams = @{Method = "Get"
                  Headers = $Script:DefaultHeader
    }
    if     ($Items) {
        $uri = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listid/items?expand=fields"
        if ($List.drive) { $uri += ',driveItem' } #trying to expand driveItem in drive-less lists causes an error.
        Write-Progress -Activity 'Getting list items'
        $listitems = (Invoke-RestMethod @webParams -uri $uri).value
        Write-Progress -Activity 'Getting list items' -Completed
        if ($Property) {
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$Property)
            $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            foreach ($i in $listitems) {
                Add-Member  -InputObject $i.fields -MemberType MemberSet    -Name PSStandardMembers -Value $PSStandardMembers -PassThru |
                Add-Member  -MemberType NoteProperty -Name SiteID -Value $siteID -PassThru |
                Add-Member  -MemberType NoteProperty -Name Listid -Value $listid -PassThru |
                Add-Member  -MemberType NoteProperty -Name Itemid -Value $i.id   -PassThru
            }
            return
        }
        else {
            foreach ($i in $listitems) {
                if ($i.driveItem) {$i.driveitem.pstypeNames.add('GraphDriveItem')}
                Add-Member  -InputObject $i -MemberType NoteProperty -Name siteID -Value $siteID
                Add-Member  -InputObject $i -MemberType NoteProperty -Name ListID -Value $listid
            }
            return $listitems
        }
    }
    #If we were asked for listItems we will have returned in the previous if. From here on we return objects for one or more list(s)
    elseif ($List)  {
        Write-Progress -Activity 'Getting list details'
        $result =  Invoke-RestMethod @webParams -uri "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listid`?expand=columns,contenttypes,drive,items(expand=fields)"
        Write-Progress -Activity 'Getting list details' -Completed
    }
    else   {
        Write-Progress -Activity 'Getting Site Lists'
        $webParams['uri'] =  "https://graph.microsoft.com/v1.0/sites/$siteID/lists?expand=columns,contenttypes,drive"
        if ($Hidden) {
          $webParams['uri'] += '&$select=system,createdDateTime,description,eTag,id,lastModifiedDateTime,name,webUrl,displayName,createdBy,lastModifiedBy,list'
        }
        $result = (Invoke-RestMethod @webParams).Value
        Write-Progress -Activity 'Getting Site Lists' -Completed
    }
    if     ($ColumnList) {
        $result = $result | Select-Object -ExpandProperty Columns
        foreach ($r in $result) {
            $r.pstypeNames.add('GraphColumn')
            Add-Member -InputObject $r -MemberType NoteProperty -Name SiteID -Value $Siteid
            Add-Member -InputObject $r -MemberType NoteProperty -Name ListID -Value $listid
        }
        return $result
    }
    else   {
        ForEach ($r in $result) {
            $r.pstypeNames.add('GraphList')
            if ($r.drive) {$r.drive.pstypeNames.add('GraphDrive')}
            Add-Member -InputObject $r  -MemberType NoteProperty -Name SiteID -Value $siteID
        }
        return $result
    }
}

<#
Update list item      PATCH https://graph.microsoft.com/v1.0/sites/{site-id}/lists/{list-id}/items/{item-id}   https://docs.microsoft.com/en-us/graph/api/listitem-update?view=graph-rest-1.0

Versions
GET /sites/{site-id}/items/{item-id}/versions
GET /sites/{site-id}/lists/{list-id}/items/{item-id}/versions
#>

function New-GraphList                {
    <#
      .Synopsis
        Creates a new sharepoint list
      .Description
        This posts a new item to https://graph.microsoft.com/v1.0/sites/{id}/lists
        which requires consent to use the Sites.Manage.All scope.
        The API allows lists to be created - but not with content types, only as a defined
        collection of columns. There is no PATCH or DELETE support so there there are is
        no Set- or Remove- function to go with the New-
      .Example
        >$site            = Get-GraphUser -Teams -Name Consultants | Get-GraphTeam -site
        >$textcolumndef   = New-GraphTextColumn -TextType plain
        >$column1         = New-GraphColumn -Name Author -ColumnDefinition $textcolumndef
        >$numberColumnDef = New-GraphNumberColumn
        >$column2         = New-GraphColumn -Name PageCount -ColumnDefinition $numberColumnDef
        >$booksList       = New-GraphList   -Name NewBooks  -Columns $column1,$column2 -Site $site -Template genericList
        >Start $bookslist.weburl

        This builds the example at https://docs.microsoft.com/en-us/graph/api/list-create?view=graph-rest-1.0
        The first line gets the sharepoint site for a team the current user is a member of named 'consultants'
        The second creates the column settings for a text column and the thrid builds a named column
        definition with that specification. Lines 4 and 5 define a Number column
        And line 6 creates a new generic list on the site found in line 1; the list is named 'books'
        and has columns title (as all generic list items do), Author and Page count (the latter two being
        defined in lines 2-5). The final line opens the new list in a browser.
      .Example
        >New-GraphList -Name books -Columns (ListColumn author (TextColumn)),(ListColumn pagecount (NumberColumn))  -Site $site
        This example does the same task as the previous one but leaves out some default parameters :
        text columns default to plain text and genericList is the default template.
        The columns are specifed using Aliases to give sort of DSL: "(listcolumn author (textColumn))" is eqivalent to
        (New-GraphColumn -Name Author -ColumnDefinition (New-TextColumn)  )
      .Example
        >
        >$cols = 'AssignedTo', 'IssueStatus',  'TaskDueDate',   'V3Comments' | foreach {Get-GraphSiteColumn -name $_}
        >$cols += Get-GraphSiteColumn -Name 'priority' -ColumnGroup 'Core Task and Issue Columns'
        >New-GraphList -Name 'Problem Tracking' -Columns $cols  -Site $site -Template genericList

        This gets a set of pre-existing columns and uses them to create a new list.
        The first line gets columns with unique names. "Priority" is defined in multiple groups,
        so the second line qualifies which version ot wants
        And the third line uses the columns to create a list.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #The Name of the list
        [parameter(Mandatory=$true,Position=0)]
        [Alias('Name')]
        $DisplayName,
        #The site where the list is to be added either as an ID or as a site object containing an ID
        $Site,
        #The base list template used in creating the list. This is validated against typical list names, though you made need to modify the code to add others
        [ValidateSet('documentLibrary', 'genericList', 'tasks', 'survey', 'links', 'announcements', 'contacts')]
        [string]$Template = 'genericList',
        #A longer description for the list
        $Description ,
        #Columns to use in the list
        $Columns,
        #Create the list hidden
        [switch]$Hidden
    )
    if     ($Site.ID)           {$siteID = $Site.id}
    elseif ($site -is [string]) {$siteID = $site }
    else   {Write-Warning -Message 'Could not determine the site ID'; Return} 
    Connect-MSGraph
    $WebParams = @{ 'URI'         = "https://graph.microsoft.com/v1.0/sites/$siteID/lists"
                    'Method'      = 'Post'
                    'Headers'     =  $DefaultHeader  
                    'ContentType' = 'application/json'
    }
    $settings  = @{
        'displayName'             = $DisplayName;
        'list'   = @{
            'contentTypesEnabled' = [bool]$ContentTypes;
            'hidden'              = [bool]$Hidden
            'template'            = $Template
        }
    }
    if ($Description) {
        $settings['description']  = $Description
    }
    if ($Columns) {
        $settings['columns']      = @() + $Columns
    }
    if ($ContentTypes) {
        $settings['contentTypes'] = @()
        $i = 1
        foreach ($ct in $contentTypes) {
            $ct['order'] = @{'position' = $i ; 'default' = ($i -eq 1) }
            $settings['contentTypes'] += $ct
         }
    }
    $json = ConvertTo-Json $settings  -Depth 10
    Write-Debug $Json
    if ($Force -or $PSCmdlet.ShouldProcess($DisplayName,"Add list to site $($site.name)")) {
        $result = Invoke-RestMethod @WebParams -body $json 
        Add-Member -InputObject $result -MemberType NoteProperty -Name SiteID -Value $siteID
        $result.pstypeNames.add('GraphList')
        return $result
    }
}

function Add-GraphListItem            {
    <#
      .Synopsis
        Adds an item to a SharePoint List
      .Description
        This posts a new item to https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items
        which requires consent to use the Sites.ReadWrite.All scope
        Posting to a list is quite basic - it is a set of Name-ValuePairs and
        FIELD NAMES ARE CASE SENSITIVE. If you get a 400 error from the server the
        first thing to check is the names of the fields. It does not appear to be possible to
        post certain types of field - lookup and Person/Group being the major issues.
        The command will try to post what it is given, but it makes no attempt at validating it!
     .Example
     >
     >$myteamsite = Get-GraphTeam -Site |select -first 1
     >$problemslist = $myteamsite.lists.where({$_.name -like "problem*"})
     >Add-GraphListItem  -List $problemslist -Fields @{Title='Demo Item';IssueStatus='Active';Priority='(2) Normal';}

     The first command gets a team site which has a list named "Problem reports"
     The second line gets that list
     The third creates a list item with Title, IssueStatus and Priority fields.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        #The list to add to; this can be an ID, or list object with an ID, and a site ID
        [parameter(Mandatory=$true,Position=0)]
        $List,
        #The item property values in a hash table as @{col1=$value1; col2='Value2'; col3=33}
        [parameter(Mandatory=$true,Position=1)]
        [hashtable]$Fields,
        #If the list parameter does not contain a .SiteID property allows the site to specified as an ID or object
        $Site,
        #If specified the new item will be returned, otherwise it is created silently.
        [Alias('PT')]
        [switch]$Passthru,
        #If specified the item will be added without prompting for confirmation (this is the default unless confirm preference is changed)
        [switch]$Force
    )
    if     ($Site.ID)     {$siteid = $Site.ID}
    elseif ($Site)        {$siteid = $Site}
    elseif ($List.siteid) {$siteid = $List.siteid}
    else   {throw 'The site could not be determined from the list; please specify the site explicitly.' ; return}
    if     ($List.id)     {$listid = $List.ID}  else {$listID = $List}
    Connect-MSGraph
    $webParams = @{
            'Method'      =  'Post'
            'Headers'     =   $Script:DefaultHeader
            'URI'         =  "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items"
            'ContentType' = 'application/json'
    }
    $Settings = @{'fields'=$Fields}
    $json = ConvertTo-Json $settings
    Write-Debug $Json
    if ($Force -or $PSCmdlet.ShouldProcess($List.name,'Add item')) {
        $result = Invoke-RestMethod @webParams -Body $json
        if ($Passthru) {
            $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet',[string[]]$Fields.Keys)
            $psStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
            Add-Member -InputObject $result.fields -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers -PassThru
        }
    }
}

function Set-GraphListItem            {
    <#
      .Synopsis
        Updates an item in a SharePoint List
      .Description
        This Patches an existing item at https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items{id}/Fields
        which requires consent to use the Sites.ReadWrite.All scope
        Caveats in Add-GraphListItem apply to Set-GraphListItem.
      .link
        Add-GraphListItem
      .Example
      >
      >$problemitems = get-graphlist $problemslist -Items -Property title,issuestatus,AssignedToLookupID,priority
      >$problemitems[2] | Set-GraphListItem -Fields @{Priority='(2) Normal'}

      The first line gets the items from a list , and the second updates the Priority field of the third one
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        #The item to update; this can be an ID or an object with an ID, and a list and site ID as well
        [parameter(ValueFromPipeline=$true, Mandatory=$true,Position=0)]
        $Item,
        #The item property values in a hash table as @{col1=$value1; col2='Value2'; col3=33}
        [parameter(Mandatory=$true,Position=1)]
        [hashtable]$Fields,
        #If the item does not contain the list, the list to delete from an ID, or list object with an ID, and a site ID
        $List,
        #If there is no site id in the item or list parameter allows the site to specified as an ID or object
        $Site,
        #If specified, passes through the server's confirmation of the update
        [Alias('PT')]
        [switch]$Passthru,
        #If specified the item will be updated without prompting for confirmation
        [switch]$Force
    )
    if     ($item.SiteId) {$siteid = $item.SiteID}
    elseif ($List.siteid) {$siteid = $List.siteid}
    elseif ($Site.id)     {$siteid = $Site.id}
    elseif ($Site)        {$siteid = $Site}
    else   {throw 'The site could not be determined; please specify the site explicitly.' ; return}
    if     ($item.Listid) {$listid = $Item.ListID}
    elseif ($List.id)     {$listid = $List.ID}
    elseif ($listid)      {$listID = $List}
    else   {throw 'The List could not be determined; please specify the list explicity' ; return}
    if     ($Item.id)     {$item   = $Item.id}
    Connect-MSGraph
    $webParams = @{
            'Method'      =  'Patch'
            'Headers'     =   $Script:DefaultHeader
            'URI'         =  "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items/$Item/fields"
            'ContentType' = 'application/json'
    }

    $json = ConvertTo-Json $Fields
    Write-Debug $Json
    if ($Force -or $PSCmdlet.ShouldProcess($List.name,'Update item')) {
        $result = Invoke-RestMethod @webParams -Body $json
        if ($Passthru) { return $result}
    }
}

function Remove-GraphListItem         {
    <#
      .Synopsis
        Deletes an item from a SharePoint List
      .Description
        This Deletes an item at https://graph.microsoft.com/v1.0/sites/{id}/lists{id}/items{id}
        which requires consent to use the Sites.ReadWrite.All scope
      .Example
        >
        >$problemitems = get-graphlist $problemslist -Items -Property title,issuestatus,AssignedToLookupID,priority
        >$problemitems[4] | Remove-GraphListItem

          The first line gets the items from a list , and the second line removes the fifth one
    #>

    [cmdletbinding(SupportsShouldProcess=$true,ConfirmImpact='High')]
    param(
        #The item to remove; this can be an ID or an object with an ID, and a list and site ID as well
        [parameter(ValueFromPipeline=$true, Mandatory=$true,Position=0)]
        $Item,
        #If the item does not contain the list, the list to delete from an ID, or list object with an ID, and a site ID
        $List,
        #If there is no site id in the item or list parameter allows the site to specified as an ID or object
        $Site,
        #If specified the item will be deleted without prompting for confirmation (prompting is the default)
        [switch]$Force
    )
    if     ($item.SiteId) {$siteid = $item.SiteID}
    elseif ($List.siteid) {$siteid = $List.siteid}
    elseif ($Site.id)     {$siteid = $Site.id}
    elseif ($Site)        {$siteid = $Site}
    else   {throw 'The site could not be determined; please specify the site explicitly.' ; return}
    if     ($item.Listid) {$listid = $Item.ListID}
    elseif ($List.id)     {$listid = $List.ID}
    elseif ($listid)      {$listID = $List}
    else   {throw 'The List could not be determined; please specify the list explicity' ; return}
    if     ($Item.id)     {$item   = $Item.id}
    Connect-MSGraph
    $webParams = @{
            'Method'      =  'DELETE'
            'Headers'     =   $Script:DefaultHeader
            'URI'         =  "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items/$Item"
            'ContentType' = 'application/json'
    }
    if ($force -or $PSCmdlet.ShouldProcess($item,'Delete List Item') ){ Invoke-RestMethod @webParams }
}

function New-GraphColumn              {
    <#
      .synopsis
        Create a new Column definition for a sharepoint list
      .description
        New-GraphList uses column definitions to set up a new list.
        Each column has a name, description, default and one of the properties from the following list
        boolean, calculated, choice, currency, dateTime, lookup, number, personOrGroup or text
        Flags can also be set to say if the column is indexed, Readonly and/or required.
        Existing Columns defined in the site can be fetched with Get-GraphSiteColumn
        New-GraphColumn defines a new column to be included in a list, and a typical list will need
        multiple columns, which may be a mixture of new and existing columns.
        The specifics of each of the column types is handled by a new-{typeName}Column command.
        Examples appear in New-GraphList
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/columndefinition?view=graph-rest-1.0
      .link
        New-GraphList
    #>
    [CmdletBinding(DefaultParameterSetName='None')]
    [Alias('ListColumn')]
    param (
        [Parameter(Mandatory=$true,Position=0)]
        # The API-facing name of the column as it appears in the fields on a listItem. For the user-facing name, see displayName.
        [string]$Name,
        #A definition created with on of the New-*Column commands for a text, currency, boolean etc
        [Parameter(Mandatory=$true,Position=1)]
        [hashtable]$ColumnDefinition ,
        #For site columns, the name of the group this column belongs to. Helps organize related columns.
        [string]$ColumnGroup,
        #The user-facing description of the column.
        [string]$Description,
        #The user-facing name of the column.
        [string]$DisplayName,
        #Fills in the default value using a formula
        [Parameter(ParameterSetName='DefaultbyFormula',Mandatory=$true)]
        [string]$DefaultValueFormula,
        #Fills in the defaultt value using a fixed value
        [Parameter(ParameterSetName='DefaultbyValue',Mandatory=$true)]
        [string]$DefaultValueString ,
        #If specified the column is indexed to help the perfomance of searching and grouping.
        [bool]$Indexed,
        #Specifies whether the column values can be modified.
        [bool]$ReadOnly,
        #  Specifies whether the column value is not optional.
        [bool]$Required,
        #If true, no two list items may have the same value for this column.
        [bool]$EnforceUniqueValues,
        # Specifies whether the column is displayed in the user interface.
        [boolean]$Hidden
    )
    $Settings = $ColumnDefinition + @{
                           'name'                = $Name ;
                           'indexed'             = [bool]$Indexed
                           'readOnly'            = [bool]$ReadOnly
                           'required'            = [bool]$Required
                           'enforceUniqueValues' = [bool]$EnforceUniqueValues
                           'hidden'              = [bool]$Hidden
    }

    if ($ColumnGroup) { $settings['columnGroup'] = $ColumnGroup }
    if ($Description) { $settings['description'] = $Description }
    if ($DisplayName) { $settings['displayName'] = $DisplayName }
    else              { $settings['displayName'] = $Name }
    if ($DefaultValueFormula) {
                        $settings['defaultValue']= @{'formula' = $DefaultValueFormula}
    }
    elseif ($DefaultValueString) {
                        $settings['defaultValue']= @{'value' = $DefaultValueString}
    }

    return $Settings
}

#region create all the column definitions used in a column
function New-GraphBooleanColumn       {
    <#
      .synopsis
        Creates a definition of a Sharepoint calculated column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/calculatedcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('BooleanColumn')]
    param (
    )
    return @{'boolean' = @{} }
}

function New-GraphCalculatedColumn    {
    <#
      .synopsis
        Creates a definition of a Sharepoint calculated column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/calculatedcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('CalculatedColumn')]
    param (
        #The formula used to calculate the value.
        $Formula ,
        #Should the value be presented as a date only or a date and time
        [ValidateSet( 'dateOnly', 'dateTime')]
        $Format = 'dateTime',
        # Should the result be treated as Number, text, date, Currency or boolean
        [ValidateSet( 'boolean','currency','dateTime', 'number', 'text')]
        $OutputType = 'text'
    )
    $columnSettings = @{
        'formula'                        = $Formula
        'ouputType'                      = $OutputType
    }
    if ($OutputType -eq 'dateTime') {
        $columnSettings['format']        = $Format
    }
    return @{'calculated' = $columnSettings}
}

function New-GraphChoiceColumn        {
    <#
      .synopsis
        Creates a definition of a Sharepoint choice column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('ChoiceColumn')]
    param (
        #The list of values available for this column..
        [Parameter(Mandatory=$true,Position=0)]
        [string[]]$Choices,
         #How the choices are to be presented in the UX, defaults to dropdown menu
        [ValidateSet('checkBoxes', 'dropDownMenu', 'radioButtons')]
        [string]$DisplayAs ='dropDownMenu',
        #Specified to indicates that values in the column should be able to exceed the standard limit of 255 characters.
        [switch]$AllowTextEntry
    )
    return @{'choice' =  @{
        'allowTextEntry'                 = [bool]$AllowTextEntry
        'choices'                        = @() + $Choices
        'displayAs'                      = $DisplayAs
    }}
}

function New-GraphCurrencyColumn      {
    <#
      .synopsis
        Creates a definition of a Sharepoint datetime column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/datetimecolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('CurrencyColumn')]
    param (
        $Locale = (Get-Culture)
    )

    if ($Locale -is [System.Globalization.CultureInfo] ) {
        return @{'currency' =  @{'locale' = $Locale.name }}
    }
    elseif ( [System.Globalization.CultureInfo]::GetCultureInfo($Locale).displayname -Notmatch 'Unknown' ) {
        return @{'currency' =   @{'locale' = $Locale}}
    }
    else {throw "$locale is not a known language"}
}

function New-GraphDateTimeColumn      {
    <#
      .synopsis
        Creates a definition of a Sharepoint datetime column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/datetimecolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('DateTimeColumn')]
    param (
        #Should the value be presented as a date only or a date and time
        [ValidateSet( 'dateOnly', 'dateTime')]
        $Format = 'dateTime',
        # Should the UX use default rendering or relative representation (eg. "today at 3:00 PM") or the standard absolute representation (eg. "5/10/2017 3:20 PM")
        [ValidateSet( 'default', 'friendly', 'standard')]
        $DisplayAs = 'default'
    )
    return @{'datetime' = @{
        'displayAs'                      = $DisplayAs
        'format'                         = $Format
      }  }
}

function New-GraphLookupColumn        {
    <#
      .synopsis
        Creates a definition of a Sharepoint lookup column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/lookupcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('LookupColumn')]
    param (
        #The unique identifier of the lookup source list.
        [Parameter(Mandatory=$true,Position=0)]
        [string]$ListId,
        #The name of the lookup source column.
        [Parameter(Mandatory=$true,Position=0)]
        [string]$ColumnName,
        #If specified, this column is a secondary lookup, pulling an additional field from the list item looked up by the primary lookup. Use the list item looked up by the primary as the source for the column named here
        [string]$PrimaryLookupColumnId,
        #If Specified allows multiple/values to be specified
        [switch]$MultipleSelection,
        #Specified to indicates that values in the column should be able to exceed the standard limit of 255 characters.
        [switch]$AllowUnlimitedLength
    )
    $columnSettings = @{
        'allowMultipleValues'            = [bool]$MultipleSelection
        'allowUnlimitedLength'           = $AllowUnlimitedLength
        'columnName'                     = $ColumnName
        'listId'                         = $ListId
    }
    if ($IncludeGroups) {
          $columnSettings['primaryLookupColumnId'] = $PrimaryLookupColumnId
    }
    return @{'lookup' = $columnSettings}
}

function New-GraphNumberColumn        {
    <#
      .synopsis
        Creates a definition of a Sharepoint number column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/numbercolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('NumberColumn')]
    param (
        #How the value should be presented in the UX, number by default, the only other choice is percentage
        [ValidateSet('number', 'percentage')]
        $DisplayAs = 'number',
        #How many decimal places to display Auto, None, or the numbers one to five in words
        [ValidateSet('automatic', 'none', 'one', 'two', 'three', 'four', 'five')]
        [string]$DecimalPlaces = 'automatic',
        #Maximum permitted value
        [double]$Max,
        #Maximum permitted value
        [double]$Min
    )
    $columnSettings = @{
        'decimalPlaces'                 = $DecimalPlaces
        'displayAs'                     = $DisplayAs
    }
    if ($Max) {
        $columnSettings['maximum']      = $Max
    }
    if ($Min) {
        $columnSettings['minimum']      = $min
  }
    return @{'number' = $columnSettings}
}

function New-GraphPersonOrGroupColumn {
    <#
      .synopsis
        Creates a definition of a Sharepoint person or group column
      .link
        https://docs.microsoft.com/en-us/graph/api/resources/personorgroupcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('PersonColumn')]
    param (
        #If Specified allows multiple/users to be specified
        [switch]$MultipleSelection,
        #Chooses how the name should be displayed; the default is to show name and presence, but it can first name, title, mail etc.
        [ValidateSet('Account', 'department  firstName', 'id', 'lastName', 'mobilePhone', 'name', 'nameWithPictureAndDetails', 'nameWithPresence',
                      'office', 'pictureOnly36x36', 'pictureOnly48x48', 'pictureOnly72x72', 'sipAddress', 'title', 'userName', 'workEmail', 'workPhone.')]
        [string]$DisplayAs = 'nameWithPresence',
        #If Specified allows groups to be selected as well as users
        [switch]$IncludeGroups
    )
    $columnSettings = @{
        'allowMultipleSelection'            = [bool]$MultipleSelection
        'displayAs'                         = $DisplayAs
    }
    if ($IncludeGroups) {
          $columnSettings['chooseFromType'] = 'peopleAndGroups'
    }
    else {$columnSettings['chooseFromType'] = 'peopleOnly' }

    return @{'personOrGroup' = $columnSettings}
}

function New-GraphTextColumn          {
    <#
      .Synopsis
        Creates a definition of a sharepoint text column
      .link
      https://docs.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
    #>
    [CmdletBinding()]
    [Alias('TextColumn')]
    param (
        #Text is single line unless multiline is specified.
        [Switch]$MultiLine,
        #A new entry replaces exisitng text unless append is specified
        [Switch]$Append,
        #The type of text being stored - plain or richText (plain by default)
        [ValidateSet('plain','richText')]
        [string]$TextType = 'plain' ,
        #The maximum number of characters for the value.
        [int32]$MaxLength,
        #The size of the text box.
        [int32]$LinesForEditing
    )
    $columnSettings = @{
        'allowMultipleLines'                = [bool]$MultiLine
        'appendChangesToExistingText'       = [bool]$Append
        'textType'                          = $TextType
    }
    if ($MaxLength) {
        $columnSettings['maxLength']       = $MaxLength
    }
    if ($TextboxSize) {
        $columnSettings['linesForEditing'] = $TextboxSize
    }
    return  @{'text' = $columnSettings}
}
#endregion

function New-GraphContentType  {
    [cmdletbinding()]
    param (
        #The ID of the contenttype
        [parameter(Mandatory=$true)]
        [string]$ParentID,
        #The name of the content type that the list will display
        [parameter(Mandatory=$true)]
        [string]$Name,
        # the content type cannot be modified unless this value is first set to false
        [Switch]$ReadOnly,
        # If Specified he content type cannot be modified by users or through push-down operations. Only site collection administrators can seal or unseal content types.
        [Switch]$Sealed
    )
    @{  'name'     = $Name
        'id'       = $ParentID
        'readOnly' = [bool]$ReadOnly
        'sealed'   = [bool]$Sealed
    }
}

function Get-GraphSiteColumn   {
    <#
      .synopsis
        Gets a column which is defined for the whole site.
    #>
    [cmdletbinding(DefaultParameterSetName='None')]
    param (
    #Selects column(s) by name (and possibly group)
    [Parameter(ParameterSetName='Terms',Position=0, ValueFromPipeline=$true)]
    [String]$Name,
    #Selects column(s) by group (and possibly by name)
    [Parameter(ParameterSetName='Terms',Position=1)]
    [String]$ColumnGroup,
    #Selects a column by unique ID
    [Parameter(ParameterSetName='Terms',Position=2)]
    [string]$ID,
    #Allows a custom where clause instead of Name and/or group and/or ID
    [Parameter(ParameterSetName='WhereClause')]
    [scriptblock]$Where,
    <#Get-GraphSiteColumn is intended to return one column to used when creating a new list, so
    if multiple columns are returned that would be an error (i.e. two columns have the
    same name and group wasn't given.) If -allowMultiple is specified it is *not* treated as an error #>
    [switch]$AllowMultiple
    )
    begin {
        Connect-MSGraph
        if (-not $script:RootSiteColumns) {
            Write-Progress -Activity "Getting list of columns for the root site" 
            $script:RootSiteColumns = (Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/sites/root/columns" -Headers $DefaultHeader).value
            Write-Progress -Activity "Getting list of columns for the root site" -Completed
        }
    }
    process {
        if (-not $Where) {
            $whereText = ''
                if ($id)           {$whereText += "`$_.id -eq '$id' -and" }
                if ($ColumnGroup)  {$whereText += "`$_.ColumnGroup -like '$ColumnGroup' -and"}
                if ($Name)         {$whereText += "(`$_.Name -like '$Name' -or `$_.displayname -like '$Name')"}
                $wheretext = $whereText -replace ' -and$',''
                $where     = [scriptblock]::Create("$whereText")
        }
        $result = $script:RootSiteColumns.where($where)
        if ($result.count -gt 1 -and -not $AllowMultiple) {throw 'More than one result was found and -AllowMultiple was not specified'}
        else {return $result}
    }
}

function Get-GraphSiteUserList {
    <#
      .Synopsis
        Gets the Users list for a [team] site
    #>
    [cmdletbinding()]
    param (
        #The [team] Site whose user-list will be fetched
        [parameter(ValueFromPipeline=$true,Position=0,Mandatory=$True)]
        $Site
    )
    Connect-MSGraph
    #If we get a list where it should be a site, but it has a site ID ... use that.
    if ($site.siteid) {$site=$site.Siteid}
    Write-Progress -Activity 'Getting Site Users' -CurrentOperation 'Finding list'
    $list = Get-GraphList -Site $Site -Hidden  | Where-Object name -eq 'users'
    Write-Progress -Activity 'Getting Site Users' -CurrentOperation 'Getting users'
    $usersAndGroups = Get-Graphlist -List $list -Items -Property id,contenttype,title,name
    Write-Progress -Activity 'Getting Site Users' -Completed
    $usersAndGroups.where({$_.contentType -eq 'person' -and  $_.name -match "@"}) |
        Select-Object -Property id, Title, @{n='Account';e={$_.name -replace '^.*\|',''}}, Name
}

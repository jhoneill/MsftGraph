using namespace System.Management.Automation
using namespace Microsoft.Graph.PowerShell.Models

#To do allow Copy-GraphOneNotePage to specify a destination-section by name (with a completer - which requires a destination notebook parameter which may cause a conflict with page which doesn't have the notebook)
function Set-GraphOneNoteHome   {
    <#
      .synopsis
        Sets a default notebook (and optionally section). Set to $Null to clear the setting
      .example
        >Get-GraphGroup 'Consultants' -Notebooks | Get-GraphOneNoteBook -SectionName general*  | Set-GraphOneNoteHome -Verbose
        The first command in the pipeline gets the notebook for the consultants group ,
        the second finds the section in the notebook with an display name beginning "general"
        and the third sets the default section for Add-FileToGraphOneNote, Add-GraphOneNotePage,
        Get-GraphOneNotePage, and Out-GraphOneNote to the this section, and sets the
        default Notebook for All the GraphOneNoteBook and all the GraphOneNoteSection commands
        to the consultants group's notebook.
    #>

    param    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [AllowNull()]
        #A note book or notebook section to set as the default location for oneNoteCommands. Passing Null will clear the default.
        $Notebook
    )
    @("*GraphOneNoteBook*:Notebook", '*GraphOneNoteSection*:Notebook', "Add-FileToGraphOneNote:Section",
    "Add-GraphOneNotePage:Section", "Get-GraphOneNotePage:Section", "Out-GraphOneNote:Section") | ForEach-Object {
       $null = $Global:PSDefaultParameterValues.Remove($_)
    }

    if ($notebook -is [MicrosoftGraphOnenoteSection]) {
        Write-Verbose "Setting Default Section to $($notebook.displayname) ... "
       $Global:PSDefaultParameterValues["Add-FileToGraphOneNote:Section"] = $Notebook
       $Global:PSDefaultParameterValues["Add-GraphOneNotePage:Section"]   = $Notebook
       $Global:PSDefaultParameterValues["Get-GraphOneNotePage:Section"]   = $Notebook
       $Global:PSDefaultParameterValues["Out-GraphOneNote:Section"]       = $Notebook
       $Notebook = $Notebook.parentNotebook
    }
   if ($Notebook -is [Microsoft.Graph.PowerShell.Models.MicrosoftGraphNotebook]) {
       Write-Verbose "Setting Default Notebook to $($notebook.displayname)"
       $Global:PSDefaultParameterValues["*GraphOneNoteBook*:Notebook"]    = $Notebook
       $Global:PSDefaultParameterValues['*GraphOneNoteSection*:Notebook'] = $Notebook
       $Global:PSDefaultParameterValues["*GraphOneNotePage*:Notebook"]    = $Notebook
   }
}

function Get-GraphOneNoteBook   {
    <#
      .Synopsis
        Gets notebook objects or sections of notebooks
      .Description
        If run with no parameters it will return the current user's personal notebooks.
        If run with just a -Notebook parameter it will return that notebook (which might belong to a group)
        If run with -Notebook and -Sections it will return the sections in that notebook,
        And if run with just -Sections it will return all the sections in the user's personal notebooks.
      .Example
        >Get-GraphOneNoteBook   team
        Looks for a workbook with a displayname begining "team" in the users workbooks. the search is case insensitive.
      .Example
        >Get-GraphOneNoteBook  -SectionName Powershell
        Finds a "PowerShell" secion in any of the users workbooks. Again the search is case insensitive
      .Example
        >Get-GraphTeam 'Consultants' -Notebooks | Set-GraphHomeNotebook
        >Get-GraphOneNoteBook -AllSections
        The first command changes the default notebook and selects different sections from the the previous command
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Alias('Get-GraphNoteBook')]
    param    (
        #A graph URI pointing to the notebook, or a notebook object where the .self property is a graph URI...
        $Notebook ,
        [Parameter(ValueFromPipeline=$true,DontShow=$true)]
        $InputObject,
        #If specified returns the sections of the notebook.
        [switch]$AllSections,
        #if specified filters the returned objects by to those with names begining with ...
        [ArgumentCompleter([OneNoteSectionCompleter])]
        [string]$SectionName
    )
    process  {
        if ($InputObject) {$Notebook = $InputObject}
        $msg = "Getting Information about Notebook $($Notebook.DisplayName)"
        Write-Progress $msg
        #convert notebook object with self property to a string. If we don't have a string ager this something is wrong
        if  ($Notebook.self) {$Notebook=$Notebook.self}
        if  ($Notebook -and $Notebook -isnot [string] ) {
            Write-Warning "Invalid notebook parameter" ; return
        }
        $webparams = @{
            'AsType'                = ([MicrosoftGraphNotebook])
            'ExcludeProperty'       = '@odata.context', 'sections@odata.context'
        }
        #If it didn't come in as sting or an object with a self parameter bail out.
        #if it is a path to a notebook use that
        if ($Notebook -and $Notebook -match "$GraphUri/.*/onenote/notebooks/.+") {
                $webparams['uri']       = "$Notebook`?`$expand=Sections"
            }
        else {
            $webparams['valueonly'] = $true
            $webparams['uri']       = "$GraphUri/me/onenote/notebooks?`$expand=Sections"
            if ($Notebook) { #it had a notebook parameter as a string but it wasn't a path so look for it by name
                $webparams['uri']  += ('&$filter=startswith(tolower(displayname),''{0}'')' -f ($Notebook.ToLower() -replace '\*$',''))
            }
        }
        $response = Invoke-GraphRequest @webparams
        #Sections fetched this way won't have parentNotebook, so make sure it is available when needed
        foreach ($bookobj in $response) {
            foreach ($s in $bookobj.sections) { $s.parentNotebook = $bookobj }
            if     ($AllSections) {$bookobj.Sections}
            elseif ($SectionName) {$bookobj.Sections.where({$_.displayname -like $SectionName})}
            else                  {$bookobj}
        }
        Write-Progress $msg -Completed
    }
}

function Get-GraphOneNoteSection {
    <#
      .Synopsis
        Gets details of  sections in OneNote notebooks or their pages
     .Description
        This command interogates  https://graph.microsoft.com/v1.0
            /users/{id}/onenote/notebooks/{id}/sections
        or /groups/{id}/onenote/notebooks/{id}/sections
        or  /sites/{id}/onenote/notebooks/{id}/sections
        which requires consent to use the Notes.Create or Notes.Read scope or better.
        If given a Notebook parameter it returns the sections in the notebook.
        If given a section parameter it either returns details of the section, or
        if the -Pages or -Name Parameters are given returns pages from the section
      .Example
        >$notebook = Get-GraphTeam  consultants -Notebooks
        >$notebook.sections[0]  | Get-GraphOneNoteSection  -PageTitle change
        The first line gets the Notebooks object for the 'consultants' team. This object
        has a 'sections' collection. The second line uses pipes a member of this collection as the
        into Get-GraphOneNoteSection to return the pages in the first section, with the title begining "change".
      .Example
        >Get-GraphOneNoteSection private -notebook $notebook -allpages
        In this example the notebook used in the first example is passed as a notebook is piped into command to get a section, by contrast with the previous section

      .Example
        >Get-GraphOneNoteSection -Section $section -Pages -Name "test" | Remove-GraphOneNotePage -Force
        >Gets all pages with names that begin 'Test...' and removes
      $section may be the a section object (from the Sections collection of a notebook object, or
      form Get-GraphOneNotebook -Sections ) or the URL for a section.
    #>
    [cmdletbinding()]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenotePage],ParameterSetName='Pages')]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenoteSection])]
    [Alias('Get-GraphNoteBookSection')]
    param    (
        #A graph URI pointing to the section, or a section object where the .self property is a graph URI or a section name...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([OneNoteSectionCompleter])]
        $Section ,
        [Parameter()]
        #The notebook to query for section(s) if sections is empty or contains a name
        $Notebook ,
        #If specified, returns the pages in the section(s).
        [switch]$AllPages,
        #If specified filters pages or Sections to those with names beginning ...
        [string]$PageTitle
    )
    process  {
        $msg = "Getting Information about Notebook section $($section.DisplayName)"
        Write-Progress $msg
        if     ($Section -is [MicrosoftGraphNotebook]) {$Notebook = $Section}
        elseif ($Section.self) {$Section = $Section.self}
        if     ($Notebook -and  ($Section -isnot [string] -or $section -Notmatch "^$graphuri.*/onenote/sections/.+")) {
            #A notebook has sections URL we'll use it. If not if it's an object with a self parameter try with that, otherwise if it is a string, assume it's the URI for the notebook
            if     ($Notebook.sectionsUrl)  {$uri  = $Notebook.sectionsUrl  +  '?$expand=parentNotebook'}
            elseif ($Notebook.self)         {$uri  = $Notebook.self + '/sections?$expand=parentNotebook'}
            elseif ($Notebook -is [string]) {$uri  = $Notebook      + '/sections?$expand=parentNotebook'}
            else   {Write-warning -Message 'Could not process the notebook parameter provided' }
            if     ($Section -is [string])  {$uri += ('&$filter=startswith(tolower(displayname),''{0}'')' -f ($Section.ToLower() -replace '\*$','')) }
            $results = Invoke-GraphRequest -Uri $uri -ValueOnly  -AsType ([MicrosoftGraphOnenoteSection]) -ExcludeProperty 'parentSectionGroup@odata.context',  'parentNotebook@odata.context'
            if      ($AllPages) {
                  $results | Get-GraphOneNoteSection -AllPages
                  return
            }
            elseif ($PageTitle) {
                  $results | Get-GraphOneNoteSection -PageTitle $PageTitle
                  return
            }
            else {return $results}
        }
        if     ($Section -isnot [string]  -or $Section -Notmatch "^$graphuri.*/onenote/sections/.+") {
                    Write-Warning 'Can not process the Section Parameter' ; Return
        }
        if     ($PageTitle -or $AllPages) {
            # $expand in the REST API ignores ParentNotebook so if it is empty we'll populate it
            if ($PageTitle) {$uri =  $Section + ('/Pages?$expand=parentSection,ParentNotebook&$filter=startswith(tolower(title),''{0}'')'  -f
                                                  $PageTitle.ToLower() ) }
            else            {$uri =  $Section +  '/Pages?$expand=parentSection,ParentNotebook'}
            Invoke-GraphRequest -Uri $uri -ValueOnly  -AsType ([MicrosoftGraphOnenotepage]) -PropertyNotMatch '@odata.context' |
                ForEach-Object {
                        if ((-not $_.ParentNotebook.DisplayName) -and $PSBoundParameters.section.parentnotebook) {
                                  $_.parentNotebook = $PSBoundParameters.section.parentnotebook
                        }
                $_
            }
            return
        }
        else   {
            Invoke-GraphRequest -Uri  ($Section + '?$expand=parentNotebook')  -AsType ([MicrosoftGraphOnenoteSection]) -ExcludeProperty 'parentSectionGroup@odata.context',  'parentNotebook@odata.context', '@odata.context'
        }
        Write-Progress $msg -Completed
    }

}

function New-GraphOneNoteSection {
    <#
      .Synopsis
        Adds a section to a OneNote notebook
      .Description
        This command Posts to  https://graph.microsoft.com/v1.0
            /users/{id}/onenote/notebooks/{id}/sections
        or /groups/{id}/onenote/notebooks/{id}/sections
        or  /sites/{id}/onenote/notebooks/{id}/sections
        which requires consent to use the Notes.Create or Notes.ReadWrite scope or better.
      .OUTPUTS
        Returns an object representing the new section
     .Example
        >
        >$notebook = Get-GraphTeam -ByName accounts -Notebooks
        >$section = New-GraphOneNoteSection -Notebook $notebook -SectionName "FY-19 Year End"
        >Add-GraphOneNotePage -Section $section -HTMLPage '<html><head><title>Welcome</Title></head><body><p>This section is ready for you to add your pages.</p></body></html>'

        The first command gets the team notebook for the account team; the second adds a section to it
        and the third adds a welcome page to the new section.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias('New-GraphNoteBookSection')]
    param   (
        #A graph URI pointing to the notebook, or a notebook object, this can be set by Set-GraphOneNoteHome
        [Parameter(Mandatory=$true)]
        $Notebook ,
        #Name for the new section.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $SectionName,
        #If specified, the command will run without asking for confirmation; this is the default unless Confirm Preference has been set
        [switch]$Force
    )
    begin   {
        $webParams = @{'Method'           = 'Post'
                       'ContentType'      = 'application/json'
                       'AsType'           =  [MicrosoftGraphOnenoteSection]
                       'ExcludeProperty' = @('parentSectionGroup@odata.context',  'parentNotebook@odata.context', '@odata.context')
        }
        if     ($Notebook.sectionsUrl)            {$webparams['uri'] = $Notebook.sectionsUrl}
        elseif ($Notebook.self)                   {$webparams['uri'] = $Notebook.self + "/sections"}
        elseif ($Notebook -isnot [String])        {Write-Warning -Message 'Could not process the Notebook parameter'; Return }
        elseif ($notebook -notmatch "/sections$") {$webparams['uri'] = $Notebook + "/sections"}
        else                                      {$webparams['uri'] = $Notebook }
    }
    process {
        $webparams['body']  = ConvertTo-Json @{"displayName" = $sectionName}
        Write-Debug $webparams['body']
        if ($Force -or $PSCmdlet.ShouldProcess($SectionName,"Add section to Notebook $($Notebook.displayname)")) {
            $sectionobj = Invoke-GraphRequest @webParams
            if ($Notebook -is [MicrosoftGraphNotebook]) {$sectionobj.parentNotebook = $Notebook}
            return $sectionobj
        }
    }
}

function Get-GraphOneNotePage    {
    <#
      .Synopsis
        Gets a OneNote page's metadata or content
      .Description
        This command interogates  https://graph.microsoft.com/v1.0
            /users/{id}/onenote/notebooks/{id}/sections/{id}/pages
        or /groups/{id}/onenote/notebooks/{id}/sections/{id}/pages
        or  /sites/{id}/onenote/notebooks/{id}/sections/{id}/pages
        which requires consent to use the  Notes.Read scope or better.
        It can get either the page metadata, the page content, or
        the page content marked up with IDs to update the page.
    #>
    [cmdletbinding(DefaultParameterSetName='None')]
    [outputtype([Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenotePage])]
    [Alias('Get-GraphNoteBookPage')]
    param   (
        #A graph URI pointing to the page, or a page object where the .self property is a graph URI...
        [Parameter(ParameterSetName='Page',              Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [Parameter(ParameterSetName='PageContent',       Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [Parameter(ParameterSetName='PageContentWithIDs',Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        [Parameter(ParameterSetName='PagePreview',       Mandatory=$true, ValueFromPipeline=$true,Position=0)]
        $Page,

        #A graph URI pointing to a notebook, or a notebook object. this can be set by Set-GraphOneNoteHome
        [Parameter()]
        $Notebook,

        #A graph URI pointing to a section, or a Section object  this can be set by Set-GraphOneNoteHome
        [Parameter()]
        [ArgumentCompleter([OneNoteSectionCompleter])]
        $Section,

        #If specified returns the contents of the page. Ignored if ContentWithIDs is specified
        [Parameter(ParameterSetName='PageContent',Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [switch]$Content,

        #If specified returs the contents with guids for each section where content can be inserted.
        [Parameter(ParameterSetName='PageContent')]
        [Parameter(ParameterSetName='PageContentWithIDs',Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [switch]$ContentWithIDs,

        #If specified returs a text preview of the page
        [Parameter(ParameterSetName='PagePreview',Mandatory=$true,ValueFromPipelineByPropertyName=$true)]
        [switch]$PreviewText,

        #If specified writes the preview or content to a file
        [Parameter(ParameterSetName='PageContent')]
        [Parameter(ParameterSetName='PageContentWithIDs')]
        [Parameter(ParameterSetName='PagePreview')]
        $SavePath
    )
    process {
        $webparams = @{
           'AsType'           =  ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphOnenotePage])
           'PropertyNotMatch' = '@odata'
        }
        if     ($Page -is [MicrosoftGraphOnenoteSection]) {$Section = $Page}
        elseif ($Page.self)            {$Page=$Page.self}
        if     ($Section -and ($Page -isnot [string] -or $page -notmatch "$graphuri/.*/onenote/pages/.+" )) {
                if     ($Section.self) {
                        $webParams['uri'] =   "$($Section.self)/pages?`$expand=parentsection(`$expand=ParentNotebook)"
                }
                elseif ($Section -is [string] -and $Section -match "^$graphuri.*/onenote/sections/.+") {
                        $webParams['uri'] =   "$Section/pages?`$expand=parentsection(`$expand=ParentNotebook)"
                }
                elseif ($Section -is [string] -and $Notebook) {
                        $Section = Get-GraphOneNoteSection -Section $Section -Notebook $Notebook
                         if (-not $Section -or $Section.Count -gt 1) {
                            Write-Warning "Could not resolve $Section to a unique section in the notebook."
                        }
                        else {$webParams['uri'] =   "$($Section.self)/pages?`$expand=parentsection(`$expand=ParentNotebook)"}
                }
                else {Write-Warning "Can't resolve the section given"; return }
                if ($Page -is [string]) {
                    $webParams['uri'] += '&$filter=startswith(tolower(title),''{0}'')' -f ($Page.ToLower() -replace '\*$','')
                }
                $pages     = Invoke-GraphRequest @webparams -ValueOnly
                if ($ContentWithIDs -or $Content -or $PreviewText ) {
                      $null = $PSBoundParameters.Remove('Page'), $PSBoundParameters.Remove('Notebook'), $PSBoundParameters.Remove('section')
                      $pages | Get-GraphOneNotePage @PSBoundParameters
                      return
                }
                else {
                    foreach ($p in $pages)   {
                        if ((-not $p.ParentNotebook.DisplayName) -and $Section.parentnotebook) {
                                  $p.parentNotebook = $Section.parentnotebook
                        }
                        $p
                    }
                    return
                }
        }
        elseif ($Page -isnot [string] -or $page -notmatch "$graphuri/.*/onenote/pages/.+") {
                Write-Warning -Message 'Could not process the page parameter' ; return
        }
        if     ($PreviewText -and -not $PSBoundParameters['savepath']) {
               return (Invoke-GraphRequest -Uri  "$Page/Preview").previewText
        }
        elseif ($PreviewText)          {
               (Invoke-GraphRequest -Uri  "$Page/Preview").previewText | Out-File $savePath
               return
        }
        elseif (-not ($ContentWithIDs -or $Content))           {
          Invoke-GraphRequest -Uri $page @webparams
          return
        }
        elseif (      $ContentWithIDs) {$uri = "$Page/Content?includeIDs=true"}
        else                           {$uri ="$Page/Content"}

        #Page content breaks the JSON parser, so ask for it to go to a file instead. If didn't want a file, read the file back and delete it.
        if     (-not $PSBoundParameters['savepath']) {$SavePath = [system.io.path]::GetTempFileName() }
        Invoke-GraphRequest -Uri  $uri  -OutputFilePath $SavePath
        if  (-not $PSBoundParameters['savepath']) {Get-Content $SavePath; Remove-Item $SavePath}
        #should return the outer xml property as this is HTML in an XML document. Check what else it comes back as
    }
}

function Copy-GraphOneNotePage   {
<#
  .synopsis
    Copies a one note page to a different section in the same notebook.
#>
    [Alias('Copy-GraphNoteBookPage')]
    param   (
        #The page to be copied.
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
        $Page,
        #The section the it will be copied to
        [Parameter(Position=1,Mandatory=$true)]
        $DestinationSection,
        #The group which owns the notebook if required
        $GroupID,
        [switch]$Wait
    )
    process {
        $webparams =@{
            'Method'          = 'POST'
            'ContentType'     = 'application/json'
            'ExcludeProperty' = @('@odata.context','@odata.type')
            'AsType'          = ([MicrosoftGraphOnenoteOperation])

        }
        if     ($Page.self)            {$webparams['uri']="$($Page.self)/copyToSection"}
        elseif ($Page -is [string])    {$webparams['uri']="$Page/copyToSection"}
        else   {Write-Warning -Message 'Could not process the page parameter' ; return}

        if     ($DestinationSection.id)           {$settings = @{id=$DestinationSection.id }}
        elseif ($DestinationSection -is [string]) {$settings = @{id=$DestinationSection}}
        else   {Write-Warning -Message 'Could not process the page parameter' ; return}
        #if the group GUID in the path to self in the destination use that, otherwise look for a group ID
        if     ($DestinationSection.Self -and
                $DestinationSection.Self -match "groups/([0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12})/") {
                                                   $settings['groupId'] = $Matches[1]}
        elseif ($GroupID.id)                      {$settings['groupId'] = $GroupID.ID}
        elseif ($GroupID -is [string] )           {$settings['groupId'] = $GroupID }
        elseif ($GroupID) {Write-Warning -Message 'Could not process the page parameter' ; return}

        $webparams['body'] = ConvertTo-Json $settings
        Write-Debug $webparams['body']
        $Op = Invoke-GraphRequest @webparams
        if ($Wait -and $op -and $DestinationSection.self -match "^(http.*/onenote/).*$") {
            $op2 = $op
            while ($op2.status -ne 'completed'){
                Write-Progress -Activity "Copying OneNote" -Status $op2.status
                Start-Sleep -Seconds 2
                $op2 = Invoke-GraphRequest "$($Matches[1])operations/$($op.id)" -ExcludeProperty @('@odata.context','@odata.type') -AsType ([MicrosoftGraphOnenoteOperation])
            }
            $op2
        }
        else {$op}
    }
}

<#
see  https://docs.microsoft.com/en-us/graph/api/section-copytonotebook?view=graph-rest-1.0&tabs=http
Currenly /groups/{id}/onenote/sections/{id}/copyToNotebook gives a 404 error
function Copy-GraphOneNoteSection {

    param   (
        #The section to be copied.
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
         [ArgumentCompleter([OneNoteSectionCompleter])]
        $Section,
        #The section the it will be copied to
        [Parameter(Position=1,Mandatory=$true)]
        $DestinationNotebook,

        #The name of the copy. Defaults to the name of the existing item.
        $NewName,

        #The group which owns the notebook if required
        $GroupID,
        [switch]$Wait
    )
    process {
        $webparams =@{
            'Method'          = 'POST'
            'ContentType'     = 'application/json'
            'ExcludeProperty' = @('@odata.context','@odata.type')
            'AsType'          = ([MicrosoftGraphOnenoteOperation])

        }
        if     ($Section.self)          {$webparams['uri']="$($Section.self)/copyToNotebook"}
        elseif ($Section -is [string])  {$webparams['uri']="$Section/copyToSection"}
        else   {Write-Warning -Message 'Could not process the page parameter' ; return}

        if     ($DestinationNotebook.id)           {$settings = @{id=$DestinationNotebook.id }}
        elseif ($DestinationNotebook -is [string]) {$settings = @{id=$DestinationNotebook}}
        else   {Write-Warning -Message 'Could not process the page parameter' ; return}
        #if the group GUID in the path to self in the destination use that, otherwise look for a group ID
        if     ($DestinationNotebook.Self -and
                $DestinationNotebook.Self -match "groups/([0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12})/") {
                                                    $settings['groupId']  = $Matches[1]}
        elseif ($GroupID.id)                       {$settings['groupId']  = $GroupID.ID}
        elseif ($GroupID -is [string] )            {$settings['groupId']  = $GroupID }
        elseif ($GroupID) {Write-Warning -Message 'Could not process the page parameter' ; return}

        if     ($NewName)                          {$settings['renameAs'] = $Matches[1]}

        $webparams['body'] = ConvertTo-Json $settings
        Write-Debug $webparams['body']

        $Op = Invoke-GraphRequest @webparams
        if ($Wait -and $op -and $DestinationSection.self -match "^(http.*/onenote/).*$") {
            $op2 = $op
            while ($op2.status -ne 'completed'){
                Write-Progress -Activity "Copying OneNote" -Status $op2.status
                Start-Sleep -Seconds 2
                $op2 = Invoke-GraphRequest "$($Matches[1])operations/$($op.id)" -ExcludeProperty @('@odata.context','@odata.type') -AsType ([MicrosoftGraphOnenoteOperation])
            }
            $op2
        }
        else {$op}
    }
}
#>
function Add-GraphOneNotePage    {
    <#
      .synopsis
        Adds a page (in HTML format) to an existing OneNote Section
      .description
        This posts to https://graph.microsoft.com/v1.0
            /users/{id}/onenote/sections/{id}/pages
        or /groups/{id}/onenote/sections/{id}/pages
        or  /sites/{id}/onenote/sections/{id}/pages
        which requires consent to use the Notes.Create or Notes.ReadWrite scope or better.
        To recognise the title the page needs to be in HTML with a head tag like this
        <html>
            <head>
                <title>A page</title>
                <meta name="created" content="2015-07-22T09:00:00-08:00" />
            </head>
            <body>
                <p>Here's Some text</p>
            </body>
        </html>
      .Example
        >Add-GraphOneNotePage -Section $section -HTMLPage '<html><head><title>Test Page</Title></head><body><p>Sample Paragraph</p></body></html>'
        With $Section already defined this adds a simple page, with a title and a short body.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias('Add-GraphNoteBookPage')]
    param   (
        #The section either as a URL or or as section object, which contains a self URL or a pages URL  this can be set by Set-GraphOneNoteHome
        [Parameter(Mandatory=$true)]
        $Section ,
        #The content of the page formatted as HTML
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $HTMLPage,
        #By default this is "text/html" - but if the content is multipart use "multipart/form-data; boundary={MARKER}"
        $ContentType = 'text/html',
        #If specified, the command will run without asking for confirmation; this is the default unless Confirm Preference has been set
        [switch]$Force,
        #Normally the page is added 'silently'. If passthru is specified, an object describing the new page will be returned.
        [Alias('PT')]
        [switch]$PassThru
    )

    if     ($Section.pagesURL)            {$uri = $Section.PagesUrl}
    elseif ($Section.self)                {$uri = $Section.Self + '/pages'}
    elseif ($Section -isnot [String])     {Write-Warning -Message 'Could not process the Section parameter'; Return }
    elseif ($Section -notmatch "/pages$") {$uri = $Section + '/pages'}
    else                                  {$uri = $Section}
    $webParams = @{'Method'          = 'Post'
                   'ContentType'     = $ContentType
                   'Body'            = $HTMLPage
                   'uri'             = $URI
                   'ExcludeProperty' = '@odata.context'
                   'AsType'          = ([MicrosoftGraphOnenotePage])
    }

    if ($Force -or $PSCmdlet.ShouldProcess($Section.DisplayName,'Add page to OneNote Section')) {
        $result =  Invoke-GraphRequest @webParams
        If ($PassThru) {
            if ($Section -is [MicrosoftGraphOnenoteSection]) {$result.ParentSection = $section}
            if ($section.parentnotebook.DisplayName)  {$result.parentNoteBook = $section.parentNotebook}
            return $result
        }
    }
}

function Add-FileToGraphOneNote  {
    <#
      .Synopsis
        Adds a file to a new OneNote page
      .DESCRIPTION
        Adds a file to a new one page. If the file is an image, the it will be rendered on the page
        Other files will be embedded. OneNote can render some types (e.g. PDF)
        This builds very simple HTML, which can be updated later.
        For more sophistaced pages use Add-GraphOneNotePage - with -HTMLPage as a byte array and
        specify a contentType of "multipart/form-data; boundary={MARKER}"
      .INPUTS
        A file to be sent to OneNote
      .example
        >
        >Add-FileToGraphOneNote -Path .\Modules\MsftGraph\Examples\upload.jpg -Title "Demo" -Section $notebook.sections[0] `
                  -PreContent "<h1>QR Code for the GIT repo</h1>" -PostContent "<b>Share and Enjoy</b>" -PassThru

        $Notebook holds a notebook object with one or more section(s). The command adds a page in the first section,
        titles it "Demo", and puts upload.jpg on it with formatted text before and after the image.
      .link
        Add-GraphOneNotePage.
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    [Alias('Add-FileToGraphNoteBook')]
    param   (
        #The file to upload to OneNote
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
        $Path ,
        #Title for the page. If not specified the file name will be used.
        [String]$Title,
        #Section to post to -  this can be set by Set-GraphOneNoteHome
        $Section,
        #Specifies text to add before the embedded object. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PreContent,
        #Specifies text to add after the embedded object. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PostContent,
        #A recognized mime type for the embedded file. on Windows the command will try to determine this from the file extension.
        [string]$MimeType,
        #Normally the page containing the file is added 'silently'. If passthru is specified, an object describing the new page will be returned.
        [Alias('PT')]
        [switch]$PassThru,
        #If specified the command will not pause for conformation, this is the default unless $ConfirmPreference is modified,
        [switch]$Force
    )
    process {
        #region set the URI - based on $Section - and other parameters used to send the page
        $webParams = @{ 'Method'          = 'Post'
                        'ContentType'     = 'multipart/form-data; boundary=MyAppPartBoundary'
                        'ExcludeProperty' = '@odata.context'
                        'AsType'          = ([MicrosoftGraphOnenotePage])
        }
        #if we got a section object use its pages URL, otherwise if we got a string without pages on the end, add pages, otherwise use section as is
        if     (-not $section )                  {throw [ParameterBindingException]::new("Section parameter is required")}
        elseif ($Section.pagesURL)               {$webParams['uri'] =  $Section.pagesURL}
        elseif ($Section -is [string] -and
                $Section -notmatch "/pages$")    {$webParams['uri'] = ($Section -replace '/$','')  + "/pages"}
        elseif ($Section -is [string])           {$webParams['uri'] =  $Section}
        else   {Write-Warning -Message 'Could not process the -Section paramater' ; return}
        if     ($webParams['uri']-notmatch "/onenote/sections/") {Write-Warning -Message "That does not appear to be a valid section" ; return}
        #endregion
        #region read file into a data block in HTML
        $i = Get-Item -Path $Path
        if ($i.count -ne 1) {Write-Warning "The path must match exactly one file. $path matches $($i.count)." ; return  }

        if ([System.Environment]::OSVersion -match "win" -and -not $MimeType) {
            $MimeType          =  (Get-ItemProperty -Path (Join-Path "HKLM:\SOFTWARE\Classes\" $I.Extension)  -Name "content type")."Content type"
        }
        if (-not $MimeType) {Write-Warning "The Mime type could not be determined automatically. Please specify the mimetype for '$path' with -MimeType." ; return  }

        [String]$filename      =      $i.Name
        [byte[]]$array         = [System.IO.File]::ReadAllBytes($i.fullName)

        if ($MimeType -match "image") {
                $imgTag         = '<img src="name:FileBlock" width="500"/>'
        }
        else      {
                $imgTag        = '<img data-render-src="name:FileBlock" width="1024"/>'
                $objectTag     = '<p align="center"><object data-attachment="{1}" data="name:FileBlock" type="{0}" /></p>' -f $mimetype, $filename
        }
        if ($Title) {$tTag     = [System.Web.HttpUtility]::HtmlEncode($Title)}
        else        {$tTag     =  $i.Name}
        [byte[]]$myhtml        = ([byte[]][char[]]( @"
--MyAppPartBoundary
Content-Disposition:form-data; name="Presentation"
Content-type:text/html

<!DOCTYPE html>
<html>
 <head><title>$tTag</title></head>
 <body>$PreContent<p>$imgTag</p>$objectTag $PostContent</body>
</html>

--MyAppPartBoundary
Content-Disposition:form-data; name="FileBlock"
Content-type:$mimetype
`r`n
"@ ))  + $array + ([byte[]][char[]]"`r`n--MyAppPartBoundary--`r`n")
        #endregion
        #region Send it - return the new page if -passthru was given
        if ($Force -or $PSCmdlet.ShouldProcess($tTag,'Add page to OneNote Section')) {
            $result =  Invoke-GraphRequest @webParams -Body $myhtml
            If ($PassThru) {
                if ($Section -is [MicrosoftGraphOnenoteSection]) {$result.ParentSection = $section}
                if ($section.parentnotebook.DisplayName)  {$result.parentNoteBook = $section.parentNotebook}
                return $result
            }
        }
        #endregion
    }
}

function Update-GraphOneNotePage {
    <#
        .Synopsis
            Update a OneNote page
        .Description
            This command makes PATCH requests to https://graph.microsoft.com/v1.0
                /users/{id}/onenote/sections/{id}/pages/{id}/content
            or /groups/{id}/onenote/sections/{id}/pages/{id}/content
            or  /sites/{id}/onenote/sections/{id}/pages/{id}/content
            which requires consent to use the Notes.ReadWrite  scope or better.
            To understand the use of Target, action & Postion and what needs to
            be in content for different scenarios, read the MSFT page at the link ...
        .link
            https://docs.microsoft.com/en-gb/graph/onenote-update-page
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Alias('Update-GraphNoteBookPage')]
    param   (
        #A graph URI pointing to the page, or a page object where the .self property is a graph URI...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Page,
        #The action to perform on the target element.
        [ValidateSet('replace', 'append', 'delete', 'insert', 'prepend')]
        [String]$Action =  'append' ,
        # A string of well-formed HTML to add to the page, and any image or file binary data.
        [Parameter(Mandatory=$true)]
        [String]$Content,
        #The location to add the supplied content, relative to the target element.
        [ValidateSet('after','before')]
        [String]$Position,
        #The element to update. Must be the #<data-id> or the generated <id> of the element, or the body or title keyword.
        [String]
        $Target = 'body',
        #If specified, the page is updated without prompting.
        [switch]$Force
    )
    process {
         #If the content contains binary data, the request must be sent using the multipart/form-data content type with a "Commands" part.
        if     ($Page.self)          {$uri = $Page.self}
        elseif ($Page -is [String])  {$uri = $Page}
        Write-Progress -Activity 'Updating Page' -Status 'Checking exsiting page'
        try {
            $result = Invoke-GraphRequest -Uri  $URI
        }
        catch {
            Write-Progress -Activity 'Updating Page' -Completed
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Warning "Could not find the page" ; return
            }
            else {throw $_ ; return}
        }
        $settings = @{
            'target'   = $Target   #body by default
            'action'   = $Action;  #Append by default
            'content'  = $Content;
        }

        if ($Position) {$settings['position'] = $Position}

        $json = ConvertTo-Json @($settings)
        Write-Debug $json
        if ($Force -or $PSCmdlet.ShouldProcess($result.title, 'Update Onenote Page')) {
            Write-Progress -Activity 'Updating Page'  -Status 'Applying changes'
            $result = Invoke-GraphRequest -Method Patch -Uri  "$URI/content"  -ContentType "application/json" -Body $json
            Write-Progress -Activity 'Updating Page' -Completed
        }
    }
}

function Remove-GraphOneNotePage {
    <#
      .Synopsis
        Removes a OneNote page
      .Description
           This command makes DELETE requests to https://graph.microsoft.com/v1.0
                /users/{id}/onenote/sections/{id}/pages/{id}
            or /groups/{id}/onenote/sections/{id}/pages/{id}
            or  /sites/{id}/onenote/sections/{id}/pages/{id}
            which requires consent to use the Notes.ReadWrite scope or better.
      .Example
         >Get-GraphUser -Teams -Name Consultants | Get-GraphTeam  -Notebooks |
            Get-GraphOneNoteBook -Sections -Name General | Get-GraphOneNoteSection -Pages -Name process | Remove-GraphOneNotePage
        finds a team named "consultants" which has the current user as a member, finds its notebook, finds a section named General
        within this sectioned finds page names that begin "process..." and removes them
    #>
    [cmdletbinding(ConfirmImpact='High',SupportsShouldProcess=$true)]
    [Alias('Remove-GraphNoteBookPage')]
    param   (
        #A graph URI pointing to the page, or a page object where the .self property is a graph URI...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Page,
        #If specified, the page is deleted without prompting.
        [switch]$Force
    )
    process {
        if     ($Page.self)         {$uri = $Page.self}
        elseif ($Page -is [string]) {$uri = $Page}
        else   {Write-Warning -Message 'Could not process the Page parameter' ; return}
        Write-Progress -Activity 'Deleting OneNote page(s)' -Status 'Checking page'
        try {
            $result = Invoke-GraphRequest -Uri  $uri
        }
        catch {
            Write-Progress -Activity 'Deleting OneNote page(s)' -Completed
            if ($_.exception.response.statuscode.value__ -eq 404) {
                Write-Warning "Could not find the page, it may have been deleted already." ; return
            }
            else {throw $_ ; return}
        }
        if ($Force -or $PSCmdlet.ShouldProcess($result.title, 'Delete Onenote Page')) {
            Write-Progress -Activity 'Deleting OneNote page(s)' -Status 'Deleting' -CurrentOperation $result.title
            $result = Invoke-GraphRequest  -Uri  $uri -Method  Delete
        }
        Write-Progress -Activity 'Deleting OneNote page(s)' -Completed
    }
}

function Out-GraphOneNote        {
    <#
      .Synopsis
        Output to a new OneNote page
      .INPUTS
        You can pipe any .NET object to Out-GraphOneNote
     .EXAMPLE
        Generates a page
      .EXAMPLE
        start ( Get-process  | Out-GraphOneNote -Title "Processes @ $(get-date)" -property Name,Handles,NPM,PM,VM,WS -passthru ).links.oneNoteWebUrl.href
        Generates a page in the default section (using the environment variable DefaultOneNoteSection) and opens it in a web browser.
    #>
    [CmdletBinding(DefaultParameterSetName='Page')]
    [Alias('Out-GraphNoteBook')]
    param   (
        #Specifies the objects to be represented in HTML.
        [parameter(ValueFromPipeline=$true)]
        [psobject]$InputObject,
        #Includes the specified properties of the objects in the output
        [Parameter(Position=0)]
        [String[]]$Property = @('*'),
        #The section where the content will be created: to this can be set by Set-GraphOneNoteHome
        $Section ,
        [Parameter(ParameterSetName='Page', Position=4)]
        #Specifies the text to add after the opening <BODY> tag. By default, there is no text in that position.
        [string[]]$Body,
        #Specifies the content of the <HEAD> tag. The default is "<title>HTML TABLE</title>".  If you use the Head parameter, the Title parameter is ignored.
        [Parameter(ParameterSetName='Page', Position=2)]
        [string[]]$Head,
        #Specifies a title for the Page.
        [Parameter(ParameterSetName='Page', Position=3)]
        [ValidateNotNullOrEmpty()][string]$Title,
        #Determines whether the object is formatted as a table or a list. Valid values are TABLE and LIST. The default value is TABLE.
        [ValidateSet('Table','List')][string]$As = 'Table',
        #Generates only an HTML table. The HTML, HEAD, TITLE, and BODY tags are omitted.
        [Parameter(ParameterSetName='Fragment')]
        [switch]$Fragment,
        [String[]]$ExcludeProperty,
        # Specifies text to add before the opening <TABLE> tag. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PreContent,
        #Specifies text to add after the closing </TABLE> tag. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PostContent,
        #Normally the page is added 'silently'. If passthru is specified, an object describing the new page will be returned.
        [Alias('PT')]
        [switch] $PassThru
    )
    #collect whatever comes in as Input object and process in the end block
    begin   { $stuff = @() }
    process { $Stuff = $Stuff + $InputObject}
    end     {
        #region gather the parameters for the API call - built URI from $section
         #if we got a section object use its pages URL, otherwise if we got a string without pages on the end, add pages, otherwise use section as is
        if     (-not $section )                  {throw [ParameterBindingException]::new('Section parameter is required')}
        $webParams = @{ 'Method'          = 'Post'
                        'ContentType'     = 'text/html'
                        'ExcludeProperty' = '@odata.context'
                        'AsType'          = ([MicrosoftGraphOnenotePage])
        }
        if ($Section.pagesURL) {
                $webParams['uri'] = $Section.pagesURL
        }
        elseif ($Section -is [string] -and $Section -notmatch '/pages$') {
                $webParams['uri'] = ($Section -replace '/$','')  + '/pages'
        }
        else   {$webParams['uri'] = $Section}
        if     ($webParams['uri']-notmatch '/onenote/sections/') {Write-Warning -Message 'That does not appear to be a valid section' ; return}
        #end region
        #region generate the HTML  - filtering the input properties as needed
        if ($PSBoundParameters['Property','ExcludeProperty']) {
            $Stuff = $stuff | Select-Object -Property $Property -ExcludeProperty $ExcludeProperty
            [void]$PSBoundParameters.Remove('Property')
            [void]$PSBoundParameters.Remove('ExcludeProperty')
        }

        [void]$PSBoundParameters.Remove('Section')
        [void]$PSBoundParameters.Remove('InputObject')
        [void]$PSBoundParameters.Remove('PassThru')
        if (-not $Title)    {$PSBoundParameters.Add('Title', ( $MyInvocation.Line + '  -  ' +  (Get-Date))) }
        $webParams['body'] = $Stuff | ConvertTo-Html  @PSBoundParameters
        #end region

        #Make the call, returning the URL of the new page.
        $result = Invoke-GraphRequest @webParams
        If ($PassThru) {
                if ($Section -is [MicrosoftGraphOnenoteSection]) {$result.ParentSection = $section}
                if ($section.parentnotebook.DisplayName)  {$result.parentNoteBook = $section.parentNotebook}
                return $result
        }
    }
}

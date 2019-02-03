Function Get-GraphOneNoteBook    {
    <#
      .Synopsis
        Gets notebook objects or sections of notebooks
      .Description
        If run with no parameters it will return the current user's personal notebooks.
        If run with just a -Notebook parameter it will return that notebook (which might belong to a group)
        If run with -Notebook and -Sections it will return the sections in that notebook,
        And if run with just -Sections it will return all the sections in the user's personal notebooks.
      .Example
       >Get-Graphuser -teams | select -First 1 |  Get-GraphTeam -Notebooks | select -first 1 | Get-GraphOneNoteBook -Sections | ft DisplayName, @{n="Notebook";e={$_.parentNotebook.DisplayName}}
        Gets the first team for a user; gets the first notebook for that team and gets its sections, which are formatted as a table.
      .Example
        >Get-GraphOneNoteBook  -name general
        Gets a users workbook with the name "General", "GENERAL", "general" - the search is case insensitive.
      .Example
        >Get-GraphOneNoteBook  -Sections -name Powershell
        Finds a PowerShell secion in any of the users workbooks. Again the search is casse insensitive
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    Param   (
        #A graph URI pointing to the notebook, or a notebook object where the .self property is a graph URI...
        [Parameter(ValueFromPipeline=$true)]
        $Notebook ,
        #If specified returns the sections of the notebook.
        [switch]$Sections,
        #if specified filters the returned objects by to those with names begining with ...
        [string]$Name
    )
    Begin   {
        Connect-MSGraph
        $webParams = @{Method  = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    Process {
        if ($Notebook.self) {$Notebook=$Notebook.self}
        if ($Name) {$Name = '?$filter=startswith(tolower(displayname),''{0}'')' -f ($Name.ToLower() -replace '\*$','') }

        #Combinations of params. Just notebook, Just Sections or both (with or without name), neither
        if     ($notebook -and -not $sections){
            Write-Progress "Getting Notebook Information"
            $n =  (Invoke-RestMethod @webParams -Uri  ("$Notebook`?`$expand=Sections" + ($Name -replace "^\?","&")))
            $n.pstypeNames.Add("GraphOneNoteBook")
            foreach ($s in $n.sections) {
                $s.pstypeNames.add("GraphOneNoteSection")
                Add-Member -InputObject $s -Name ParentNotebookID -MemberType NoteProperty -Value $n.id
            }
            Write-Progress "Getting Notebook Information" -Completed
            return $n
        }
        elseif ($sections) {
            if   ($notebook) {$results =  Invoke-RestMethod @webParams -Uri ("$Notebook/sections" + $Name) }
            else {$results =  Invoke-RestMethod @webParams -Uri ('https://graph.microsoft.com/v1.0/me/onenote/Sections' +  $Name) }
            $sectionList = $results.value
            foreach ($s in $sectionList) {
                $s.pstypeNames.add("GraphOneNoteSection")
                Add-Member -InputObject $s -Name ParentNotebookID -MemberType ScriptProperty -Value {$this.ParentNotebook.id}
            }
            return $sectionList
        }
        else                              {
            $n =  (Invoke-RestMethod @webParams -Uri ('https://graph.microsoft.com/v1.0/me/onenote/notebooks?`$expand=Sections' + ($Name -replace '^\?','&') ))
            if($n.value) {
                foreach ($item in $n.value) {
                    $item.pstypeNames.Add('GraphOneNoteBook')
                    foreach ($s in $item.sections) {
                        $s.pstypeNames.add('GraphOneNoteSection')
                        Add-Member -InputObject $s -Name ParentNotebookID -MemberType NoteProperty -Value $item.id
                    }
                }
                return $n.value
            }
            elseif ($n.self) {
                $n.pstypeNames.Add('GraphOneNoteBook')
                foreach ($s in $n.sections) {$s.pstypeNames.add('GraphOneNoteSection')}
                return $n
            }
        }
    }
}

Function Get-GraphOneNoteSection {
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
        >
        >$notebook = Get-GraphTeam -ByName accounts -Notebooks
        >Get-GraphOneNoteSection -Pages $notebook.sections[0]

        The first line gets the Notebooks object for the Accounts team. This has a sections
        collection. The second line gets the pages in the first section.
      .Example
      >Get-GraphOneNoteSection -Section $section -Pages -Name "test" | Remove-GraphOneNotePage -Force
      Gets all pages with names that begin 'Test...' and removes
      $section may be the a section object (from the Sections collection of a notebook object, or
      form Get-GraphOneNotebook -Sections ) or the URL for a section.
    #>
    [cmdletbinding()]
    Param   (
        #A graph URI pointing to the section, or a section object where the .self property is a graph URI...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,ParameterSetName='Sections',Position=0)]
        $Section ,
        [Parameter(ParameterSetName='Notebook')]
        $Notebook ,
        #If Specified returns the pages in the section.
        [Parameter(ParameterSetName='Sections',Position=1)]
        [switch]$Pages,
        #If specified filters pages or Sections to those with names beginning ...
        [string]$Name
    )
    Begin   {
        Connect-MSGraph
        $webParams = @{Method  = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    Process {
        if     ($Notebook) {
            #A notebook has sections URL we'll use it. If not if it's an object with a self parameter try with that, otherwise if it is a string, assume it's the URI for the notebook
            if     ($Notebook.sectionsUrl)  {$uri  = $Notebook.sectionsUrl}
            elseif ($Notebook.self)         {$uri  = $Notebook.self +"/sections"}
            elseif ($Notebook -is [string]) {$uri  = $Notebook      +"/sections"}
            else   {Write-warning -Message 'Could not process the notebook parameter provided'}
            if     ($Name)                  {$uri += ('?$filter=startswith(tolower(displayname),''{0}'')' -f ($Name.ToLower() -replace '\*$','')) }

            $results =  Invoke-RestMethod @webParams -Uri $uri
            $sectionList = $results.value
            foreach ($s in $sectionList) {
                $s.pstypeNames.add("GraphOneNoteSection")
                Add-Member -InputObject $s -Name ParentNotebookID -MemberType ScriptProperty -Value {$this.ParentNotebook.id}
            }
            return $sectionList
        }
        if     ($Section.self)         {$uri = $Section.self}
        elseif ($Section -is [string]) {$uri = $Section}
        else   {Write-Warning 'Can not process the Section Parameter' ; Return }
        if     ($Name -or $Pages) {
            if ($Name)     {$uri =  "$uri/Pages?`$filter=startswith(tolower(title),'$Name')" }
            else           {$uri =  "$uri/Pages"}
            $p = (Invoke-RestMethod @webParams -Uri  $uri).value
            foreach ($page in $p) {$page.pstypeNames.add("GraphOneNotePage")}
            return   $p
        }
        else   {
            $result  = Invoke-RestMethod @webParams -Uri  $uri
            $result.pstypeNames.add("GraphOneNoteSection")
            Add-Member -InputObject $result -Name ParentNotebookID -MemberType ScriptProperty -Value {$this.ParentNotebook.id}
            return $result
        }
    }
}

Function New-GraphOneNoteSection {
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
    Param   (
        #A graph URI pointing to the notebook, or a notebook object
        [Parameter(Mandatory=$true)]
        $Notebook ,
        #Name for the new section.
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $SectionName,
        #If specified, the command will run without asking for confirmation; this is the default unless Confirm Preference has been set
        [switch]$Force
    )
    Begin   {
        Connect-MSGraph
        $webParams = @{'Method'      = 'Post'
                       'Headers'     = $Script:DefaultHeader
                       'ContentType' = 'application/json'
        }
        if     ($Notebook.sectionsUrl)            {$uri = $Notebook.sectionsUrl}
        elseif ($Notebook.self)                   {$uri = $Notebook.self + "/sections"}
        elseif ($Notebook -isnot [String])        {Write-Warning -Message 'Could not process the Notebook parameter'; Return }
        elseif ($notebook -notmatch "/sections$") {$uri = $Notebook + "/sections"}
        else                                      {$uri = $Notebook }
    }
    Process {
        $json = ConvertTo-Json @{"displayName" = $sectionName}
        Write-Debug $json
        if ($Force -or $PSCmdlet.ShouldProcess($SectionName,"Add section to Notebook $($Notebook.displayname)")) {
            $result = Invoke-RestMethod @webParams -Uri $uri -Body $json
            $result.pstypenames.add('GraphOneNoteSection')
            #Some things need the section's parent ID, if we were passed notebook as an object, add its ID to the section object
            if     ($Notebook.id)  {
                Add-Member -InputObject $result -MemberType NoteProperty -Name ParentNotebookID -Value $Notebook.id
            }
            return $result
        }
    }
}

Function Get-GraphOneNotePage    {
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
    [cmdletbinding()]
    Param   (
        #A graph URI pointing to the page, or a page object where the .self property is a graph URI...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Page,
        #If specified returns the contents of the page. Ignored if ContentWithIDs is specified
        [switch]$Content,
        #If specified returs the contents with guids for each section where content can be inserted.
        [switch]$ContentWithIDs
    )
    Begin   {
        Connect-MSGraph
        $webParams = @{'Method'  = 'Get';
                       'Headers' = $Script:DefaultHeader
        }
    }
    Process {
        if     ($Page.self)        {$uri=$Page.self}
        elseif ($page-is [string]) {$uri=$Page}
        else   {Write-Warning -Message 'Could not process the page parameter' ; return}
        #Normally we want Invoke-RestMethod, but here we want the unprocessed content.
        if      ($ContentWithIDs) {(Invoke-WebRequest @webParams -Uri  "$uri/Content?includeIDs=true").Content}
        elseif  ($Content)        {(Invoke-WebRequest @webParams -Uri  "$uri/Content").Content}
        #should return the outer xml property as this is HTML in an XML document. Check what else it comes back as
        else           {
            $result = Invoke-RestMethod @webParams -Uri  $uri
            $result.pstypeNames.add("GraphOneNotePage")
            return $result
        }
    }
}

Function Add-GraphOneNotePage    {
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
    Param (
        #The section either as a URL or or as section object, which contrains a self URL or a pages URL
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
    Connect-MSGraph
    if     ($Section.pagesURL)            {$uri = $Section.PagesUrl}
    elseif ($Section.self)                {$uri = $Section.Self + '/pages'}
    elseif ($Section -isnot [String])     {Write-Warning -Message 'Could not process the Section parameter'; Return }
    elseif ($Section -notmatch "/pages$") {$uri = $Section + '/pages'}
    else                                  {$uri = $Section}
    $webParams = @{'Method'      = 'Post'
                   'Headers'     = $Script:DefaultHeader
                   'ContentType' = $ContentType
                   'Body'        = $HTMLPage
    }

    if ($Force -or $PSCmdlet.ShouldProcess($Section.DisplayName,'Add page to OneNote Section')) {
        $result =  Invoke-WebRequest @webParams -uri $uri
        Write-Verbose  -Message "Return status was $($result.StatusCode)/$($result.StatusDescription)"
        If ($PassThru) {
            $p = ConvertFrom-Json $result.Content
            $p.pstypeNames.add('GraphOneNotePage')
            Add-Member -InputObject $p -MemberType NoteProperty -Name 'ParentSection' -Value $Section
            return $p
        }
    }
}

Function Add-FileToGraphOneNote  {
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
      .link
        Add-GraphOneNotePage.
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param   (
        #The file to upload to OneNote
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
        $Path ,
        #Title for the page. If not specified the file name will be used.
        [String]$Title,
        #Section to post to.
        $Section,
        #Specifies text to add before the embedded object. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PreContent,
        #Specifies text to add after the embedded object. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PostContent,
        #Normally the page containing the file is added 'silently'. If passthru is specified, an object describing the new page will be returned.
        [Alias('PT')]
        [switch]$PassThru,
        #If specified the command will not pause for conformation, this is the default unless $ConfirmPreference is modified,
        [switch]$Force
    )
    Begin   {
        $webParams = @{ 'Method'      = 'Post'
                        'Headers'     = $Script:DefaultHeader
                        'ContentType' = 'multipart/form-data; boundary=MyAppPartBoundary'
        }
    }
    Process {
        #If section wasn't passed but we have it in an enviroment variable use that
        if     (-not $Section -and
                     $env:DefaultOneNoteSection) {$Section = $env:DefaultOneNoteSection}
        elseif (-not $section )                  {throw "Section parameter is required"}
        #if we got a section object use its pages URL, otherwise if we got a string without pages on the end, add pages, otherwise use section as is
        if     ($Section.pagesURL)               {$webParams['uri'] = $Section.pagesURL}
        elseif ($Section -is [string] -and
                $Section -notmatch "/pages$")    {$webParams['uri'] = ($Section -replace '/$','')  + "/pages"}
        elseif ($Section -is [string])           {$webParams['uri'] = $Section}
        else   {Write-Warning -Message 'Could not process the -Section paramater' ; return}
        #check we have a valid URI for posting to
        if     ($webParams['uri']-notmatch "/onenote/sections/") {Write-Warning -Message "That does not appear to be a valid section" ; return}

        #region read file
        $i = Get-Item -Path $Path
        if ($i.count -ne 1) {Write-Warning "The path must be exactly one file. $path matches $($i.count)." ; return  }
        #Not sure where this came from and why I don't just use [byte[]]$array = [System.IO.File]::ReadAllBytes($i.fullName)
        [String]$filename      =      $i.Name
        [byte[]]$array         = ,0 * $i.length
        $stream                =      $i.OpenRead()
        [void]$stream.Read($array, 0, $i.Length)
        $stream.Close()
        #endregion
        #region   Prepare Data to send
        $mimetype           =  (Get-ItemProperty -Path (Join-Path "HKLM:\SOFTWARE\Classes\" $I.Extension)  -Name "content type")."Content type"
        if ($mimetype -match "image") {
                   $imgTag  = '<img src="name:MyAppFileBlockName" width="500"/>'}
        else      {$imgTag  = '<img data-render-src="name:MyAppFileBlockName" width="1024"/>'
                $objectTag  = '<p align="center"><object data-attachment="{1}" data="name:MyAppFileBlockName" type="{0}" /></p>' -f $mimetype,$filename}
        if ($Title) {$tTag  = [System.Web.HttpUtility]::HtmlEncode($Title)}
        else        {$tTag  =  $i.Name}
        [byte[]]$myhtml     = ([byte[]][char[]]( @"
--MyAppPartBoundary
Content-Disposition:form-data; name="Presentation"
Content-type:text/html

<!DOCTYPE html>
<html>
 <head><title>$tTag</title></head>
 <body>$PreContent<p>$imgTag</p>$objectTag $PostContent</body>
</html>

--MyAppPartBoundary
Content-Disposition:form-data; name="MyAppFileBlockName"
Content-type:$mimetype
`r`n
"@ ))  + $array + ([byte[]][char[]]"`r`n--MyAppPartBoundary--`r`n")
#endregion

#Send it
        if ($Force -or $PSCmdlet.ShouldProcess($tTag,'Add page to OneNote Section')) {
            $result =  Invoke-WebRequest @webParams -Body $myhtml
            Write-Verbose  -Message "Return status was $($result.StatusCode)/$($result.StatusDescription)"
            If ($PassThru) {
                $p = ConvertFrom-Json $result.Content
                $p.pstypeNames.add('GraphOneNotePage')
                Add-Member -InputObject $p -MemberType NoteProperty -Name 'ParentSection' -Value $Section
                return $p
            }
        }
    }
}

Function Update-GraphOneNotePage {
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
    Param   (
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
    Begin   {
        Connect-MSGraph
    }
    Process {
         #If the content contains binary data, the request must be sent using the multipart/form-data content type with a "Commands" part.
        if     ($Page.self)          {$uri = $Page.self}
        elseif ($Page -is [String])  {$uri = $Page}
        try {
            Write-Progress -Activity 'Updating Page' -Status 'Checking exsiting page'
            $result = Invoke-RestMethod  -Headers $Script:DefaultHeader -Uri  $URI -Method  Get
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
            $result = Invoke-WebRequest -Method Patch -Uri  "$URI/content" -Headers  $Script:DefaultHeader -ContentType "application/json" -Body $json
            Write-Progress -Activity 'Updating Page' -Completed
            Write-Verbose  -Message "Update response was $($result.statuscode)/$($result.statusDescription)"
        }
    }
}

Function Remove-GraphOneNotePage {
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
    Param   (
        #A graph URI pointing to the page, or a page object where the .self property is a graph URI...
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        $Page,
        #If specified, the page is deleted without prompting.
        [switch]$Force
    )
    Begin   {
        Connect-MSGraph
    }
    Process {
        if     ($Page.self)         {$uri = $Page.self}
        elseif ($Page -is [string]) {$uri = $Page}
        else   {Write-Warning -Message 'Could not process the Page parameter' ; return}
        try {
            Write-Progress -Activity 'Deleting OneNote page(s)' -Status 'Checking page'
            $result = Invoke-RestMethod  -Headers $Script:DefaultHeader -Uri  $uri -Method  Get
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
            $result = Invoke-WebRequest  -Headers  $Script:DefaultHeader -Uri  $uri -Method  Delete
            Write-Verbose -Message "Delete response was $($result.statuscode) $($result.statusDescription)"
        }
        Write-Progress -Activity 'Deleting OneNote page(s)' -Completed
    }
}

Function Out-GraphOneNote        {
    <#
      .Synopsis
        Output to a new OneNote page
      .INPUTS
        You can pipe any .NET object to Out-OneNoteLive
     .EXAMPLE
        Generates a page
      .EXAMPLE
        start ( Get-Process | Out-OneNoteLive -Title "Processes @ $(get-date)" -property Name,Handles,NPM,PM,VM,WS )
        Generates a page and opens it
    #>
    [CmdletBinding(DefaultParameterSetName='Page')]
    Param   (
        #Specifies the objects to be represented in HTML.
        [parameter(ValueFromPipeline=$true)]
        [psobject]$InputObject,
        #Includes the specified properties of the objects in the output
        [Parameter(Position=0)]
        [System.Object[]]
        $Property,
        #The section to the content to this can be set in an environment variable DefaultOneNoteSection.
        $Section,
        [Parameter(ParameterSetName='Page', Position=3)]
        #Specifies the text to add after the opening <BODY> tag. By default, there is no text in that position.
        [string[]]$Body,
        #Specifies the content of the <HEAD> tag. The default is "<title>HTML TABLE</title>".  If you use the Head parameter, the Title parameter is ignored.
        [Parameter(ParameterSetName='Page', Position=1)]
        [string[]]
        $Head,
        #Specifies a title for the Page.
        [Parameter(ParameterSetName='Page', Position=2)]
        [ValidateNotNullOrEmpty()][string]$Title,
        #Determines whether the object is formatted as a table or a list. Valid values are TABLE and LIST. The default value is TABLE.
        [ValidateSet('Table','List')][string]$As = 'Table',
        #Generates only an HTML table. The HTML, HEAD, TITLE, and BODY tags are omitted.
        [Parameter(ParameterSetName='Fragment')]
        [ValidateNotNullOrEmpty()][switch]$Fragment,
        # Specifies text to add before the opening <TABLE> tag. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PreContent,
        #Specifies text to add after the closing </TABLE> tag. By default, there is no text in that position.
        [ValidateNotNullOrEmpty()][string[]]$PostContent
    )
    Begin   { $stuff = @() }
    Process { $Stuff = $Stuff + $InputObject}
    End     {
        Connect-MSGraph
        $webParams = @{ Method      = "Post"
                        Headers     = $Script:DefaultHeader
                        ContentType ="text/html"
        }
        #If section wasn't passed but we have it in an enviroment variable use that
        if(-not $Section -and $env:DefaultOneNoteSection) {$Section = $env:DefaultOneNoteSection}
        elseif(-not $section ) {throw "Section parameter is required"}
        #if we got a section object use its pages URL, otherwise if we got a string without pages on the end, add pages, otherwise use section as is
        if ($Section.pagesURL) {$webParams['uri'] = $Section.pagesURL}
        elseif ($Section -is [string] -and $Section -notmatch "/pages$") {$webParams['uri'] = ($Section -replace '/$','')  + "/pages"}
        else   {$webParams['uri'] = $Section}

        #check we have a valid URI for posting to
        if ($webParams['uri']-notmatch "/onenote/sections/") {Write-Warning -Message "That does not appear to be a valid section" ; return}

        #Generate HTML
        [void]$PSBoundParameters.Remove("Section")
        [void]$PSBoundParameters.Remove("InputObject")
        if (-not $Title)    {$PSBoundParameters.Add("Title", ( $MyInvocation.Line + "  -  " +  (Get-Date))) }
        $webParams['body'] = $Stuff | ConvertTo-Html  @PSBoundParameters
        #And post it, returning the URL of the page.
        (Invoke-RestMethod @webParams ).links.onenoteWebUrl.href
    }
}

Function Add-GraphOneNoteTab     {
    <#
      .Synopsis
        Adds a tab in a Teams channel for a OneNote section or Notebook
      .Description
        This posts to https://graph.microsoft.com/v1.0/teams/{id}/channels/{id}/tabs
        which requires consent to use the Group.ReadWrite.All scope.
        The Notebook Parameter has an alias of 'Section' and will accept either
        a OneNote Notebook object (or its 'Self' URI - which requires the tab name to be
        set explicitly) or a Section object. If the notebook is specified it opens at the
        first section.
      .Example
        >
        > $section = Get-GraphTeam -ByName accounts -Notebooks | Select-Object -ExpandProperty sections  | where displayname -like "FY-19*"
        > $channel = Get-GraphTeam -ByName accounts -Channels -ChannelName 'year-end'
        > Add-GraphOneNoteTab  $section $channel -TabLabel "FY-19 Notes"

        The first command gets the Notebook for the Accounts team and finds the "FY-19 Year End" section
        The second command gets the channels for the same team and finds the "Year end" channel
        The Third command creates a tab in the channel named 'FY-19 Notes' which opens the team notebook
        at its 'FY-19 Year End' section.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        #The Notebook or Section to associate with the tab
        [Parameter(Mandatory=$true,Position=0)]
        [Alias('Section')]
        $Notebook,
        #An ID or Channel object which may contain the team ID; the tab will be created in this channel
        [Parameter(Mandatory=$true, Position=1)]
        $Channel,
        #A team ID, or a team object if the team can't be found from the the channel
        $Team,
        #The label for the tab, if left blank the name of the Notebook or Section will be sued
        $TabLabel,
        #Normally the tab is added 'silently'. If passthru is specified, an object describing the new tab will be returned.
        [Alias('PT')]
        [switch]$PassThru,
        #If Specified the tab will be added without pausing for confirmation, this is the default unless $ConfirmPreference has been set.
        $Force
    )
    Connect-MSGraph
    if (-not $Script:WorkOrSchool) {Write-Warning   -Message "This command only works when you are logged in with a work or school account." ; return    }
    if       ($Channel.Team)           {$Team     = $Channel.Team }
    elseif   ($Team.id)                {$Team     = $Team.id}
    elseif   ($team -isnot [string])   {Write-Warning 'Unable to determine the team, please specify it explicitly'; return}
    if       ($Channel.id) {           $Channel   = $Channel.id }
    elseif   ($Channel-isnot [string]) {Write-Warning 'Unable to determine the channel'; return}
    if       (-not $TabLabel -and
                $notebook.displayName) {$TabLabel = $Notebook.displayName}
    elseif   (-not $TabLabel)          {Write-warning 'Unable to determin a name for the tab, please specify one explicitly'; return}

    $webparams = @{'Method'       = 'Post';
                   'Uri'          = "https://graph.microsoft.com/beta/teams/$team/channels/$channel/tabs" ;
                   'Headers'      =  $Script:DefaultHeader;
                   'ContentType'  = 'application/json'
    }
    #This bit had to be reverse engineered, from a beta version of the API, so if it works past next week, be happy.
    #If the "Notebook" object is actually a section, and it was fetched by one of the module commands (get-GraphTeam -notebook, or get-graphNotebook -section)
    #then $Notebook it will have a a parentNotebook ID. This IF..Else is to make sure we have the real notebook ID, and catch a sectionID if there is one.
    if   ($Notebook.parentNotebookid) {
                    $ParamsPt2    = '&notebookSource=PickSection&sectionId='+ $Notebook.id
                    $NotebookID   = $Notebook.parentNotebookid
          }
    else  {         $ParamsPt2    = '&notebookSource=New'
                    $NotebookID   = $Notebook.id }

    #if $Notebook is a section its url will end ?wd=(something). We need to split this off the URL and re-use it. The () need to be unescapted too,
    if ($notebook.links.oneNoteWebUrl.href -match '\?(wd=.*$)') {
                $ParamsPt2       += '&' + ( $Matches[1] -replace '%28','(' -replace '%29',')' )
                $OnenoteWebUrl    = $notebook.links.oneNoteWebUrl.href  -replace  '\?wd=.*$', ''
    }
    else      { $OnenoteWebUrl    = $notebook.links.oneNoteWebUrl.href}

    #We need the teamsite URL for the team who owns this channel, and the URL to the the Notebook. Both need to be escaped.
    $OnenoteWebUrl  = $OnenoteWebUrl                           -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'
    $siteUrl        = (Get-GraphTeam -Team $Team -Site).webUrl -replace "%", "%25" -replace '/','%2F' -replace ':','%3A'

    #Now we need to build up the mother and father of all URIs It contains the ID and URL for the notebook (not section). The Name, the teamsite. And Section specifics if applicable.
    $URIParams      = "?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&ui={locale}&tenantId={tid}"+
                      "&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$Team%2Fnotes%2Fnotebooks%2F"+ $NotebookID   +
                      "&oneNoteWebUrl=" + $oneNoteWebUrl +
                      "&notebookName="  + [uri]::EscapeDataString( $notebook.displayName ) +
                      "&siteUrl="       + $SiteUrl +
                      $ParamsPt2

    #Now we can create the JSON. Such information as there is can be found at https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs
    $json = ConvertTo-Json ([ordered]@{
                'TeamsAppId'      = '0d820ecd-def2-4297-adad-78056cde7c78'
                'name'            = $TabLabel
                'configuration'   = [ordered]@{
                    'entityId'    = ((New-Guid).tostring() + "_" +  $Notebook.ID)
                    'contentUrl'  = "https://www.onenote.com/teams/TabContent" + $URIParams
                    'removeUrl'   = "https://www.onenote.com/teams/TabRemove"  + $URIParams
                    'websiteUrl'  = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"
                }})
    $json= $json  -replace "\\u0026","&"
    Write-Debug $json
    if ($Force -or $PSCmdlet.ShouldProcess($TabLabel,"Add Tab")) {
        $result = Invoke-RestMethod @webParams -body $json
        if ($PassThru) {
            $result.pstypeNames.add('GraphTab')
            #Giving a type name formats things nicely, but need to set the name to be used when the tab is displayed
            Add-Member -InputObject $result -MemberType NoteProperty -Name teamsAppName -Value 'OneNote'
            return $result
        }
    }
}

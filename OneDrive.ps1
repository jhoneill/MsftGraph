using namespace Microsoft.Graph.PowerShell.Models

function Get-GraphDrive        {
    <#
      .Synopsis
        Gets information about a OneDrive volume
      .Description
        Run with no parameters this gets information about the volume for the current user.
        It can get information about another volume by specifying -Drive
        It can get information about the root folders, or the contents of a specific folder,
        or a link to a special folder  or recent items.
      .Example
        >Get-GraphDrive
        Returns the drive for the current user.
      .Example
        >get-graphdrive -Subfolders
        Returns the root folders in the the current user's drive. Formatting is defined
        to display the items like a normal directory, but other properties are also available.
      .Example
        >
        >$d = get-graphteam -Drive | select -first 1
        >get-graphdrive -Drive $d -SpecialFolder Documents

        The first line gets the first team drive for a user, the second gets
        the items in its Documents folder
      .Example
        >get-graphdrive -Drive $d -FolderPath general
        This example uses the team drive found in the previous one and gets the contents of the team's "General" folder
      .Example
        >get-graphdrive -Drive $d -itemPath general
        Instead of getting the the items in the General folder, this returns an object representing the folder itself
      .Example
        >Get-GraphDrive -Search preferredLanguage -FolderPath 'root:/Scripts'
        This does a freetext search of "preferredLanguage" in the scripts folder; because no drive is
        specified this folder is on the current user's drive.
        Note that searches do not return the parent path if you need to find the folder path you can do
        get-graphitem [-drive {drive}] -itemid with either the item's own ID or its parent's ID.
      .Example
        >$folder = (get-graphuser -Drive).root.children | where name -eq scripts
        >get-graphdrive -Drive $folder.parentReference.driveId -FolderID $folder.id
        The first command gets the users drive, and looks for a known folder as a child item in the drive-root.
        This folder can't be piped into get-graphdrive, so the drive id and folder id are passed.
        In this case the drive ID could be ommitted because the default is to use the user's home drive
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write-warning could be used, ut the is informational non-output.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Drive cache is intended to be accessible outside the module.')]
    param   (
        #The drive to examine - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Drive = 'me/Drive',

        #If specified gets the items in a folder by the path from {drive}/root:
        [Parameter(Mandatory=$true, ParameterSetName='FolderName',Position=0)]
        [ArgumentCompleter([OneDriveFolderCompleter])]
        [Alias("Path")]
        [String]$FolderPath,

        #If Specified gets the items in a folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='FolderID')]
        [String]$FolderID,

        #If specified gets one of the special folders (Documents, photos etc) in the drive. If they don't already exist the server appears to create them.
        [Parameter(Mandatory=$true, ParameterSetName='Special')]
        [ValidateSet('Apps','Attachments','CameraRoll','Documents','Music','Photos','Public')]
        [String]$SpecialFolder,

        #If specified gets recent items in the drive
        [Parameter(Mandatory=$true, ParameterSetName='Recent')]
        [switch]$Recent,

        #If Specified gets items shared with the user
        [Parameter(Mandatory=$true, ParameterSetName='Shared')]
        [switch]$SharedWithMe ,

        [Parameter(ParameterSetName='FolderID')]
        [Parameter(ParameterSetName='FolderName')]
        [Parameter(ParameterSetName='Special')]
        [Parameter(ParameterSetName='Shared')]
        [Parameter(ParameterSetName='Recent')]
        [string]$Include,

        #Enables a free text search of the selected content
        [Parameter(ParameterSetName='RootSearch')]
        [Parameter(ParameterSetName='Shared')]
        [Parameter(ParameterSetName='FolderID')]
        [Parameter(ParameterSetName='FolderName')]
        [string]$Search,

        #If specified returns the subfolders - if no FolderPath or FolderID is given will return folders of the root drive
        [Parameter(ParameterSetName='RootFolders')]
        [Parameter(ParameterSetName='FolderID')]
        [Parameter(ParameterSetName='FolderName')]
        [Parameter(ParameterSetName='None')]
        [Switch]$Subfolders,

        #if specified gets a file or folder by the path from {drive}/root:
        [ArgumentCompleter([OneDrivePathCompleter])]
        [Parameter(Mandatory=$true, ParameterSetName='ItemName')]
        [String]$ItemPath,

        #If Specified gets the a file or folder item by ID
        [Parameter(Mandatory=$true, ParameterSetName='ItemID',ValueFromPipelineByPropertyName=$true)]
        [String]$ItemID,

        #If specified does not display a message when a folder is empty
        [switch]$Quiet
    )
    process {
        #region Sort out the Drive - it might be "me/drives" (the default), "drives/drive-id", "drive-id" or a drive object with an ID.
        #       Fix up the last two; check the drive is accessible and then cache the id --> name
        if     ($Drive.drive)            {$Drive = "drives/$($Drive.drive)"}
        elseif ($Drive.id)               {$Drive = "drives/$($Drive.id)"}
        elseif ($Drive -notmatch './.')  {$Drive = "drives/$Drive"      }
        #Strip leading and trailing / from $drive so it fits in the URI template.
        $Drive = $Drive -replace '/$','' -replace '^/',''
        try   {
            if  (ContextHas -WorkOrSchool) {$uri = "$GraphUri/$Drive`?`$expand=root(`$expand=children)"}
            else                           {$uri = "$GrpahUri/$Drive"} #The expand fails on consumer one drive
            $driveObj  =  Invoke-GraphRequest -uri $uri #Don't convert to a type yet
            $global:DriveCache[$driveObj.id] = $driveObj.name
        }
        catch {
            $driveObj = $null
            throw ("Error trying to get drive $drive - the code was " + $_.exception.response.statuscode.value__  ) ; return
        }
        #endregion
        #region Getting a single item (file or folder) by ID or by path.
        # Make something we can insert in a REST URI
        if     ($ItemID   -and $ItemID -Match '^/?items')       {$ItemID   =  $ItemID  -replace '^/?(.*)/?$',            '$1' } #Allow "items/{id}" Strip off any leading or trailing /
        elseif ($ItemID   )                                     {$ItemID   =  $ItemID  -replace '^/?(.*)/?$',      'items/$1' } #Allow "{id}". Strip off any leading or trailing / and prepend "items/"
        elseif ($Itempath -in @("root:", "/", "root:/") )       {$ItemID   = 'root' }                                           #Convert "root:", "/" or root:/" to "root"
        elseif ($ItemPath -and $ItemPath -Match '^/?root:')     {$ItemID   =  $ItemPath -replace '^/?(.*?)[:/]*$',       '$1:'} #Allow "[/]root:{path}" strip any leading / or trailing / or : and append ":"
        elseif ($ItemPath )                                     {$ItemID   =  $ItemPath -replace '^/?(.*?)[:/]*$', 'root:/$1:'} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"
        #if we had an item ID or built an itemID string from the path, get the item, add a type and return it
        if     ($ItemID ) {
            try   {
                Invoke-GraphRequest -Uri "$GraphUri/$Drive/$ItemID" -ExcludeProperty '@odata.context', '@microsoft.graph.downloadUrl' -AsType ([MicrosoftGraphDriveItem])  |
                    Add-Member  -PassThru -MemberType AliasProperty -Name ItemID -Value 'id'           |
                    Add-Member  -PassThru -NotePropertyName Drive -NotePropertyValue $driveObj.id
                return
            }
            catch {
                if ($_.exception.response.statuscode.value__ -eq 404) {
                     Write-Warning -Message "Item Not found" ; return
                }
                #we got something other than a 404 error
                else {Write-Warning -Message $_.exception.tostring() ; return }
            }
        }
        #endregion
        #region Getting collections of items either in special folders by name, normal folders by path/id, recent items, or items "shared with me".
        #if we got a folder path or ID, search for its items; first make make sure we can insert it into the URL
        if    (($Search) -and -not
               ($FolderID   -or $FolderPath) )                  {$FolderID = 'root'}                                               #If We were asked to search but not told where, choose "root"
        elseif ($FolderID  -and $FolderID -Match '^/?items')    {$FolderID =  $FolderID   -replace '^/?(.*)/?$',             '$1'} #Other processing mirrors items above.
        elseif ($FolderID )                                     {$FolderID =  $FolderID   -replace '^/?(.*)/?$',       'items/$1'}
        elseIf ($FolderPath -in @("root:", "/", "root:/") )     {$FolderID = 'root' }
        elseif ($FolderPath -and $FolderPath -Match '^/?root:') {$FolderID =  $FolderPath -replace '^/?(.*?)[:/]*$',       '$1:' }
        elseif ($FolderPath )                                   {$FolderID =  $FolderPath -replace '^/?(.*?)[:/]*$', 'root:/$1:' }
        elseif ($SpecialFolder)                                 {$FolderID = "special/$SpecialFolder"                            }

        if     ($FolderID -or $SharedWithMe -or $Recent) {
            if     ($FolderID     -and $Search) {$uri  = "$GraphUri/$Drive/$FolderID/search(q='$Search')?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file"}
            elseif ($FolderID                 ) {$uri  = "$GraphUri/$Drive/$FolderID/children?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file"    }
            elseif ($SharedWithMe -and $search) {}  #can these be combined ?
            elseif ($SharedWithMe             ) {$uri  = "$GraphUri/me/drive/SharedWithMe?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file"        }
            elseif ($Search                   ) {$uri  = "$GraphUri/me/drive/search(q='$Search')?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file" }  #me or $drive
            elseif ($Recent                   ) {$uri  = "$GraphUri/$Drive/recent?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file"                }  #Me or $drive

            if     ($Include -match '\*.*\*'  ) {Write-Warning "Graph doesn't do *something* searches"}
            elseif ($Include -match '\*$'     ) {$uri += "&`$filter=startswith(name,'{0}')" -f ($Include -replace '\*$','') }
            elseif ($Include -match '^\*'     ) {$uri += "&`$filter=endswith(name,'{0}')"   -f ($Include -replace '^\*','') }
            elseif ($Include -match '\*'      ) {Write-Warning "Graph doesn't do something*something searches"}
            elseif ($Include -match '\*'      ) {Write-Warning "Graph doesn't do something*something searches"}

            try    {$Children = Invoke-GraphRequest $uri -ValueOnly -ExcludeProperty '@odata.etag', '@odata.type' -AsType ([MicrosoftGraphDriveItem])  |
                                    Where-Object {$_.folder -or -not $Subfolders}

                    $Children | Add-Member  -PassThru -MemberType AliasProperty -Name ItemID -Value 'id'      |
                                Add-Member  -PassThru -NotePropertyName Drive -NotePropertyValue $driveObj.id |
                                Sort-Object -Property name
            }
            catch  {
                    if ($_.exception.response.statuscode.value__ -eq 404) {
                          Write-Warning -Message "Not found" ;return
                    }
                    else {Write-Warning -Message $_.exception.tostring() ; return}
            }
            #This may have returned nothing or no subfolders.
        }
        #endregion
        #region Getting the drive - either the drive object itself , or the folders in its root.
        elseif ($Subfolders) {
            $children = $driveObj.root.children.where({$_.folder})
            $children | ForEach-Object {New-Object -TypeName MicrosoftGraphDriveItem -Property $_} |
                Add-Member  -PassThru -MemberType AliasProperty -Name ItemID -Value 'id'           |
                Add-Member  -PassThru -NotePropertyName Drive -NotePropertyValue $driveObj.id      |
                    Sort-Object -Property Name
        }
        else                 {
            [void]$driveObj.remove('root@odata.context')
            [void]$driveObj.remove('@odata.context')
            New-object -TypeName MicrosoftGraphDrive -Property $driveObj |
                Add-Member  -PassThru -MemberType AliasProperty -Name Drive -Value 'id'
            return
        }
        #endregion

        #region items and drives will explictly return result. If we were looking in a folder with no suitable children, tell the user the lack of result is OK
        if (-not $children -and -not $Quiet) { Write-Host  "Folder exists, but had nothing to return."}
        #endregion

        <#
        see https://docs.microsoft.com/en-gb/graph/api/driveitem-list-children?view=graph-rest-1.0
        --- We can also get
        https://graph.microsoft.com/v1.0/me/drive/root:/scripts/type-info.xlsx:/content?format=pdf    -OutFile \temp\pictures.pdf)
        https://graph.microsoft.com/v1.0/me/drive/items/<id>/lastModifiedByUser
        https://graph.microsoft.com/v1.0/mary@contoso.com/drive/root/children
        https://graph.microsoft.com/v1.0/mary@contoso.com/drive/items/<id>/lastModifiedByUser/manager
        #>
    }
}

function New-GraphFolder       {
    <#
      .synopsis
        Creates a new folder on OneDrive.
      .description
        By default this will create a new folder on the user's one drive, and if the no Parent ID is specified
        the folder will be created in the root of the drive.
      .Example
        >New-GraphFolder -Path '/Documents/Project-x'
        Creates a new folder named "Project x" in the current users Documents folder
       .Example
        >New-GraphFolder -Path 'root:/Documents/Project-Y'
        Creates a new folder named "Project Y" in the current users Documents folder
        Note that tab completion will change /Projects/ to root:/Projects
      .Example
        >
        >$drive = Get-GraphTeam -ByName Consultants -Drive
        >New-GraphFolder -Drive $drive -Path 'root:/Documents/Project Firebird/Planning'
        Gets the drive for the Consultants team; and adds a subfolder under documents.
        As in the previous examples root:/ is how tab completion would render the path, but
        '/Documents/Project Firebird/Planning' works just as well.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Drive cache is intended to be accessible outside the module.')]
    param   (
        #The name for the new folder
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [Alias('FolderPath')]
        [ArgumentCompleter([OneDriveFolderCompleter])]
        [string]$Path,
        #The drive holding the new folder - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [Parameter()]
        $Drive = 'me/Drive'
    )
    begin   {
        #  Sort out the Drive - it might be "me/drives" (the default), "drives/drive-id", "drive-id" or a drive object with an ID.
        #       Fix up the last two; check the drive is accessible and then cache the id --> name
        if     ($Drive.id)               {$drive = "drives/$($Drive.id)"}
        elseif ($Drive -notmatch './.')  {$drive = "drives/$Drive"      }
        #Strip leading and trailing / from $drive so it fits in the URI template.
        $Drive = $Drive -replace '/$','' -replace '^/',''
        try   {
            $driveObj  =  (Invoke-GraphRequest -Method GET -Uri "$GraphUri/$Drive")
            $global:DriveCache[$driveObj.id] = $driveObj.name
        }
        catch {
            throw ('Error trying to get drive $drive - the code was ' + $_.exception.response.statuscode.value__  ) ; return
        }
    }
    process {
        $settings   = @{'name'  = $path -replace '^.*/(.+?)/?$' , '$1'  #Strip any leading or trailing / keep everything after the last /
                        'folder' = @{} ;
                        '@microsoft.graph.conflictBehavior'= 'fail'
        }
        #Strip any leading or trailing / and everything after the last /
        $parentpath = $path -replace '^/?(.*)/.+?/?$' , '$1'

        if     ($parentpath -in @("", "root:", "/", "root:/") )   {$parentpath = 'root' }                                           #Convert "root:", "/" or root:/" to "root"
        elseif ($parentpath -Match '^/?root:')                    {$parentpath =  $parentpath -replace '^/?(.*?)[:/]*$',       '$1:'} #Allow "[/]root:{path}" strip any leading / or trailing / or : and append ":"
        else                                                      {$parentpath =  $parentpath -replace '^/?(.*?)[:/]*$', 'root:/$1:'} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"

        $webparams = @{
            uri             = "$GraphUri/$Drive/$parentPath/children"
            Method          = 'Post'
            body            = (ConvertTo-Json $settings)
            ContentType     = "application/json"
            AsType          =([MicrosoftGraphDriveItem])
            ExcludeProperty = @('@odata.context', '@microsoft.graph.downloadUrl')

        }
        Write-Debug $webparams.body
        if ($PSCmdlet.ShouldProcess($parentPath, "Create new OneDrive folder '$($settings.Name)'")) {
            try { Invoke-GraphRequest @webparams}
            Catch {
                if ($_.exception.response.statuscode.value__ -eq 409) {
                    Write-Warning -Message "A Confilict error was returned. The folder probably exists already."
                }
                else {throw $_ }
            }
        }
    }
}

function Show-GraphFolder      {
    <#
      .synopsis
        Opens a OneDrive folder in a browser
      .Example
        Show-GraphFolder -Path 'root:/Documents'
        Opens the documents folder from the current user's drive in the default browser
        Note that root:/documents is how tab completion will render the path, but
        /documents is equally valid
      .Example
        >
        >$drive = Get-GraphTeam -ByName Consultants -Drive
        >Show-GraphFolder -Path 'root:/Documents' -drive $drive
        Finds the drive for the consultants team, and opens its
        documents folder in the default browser
    #>
    [CmdletBinding(DefaultParameterSetName='FolderName')]
    param   (
        #If Specified gets the  folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='FolderName',Position=0)]
        [Alias("FolderPath")]
        [ArgumentCompleter([OneDriveFolderCompleter])]
        [String]$Path,
        #If Specified gets the  folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='FolderID')]
        [String]$FolderID,
        #The Drive containing the path .
        $Drive = 'me/Drive'
    )
    process {
        if ($Path.weburl)     {Start-Process $Path.weburl ; return}
        elseif ($Path.id)     {$FolderID = $Path.id}
        elseif ($Path)        {
            $item = Get-GraphDrive -ItemPath $Path -Drive $Drive
            if ($item.weburl) {Start-Process $item.weburl ; return}
        }
        if ($FolderID) {
            $item = Get-GraphDrive -ItemID $FolderID -Drive $Drive
            if ($item.weburl) {Start-Process $item.weburl ; return}
        }
   }
}

function Copy-ToGraphFolder    {
    <#
      .synopsis
        Copies filse from the local computer to one drive
      .example
        >
        >$teamdrive = Get-GraphTeam -ByName Consultants -Drive
        >dir *.xlsx |  Copy-ToGraphFolder -Drive $teamdrive -Destination 'root:/Documents'
        The first command gets the drive for a team; the second finds
        .xlsx files in the current directory, and copies them to the Documents folder
        on the team's drive.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #location of file on the local machine
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Path,
        #Location file should be copied to can be in the form "/files/" to copy to users "files" folder, or "/drives/{id}/root:/folder/Subfolder" to select another drive
        [Parameter(Mandatory=$true)]
        [ArgumentCompleter([OneDriveFolderCompleter])]
        [string]$Destination,
        #The drive, by default the current user's OneDrive.
        $Drive = 'me/Drive',
        #Mime type of file
        [String]$ContentType,
        #Specifies what to do if the file already exists.
        [ValidateSet('replace', 'fail' ,'rename')]
        $ConflictBehavior = 'replace',
        #if specified disables quick updates and uses resumable ones. Forced to true of conflict behavior is set to "fail"
        [switch]$ForceResumable
    )
    process {
        #Ensure Path gives us something we can upload
        $uploadItem = Get-Item -Path $Path
        if (-not $uploadItem)         {Write-Warning -Message "Could not find $Path"          ; return }
        if ($uploadItem.Count -gt 1 ) {Write-Warning -Message "$Path returns multiple items." ; return }

        #Byte range and Settings are only needed for "resumable upload"
        $rangeText = "bytes 0-" + ($uploadItem.length -1) + "/" + $uploadItem.Length
        $settings  = @{
            'item' = [ordered]@{
                        '@microsoft.graph.conflictBehavior'= $ConflictBehavior
                        'name'                             = $uploadItem.Name
                        'fileSystemInfo'                   = [ordered]@{
                        'lastModifiedDateTime'             = $uploadItem.LastWriteTimeUtc.ToString("yyyy-MM-ddTHH:mm:ss'Z'") #'o' might work for ISO format here
                        }
            }
        }

        #Content type is only needed for "quick upload"
        if (-not $ContentType) {
            $reg = Get-ItemProperty -Path "HKLM:\SOFTWARE\Classes\$($uploadItem.extension)"
            if ($reg.'Content Type') {
                $ContentType= $reg.'Content Type'
                Write-Verbose -Message "COPY-TOGRAPHFOLDER: Selected content type of $contentType for a $($UploadItem.Extension) file."
            }
            else {$ContentType = "application/octet-stream"}
        }

        #region Figure out what the URI should be
        #was destination writen out in full as drives/{id}/root:/{path} ? (with or without a leading /)
        if     ($Destination.parentReference -and
                $Destination.fileSystemInfo)                   { $uri  = "$GraphUri/drives/" + $Destination.parentReference.driveId + "/items/" + $Destination.id}
        elseif ($Destination -match '/?drives.*:/\w')           {$uri  = "$GraphUri/" + ($Destination -replace '^/','') }
        else { #We didn't get the drive in the destination, so is it an object, a partial path "drives/id"  or "me/drive", or just an ID
            if     ($Drive.id)                                  {$uri  = "$GraphUri/drives/$($Drive.id)/"}
            elseif ($Drive -match './.')                        {$uri  = "$GraphUri/$Drive/"             }
            elseif ($Drive)                                     {$uri  = "$GraphUri/drives/$Drive/"      }
            # the root might be "/" root: or root:/ (/root will be assumed to be a folder) anywhere else we can bolt on to the URI. We may need to put root: in front and strip leading /
            If     ($Destination -in @("root:", "/", "root:/")) {$uri += "root/"                                       }
            elseif ($Destination -Match '^/?root:')             {$uri += ($Destination -replace '^/', '')              }
            else                                                {$uri += ($Destination -replace '^/?(.*$)', 'root:/$1')}
        }

        #if URI ends with / so that's a directory, easy. .
        if    ($uri -match '/$') {$uri = $uri + $uploadItem.Name }
        else  {     # Otherwise see if we have a directory, or a file path with a valid parent
            try   { # Does the Path exist ? Try to get it and catch the 404 error that will result if destination points to a new file
                $x = Invoke-GraphRequest -Uri $uri
                #The path exists ... is it a folder or a file  ?
                if ($x.folder) {
                      $uri = $uri +'/' + $uploadItem.Name
                      Write-Verbose -Message "COPY-TOGRAPHFOLDER: $Destination appears to be a folder, will upload to a file named $($uploadItem.Name) in it."
                }
                else           {
                      #It's a file make sure the name in the JSON matches the the name in the URI
                      $settings['item'].name = $uri -replace '^.*/','' #get rid of everything up to the last slash (greedy regex)
                      Write-Warning -Message "$Destination exists as a file."
                      if ($ConflictBehavior -eq 'fail') {return}
                }
            }
            catch {
                if ($_.exception.response.statuscode.value__ -eq 404) {
                    #We couldn't find $uri - this is expected if we have been given the path to a file.
                    $folderURI  = $uri -replace "/[^/]*?$",""   #the last slash and everything after it (lazy regex)
                    Write-Verbose -Message "COPY-TOGRAPHFOLDER: $uri was not found,  checking for $folderURI"
                    try {
                        $x = Invoke-GraphRequest -Uri $folderuri
                        if ($x.folder) {
                                $settings['item'].name = $uri -replace '^.*/','' #get rid of everything up to the last slash (greedy regex)
                                Write-Verbose -Message "COPY-TOGRAPHFOLDER: $folderURI is a valid folder; will upload as a new file."
                        }
                        else  { Write-Warning -Message "There was a problem with $Destination as a target path. Neither it nor its parent look like valid folders."; return}
                    }
                    catch     { Write-Warning -Message "There was a problem with $Destination as a target path. Neither it nor its parent look like valid folders."; return}
                }
                #we got something other than a 404 error
                else          { Write-Warning -Message $_.exception.tostring() ; return}
            }
        }
        #endregion

        #If we don't want to overwrite small files, the easiest way is to a resumable update which will check if the file exists
        if ($ConflictBehavior -eq 'fail') {$ForceResumable = $true}
        if ($PSCmdlet.ShouldProcess($uploadItem.FullName,"Upload file")){
            if ($uploadItem.Length -lt 3.5mb -and -not $ForceResumable) {
                $webparams = @{
                    'Method'          = 'Put'
                    'URI'             =  ($uri + ":/Content")
                    'InputFilePath'   =  $uploadItem.FullName
                    'ContentType'     =   $ContentType
                    'AsType'          = ([MicrosoftGraphDriveItem])
                    'ExcludeProperty' = @('@odata.context', '@microsoft.graph.downloadUrl')
                }
                Invoke-GraphRequest @webparams |
                    Add-Member  -PassThru -MemberType AliasProperty -Name ItemID -Value 'id'  |
                    Add-Member  -PassThru -NotePropertyName Drive -NotePropertyValue $drive
            }
            else {
                $body               = ConvertTo-Json $settings
                try   {
                    $UploadSession  = Invoke-GraphRequest -Method Post -Uri ($uri + ":/createUploadSession") -Body $body -ContentType "application/json"
                }
                catch {
                    if ($_.exception.response.statuscode.value__ -eq 409) {
                        Write-Warning -Message "Uploading to $Destination responed 'Conflict'. This is expected if you chose 'Conflict FAIL' and the file exists" ; return
                    }
                    #we got something other than a conflict error
                    else          { Write-Warning -Message $_.exception.tostring() ; return}
                }
                if (-not $UploadSession.uploadUrl) {Write-Warning -Message 'Server did not provide an upload destination' ; return}
                else                               {Write-Verbose -Message "COPY-TOGRAPHFOLDER: Have an upload session until $($UploadSession.expirationDateTime)" }
                $oldprogressPref    = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                $result             = Invoke-WebRequest  -Method Put -Uri $UploadSession.uploadUrl -InFile $uploadItem.FullName -ContentType "application/octet-stream" -Headers @{"Content-Range"=$RangeText}
                $resultHash         = ConvertFrom-Json $result.content -AsHashtable
                $keysToRemove        = $resultHash.Keys.where({$_ -match '@'})
                foreach ($k in $keysToRemove) {[void]$resultHash.Remove($k)}
                New-Object -TypeName MicrosoftGraphDriveItem -Property $resultHash |
                    Add-Member  -PassThru -MemberType AliasProperty -Name ItemID -Value 'id'  |
                    Add-Member  -PassThru -NotePropertyName Drive -NotePropertyValue $drive
                $ProgressPreference = $oldprogressPref
            }
        }
    }
}

function Copy-FromGraphFolder  {
    <#
      .Synopsis
        Copies files from OneDrive to the local computer
      .Example
        >Copy-FromGraphFolder -Path 'root:/Scripts/Type-Info.xlsx' -Destination c:\temp
        Copies a single file from a "scripts" directory on the user's drive to c:\temp.
      .Example
        >
        >$drive = Get-GraphTeam -ByName Consultants -Drive
        >Get-GraphDrive -Drive $drive -FolderPath 'root:/Documents/Project Firebird/Planning' | Copy-FromGraphFolder -Destination c:\temp
        Gets all the files in a folder on a teams drive and copies them to C:\Temp.
    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param   (
        #The path to the file on one drive
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [ArgumentCompleter([OneDrivePathCompleter])]
        $Path,
        #The drive, by default the current user's OneDrive.
        [Parameter()]
        $Drive = 'me/Drive',
        #The destination on the local computer
        [Parameter(Position=2)]
        $Destination = $pwd,
        #If specified prevents an existing file from being overwritten.
        [Switch]$NoClobber,
        #If Specified the destination file will be returned (similar to Copy-Item)
        [Alias('PT')]
        [Switch]$Passthru
    )
    process {
        <#
        We can download from
        /drives/{drive-id}/items/{parent-id}:/{filename}
        /drives/{drive-id}/root:/{folder-path}/{filename}q
        /groups/{group-id}/drive/items/{parent-id}:/{filename}
        /me/drive/items/{parent-id}:/{filename}
        /me/drive/root:/{folder-path}/{filename}
        /sites/{site-id}/drive/items/{parent-id}:/{filename}
        /users/{user-id}/drive/items/{parent-id}:/{filename}
        #>
        if ($Path.name -and $Path.'@microsoft.graph.downloadUrl') {$sourceDetails = $Path}
        else {
            if     ($Path.parentReference -and
                    $Path.fileSystemInfo         ) {$uri  = "$GraphUri/drives/"+ $Path.parentReference.driveId + "/items/" + $path.id}
            elseif ($Path -isnot [string]        ) {throw 'An invalid Object was passed as a Path Parameter' ; return }
            elseif ($Path -match '/?drives.*:/\w') {$uri  = '$GraphUri/' + ($path -replace '^/','') }
            else { #We didn't get the drive in the destination, so is it an object, a partial path "drives/id"  or "me/drive", or just an ID
                if     ($Drive.id)                 {$uri  = "$GraphUri/drives/$($Drive.id)/"}
                elseif ($Drive -match './.')       {$uri  = "$GraphUri/$Drive/"             }
                elseif ($Drive)                    {$uri  = "$GraphUri/drives/$Drive/"      }
                if     ($path -Match '^/?root:')   {$uri += ($path -replace '^/', '')              }
                else                               {$uri += ($path -replace '^/?(.*$)', 'root:/$1')}
            }

            #Get the item. The result should have a downloadURL as property.
            try   {$sourceDetails  = Invoke-GraphRequest $uri}
            catch {Write-warning -Message "Error trying to get $uri"; return }
        }
        if (-not $sourceDetails) {Write-warning -Message 'Could not get soruce file'; return}
        if  ( Test-Path -Path $Destination -PathType Container            ) {
              $Destination = Join-Path -Path $Destination -ChildPath $sourceDetails.name
        }
        if  ((Test-Path -Path $Destination -PathType Leaf) -and $NoClobber) {
              Write-Warning "$Destination Exists, and -NoClobber was specified";
              return
        }
        if  ((Test-path -Path $Destination -IsValid      ) -and $sourceDetails.'@microsoft.graph.downloadUrl') {
              if ($pscmdlet.ShouldProcess($Destination,"Copy file to")){
                Invoke-GraphRequest -Method get -Uri $sourceDetails.'@microsoft.graph.downloadUrl' -OutputFilePath $Destination
              if ($Passthru) {Get-Item -Path  $Destination}}
        }
        elseif (-not $sourceDetails.'@microsoft.graph.downloadUrl') {
              Write-Warning -Message "Could not get the download url for $path"
        }
        else {Write-Warning -Message "$Destination is not a valid path."}
    }
}

function Set-GraphHomeDrive    {
    param    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphDrive]$Drive
    )
    $Global:PSDefaultParameterValues["*GraphFolder:Drive"]    = $Drive
    $Global:PSDefaultParameterValues["*GraphWorkBook:Drive"]  = $Drive
    $Global:PSDefaultParameterValues["*GraphWorksheet:Drive"] = $Drive
    $Global:PSDefaultParameterValues["Get-GraphDrive:Drive"]  = $Drive
}

Function Get-GraphWorkBook     {
    param    (
        #The drive to examine - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [parameter(ValueFromPipelineByPropertyName=$true)]
        $Drive = 'me/Drive',

        #if specified gets a file or folder by the path from {drive}/root:
        [ArgumentCompleter([OneDrivePathCompleter])]
        [Parameter(ValueFromPipeline=$true,Mandatory=$true, ParameterSetName='ItemName',Position=0)]
        $ItemPath,

        #If Specified gets the a file or folder item by ID
        [Parameter(Mandatory=$true, ParameterSetName='ItemID',ValueFromPipelineByPropertyName=$true)]
        $ItemID
    )
    process {
        if     ($ItemPath.drive)      {$Drive  = $ItemPath.drive}    #we might get a file with drive as a property.
        elseif ($Drive.drive)         {$Drive  = $Drive.drive   }    #Or a drive object or drive as a string
        elseif ($Drive.id)            {$Drive  = $Drive.id      }
        elseif ($Drive -is [string])  {$Drive  = $Drive -replace "^/?(drives/)?(.*?)/?$",'$2'}   #Allow "{id}" or "drives/{id}" with any leading or trailing /

        if     ($ItemPath.itemId)     {$ItemID = $ItemPath.ItemId}   #we might get the item as an object with an ItemID or ID or the path as string or an ID
        elseif ($ItemPath.Id)         {$ItemID = $ItemPath.id}
        elseif ($ItemPath )           {$ItemID = (Invoke-GraphRequest ("$GraphUri/drives/$drive" + ($ItemPath -replace  '^/?(root:/)?(.*?)[:/]*$','/root:/$2:' ) )).id} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"
        elseif ($ItemID -is [string]) {$ItemID = $ItemID  -replace '^/?(items/)?(.*?)/?$', '$2' } #Allow "{id}" or "items/{id}" with any leading or trailing /

        $webparams = @{
          'Uri'              =  "$GraphUri/drives/$drive/items/$ItemID/workbook?`$expand=application,comments,functions,names,operations,tables,worksheets"
          'PropertyNotMatch' =  '@odata'
          'AsType'           = ([Microsoft.Graph.PowerShell.Models.MicrosoftGraphWorkbook])
        }
        Invoke-GraphRequest @webparams |
            Add-Member -PassThru -NotePropertyName ItemId -NotePropertyValue $ItemID |
            Add-Member -PassThru -NotePropertyName Drive  -NotePropertyValue $Drive
    }
}

Function Import-GraphWorksheet {
    param    (
        #The drive to examine - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [parameter(ValueFromPipelineByPropertyName=$true)]
        $Drive = 'me/Drive',

        #if specified gets a file or folder by the path from {drive}/root:
        [ArgumentCompleter([OneDrivePathCompleter])]
        [Parameter(Mandatory=$true, ParameterSetName='ItemName',Position=0)]
        $ItemPath,

        #If Specified gets the a file or folder item by ID
        [Parameter(Mandatory=$true, ParameterSetName='ItemID',ValueFromPipelineByPropertyName=$true)]
        $ItemID,

        #The name of a workSheet within the file
        [alias('WorkSheetName')]
        [Parameter(Position=1)]
        [string]$SheetName = 'Sheet1',

        [switch]$AsHashTable
    )
    process {
        if     ($ItemPath.drive)      {$Drive  = $ItemPath.drive}    #we might get a file with drive as a property.
        elseif ($Drive.drive)         {$Drive  = $Drive.drive   }    #Or a drive object or drive as a string
        elseif ($Drive.id)            {$Drive  = $Drive.id      }
        elseif ($Drive -is [string])  {$Drive  = $Drive -replace "^/?(drives/)?(.*?)/?$",'$2'}   #Allow "{id}" or "drives/{id}" with any leading or trailing /

        if     ($ItemPath.itemId)     {$ItemID = $ItemPath.ItemId}   #we might get the item as an object with an ItemID or ID or the path as string or an ID
        elseif ($ItemPath.Id)         {$ItemID = $ItemPath.id}
        elseif ($ItemPath )           {$ItemID = (Invoke-GraphRequest ("$GraphUri/drives/$Drive" + ($ItemPath -replace  '^/?(root:/)?(.*?)[:/]*$','/root:/$2:' ) )).id} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"
        elseif ($ItemID -is [string]) {$ItemID = $ItemID  -replace '^/?(items/)?(.*?)/?$', '$2' } #Allow "{id}" or "items/{id}" with any leading or trailing /

        <#
            information about the sheet instead of data in the sheet ...
            * expanding charts seems to cause an issue at least if a pivot chart is present.
            $uri = "$GraphUri/drives/$drive/items/$ItemID/workbook/worksheets/$sheetName`?`$expand=charts"
            $uri = "$GraphUri/drives/$drive/items/$ItemID/workbook/worksheets/$sheetName`?`$expand=names,tables"
            Invoke-GraphRequest -Uri $uri -PropertyNotMatch odata -AsType ([MicrosoftGraphWorkbookWorksheet])
            Can get info about charts with
            $GraphUri/drives/$drive/items/$ItemID/workbook/worksheets/$sheetName/charts/{{id}}}}`$expand=series,legend,title,format,datalabels,axes
        #>
        $sheet    =  Invoke-GraphRequest "$GraphUri/drives/$Drive/items/$ItemID/workbook/worksheets/$SheetName/usedrange?`$select=Values,address,rowcount,columncount,valueTypes"
        #can select a cell range using  "$GraphUri/drives/$drive/items/$ItemID/workbook/worksheets/$SheetName/range(address='A1:B2')"
        #Or sheet-scoped named range    "$GraphUri/drives/$drive/items/$ItemID/workbook/worksheets/$SheetName/names/$RANGEName/range"
        #Or a table name                "$GraphUri/drives/$drive/items/$ItemID/workbook/tables/$TABLEName/range"
        <# Create ranges with
            $body = ' {   "name": "<<RANGENAME>>",   "reference": "=<<SHEETNAME>$A$1:$E$5",   "comment": "" }
            #workbook-scoped range
            Invoke-GraphRequest "$GraphUri/drives/$($f.drive)/items/$($f.ItemID)/workbook/names/add" -method post -body $body -ContentType "application/json"
            # or sheet-scoped range    $GraphUri/drives/$($f.drive)/items/$($f.ItemID)/workbook/worksheets/Processes/names/add"
        #>
        $headings = $sheet.Values[0] #headings is row 0 - then look at rows 1..end
        foreach ($r in 1..($sheet.rowCount-1)) {
            $c = 0  #r and c are row and column.
            #Starting at column zero, for each heading read cell into hashTable[heading] move to next column
            $newrow = [ordered]@{}
            foreach ($h in $headings) {
                if ($sheet.numberFormat[1][2] -match '[dmy]+\W*[dmy]+\W*[dmy]') {
                      $newrow[$h] = [datetime]::FromOADate($sheet.values[$r][$c])
                }
                else {$newrow[$h] = $sheet.values[$r][$c]}
                $c++
            }
            if ($AsHashTable)         { $newrow}
            else      { [pscustomObject]$newrow}
        }
    }
}

function Export-GraphWorkSheet {
    param   (
        #The drive to examine - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        $Drive = 'me/Drive',

        #if specified gets a file or folder by the path from {drive}/root:
        [ArgumentCompleter([OneDrivePathCompleter])]
        [Parameter(Mandatory=$true, ParameterSetName='ItemName',Position=0)]
        $ItemPath,

        #If Specified gets the a file or folder item by ID
        [Parameter(Mandatory=$true, ParameterSetName='ItemID')]
        $ItemID,

        #The name of a workSheet within the file
        [alias('WorkSheetName')]
        [Parameter(Position=1)]
        [string]$SheetName = 'Sheet1',

        [Parameter(ValueFromPipeline=$true)]
        $InputObject,

        [switch]$NoHeader,

        [Parameter(ParameterSetName='Table')]
        [switch]$AsTable,
        [Parameter(ParameterSetName='Range')]
        [string]$RangeName,
        [string]$dateFormat = ([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern + " " + [cultureinfo]::CurrentCulture.DateTimeFormat.ShortTimePattern)
    )
    begin   {
        if     ($ItemPath.drive)      {$Drive  = $ItemPath.drive}    #we might get a file with drive as a property.
        elseif ($Drive.drive)         {$Drive  = $Drive.drive   }    #Or a drive object or drive as a string
        elseif ($Drive.id)            {$Drive  = $Drive.id      }
        elseif ($Drive -is [string])  {$Drive  = $Drive -replace "^/?(drives/)?(.*?)/?$",'$2'}   #Allow "{id}" or "drives/{id}" with any leading or trailing /

        if     ($ItemPath.itemId)     {$ItemID = $ItemPath.ItemId}   #we might get the item as an object with an ItemID or ID or the path as string or an ID
        elseif ($ItemPath.Id)         {$ItemID = $ItemPath.id}
        elseif ($ItemPath )           {$ItemID = (Invoke-GraphRequest ("$GraphUri/drives/$Drive" + ($ItemPath -replace  '^/?(root:/)?(.*?)[:/]*$','/root:/$2:' ) )).id} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"
        elseif ($ItemID -is [string]) {$ItemID = $ItemID  -replace '^/?(items/)?(.*?)/?$', '$2' } #Allow "{id}" or "items/{id}" with any leading or trailing /

        $workBookUrl = "$GraphUri/drives/$drive/items/$ItemID/workbook/"
        try {
            if ($sheetname -notin (Invoke-GraphRequest "$workBookUrl/worksheets?`$select=name" -ValueOnly | ForEach-Object {$_.name})) {
                        $null = Invoke-GraphRequest "$workBookUrl/worksheets" -Method post  -ContentType "application/json" -Body ('{{ "name": "{0}" }}' -f $sheetName)
            }
        }
        catch { throw [System.IO.FileNotFoundException]::new('Could not connect to the given XLSX file') ; break       }

        $outValues  = New-Object System.Collections.ArrayList
        $outFormats = New-Object System.Collections.ArrayList
        $Header     = $null
        $row        = 1
        $firstTimeThru = $true
    }
    process {
        if ($null -eq $InputObject) {$row += 1}
        foreach ($TargetData in $InputObject) {
            if ($firstTimeThru) {
                $firstTimeThru = $false
                $isDataTypeValueType = ($null -eq $TargetData) -or ($TargetData -is [ValueType]) -or ($TargetData -is [string])
                if ($isDataTypeValueType ) {
                    $Header = @(".")       # dummy value to make sure we go through the "for each name in $header"
                    if (-not $Append) {$row -= 1} # By default row will be 1, it is incremented before inserting values (so it ends pointing at final row.);  si first data row is 2 - move back up 1 if there is no header .
                }
            }

            if (-not $Header) {
                $Header = $TargetData.PSObject.Properties.Name
                foreach ($exclusion in $ExcludeProperty) {$Header = $Header -notlike $exclusion}
                if ($NoHeader) { $row -= 1 }
                else {
                    $formatrow  = New-Object System.Collections.ArrayList
                    foreach ($h in $Header) {$null = $formatrow.add('General')}
                    $null = $OutValues.Add($Header)
                    $Null = $OutFormats.Add($formatrow)
                }
            }
            $row += 1

            $formatrow  = New-Object System.Collections.ArrayList
            $datarow    = New-Object System.Collections.ArrayList
            foreach ($Name in $Header) {
                if   ($isDataTypeValueType) {$v = $TargetData}
                else {$v = $TargetData.$Name}
                    if     ($v -is    [DateTime]) {
                            $null = $datarow.add($v.ToOADate()) , $formatrow.add($dateFormat)
                    }
                    elseif ($v -is    [TimeSpan]) {
                            $null = $datarow.add($v.totalDays)  , $formatrow.add('[h]:mm:ss')
                    }
                    elseif ($v -is    [System.ValueType] -or $v -is [string] -or $null -eq $v) {
                            $null = $datarow.add($v)            , $formatrow.add($null)
                    }
                    else{   $null = $datarow.add($v.tostring()) , $formatrow.add($null) }
            }
            $null = $OutValues.Add($datarow) , $OutFormats.Add($formatrow)
        }
    }
    end     {
        $dividend = $Header.Count
        $columnName = New-Object System.Collections.ArrayList
        while($dividend -gt 0) {
            $modulo    =  ($dividend -1)% 26
            $columnName.Insert(0,([char](65 + $modulo)))
            $dividend    = ($dividend - $modulo -1) /26
        }
        $range = '$A$1:${0}${1}' -f ($columnName -join ''), $row
        $body = convertto-Json -depth 5 @{values = $OutValues ;   numberFormat = $OutFormats}  -Compress

        $null = Invoke-GraphRequest "$workBookUrl/worksheets/$sheetName/range(address='$range')" -Method PATCH -ContentType "application/json" -Body $body

        if ($AsTable) {
            $null = Invoke-GraphRequest "$workBookUrl/worksheets/$sheetName/tables/`$/add" -body ('{{"address": "{0}!{1}" , "hasheaders": true}}' -f $SheetName,$range)  -ContentType "application/json" -method POST
        }
        elseif ($RangeName){
            $null = Invoke-GraphRequest "$workBookUrl/names/add" -Method POST -Body ('{{"name": "{2}",   "reference": "={0}!{1}" }}' -f $SheetName,$range,$RangeName )  -ContentType "application/json"
        }
        if ($Show) {
            Start-Process (Invoke-GraphRequest ("$workBookUrl" -replace "/workbook/?$") ).weburl
        }
    }
}

function New-GraphWorkBook     {
    param    (
        #Location file should be copied to can be in the form "/files/" to copy to users "files" folder, or "/drives/{id}/root:/folder/Subfolder" to select another drive
        [Parameter(Mandatory=$true,position=0)]
        [ArgumentCompleter([OneDriveFolderCompleter])]
        [string]$Destination,
        #The drive, by default the current user's OneDrive.
        $Drive = 'me/Drive',
        $TemplatePath
    )

    if (-not $TemplatePath) {$TemplatePath = (Join-Path $PSScriptRoot 'Blank.xlsx')}
    Copy-ToGraphFolder -Path $TemplatePath -Destination $Destination -Drive $Drive -ConflictBehavior fail -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

}
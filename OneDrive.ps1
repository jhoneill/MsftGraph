function Get-GraphDrive {
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
        The first command gets the users drive, and looks for a known folder as a child item in the drive-roo.
        This folder can't be piped into get-graphdrive, so the drive id and folder id are passed.
        In this case the drive ID could be ommitted because the default is to use the user's home drive
    #>
    [cmdletbinding(DefaultParameterSetName="None")]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Write-warning could be used, but the is informational non-output.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Drive cache is intended to be accessible outside the module.')]
    param   (
        #The drive to examine - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [parameter(ValueFromPipeline=$true)]
        $Drive = 'me/Drive',
        #if specified gets the items in a folder by the path from {drive}/root:
        [Parameter(Mandatory=$true, ParameterSetName='FolderName',Position=0)]
        [Alias("Path")]
        [String]$FolderPath,
        #If Specified gets the items in a folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='FolderID')]
        [String]$FolderID,
        [ValidateSet('Apps','Attachments','CameraRoll','Documents','Music','Photos','Public')]
        #If specified returns the subfolders - if no FolderPath or FolderID is given will return folders of the root drive
        [Parameter(Mandatory=$true, ParameterSetName='Special')]
        [String]$SpecialFolder,
        #If specified gets recent items in the drive
        [Parameter(Mandatory=$true, ParameterSetName='Recent')]
        [switch]$Recent ,
        #If Specified gets items shared with the user
        [Parameter(Mandatory=$true, ParameterSetName='Shared')]
        [switch]$SharedWithMe ,
        #Enables a free text search of the selected content
        [Parameter(ParameterSetName='RootSearch')]
        [Parameter(ParameterSetName='Shared')]
        [Parameter(ParameterSetName='FolderID')]
        [Parameter(ParameterSetName='FolderName')]
        [string]$Search,
        #If specified gets one of the special folders (Documents, photos etc) in the drive. If they don't already exist the server appears to create them.
        [Parameter(ParameterSetName='RootFolders')]
        [Parameter(ParameterSetName='FolderID')]
        [Parameter(ParameterSetName='FolderName')]
        [Parameter(ParameterSetName='None')]
        [Switch]$Subfolders,
        #if specified gets the items in a folder by the path from {drive}/root:
        [Parameter(Mandatory=$true, ParameterSetName='ItemName')]
        [String]$ItemPath,
        #If Specified gets the items in a folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='ItemID')]
        [String]$ItemID
    )
    begin   {
        Connect-MSGraph
        $webParams = @{Method  = "Get"
                       Headers = $Script:DefaultHeader
        }
    }
    process {
        #region Sort out the Drive - it might be "me/drives" (the default), "drives/drive-id", "drive-id" or a drive object with an ID.
        #       Fix up the last two; check the drive is accessible and then cache the id --> name
        if     ($Drive.id)               {$drive = "drives/$($drive.id)"}
        elseif ($Drive -notmatch './.')  {$drive = "drives/$drive"      }
        #Strip leading and trailing / from $drive so it fits in the URI template.
        $Drive = $Drive -replace '/$','' -replace '^/',''
        try   {
            if ($script:WorkOrSchool) {$uri = "https://graph.microsoft.com/v1.0/$Drive`?`$expand=root(`$expand=children)"}
            else                      {$uri = "https://graph.microsoft.com/v1.0/$Drive"} #The expand fails on consumer one drive
            $driveObj  =  (Invoke-RestMethod @webParams -uri $uri )
            $driveObj.pstypenames.Add('GraphDrive')
            $global:drivecache[$driveObj.id] = $driveObj.name
        }
        catch {
            $drive = $null
            throw ('Error trying to get drive $drive - the code was ' + $_.exception.response.statuscode.value__  ) ; return
        }
        #endregion

        #region Getting a single item (file or folder) by ID or by path.
        # Make something we can insert in a REST URI
        if     ($ItemID   -and $ItemID -Match '^/?items')   {$ItemID =  $ItemID  -replace '^/?(.*)/?$',            '$1' } #Allow "items/{id}" Strip off any leading or trailing /
        elseif ($ItemID   )                                 {$ItemID =  $ItemID  -replace '^/?(.*)/?$',      'items/$1' } #Allow "{id}". Strip off any leading or trailing / and prepend "items/"
        elseIf ($Itempath -in @("root:", "/", "root:/") )   {$ItemID = 'root' }                                           #Convert "root:", "/" or root:/" to "root"
        elseif ($ItemPath -and $ItemPath -Match '^/?root:') {$ItemID =  $ItemPath -replace '^/?(.*?)[:/]*$',       '$1:'} #Allow "[/]root:{path}" strip any leading / or trailing / or : and append ":"
        elseif ($ItemPath )                                 {$ItemID =  $ItemPath -replace '^/?(.*?)[:/]*$', 'root:/$1:'} #Allow "{path}", strip any leading / or trailing / or : and place between "root:/" and ":"
        #if we had an item ID or built an itemID string from the path, get the item, add a type and return it
        if     ($ItemID ) {
            try   {$item = Invoke-RestMethod @WebParams -Uri "https://graph.microsoft.com/v1.0/$Drive/$ItemID" }
            catch {
                if ($_.exception.response.statuscode.value__ -eq 404) {
                     Write-Warning -Message "Item Not found" ; return
                }
                #we got something other than a 404 error
                else {Write-Warning -Message $_.exception.tostring() ; return }
            }
            $item.pstypenames.add('GraphDriveItem')
            return $item
        }
        #endregion
        #region Getting collections of items either in special folders by name, normal folders by path/id, recent items, or items "shared with me".
        #if we got a folder path or ID, search for its items; first make make sure we can insert it into the URL
        if    (($search) -and -not
               ($FolderID   -or $FolderPath) )                  {$FolderID = 'root'}                                               #If We were asked to search but not told where, choose "root"
        elseif ($FolderID  -and $FolderID -Match '^/?items')    {$FolderID =  $FolderID   -replace '^/?(.*)/?$',             '$1'} #Other processing mirrors items above.
        elseif ($FolderID )                                     {$FolderID =  $FolderID   -replace '^/?(.*)/?$',       'items/$1'}
        elseIf ($FolderPath -in @("root:", "/", "root:/") )     {$FolderID = 'root' }
        elseif ($FolderPath -and $FolderPath -Match '^/?root:') {$FolderID =  $FolderPath -replace '^/?(.*?)[:/]*$',       '$1:' }
        elseif ($FolderPath )                                   {$FolderID =  $FolderPath -replace '^/?(.*?)[:/]*$', 'root:/$1:' }
        elseif ($SpecialFolder)                                 {$FolderID = "special/$SpecialFolder"                            }

        if ($FolderID -or $SharedWithMe -or $Recent) {
            if     ($FolderID -and $Search)     {$webParams['URI']=  "https://graph.microsoft.com/v1.0/$Drive/$FolderID/search(q='$search')?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file"}
            elseif ($FolderID             )     {$webParams['URI'] = "https://graph.microsoft.com/v1.0/$Drive/$FolderID/children?`$Select=Name,Id,folder,Size,Weburl,specialfolder,parentReference,fileSystemInfo,folder,file" }
            elseif ($SharedWithMe -and $search) {}  #can these be combined ?
            elseif ($SharedWithMe             ) {$webParams['URI'] = "https://graph.microsoft.com/v1.0/me/Drive/SharedWithMe"        }
            elseif ($Search                   ) {$webParams['URI'] = "https://graph.microsoft.com/v1.0/me/drive/search(q='$Search')" }  #me or $drive
            elseif ($Recent                   ) {$webParams['URI'] = "https://graph.microsoft.com/v1.0/$Drive/recent"                }  #Me or $drive
            try    {$children = (Invoke-RestMethod @WebParams).value }
            catch  {
                    if ($_.exception.response.statuscode.value__ -eq 404) {
                          Write-Warning -Message "Not found" ;return
                    }
                    else {Write-Warning -Message $_.exception.tostring() ; return}
            }
            if ($Subfolders) {$children.where({$_.folder}) | Sort-Object -Property name}
        }
        #endregion
        #region Getting the drive - either the drive object itself , or the folders in its root.
        elseif ($Subfolders) {
            $children = $driveObj.root.children.where({$_.folder}) | Sort-Object -Property Name
        }
        else             {
            foreach ($c in $driveObj.children) {$c.pstypenames.Add("GraphDriveItem")}
            return $driveObj
        }
        #endregion

        #The above will either have left a collection of items in $children, or explictly returned a result.
        #region return any collection of items - filtered to subfolders if required. Tell the user if the folder is empty but send nothing to the pipeline
        if (-not $children) { Write-Host  "Folder exists, but is empty."}
        else  {
                foreach ($c in $children ) {$c.pstypenames.Add("GraphDriveItem")  }

                $children  | Sort-Object -Property @{e={$null -eq $_.folder}},name
        }
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

function New-GraphFolder {
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
    param(
        #The name for the new folder
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)]
        [Alias('FolderPath')]
        [string]$Path,
        #The drive hold the new folder - defaults to the user's OneDrive but can be a shared one e.g. Drives/{ID}
        [Parameter()]
        $Drive = 'me/Drive'
    )
    begin {
        Connect-MSGraph
        $webParams = @{Headers= $Script:DefaultHeader }

        #  Sort out the Drive - it might be "me/drives" (the default), "drives/drive-id", "drive-id" or a drive object with an ID.
        #       Fix up the last two; check the drive is accessible and then cache the id --> name
        if     ($Drive.id)               {$drive = "drives/$($drive.id)"}
        elseif ($Drive -notmatch './.')  {$drive = "drives/$drive"      }
        #Strip leading and trailing / from $drive so it fits in the URI template.
        $Drive = $Drive -replace '/$','' -replace '^/',''
        try   {
            $driveObj  =  (Invoke-RestMethod @webParams -Method GET -Uri "https://graph.microsoft.com/v1.0/$Drive")
            $global:drivecache[$driveObj.id] = $driveObj.name
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

        $body = ConvertTo-Json $settings
        Write-Debug $body
        if ($PSCmdlet.ShouldProcess($parentPath, "Create new OneDrive folder '$($settings.Name)'")) {
            try {
                $newFolder = Invoke-RestMethod @webParams -Method Post -uri "https://graph.microsoft.com/v1.0/$Drive/$parentPath/children" -ContentType "application/json" -Body $body
                $newFolder.Pstypenames.add("GraphDriveItem")
                return $newFolder
            }
            Catch {
                if ($_.exception.response.statuscode.value__ -eq 409) {
                    Write-Warning -Message "A Confilict error was returned. The folder probably exists already"
                }
                else {throw $_ }
            }
        }
    }
}

function Show-GraphFolder {
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
    param(
        #If Specified gets the  folder by folder ID
        [Parameter(Mandatory=$true, ParameterSetName='FolderName')]
        [Alias("FolderPath")]
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

function Copy-ToGraphFolder {
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
    param(
        #location of file on the local machine
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Path,
        #Location file should be copied to can be in the form "/files/" to copy to users "files" folder, or "/drives/{id}/root:/folder/Subfolder" to select another drive
        [Parameter(Mandatory=$true)]
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

    begin  {
        Connect-MSGraph
    }
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
                Write-Verbose -Message "Selected content type of $contentType for a $($UploadItem.Extension) file."
            }
            else {$ContentType = "application/octet-stream"}
        }

        #region Figure out what the URI should be
        #was destination writen out in full as drives/{id}/root:/{path} ? (with or without a leading /)
        if     ($Destination.parentReference -and
                $Destination.fileSystemInfo)                   { $uri = "https://graph.microsoft.com/v1.0/drives/"+ $Destination.parentReference.driveId + "/items/" + $Destination.id}
        elseif ($Destination -match '/?drives.*:/\w')           {$uri = 'https://graph.microsoft.com/v1.0/' + ($Destination -replace '^/','') }
        else { #We didn't get the drive in the destination, so is it an object, a partial path "drives/id"  or "me/drive", or just an ID
            if     ($Drive.id)                                  {$uri  = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/"}
            elseif ($Drive -match './.')                        {$uri  = "https://graph.microsoft.com/v1.0/$drive/"             }
            elseif ($Drive)                                     {$uri  = "https://graph.microsoft.com/v1.0/drives/$drive/"      }
            # the root might be "/" root: or root:/ (/root will be assumed to be a folder) anywhere else we can bolt on to the URI. We may need to put root: in front and strip leading /
            If     ($Destination -in @("root:", "/", "root:/")) {$uri += "root/"                                       }
            elseif ($Destination -Match '^/?root:')             {$uri += ($Destination -replace '^/', '')              }
            else                                                {$uri += ($Destination -replace '^/?(.*$)', 'root:/$1')}
        }

        #if URI ends with / so that's a directory, easy. .
        if    ($uri -match '/$') {$uri = $uri + $uploadItem.Name }
        else  {     # Otherwise see if we have a directory, or a file path with a valid parent
            try   { # Does the Path exist ? Try to get it and catch the 404 error that will result if destination points to a new file
                $x = Invoke-RestMethod -Method get -Headers $Script:DefaultHeader -Uri $uri
                #The path exists ... is it a folder or a file  ?
                if ($x.folder) {
                      $uri = $uri +'/' + $uploadItem.Name
                      Write-Verbose -Message "$Destination appears to be a folder, will upload to a file named $($uploadItem.Name) in it."
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
                    Write-Verbose -Message "$uri was not found,  checking for $folderURI"
                    try {
                        $x = Invoke-RestMethod -Method get -Headers $Script:DefaultHeader -Uri $folderuri
                        if ($x.folder) {
                                $settings['item'].name = $uri -replace '^.*/','' #get rid of everything up to the last slash (greedy regex)
                                Write-Verbose -Message "$folderURI is a valid folder; will upload as a new file."
                        }
                        else  { Write-Warning -Message "There was a problem with $Destination as a target path. Neither it nor its parent look like valid folders."; return}
                    }
                    Catch     { Write-Warning -Message "There was a problem with $Destination as a target path. Neither it nor its parent look like valid folders."; return}
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
                $result             = Invoke-RestMethod -Method Put  -headers @{Authorization = "Bearer $AccessToken"} -Uri ($uri + ":/Content") -InFile $uploadItem.FullName -ContentType $ContentType
            }
            else {
                $body               = ConvertTo-Json $settings
                try   {
                    $UploadSession  = Invoke-RestMethod -Method Post -headers @{Authorization = "Bearer $AccessToken"} -Uri ($uri + ":/createUploadSession") -Body $body  -ContentType "application/json"
                }
                catch {
                    if ($_.exception.response.statuscode.value__ -eq 409) {
                        Write-Warning -Message "Uploading to $Destination responed 'Conflict'. This is expected if you chose 'Conflict FAIL' and the file exists" ; return
                    }
                    #we got something other than a conflict error
                    else          { Write-Warning -Message $_.exception.tostring() ; return}
                }
                if (-not $UploadSession.uploadUrl) {Write-Warning -Message 'Server did not provide an upload destination' ; return}
                else                               {Write-Verbose -Message "Have an upload session until $($SessionConnection.expirationDateTime)" }
                $oldprogressPref    = $ProgressPreference
                $ProgressPreference = 'SilentlyContinue'
                $result             = Invoke-RestMethod -Method Put -Uri $UploadSession.uploadUrl -InFile $uploadItem.FullName -ContentType "application/octet-stream" -Headers @{"Content-Range"=$RangeText}
                $ProgressPreference =$oldprogressPref
            }
            $result.pstypenames.Add("GraphDriveItem")
            return $result
        }
    }
}

function Copy-FromGraphFolder {
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
    param(
        #The path to the file on one drive
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $Path,
        #The drive, by default the current user's OneDrive.
        $Drive = 'me/Drive',
        #The destination on the local computer
        $Destination = $pwd,
        #If specified prevents an existing file from being overwritten.
        [Switch]$NoClobber,
        #If Specified the destination file will be returned (similar to Copy-Item)
        [Alias('PT')]
        [Switch]$Passthru
    )
    begin    {
                 Connect-MSGraph
                 $webParams = @{Method  = "Get"
                               Headers = $Script:DefaultHeader
        }
    }
    process  {
        <#
        We can download from
        /drives/{drive-id}/items/{parent-id}:/{filename}
        /drives/{drive-id}/root:/{folder-path}/{filename}
        /groups/{group-id}/drive/items/{parent-id}:/{filename}
        /me/drive/items/{parent-id}:/{filename}
        /me/drive/root:/{folder-path}/{filename}
        /sites/{site-id}/drive/items/{parent-id}:/{filename}
        /users/{user-id}/drive/items/{parent-id}:/{filename}
        #>
        if ($Path.name -and $Path.'@microsoft.graph.downloadUrl') {$sourceDetails = $Path}
        else {
            if     ($Path.parentReference -and
                    $Path.fileSystemInfo         ) {$webparams['uri']  = "https://graph.microsoft.com/v1.0/drives/"+ $Path.parentReference.driveId + "/items/" + $path.id}
            elseif ($Path -isnot [string]        ) {throw 'An invalid Object was passed as a Path Parameter' ; return }
            elseif ($Path -match '/?drives.*:/\w') {$webparams['uri']  = 'https://graph.microsoft.com/v1.0/' + ($path -replace '^/','') }
            else { #We didn't get the drive in the destination, so is it an object, a partial path "drives/id"  or "me/drive", or just an ID
                if     ($Drive.id)                 {$webparams['uri']  = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/"}
                elseif ($Drive -match './.')       {$webparams['uri']  = "https://graph.microsoft.com/v1.0/$drive/"             }
                elseif ($Drive)                    {$webparams['uri']  = "https://graph.microsoft.com/v1.0/drives/$drive/"      }
                if     ($path -Match '^/?root:')   {$webparams['uri'] += ($path -replace '^/', '')              }
                else                               {$webparams['uri'] += ($path -replace '^/?(.*$)', 'root:/$1')}
            }

            #Get the item. The result should have a downloadURL as property.
            try   {$sourceDetails  = Invoke-RestMethod @webParams }
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
                Invoke-WebRequest -Method get -Uri $sourceDetails.'@microsoft.graph.downloadUrl' -OutFile $Destination
              if ($Passthru) {Get-Item -Path  $Destination}}
        }
        elseif (-not $sourceDetails.'@microsoft.graph.downloadUrl') {
              Write-Warning -Message "Could not get the download url for $path"
        }
        else {Write-Warning -Message "$Destination is not a valid path."}
    }
}

function FileCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $Drive       = $fakeBoundParameter['Drive']
     if     (-not $Drive -or $drive -eq "/")             {$uri =   'https://graph.microsoft.com/v1.0/me/Drive'}
    elseif ($Drive.id)               {$uri =   "https://graph.microsoft.com/v1.0/drives/$($Drive.id)"}
    elseif ($Drive -notmatch './.')  {$uri =   'https://graph.microsoft.com/v1.0/drives/' + $Drive -replace '/$','' -replace '^/' }
    else                             {$uri =   'https://graph.microsoft.com/v1.0/drives/' + $Drive -replace '/$','' -replace '^/' }
    #strip quotes from word to complete - replace " or ' with nothing
    $wordToComplete = $wordToComplete -replace '"|''', ''


    #if it is not */something/* (and that includes nothing, / or root:/) or if it is /root:/ or /root/   use "/root"                (no :path:)
    #if it is root:/something/more   or /root:/something/more  or /something/more  ......................use "/root:/something:"    (with:path ignore a part-completed final item)
    #if it is root:/something/more/  or /root:/something/more/ or /something/more/ --------------------- use /root:/something/more: (with:path just drop the final /)
    If     ($wordToComplete -notmatch "/.+/" -or $wordToComplete -eq "/root:?/" ) {$uri +=  '/root' }
    elseif ($wordToComplete -Match '^/?root:')                                    {$Uri +=  $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:' } #catch after any leading / and before final /; and sandwich between / and :
    else                                                                          {$uri +=  $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:' } #catch after any leading / and before final /; and sandwich between /root/ and :

    #So the uri is now either /root   or /root:/{path}: where path is a complete folder name. we its children, but only the folders, only a couple of columns, in name order
    $uri +=    '/children?$select=Name,ParentReference'

    (Invoke-RestMethod -Method get -headers @{Authorization = "Bearer $AccessToken"} -Uri $uri ).value | Sort-Object -Property Name | #it would be better to order-by at the server, but consumer one drive doesn't support it.
        ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$p'", $p, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $p
            }
        }
}

function FolderCompletion {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $Drive       = $fakeBoundParameter['Drive']
    if     (-not $Drive)             {$uri =   'https://graph.microsoft.com/v1.0/me/Drive'}
    elseif ($Drive.id)               {$uri =   "https://graph.microsoft.com/v1.0/drives/$($Drive.id)"}
    elseif ($Drive -notmatch './.')  {$uri =   'https://graph.microsoft.com/v1.0/drives/' + $Drive -replace '/$','' -replace '^/' }
    else                             {$uri =   'https://graph.microsoft.com/v1.0/drives/' + $Drive -replace '/$','' -replace '^/' }
    #strip quotes from word to complete - replace " or ' with nothing
    $wordToComplete = $wordToComplete -replace '"|''', ''

    #if it is not */something/* (and that includes nothing, / or root:/) or if it is /root:/ or /root/   use "/root"                (no :path:)
    #if it is root:/something/more   or /root:/something/more  or /something/more  ......................use "/root:/something:"    (with:path ignore a part-completed final item)
    #if it is root:/something/more/  or /root:/something/more/ or /something/more/ --------------------- use /root:/something/more: (with:path just drop the final /)
    If     ($wordToComplete -notmatch "/.+/" -or $wordToComplete -eq "/root:?/" ) {$uri +=  '/root' }
    elseif ($wordToComplete -Match '^/?root:')                                    {$Uri +=  $wordToComplete -replace '^/?(.*)/.*?$',      '/$1:' } #catch after any leading / and before final /; and sandwich between / and :
    else                                                                          {$uri +=  $wordToComplete -replace '^/?(.*)/.*?$','/root:/$1:' } #catch after any leading / and before final /; and sandwich between /root/ and :

    #So the uri is now either /root   or /root:/{path}: where path is a complete folder name. we its children, but only the folders, only a couple of columns, in name order
    $uri +=    '/children?$filter=folder ne null&$select=Name,ParentReference'

    (Invoke-RestMethod -Method get -headers @{Authorization = "Bearer $AccessToken"} -Uri $uri ).value | Sort-Object -Property Name |  #it would be better to order-by at the server, but consumer one drive doesn't support it.
        ForEach-Object {
            $P = ($_.parentReference.path -replace "/drive/|/drives/.*?/","" ) + "/" + $_.name
            if ($P -like "*$wordToComplete*") {
                New-Object -TypeName System.Management.Automation.CompletionResult -ArgumentList "'$p'", $p, ([System.Management.Automation.CompletionResultType]::ParameterValue) , $p
            }
        }
}
#In PowerShell 3 and 4 Register-ArgumentCompleter is part of TabExpansion ++. From V5 it is part of Powershell.core
if (Get-Command -ErrorAction SilentlyContinue -name Register-ArgumentCompleter) {
 Register-ArgumentCompleter -CommandName 'Copy-FromGraphFolder' -ParameterName 'Path'        -ScriptBlock $Function:FileCompletion
 Register-ArgumentCompleter -CommandName 'Get-GraphDrive'       -ParameterName 'ItemPath'    -ScriptBlock $Function:FileCompletion
 Register-ArgumentCompleter -CommandName 'Get-GraphDrive'       -ParameterName 'FolderPath'  -ScriptBlock $Function:FolderCompletion
 Register-ArgumentCompleter -CommandName 'New-GraphFolder'      -ParameterName 'Path'        -ScriptBlock $Function:FolderCompletion
 Register-ArgumentCompleter -CommandName 'Show-GraphFolder'     -ParameterName 'Path'        -ScriptBlock $Function:FolderCompletion
 Register-ArgumentCompleter -CommandName 'Copy-ToGraphFolder'   -ParameterName 'Destination' -ScriptBlock $Function:FolderCompletion
}

$boundary='MyAppPartBoundary'
$myhtml     =  @"
--$boundary
Content-Disposition:form-data; name="Presentation"
Content-type:text/html

<!DOCTYPE html>
<html>
 <head><title>Page with image 9</title></head>
 <body>
        <img width="500" src="name:MyAppFileBlockName" />
 </body>
</html>

"@

$content = [byte[]][char[]]$myhtml 

#foreach attachment , set mimetype and name,  
$Content += [byte[]][char[]]@"
--$boundary
Content-Disposition:form-data; name="MyAppFileBlockName"
Content-type:image/jpeg


"@
$Content += [System.IO.File]::ReadAllBytes((Resolve-Path .\upload.jpg))
$Content += ([byte[]][char[]]"`r`n--$boundary--`r`n")

$NewPage = Add-GraphOneNotePage -Section $notebook.sections[0]  -ContentType "multipart/form-data; boundary=$boundary" -HTMLPage $content -PassThru

start $NewPage.links.oneNoteWebUrl.href
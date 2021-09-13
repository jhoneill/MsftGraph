push-location C:\Users\P10506111\downloads\msgraph-sdk-powershell-dev\openApiDocs\v1.0\
  $lines = (Get-Content -raw *.yml | Measure-Object -line).Lines
  $files = (Get-ChildItem *.yml | Measure-Object).Count
  $paths = @{}
  $schemas  = @{}
  $methods  = @()
  Get-ChildItem *.yml | ForEach-Object {
     $y = ConvertFrom-Yaml (Get-Content $_ -Raw)
     foreach ($k in $y.paths.keys) {$paths[$k]= $true ; $methods += $y.paths[$k].Keys}
     foreach ($k in $y.components.schemas.keys) {$schemas[$k]= $true}
  }
    "$files files, of $lines lines defining $($schemas.keys.count) objects and $($paths.Keys.Count) RestAPI Paths"
    $methods | Group-Object -NoElement
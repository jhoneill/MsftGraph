Import-Module Microsoft.Graph.PlusPlus
Connect-Graph -fromazure <#This will fail#>
Connect-AzAccount <#connect to azure and try again#>
Connect-Graph -FromAzureSession
GWhoAmI
Get-Alias gwhoami
Get-GraphUser
Get-GraphUser | ft Organization # shows the organization property set
Get-GraphGroup 'Accounts'  # an existing teams-enabled group
Get-GraphGroup 'Accounts' -Drive # #Some things don't work with an azure token.
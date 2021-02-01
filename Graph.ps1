#Most functions for this are in files based on the application where they would surface
# OneDrive, OneNote, Outlook-Calendar, Outlook-Contacts, Outlook-Mail, Planner, SharePoint, Teams.
# Those in this file don't belong to an application.

<#
 Others to explore
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directoryRoles").value | ft id,displayname,description               https://docs.microsoft.com/en-us/graph/api/directoryrole-list?view=graph-rest-1.0
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/directoryRoleTemplates").value | sort displayname |  ft id,displayname,description
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/groupsettingTemplates").value | ft displayname,description -wrap -aut
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices").value | ft approximateLastSignInDateTime,displayName,operatingsystem,operatingsystemversion
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices/c2221377-d362-42e7-8e16-e7d6abf80e61/registeredOwners").value
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/devices/c2221377-d362-42e7-8e16-e7d6abf80e61/memberof").value
 (irm -Method Get -headers $Script:DefaultHeader -Uri "https://graph.microsoft.com/v1.0/me/owneddevices").value | ft displayname,operatingsystemversion,trusttype
#>

#Get-PSCallStack | Out-File -Append ~\graph.txt

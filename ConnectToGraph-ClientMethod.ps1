#Documentation:
# https://docs.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0&tabs=http

# Application (client) ID, tenant Name and secret
$clientId = "13da4e98-1f6f-4e01-b6ae-68b79d477669"
$tenantName = "M365x298804.onmicrosoft.com"
$clientSecret = "pq2BH20~c12H~SIbRX9I5~_5Ai_1R~Cn-d"
$resource = "https://graph.microsoft.com/"


$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 

$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

#SAMPLES:

#1) Get All Groups:

#$apiUrl = 'https://graph.microsoft.com/v1.0/Groups/'
#$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
#$Groups = ($Data | select-object Value).Value
#$Groups | Format-Table DisplayName, Description -AutoSize

#2) List All Teams:

#$apiUrl = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
#$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
#$Teams = ($Data | select-object Value).Value | Select-Object displayName, id, description
#$Teams

#3) List All Team Channels:

#$apiUrl = "https://graph.microsoft.com/beta/teams/d386ef45-b535-4a60-a683-4fd9a55ca4d8/channels"
#$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
#$TeamChannels = ($Data | select-object Value).Value
#$TeamChannels

#4) Post to Channel:
#$apiUrl = "https://graph.microsoft.com/beta/teams/d386ef45-b535-4a60-a683-4fd9a55ca4d8/channels/19:51454a08de90466db714a0d3e8efeb90@thread.tacv2/messages"
#$body = @"
#{ 
#"body": { 
#    "content": "Hello Management team! This is being sent to you from PowerShell!" 
#     } 
# }
#"@

#Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Body $body -Method Post -ContentType 'application/json'

#5) Create a Group:
#$apiUrl = "https://graph.microsoft.com/v1.0/groups"

#$body = @"
#{
#  "description": "Self help community for library",
#  "displayName": "Library Assist",
#  "groupTypes": [
#    "Unified"
#  ],
#  "mailEnabled": true,
#  "mailNickname": "library",
#  "securityEnabled": false
#}
#"@
#Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Body $body -Method Post -ContentType 'application/json'

#6) List channel messages

#$teamsid= "d386ef45-b535-4a60-a683-4fd9a55ca4d8";
#$channelid= "19:51454a08de90466db714a0d3e8efeb90@thread.tacv2";

#$apiUrl = "https://graph.microsoft.com/beta/teams/" + $teamsid + "/channels/" + $channelid + "/messages"

#$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get

#$Messages = ($Data | select-object Value).Value | Select-Object id, messageType, createdDateTime, importance, mentions

#$Messages

#$ListOfApps | export-csv -notypeinformation -encoding 'utf8' -path ('.\TeamsApps.csv')


#7) Get Members of a Team

$teamsid = "d386ef45-b535-4a60-a683-4fd9a55ca4d8";

$apiUrl = "https://graph.microsoft.com/v1.0/groups/" + $teamsid  + "/members";

$Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get

$members = ($Data | select-object Value).Value | Select-Object id,displayName,jobTitle,mail,mobilePhone,officeLocation,userPrincipalName,businessPhones

$members | export-csv -notypeinformation -encoding 'utf8' -path ('.\TeamsMembers.csv')
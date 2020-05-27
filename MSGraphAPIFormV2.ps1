<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    MSGraphAPIForms
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MicrosoftGraphAPIQueries        = New-Object system.Windows.Forms.Form
$MicrosoftGraphAPIQueries.ClientSize  = '986,792'
$MicrosoftGraphAPIQueries.text   = "Microsoft Graph API Queries Form"
$MicrosoftGraphAPIQueries.TopMost  = $false

$GbTenantSettings                = New-Object system.Windows.Forms.Groupbox
$GbTenantSettings.height         = 173
$GbTenantSettings.width          = 445
$GbTenantSettings.text           = "Tenant Settings"
$GbTenantSettings.location       = New-Object System.Drawing.Point(33,38)

$LblClientId                     = New-Object system.Windows.Forms.Label
$LblClientId.text                = "Client Id:"
$LblClientId.AutoSize            = $true
$LblClientId.width               = 25
$LblClientId.height              = 10
$LblClientId.location            = New-Object System.Drawing.Point(31,34)
$LblClientId.Font                = 'Microsoft Sans Serif,10'

$LblTenantName                   = New-Object system.Windows.Forms.Label
$LblTenantName.text              = "Tenant Name: "
$LblTenantName.AutoSize          = $true
$LblTenantName.width             = 25
$LblTenantName.height            = 10
$LblTenantName.location          = New-Object System.Drawing.Point(30,65)
$LblTenantName.Font              = 'Microsoft Sans Serif,10'

$LblClientSecret                 = New-Object system.Windows.Forms.Label
$LblClientSecret.text            = "Client Secret"
$LblClientSecret.AutoSize        = $true
$LblClientSecret.width           = 25
$LblClientSecret.height          = 10
$LblClientSecret.location        = New-Object System.Drawing.Point(31,97)
$LblClientSecret.Font            = 'Microsoft Sans Serif,10'

$TxtClientId                     = New-Object system.Windows.Forms.TextBox
$TxtClientId.multiline           = $false
$TxtClientId.text                = "13da4e98-1f6f-4e01-b6ae-68b79d477669"
$TxtClientId.width               = 287
$TxtClientId.height              = 20
$TxtClientId.location            = New-Object System.Drawing.Point(148,32)
$TxtClientId.Font                = 'Microsoft Sans Serif,10'

$TxtClientSecret                 = New-Object system.Windows.Forms.MaskedTextBox
$TxtClientSecret.multiline       = $false
$TxtClientSecret.text            = "pq2BH20~c12H~SIbRX9I5~_5Ai_1R~Cn-d"
$TxtClientSecret.width           = 285
$TxtClientSecret.height          = 20
$TxtClientSecret.location        = New-Object System.Drawing.Point(148,93)
$TxtClientSecret.Font            = 'Microsoft Sans Serif,10'

$TxtTenantName                   = New-Object system.Windows.Forms.TextBox
$TxtTenantName.multiline         = $false
$TxtTenantName.text              = "M365x298804.onmicrosoft.com"
$TxtTenantName.width             = 285
$TxtTenantName.height            = 20
$TxtTenantName.location          = New-Object System.Drawing.Point(148,60)
$TxtTenantName.Font              = 'Microsoft Sans Serif,10'

$LblResource                     = New-Object system.Windows.Forms.Label
$LblResource.text                = "Resource"
$LblResource.AutoSize            = $true
$LblResource.width               = 25
$LblResource.height              = 10
$LblResource.location            = New-Object System.Drawing.Point(34,128)
$LblResource.Font                = 'Microsoft Sans Serif,10'

$TxtResource                     = New-Object system.Windows.Forms.TextBox
$TxtResource.multiline           = $false
$TxtResource.text                = "https://graph.microsoft.com/"
$TxtResource.width               = 284
$TxtResource.height              = 20
$TxtResource.enabled             = $false
$TxtResource.location            = New-Object System.Drawing.Point(148,123)
$TxtResource.Font                = 'Microsoft Sans Serif,10'

$GbMSTeamsQueries                = New-Object system.Windows.Forms.Groupbox
$GbMSTeamsQueries.height         = 151
$GbMSTeamsQueries.width          = 933
$GbMSTeamsQueries.text           = "Microsoft Teams Queries"
$GbMSTeamsQueries.location       = New-Object System.Drawing.Point(33,251)

$BtnGetData                      = New-Object system.Windows.Forms.Button
$BtnGetData.text                 = "Get Data"
$BtnGetData.width                = 171
$BtnGetData.height               = 30
$BtnGetData.location             = New-Object System.Drawing.Point(752,19)
$BtnGetData.Font                 = 'Microsoft Sans Serif,10'

$RbtnListTeams                   = New-Object system.Windows.Forms.RadioButton
$RbtnListTeams.text              = "List All MS Teams"
$RbtnListTeams.AutoSize          = $true
$RbtnListTeams.width             = 104
$RbtnListTeams.height            = 20
$RbtnListTeams.location          = New-Object System.Drawing.Point(27,23)
$RbtnListTeams.Font              = 'Microsoft Sans Serif,10'
$RbtnListTeams.Checked           = $true

$RbtnListChannels                = New-Object system.Windows.Forms.RadioButton
$RbtnListChannels.text           = "List All Team Channels"
$RbtnListChannels.AutoSize       = $true
$RbtnListChannels.width          = 104
$RbtnListChannels.height         = 20
$RbtnListChannels.location       = New-Object System.Drawing.Point(27,49)
$RbtnListChannels.Font           = 'Microsoft Sans Serif,10'

$RbtnGetMembers                  = New-Object system.Windows.Forms.RadioButton
$RbtnGetMembers.text             = "Get Members of a MS Team"
$RbtnGetMembers.AutoSize         = $true
$RbtnGetMembers.width            = 104
$RbtnGetMembers.height           = 20
$RbtnGetMembers.location         = New-Object System.Drawing.Point(27,73)
$RbtnGetMembers.Font             = 'Microsoft Sans Serif,10'

$GbTeamsSettings                 = New-Object system.Windows.Forms.Groupbox
$GbTeamsSettings.height          = 171
$GbTeamsSettings.width           = 480
$GbTeamsSettings.text            = "MS Teams Settings"
$GbTeamsSettings.location        = New-Object System.Drawing.Point(488,40)

$DgvTeams                    = New-Object system.Windows.Forms.DataGridView
$DgvTeams.width              = 931
$DgvTeams.height             = 340
$DgvTeams.location           = New-Object System.Drawing.Point(35,421)
$DgvTeams.ColumnCount        = 6
$DgvTeams.Visible            = $false

$LblTeamsId                      = New-Object system.Windows.Forms.Label
$LblTeamsId.text                 = "Teams Id"
$LblTeamsId.AutoSize             = $true
$LblTeamsId.width                = 25
$LblTeamsId.height               = 10
$LblTeamsId.location             = New-Object System.Drawing.Point(15,28)
$LblTeamsId.Font                 = 'Microsoft Sans Serif,10'

$LblChannelId                    = New-Object system.Windows.Forms.Label
$LblChannelId.text               = "Channel Id"
$LblChannelId.AutoSize           = $true
$LblChannelId.width              = 25
$LblChannelId.height             = 10
$LblChannelId.location           = New-Object System.Drawing.Point(17,61)
$LblChannelId.Font               = 'Microsoft Sans Serif,10'

$TxtTeamsId                      = New-Object system.Windows.Forms.TextBox
$TxtTeamsId.multiline            = $false
$TxtTeamsId.text                 = "d386ef45-b535-4a60-a683-4fd9a55ca4d8"
$TxtTeamsId.width                = 352
$TxtTeamsId.height               = 20
$TxtTeamsId.location             = New-Object System.Drawing.Point(117,23)
$TxtTeamsId.Font                 = 'Microsoft Sans Serif,10'

$TxtChannelId                    = New-Object system.Windows.Forms.TextBox
$TxtChannelId.multiline          = $false
$TxtChannelId.text               = "19:51454a08de90466db714a0d3e8efeb90@thread.tacv2"
$TxtChannelId.width              = 352
$TxtChannelId.height             = 20
$TxtChannelId.location           = New-Object System.Drawing.Point(117,55)
$TxtChannelId.Font               = 'Microsoft Sans Serif,10'

$RbtnListTabs                    = New-Object system.Windows.Forms.RadioButton
$RbtnListTabs.text               = "List Tabs in a Channel"
$RbtnListTabs.AutoSize           = $true
$RbtnListTabs.width              = 104
$RbtnListTabs.height             = 20
$RbtnListTabs.location           = New-Object System.Drawing.Point(285,19)
$RbtnListTabs.Font               = 'Microsoft Sans Serif,10'

$RbtnListApps                    = New-Object system.Windows.Forms.RadioButton
$RbtnListApps.text               = "List installed apps in a Team"
$RbtnListApps.AutoSize           = $true
$RbtnListApps.width              = 104
$RbtnListApps.height             = 20
$RbtnListApps.location           = New-Object System.Drawing.Point(284,45)
$RbtnListApps.Font               = 'Microsoft Sans Serif,10'

$RbtnListItems                   = New-Object system.Windows.Forms.RadioButton
$RbtnListItems.text              = "List Items in a Drive"
$RbtnListItems.AutoSize          = $true
$RbtnListItems.width             = 104
$RbtnListItems.height            = 20
$RbtnListItems.location          = New-Object System.Drawing.Point(285,71)
$RbtnListItems.Font              = 'Microsoft Sans Serif,10'

$BtnExportToExcel                = New-Object system.Windows.Forms.Button
$BtnExportToExcel.text           = "Export to Excel"
$BtnExportToExcel.width          = 172
$BtnExportToExcel.height         = 30
$BtnExportToExcel.location       = New-Object System.Drawing.Point(752,64)
$BtnExportToExcel.Font           = 'Microsoft Sans Serif,10'

$MicrosoftGraphAPIQueries.controls.AddRange(@($GbTenantSettings,$GbMSTeamsQueries,$GbTeamsSettings,$DgvTeams))
$GbTenantSettings.controls.AddRange(@($LblClientId,$LblTenantName,$LblClientSecret,$TxtClientId,$TxtClientSecret,$TxtTenantName,$LblResource,$TxtResource))
$GbMSTeamsQueries.controls.AddRange(@($BtnGetData,$RbtnListTeams,$RbtnListChannels,$RbtnGetMembers,$RbtnListTabs,$RbtnListApps,$RbtnListItems,$BtnExportToExcel))
$GbTeamsSettings.controls.AddRange(@($LblTeamsId,$LblChannelId,$TxtTeamsId,$TxtChannelId))

#----------------------FUNCTIONS-------------------
function TeamsQueries {
    
    #Getting Tenant Settings
    $clientId = $TxtClientId.Text
    $tenantName = $TxtTenantName.Text
    $clientSecret = $TxtClientSecret.Text
    $resource = $TxtResource.Text

    #Getting Client Token
    $ReqTokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        client_Id     = $clientId
        Client_Secret = $clientSecret
    } 

    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

    #MS Team Queries
    
    #1)List All MS Teams 
    if($RbtnListTeams.Checked){
        $apiUrl = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object displayName, id, description

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "displayName";
        $DgvTeams.Columns[1].Name = "id";
        $DgvTeams.Columns[2].Name = "description";
        $DgvTeams.Columns["displayName"].Width = 300;
        $DgvTeams.Columns["id"].Width = 300;
        $DgvTeams.Columns["description"].Width = 300;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.displayName,$row.id,$row.description)
        }

        $DgvTeams.Visible = $true
    }
    #2) List All Team Channels
    elseif($RbtnListChannels.Checked){
        #Variables
        $teamsid = $TxtTeamsId.Text;

        $apiUrl = "https://graph.microsoft.com/beta/teams/" + $teamsid + "/channels"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $TeamChannels = ($Data | select-object Value).Value | Select-Object displayName,description,webUrl,membershipType,id

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $TeamChannels
        
        $DgvTeams.Columns[0].Name = "displayName";
        $DgvTeams.Columns[1].Name = "id";
        $DgvTeams.Columns[2].Name = "description";
        $DgvTeams.Columns[3].Name = "webUrl";
        $DgvTeams.Columns[4].Name = "membershipType";
        $DgvTeams.Columns["displayName"].Width = 300;
        $DgvTeams.Columns["id"].Width = 450;
        $DgvTeams.Columns["description"].Width = 600;
        $DgvTeams.Columns["webUrl"].Width = 600;
        $DgvTeams.Columns["membershipType"].Width = 100;
        
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.displayName,$row.id,$row.description,$row.webUrl,$row.membershipType)
        }

        $DgvTeams.Visible = $true
        
    }
    #3) Get Team Members
    elseif($RbtnGetMembers.Checked){
        #Variables
        $teamsid = $TxtTeamsId.Text;

        $apiUrl = 'https://graph.microsoft.com/v1.0/groups/' + $teamsid + '/members'
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $MembersOfTeam = ($Data | select-object Value).Value | Select-Object displayName, jobTitle, mail, mobilePhone, officeLocation, userPrincipalName
        
        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $MembersOfTeam
        
        $DgvTeams.Columns[0].Name = "displayName";
        $DgvTeams.Columns[1].Name = "jobTitle";
        $DgvTeams.Columns[2].Name = "mail";
        $DgvTeams.Columns[3].Name = "mobilePhone";
        $DgvTeams.Columns[4].Name = "officeLocation";
        $DgvTeams.Columns[5].Name = "userPrincipalName";
        $DgvTeams.Columns["displayName"].Width = 200;
        $DgvTeams.Columns["jobTitle"].Width = 200;
        $DgvTeams.Columns["mail"].Width = 350;
        $DgvTeams.Columns["mobilePhone"].Width = 150;
        $DgvTeams.Columns["officeLocation"].Width = 150;
        $DgvTeams.Columns["userPrincipalName"].Width = 350;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.displayName,$row.jobTitle,$row.mail,$row.mobilePhone,$row.officeLocation,$row.userPrincipalName)
        }

        $DgvTeams.Visible = $true
    
    }
    #4) Get Channel Tabs
    elseif($RbtnListTabs.Checked){
        #Variables
        $teamsid = $TxtTeamsId.Text;
        $channelid = $TxtChannelId.Text;

        $apiUrl = 'https://graph.microsoft.com/v1.0/teams/' + $teamsid + '/channels/' + $channelid + '/tabs?$expand=teamsApp'
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $tabs = ($Data | select-object Value).Value | Select-Object id, displayName, configuration

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $tabs
        
        $DgvTeams.Columns[0].Name = "id";
        $DgvTeams.Columns[1].Name = "displayName";
        $DgvTeams.Columns[2].Name = "configuration";
        $DgvTeams.Columns["id"].Width = 300;
        $DgvTeams.Columns["displayName"].Width = 200;
        $DgvTeams.Columns["configuration"].Width = 700;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.id,$row.displayName,$row.configuration)
        }

        $DgvTeams.Visible = $true

    
    }#elseif($RbtnListApps.Checked){
    
    


    #}elseif($RbtnListItems.Checked){
    
    
    }



}
#--------------------------------------------------

#----------------------EVENTS----------------------

$BtnGetData.Add_Click({TeamsQueries})


#--------------------------------------------------

[void]$MicrosoftGraphAPIQueries.ShowDialog()
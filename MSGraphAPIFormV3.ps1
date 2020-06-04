<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    MSGraphAPIForms
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MicrosoftGraphAPIQueries        = New-Object system.Windows.Forms.Form
$MicrosoftGraphAPIQueries.ClientSize  = '1333,946'
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
$TxtClientId.text                = "a919e14e-d9ed-4c1d-babd-b1a19cb50363"
$TxtClientId.width               = 287
$TxtClientId.height              = 20
$TxtClientId.location            = New-Object System.Drawing.Point(148,32)
$TxtClientId.Font                = 'Microsoft Sans Serif,10'

$TxtClientSecret                 = New-Object system.Windows.Forms.MaskedTextBox
$TxtClientSecret.multiline       = $false
$TxtClientSecret.text            = "BDi0722d4-3Lhx25QJlk7jdpYxf~Q-Y3~j"
$TxtClientSecret.width           = 285
$TxtClientSecret.height          = 20
$TxtClientSecret.location        = New-Object System.Drawing.Point(148,93)
$TxtClientSecret.Font            = 'Microsoft Sans Serif,10'

$TxtTenantName                   = New-Object system.Windows.Forms.TextBox
$TxtTenantName.multiline         = $false
$TxtTenantName.text              = "M365x797062.onmicrosoft.com"
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
$GbMSTeamsQueries.height         = 111
$GbMSTeamsQueries.width          = 1268
$GbMSTeamsQueries.text           = "Microsoft Teams Queries"
$GbMSTeamsQueries.location       = New-Object System.Drawing.Point(34,226)

$BtnGetData                      = New-Object system.Windows.Forms.Button
$BtnGetData.text                 = "Get Data"
$BtnGetData.width                = 171
$BtnGetData.height               = 30
$BtnGetData.location             = New-Object System.Drawing.Point(1037,21)
$BtnGetData.Font                 = 'Microsoft Sans Serif,10'

$RbtnListTeams                   = New-Object system.Windows.Forms.RadioButton
$RbtnListTeams.text              = "List All MS Teams"
$RbtnListTeams.AutoSize          = $true
$RbtnListTeams.width             = 104
$RbtnListTeams.height            = 20
$RbtnListTeams.location          = New-Object System.Drawing.Point(27,23)
$RbtnListTeams.Font              = 'Microsoft Sans Serif,10'

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
$GbTeamsSettings.width           = 816
$GbTeamsSettings.text            = "MS Teams Settings"
$GbTeamsSettings.location        = New-Object System.Drawing.Point(488,40)

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
$TxtTeamsId.text                 = "5c0c45da-b203-4f3e-86ad-2235a0e7ac46"
$TxtTeamsId.width                = 352
$TxtTeamsId.height               = 20
$TxtTeamsId.location             = New-Object System.Drawing.Point(117,23)
$TxtTeamsId.Font                 = 'Microsoft Sans Serif,10'

$TxtChannelId                    = New-Object system.Windows.Forms.TextBox
$TxtChannelId.multiline          = $false
$TxtChannelId.text               = "19:859c3225abae4a03bae90dc804a9bdb9@thread.tacv2"
$TxtChannelId.width              = 352
$TxtChannelId.height             = 20
$TxtChannelId.location           = New-Object System.Drawing.Point(117,55)
$TxtChannelId.Font               = 'Microsoft Sans Serif,10'

$RbtnListTabs                    = New-Object system.Windows.Forms.RadioButton
$RbtnListTabs.text               = "List Tabs in a Channel"
$RbtnListTabs.AutoSize           = $true
$RbtnListTabs.width              = 104
$RbtnListTabs.height             = 20
$RbtnListTabs.location           = New-Object System.Drawing.Point(482,18)
$RbtnListTabs.Font               = 'Microsoft Sans Serif,10'

$BtnExportToExcel                = New-Object system.Windows.Forms.Button
$BtnExportToExcel.text           = "Export to Excel"
$BtnExportToExcel.width          = 172
$BtnExportToExcel.height         = 30
$BtnExportToExcel.location       = New-Object System.Drawing.Point(1036,64)
$BtnExportToExcel.Font           = 'Microsoft Sans Serif,10'
$BtnExportToExcel.Visible      = $false

$GbMSTeamsUsage                  = New-Object system.Windows.Forms.Groupbox
$GbMSTeamsUsage.height           = 100
$GbMSTeamsUsage.width            = 1269
$GbMSTeamsUsage.text             = "Microsoft Teams Usage Reports"
$GbMSTeamsUsage.location         = New-Object System.Drawing.Point(34,352)

$GbSPOUsage                      = New-Object system.Windows.Forms.Groupbox
$GbSPOUsage.height               = 100
$GbSPOUsage.width                = 1271
$GbSPOUsage.text                 = "SharePoint Online Usage Reports"
$GbSPOUsage.location             = New-Object System.Drawing.Point(32,470)

$RbtnTeamsUsage01                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage01.text           = "MS Teams user activity  by user"
$RbtnTeamsUsage01.AutoSize       = $true
$RbtnTeamsUsage01.width          = 104
$RbtnTeamsUsage01.height         = 20
$RbtnTeamsUsage01.location       = New-Object System.Drawing.Point(31,17)
$RbtnTeamsUsage01.Font           = 'Microsoft Sans Serif,10'

$RbtnTeamsUsage03                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage03.text           = "Number of users by activity type"
$RbtnTeamsUsage03.AutoSize       = $true
$RbtnTeamsUsage03.width          = 104
$RbtnTeamsUsage03.height         = 20
$RbtnTeamsUsage03.location       = New-Object System.Drawing.Point(30,68)
$RbtnTeamsUsage03.Font           = 'Microsoft Sans Serif,10'

$RbtnTeamsUsage02                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage02.text           = "MS Teams activities by activity type"
$RbtnTeamsUsage02.AutoSize       = $true
$RbtnTeamsUsage02.width          = 104
$RbtnTeamsUsage02.height         = 20
$RbtnTeamsUsage02.location       = New-Object System.Drawing.Point(30,43)
$RbtnTeamsUsage02.Font           = 'Microsoft Sans Serif,10'

$RbtnTeamsUsage04                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage04.text           = "MS Teams device usage by user"
$RbtnTeamsUsage04.AutoSize       = $true
$RbtnTeamsUsage04.width          = 104
$RbtnTeamsUsage04.height         = 20
$RbtnTeamsUsage04.location       = New-Object System.Drawing.Point(482,14)
$RbtnTeamsUsage04.Font           = 'Microsoft Sans Serif,10'

$RbtnTeamsUsage05                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage05.text           = "Number of daily unique users by device type"
$RbtnTeamsUsage05.AutoSize       = $true
$RbtnTeamsUsage05.width          = 104
$RbtnTeamsUsage05.height         = 20
$RbtnTeamsUsage05.location       = New-Object System.Drawing.Point(482,39)
$RbtnTeamsUsage05.Font           = 'Microsoft Sans Serif,10'

$RbtnTeamsUsage06                = New-Object system.Windows.Forms.RadioButton
$RbtnTeamsUsage06.text           = "Unique users by device type"
$RbtnTeamsUsage06.AutoSize       = $true
$RbtnTeamsUsage06.width          = 104
$RbtnTeamsUsage06.height         = 20
$RbtnTeamsUsage06.location       = New-Object System.Drawing.Point(482,67)
$RbtnTeamsUsage06.Font           = 'Microsoft Sans Serif,10'

$RbtnSPOUsage01                  = New-Object system.Windows.Forms.RadioButton
$RbtnSPOUsage01.text             = "SharePoint activity by user"
$RbtnSPOUsage01.AutoSize         = $true
$RbtnSPOUsage01.width            = 104
$RbtnSPOUsage01.height           = 20
$RbtnSPOUsage01.location         = New-Object System.Drawing.Point(35,25)
$RbtnSPOUsage01.Font             = 'Microsoft Sans Serif,10'

$RbtnSPOUsage02                  = New-Object system.Windows.Forms.RadioButton
$RbtnSPOUsage02.text             = "Trend in the number of active users"
$RbtnSPOUsage02.AutoSize         = $true
$RbtnSPOUsage02.width            = 104
$RbtnSPOUsage02.height           = 20
$RbtnSPOUsage02.location         = New-Object System.Drawing.Point(34,50)
$RbtnSPOUsage02.Font             = 'Microsoft Sans Serif,10'

$RbtnSPOUsage03                  = New-Object system.Windows.Forms.RadioButton
$RbtnSPOUsage03.text             = "Number of unique users who interacted with files in SPO"
$RbtnSPOUsage03.AutoSize         = $true
$RbtnSPOUsage03.width            = 104
$RbtnSPOUsage03.height           = 20
$RbtnSPOUsage03.location         = New-Object System.Drawing.Point(485,16)
$RbtnSPOUsage03.Font             = 'Microsoft Sans Serif,10'

$RbtnSPOUsage04                  = New-Object system.Windows.Forms.RadioButton
$RbtnSPOUsage04.text             = "Number of unique pages visited by users"
$RbtnSPOUsage04.AutoSize         = $true
$RbtnSPOUsage04.width            = 104
$RbtnSPOUsage04.height           = 20
$RbtnSPOUsage04.location         = New-Object System.Drawing.Point(484,44)
$RbtnSPOUsage04.Font             = 'Microsoft Sans Serif,10'

$BtnGetDataTeamsUsage            = New-Object system.Windows.Forms.Button
$BtnGetDataTeamsUsage.text       = "Get Data"
$BtnGetDataTeamsUsage.width      = 170
$BtnGetDataTeamsUsage.height     = 30
$BtnGetDataTeamsUsage.location   = New-Object System.Drawing.Point(1041,16)
$BtnGetDataTeamsUsage.Font       = 'Microsoft Sans Serif,10'

$BtnExportTeamsUsage             = New-Object system.Windows.Forms.Button
$BtnExportTeamsUsage.text        = "Export to Excel"
$BtnExportTeamsUsage.width       = 171
$BtnExportTeamsUsage.height      = 30
$BtnExportTeamsUsage.location    = New-Object System.Drawing.Point(1042,58)
$BtnExportTeamsUsage.Font        = 'Microsoft Sans Serif,10'
$BtnExportTeamsUsage.Visible      = $false

$BtnGetDataSPOUsage              = New-Object system.Windows.Forms.Button
$BtnGetDataSPOUsage.text         = "Get Data"
$BtnGetDataSPOUsage.width        = 167
$BtnGetDataSPOUsage.height       = 30
$BtnGetDataSPOUsage.location     = New-Object System.Drawing.Point(1048,16)
$BtnGetDataSPOUsage.Font         = 'Microsoft Sans Serif,10'

$BtnExportTSPOUsage              = New-Object system.Windows.Forms.Button
$BtnExportTSPOUsage.text         = "Export to Excel"
$BtnExportTSPOUsage.width        = 167
$BtnExportTSPOUsage.height       = 30
$BtnExportTSPOUsage.location     = New-Object System.Drawing.Point(1049,60)
$BtnExportTSPOUsage.Font         = 'Microsoft Sans Serif,10'
$BtnExportTSPOUsage.Visible      = $false 

$DgvTeams                    = New-Object system.Windows.Forms.DataGridView
$DgvTeams.width              = 1272
$DgvTeams.height             = 340
$DgvTeams.location           = New-Object System.Drawing.Point(36,583)
$DgvTeams.ColumnCount        = 9
$DgvTeams.Visible            = $false

$MicrosoftGraphAPIQueries.controls.AddRange(@($GbTenantSettings,$GbMSTeamsQueries,$GbTeamsSettings,$GbMSTeamsUsage,$GbSPOUsage,$DgvTeams))
$GbTenantSettings.controls.AddRange(@($LblClientId,$LblTenantName,$LblClientSecret,$TxtClientId,$TxtClientSecret,$TxtTenantName,$LblResource,$TxtResource))
$GbMSTeamsQueries.controls.AddRange(@($BtnGetData,$RbtnListTeams,$RbtnListChannels,$RbtnGetMembers,$RbtnListTabs,$BtnExportToExcel))
$GbTeamsSettings.controls.AddRange(@($LblTeamsId,$LblChannelId,$TxtTeamsId,$TxtChannelId))
$GbMSTeamsUsage.controls.AddRange(@($RbtnTeamsUsage01,$RbtnTeamsUsage03,$RbtnTeamsUsage02,$RbtnTeamsUsage04,$RbtnTeamsUsage05,$RbtnTeamsUsage06,$BtnGetDataTeamsUsage,$BtnExportTeamsUsage))
$GbSPOUsage.controls.AddRange(@($RbtnSPOUsage01,$RbtnSPOUsage02,$RbtnSPOUsage03,$RbtnSPOUsage04,$BtnGetDataSPOUsage,$BtnExportTSPOUsage))

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

    
    }
    else{
        [System.Windows.MessageBox]::Show('You need to select an option','Graph API Form','OK','Information')
    }    
    
    }

function TeamsUsageQueries{

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

    #MS Team Usage Queries
    
    if($RbtnTeamsUsage01.Checked){

        #$period = "D7"
        #$apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsUserActivityUserDetail(period='" + $period + "')?$" + "format=application/json"
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsUserActivityUserDetail(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, userPrincipalName, lastActivityDate, isDeleted, assignedProducts, teamChatMessageCount

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "userPrincipalName";
        $DgvTeams.Columns[2].Name = "lastActivityDate";
        $DgvTeams.Columns[3].Name = "isDeleted";
        $DgvTeams.Columns[4].Name = "assignedProducts";
        $DgvTeams.Columns[5].Name = "teamChatMessageCount";
        $DgvTeams.Columns[6].Name = "";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["userPrincipalName"].Width = 300;
        $DgvTeams.Columns["lastActivityDate"].Width = 300;
        $DgvTeams.Columns["isDeleted"].Width = 300;
        $DgvTeams.Columns["assignedProducts"].Width = 300;
        $DgvTeams.Columns["teamChatMessageCount"].Width = 300;
        $DgvTeams.Columns[6].Width = 50;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.userPrincipalName,$row.lastActivityDate,$row.isDeleted,$row.assignedProducts,$row.teamChatMessageCount)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnTeamsUsage02.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsUserActivityCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, reportDate, teamChatMessages, privateChatMessages, calls, meetings

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "reportDate";
        $DgvTeams.Columns[2].Name = "teamChatMessages";
        $DgvTeams.Columns[3].Name = "privateChatMessages";
        $DgvTeams.Columns[4].Name = "calls";
        $DgvTeams.Columns[5].Name = "meetings";
        $DgvTeams.Columns[6].Name = "";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns["teamChatMessages"].Width = 300;
        $DgvTeams.Columns["privateChatMessages"].Width = 300;
        $DgvTeams.Columns["calls"].Width = 300;
        $DgvTeams.Columns["meetings"].Width = 300;
        $DgvTeams.Columns[6].Width = 50;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.reportDate,$row.teamChatMessages,$row.privateChatMessages,$row.calls,$row.meetings)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnTeamsUsage03.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsUserActivityUserCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, reportDate, teamChatMessages, privateChatMessages, calls, meetings, otherActions

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "reportDate";
        $DgvTeams.Columns[2].Name = "teamChatMessages";
        $DgvTeams.Columns[3].Name = "privateChatMessages";
        $DgvTeams.Columns[4].Name = "calls";
        $DgvTeams.Columns[5].Name = "meetings";
        $DgvTeams.Columns[6].Name = "otherActions";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns["teamChatMessages"].Width = 300;
        $DgvTeams.Columns["privateChatMessages"].Width = 300;
        $DgvTeams.Columns["calls"].Width = 300;
        $DgvTeams.Columns["meetings"].Width = 300;
        $DgvTeams.Columns["otherActions"].Width = 300;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.reportDate,$row.teamChatMessages,$row.privateChatMessages,$row.calls,$row.meetings,$row.otherActions)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnTeamsUsage04.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsDeviceUsageUserDetail(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, userPrincipalName, lastActivityDate, usedWeb, usedWindowsPhone, usediOS, usedMac, usedAndroidPhone, usedWindows

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "userPrincipalName";
        $DgvTeams.Columns[2].Name = "lastActivityDate";
        $DgvTeams.Columns[3].Name = "usedWeb";
        $DgvTeams.Columns[4].Name = "usedWindowsPhone";
        $DgvTeams.Columns[5].Name = "usediOS";
        $DgvTeams.Columns[6].Name = "usedMac";
        $DgvTeams.Columns[7].Name = "usedAndroidPhone";
        $DgvTeams.Columns[8].Name = "usedWindows";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["userPrincipalName"].Width = 300;
        $DgvTeams.Columns["lastActivityDate"].Width = 300;
        $DgvTeams.Columns["usedWeb"].Width = 300;
        $DgvTeams.Columns["usedWindowsPhone"].Width = 300;
        $DgvTeams.Columns["usediOS"].Width = 300;
        $DgvTeams.Columns["usedMac"].Width = 300;
        $DgvTeams.Columns["usedAndroidPhone"].Width = 300;
        $DgvTeams.Columns["usedWindows"].Width = 300;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.userPrincipalName,$row.lastActivityDate,$row.usedWeb,$row.usedWindowsPhone,$row.usediOS,$row.usedMac,$row.usedAndroidPhone,$row.usedWindows)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnTeamsUsage05.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsDeviceUsageUserCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, web, windowsPhone, androidPhone, ios, mac, windows, reportDate

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "web";
        $DgvTeams.Columns[2].Name = "windowsPhone";
        $DgvTeams.Columns[3].Name = "androidPhone";
        $DgvTeams.Columns[4].Name = "ios";
        $DgvTeams.Columns[5].Name = "mac";
        $DgvTeams.Columns[6].Name = "windows";
        $DgvTeams.Columns[7].Name = "reportDate";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["web"].Width = 300;
        $DgvTeams.Columns["windowsPhone"].Width = 300;
        $DgvTeams.Columns["androidPhone"].Width = 300;
        $DgvTeams.Columns["ios"].Width = 300;
        $DgvTeams.Columns["mac"].Width = 300;
        $DgvTeams.Columns["windows"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.web,$row.windowsPhone,$row.androidPhone,$row.ios,$row.mac,$row.windows,$row.reportDate)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnTeamsUsage06.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getTeamsDeviceUsageDistributionUserCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, web, windowsPhone, androidPhone, ios, mac, windows

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "web";
        $DgvTeams.Columns[2].Name = "windowsPhone";
        $DgvTeams.Columns[3].Name = "androidPhone";
        $DgvTeams.Columns[4].Name = "ios";
        $DgvTeams.Columns[5].Name = "mac";
        $DgvTeams.Columns[6].Name = "windows";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["web"].Width = 300;
        $DgvTeams.Columns["windowsPhone"].Width = 300;
        $DgvTeams.Columns["androidPhone"].Width = 300;
        $DgvTeams.Columns["ios"].Width = 300;
        $DgvTeams.Columns["mac"].Width = 300;
        $DgvTeams.Columns["windows"].Width = 300;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.web,$row.windowsPhone,$row.androidPhone,$row.ios,$row.mac,$row.windows)
        }

        $DgvTeams.Visible = $true
    }
    else{
        [System.Windows.MessageBox]::Show('You need to select an option','Graph API Form','OK','Information')
    }  
   
}

function SPOUsageQueries{

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

    #SPO Usage Queries
    if($RbtnSPOUsage01.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getSharePointActivityUserDetail(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, userPrincipalName, lastActivityDate, viewedOrEditedFileCount, syncedFileCount, sharedInternallyFileCount, sharedExternallyFileCount, visitedPageCount, assignedProducts

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "userPrincipalName";
        $DgvTeams.Columns[2].Name = "lastActivityDate";
        $DgvTeams.Columns[3].Name = "viewedOrEditedFileCount";
        $DgvTeams.Columns[4].Name = "syncedFileCount";
        $DgvTeams.Columns[5].Name = "sharedInternallyFileCount";
        $DgvTeams.Columns[6].Name = "sharedExternallyFileCount";
        $DgvTeams.Columns[7].Name = "visitedPageCount";
        $DgvTeams.Columns[8].Name = "assignedProducts";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["userPrincipalName"].Width = 300;
        $DgvTeams.Columns["lastActivityDate"].Width = 300;
        $DgvTeams.Columns["viewedOrEditedFileCount"].Width = 300;
        $DgvTeams.Columns["syncedFileCount"].Width = 300;
        $DgvTeams.Columns["sharedInternallyFileCount"].Width = 300;
        $DgvTeams.Columns["sharedExternallyFileCount"].Width = 300;
        $DgvTeams.Columns["visitedPageCount"].Width = 300;
        $DgvTeams.Columns["assignedProducts"].Width = 300;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.userPrincipalName,$row.lastActivityDate,$row.viewedOrEditedFileCount,$row.syncedFileCount,$row.sharedInternallyFileCount,$row.sharedExternallyFileCount,$row.visitedPageCount,$row.assignedProducts)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnSPOUsage02.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getSharePointActivityFileCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, viewedOrEdited, synced, sharedInternally, sharedExternally, reportDate

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "viewedOrEdited";
        $DgvTeams.Columns[2].Name = "synced";
        $DgvTeams.Columns[3].Name = "sharedInternally";
        $DgvTeams.Columns[4].Name = "sharedExternally";
        $DgvTeams.Columns[5].Name = "reportDate";
        $DgvTeams.Columns[6].Name = "";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["viewedOrEdited"].Width = 300;
        $DgvTeams.Columns["synced"].Width = 300;
        $DgvTeams.Columns["sharedInternally"].Width = 300;
        $DgvTeams.Columns["sharedExternally"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns[6].Width = 50;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.viewedOrEdited,$row.synced,$row.sharedInternally,$row.sharedExternally,$row.reportDate)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnSPOUsage03.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getSharePointActivityUserCounts(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, visitedPage, viewedOrEdited, synced, sharedInternally, sharedExternally, reportDate

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "visitedPage";
        $DgvTeams.Columns[2].Name = "viewedOrEdited";
        $DgvTeams.Columns[3].Name = "synced";
        $DgvTeams.Columns[4].Name = "sharedInternally";
        $DgvTeams.Columns[5].Name = "sharedExternally";
        $DgvTeams.Columns[6].Name = "reportDate";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["visitedPage"].Width = 300;
        $DgvTeams.Columns["viewedOrEdited"].Width = 300;
        $DgvTeams.Columns["synced"].Width = 300;
        $DgvTeams.Columns["sharedInternally"].Width = 300;
        $DgvTeams.Columns["sharedExternally"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.visitedPage,$row.viewedOrEdited,$row.synced,$row.sharedInternally,$row.sharedExternally,$row.reportDate)
        }

        $DgvTeams.Visible = $true
    }
    elseif($RbtnSPOUsage04.Checked){
        
        $apiUrl = "https://graph.microsoft.com/beta/reports/getSharePointActivityPages(period='D7')?$" + "format=application/json"
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)"} -Uri $apiUrl -Method Get
        $Teams = ($Data | select-object Value).Value | Select-Object reportRefreshDate, visitedPageCount, reportDate

        #DataGridView Databinding 
        $DgvTeams.Rows.Clear()
        $DgvTeamsData = $Teams
        
        $DgvTeams.Columns[0].Name = "reportRefreshDate";
        $DgvTeams.Columns[1].Name = "visitedPageCount";
        $DgvTeams.Columns[2].Name = "reportDate";
        $DgvTeams.Columns[3].Name = "";
        $DgvTeams.Columns[4].Name = "";
        $DgvTeams.Columns[5].Name = "";
        $DgvTeams.Columns[6].Name = "";
        $DgvTeams.Columns[7].Name = "";
        $DgvTeams.Columns[8].Name = "";
        $DgvTeams.Columns["reportRefreshDate"].Width = 300;
        $DgvTeams.Columns["visitedPageCount"].Width = 300;
        $DgvTeams.Columns["reportDate"].Width = 300;
        $DgvTeams.Columns[3].Width = 50;
        $DgvTeams.Columns[4].Width = 50;
        $DgvTeams.Columns[5].Width = 50;
        $DgvTeams.Columns[6].Width = 50;
        $DgvTeams.Columns[7].Width = 50;
        $DgvTeams.Columns[8].Width = 50;
        $DgvTeams.ColumnHeadersVisible = $true
        foreach ($row in $DgvTeamsData){
            $DgvTeams.Rows.Add($row.reportRefreshDate,$row.visitedPageCount,$row.reportDate)
        }

        $DgvTeams.Visible = $true
    }
    else{
        [System.Windows.MessageBox]::Show('You need to select an option','Graph API Form','OK','Information')
    }  
}

#--------------------------------------------------

#----------------------EVENTS----------------------

$BtnGetData.Add_Click({TeamsQueries})
$BtnGetDataTeamsUsage.Add_Click({TeamsUsageQueries})
$BtnGetDataSPOUsage.Add_Click({SPOUsageQueries})

#--------------------------------------------------


[void]$MicrosoftGraphAPIQueries.ShowDialog()
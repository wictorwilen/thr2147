#-----------------------------------------------------------------------#
# Microsoft Ignite 2018 - THR2147
# Wictor Wilén
# wictor@wictorwilen.se
# @wictor
# Scripts available at https://askwictor.com/teams-scripts
#-----------------------------------------------------------------------#

# Install the Microsoft Teams PowerShell
# Current available version 0.9.2
Install-Module -Name MicrosoftTeams

## Or
Install-Module -Name MicrosoftTeams -Scope CurrentUser

# List cmdlets in the module

Get-TeamHelp

# Connecto to Microsoft Teams
## Can connect using credentials, login popup, or AAD app access token or certificate
## Returns info about your tenant (id, etc)
Connect-MicrosoftTeams

# Disconnect
Disconnect-MicrosoftTeams

# List teams
# Note: 0.9.2 only shows the teams the connected user is a member of
Get-Team 

(Get-Team).Count

Get-Team -User w.wilen@wictordev.onmicrosoft.com

(Get-Team -User w.wilen@wictordev.onmicrosoft.com).Count

Get-Team -User j.doe@wictordev.onmicrosoft.com

# Get a specific team

$team = Get-Team | Where-Object {$_.DisplayName -eq "The new team"} 
$GroupId = $team.GroupId
$GroupId

# Get info about a specific team
Get-TeamChannel -GroupId $GroupId
Get-TeamFunSettings -GroupId $GroupId
Get-TeamGuestSettings -GroupId $GroupId
Get-TeamMemberSettings -GroupId $GroupId
Get-TeamMessagingSettings -GroupId $GroupId

# List all members
Get-TeamUser -GroupId $GroupId

# List a specific role
Get-TeamUser -GroupId $GroupId -Role "Owner"


# Create a team
$team = New-Team -DisplayName "Ignite demo team" -Description "This is a brand new team" -Alias "IgniteDemoTeam" -AccessType Public -AddCreatorAsMember $true
$team
$GroupId = $team.GroupId
$GroupId

Get-TeamUser -GroupId $GroupId

# Set the team picture
Set-TeamPicture -GroupId $GroupId -ImagePath .\ignite-logo.png

# Update a setting
Set-TeamMessagingSettings -GroupId $GroupId -AllowUserDeleteMessages $false
Get-TeamMessagingSettings -GroupId $GroupId

# Create a chennel
New-TeamChannel -GroupId $GroupId -DisplayName "New Channel" -Description "This is a new channel"
Get-TeamChannel -GroupId $GroupId

# Upadte a channel
Set-TeamChannel -GroupId $GroupId -CurrentDisplayName "New Channel" -NewDisplayName "A better name" -Description "A better description"

# Add a user
Add-TeamUser -GroupId $GroupId -User "j.doe@wictordev.onmicrosoft.com" -Role Member
Get-TeamUser -GroupId $GroupId

# Disconnect
Disconnect-MicrosoftTeams

# With certificate (not all functionality is currently supported)
# Import-PfxCertificate -FilePath .\cert.pfx -CertStoreLocation Cert:\LocalMachine\My 

# Requires an Azure AD application, with a certificate and the following permissions TODO
$applicationId = '23fb742e-8565-49bd-9cb0-1a89919c49db'
$thumbprint = 'd802ebab09e3a9788e9d1a234ea4ac98d173c6c3'
$tenantId = 'dbb0bda5-c207-4515-bd5f-fb7bf7d85616'
Connect-MicrosoftTeams -TenantId $tenantId -CertificateThumbprint $thumbprint -ApplicationId $applicationId

Get-Team 

(Get-Team).Count

Get-Team -User w.wilen@wictordev.onmicrosoft.com

(Get-Team -User w.wilen@wictordev.onmicrosoft.com).Count

Get-Team -User j.doe@wictordev.onmicrosoft.com


# With cert (fails)
$team = New-Team -DisplayName "Ignite Cert demo" -Description "This is a brand new team" -Alias "IgniteCertDemo" -AccessType Public 

# WOrkaround for certs
npm install @pnp/office365-cli@next --global #Cert auth will come in next version (1.7+)
$applicationId = '23fb742e-8565-49bd-9cb0-1a89919c49db'
$thumbprint = 'd802ebab09e3a9788e9d1a234ea4ac98d173c6c3'
$env:OFFICE365CLI_TENANT = 'wictordev.onmicrosoft.com'
o365 graph connect --authType certificate --certificateFile cert.key --applicationId $applicationId --thumbprint $thumbprint #--debug
$groups = o365 graph o365group list --output json | ConvertFrom-Json

$group = $groups | Where-Object {$_.displayName -eq "Group987"} 
$groupId = $group.id
$GroupId

$team = New-Team -Group $GroupId
$team

Get-Team

# will fail
New-TeamChannel -GroupId $groupId -DisplayName "Certificate" -Description "Created by cert auth"


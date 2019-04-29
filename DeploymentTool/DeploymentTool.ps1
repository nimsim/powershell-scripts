#BEFORE RUNNING SCRIPT, MAKE SURE YOU ARE RUNNING POWERSHELL IN ADMIN mode 
$ErrorPreference='Stop'

#Setting the ExecutionPolicy to BYPASS. Note: This won't work because running scripts is by default restricted in PowerShell. 
#See OneNote for more information: https://microsoft.sharepoint.com/teams/OfficePeople/_layouts/OneNote.aspx?id=%2Fteams%2FOfficePeople%2FSiteAssets%2FOffice%20People%20Notebook&wd=target%28Projects%2FTest%20Lab.one%7CA70B945C-4284-48F6-983A-30A0C55E78C1%2FTest%20Users%20deploy%20O365%7CDCD4BAE6-936C-4944-A922-E9BED7B7B584%2F%29
$execPolicy = Get-ExecutionPolicy
if ($execPolicy -eq "AllSigned" -or $execPolicy -eq "Default" -or $execPolicy -eq "RemoteSigned" -or $execPolicy -eq "Restricted") {
    Write-Host "Execution Policy is not set to bypass, setting it temporarily"
    Try {
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Force
    }
    Catch {
    Write-Host -ForegroundColor Red "Could not set Execution Policy to Bypass, please run the script in Administrator mode"
    Write-Host -ForegroundColor Red "Exiting script.."
    Break
    }
}

# Path for transcript
$path = [environment]::getfolderpath(“mydocuments”)
$path = "$path\DeploymentTool"
$createdDate = get-date -UFormat "%y%m%d%H%M"
If(!(test-path $path))
{
      New-Item -ItemType Directory -Force -Path $path
}
Start-Transcript -Path "$path\DeploymentTool_$createdDate.txt" -NoClobber -Append
Write-Host -ForegroundColor Yellow "Starting transcript and saving the logfile at $path"

# Install necessary module for GUI
Write-Host -ForegroundColor Green "Checking if Anybox is installed"
if (Get-InstalledModule -Name Anybox -ErrorAction SilentlyContinue){
    Write-Host -ForegroundColor Green "Anybox Powershell Module exists"
    $online = Find-Module -Name Anybox
    $local = Get-InstalledModule -Name Anybox
    if ($online.Version -gt $local.Version)
    {
    Write-Host -ForegroundColor Green "Updating Anybox Powershell module to latest version"
    Update-Module Anybox -Force -Confirm:$false
    }
} else {
    Write-Host -ForegroundColor Red "Anybox Powershell does not exist, installing now"
    Install-module -name Anybox -Force -AllowClobber
    Write-Host -ForegroundColor Green "Anybox Powershell Module is now installed"
}

#Import Anyboxy Module
Import-Module -name Anybox

$anybox = New-Object AnyBox.AnyBox
$anybox.ContentAlignment = 'Center'
# E-MailAddress Input for tagging resource with correct user
$anybox.Prompts += New-AnyBoxPrompt -Name "EmailAdr" -InputType Text -Message "Please write your E-mailAddress" -ValidateNotEmpty -Group "Usage info"
# Purpose of usage
$anybox.Prompts += New-AnyBoxPrompt -Name "Purpose" -InputType Text -Message "Please tell us why you are using the Deployment Tool" -ValidateNotEmpty -Group "Usage info"
$anybox.Buttons += New-AnyBoxButton -Name 'office' -Text 'Office 365 Deployment'
$anybox.Buttons += New-AnyBoxButton -Name 'OnPrem' -Text 'OnPrem Deployment'
$anybox.Buttons += New-AnyBoxButton -Name 'Cancel' -IsCancel -Text 'Cancel'
$anybox.Title = 'OP Testlabs Deployment Tool'
$anybox.Icon = 'Question'
$anybox.Comment += "Transcript can be found at $path" 
$anybox.Prompts += New-AnyBoxPrompt -Name "urlOneNote" -InputType Link -Message "OneNote Deployment guide" -DefaultValue 'onenote:https://microsoft.sharepoint.com/teams/OfficePeople/SiteAssets/Office%20People%20Notebook/Projects/Test%20Lab.one#Guides%20Hybrid&section-id={A70B945C-4284-48F6-983A-30A0C55E78C1}&page-id={F2CE8291-CE63-4B6E-9AE0-4431C555929F}&end'

$choiceOnPremCloud = $anybox | Show-AnyBox


# Teams Webhook Usage Channel 
$TeamsChannelUri = "https://outlook.office.com/webhook/c51fbb48-de5b-4fa9-a470-71f546aef43f@72f988bf-86f1-41af-91ab-2d7cd011db47/IncomingWebhook/f42496660a0c4a89ac7f806a0bfdae45/a7930336-98e4-4b54-a4a0-cf6e5a1286b9"
 
$BodyTemplate = @"
    {
        "@type": "MessageCard",
        "@context": "https://schema.org/extensions",
        "summary": "DeploymentToolUsed-Notification",
        "themeColor": "D778D7",
        "title": "Deployment Tool: Resources created",
         "sections": [
            {
            
                "facts": [
                    {
                        "name": "Username:",
                        "value": "DOMAIN_USERNAME"
                    },
                    {
                        "name": "Purpose:",
                        "value": "PURPOSE"
                    },
                    {
                        "name": "Time:",
                        "value": "DATETIME"
                    },
                    {
                        "name": "PC Name:",
                        "value": "COMPUTERNAME"
                    }
                ],
                "text": "A user has used the Deployment Tool!"
            }
        ]
    }
"@
 
 
        $body = $BodyTemplate.Replace("DOMAIN_USERNAME","$($choiceOnPremCloud['emailadr'])").Replace("PURPOSE","$($choiceOnPremCloud['purpose'])").Replace("DATETIME",$(Get-Date)).Replace("COMPUTERNAME","$env:COMPUTERNAME")
        Invoke-RestMethod -uri $TeamsChannelUri -Method Post -body $body -ContentType 'application/json'

        # abort if cancel is chosen
        if ($choiceOnPremCloud['Cancel']) {
            Show-AnyBox -Message "You have chosen to cancel the Deployment Tool. Script will now end." -Buttons 'OK'
            Break
        }

####################################################
#                                                  #
#      Office365 section below                     #
#      This part of the script works on O365       #
#      deployment. The script installs modules     #
#      for Sharepoint PnP, MsOnline and            #
#      connects to Exchange Online                 # 
#                                                  #
#      This part of the script uses a CSV which    #
#      is pre-configured (editable) and creates    #
#      20 users fully licensed with all attributes #
#      added in Sharepoint, Exchange, and AAD      #
#                                                  #
#      Contact Margrete Sævareid                   #
#      margrets@microsoft.com if you have any      #
#      questions.                                  #
#                                                  #
####################################################

if ($choiceOnPremCloud['office']) {
    $anybox = New-Object AnyBox.AnyBox
    $anybox.Message += 'If you have created a Trial- or Demo-tenant, you can proceed below. If not, create one and proceed below.'
    $anybox.Message += ''
    $anybox.Message += 'This tool creates 20 licensed users in a tenant based on a userlist.', 'You can edit the userlist as you like, but know that the default userlist has approved'
    $anybox.Message += 'names, numbers and addresses. Important for screenshots shared externally'
      
    $anybox.Prompts = @(
      # Tenant admin username
      New-AnyBoxPrompt -Name "userName" -InputType Text -Message "Tenant Admin Username" -ValidateNotEmpty -DefaultValue 'admin@xyz.onmicrosoft.com' -Group 'Connection Info' -Alignment Left
      # Tenant admin password
      New-AnyBoxPrompt -Name "password" -InputType Password -Message "Tenant Admin Password" -ValidateNotEmpty -Group 'Connection Info' -Alignment Left
      # Implement Address Book Policy?
      New-AnyBoxPrompt -Name "adrBookPolicy" -InputType Text -Message "Would you like to implement an Address Book Policy?" -ValidateSet 'Yes','No' -DefaultValue 'Yes' -Group 'Customization' -Alignment Left
      # Wipe users already in tenant?
      New-AnyBoxPrompt -Name "userWipe" -InputType Text -Message "Would you like to wipe all previous tenant users (except Admin)?" -ValidateSet 'Yes','No' -DefaultValue 'Yes' -Group 'Customization' -Alignment Left
      # Location of userlist
      New-AnyBoxPrompt -Name "userList" -InputType FileOpen -Message "Location of userlist CSV-file" -ValidateNotEmpty -Group 'Customization' -Alignment Left
      # delimiter for userlist
      New-AnyBoxPrompt -Name "csvdelimiter" -InputType Text -Message "Type of delimiter for the userlist" -ValidateSet ',',';','-','|' -DefaultValue ',' -Group 'Customization' -Alignment Left
      # link to OneNote
      New-AnyBoxPrompt -Name "urlOneNote" -InputType Link -Message "OneNote Deployment guide" -DefaultValue 'onenote:https://microsoft.sharepoint.com/teams/OfficePeople/SiteAssets/Office%20People%20Notebook/Projects/Test%20Lab.one#Guides%20Hybrid&section-id={A70B945C-4284-48F6-983A-30A0C55E78C1}&page-id={F2CE8291-CE63-4B6E-9AE0-4431C555929F}&end'
    )

    $anybox.ContentAlignment = 'Center'
    $anybox.Buttons += New-AnyBoxButton -Name 'Submit' -IsDefault -Text 'Submit'
    $anybox.Buttons += New-AnyBoxButton -Name 'Cancel' -IsCancel -Text 'Cancel'
    $anybox.Title = 'OP Testlabs'
    $anybox.Icon = 'Question'
    $anybox.Comment += "Transcript can be found at $path" 

    $choiceDeployCloud = $anybox | Show-AnyBox

        # abort if cancel is chosen
        if ($choiceDeployCloud['submit'] -eq $false) {
            Show-AnyBox -Message "You have chosen to cancel the deployment. Script will now end." -Buttons 'OK'
            Break
        }

        if ($choiceDeployCloud['submit']) {
            # Create Credentials for Office 365 based on input from user
            $O365Cred = New-Object System.Management.Automation.PSCredential ($choiceDeployCloud['userName'], $choiceDeployCloud['password'])
            
            #Tenant is unique for every new testenvironment. Grab this by splitting the username of the user logging on O365 (admin). 
            #FQDN for full range "contoso.onmicrosoft.com"
            #Name for only "Contoso
            $tenantFQDN = ($O365Cred.username -split("@"))[1]
            $tenantName = ($O365Cred.username -split("@")).split(".")[1]

            Show-AnyBox -Message 'The tool will now try to connect to your tenant with ',"the username $($choiceDeployCloud['userName'])",'','First we need to install the required modules for:','Microsoft Online (O365) and Sharepoint Online' -Buttons 'OK'
            
            Write-Host -ForegroundColor Green "Checking if Sharepoint PnP module is installed"
            if ($SharepointPnP = Get-InstalledModule -Name SharepointPnPPowershellOnline){
                Write-Host -ForegroundColor Green "Sharepoint PnP Module exists"
                $online = Find-Module -Name SharepointPnPPowershellOnline
                $local = Get-InstalledModule -Name SharepointPnPPowershellOnline
                
                if ($online.Version -gt $local.Version)
                    {
                    Write-Host -ForegroundColor Green "Updating Sharepoint PnP module to latest version"
                    Update-Module SharepointPnPPowershell* -Force -Confirm:$false
                    }
                } else {
                    Write-Host -ForegroundColor Red "Sharepoint PnP does not exist, installing now"
                    Install-module -name SharepointPnPPowershellOnline -Force
                    Write-Host -ForegroundColor Green "Sharepoint PnP Module is now installed"
                }

            Write-Host -ForegroundColor Green "Checking if  Microsoft Online module is installed"
            if ($msonline = Get-InstalledModule -Name MSOnline) {
                Write-Host -ForegroundColor Green "MSOnline Module exists"
                $online = Find-Module -Name MSOnline
                $local = Get-InstalledModule -Name MSOnline
                
                if ($online.Version -gt $local.Version)
                    {
                    Write-Host -ForegroundColor Green "Updating Microsoft Online module to latest version"
                    Update-Module MSOnline* -Force -Confirm:$false
                    }
                } else {
                    Write-Host -ForegroundColor Red "MSOnline Module does not exist, installing now"
                    Install-module -name MSOnline -Force
                    Write-Host -ForegroundColor Green "MSOnline Module is now installed"
                }
            
            # import the required modules
            $session = Get-PSSession
            Import-Module MSOnline,SharepointPnPPowershellOnline
            
            If (($session.ComputerName -eq "ps.outlook.com") -and ($session.State -eq "Opened")) {
            Remove-PSSession -Session $o365Session
            Write-Host "Refreshing connection to MS Online, Exchange Online, and Sharepoint PnP Online for new roles"
        
                Try
                #Sharepoint PnP Connection
                {
                    Connect-PnPOnline –Url https://$tenantName.sharepoint.com –Credentials $O365Cred
                }
                Catch
                {
                    Write-Host -ForegroundColor Red "Could not connect to Sharepoint Online - Are you sure it is provisioned? `nIf this is an SDFv2-tenant, it might take up to 300 minutes to provision Sharepoint"
                    Break
                }

                Try
                #Exchange Online Connection
                {
                    $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $O365Cred -Authentication Basic -AllowRedirection
                    Import-PSSession $O365Session -AllowClobber
                }
                Catch
                {
                    Write-Host -ForegroundColor Red "Could not connect to Exchange Online - Are you sure it is provisioned? `nIf this is an SDFv2-tenant, it might take up to 300 minutes to provision Exchange"
                    Write-Host -ForegroundColor DarkMagenta "We failed to read $_.Exception.Itemname. The errer message was $_.Exception.Message"
                    Break
                }
                Try
                #Office 365 Connection
                {
                    Connect-MsolService –Credential $O365Cred
                }
                    Catch
                {
                    Write-Host -ForegroundColor Red "Could not connect to Microsoft Online - Are you sure it is provisioned?"
                    Break
                }
                } else {
                    
                    Write-Host -ForegroundColor Cyan "Connecting to MS Online, Exchange Online, and Sharepoint PnP Online"
    
                    Try
                    #Sharepoint PnP Connection
                    {
                        Connect-PnPOnline –Url https://$tenantName.sharepoint.com –Credentials $O365Cred
        
                    }
                    Catch
                    {
                    Write-Host -ForegroundColor Red "Could not connect to Sharepoint Online - Are you sure it is provisioned? `nIf this is an SDFv2-tenant, it might take up to 300 minutes to provision Sharepoint"
                    Break
                    }

                    Try
                    #Exchange Online Connection
                    {
                        $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $O365Cred -Authentication Basic -AllowRedirection 
                        Import-PSSession $O365Session -AllowClobber
                    }
                    Catch
                    {
                        Write-Host -ForegroundColor Red "Could not connect to Exchange Online - Are you sure it is provisioned? `nIf this is an SDFv2-tenant, it might take up to 300 minutes to provision Exchange"
                        Write-Host -ForegroundColor DarkMagenta "We failed to read $_.Exception.Itemname. The error message was $_.Exception.Message"
                        Break
                    }
        
                    Try
                    {
                    #Office 365 Connection
                        Connect-MsolService –Credential $O365Cred
                    }  
                    Catch
                    {
                        Write-Host -ForegroundColor Red "Could not connect to Exchange Online - Are you sure it is provisioned? `nIf this is an SDFv2-tenant, it might take up to 300 minutes to provision Exchange"
                        Write-Host -ForegroundColor DarkMagenta "We failed to read $_.Exception.Itemname. The error message was $_.Exception.Message"
                        Break
                    }
                }
        
            #Valid subscriptions available
            $subsAvailable = Get-MsolAccountSku | ?{$_.accountSkuID -like "*ENTERPRISEPACK" -or $_.accountSkuID -like "*ENTERPRISEPREMIUM" -or $_.accountSkuID -like "*SPE_E3" -or $_.accountSkuID -like "SMB_BUSINESS*"}
            $subscriptions = $subsAvailable.AccountSkuId

            $anybox = New-Object AnyBox.AnyBox

            $anybox.Message += 'These are the valid licenses found on your tenant.'
            $anybox.Message += ''
            $anybox.Message += 'Please pick a valid license before continuing'
            $anybox.Prompts = @(
            # Valid Office 365 Subscriptions available on tenant 
              New-AnyBoxPrompt -Name "subscriptions" -InputType Text -Message "Subscriptions available" -ValidateSet @($subscriptions)
              # link to OneNote
              New-AnyBoxPrompt -Name "urlOneNote" -InputType Link -Message "OneNote Deployment guide" -DefaultValue 'onenote:https://microsoft.sharepoint.com/teams/OfficePeople/SiteAssets/Office%20People%20Notebook/Projects/Test%20Lab.one#Guides%20Hybrid&section-id={A70B945C-4284-48F6-983A-30A0C55E78C1}&page-id={F2CE8291-CE63-4B6E-9AE0-4431C555929F}&end'
            )

            $anybox.ContentAlignment = 'Center'
            $anybox.Buttons += New-AnyBoxButton -Name 'Submit' -IsDefault -Text 'Submit'
            $anybox.Buttons += New-AnyBoxButton -Name 'Cancel' -IsCancel -Text 'Cancel'
            $anybox.Title = 'OP Testlabs'
            $anybox.Icon = 'Question'
            $anybox.Comment += "Transcript can be found at $path" 

            $choiceSubscriptions = $anybox | Show-AnyBox

            #Check if Role is added to Organization Management. If not, add it.
            Write-Host -ForegroundColor Cyan "Enabling Exchange Organization Customization, this will take a few minutes"
            Try
            {
                Enable-OrganizationCustomization -ErrorAction SilentlyContinue -ErrorVariable $orgCustomizeError
            }
            Catch
            {
                Write-warning "Could not enable Exchange Organization Customization - Is Exchange provisioned? If this is an SDFv2-tenant, it might take up to 300 minutes to provision Exchange"
            }

            Try
            {
                Write-Host -ForegroundColor Cyan "Enabling Address Lists role for your administrator"
                $roleAssign=New-ManagementRoleAssignment -SecurityGroup "Organization Management" -Role "Address Lists" -ErrorAction SilentlyContinue
            }
            Catch
            {
                Write-Warning  "Could not add role Address Lists to your user - Is Organization Customization enabled? If this is an SDFv2-tenant, it might take up to 300 minutes to provision Exchange"
                Write-Warning  "You can try manually by running Enable-OrganizationCustomization"
            }
            #}

            if ($orgCustomizeError) {
            Write-Warning -ForegroundColor Yellow "Script failed to activate OrganizationCustomization. This is only important for Address Book Policies`nSee FAQ in OneNote for help"
            Start-sleep -Seconds 4
            }

            if ($choiceDeployCloud['userwipe'] -eq "yes") {
                #grab admins for exclusion of removal
                $role = Get-MsolRole -RoleName "Company Administrator"
                $adm = @()
                $admins = Get-MsolRoleMember -RoleObjectId $role.ObjectId
                    foreach ($admin in $admins) {
                        if($Admin.EmailAddress -ne $null){
                            $adm+= $admin.EmailAddress
                        }
                    }
              # proceed removal of users
              write-host -ForegroundColor DarkYellow "You have chosen to remove all users from the tenant"
              write-host -ForegroundColor DarkYellow "Removing all users" 
              Get-MsolUser -All | where {$_.UserPrincipalName -notcontains $adm} | Remove-MsolUser -Force
              Get-MsolUser -ReturnDeletedUsers -All | Remove-MsolUser -RemoveFromRecycleBin -Force -ErrorAction SilentlyContinue
              Get-MsolGroup -All | Remove-MsolGroup -Force
              Get-DistributionGroup | Remove-DistributionGroup -Confirm:$false
              Get-UnifiedGroup | Remove-UnifiedGroup -Confirm:$false
              Get-MailContact | Remove-MailContact -Confirm:$false
            }

            #Create Users and Contacts from CSV
            $users = import-csv -Path $choiceDeployCloud['userList'] -Delimiter $choiceDeployCloud['csvdelimiter'] -Encoding UTF8 
            foreach ($user in $users) 
                {
                if ($user.External -eq "True") {
                 Write-Host -ForegroundColor DarkYellow "Creating External User" $user.DisplayName 
                 $userCreate=New-MailContact -Name $user.DisplayName -ExternalEmailAddress $user.UserPrincipalName 
                    } else {
                        if ($user.UserPrincipalName -like "sdfadm*"){
                        Write-Host -ForegroundColor Yellow "Creating administrator" $user.DisplayName "with password" $user.Password
                        $userCreate=New-MsolUser -UserPrincipalName ($user.UserPrincipalName+"@"+$tenantfqdn) -UsageLocation $user.UsageLocation -DisplayName $user.DisplayName -Password $user.Password -ForceChangePassword $false
                        } else {
                            Write-Host -ForegroundColor Yellow "Creating & Licensing user" $user.DisplayName "in Department" $user.Department "with manager" $user.manager  
                            $userCreate=New-MsolUser -UserPrincipalName ($user.UserPrincipalName+"@"+$tenantfqdn) -Department $user.Department -Country $user.Country -StreetAddress $user.StreetAddress -State $user.State -City $user.City -PostalCode $user.PostalCode -AlternateMobilePhones $user.AlternativeMobilePhone -MobilePhone $user.MobilePhone -PhoneNumber $user.HomePhone -Fax $user.Fax -FirstName $user.FirstName -LastName $user.LastName -Office $user.Office -UsageLocation $user.UsageLocation -DisplayName $user.DisplayName -Title $user.Title -Password $user.Password -ForceChangePassword $false -LicenseAssignment $choiceSubscriptions['subscriptions']
                        }
                    }
                }
            $roleName="Company Administrator"
            $adminUser = (Get-MsolUser | Where DisplayName -like "SDF Admin")
            Write-Host -ForegroundColor Yellow "Adding Admin-rights to User" $adminUser.DisplayName
            Add-MsolRoleMember -RoleMemberEmailAddress $adminUser.UserPrincipalName -RoleName $roleName

            #Divide members by department for Distribution List Membership
            $OpsGroup = $users | ?{$_.Department -like "Ops"}
            $OpsGroupUsers = $OpsGroup.UserPrincipalName.split(",")
            $AdmGroup = $users | ?{$_.Department -like "Adm"}
            $AdmGroupUsers = $AdmGroup.UserPrincipalName.split(",")
            
            #Create three Distributionlists, add members divided by departments above, and make one Distributionlist HIDDEN
            Try
            {
                $distribCreate=New-DistributionGroup -Name OpsHiddenDL -DisplayName "OpsHiddenDL" 
                Write-Host -ForegroundColor DarkCyan "Creating Hidden Distributionlist - OpsHiddenDL"
                $distribCreate=New-DistributionGroup -Name OpsDistribution -DisplayName "OpsDistribution" 
                Write-Host -ForegroundColor DarkCyan "Creating Distributionlist - OpsDistribution"
                $distribCreate=New-DistributionGroup -Name AdmDistribution -DisplayName "AdmDistribution"
                Write-Host -ForegroundColor DarkCyan "Creating Distributionlist - AdmDistribution" 
            }

            Catch
            {
                Write-Host -ForegroundColor Red "Could not create Distribution lists"
                Break
            }
            #Create three Office365 groups, add members divided by departments above, and make one O365-group HIDDEN

                Try
            {
            $o365Create=New-UnifiedGroup -DisplayName "OpsGroup" -Alias "OpsGroup" 
            Write-Host -ForegroundColor DarkCyan "Office 365 Group (Unified) OpsGroup has been created"
            $o365Create=New-UnifiedGroup -DisplayName "OpsGroupHidden" -Alias "OpsGroupHidde" 
            Write-Host -ForegroundColor DarkCyan "Hidden Office 365 Group (Unified) OpsGroupHidden has been created"
            $o365Create=New-UnifiedGroup -DisplayName "AdmGroup" -Alias "AdmGroup" 
            Write-Host -ForegroundColor DarkCyan "Office 365 Group (Unified) AdmGroup has been created"
            }
                Catch
            {
                Write-Host -ForegroundColor Red "Could not create Office 365 Groups"
                Break
            }
            #Add photos,managers, EXODS attributes, and hide mailboxes set to true in sheet
            write-host -ForegroundColor DarkCyan "Checking to see if groups created exist before trying to add Profile Pictures, and adding members to O365 Groups & Managers"
            Write-Host -ForegroundColor Green "Refreshing Exchange-connection to avoid Photo-bug"
            
            # refresh Exchange connection because of photo-bug with no fix!
            Remove-PSSession -Session $o365Session
            $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxymethod=rps -Credential $O365Cred -Authentication Basic -AllowRedirection
            Import-PSSession $O365Session -AllowClobber
            Write-Host -Foregroundcolor Green "Exchange-connection refreshed, continuing.."

            $folderPath = Split-Path -Path $choiceDeployCloud['userList']
            $ProfilePics = "$folderPath\ProfilePics\"
            $i=0
            $count = ($users.photo -like "*.png").count
            write-host -ForegroundColor DarkCyan "Adding Pictures for the users"
            foreach ($user in $users) 
                {
                if (!($user.External -eq "True" -or $User.UserPrincipalName -eq "sdfadm")) {
                    $userExists = Get-User $user.UserPrincipalName -ErrorAction SilentlyContinue
                    while ($userExists -eq $null) {
                        Write-Host -ForegroundColor Yellow "User" $user.UserPrincipalName "is not provisioned yet, sleeping"
                        Start-Sleep -Seconds 10
                        $userExists = Get-User $user.UserPrincipalName -ErrorAction SilentlyContinue
                        }
                $i++
                $Pic = $ProfilePics+$user.Photo
                Write-Host -ForegroundColor DarkCyan "Uploading picture for" $user.DisplayName "number" $i "out of $count"
                Write-Progress -Activity "Uploading photos [$count]..." -Status $i
                Try 
                {
                    $photoshoot=Set-UserPhoto $user.UserPrincipalName -PictureData ([System.IO.File]::ReadAllBytes("$Pic")) -Confirm:$false -ErrorVariable $errorPhoto

                }
                Catch {
                    Write-Warning "Error on setting UserPhoto for User $user.UserPrincipalName"
                    }
                }
            }
            Try
            {    
                Write-Host -ForegroundColor DarkCyan "Adding members to GroupOps Office365-Group"
                $add=Add-UnifiedGroupLinks -Identity "OpsGroup" -LinkType members -Links $OpsGroupUsers
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not add members to Office 365 Group GroupOps - Does it exist?"
                Write-Host -ForegroundColor DarkMagenta "We failed to read $_.Exception.Itemname. The error message was $_.Exception.Message"
                Break
            }
            Try
            {
                Write-Host -ForegroundColor DarkCyan "Adding members to GroupHiddenOps Office365-Group"
                $add=Add-UnifiedGroupLinks -Identity "OpsGroupHidden" -LinkType members -Links $OpsGroupUsers
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not add members to Office 365 Group GroupHiddenOps - Does it exist?"
                Write-Host -ForegroundColor DarkMagenta "$_.Exception.Itemname. The error message was $_.Exception.Message"
                Break
            }
            Try
            {
                Write-Host -ForegroundColor DarkCyan "Adding members to GroupAdm Office365-Group"
                $add=Add-UnifiedGroupLinks -Identity "AdmGroup" -LinkType members -Links $AdmGroupUsers
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not add members to Office 365 Group GroupAdm - Does it exist?"
                Write-Host -ForegroundColor DarkMagenta "$_.Exception.Itemname. The error message was $_.Exception.Message"
                Break
            }

            Try
            {
                Write-Host -ForegroundColor DarkCyan "Hiding GroupHiddenOps (Hidden Office365-Group)"
                $add=Set-UnifiedGroup -Identity "OpsGroupHidden" -HiddenFromAddressListsEnabled $True
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not hide Office 365 Group GroupHiddenOps - Does it exist?"
                Write-Host -ForegroundColor DarkMagenta "$_.Exception.Itemname. The error message was $_.Exception.Message"
                Break
            }

            Try
            {
                Write-Host -ForegroundColor DarkCyan "Hiding OpsHiddenDL (Hidden Distributionlist)"
                $add=Set-DistributionGroup -Identity OpsHiddenDL -HiddenFromAddressListsEnabled $true
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not hide Distribution list OpsHiddenDL - Does it exist?"
                Write-Host -ForegroundColor DarkMagenta "$_.Exception.Itemname. The error message was $_.Exception.Message"
                Break
            }

            Try
            {
            Write-Host -ForegroundColor DarkCyan "Adding members to OpsHiddenDL and OpsDistribution Distributionlists"
            foreach ($dlOps in $OpsGroupUsers)
                {
                Add-DistributionGroupMember -Identity OpsHiddenDL -Member $dlOps
                Add-DistributionGroupMember -Identity OpsDistribution -Member $dlOps
                }
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not add members to Distribution lists OpsDistribution and OpsHiddenDL - Do they exist?"
                Break
            }
            Try
            {
            Write-Host -ForegroundColor DarkCyan "Adding members to AdmDistribution Distributionlists"
            foreach ($dlAdm in $AdmGroupUsers)
                {
                Add-DistributionGroupMember -Identity AdmDistribution -Member $dlAdm
                }
            }
            Catch
            {
                Write-Host -ForegroundColor Red "Could not add members to Distribution list AdmDistribution - Does it exist?"
                Break
            }

            # exchange attributes from userlist added beneath
            foreach ($user in $users) {
                if (! ($user.UserPrincipalName -like "sdfadm" -or $user.External -eq "True" -or $user.Title -like "CEO")) {
                    Set-User -identity $user.UserPrincipalName -Company $user.Company -Manager $user.Manager -Initials $user.Initials -Pager $user.Pager -WebPage $user.WebPage  -OtherHomePhone $user.OtherHomePhone -OtherTelephone $user.OtherTelephone -OtherFax $user.OtherFax -Phone $user.Phone -Fax $user.Fax -MobilePhone $user.MobilePhone -HomePhone $user.HomePhone
                    Write-Host -ForegroundColor Cyan "Adding extra properties to user (Phone, fax, pager, etc):" $user.UserPrincipalName
                } Else {
                    if ($user.Title -like "CEO") {
                        Set-User -identity $user.UserPrincipalName -Company $user.Company  -Initials $user.Initials -Pager $user.Pager -WebPage $user.WebPage  -OtherHomePhone $user.OtherHomePhone -OtherTelephone $user.OtherTelephone -OtherFax $user.OtherFax -Phone $user.Phone -Fax $user.Fax -MobilePhone $user.MobilePhone -HomePhone $user.HomePhone
                        Write-Host -ForegroundColor Cyan "Adding extra properties to user (Phone, fax, pager, etc):" $user.UserPrincipalName
                    }
                }
     
                 if ($user.AssistantName.length -gt 2){
                    Set-User -identity $user.UserPrincipalName -AssistantName $user.AssistantName
                    Write-Host -ForegroundColor DarkCyan "Adding assistant to this user"
                }


                if ($user.HiddenMB -gt 2) {
                    Set-Mailbox -Identity $user.UserPrincipalName -HiddenFromAddressListsEnabled $true
                    Write-Host -ForegroundColor DarkCyan "Hiding mailbox on this user"
                }
            }

            # Sharepoint attributes section
            $PnPfile = "$folderpath\PnPAttributes.csv"
                $wshell = New-Object -ComObject Wscript.Shell  
                try {  
                    $ctx = Get-PnPContext  
                } catch {  
                    $wshell.Popup("Please connect to tenant admin site!", 0, "Done", 0x1)  
                }  
                if ($ctx) {  
                    $UserData = Import-Csv $PnPfile -Encoding UTF8 -Delimiter $choiceDeployCloud['csvdelimiter']

                    $rows = $UserData | measure  

                    $ColumnName = $userdata[0].psobject.properties.name
         
                    for ($i = 0; $i -lt $rows.Count; $i++) {  
                        $UserPrincipalName = $UserData[$i].UserPrincipalName
             
                        Write-Host "Updating Sharepoint-data for $UserPrincipalName"  

                        for ($j = 1; $j -lt $ColumnName.Count; $j++) {  
                            $value = $UserData[$i].($ColumnName[$j])  
                            if (($value.Length -ge 3) -and(($value.Substring(0, 3) -eq "i:0") -or($value.SubString(0, 3) -eq "c:0"))) {  
                               # if ($ColumnName.Name -ne "UserPrincipalName") {
                                Set-PnPUserProfileProperty -Account "$UserPrincipalName@$tenantfqdn" -PropertyName $ColumnName[$j] -Values $value -ErrorAction SilentlyContinue -Verbose
                               # } 
                            } else {# split the string using the | as a delimiter and load the values into the field. 
                               # if ($ColumnName.Name -ne "UserPrincipalName") { 
                                Set-PnPUserProfileProperty -Account "$UserPrincipalName@$tenantfqdn" -PropertyName $ColumnName[$j] -Values $value.Split("|") -ErrorAction SilentlyContinue 
                               # } 
                            }  
                            if ($?) {  
                                Write-Host "Set $($ColumnName[$j]) --> $($UserData[$i].($ColumnName[$j]))." -ForegroundColor Green  
                            } else {  
                                Write-Host " Could not set $($ColumnName[$j]) --> $($UserData[$i].($ColumnName[$j])). $($error[0].Exception.message)" -ForegroundColor Red  
                            }  
                        }  
                    }        
                }
            
            # addressbook policy section
            If ($choiceDeployCloud['adrBookPolicy'] -eq "Yes") {
                Write-Host -ForegroundColor Yellow "You chose Address Book Policy deployment"    
                Remove-PSSession -Session $o365Session
                Write-Host "Refreshing connection to Exchange Online for new roles"
                $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
                Import-PSSession $O365Session -AllowClobber
                Connect-MsolService –Credential $O365Cred

                #Check if Role is added to Organization Management, if not...add it
                $orgRoles = Get-ManagementRoleAssignment -RoleAssignee "Organization Management"
                if ( $orgRoles.role -like "*Address Lists") {
                    write-host -ForegroundColor Green -BackgroundColor DarkGray "User already has Address List role, continuing"
                } else {
                    Write-Host -ForegroundColor Red -BackgroundColor White "WARNING  `nThis will most likely fail, since there is a delay between adding the permission, and being able to use the cmdlets below  `nAdding the user to Address Lists to be able to configure ABP" 
                    New-ManagementRoleAssignment -SecurityGroup "Organization Management" -Role "Address Lists" 
                }

            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating the Address Book Policy"
            $tenantFQDN = ($O365Cred.UserName -split("@"))[1]

            #Adding wildcard strings for Departments
            $firstDepWild = "$firstDep*"
            $secondDepWild = "$secondDep*"

            #Check to see if ABPRouting is disabled (default), enable if not
            $transportConfig = Get-TransportConfig
            if (!($transportConfig.AddressBookPolicyRoutingEnabled -match "true")) {
                Write-Host -ForegroundColor Red -BackgroundColor White "TransportConfig AddressBookPolicyRouting is not enabled. Enabling now"
                #Enable Routing for segmentation
                Set-TransportConfig -AddressBookPolicyRoutingEnabled $true
                Write-Host -ForegroundColor green -BackgroundColor DarkGray "TransportConfig AddressBookPolicyRouting is now enabled on tenant. TransportConfig is set to"$transportConfig.AddressBookPolicyRoutingEnabled
                } Else {
                Write-Host "TransportConfig already enabled on tenant. TransportConfig is set to"$transportConfig.AddressBookPolicyRoutingEnabled
                }


            #Create Global Address List based on Department (this can be changed)
            if (!(Get-GlobalAddressList -Identity GAL_Administrative -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Global Address List (GAL) based on department:"$firstDep
            $GalADm=New-GlobalAddressList -Name GAL_Administrative -RecipientFilter "Department -like '$firstDep'"
            } Else {
            Write-Host "GlobalAddressList GAL_Administrative already exists, continuing"
            }

            #Create New Address list based on users in department
            if (!(Get-AddressList -Identity "All Administrative Users" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List for all users based on department:"$firstDep
            $adrList=New-AddressList -Name "All Administrative Users" -RecipientFilter "Department -like '$firstDep'"
            } Else {
            Write-Host "AddressList All Administrative Users already exists, continuing"
            }

            #Roomlist is required, therefore we create one in our test-tenant
            if (!(Get-AddressList -Identity "All Administrative Rooms" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List based on rooms in tenant for department:"$firstDep
            $adrList=New-AddressList -Name "All Administrative Rooms" -RecipientFilter {((Alias -ne $null) -and (((RecipientDisplayType -eq 'ConferenceRoomMailbox') -or (RecipientDisplayType -eq 'SyncedConferenceRoomMailbox'))))}
            } Else {
            Write-Host "AddressList All Administrative Rooms already exists, continuing"
            }

            #Create AddressList for Distributiongroups
            if (!(Get-AddressList -Identity "All Administrative DistributionLists" -ErrorAction Ignore)){
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List based on Distributiongroups in tenant for department:"$firstDep
            $adrList=New-AddressList -name "All Administrative DistributionLists" -RecipientFilter "Name -like '$firstDepWild' -and RecipientTypeDetails -eq 'MailUniversalDistributionGroup'"
            } Else {
            Write-Host "AddressList All Administrative DistributionLists already exists, continuing"
            }

            #New Offline Address Book for Administrative users
            if (!(Get-OfflineAddressBook -Identity "Administrative Offline Address Book" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Offline Address Book (OAB) for department:"$firstDep
            $offAdrBook=New-OfflineAddressBook -Name "Administrative Offline Address Book" -AddressLists "Gal_Administrative"
            } Else {
            Write-Host "Offline Address Book Administrative Offline Address Book already exists, continuing"
            }

            #New Address Book Policy for Administrative users (ABP)
            if (!(Get-AddressBookPolicy -Identity "Administrative ABP" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating unassigned Address Book Policy based on department:"$firstDep
            $adrBookPolicy=New-AddressBookPolicy -Name "Administrative ABP" -AddressLists "All Administrative Users”,"All Administrative DistributionLists" -RoomList "All Administrative Rooms" -OfflineAddressBook "Administrative Offline Address Book" -GlobalAddressList "GAL_Administrative"
            } Else {
            Write-Host "AddressBookPolicy Administrative ABP already exists, continuing"
            }

            #Create Global Address List based on Department (this can be changed)
            if (!(Get-GlobalAddressList -Identity GAL_Operations -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Gloobal Address List (GAL) based on department:"$secondDep
            $galOps=New-GlobalAddressList -Name GAL_Operations -RecipientFilter "Department -like '$secondDep'"
            } Else {
            Write-Host "GlobalAddressList GAL_Operations already exists, continuing"
            }

            #Create New Address list based on users in department
            if (!(Get-AddressList -Identity "All Operations Users" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List based on department:"$secondDep
            $adrList=New-AddressList -Name "All Operations Users" -RecipientFilter "Department -like '$secondDep'"
            } Else {
            Write-Host "AddressList All Operations Users already exists, continuing"
            }

            #Roomlist is required, therefore we create one in our test-tenant
            if (!(Get-AddressList -Identity "All Operations Rooms" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List based on rooms in tenant for department:"$secondDep
            $adrList=New-AddressList -Name "All Operations Rooms" -RecipientFilter {((Alias -ne $null) -and (((RecipientDisplayType -eq 'ConferenceRoomMailbox') -or (RecipientDisplayType -eq 'SyncedConferenceRoomMailbox'))))}
            } Else {
            Write-Host "AddressList All Operations Rooms already exists, continuing"
            }

            #Create AddressList for Distributiongroups
            if (!(Get-AddressList -Identity "All Operations DistributionLists" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Address List based on Distributiongroups in tenant for department:"$secondDep
            $adrList=New-AddressList -Name "All Operations DistributionLists" -RecipientFilter "Name -like '$secondDepWild' -and RecipientTypeDetails -eq 'MailUniversalDistributionGroup'"
            } Else {
            Write-Host "AddressList All Operations DistributionLists already exists, continuing"
            }

            #New Offline Address Book for Operations
            if (!(Get-OfflineAddressBook -Identity "Operations Offline Address Book" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating new Offline Address Book (OAB) for department:"$secondDep
            $offAdrBook=New-OfflineAddressBook -Name "Operations Offline Address Book" -AddressLists "Gal_Operations"
            } Else {
            Write-Host "Offline Address Book Operations Offline Address Book already exists, continuing"
            }

            #New Address Book Policy for Administrative (ABP)
            if (!(Get-AddressBookPolicy -Identity "Operations ABP" -ErrorAction Ignore)) {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Creating unassigned Address Book Policy based on department:"$secondDep
            $adrBookPolicy=New-AddressBookPolicy -Name "Operations ABP" -AddressLists "All Operations Users”,"All Operations Distributionlists" -RoomList "All Operations Rooms" -OfflineAddressBook "Operations Offline Address Book" -GlobalAddressList "GAL_Operations"
            } Else {
            Write-Host "AddressBookPolicy Operations ABP already exists, continuing"
            }

            $allUserMbx = Get-User -RecipientTypeDetails UserMailbox -Resultsize unlimited
            #assign "DL/Adm ABP" Address Book Policy to DL/Adm Department
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Assigning Address Book Policy based on department:"$firstDep
            if ((($allUserMbx|?{($_.Department -match $firstDep)}).Count) -lt 1) {
            Write-Host 'There are no users with Department "Adm" in your tenant, have you run CreateUsers.ps1?'
            } Else {
            $allUserMbx | ?{($_.Department -match $firstDep)} | Set-Mailbox -AddressBookPolicy “Administrative ABP”
            }
            #assign "Drift ABP" Address Book Policy to Drift users
            if ((($allUserMbx|?{($_.Department -match $secondDep)}).Count) -lt 1) {
            Write-Host 'There are no users with Department "Adm" in your tenant, have you run CreateUsers.ps1?'
            } Else {
            Write-Host -ForegroundColor green -BackgroundColor DarkGray "Assigning Address Book Policy based on department:"$secondDep
            $allUserMbx | ?{($_.Department -match $secondDep)} | Set-Mailbox -AddressBookPolicy “Operations ABP”
            }

            #Tickle the mailboxes so they update
            $mailboxes = Get-Mailbox -Resultsize Unlimited
            $count = $mailboxes.count
            $i=0

            Write-Host
            Write-Host "Mailboxes Found:" $count

            foreach($mailbox in $mailboxes){
                $i++
                Set-Mailbox $mailbox.alias -SimpleDisplayName $mailbox.SimpleDisplayName -WarningAction silentlyContinue
                Write-Progress -Activity "Tickling Mailboxes [$count]..." -Status $i
            }

            $mailusers = Get-MailUser -Resultsize Unlimited
            $count = $mailusers.count
            $i=0

            Write-Host
            Write-Host "Mail Users Found:" $count

            foreach($mailuser in $mailusers){
                $i++
                Set-MailUser $mailuser.alias -SimpleDisplayName $mailuser.SimpleDisplayName -WarningAction silentlyContinue
                Write-Progress -Activity "Tickling Mail Users [$count]..." -Status $i
            }

            $distgroups = Get-DistributionGroup -Resultsize Unlimited
            $count = $distgroups.count
            $i=0

            Write-Host
            Write-Host "Distribution Groups Found:" $count

            foreach($distgroup in $distgroups){
                $i++
                Set-DistributionGroup $distgroup.alias -SimpleDisplayName $distgroup.SimpleDisplayName -WarningAction silentlyContinue
                Write-Progress -Activity "Tickling Distribution Groups. [$count].." -Status $i
            }

            Write-Host
            Write-Host "Tickling Complete"

            #Last two users in userlist are hidden from addresslist in Exchange Online

            Write-Host "To verify that the Global Addresslist and Addresslist is populated: "
            Write-Host
            Write-Host -ForegroundColor DarkCyan "AddressList All Administrative Users"
            $allusers = Get-AddressList "All Administrative Users"
            Get-Recipient -RecipientPreviewFilter $allusers.RecipientFilter | Select-Object DisplayName

            Write-Host -ForegroundColor DarkCyan "AddressList All Operations Users"
            $allusers = Get-AddressList "All Operations Users"
            Get-Recipient -RecipientPreviewFilter $allusers.RecipientFilter | Select-Object DisplayName

            Write-Host -ForegroundColor DarkCyan "AddressList All Administrative DistributionLists"
            $allusers = Get-AddressList "All Administrative Distributionlists"
            Get-Recipient -RecipientPreviewFilter $allusers.RecipientFilter | Select-Object DisplayName

            Write-Host -ForegroundColor DarkCyan "AddressList All Operations DistributionLists"
            $allusers = Get-AddressList "All Operations Distributionlists"
            Get-Recipient -RecipientPreviewFilter $allusers.RecipientFilter | Select-Object DisplayName


            Write-Host -ForegroundColor DarkCyan "GAL_Administrative Members"
            $GAL = Get-GlobalAddressList -Identity GAL_Administrative; $admRecipients=Get-Recipient -ResultSize unlimited -RecipientPreviewFilter $GAL.RecipientFilter | select DisplayName, PrimarySMTPAddress
            foreach ($admRecipient in $admRecipients){
                Write-Host -ForegroundColor Yellow $admRecipient.DisplayName 
            }
            Write-Host -ForegroundColor DarkCyan "GAL_Operations Members"
            $GAL = Get-GlobalAddressList -Identity GAL_Operations; $opsRecipients=Get-Recipient -ResultSize unlimited -RecipientPreviewFilter $GAL.RecipientFilter | select DisplayName, PrimarySMTPAddress
            foreach ($opsRecipient in $opsRecipients){
                Write-Host -ForegroundColor Yellow $opsRecipient.DisplayName
            }
        }
    Show-AnyBox -Message 'User Deployment complete! You can now log on any user on the userlist' -Buttons 'OK' -Icon 
    }
}

###################################################
#                                                 #
#      Azure Deployment for AD, Exchange and      #
#      Sharepoint below. Do not share the app     #
#      secret "securePass" as it is a Service     #
#      Principal in Azure with contributor        #
#      rights to Office People Azure.             #
#                                                 #
#      In other words it has unlimited rights     #
#      to deploy as many resources as it wants    #
#                                                 #
#      Contact Margrete Sævareid                  #
#      margrets@microsoft.com if you have any     #
#      questions.                                 #
#                                                 #
###################################################

if ($choiceOnPremCloud['onPrem']) {
    Write-Host -ForegroundColor Green "Checking if Azure Powershell is installed"
    if (Get-InstalledModule -Name az -ErrorAction SilentlyContinue){
        Write-Host -ForegroundColor Green "Azure Powershell Module exists "
        $online = Find-Module -Name az
        $local = Get-InstalledModule -Name az
        if ($online.Version -gt $local.Version)
        {
        Write-Host -ForegroundColor Green "Updating Azure Powershell module to latest version"
        Update-Module az -Force -Confirm:$false
        }
    } else {
        Write-Host -ForegroundColor Red "Azure Powershell does not exist, installing now"
        Install-module -name az -Force -AllowClobber
        Write-Host -ForegroundColor Green "Azure Powershell Module is now installed"
    }

    #Import Azure Module
    Import-Module -name Az


    #Service Principal AppID
    $appID = ""
    #Service Principal Secret
    $securePass = "" | ConvertTo-SecureString -AsPlainText -Force
    # azure tenantID
    $tenant = ""
    # Azure subscriptionID
    $azSubscriptionID = ""
    Write-Warning "Service Principal secret key is valid until 01/04/2021. Script needs a new secret for ApplicationID: $appID`nName of application is:FastSP"
    #Create SP Credential
    $azCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $appID, $securePass

    #Connect Azure
    Write-Host -ForegroundColor Yellow "Connecting to Microsoft Azure..."
    Try 
    {
        Connect-AzAccount -ServicePrincipal -Credential $azCred -Tenant $tenant
    }
    Catch
    {
    Write-Warning "Script failed to connect to Azure." 
    Break
    }
    Write-Host -ForegroundColor Green "Connected!"

    #Select valid Office365 People Subscription for Resource deployments
    Write-Host -ForegroundColor Yellow "Connecting to Office People subscription"
    $subscription = Select-AzSubscription -Subscription $azSubscriptionID
    $userName = $env:USERNAME
    $createdDate = get-date -UFormat "%d.%m.%y"
    $countries = (Get-AzLocation).Location
    $anybox = New-Object AnyBox.AnyBox

    $anybox.Prompts = @(
      # Domainname Input with default MSFast-owned domain.
      New-AnyBoxPrompt -Name "domainName" -InputType Text -Message "What is your domainname?" -DefaultValue "msfastlabs.com"
      # Environment-type dropdown menu
      New-AnyBoxPrompt -Name "environment" -InputType Text -Message "What are your environment needs?" -ValidateSet 'Exchange Hybrid','Sharepoint OnPrem','Both Exchange & Sharepoint' -DefaultValue 'Exchange Hybrid'
      # Azure datacenter default choices
      New-AnyBoxPrompt -Name "rgLocation" -InputType Text -Message "Which Azure Datacenter would you to deploy the environment to?" -ValidateSet @($countries) -DefaultValue 'northeurope'
      # Azure datacenter default choices
      New-AnyBoxPrompt -Name "tenantType" -InputType Text -Message "Will your Office 365 tenant be a Production Environment or SDFv2?" -ValidateSet 'Production', 'SDFv2' -ShowSetAs Radio_Wide -ValidateNotEmpty
      # date output for later automation/governance in azure
      New-AnyBoxPrompt -Name "date" -InputType Text -Message "Date" -DefaultValue $createdDate -ReadOnly
      # link to OneNote
      New-AnyBoxPrompt -Name "urlOneNote" -InputType Link -Message "OneNote Deployment guide" -DefaultValue 'onenote:https://microsoft.sharepoint.com/teams/OfficePeople/SiteAssets/Office%20People%20Notebook/Projects/Test%20Lab.one#Guides%20Hybrid&section-id={A70B945C-4284-48F6-983A-30A0C55E78C1}&page-id={F2CE8291-CE63-4B6E-9AE0-4431C555929F}&end'
    )

    $anybox.ContentAlignment = 'Center'
    $anybox.Buttons += New-AnyBoxButton -Name 'Submit' -IsDefault -Text 'Submit'
    $anybox.Buttons += New-AnyBoxButton -Name 'Cancel' -IsCancel -Text 'Cancel'
    $anybox.Title = 'OP Testlabs'
    $anybox.Icon = 'Question'
    $anybox.Comment += 'WARNING: Non-default domain (msfastlabs.com) require user input on DNS after script finish'
    $anybox.Comment += "Transcript can be found at $path" 

    $choiceDeploy = $anybox | Show-AnyBox

    # abort if cancel is chosen
    if ($choiceDeploy['submit'] -eq $false) {
    Show-AnyBox -Message "You have chosen to cancel the deployment. Script will now end." -Buttons 'OK'
    Break
    }

    $i = 1

    # Exchange Hybrid Resource Group creation in Azure
    if ($choiceDeploy['tenantType'] -eq "Production" -and $choiceDeploy['environment'] -eq "Exchange Hybrid" )
    {
        $resourceGroupName = "RGProdExch$i"
        while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
            {
                $i++
                $resourceGroupName = "RGProdExch$i"
            }
        Write-Host -ForegroundColor Green "Setting the Resource Group Name to the Prod naming standard $resourceGroupName" 
        $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
    }
        elseif ($choiceDeploy['tenantType'] -eq "SDFv2" -and $choiceDeploy['environment'] -eq "Exchange Hybrid" ) 
        {
            $resourceGroupName = "RGSDFExch$i"
            while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
                {
                    $i++
                    $resourceGroupName = "RGSDFExch$i"
                } 
            Write-Host -ForegroundColor Green "Setting the Resource Group Name to the SDFv2 naming standard $resourceGroupName"
            $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
        }

    # Sharepoint Resource Group creation in Azure
    if ($choiceDeploy['tenantType'] -eq "Production" -and $choiceDeploy['environment'] -eq "Sharepoint OnPrem" )
    {
        $resourceGroupName = "RGProdShPo$i"
        while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
            {
                $i++
                $resourceGroupName = "RGProdShPo$i"
            }
        Write-Host -ForegroundColor Green "Setting the Resource Group Name to the Prod naming standard $resourceGroupName" 
        $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
    }
        elseif ($choiceDeploy['tenantType'] -eq "SDFv2" -and $choiceDeploy['environment'] -eq "Sharepoint") 
        {
            $resourceGroupName = "RGSDFShPo$i"
            while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
                {
                    $i++
                    $resourceGroupName = "RGSDFShPo$i"
                } 
            Write-Host -ForegroundColor Green "Setting the Resource Group Name to the SDFv2 naming standard $resourceGroupName"
            $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
        }

    # Full environment SP+Exch Resource Group creation in Azure
    if ($choiceDeploy['tenantType'] -eq "Production" -and $choiceDeploy['environment'] -eq "Both Exchange & Sharepoint" )
    {
        $resourceGroupName = "RGProdFull$i"
        while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
            {
                $i++
                $resourceGroupName = "RGProdFull$i"
            }
        Write-Host -ForegroundColor Green "Setting the Resource Group Name to the Prod naming standard $resourceGroupName" 
        $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
    }
        elseif ($choiceDeploy['tenantType'] -eq "SDFv2" -and $choiceDeploy['environment'] -eq "Both Exchange & Sharepoint" ) 
        {
            $resourceGroupName = "RGSDFFull$i"
            while(Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue)
                {
                    $i++
                    $resourceGroupName = "RGSDFFull$i"
                } 
            Write-Host -ForegroundColor Green "Setting the Resource Group Name to the SDFv2 naming standard $resourceGroupName"
            $azResourceGroup = New-AzResourceGroup -Name $resourceGroupName -Location $choiceDeploy['rgLocation'] -Tag @{PCUsername=$userName;Email=$choiceOnPremCloud['emailAdr'];CreatedDate=$choiceDeploy['date']}
        }

    # If domain is default (msfastlabs.com), create a subdomain based on resourcegroupname. This if for 
    if ($choiceDeploy['domainName'] -eq "msfastlabs.com")
    {
        $choiceDeploy['domainName'] = "$resourceGroupName.$($choiceDeploy['domainName'])"
    }

    # remove RG from resourcenames
    $resourcenames = $resourceGroupName.ToLower()
    $resourcenames = $resourcenames.Replace('rg','')

    # Deploy Exchange
    if ($choiceDeploy['environment'] -eq "Exchange Hybrid")
        {
            $deploy = Show-AnyBox -Title "Exchange  Deployment" -Message "Deploying Exchange and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 90 minutes." -Buttons 'Ok','Cancel' -WindowStyle ToolWindow
            if ($deploy['Ok'])
        {
            Write-Host -Foregroundcolor Green "Deploying Exchange and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 90 minutes."
            New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri https://raw.githubusercontent.com/nimsim/Templates/master/Exchange/azuredeploy.json -domainName $choiceDeploy['domainName']
        
            $domainName = $choicedeploy['domainName']

            $IPadr = $null

            $IPadr = (Get-AzPublicIpAddress -ResourceGroupName $resourceGroupName).IpAddress
            Write-Host -ForegroundColor Green "Deployment completed! The IP address is" $IPadr
            Write-Host -ForegroundColor Green "You can now log on to the server by connecting to:" $IPadr":65221"
            Write-Host -ForegroundColor Green "Username is $domainname\SDFadm"


            $dnsInfo = @(
                @{Hostname="$domainName";Type="MX";PointsTo="mail.$domainName"},
                @{Hostname="$null";Type="TXT";PointsTo="v=spf1 ip4:$IPadr include:spf.protection.outlook.com -all"},
                @{Hostname="autodiscover.$domainName";Type="CNAME";PointsTo="mail.$domainName"},
                @{Hostname="owa.$domainName";Type="CNAME";PointsTo="mail.$domainName"},
                @{Hostname="mail.$domainName";Type="A-Record";PointsTo="$IPadr"}) | % { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru}

            $dnsInfo | Export-Csv -LiteralPath C:\Temp\DNSInfo$ipadr.csv -NoTypeInformation
            Write-Host -ForegroundColor Yellow "The following DNS-records are required in the domain $domainName DNS"
            $dnsInfo
        
            } else 
                {
                    Show-AnyBox -Title 'Deplyoment Aborted' -Message "The deployment has been cancelled" -Buttons 'Ok'
                    break
                }
        }

    # Deploy Sharepoint
    if ($choiceDeploy['environment'] -eq "Sharepoint")
        {
            $deploy = Show-AnyBox -Title "Sharepoint Deployment" -Message "Deploying Sharepoint and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 3 hours." -Buttons 'Ok','Cancel' -WindowStyle SingleBorderWindow
            if ($deploy['Ok'])
            {
                Write-Host -Foregroundcolor Green "Deploying Sharepoint and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 3 hours."
                New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri https://raw.githubusercontent.com/nimsim/Templates/dev/Sharepoint/azuredeploy.json -domainName $choiceDeploy['domainName']
            } else 
                {
                    Show-AnyBox -Title 'Deplyoment Aborted' -Message "The deployment has been cancelled" -Buttons 'Ok'
                    break
                }
        }
    if ($choiceDeploy['environment'] -eq "Both Exchange & Sharepoint")
        {
            $deploy = Show-AnyBox -Title "Exchange and Sharepoint Deployment" -Message "Deploying Exchange, Sharepoint, and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 4 hours." -Buttons 'Ok','Cancel' -WindowStyle SingleBorderWindow
            if ($deploy['Ok'])
            {
            #Write-Host -Foregroundcolor Green "Deploying Exchange, Sharepoint, and Active Directory resources to ResourceGroup: $resourceGroupName.`nThis may take up to 3 hours."
            #New-AzResourceGroupDeployment -ResourceGroupName $resourceGroupName -TemplateUri https://raw.githubusercontent.com/nimsim/Templates/dev/sharepoint/azuredeploy.json -domainName $choiceDeploy['domainName']
            Show-AnyBox -Title 'Placeholder' -Message "this doesn't do anything yet.." -buttons 'OK'
            } else 
                {
                    Show-AnyBox -Title 'Deplyoment Aborted' -Message "The deployment has been cancelled" -Buttons 'Ok'
                    break
                }
        }

    # Add Send Connector manually [NOT WORKING] Need to do lots more than a simple line
    $fqdn = "vm$resourceGroupName.$domainName"
    # Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName -VMName "vm$resourceGroupName" -CommandId RunPowerShellScript 


    $azureVMs = (Get-AzVM -ResourceGroupName $resourceGroupName).Name
    # building choice GUI for connecting to servers 
    $anybox = New-Object AnyBox.AnyBox

    $anybox.Message = 'Deployment completed with no errors!', 'Would you like to connect to the VM?'
    $anybox.Prompts += New-AnyBoxPrompt -Name 'srvName' -Message "Which server?" -ValidateSet @($azureVMs)
    $anybox.Prompts += New-AnyBoxPrompt -Name 'saveLocation' -Message 'Save RDP-files to folder' -InputType FolderOpen -ReadOnly

    $anybox.Buttons = @(
        New-AnyBoxButton -Name 'connect' -Text 'Connect' -IsDefault
        New-AnyBoxButton -Name 'save' -Text 'Save' 
        New-AnyBoxButton -Name 'cancel' -Text 'Cancel' -IsCancel
    )

    # show GUI.
    $finishedDeploy = $anybox | Show-AnyBox

    # act on response
    if ($finishedDeploy['save'] -eq $true) {
        # Folder location to variable
        $folderLocation=$finishedDeploy['saveLocation']
        # for each VM, store them in the folderlocation
        foreach ($VM in $azureVMs)
            {
                Get-AzRemoteDesktopFile -ResourceGroupName $resourceGroupName -Name $vm -LocalPath "$folderLocation\$VM.rdp" -ErrorAction SilentlyContinue
            }
            $choiceSave = Show-AnyBox -Message "your RDP files have been saved to the following location:", "$folderLocation\" -Buttons 'OK', "Open Folder"
            if ($choiceSave['Open Folder'] -eq $true) {
                Invoke-Item $folderLocation
                }
    } elseif ($finishedDeploy['connect'] -eq $true) {
            Get-AzRemoteDesktopFile -ResourceGroupName $resourceGroupName -Name $finishedDeploy['srvName'] -Launch 
    
        } else {
        Show-AnyBox -Message "You have chosen to cancel." -Buttons 'OK'
        Break
        }
}

Stop-Transcript

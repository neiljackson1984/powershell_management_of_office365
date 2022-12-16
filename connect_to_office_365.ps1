
#To get pre-requisites:
# Install-Module -Confirm:$false -Force -Name 'AzureAD', 'ExchangeOnlineManagement', 'PnP.PowerShell'
# Install-Module -Confirm:$false -Force -Name 'AzureADPreview', 'ExchangeOnlineManagement', 'PnP.PowerShell'
# UnInstall-Module -Confirm:$false -Force -Name 'AzureAD'
# UnInstall-Module -Confirm:$false -Force -Name 'AzureADPreview'
# to make this work with Powershell Core (which, as of 2021-10-26, does not work out of the box with the AzureAD module), install the following special version of the AzureAD module as follows:
# (thanks to https://endjin.com/blog/2019/05/how-to-use-the-azuread-module-in-powershell-core)
###    # Check if test gallery is registered
###    $packageSource = Get-PackageSource -Name 'Posh Test Gallery'
###    if (!$packageSource)
###    {
###    	$packageSource = Register-PackageSource -Trusted -ProviderName 'PowerShellGet' -Name 'Posh Test Gallery' -Location 'https://www.poshtestgallery.com/api/v2/'
###    }
###    
###    # Check if module is installed
###    $module = Get-Module 'AzureAD.Standard.Preview' -ListAvailable -ErrorAction SilentlyContinue
###    
###    if (!$module) 
###    {
###      Write-Host "Installing module AzureAD.Standard.Preview ..."
###      $module = Install-Module -Name 'AzureAD.Standard.Preview' -Force -Scope CurrentUser -SkipPublisherCheck -AllowClobber 
###      Write-Host "Module installed"
###    }
# when I attempt connect-azuread in powershell core (even when I am using the version of connect-azuread from the AzureAD.Standard.Preview module),
# I encounter the error "Certificate based authentication is not supported in netcore version."
# I take that as the nail in the coffin of the hope of using this script from within powershell core (for now).
# Install-Module -Confirm:$false -Force -Name 'AzureAD', 'ExchangeOnlineManagement', 'PnP.PowerShell'
# TODO (potentially): check the version of powershell that we are running under and throw some kind of error or warning if we notice that
# we are running under powershell core, because the AzureAD module does not quite work correctly under powershell core, it seems.


# # update 2022-09-16:
# # to get prerequisistes:
# #   AzureADPreview (and AzureAD) STILL does not work completely correctly under powershell core.
# #   The -UseWindowsPowerShell  option of powershell core's Import-Module function
# #   seemed promising as a way to use the windowsPowershell module from within powershell core,
# #   , but the serializing of the command output is a dealbreaker.  Therefore, we are STILL
# #   constrained to use windowsPowershell and not powershell core.
# powershell -c "Install-Module -Confirm:0 -Force -Name AzureADPreview"
# powershell -c "Install-Module -Confirm:0 -Force -Name ExchangeOnlineManagement -AllowPrerelease"
# powershell -c "Install-Module -Confirm:0 -Force -Name PnP.PowerShell"

# powershell -c "Install-Module -Confirm:0 -Force -Name Microsoft.Graph"; pwsh -c "Install-Module -Confirm:0 -Force -Name Microsoft.Graph"


# the AzureADPreview module is being deprecated, and replaced with "Microsoft Graph Powershell"
# see https://learn.microsoft.com/en-us/powershell/azure/active-directory/migration-faq?view=azureadps-2.0 
# see https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0
# see https://practical365.com/connect-microsoft-graph-powershell-sdk/


[CmdletBinding()]
Param (

    
    
    [Parameter(HelpMessage=
        @"
The path of the configuration file.
"@
    )]
    [String]$pathOfTheConfigurationFile = "config.json" # (Join-Path $PSScriptRoot "config.json")
)

# Import-Module -Name 'AzureAD'  -UseWindowsPowerShell -ErrorAction SilentlyContinue
# Import-Module -Name 'AzureADPreview'   -UseWindowsPowerShell 
Import-Module -Name 'AzureADPreview'   
Import-Module -Name 'ExchangeOnlineManagement'
Import-Module -Name 'PnP.PowerShell'
# Import-Module -Name 'Microsoft.Graph'

$certificateStorageLocation = "cert:\localmachine\my"


.{$roleSpecifications = `
    @(
        @{
            displayNameOfTargetServicePrincipal = 'Windows Azure Active Directory';
            namesOfAppRoles = @(
                'Policy.Read.All',
                'Directory.Read.All',
                'Domain.ReadWrite.All',
                'Directory.ReadWrite.All',
                'Device.ReadWrite.All',
                'Member.Read.Hidden',
                'Application.ReadWrite.OwnedBy',
                'Application.ReadWrite.All'
            )
        },
        @{
            displayNameOfTargetServicePrincipal = 'Office 365 Exchange Online';
            namesOfAppRoles = @(
                'Exchange.ManageAsApp'
            )
        },
        @{
            displayNameOfTargetServicePrincipal = 'Office 365 Management APIs';
            namesOfAppRoles = @(
                'ServiceHealth.Read',
                'ActivityFeed.Read',
                'ActivityFeed.ReadDlp'
            )
        },
        @{
            displayNameOfTargetServicePrincipal = 'Microsoft Graph';
            namesOfAppRoles = @(
                'Sites.Selected',
                'ChatMember.ReadWrite.All',
                'DataLossPreventionPolicy.Evaluate',
                'SensitivityLabel.Evaluate',
                'APIConnectors.ReadWrite.All',
                'TeamsTab.ReadWriteForUser.All',
                'TeamsTab.ReadWriteForChat.All',
                'Policy.Read.ConditionalAccess',
                'ShortNotes.ReadWrite.All',
                'ServiceMessage.Read.All',
                'TeamMember.ReadWriteNonOwnerRole.All',
                'TeamsAppInstallation.ReadWriteSelfForUser.All',
                'TeamsAppInstallation.ReadWriteSelfForTeam.All',
                'TeamsAppInstallation.ReadWriteSelfForChat.All',
                'TeamsAppInstallation.ReadForUser.All',
                'TeamsAppInstallation.ReadForChat.All',
                'Teamwork.Migrate.All',
                'PrintJob.ReadWriteBasic.All',
                'PrintJob.Read.All',
                'PrintJob.Manage.All',
                'Printer.ReadWrite.All',
                'Printer.Read.All',
                'Policy.ReadWrite.PermissionGrant',
                'Policy.Read.PermissionGrant',
                'Policy.ReadWrite.AuthenticationMethod',
                'Policy.ReadWrite.AuthenticationFlows',
                'TeamMember.Read.All',
                'TeamSettings.ReadWrite.All',
                'Channel.ReadBasic.All',
                'ChannelSettings.Read.All',
                'UserShiftPreferences.Read.All',
                'Device.Read.All',
                'Policy.ReadWrite.ApplicationConfiguration',
                'TeamsTab.ReadWrite.All',
                'TeamsTab.Read.All',
                'TeamsTab.Create',
                'UserAuthenticationMethod.Read.All',
                'UserAuthenticationMethod.ReadWrite.All',
                'Policy.ReadWrite.ConditionalAccess',
                'Schedule.ReadWrite.All',
                'BitlockerKey.ReadBasic.All',
                'BitlockerKey.Read.All',
                'TeamsApp.Read.All',
                'ApprovalRequest.ReadWrite.CustomerLockbox',
                'PrivilegedAccess.Read.AzureAD',
                'TeamsActivity.Send',
                'TeamsActivity.Read.All',
                'DelegatedPermissionGrant.ReadWrite.All',
                'OrgContact.Read.All',
                'Calls.InitiateGroupCall.All',
                'Calls.JoinGroupCall.All',
                'Calls.JoinGroupCallAsGuest.All',
                'OnlineMeetings.Read.All',
                'OnlineMeetings.ReadWrite.All',
                'IdentityUserFlow.ReadWrite.All',
                'Calendars.Read',
                'Device.ReadWrite.All',
                'Directory.ReadWrite.All',
                'Group.Read.All',
                'Mail.ReadWrite',
                'MailboxSettings.Read',
                'Domain.ReadWrite.All',
                'Application.ReadWrite.All',
                'Chat.UpdatePolicyViolation.All',
                'People.Read.All',
                'AccessReview.ReadWrite.All',
                'Application.ReadWrite.OwnedBy',
                'User.ReadWrite.All',
                'EduAdministration.Read.All',
                'EduAssignments.ReadWrite.All',
                'EduAssignments.ReadWriteBasic.All',
                'EduRoster.Read.All',
                'IdentityRiskyUser.ReadWrite.All',
                'IdentityRiskEvent.ReadWrite.All',
                'SecurityEvents.Read.All',
                'Sites.Read.All',
                'SecurityActions.ReadWrite.All',
                'ThreatIndicators.ReadWrite.OwnedBy',
                'AdministrativeUnit.Read.All',
                'OnPremisesPublishingProfiles.ReadWrite.All',
                'DeviceManagementServiceConfig.Read.All',
                'DeviceManagementManagedDevices.Read.All',
                'AccessReview.ReadWrite.Membership',
                'Place.Read.All',
                'RoleManagement.Read.Directory',
                'Sites.ReadWrite.All',
                'Mail.ReadBasic.All'
            )
        },
        @{
            displayNameOfTargetServicePrincipal = 'Office 365 SharePoint Online';
            namesOfAppRoles = @(
                'Sites.FullControl.All',
                'TermStore.ReadWrite.All',
                'User.ReadWriteAll'
            )
        }
    )
    
        
    if($false){ # how to construct $roleSpecifications programmatically, if needed:
        #=======================================
        # we can look up the proper/allowed $roleSpecifications by doing the following in a tenant that is already properly set up.
        # this assumes that $servicePrincipal is the service principal for the app that we have created.

        $targetServicePrincipals = `
            Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId | 
            select -Unique ResourceId |
            foreach-object { (Get-AzureADObjectByObjectId -ObjectIds @($_.ResourceId  )) }

        $roleSpecifications = @()

        foreach ($targetServicePrincipal in $targetServicePrincipals){
            $roleSpecification = @{
                displayNameOfTargetServicePrincipal = $targetServicePrincipal.DisplayName;
                namesOfAppRoles = @()
            }

            
            $appRoleIds = Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId |
            where {$_.ResourceId -eq $targetServicePrincipal.ObjectId} |
            foreach-object {$_.Id}
            
            $appRoles = $targetServicePrincipal.AppRoles | where { $_.Id -in $appRoleIds }

            $appRoles | foreach-object {"`t" + $_.Value}

            $roleSpecification.namesOfAppRoles = $appRoles | foreach-object {$_.Value}
            
            # $roleSpecifications.Add ($roleSpecification)
            $roleSpecifications += $roleSpecification
            
        }




        $roleSpecificationsExpression = `
            "`$roleSpecifications = @(`n" `
            + (
                (
                    &{
                        foreach ($roleSpecification in $roleSpecifications){      
                            "`t"*1 + "@{`n" `
                            + "`t"*2 +   "displayNameOfTargetServicePrincipal" + " = " + "'" + $roleSpecification.displayNameOfTargetServicePrincipal + "'" + ";" + "`n" `
                            + "`t"*2 +   "namesOfAppRoles" + " = " + "@(" + "`n" `
                            + (($roleSpecification.namesOfAppRoles | foreach-object {"`t"*3 + "'" + $_ + "'"}) -Join ",`n") + "`n" `
                            + "`t"*2 +   ")" + "`n" `
                            + "`t"*1 + "}"
                        } 
                    }
                ) -Join ",`n"
            )  `
            + "`n" + ")`n"

        $roleSpecificationsExpression | Write-Output
        
        
        #=== LEFTOVERS: 
        # $targetServicePrincipal = Get-AzureADServicePrincipal -Filter "DisplayName eq 'Microsoft Graph'"
        # $namesOfAllAvailableAppPermissions = $targetServicePrincipal.AppRoles | foreach-object {$_.Value}

        # #or, while working in some other tenant that is set up properly, by doing
        # $namesOfAppPermissionsThatWeWant = `
            # Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId |
            # where {$_.ResourceId -eq $targetServicePrincipal.ObjectId} -PipelineVariable roleAssignment |
            # foreach-object { ($targetServicePrincipal.AppRoles | where {$_.Id -eq $roleAssignment.Id}).Value }
        
        # $namesOfTargetServicePrincipals =  `
            # Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId | 
            # select -Unique ResourceId |
            # foreach-object { (Get-AzureADObjectByObjectId -ObjectIds @($_.ResourceId  )).DisplayName}
    }


    
    
}

#attempt to read configuration from the configuration file
try {
    $configuration = Get-Content -Raw $pathOfTheConfigurationFile | ConvertFrom-JSON
} catch {
    Write-Output "Failed to read configuration parameters from the configuration file."
    Remove-Variable configuration -ErrorAction SilentlyContinue
}

if(! $configuration){
    Write-Output "Constructing fresh configuration."
        
    .{Function GrantAllThePermissionsWeWant
        # thanks to https://stackoverflow.com/questions/61457429/how-to-add-api-permissions-to-an-azure-app-registration-using-powershell
        {
            param
            (
                [string] $targetServicePrincipalName,
                $appPermissionsRequired,
                $childApp,
                $spForApp
            )

            $targetSp = Get-AzureADServicePrincipal -Filter "DisplayName eq '$($targetServicePrincipalName)'"

            # Iterate Permissions array
            Write-Output -InputObject ('Retrieve Role Assignments objects')
            $RoleAssignments = @()
            Foreach ($AppPermission in $appPermissionsRequired) {
                $RoleAssignment = $targetSp.AppRoles | Where-Object { $_.Value -eq $AppPermission}
                $RoleAssignments += $RoleAssignment
            }

            $ResourceAccessObjects = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]'
            foreach ($RoleAssignment in $RoleAssignments) {
                $resourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess"
                $resourceAccess.Id = $RoleAssignment.Id
                $resourceAccess.Type = 'Role'
                $ResourceAccessObjects.Add($resourceAccess)
            }
            $requiredResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
            $requiredResourceAccess.ResourceAppId = $targetSp.AppId
            $requiredResourceAccess.ResourceAccess = $ResourceAccessObjects

            # set the required resource access
            #actually, we want to append to the app's RequiredResourceAccessList, not overwrite it.
            $initialRequiredResourceAccessList = (Get-AzureADObjectByObjectId -ObjectId $application.ObjectId).RequiredResourceAccess
            $newRequiredResourceAccessList = $initialRequiredResourceAccessList + $requiredResourceAccess
            
            Set-AzureADApplication -ObjectId $childApp.ObjectId -RequiredResourceAccess $newRequiredResourceAccessList
            Start-Sleep -s 1

            # grant the required resource access
            foreach ($RoleAssignment in $RoleAssignments) {
                Write-Output -InputObject ('Granting admin consent for App Role: {0}' -f $($RoleAssignment.Value))
                New-AzureADServiceAppRoleAssignment -ObjectId $spForApp.ObjectId -Id $RoleAssignment.Id -PrincipalId $spForApp.ObjectId -ResourceId $targetSp.ObjectId
                # Start-Sleep -s 1
            }
            
            #TO-do: see if we can get rid of, or at least reduce, the above sleeps.
        }
    }

    Connect-AzureAD
    #following along with instructions at: https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps

    # Create the self signed cert
    
    # construct (or load existing from file) a $certificate, and ensure that the $certificate is installed in the $certificateStorageLocation for later use.
    $certificate = $null
    
    # $pathOfPfxFile = (Join-Path $PSScriptRoot "certificate.pfx")
    # $passwordOfthePfxFile = ""
    
    if($pathOfPfxFile){
        $securePassword =  $( if( $passwordOfthePfxFile ) {(ConvertTo-SecureString -String $passwordOfthePfxFile -AsPlainText -Force)} else {(New-Object System.Security.SecureString)}  )
        try {
            $certificate = Import-PfxCertificate -FilePath $pathOfPfxFile -Password $securePassword -CertStoreLocation $certificateStorageLocation
        } catch {
            Write-Output "Failed to import the certificate from the certificate file"
            # Remove-Variable certificate -ErrorAction SilentlyContinue
            $certificate = $null
        }
    }
    
    if(!$certificate){
        Write-Output "constructing fresh certificate"
        $currentDate = Get-Date
        $endDate = $currentDate.AddYears(10)
        $notAfter = $endDate.AddYears(10)

        $certificate = (New-SelfSignedCertificate -CertStoreLocation $certificateStorageLocation -DnsName com.foo.bar -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter)
        # Export-PfxCertificate -cert $certificate -Password $securePassword -FilePath $pathOfPfxFile
        # 2021-10-26: I have decided to no longer export the certificate to a file -- it should suffice, and will be more secure, to have $certificateStorageLocation be the only place where the certificate's private key is stored.
    }

    $displayNameOfApplication = (Get-AzureADCurrentSessionInfo).Account.ToString() + "_powershell_management"
    
    # Get the Azure Active Directory Application, creating it if it does not already exist.
    $application = Get-AzureADApplication -SearchString $displayNameOfApplication
    if (! $application) {
        $s = @{
            DisplayName                 = $displayNameOfApplication 
            Homepage                    = "https://localhost" 
            ReplyUrls                   = @("https://localhost") 
            IdentifierUris              = ('https://{0}/{1}' -f ((Get-AzureADTenantDetail).VerifiedDomains)[0].Name, $displayNameOfApplication) 
            # Oauth2AllowImplicitFlow     = True
        }; $application = New-AzureADApplication @s 
    }
    else {
        Write-Output -InputObject ('App Registration {0} already exists' -f $displayNameOfApplication)
    }
    
    # Get the service principal associated with $application, creating it if it does not already exist.
    $servicePrincipal = Get-AzureADServicePrincipal -Filter ("appId eq '" + $application.AppId + "'")
    if(! $servicePrincipal){
        $servicePrincipal = New-AzureADServicePrincipal -AppId $application.AppId
    }  else {
        Write-Output -InputObject ('Service Principal {0} already exists' -f $servicePrincipal)
    }
    
    #ensure that the service principal has global admin permissions to the current tenant
    $globalAdminAzureAdDirectoryRole =  Get-AzureADDirectoryRole | where {$_.DisplayName -eq "Global Administrator"}
    if(!$globalAdminAzureAdDirectoryRole){
        $globalAdminAzureAdDirectoryRole =  Get-AzureADDirectoryRole | where {$_.DisplayName -eq "Company Administrator"}
        # for reasons unknown, in some tenants, the displayname of the global admin role is "Company Administrator"
    }
    $azureADDirectoryRoleMember = Get-AzureADDirectoryRoleMember -ObjectId $globalAdminAzureAdDirectoryRole.ObjectId | where {$_.ObjectId -eq $servicePrincipal.ObjectId}
    # iff. $servicePrincipal has the global admin permission, then $azureADDirectoryRoleMember will be $servicePrincipal, otherwise will be null
    if(! $azureADDirectoryRoleMember ){
        Add-AzureADDirectoryRoleMember -ObjectId $globalAdminAzureAdDirectoryRole.ObjectId -RefObjectId $servicePrincipal.ObjectId 
    } else {
        Write-Output -InputObject ('the service principal already has global admin permissions.')
    }
    # we could have probably gotten away simply wrapping Add-AzureADDirectoryRoleMember in a try/catch statement.
    
    #ensure that our public key is installed in our application
    $keyCredential = Get-AzureADApplicationKeyCredential -ObjectId $application.ObjectId | where { ($_.ToJson() | ConvertFrom-JSON).customKeyIdentifier -eq $certificate.Thumbprint }
    if(!$keyCredential){
        $keyCredential = New-AzureADApplicationKeyCredential -ObjectId $application.ObjectId -StartDate $currentDate -EndDate $endDate -Type AsymmetricX509Cert -Usage Verify -Value ([System.Convert]::ToBase64String($certificate.GetRawCertData()))
    } else {
        Write-Output -InputObject ('keyCredential {0} already exists' -f $keyCredential)
    }
    
    #grant all the required approles (as defined by $roleSpecifications) to our app's service principal
    foreach ( $roleSpecification in $roleSpecifications){
        GrantAllThePermissionsWeWant `
            -childApp $application `
            -spForApp $servicePrincipal `
            -targetServicePrincipalName $roleSpecification.displayNameOfTargetServicePrincipal `
            -appPermissionsRequired $roleSpecification.namesOfAppRoles
    }

    $configuration = @{
        tenantId = (Get-AzureADTenantDetail).ObjectId;
        applicationAppId = $application.AppId;
        certificateThumbprint = $certificate.Thumbprint;
    } | ConvertTo-JSON | Out-File $pathOfTheConfigurationFile
    
    Disconnect-AzureAD
    
    $configuration = Get-Content -Raw $pathOfTheConfigurationFile | ConvertFrom-JSON
}

#at this point, we expect to have a valid $configuration and can proceed with making the connection:

# to-do: confirm that the certificate specified in the configuration file is accessible from the certificate store.  If not, 
# attempt to load the certificate from the pfx file, if the pfx file exists.

# if($azureConnection.Account -eq $null){
# if(-not (& {
# try{[Microsoft.Open.Azure.AD.CommonLibrary.AzureSession]::AccessTokens}
# catch{ $null}
# })){
if(-not (& {
try{Get-MgOrganization}
catch{ $null}
})){
    
    
    Write-Host "about to do Connect-MgGraph"
    Select-MgProfile -Name Beta
    $s = @{
        ApplicationId           = $configuration.applicationAppId 
        CertificateThumbprint   = $configuration.certificateThumbprint 
        TenantId                = $configuration.tenantId 
    }; Connect-MgGraph @s 
    Write-Host "done"


    # Write-Host "about to do Connect-AzureAD"
    # $s = @{
    #     ApplicationId           = $configuration.applicationAppId 
    #     CertificateThumbprint   = $configuration.certificateThumbprint 
    #     TenantId                = $configuration.tenantId 
    # }; $azureConnection = Connect-AzureAD @s 
    # Write-Host "done"



    #ideally, we should do a separate test for connection for each of the modules (AzureAD, Exchange, and Sharepoint).
    # However, as a hack, I am only looking at the AzureAD module.

    # Install-Module -Name ExchangeOnlineManagement -RequiredVersion 2.0.5 
    # Install-Module -Name ExchangeOnlineManagement -AllowPrerelease -Confirm:$false -Force
    # Install-Module -Name ExchangeOnlineManagement -AllowPrerelease -Confirm:$false -Force -Scope CurrentUser
    Write-Host "about to do Connect-ExchangeOnline"
    $s = @{
        AppID                   = $configuration.applicationAppId  
        CertificateThumbprint   = $configuration.certificateThumbprint 
        Organization            = ((Get-AzureADTenantDetail).VerifiedDomains | where {$_.Initial -eq $true}).Name
        ShowBanner              = $false
    }
    Write-Host "arguments are $($s | out-string)"
    Connect-ExchangeOnline @s
    Write-Host "done"


    # connect to "Security & Compliance PowerShell in a Microsoft 365 organization."
    # Write-Host "about to do Connect-IPPSSession "
    # $s = @{
    #     AppID                   = $configuration.applicationAppId  
    #     CertificateThumbprint   = $configuration.certificateThumbprint 
    #     Organization            = ((Get-AzureADTenantDetail).VerifiedDomains | where {$_.Initial -eq $true}).Name
    # }
    # Write-Host "arguments are $($s | out-string)"
    # Connect-IPPSSession @s
    # Write-Host "done"

    # Connect-IPPSSession does not seem to be working properly with 
    # unattended app-based authentication.  Connect-IPPSSession tends to 
    # launch a username and apssword prompt (and then fails when the oauth redirect url doesn't match).
    # It appears that connect-ipppssession is a wrapper around connect-exchangeonline.  
    # connect-ippssession calls connect-exchangeonline with 
    # a couple of parameters specified:
    # -UseRPSSession:$true
    # -ShowBanner:$false
    # -ConnectionUri 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId' 
    # -AzureADAuthorizationEndpointUri 'https://login.microsoftonline.com/organizations'
    
    Write-Host "about to do our own equivalent of 'Connect-IPPSSession' "
    $s = @{
        AppID                               = $configuration.applicationAppId  
        CertificateThumbprint               = $configuration.certificateThumbprint 
        Organization                        = ((Get-AzureADTenantDetail).VerifiedDomains | where {$_.Initial -eq $true}).Name
        UseRPSSession                       = $true
        ShowBanner                          = $false
        ConnectionUri                       = 'https://ps.compliance.protection.outlook.com/PowerShell-LiveId' 
        AzureADAuthorizationEndpointUri     = 'https://login.microsoftonline.com/organizations'
    }
    Write-Host "arguments are $($s | out-string)"
    Connect-ExchangeOnline @s
    Write-Host "done"







    $sharepointServiceUrl="https://" +  (((Get-AzureAdDomain | where-object {$_.IsInitial}).Name) -Split '\.')[0] + "-admin.sharepoint.com"

    # $s=@{
    #     Url=$sharepointServiceUrl
    #     # Credential=
    # }; Connect-SPOService @s

    # Connect-PnPOnline `
        # -ClientId $configuration.applicationAppId  `
        # -Tenant (Get-AzureAdDomain | where-object {$_.IsInitial}).Name `
        # -Thumbprint $configuration.certificateThumbprint 
        
    # Install-Module -Name "PnP.PowerShell"   
    Write-Host "about to do Connect-PnPOnline"    
    Connect-PnPOnline `
        -Url ( "https://" +  (((Get-AzureAdDomain | where-object {$_.IsInitial}).Name) -Split '\.')[0] + ".sharepoint.com") `
        -ClientId $configuration.applicationAppId  `
        -Tenant $configuration.tenantId `
        -Thumbprint $configuration.certificateThumbprint 
    Write-Host "done"    
    $application = Get-AzureADApplication -SearchString $application.DisplayName
    
} else {
    Write-Host "It seems that a connection to AzureAD already exists, so we will not bother attempting to reconnect to AzureAD (or ExchangeOnline, or Sharepoint)"
}

# exit     



# [System.Text.Encoding]::ASCII.GetString((Get-AzureADApplicationKeyCredential -ObjectId $application.ObjectId  ).CustomKeyIdentifier)
# Get-AzureADServicePrincipalKeyCredential -ObjectId $servicePrincipal.ObjectId
# # Create the Service Principal and connect it to the Application
# $servicePrincipal = New-AzureADServicePrincipal -AppId $application.AppId



# # Give the Service Principal global admin access to the current tenant (Get-AzureADDirectoryRole)
# Add-AzureADDirectoryRoleMember -ObjectId $globalAdminAzureAdDirectoryRole.ObjectId -RefObjectId $servicePrincipal.ObjectId 

# Remove-AzureADDirectoryRoleMember -ObjectId $globalAdminAzureAdDirectoryRole.ObjectId -MemberId $servicePrincipal.ObjectId

# Get-AzureADApplicationOwner -ObjectId $application.ObjectId

# $result = `
    # $namesOfTargetServicePrincipals -PipelineVariable nameOfTargetServicePrincipal | 
    # foreach-object { 

        
        # @( 
            # $nameOfTargetServicePrincipal , 
            
            
        # ) 
    # }

# $targetServicePrincipal = Get-AzureADServicePrincipal -Filter "DisplayName eq 'Microsoft Graph'"
# # $targetAppRole = $targetServicePrincipal.AppRoles[0]
# $targetAppRole = $targetServicePrincipal.AppRoles | where {$_.Value -eq "Sites.Selected"}


# New-AzureADServiceAppRoleAssignment 
    # -ResourceId # this is the id of the 'resource' (i.e. the service principal for the app whose api we want to access)
    # -Id # this is the id of one of the Microsoft.Open.AzureAD.Model.AppRole objects in the resource's AppRoles property.
    # -PrincipalId # this is the id of the service principal for our app (i.e. the service principal to whom we are granting permissions.)
    # -ObjectId # I don't know what the purpose of this argument is
    
# New-AzureADServiceAppRoleAssignment `
    # -ResourceId $targetServicePrincipal.ObjectId `
    # -Id  $targetAppRole.Id `
    # -PrincipalId  $servicePrincipal.ObjectId `
    # -ObjectId ([Guid]::Empty)
        
# $result = New-AzureADServiceAppRoleAssignment `
    # -ResourceId $targetServicePrincipal.ObjectId `
    # -Id  $targetAppRole.Id `
    # -PrincipalId  $servicePrincipal.ObjectId `
    # -ObjectId $servicePrincipal.ObjectId        

# $requiredResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
# $requiredResourceAccess.ResourceAppId = $targetSp.AppId
# $requiredResourceAccess.ResourceAccess = $ResourceAccessObjects

# # set the required resource access
# Set-AzureADApplication -ObjectId $childApp.ObjectId -RequiredResourceAccess $requiredResourceAccess


# #result is of type Microsoft.Open.AzureAD.Model.AppRoleAssignment, and the newly-created 'role assignment' (aka permission) appears in the 'Other permissions' section (not in the 'configured permissions') of the app's "api permissions' page in the azure ad web interface.    
# # also, the list returned by (Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId) remains empty.
# Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId -All $true

# $roleAssignment = (Get-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId)[0]
# Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.Id  )
# Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.PrincipalId  )
# Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.ResourceId  )
# Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.ObjectId  )

# Get-AzureADObjectByObjectId -ObjectIds @($application.AppId  )

# (Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.ResourceId  )).AppRoles | Where {$_.Id -eq $roleAssignment.Id}
# (Get-AzureADObjectByObjectId -ObjectIds @($roleAssignment.ResourceId  )).AppRoles | Where {$_.Id -eq $roleAssignment.ObjectId}

# #add api permissions:
# # see (https://stackoverflow.com/questions/61457429/how-to-add-api-permissions-to-an-azure-app-registration-using-powershell)

# $appPermissionsRequired = ...

# # Iterate Permissions array
# Write-Output -InputObject ('Retrieve Role Assignments objects')
# $RoleAssignments = @()
# Foreach ($AppPermission in $appPermissionsRequired) {
    # $RoleAssignment = $servicePrincipal.AppRoles | Where-Object { $_.Value -eq $AppPermission}
    # $RoleAssignments += $RoleAssignment
# }

# $ResourceAccessObjects = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.ResourceAccess]'
# foreach ($RoleAssignment in $RoleAssignments) {
    # $resourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess"
    # $resourceAccess.Id = $RoleAssignment.Id
    # $resourceAccess.Type = 'Role'
    # $ResourceAccessObjects.Add($resourceAccess)
# }
# $requiredResourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
# $requiredResourceAccess.ResourceAppId = $servicePrincipal.AppId
# $requiredResourceAccess.ResourceAccess = $ResourceAccessObjects

# $requiredResourceAccessList = New-Object 'System.Collections.Generic.List[Microsoft.Open.AzureAD.Model.RequiredResourceAccess]'

# $requiredResourceAccessList.Add(...)

# # set the required resource access
# $application | Set-AzureADApplication  -RequiredResourceAccess $requiredResourceAccessList
# Start-Sleep -s 1

# # grant the required resource access
# foreach ($RoleAssignment in $RoleAssignments) {
    # Write-Output -InputObject ('Granting admin consent for App Role: {0}' -f $($RoleAssignment.Value))
    # New-AzureADServiceAppRoleAssignment -ObjectId $spForApp.ObjectId -Id $RoleAssignment.Id -PrincipalId $spForApp.ObjectId -ResourceId $servicePrincipal.ObjectId
    # Start-Sleep -s 1
# }


# GrantAllThePermissionsWeWant `
    # -targetServicePrincipalName $targetServicePrincipalName `
    # -appPermissionsRequired $appPermissionsRequired `
    # -childApp $app `
    # -spForApp $spForApp




# # Remove-AzureAdApplication -ObjectId $application.ObjectId
# # Remove-AzureADServicePrincipal -ObjectId $servicePrincipal.ObjectId
# #at this point, the configuration of our app in AzureAd is complete.
# #Collect the configuration details into an object and serialize to a file for future use by the connect_to_office_365.ps1 script

# $configuration = @{
    # tenantId = (Get-AzureADTenantDetail).ObjectId;
    # servicePrincipalId = $servicePrincipal.AppId;
    # pathOfCertificateFile = $pathOfCertificateFile;
    # passwordOfCertificateFile = $passwordOfCertificateFile;
# }





# # Get Tenant Detail
# $tenant=(Get-AzureADTenantDetail).ObjectId
# # Now you can login to Azure PowerShell with your Service Principal and Certificate
# Connect-AzureAD -TenantId $tenant.ObjectId -ApplicationId  $sp.AppId -CertificateThumbprint $thumb



# # $appId = Get-AzureADApplication -SearchString ""
# # $appId = Get-AzureADApplication | Out-String -Stream | Select-String -Pattern "autoscan"

# #Get-AzureADMSApplication

# $autoscanManagementAzureAdApp = Get-AzureADApplication -ObjectId "94bbd8b1-a0e1-468a-aa8c-c0a8e340873f"
# $servicePrincipal = Get-AzureADServicePrincipal -Filter ("appId eq '" + $autoscanManagementAzureAdApp.AppId + "'")
# $azureAdDirectoryRole =  Get-AzureADDirectoryRole | where {$_.DisplayName -eq "Company Administrator"}


# Get-AzureADDirectoryRoleMember -ObjectId $azureAdDirectoryRole.ObjectId

# # New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId    -Id $azureAdDirectoryRole.ObjectId  -PrincipalId <String>  -ResourceId <String>
# # New-AzureADServiceAppRoleAssignment -ObjectId $servicePrincipal.ObjectId   -PrincipalId $servicePrincipal.ObjectId    -Id $azureAdDirectoryRole.ObjectId  

# Add-AzureADDirectoryRoleMember -ObjectId $azureAdDirectoryRole.ObjectId  -RefObjectId $servicePrincipal.ObjectId 

# # Connect-ExchangeOnline -CertificateFilePath "J:\loberg_roofing\powershell management of Office365 for Loberg\mycert.pfx" -CertificatePassword (ConvertTo-SecureString -String "N4M%2ezK9FAkZurF" -AsPlainText -Force) -AppID "94bbd8b1-a0e1-468a-aa8c-c0a8e340873f" -Organization "appriver3651003074.onmicrosoft.com"
# # Connect-ExchangeOnline -CertificateFilePath "J:\loberg_roofing\powershell management of Office365 for Loberg\mycert.pfx" -CertificatePassword (ConvertTo-SecureString -String "N4M%2ezK9FAkZurF" -AsPlainText -Force) -AppID "27b20dbe-43b3-4185-878b-bf564f7e2a21" -Organization "lobergroofing.com"
# # Connect-ExchangeOnline -CertificateFilePath "J:\loberg_roofing\powershell management of Office365 for Loberg\mycert.pfx" -CertificatePassword (ConvertTo-SecureString -String "N4M%2ezK9FAkZurF" -AsPlainText -Force) -AppID "bcd4ec85-1ab0-4228-9078-e9484d23037c" -Organization "lobergroofing.com"
# # Connect-ExchangeOnline -CertificateFilePath "J:\loberg_roofing\powershell management of Office365 for Loberg\mycert.pfx" -CertificatePassword (ConvertTo-SecureString -String "N4M%2ezK9FAkZurF" -AsPlainText -Force) -AppID "bcd4ec85-1ab0-4228-9078-e9484d23037c" -Organization "appriver3651003074.onmicrosoft.com"



# $tenantId = "f3f4dd6b-4a3c-42b9-b6f9-e959fa1c4c25"
# $applicationClientId = "bcd4ec85-1ab0-4228-9078-e9484d23037c"
# $organization = "appriver3651003074.onmicrosoft.com"
# $pathOfCertificateFile = "J:\loberg_roofing\powershell management of Office365 for Loberg\mycert.pfx"
# $passwordOfCertificateFile = "N4M%2ezK9FAkZurF"
# # $clientSecret="FPx12~6GdAiX9xhynY1oWG~R8i_-J-GkqX"
# # $scope = "https://graph.microsoft.com/.default"
# # $grantType = "client_credentials"


# $certificate =  Import-PfxCertificate -CertStoreLocation "cert:\LocalMachine\My" -FilePath $pathOfCertificateFile -Password (ConvertTo-SecureString -String $passwordOfCertificateFile -AsPlainText -Force) 

# # Connect-ExchangeOnline -CertificateFilePath $pathOfCertificateFile -CertificatePassword (ConvertTo-SecureString -String $passwordOfCertificateFile -AsPlainText -Force) -AppID $appId -Organization $organization
# # Connect-ExchangeOnline -AppID $appId -Organization $organization -Certificate $certificate 
# Connect-ExchangeOnline -AppID $applicationClientId -Organization $organization -CertificateThumbprint $certificate.Thumbprint
# # Connect-AzureAD -TenantId $tenantId  -ApplicationId $appId -CertificateFilePath $pathOfCertificateFile -CertificatePassword (ConvertTo-SecureString -String $passwordOfCertificateFile -AsPlainText -Force) 
# Connect-AzureAD -TenantId $tenantId  -ApplicationId $applicationClientId -CertificateThumbprint $certificate.Thumbprint



# $autoscanManagementAzureAdApp = (Get-AzureADApplication -Filter ("AppId eq '" + $applicationClientId + "'"))
# $servicePrincipal = Get-AzureADServicePrincipal -Filter ("appId eq '" + $autoscanManagementAzureAdApp.AppId + "'")


# # $result = New-AzureADApplicationPasswordCredential -ObjectId $applicationClientId
# # $result = New-AzureADApplicationPasswordCredential -ObjectId $servicePrincipal.ObjectId
# # New-AzureADMSApplicationPassword -ObjectId $applicationClientId -PasswordCredential @{ displayname = "mypassword" }
# # New-AzureADMSApplicationPassword -ObjectId $servicePrincipal.ObjectId -PasswordCredential @{ displayname = "mypassword" }
# $passwordCredential = New-AzureADMSApplicationPassword -ObjectId $autoscanManagementAzureAdApp.ObjectId -PasswordCredential @{ displayname = "mypassword" }
# $clientSecret=$passwordCredential.SecretText

# write-host "Sleeping for 4 seconds to allow client secret creation in cloud" -foregroundcolor green
# start-sleep 30

# # Create a hashtable for the body, the data needed for the token request
# # The variables used are explained above
# $Body = @{
    # 'tenant' = $tenantId
    # 'client_id' = $applicationClientId
    # 'scope' = 'https://graph.microsoft.com/.default'
    # 'client_secret' = $clientSecret
    # 'grant_type' = 'client_credentials'
# }

# # Assemble a hashtable for splatting parameters, for readability
# # The tenant id is used in the uri of the request as well as the body
# $Params = @{
    # 'Uri' = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    # 'Method' = 'Post'
    # 'Body' = $Body
    # 'ContentType' = 'application/x-www-form-urlencoded'
# }

# $AuthResponse = Invoke-RestMethod @Params


# $msGraphAccessToken = $AuthResponse.access_token

 


# # Create a hashtable for the body, the data needed for the token request
# # The variables used are explained above
# $Body = @{
    # 'tenant' = $tenantId
    # 'client_id' = $applicationClientId
    # 'scope' = 'https://graph.windows.net/.default'
    # 'client_secret' = $clientSecret
    # 'grant_type' = 'client_credentials'
# }

# # Assemble a hashtable for splatting parameters, for readability
# # The tenant id is used in the uri of the request as well as the body
# $Params = @{
    # 'Uri' = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    # 'Method' = 'Post'
    # 'Body' = $Body
    # 'ContentType' = 'application/x-www-form-urlencoded'
# }

# $AuthResponse = Invoke-RestMethod @Params
# $adGraphAccessToken = $AuthResponse.access_token




# Connect-MsolService -AdGraphAccessToken $adGraphAccessToken -MsGraphAccessToken $msGraphAccessToken
 
# $secureCredential = New-Object System.Management.Automation.PSCredential ($applicationClientId, (ConvertTo-SecureString $clientSecret -AsPlainText -Force))
# Connect-MsolService -Credential $secureCredential 

# $secureCredential = New-Object System.Management.Automation.PSCredential ($servicePrincipal.ObjectId, (ConvertTo-SecureString $clientSecret -AsPlainText -Force))
# Connect-MsolService -Credential $secureCredential

 # Connect-MsolService -AccessToken $adGraphAccessToken 
 # Connect-MsolService -AccessToken $msGraphAccessToken


# Connect-MsolService -AdGraphAccessToken $adGraphAccessToken 
# Connect-MsolService -MsGraphAccessToken $msGraphAccessToken



# Connect-MsolService -AdGraphAccessToken  $msGraphAccessToken -MsGraphAccessToken  $adGraphAccessToken
# Connect-MsolService  -MsGraphAccessToken  $adGraphAccessToken
# Connect-MsolService -AdGraphAccessToken  $msGraphAccessToken 

# Connect-MsolService AdGraphAccessToken  $msGraphAccessToken -MsGraphAccessToken  $msGraphAccessToken
# Connect-MsolService -AdGraphAccessToken  $adGraphAccessToken -MsGraphAccessToken  $adGraphAccessToken


# Set-Clipboard -Value $adGraphAccessToken
# Set-Clipboard -Value $msGraphAccessToken

# # serviceprincipal's objectId is 27b20dbe-43b3-4185-878b-bf564f7e2a21


# # Get-Command Export-PfxCertificate  | fl

# # there are good instructions about how to automate the initial setup of the app permissions and certificate creation at https://docs.microsoft.com/en-us/powershell/module/azuread/connect-azuread?view=azureadps-2.0

# $ApplicationId         = 'xxxx-xxxx-xxxx-xxxx-xxx'
# $ApplicationSecret     = 'YOURSECRET' | Convertto-SecureString -AsPlainText -Force
# $TenantID              = 'xxxxxx-xxxx-xxx-xxxx--xxx' 
# $RefreshToken          = 'LongResourcetoken'
# $ExchangeRefreshToken  = 'LongExchangeToken'
# $credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)



# $aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID 
# $graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID 

# Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken


# Get-AzureADUserOAuth2PermissionGrant $appId

# Get-AzureADServicePrincipalOAuth2PermissionGrant -ObjectId $appId
# Get-AzureADServicePrincipalOAuth2PermissionGrant -ObjectId $servicePrincipal.ObjectId

# Install-Module -Name Microsoft.Graph -Force

# Connect-Graph
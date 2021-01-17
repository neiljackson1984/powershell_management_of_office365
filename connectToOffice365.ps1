Install-Module -Force -Name "CredentialManager"
Import-Module CredentialManager
New-StoredCredential -Comment 'Microsoft Online Administrator credentials for ' -Credentials $(Get-Credential) -Target 'azure ad administrator '






# #  this is what I called to create securestring.txt, which contains the Office365 password
# #  read-host -assecurestring | convertfrom-securestring | out-file securestring.txt




# # create the credential


# $password = cat password-securestring.txt | convertto-securestring
# $username = cat username.txt
# # $O365Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $username,$password

# # Import-Module MSOnline
# # # $O365Cred = Get-Credential
    # # $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
# # Import-PSSession $O365Session -AllowClobber
# # Connect-MsolService -Credential $O365Cred

# Import-Module ExchangeOnlineManagement
# Connect-MsolService  # -Credential $O365Cred
# Connect-ExchangeOnline -UserPrincipalName $username
# # Connect-MsolService 
# # Connect-ExchangeOnline 
@rem this script sets up a powershell session to Office365.

@echo off
set directoryOfThisScript=%~dp0

REM @powershell -NoLogo -NoExit  -File ConnectToOffice365.ps1

REM Import-Module ExchangeOnlineManagement
REM Connect-MsolService  # -Credential $O365Cred
REM Connect-ExchangeOnline -UserPrincipalName $username

powershell -NoLogo -NoExit -Command "Import-Module ExchangeOnlineManagement; Connect-MsolService; Connect-ExchangeOnline"
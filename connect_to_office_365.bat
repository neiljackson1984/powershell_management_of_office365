@rem this batch file is designed to sit in a folder alongside a file config.json, that is a configuration file for the
@rem connect_to_office_365.ps1 powershell script.
@rem this batch file will launch a powershell session and call the connect_to_office365 powershell script, and pass the argument to that script to
@rem to tell that script to use the specified configuration file.

@echo off
set directoryOfThisScript=%~dp0
set pathOfConnectToOffice365PowershellScript=C:\work\powershell_management_of_office365\connect_to_office_365.ps1

powershell -NoLogo -NoExit -File "%pathOfConnectToOffice365PowershellScript%" -pathOfTheConfigurationFile "%directoryOfThisScript%config.json"
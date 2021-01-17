@rem this script sets up a powershell session to Office365.

@echo off
set directoryOfThisScript=%~dp0

powershell -NoLogo -NoExit -File "%directoryOfThisScript%connect_to_office_365.ps1"
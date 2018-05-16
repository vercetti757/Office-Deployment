@Echo Off
TITLE PLEASE DO NOT CLOSE - Office 365 Install
PUSHD %~dp0
Echo "Please do not close this window."
Echo "The install process will take 30-45 minutes to complete."
Echo "You will not be able to use the Office suite during that time."
Echo "After the installation of Office 365 is complete,"
ECHO "a prompt will appear on-screen that Office is ready for use."
Echo.
ECHO "Starting Office 365 Install"
::Show VBS Prompt
cscript.exe .\OffInstallPrompt.vbs
ECHO "Starting Office Scrubbers"
::Scrub Office 2007
Echo "Removing Office 2007"
cscript.exe .\Scrubber\OffScrub07.vbs ALL /log "C:\Logs" /quiet
::Scrub Office 2010
Echo "Removing Office 2010"
cscript.exe .\Scrubber\OffScrub10.vbs ALL /log "C:\Logs" /quiet
::Scrub Office 2013
Echo "Removing Office 2013 (MSI)"
cscript.exe .\Scrubber\OffScrub_O15msi.vbs ALL /log "C:\Logs" /quiet
::Scrub Office 2016
Echo "Removing Office 2016 (MSI)"
cscript.exe .\Scrubber\OffScrub_O16msi.vbs ALL /log "C:\Logs" /quiet
:: Remove Lync 2010
msiexec /x {81BE0B17-563B-45D4-B198-5721E6C665CD} /q /l*v "C:\Logs\Lync2010_Uninstall.log"
:: Remove Office 2007 Folder
rmdir /q /s "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Office 2007\"
:: Install Office 365
cls
Echo "Please do not close this window."
Echo "The install process will take 30-45 minutes to complete."
Echo "You will not be able to use the Office suite during that time."
Echo "After the installation of Office 365 is complete,"
ECHO "a prompt will appear on-screen that Office is ready for use."
Echo.
ECHO "Starting Office 365 Install"
setup.exe /configure Office365_Prod_Full.xml
POPD
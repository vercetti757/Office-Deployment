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
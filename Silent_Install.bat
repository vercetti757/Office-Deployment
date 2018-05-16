@Echo Off
TITLE PLEASE DO NOT CLOSE - Office 365 Install
PUSHD %~dp0
:: Install Office 365
ECHO "Starting Office 365 Install"
setup.exe /configure Office365_Prod_Silent.xml
POPD
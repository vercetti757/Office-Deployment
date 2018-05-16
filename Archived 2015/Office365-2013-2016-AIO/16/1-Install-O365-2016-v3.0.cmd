@echo off
::ADMIN CHECK AUTO ELEVATION
(NET FILE||(powershell -command Start-Process '%0' -Verb runAs -ArgumentList '%* '&EXIT /B))>NUL 2>&1
Color 3
:start
cd /D %0\..\
cls
ECHO ##########################################################################
ECHO #                 Office 365 and 2016 Offline Installer                  #
ECHO #                    Version 3.0 - Updated 09/22/2015                    #
ECHO #                             By AaronWa                                 #
ECHO #                                                                        #
echo #                           -- MAIN MENU --                              #
ECHO #      Please type a number or letter from the menu and hit enter.       #
ECHO ##########################################################################
ECHO:
echo  1 - Stand-Alone Applications (Project, Visio, Etc.)
echo  2 - Office 2016 (Student, Business, Professional, Pro Plus)
echo  3 - Office 365 (Home Premium, University, Small Business, Pro Plus)
echo  4 - Launch Office365/2013 Uninstaller Fix-It(2016 Fixit Release TBD)
echo  5 - Launch Office365/2016 Scrubber (More thorough than Fix-It)
echo  U - Update Office Setup to most curent version
echo  X - Exit Program
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 GOTO standalone
if %OPT%==2 GOTO OFFICE2016RETAIL
if %OPT%==3 GOTO OFFICE365
if %OPT%==4 GOTO O15FIXIT
if %OPT%==5 GOTO OFFICESCRUB
if %OPT%==6 GOTO OFFICE2016RETAIL
if %OPT%==U GOTO UPDATEOFFICE
if %OPT%==u GOTO UPDATEOFFICE
if %OPT%==X GOTO EXIT
if %OPT%==x GOTO EXIT

:standalone
cls
ECHO ##########################################################################
ECHO #     Office 365/2016 Offline Installer - Version 3.0 - By AaronWa       #
ECHO #                                                                        #
ECHO #                -- Stand-Alone Office 32-Bit Products --                #
ECHO ##########################################################################
echo  1 - Install Project 2016 Standard (32-Bit)
echo  2 - Install Project 2016 Professional (32-Bit)
echo  3 - Install Visio 2016 Standard (32-Bit)
echo  4 - Install Visio 2016 Professional (32-Bit)
echo  5 - Install Outlook 2016 (32-Bit)
echo  6 - Install Publisher 2016 (32-Bit)
echo  7 - Install Access 2016 (32-Bit)
echo  8 - Install Word 2016 (32-Bit)
echo  9 - Install Excel 2016 (32-Bit)
echo 10 - Install PowerPoint 2016 (32-Bit)
echo 11 - Install OneNote 2016 (32-Bit)
echo 12 - Install Onenote Free 2016 (32-Bit)
echo 13 - Install Skype for Business Entry(formerly Lync) (32-Bit)
echo 14 - Install Skype for Business(formerly Lync) (32-Bit)
echo 15 - Install SharePoint Designer 2016 (32-Bit)
echo 16 - Install OneDrive for Business 2016 (32-Bit)
echo 64 - 64-Bit Stand-Alone Installers
echo  B - (B)ack to Main Menu
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto 2016ProjectStdRetail32
if %OPT%==2 goto 2016ProjectProRetail32
if %OPT%==3 goto 2016VisioStdRetail32
if %OPT%==4 goto 2016VisioProRetail32
if %OPT%==5 goto 2016OutlookRetail32
if %OPT%==5 goto 2016PublisherRetail32
if %OPT%==7 goto 2016AccessRetail32
if %OPT%==8 goto 2016WordRetail32
if %OPT%==9 goto 2016ExcelRetail32
if %OPT%==10 goto 2016PowerPointRetail32
if %OPT%==11 goto 2016OneNoteRetail32
if %OPT%==12 goto 2016OneNoteFreeRetail32
if %OPT%==13 goto 2016SkypeForBusinessEntryRetail32
if %OPT%==14 goto 2016SkypeForBusinessRetail32
if %OPT%==15 goto 2016SharePointDesigner32
if %OPT%==16 goto GrooveRetail32
if %OPT%==64 goto standalonex64
if %OPT%==b goto Start
if %OPT%==B goto Start

:standalonex64
cls
ECHO ##########################################################################
ECHO #     Office 365/2016 Offline Installer - Version 3.0 - By AaronWa       #
ECHO #                                                                        #
ECHO #                -- Stand-Alone Office 64-Bit Products --                #
ECHO ##########################################################################
echo  1 - Install Project 2016 Standard (64-Bit)
echo  2 - Install Project 2016 Professional (64-Bit)
echo  3 - Install Visio 2016 Standard (64-Bit)
echo  4 - Install Visio 2016 Professional (64-Bit)
echo  5 - Install Outlook 2016 (64-Bit)
echo  6 - Install Publisher 2016 (64-Bit)
echo  7 - Install Access 2016 (64-Bit)
echo  8 - Install Word 2016 (64-Bit)
echo  9 - Install Excel 2016 (64-Bit)
echo 10 - Install PowerPoint 2016 (64-Bit)
echo 11 - Install OneNote 2016 (64-Bit)
echo 12 - Install Onenote Free 2016 (64-Bit)
echo 13 - Install Skype for Business Entry(formerly Lync) (64-Bit)
echo 14 - Install Skype for Business(formerly Lync) (64-Bit)
echo 15 - Install SharePoint Designer 2016 (64-Bit)
echo 16 - Install OneDrive for Business 2016 (64-Bit)
echo 32 - 32-Bit Stand-Alone Installers
echo  B - (B)ack to Main Menu
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto 2016ProjectStdRetail64
if %OPT%==2 goto 2016ProjectProRetail64
if %OPT%==3 goto 2016VisioStdRetail64
if %OPT%==4 goto 2016VisioProRetail64
if %OPT%==5 goto 2016OutlookRetail64
if %OPT%==5 goto 2016PublisherRetail64
if %OPT%==7 goto 2016AccessRetail64
if %OPT%==8 goto 2016WordRetail64
if %OPT%==9 goto 2016ExcelRetail64
if %OPT%==10 goto 2016PowerPointRetail64
if %OPT%==11 goto 2016OneNoteRetail64
if %OPT%==12 goto 2016OneNoteFreeRetail64
if %OPT%==13 goto 2016SkypeForBusinessEntryRetail64
if %OPT%==14 goto 2016SkypeForBusinessRetail64
if %OPT%==15 goto 2016SharePointDesigner64
if %OPT%==16 goto GrooveRetail64
if %OPT%==32 goto Standalone
if %OPT%==b goto Start
if %OPT%==B goto Start

:2016ProjectStdRetail32
cls
cd /D %0\..\
echo Installing Project 2016 Standard (32-Bit)......
start /min setup.exe /configure config\standalone\ProjectStdRetail32.xml
goto exit
 
:2016ProjectProRetail32
cls
cd /D %0\..\
echo Installing Project 2016 Professional (32-Bit)......
start /min setup.exe /configure config\standalone\ProjectProRetail32.xml
goto exit

:2016VisioStdRetail32
cls
cd /D %0\..\
echo Installing Visio 2016 Standard (32-Bit)......
start /min setup.exe /configure config\standalone\VisioStdRetail32.xml
goto exit
 
:2016VisioProRetail32
cls
cd /D %0\..\
echo Installing Visio 2016 Professional (32-Bit)......
start /min setup.exe /configure config\standalone\VisioProRetail32.xml
goto exit

:2016ProjectStdRetail64
cls
cd /D %0\..\
echo Installing Project 2016 Standard (64-Bit)......
start /min setup.exe /configure config\standalone\ProjectStdRetail64.xml
goto exit
 
:2016ProjectProRetail64
cls
cd /D %0\..\
echo Installing Project 2016 Professional (64-Bit)......
start /min setup.exe /configure config\standalone\ProjectProRetail64.xml
goto exit

:2016VisioStdRetail64
cls
cd /D %0\..\
echo Installing Visio 2016 Standard (64-Bit)......
start /min setup.exe /configure config\standalone\VisioStdRetail64.xml
goto exit
 
:2016VisioProRetail64
cls
cd /D %0\..\
echo Installing Visio 2016 Professional (64-Bit)......
start /min setup.exe /configure config\standalone\VisioProRetail64.xml
goto exit

:2016OutlookRetail32
cls
cd /D %0\..\
echo Installing Outlook 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\OutlookRetail32.xml
goto exit

:2016OutlookRetail64
cls
cd /D %0\..\
echo Installing Outlook 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\OutlookRetail64.xml
goto exit

:2016PublisherRetail32
cls
cd /D %0\..\
echo Installing Publisher 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\PublisherRetail32.xml
goto exit

:2016PublisherRetail64
cls
cd /D %0\..\
echo Installing Publisher 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\PublisherRetail64.xml
goto exit

:2016AccessRetail32
cls
cd /D %0\..\
echo Installing Access 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\AccessRetail32.xml
goto exit

:2016AccessRetail64
cls
cd /D %0\..\
echo Installing Access 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\AccessRetail64.xml
goto exit

:2016WordRetail32
cls
cd /D %0\..\
echo Installing Word 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\WordRetail32.xml
goto exit

:2016WordRetail64
cls
cd /D %0\..\
echo Installing Word 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\WordRetail64.xml
goto exit

:2016ExcelRetail32
cls
cd /D %0\..\
echo Installing Excel 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\ExcelRetail32.xml
goto exit

:2016ExcelRetail64
cls
cd /D %0\..\
echo Installing Excel 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\ExcelRetail64.xml
goto exit

:2016PowerPointRetail32
cls
cd /D %0\..\
echo Installing PowerPoint 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\PowerPointRetail32.xml
goto exit

:2016PowerPointRetail64
cls
cd /D %0\..\
echo Installing PowerPoint 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\PowerPointRetail64.xml
goto exit

:2016OneNoteRetail32
cls
cd /D %0\..\
echo Installing OneNote 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\OneNoteRetail32.xml
goto exit

:2016OneNoteRetail64
cls
cd /D %0\..\
echo Installing OneNote 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\OneNoteRetail64.xml
goto exit

:2016OneNoteFreeRetail32
cls
cd /D %0\..\
echo Installing OneNote (Retail Free) 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\OneNoteFreeRetail32.xml
goto exit

:2016OneNoteFreeRetail64
cls
cd /D %0\..\
echo Installing OneNote (Retail Free) 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\OneNoteFreeRetail64.xml
goto exit

:2016SkypeForBusinessEntryRetail32
cls
cd /D %0\..\
echo Installing Skype for Business Retail 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\SkypeForBusinessEntryRetail32.xml
goto exit

:2016SkypeForBusinessEntryRetail64
cls
cd /D %0\..\
echo Installing Skype for Business Retail 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\SkypeForBusinessEntryRetail64.xml
goto exit

:2016SkypeForBusinessRetail32
cls
cd /D %0\..\
echo Installing Skype for Business Retail 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\SkypeForBusinessRetail32.xml
goto exit

:2016SkypeForBusinessRetail64
cls
cd /D %0\..\
echo Installing Skype for Business Retail 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\SkypeForBusinessRetail64.xml
goto exit

:2016SharePointDesigner32
cls
cd /D %0\..\
echo Installing SPD 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\SPDRetail32.xml
goto exit

:2016SharePointDesigner64
cls
cd /D %0\..\
echo Installing SPD 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\SPDRetail64.xml
goto exit

:GrooveRetail32
cls
cd /D %0\..\
echo Installing OneDrive for Business 2016 (32-Bit)......
start /min setup.exe /configure config\standalone\GrooveRetail32.xml
goto exit

:GrooveRetail64
cls
cd /D %0\..\
echo Installing OneDrive for Business 2016 (64-Bit)......
start /min setup.exe /configure config\standalone\GrooveRetail64.xml
goto exit

:OFFICE365
cls
ECHO ##########################################################################
ECHO #               Office 365 and 2016 Offline Installer                    #
ECHO #                       Version 3.0 - By AaronWa                         #
ECHO #                                                                        #
ECHO #              -- Office 365 Subscription 32-Bit Products --             #
ECHO #      Please type a number or letter from the menu and hit enter.       #
ECHO ##########################################################################
echo:
echo  1 - Install Office 365 Home Premium/Personal/University Edition (32-Bit)
echo  2 - Install Office 365 Small Business Premium (32-Bit)
echo  3 - Install Office 365 Professional Plus Retail (32-Bit)
echo 64 - 64-Bit Office 365 Installers
echo  B - (B)ack to Main Menu

echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto O365HomePrem32
if %OPT%==2 goto O365SmallBusPrem32
if %OPT%==3 goto O365ProPlus32
if %OPT%==64 goto OFFICE365x64
if %OPT%==B goto Start
if %OPT%==b goto Start

:OFFICE365X64
cls
ECHO ##########################################################################
ECHO #               Office 365 and 2016 Offline Installer                    #
ECHO #                       Version 3.0 - By AaronWa                         #
ECHO #                                                                        #
ECHO #              -- Office 365 Subscription 64-Bit Products --             #
ECHO #      Please type a number or letter from the menu and hit enter.       #
ECHO ##########################################################################
ECHO:
echo  1 - Install Office 365 Home Premium/Personal/University Edition (64-Bit)
echo  2 - Install Office 365 Small Business Premium (64-Bit)
echo  3 - Install Office 365 Professional Plus Retail (64-Bit)
echo 32 - 32-Bit Office 365 Installers
echo  5 - (B)ack to Main Menu

echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto O365HomePrem64
if %OPT%==2 goto O365SmallBusPrem64
if %OPT%==3 goto O365ProPlus64
if %OPT%==32 goto OFFICE365
if %OPT%==5 goto Start
if %OPT%==b goto Start


:O365HomePrem32
cls
cd /D %0\..\
echo Installing Office 365 Home Premium (32-Bit)......
start /min setup.exe /configure config\O365\O365HomePrem32.xml
goto exit
 
:O365SmallBusPrem32
cls
cd /D %0\..\
echo Installing Office 365 Small Business Premium (32-Bit)......
start /min setup.exe /configure config\O365\O365SmallBusPrem32.xml
goto exit
 
:O365ProPlus32
cls
cd /D %0\..\
echo Installing Office 365 Professional Plus (32-Bit)......
start /min setup.exe /configure config\O365\O365ProPlus32.xml
goto exit

:O365HomePrem64
cls
cd /D %0\..\
echo Installing Office 365 Home Premium (64-Bit)......
start /min setup.exe /configure config\O365\O365HomePrem64.xml
goto exit
 
:O365SmallBusPrem64
cls
cd /D %0\..\
echo Installing Office 365 Small Business Premium (64-Bit)......
start /min setup.exe /configure config\O365\O365SmallBusPrem64.xml
goto exit

:O365ProPlus64
cls
cd /D %0\..\
echo Installing Office 365 Professional Plus (64-Bit)......
start /min setup.exe /configure config\O365\O365ProPlus64.xml
goto exit

:Office2016Retail
cls
ECHO ##########################################################################
ECHO #               Office 365 and 2016 Offline Installer                    #
ECHO #                       Version 3.0 - By AaronWa                         #
ECHO #                                                                        #
ECHO #           -- Office 2016 Retail Perpetual 32-Bit Products --           #
ECHO #       Please type a number or letter from the menu and hit enter.      #
ECHO ##########################################################################
ECHO:
echo  1 - Install Office 2016 Home and Student (32-Bit)
echo  2 - Install Office 2016 Home and Business (32-Bit)
echo  3 - Install Office 2016 Professional (32-Bit)
echo  4 - Install Office 2016 Professional Plus (32-Bit)
echo 64 - 64-Bit Office 2016 Installers
echo  B - (B)ack to Main Menu
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto 2016HomeStudent32
if %OPT%==2 goto 2016HomeBusiness32
if %OPT%==3 goto 2016ProRetail32
if %OPT%==4 goto 2016ProPlusRetail32
if %OPT%==64 goto Office2016Retailx64
if %OPT%==B goto Start
if %OPT%==b goto Start

:Office2016Retailx64
cls
ECHO ##########################################################################
ECHO #               Office 365 and 2016 Offline Installer                    #
ECHO #                       Version 3.0 - By AaronWa                         #
ECHO #                                                                        #
ECHO #           -- Office 2016 Retail Perpetual 64-Bit Products --           #
ECHO #       Please type a number or letter from the menu and hit enter.      #
ECHO ##########################################################################
ECHO:
echo  1 - Install Office 2016 Home and Student (64-Bit)
echo  2 - Install Office 2016 Home and Business (64-Bit)
echo  3 - Install Office 2016 Professional (64-Bit)
echo  4 - Install Office 2016 Professional Plus (64-Bit)
echo 32 - 32-Bit Office 2016 Installers
echo  B - Back to Main Menu
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto 2016HomeStudent64
if %OPT%==2 goto 2016HomeBusiness64
if %OPT%==3 goto 2016ProRetail64
if %OPT%==4 goto 2016ProPlusRetail64
if %OPT%==32 goto Office2016Retail
if %OPT%==B goto Start
if %OPT%==b goto Start


:2016HomeStudent32
cls
cd /D %0\..\
echo Installing Office 365 Home Premium (32-Bit)......
start /min setup.exe /configure config\2016\HomeStudent32.xml
goto exit
 
:2016HomeBusiness32
cls
cd /D %0\..\
echo Installing Office 2016 Home and Business (32-Bit)......
start /min setup.exe /configure config\2016\HomeBusiness32.xml
goto exit

:2016ProRetail32
cls
cd /D %0\..\
echo Installing Office 2016 Professional (32-Bit)......
start /min setup.exe /configure config\2016\ProRetail32.xml
goto exit
 
:2016ProPlusRetail32
cls
cd /D %0\..\
echo Installing Office 2016 Professional Plus (32-Bit)......
start /min setup.exe /configure config\2016\ProPlusRetail32.xml
goto exit

:2016HomeStudent64
cls
cd /D %0\..\
echo Installing Office 2016 Home and Student (64-Bit)......
start /min setup.exe /configure config\2016\HomeStudent64.xml
goto exit
 
:2016HomeBusiness64
cls
cd /D %0\..\
echo Installing Office 2016 Home and Business (64-Bit)......
start /min setup.exe /configure config\2016\HomeBusiness64.xml
goto exit

:2016ProRetail64
cls
cd /D %0\..\
echo Installing Office 2016 Professional (64-Bit)......
start /min setup.exe /configure config\2016\ProRetail64.xml
goto exit
 
:2016ProPlusRetail64
cls
cd /D %0\..\
echo Installing Office 2016 Professional Plus (64-Bit)......
start /min setup.exe /configure config\2016\ProPlusRetail64.xml
goto exit


:O15FIXIT
cls
cd /D %0\..\Uninstall\
echo Launching Office 365/2013 Removal Tool......
start O15CTRRemove.diagcab
goto start

:O16FIXIT
cls
cd /D %0\..\Uninstall\
echo Launching Office 365/2016 Removal Tool......
Echo Coming Soon!
REM start O16CTRRemove.diagcab
goto start

:OFFICESCRUB
cls
ECHO ##########################################################################
ECHO #                                                                        #
ECHO #  Warning, this script will completely remove all traces of Office 2016 #
ECHO #     and 365 C2R products. Please use with caution and try this if the  #
ECHO #              Uninstall Fixit does not work as expected.                #
ECHO #                                                                        #
ECHO ##########################################################################
Pause
cd /D %0\..\Uninstall\Troubleshoot\
echo Calling Office 365/2016 Scrubber Script......
call ForceRemoval.bat
cls
ECHO ##########################################################################
ECHO #                   Office 365/2016 Scrubber Script is complete.         #
ECHO #                          Returning to Main Menu.                       #
ECHO ##########################################################################
timeout 5
goto start

:UPDATEOFFICE
CALL :NETCHECKUPDATE
cls
ECHO ##########################################################################
ECHO #                             INSTRUCTIONS:                              #
ECHO #  To download Office setup files you must be connected to the Internet. #
ECHO #     Downloads will take 10-40 mins depending on your Internet speed.   #
ECHO #      After downloading it is reccomended to install Office while in    #
ECHO #            Airplane Mode. Script will return to the Main Menu          #
ECHO #                       when download is complete.                       #
ECHO #                                                                        #
ECHO #               -- Update Office 365/2016 Setup Files --                 #
ECHO ##########################################################################
echo:
echo  1 - Download Office Setup Files (32-Bit)
echo  2 - Download Office Setup Files (64-Bit)
echo  3 - Download Office Setup Files (32-Bit and 64-Bit)
echo  D - Delete Office Setup Files from USB (Use before updating Setup files
echo  F - Download missing CAB Files only. Fixes some setup errors.
echo  B - (B)ack to Main Menu

echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 goto UPDATEOFFICE32
if %OPT%==2 goto UPDATEOFFICE64
if %OPT%==3 goto UPDATEOFFICEALL
if %OPT%==D goto OFFICEINSTALLCLEANUP
if %OPT%==d goto OFFICEINSTALLCLEANUP
if %OPT%==5 goto Start
if %OPT%==B goto Start
if %OPT%==b goto Start
if %OPT%==F goto downloadCAB32
if %OPT%==f goto downloadCAB32

:UPDATEOFFICE32
cls
ECHO ##########################################################################
ECHO #                               Download Started                         #
ECHO #        Check the CMD Window for Errors. File download size is 1.1GB.   #
ECHO #           Please be patient, the download will take 10-20 mins         #
ECHO #                      depending on your Internet Speed.                 #
ECHO ##########################################################################
ECHO:
cd /D %0\..\
echo Creating Dumming Files before running Setup.
if not exist "Office" md "Office" >nul
if not exist "Office\Data" md "Office\Data" >nul
echo This file will be replaced automatically during first download.>Office\Data\v32.cab
Echo Done.
echo Downloading most recent Office 365/2016 version 32-Bit......
setup.exe /download %0\..\config\Update\Update32.xml
ECHO Download of Office 32-Bit Complete
ECHO All 32-Bit Office files are ready for install. Returning to Main Menu.
Timeout 7
goto start


:UPDATEOFFICE64
cls
ECHO ##########################################################################
ECHO #                               Download Started                         #
ECHO #        Check the CMD Window for Errors. File download size is 1.3GB.   #
ECHO #           Please be patient, the download will take 10-20 mins         #
ECHO #                      depending on your Internet Speed.                 #
ECHO ##########################################################################
ECHO:
cd /D %0\..\
echo Creating Dumming Files before running Setup.
if not exist "Office" md "Office" >nul
if not exist "Office\Data" md "Office\Data" >nul
echo This file will be replaced automatically during first download.>Office\Data\v64.cab
Echo Done.
echo Downloading most recent Office 365/2016 version 64-Bit......
setup.exe /download %0\..\config\Update\Update64.xml
ECHO Download of Office 64-Bit Complete
ECHO All 64-Bit Office files are ready for install. Returning to Main Menu.
timeout 7
goto start

:UPDATEOFFICEALL
cls
ECHO ##########################################################################
ECHO #                               Download Started                         #
ECHO #        Check the CMD Window for Errors. File download size is 2.4GB.   #
ECHO #           Please be patient, the download will take 10-40 mins         #
ECHO #                      depending on your Internet Speed.                 #
ECHO ##########################################################################
ECHO:
echo Downloading most recent Office 365/2016 version 32-Bit......
cd /D %0\..\
echo Creating Dumming Files before running Setup.
if not exist "Office" md "Office" >nul
if not exist "Office\Data" md "Office\Data" >nul
echo This file will be replaced automatically during first download.>Office\Data\v32.cab
echo This file will be replaced automatically during first download.>Office\Data\v64.cab
setup.exe /download %0\..\config\Update\Update32.xml
ECHO Download of Office 32-Bit Complete
echo Downloading Most Recent Office 365/2016 Version 64-Bit......
cd /D %0\..\
setup.exe /download %0\..\config\Update\Update64.xml
ECHO Download of Office 64-Bit Complete
ECHO All Office files downloaded and ready for install. Returning to Main Menu.
timeout 10
goto start


:DownloadCAB32
set "wget=config\wget.exe"
REM SET CURRENT VERSION BASED OFF FOLDER NAME
cd /D %0\..\office\data\
for /f "tokens=*" %%a in ('dir /b /od') do set O16VERS=%%a
echo Current Office Version is "%O16VERS%"
cd /D %0\..\
REM SET OFFICE BIT
set "O16BIT=32"
REM SET OFFICE COUNTRY ID
set "O16CID=1033"
REM SET OFFICE LANGUAGE
set "O16LANG=en-US"
REM SET OFFICE 16 GUID
set "O16GUID=64256AFE-F5D9-4F86-8936-8840A6A4F5BE"
REM SET DOWNLOAD DIRECTORY
set "dldir=%0\..\"
REM SET OFFICE UPDATE LOCATION
	for /f "tokens=3,4 delims=." %%a in ("%O16VERS%") do (set /a offvermajor=%%a && set /a offverminor=%%b)
	set "o16updlocid=64256AFE-F5D9-4F86-8936-8840A6A4F5BE"
	if %offvermajor% EQU 4229 if %offverminor% EQU 1015 set "o16updlocid=A1D255FE-49BB-4A88-ADE8-4EC08D98FBD4"
	set "o16downloadloc=officecdn.microsoft.com/pr/%o16updlocid%/Office/Data"
REM DOWNLOADING HELPER CAB FILES WITH WGET
	cd /d "%dldir%"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/c2rfireflydata.xml
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/c2rfireflydata.xml"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/v32.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v32.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/v32_%O16VERS%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v32_%O16VERS%.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/v64.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v64.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/v64_%O16VERS%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/v64_%O16VERS%.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/fre.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/fre.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/i320.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/i320.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/i32%O16CID%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/i32%O16CID%.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/i640.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/i640.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/i64%O16CID%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/i64%O16CID%.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/s320.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/s320.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/s32%O16CID%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/s32%O16CID%.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/s640.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/s640.cab"
	%wget% --tries=0 --timestamping --force-directories --no-host-directories --cut-dirs=2 http://%o16downloadloc%/%O16VERS%/s64%O16CID%.cab
	if %errorlevel% GEQ 1 call :wgeterror "http://%o16downloadloc%/%O16VERS%/s64%O16CID%.cab"
ECHO All Office CAB files are ready for install. Returning to Main Menu.
Timeout 7
goto start


:OFFICEINSTALLCLEANUP
cls
ECHO ##########################################################################
ECHO #                Deleting Office Setup files in "Office" folder          #
ECHO #               You will need to re-download Setup Files before          #
ECHO #                     you are able to install Office again.              #
ECHO ##########################################################################
Pause
cd /D %0\..\
RMDIR /s /q %0\..\Office\
cls
ECHO ##########################################################################
ECHO #                   Delete of Setup Files is complete.                   #
ECHO #     You must run the Update command from the menu before installing.   #
ECHO #                        Returning to Update Menu.                       #
ECHO ##########################################################################
Timeout 5
goto :UPDATEOFFICE

:NETCHECKINSTALL
REM Network check during installs have been removed
REM It confused the user and generally does not cause any failures without it.
REM Code left here in case of future use.

PING -n 1 www.bing.com|find "Reply from " >NUL
IF NOT ERRORLEVEL 1 goto :NETSUCCESSINSTALL
IF     ERRORLEVEL 1 goto :NETTRYAGAININSTALL

:NETTRYAGAININSTALL
PING -n 3 www.bing.com|find "Reply from " >NUL
IF NOT ERRORLEVEL 1 goto :NETSUCCESSINSTALL
IF     ERRORLEVEL 1 goto :NETFAILUREINSTALL

:NETSUCCESSINSTALL
cls
ECHO ##########################################################################
ECHO #                     Network Connection Detected                        #
ECHO #                                                                        # 
ECHO #           While installing Office it is reccomended to                 #
ECHO #     be in Disconnected from the Internet or in Airplane Mode           #
ECHO #  If not, Setup will automatically connect to Office.com to get the     #
ECHO #           install files instead of using the local files.              #
ECHO #                                                                        #
ECHO ##########################################################################
PAUSE
goto :EOF

:NETFAILUREINSTALL
cls
ECHO ##########################################################################
ECHO #                  You are not connected to the Internet                 #
ECHO #                                                                        #
ECHO #           Congratulations, you are Offline or in Airplane Mode.        #
ECHO #                         Proceeding to Install Menu                     #
ECHO #                                                                        #
ECHO ##########################################################################
timeout 3
goto :EOF

:NETCHECKUPDATE
ECHO Checking Internet connection, please wait...
PING -n 1 www.bing.com|find "Reply from " >NUL
IF NOT ERRORLEVEL 1 goto :NETSUCCESSUPDATE
IF     ERRORLEVEL 1 goto :NETTRYAGAINUPDATE

:NETTRYAGAINUPDATE
PING -n 3 www.bing.com|find "Reply from " >NUL
IF NOT ERRORLEVEL 1 goto :NETSUCCESSUPDATE
IF     ERRORLEVEL 1 goto :NETFAILUREUPDATE

:NETSUCCESSUPDATE
cls
ECHO ##########################################################################
ECHO #                    Internet Connection Detected                        #
ECHO #                                                                        #
ECHO #           Congratulations, you are connected to the Internet.          #
ECHO #                      Proceeding to Update Menu                         #
ECHO #                                                                        #
ECHO ##########################################################################
Timeout 3
goto :EOF

:NETFAILUREUPDATE
cls
ECHO ##########################################################################
ECHO #                You are not connected to the Internet                   #
ECHO #                                                                        # 
ECHO #                  While downloading Office setup files                  #
ECHO #                you must be connected to the Internet.                  #
ECHO #          Please check your Internet connection and try again.          #
ECHO #                                                                        #
ECHO ##########################################################################
pause
goto :EOF

:EXIT
exit

:EOF
@echo off
Color 3
:versionselection
cls
ECHO ##########################################################################
ECHO #               Office 365, 2013, 2016  Offline Installer                #
ECHO #                    Version 3.0 - Updated 09/20/2015                    #
ECHO #                             By AaronWa                                 #
ECHO #                                                                        #
echo #                           -- MAIN MENU --                              #
ECHO #      Please type a number or letter from the menu and hit enter.       #
ECHO ##########################################################################
ECHO:
echo  1 - Office 2013 and 365 (Version 15)
echo  2 - Office 2016 and 365 (Version 16)
echo:
echo  X - Exit Program
echo -----------------------------------------------------
SET /P OPT=Please choose a number or letter, and press enter:
if %OPT%==1 GOTO 2013MENU
if %OPT%==2 GOTO 2016MENU
if %OPT%==X GOTO EXIT
if %OPT%==x GOTO EXIT



:2013MENU
cd /D %0\..\15\
call 1-Install-O365-2013-v3.0.cmd
goto start

:2016MENU
cd /D %0\..\16\
call 1-Install-O365-2016-v3.0.cmd
goto start


:EXIT
exit

:EOF
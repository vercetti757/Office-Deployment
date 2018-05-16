Dim objShell

Set objShell = CreateObject("Wscript.Shell")

returncode = objShell.Popup("Please save all data and close all Office applications (Outlook, Lync, Excel, Word, Internet Explorer, etc.) *** Before clicking OK***" & chr(13) & chr(13) & "Installation will begin when you click OK, close this window, or after 15 minutes." & chr(13) & chr(13) & "Open Office applications WILL CLOSE and unsaved data WILL BE LOST." & chr(13) & chr(13) & "DO NOT log off after install begins.",600,"Microsoft Office 365 Upgrade",4144)

WScript.Quit(returncode)
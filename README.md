# Office-Deployment
Scripts used with SCCM or standalone to install Office 365 and remove previous versions of Microsoft Office.

These are examples of scripts used for deploying Office 365 to workstations. I originally used PowerShell to leverage some nice enhancements, but re-wrote these scripts in .bat and .vbs to maintain compatibility with legacy devices and machines with broken Windows Management Framework. 

In order to use these scripts for installing Office 365 you will need to download the current versions of the Office Deployment Tool from Microsoft. For licensing reasons I am not able to include the setup.exe files directly in this repo. If you are only trying to remove previous versions of Office, no further dowloads are needed.

You can find the Office 2016 Deployment Tool here:
http://www.microsoft.com/en-us/download/details.aspx?id=49117

After downloading, extract the files from the download. Place "setup.exe" in root of the folder above(this Repo).

To download setup files for Office 365, run:
setup.exe /download "Office365_Prod_Full.xml"

This will download source files that match the channel and versions listed in the corresponding XML file.

Thanks to the amazing people over at http://officedev.github.io/Office-IT-Pro-Deployment-Scripts/  They have truly raised the bar and has been a pleasure collobarating with them over the past few years.

I welcome any feedback, and feel free to use/steal/fork/merge this code as needed.

LEGAL DISCLAIMER

* This program is free software. It comes without any warranty, to the extent permitted by applicable law. You can redistribute it and/or modify it under the terms of the Do As You See Fit Public License, Version 4, as published by the author. /*

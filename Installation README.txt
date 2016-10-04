
Installation README

-- This folder contains MSI setup programs for installing the sample projects.

-- The sample projects are written for Visual Studio .NET 1.1.

-- To install a sample project simply open the folder for a chapter and double-click
   on the Setup icon in the folder. This will install the sample project under IIS
   and will set the project's start page (almost always menu.aspx).

-- The first time you build a project you will be prompted to save the solution file.

-- Each sample project must be installed separately. But you will find the installations
   to be very convenient with the MSI.

-- Prior to running a newly installed project you will have to change the SQL connection 
   string in the project's Web.config file to point to your local installed copy of the 
   Northwind database. Currently, the SQL connection strings in the Web.config files contain
   placeholder values for the settings.

-- Finally, the sample projects for Chapters 3 and 6 each require custom stored procedures
   that have been scripted out in 2 .sql script files. You will find these script files 
   located in the "SQL Scripts" folder. To install the scripts, please follow the instructions
   in the "SQL Scripts README" file.
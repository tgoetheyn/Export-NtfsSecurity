# Export-NtfsSecurity

## Descrition
The scripts creates an *.xlsx file with the NTFS security of the given UNC path and it's underlying folder.
It only displays the folders where there security is set and skips folders that only inherit all security from it's parent.

The security is displayed for al the users/groups that is set in a povottable.
The second part of the report shows the members of the groups with som aditional info per user, like description & last logontime
script alse generates a logfile in the outputpath 

## Requirements
This script needs "Quest Active Directory powershell module"
This can be found here: http://www.powershelladmin.com/wiki/Quest_ActiveRoles_Management_Shell_Download

You also need to have access tot the folders you want to analyse, so this script is best ran as a user who is member of the (domain) administrator group

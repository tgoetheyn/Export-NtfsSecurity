# Export-NtfsSecurity

## Descrition
The scripts creates an *.xlsx file with the NTFS security of the given UNC path and it's underlying folder.
It only displays the folders where there security is set and skips folders that only inherit all security from it's parent.

The security is displayed for all the users/groups that is set, in a povottable.
The second part of the report shows the members of the groups with some aditional info per user, like description & last logontime
script also generates a logfile in the outputpath.

## Requirements
This script needs the following powershell modules:

* NTFSSecurity
* ActiveDirectory
* SQLServer

You also need to have access tot the folders you want to analyse, so this script is best ran as a user who is member of the (domain) administrator group.

## Example
create a report of `\\server\data\HR` and its subfolders and save it as `\\server\NTFS-reports\HR-securityreport.xlsx`,
do not include the AD group "SupportGroup" and the AD account "ApplicationAccount".
If the group "domain users" is found, do not display it's members.

`Export-Foldersecurity -unc "\\server\data\HR" -OutputPath "\\server\NTFS-reports\HR-securityreport.xlsx" -ExcludedObjects "SupportGroup","ApplicationAccount" -NoGroupmembershipQuery "Domain users"`
 

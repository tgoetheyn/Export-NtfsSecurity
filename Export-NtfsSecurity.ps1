<# 
    .SYNOPSIS 
    creates NTFS security report as an excell file
 
    .DESCRIPTION 
    the scripts creates an *.xlsx file with the NTFS security of the given UNC path and it's underlying folder.
    it only displays the folders where there security is set and skips folders that only inherit all security from it's parent.

    the security is displayed for al the users/groups that is set in a povottable.
    the second part of the report shows the members of the groups with som aditional info per user, like description & last logontime

    script alse generates a logfile in the outputpath 

    .PARAMETER unc  
    the unc path to analyse.
    if no unc pad given, the script will ask you.
 
    .PARAMETER OutputPath 
    the path and file name of the .xlsx report file.
    if no file given, the script will ask you.

    .PARAMETER Domain 
    the domain you want to query. Default the script uses the domain of the user account used to run the script

    .PARAMETER Domainusers_Only
    only display domain accounts, no local accounts (default = true)
 
    .PARAMETER ExcludedObjects
    array of users and/or groups you want to exclude from the report

    .PARAMETER NoGroupmembershipQuery
    array of groups you want to include in the report, but whose group members you don't want to display
    the groupname is going to be printed in red in the report

    .PARAMETER GetNestedGroupMembers
    retrieve users from nested groups (default = $true)
    only goes 1 level deep, to prevent infinite looping in case of bad groupnesting

    .PARAMETER ShowLogonscripts
    append shown user properties with logonscript

    .PARAMETER CheckOrphanSecurity
    checks for inherited secuity where the parent object is missing, this is where in the advanced security tab you'll see "parent object" instead of a UNC path 
    when found, it's printed in purple in the report
    BEWARE: enabling this makes the script 10 to 20 times slower!!!
    (default = $false)

    .PARAMETER ExportToSQL
    enable export to SQL

    .PARAMETER $DBconnection
    name the SQL xml config file, must be placed in same folder as the script
    fileformat is the following

    <Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
    <Obj RefId="0">
    <TN RefId="0">
      <T>System.Management.Automation.PSCustomObject</T>
      <T>System.Object</T>
    </TN>
    <MS>
      <S N="Server">mysqlserver.mydomain.com</S>
      <S N="Database">databasename</S>
    </MS>
    </Obj>
    </Objs>

    .PARAMETER NoLogfile
    Do not create logfile (default = false)

    .PARAMETER Debug
    specifies scriptpath, for development use only

    .PARAMETER $language
    select report language.
    see "language" folder for available options

    .EXAMPLE
    Export-NtfsSecurity -Domainusers_Only $false

    start the script which asks for an UNC path to analyse and a location to save the report.
    also include local accounts

    .EXAMPLE 
    Export-NtfsSecurity -unc "\\server\data\HR" -OutputPath "\\server\NTFS-reports\HR-securityreport.xlsx" -ExcludedObjects "SupportGroup","ApplicationAccount" -NoGroupmembershipQuery "Domain users"  
 
    create a report of ""\\server\data\HR and is 's subfolders and save it as "\\server\NTFS-reports\HR-securityreport.xlsx"
    do not include the AD group "SupportGroup" and the AD account "ApplicationAccount"
    if the group "domain users" is found, do not display it's members

    .NOTES 
    This script needs "Quest Active Directory powershell module"
    This can be found here: http://www.powershelladmin.com/wiki/Quest_ActiveRoles_Management_Shell_Download

    You also need to have access tot the folders you want to analyse, so this script is best ran as a user who is memeber of the (domain) administrator group
#>

Param(
  [string]$unc = " ",
  [string]$OutputPath = " ",
  [string]$domain = $env:USERDOMAIN,
  [bool]$domainusers_Only = $true,
  [array]$ExcludedObjects = @(), 
  [array]$NoGroupmembershipQuery = @(),
  [bool]$GetNestedGroupMembers = $true,	
  [bool]$ShowLogonscripts = $false,
  [bool]$CheckOrphanSecurity = $false,
  [bool]$ExportToSQL = $true,
  [string]$DBconnection = "NTFS-Database.XML",
  [bool]$NoLogfile = $false,
  [bool]$debug = $false,
  [string]$language = "nl"
)

if ($debug -eq $true){
	$ScriptDir = "\\myserver\myfolder"
}else{
	# set working Dir and fetch it in variable
	Push-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
	$ScriptDir = Split-Path -parent $MyInvocation.MyCommand.Path
}

# import modules and stuff

#Requires -Modules NTFSSecurity, ActiveDirectory, SQLServer
Add-Type -Path "$ScriptDir\EPPlus4.0.4\EPPlus.dll"
$localize = Import-Clixml -Path "$ScriptDir\language\$language.xml"


#################### CONFIG #######################

# replace unc path by driveletter
$UncToDrive = $true
$UncToDrive_UNC = "\\dataserver\share"
$UncToDrive_Drive = "G:"

###################################################


#region Functions
Function Get-FileName($initialDirectory, $filename){  
  [void] [Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' )
  $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.Title = $localize.ExportLocation
  $OpenFileDialog.filter = "xlsx-file (*.xlsx)| *.xlsx"
  $OpenFileDialog.FileName = $filename
  $OpenFileDialog.ShowHelp = $true
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.filename
}
function ToPivotTable($Array) {
  # Function accepts array with 3 coloumns (column, row, value) and converts it to a pivottable
  $coloumnheader = @()
  $coloumnheader = @("")
  foreach ($item in $array){
    $coloumnheader += $item[0]
  }
  $coloumnheader = $coloumnheader | select -Unique | Sort-Object

  $rowheader = @()
  foreach ($item in $array){
    $rowheader += $item[1]
  }
  $rowheader = $rowheader | select -Unique | Sort-Object

  $Pivot = @{}
  $Pivot[0]= $coloumnheader
  $x=1
  foreach ($row in $rowheader){
    $Pivot[$x] = @()
    $Pivot[$x] = @($row)
    $y=0
    foreach ($col in $coloumnheader){
      if($y -eq 0){
        #skip first coloumn
        $y=1
      }else{
        $result = ""
        foreach ($i in $array){
          if (($i[1] -eq $row) -and ($i[0] -eq $col)){
            $result = $i[2]
          } 
        }
        $Pivot[$x] += $result
      }
    }
    $x=$x+1
  }
  return ,$Pivot
}
function SuperSecuritySucker{
  param ($path)
  try{
    $ntfssec = Get-NTFSAccess -Path $path -ErrorAction Stop
  }catch{
    write-log "ERROR: access denied : $path"
  }
  $Object = New-Object PSObject -Property @{
    path = $path
    sec = $ntfssec
  }
  return $Object
}
function Write-Log {
  param([string]$text)
  $logline = "$(Get-Date -Format "yyyy-MM-dd hh:mm:ss"): $text"
  if (!($NoLogfile)){
    $LogFile =  $OutputPath+"logs\"+(($unc -split "\\")[-1])+".log"
    if (!(Test-Path $($OutputPath+"logs\"))){
      $null = mkdir $OutputPath"logs\" -Force
    }
    $logline >> $LogFile
  }
  write-host $logline
}
function short-right {
  param([string]$text)	
  $shortstring = $text -replace("Modify","W")`
             -replace("FullControl","F")`
             -replace("GenericAll","F")`
             -replace("Synchronize","")`
             -replace("ReadAndExecute","R")`
             -replace("GenericExecute","")`
             -replace("GenericRead","R")`
             -replace("GenericWrite","W")`
             -replace("Delete","D")
  if ($shortstring -like "*F*") {$shortstring = "F"}
  if ($shortstring -like "*W*") {$shortstring = "W"}
  $shortstring = $shortstring -replace("^_","")`
                -replace("_$","")`
                -replace("__","_")`
                -replace("_","&")
  return $shortstring
}
function ConvertTo-ExcelCoordinate {
    <#
      .SYNOPSIS
        Convert a row and column to an Excel coordinate
      .DESCRIPTION
        Convert a row and column to an Excel coordinate
      .PARAMETER Row
        Row number
      .PARAMETER Column
        Column number
      .EXAMPLE
        ConvertTo-ExcelCoordinate -Row 1 -Column 2
        #Get Excel coordinates for Row 1, Column 2.  B1.
      .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1
        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/
      .LINK
        https://github.com/RamblingCookieMonster/PSExcel
      .FUNCTIONALITY
        Excel
    #>
    [OutputType([system.string])]
    [cmdletbinding()]
    param(
        [int]$Row,
        [int]$Column
    )

        #From http://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
        Function Get-ExcelColumn
        {
            param([int]$ColumnIndex)

            [string]$Chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

            $ColumnIndex -= 1
            [int]$Quotient = [math]::floor($ColumnIndex / 26)

            if($Quotient -gt 0)
            {
                ( Get-ExcelColumn -ColumnIndex $Quotient ) + $Chars[$ColumnIndex % 26]
            }
            else
            {
                $Chars[$ColumnIndex % 26]
            }
        }

    $ColumnIndex = Get-ExcelColumn $Column
    "$ColumnIndex$Row"
}
Function checkUNC{
  param ($unc)
  if (!(Test-Path $unc)){	
    return $false
  }elseif (!($unc -match "^\\\\\w+\\\w+")){
    return $false
  }else{
    return $true
  }
}

function Cleanup-SQL {
  param ($text)
  $clean = $text -replace ("'","''")
  return $clean
}
#endregion

cls

#Get UNC path
if ($unc -eq " "){
  Do {
    $i = 0
    $unc = read-host $localize.AskUncPath
    if ($unc -eq "exit"){exit}
    if (!(checkUNC $unc)){
      Write-Host """$unc"" $($localize.NotValidTryAgain)"
      $i++
    }
  } While ($i -ne 0)
}else{
  $unc = $unc.trim()
  if (checkUNC $unc){
      $i++
  }else{
    Write-Host """$unc"" $($localize.NotValid)"
    sleep -s 1
    exit
  }
}

#folder and path setup
if ($OutputPath -eq " "){
  $outPutFile = Get-FileName -initialDirectory ($env:USERPROFILE + "\Desktop") -filename (($unc -split "\\")[-1])
  $OutputPath = "$(Split-Path $outPutFile)\"
}else{
  if (!(Test-Path $OutputPath)){
    mkdir $OutputPath -Force
  }
  $outPutFile = $OutputPath +(($unc -split "\\")[-1])+".xlsx"
}
if (Test-Path $outPutFile){
  Remove-Item $outPutFile -Force
}

#clear logfile
if (Test-Path ($OutputPath+"logs\"+(($unc -split "\\")[-1])+".log")){
  rm ($OutputPath+"logs\"+(($unc -split "\\")[-1])+".log")
}

write-log "--------------------------------Starting script------------------------------------------"
write-log "Running script with the following options:"
write-log "OPTION: unc = $unc"
write-log "OPTION: OutputPath = $OutputPath"
write-log "OPTION: domain = $domain"
write-log "OPTION: domainusers_Only = $domainusers_Only"
write-log "OPTION: ExcludedObjects  = $ExcludedObjects "
write-log "OPTION: NoGroupmembershipQuery = $NoGroupmembershipQuery"
write-log "OPTION: GetNestedGroupMembers, = $GetNestedGroupMembers"
write-log "OPTION: ShowLogonscripts = $ShowLogonscripts"
write-log "OPTION: CheckOrphanSecurity = $CheckOrphanSecurity"
write-log "OPTION: ExportToSQL = $ExportToSQL"
write-log "OPTION: DBconnection = $DBconnection"
write-log "OPTION: NoLogfile = $NoLogfile"
write-log "OPTION: debug = $debug"
write-log "-----------------------------------------------------------------------------------------"
Write-Log "Start NTFS analyse of $unc"
Write-Log "Getting folders, this can take a while depending on the amount of folders...."

$startDate = Get-Date
$totalfolders = @()
$analyzefolders = @()
$analyzefolders += $unc

if ($CheckOrphanSecurity){
  Write-Log """CheckOrphanSecurity"" enabled, time to grab a coffee..."
  $analyzefolders += Get-ChildItem $unc -Recurse -Directory -Force -ErrorAction SilentlyContinue | select -ExpandProperty fullname
  $totalfolders = $analyzefolders.count - 1
} else {
  $totalfolders = 0
  #list folders with defined security in SDDL security string (it's super fast)
  Get-ChildItem $unc -Recurse -Directory -Force -ErrorAction SilentlyContinue | % {
    $sec = get-acl $_.fullname | select -ExpandProperty SDDL
    if (($sec -like '*D:P*') -or ($sec -like '*A;OIIO;*') -or ($sec -like '*A;CI;*') -or $sec -like ("*A;OICI;*")){
      $analyzefolders += $_.fullname
    }
    $totalfolders++
  }
}

# analyse folders, get their ACLs, and stuff them into a object and store them in a table ($dirs). 
$dirs = @()
write-log "$totalfolders subfolders found"
Write-Log "Analyzing $($analyzefolders.count) folders"
Write-Log "Start collecting folder security"
$analyzefolders | % {
  if ($_){
    try {
      $dirs += SuperSecuritySucker -path $_
    }catch{
      write-log "ERROR: cannot get security for $($_)"
    }
  }
}
$analyzefolders = $null

#open DB connection (if needed)
if ($ExportToSQL){
  $db = import-Clixml "$ScriptDir\$DBconnection"
  $ServerInstance = $db.server
  $Database = $db.Database
}

# run through the hills!!! i mean, table
$output = @()
$counter2 = 0
$founddirs = @()
if ($ExportToSQL){
  $command = "Delete FROM tblSecurity WHERE (Path LIKE '$(Cleanup-SQL $unc)%')"
  Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $database -Query $command
}
Foreach ($dir in $dirs) { 
  Foreach ($Access in $dir.sec) { 
    $Inherited = [string]$Access.IsInherited 
    $AccessShort = $Access.Account.AccountName.Split("\")[1]
    if ($UncToDrive){
      $shortpath = ($dir.path) -replace [regex]::Escape($UncToDrive_UNC),$UncToDrive_Drive
    }else{
      $shortpath = $dir.path
    }
    #show me al your non-inherited and root security, you dirty NTFS sloth!
    if (($Inherited -eq "False") -or ($dir.path -eq $unc)) {
        $rights = $Access.AccessRights -replace(", ","_")
      #i hate long names... 
        $rights = short-right $rights
      if (($Access.InheritanceFlags -eq "ContainerInherit") -and ($rights -eq "R")){
        $rights = "L"
      } 
      if ($domainusers_Only){
          if ($ExcludedObjects.Count -ne 0){
          if (($Access.Account.AccountName.split("\")[0] -eq $domain) -and ($ExcludedObjects -notcontains $Access.Account.AccountName.split("\")[1])){
              $output += ,@($AccessShort,$shortpath,$rights)
            $founddirs += $dir.path
          }
        }else{
          if ($Access.Account.AccountName.split("\")[0] -eq $domain){
              $output += ,@($AccessShort,$shortpath,$rights)
            $founddirs += $dir.path
          }
        }
      }else{
          $output += ,@($AccessShort,$shortpath,$rights)
        $founddirs += $dir.path
      }
      if ($ExportToSQL){
        # put everything in the database
        $command = "INSERT INTO tblSecurity VALUES ('$($Access.Account.Sid)', '$(Cleanup-SQL $dir.path)', '$($Access.AccessRights)', GETDATE())"
        Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $database -Query $command
      }
      } else {
      if($CheckOrphanSecurity){
        #check for bad parenting en fucked up inheritance
        if ($Access.InheritedFrom.Length -eq 0){  # <== no parent, Q_Q
          $rights = $Access.AccessRights -replace(", ","_")
            $rights = short-right $rights
          if (($Access.InheritanceFlags -eq "ContainerInherit") -and ($rights -eq "R")){
            $rights = "L"
          } 
          # Bad security get marked
          $rights = "-$rights"
          if ($domainusers_Only){
              if ($ExcludedObjects.Count -ne 0){
              if (($Access.Account.AccountName.split("\")[0] -eq $domain) -and ($ExcludedObjects -notcontains $Access.Account.AccountName.split("\")[1])){
                  $output += ,@($AccessShort,$shortpath,$rights)
              }
            }else{
              if ($Access.Account.AccountName.split("\")[0] -eq $domain){
                  $output += ,@($AccessShort,$shortpath,$rights)
              }
            }
          }else{
              $output += ,@($AccessShort,$shortpath,$rights)
          }
        }
      }
    }
  } 
  $counter2++
  Write-Progress -activity "Analyzing $($dirs.Count) Security Entries..." -status "Percent complete: " -PercentComplete (($counter2 / $dirs.Count) * 100)
} 
Write-Progress -Activity "Analyzing $($dirs.Count) Security Entries..." -Completed -Status "All done."
$dirs = $null

#get unique folders for inherit check (bad inheritance is skiped)
Write-Log "Checking folder inheritance"
$founddirs = $founddirs | select -Unique
Foreach ($foundpath in $founddirs) {
  if ($ExportToSQL){
    $command = "Delete FROM tblInheritance WHERE (Path LIKE '$(Cleanup-SQL $foundpath)')"
    Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $database -Query $command
  }
  $inbred = Get-NTFSInheritance $foundpath
  if ($inbred.AccessInheritanceEnabled){
    $TotalInherit = $localize.Yes
    $command = "INSERT INTO tblInheritance VALUES ('$(Cleanup-SQL $foundpath)', '$true')"
  }else{
    $TotalInherit = $localize.No
    $command = "INSERT INTO tblInheritance VALUES ('$(Cleanup-SQL $foundpath)', '$false')"
  }
  if ($UncToDrive){
    $shortpath = $foundpath -replace [regex]::Escape($UncToDrive_UNC),$UncToDrive_Drive
  }else{
    $shortpath = $foundpath
  }
  if ($ExportToSQL){
    $output += ,@($localize.Inherit,$shortpath,$TotalInherit)
    Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $database -Query $command
  }
}

#convert to pivottable an append to file
$result = ToPivotTable($output)

Write-Log "Start collecting groupmembership"

#get Groups
$groups =@()
foreach ($stuff in $output){
  $groups += $stuff[0]
}
$groups = $groups | select -Unique | Sort-Object
$output = $null
Write-Log "AD objects found : $($groups.count)"

#get members of group and add to array
$NoGroupmembershipQuery += $localize.Inherit
$Groupmembers = @()
$SearchNested = @()
Foreach ($gr in $Groups){
  if ($NoGroupmembershipQuery -notcontains $gr){
    try {$adobject = get-adobject -Filter {name -eq $gr -or samaccountname -eq $gr}}
    catch {write-log "ERROR: cannot find $gr"}
    If ($adobject.ObjectClass -eq "group") {
      Get-ADGroupMember $gr | % {
        if ($ExcludedObjects -notcontains $_.Name){ # ignore excluded objects
          $Groupmembers += ,@($gr,$_.Name,"x")
          $SearchNested += $_
        }
      }
      if ((Get-ADGroupMember $gr) -eq $null){$Groupmembers += ,@($gr,"","")}
    }else{
      $Groupmembers += ,@($gr,"","")
    }
  }else{
    $Groupmembers += ,@($gr,"","")
  }
}

#convert to pivottable an append to file
$resultgr = ToPivotTable($Groupmembers)
$Groupmembers = $null


if ($GetNestedGroupMembers){
  Write-Log "Start collecting nested groupmembers"
  $indirectGroups = $SearchNested | ? {$_.type -eq "group"} | select -Unique
}
Write-Log "Nested group(s) found : $($indirectGroups.count)"

#region Create Excel

Write-Log "Generating Excel file"
$newFile = New-Object System.IO.FileInfo($outPutFile)
$package = New-Object OfficeOpenXml.ExcelPackage($newFile)

# set some workbook properties (note there is no need to explicitly create the workbook)
$package.Workbook.Properties.Title = $localize.SecOverview
$package.Workbook.Properties.Author = $localize.Owner
$package.Workbook.Properties.Comments = "$($localize.ReportTitle) $unc"

# create the worksheet we will work on, this contains the spreadsheet cells
$worksheet = $package.Workbook.Worksheets.Add("NTFS Security")


# populate the cells with data
$row = 1
$Column = 1

$BorderColorConverted = [System.Drawing.Color]::Black

#row headers
$AlgGroupDetect = $false
$result[0] | % {
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  if ($_.ToString() -ne $localize.Inherit) {
    $gr = $_.ToString()
    try{
      $tmpobj = get-adobject -Filter {name -eq $gr -or samaccountname -eq $gr}
    }catch{}
    $columnheader = $tmpobj | select -ExpandProperty name
    if ($tmpobj.type -eq "user"){
      $worksheet.Cells[$cell].Style.Font.Color.SetColor([System.Drawing.Color]::blue)
    }
  }else {
    $columnheader = $localize.Inherit
  }
  $worksheet.SetValue($row,$column,$columnheader)
  $worksheet.Cells[$cell].Style.TextRotation = 90
  if (($NoGroupmembershipQuery -contains $_.ToString()) -and ($_.ToString() -ne $localize.Inherit)){
    $worksheet.Cells[$cell].Style.Font.Color.SetColor([System.Drawing.Color]::darkred)
    $AlgGroupDetect = $true
  }
  $worksheet.Column($column).Width = 3
  $Column++
}
$cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
$worksheet.Cells[$cell].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
$worksheet.Cells[$cell].Style.Border.Left.Color.SetColor($BorderColorConverted)

if ($worksheet.Row(1).Height -le 145){
  $worksheet.Row(1).Height = 145
}
$worksheet.Column(1).Width = 65
$worksheet.Column(2).Width = 5

$cell = ConvertTo-ExcelCoordinate -Row 1 -Column 1
$legende= $worksheet.Cells[$cell]
$legende.Style.WrapText = $true
$legende.Style.TextRotation = 0
$legende.Style.VerticalAlignment = "Top"
$legende.IsRichText = $true
$substr = $legende.RichText.Add($localize.LegendTitle)
$substr.Bold = $true
$substr.UnderLine = $true
$substr2 = $legende.RichText.Add($localize.LegendContent)
$substr2.Bold = $false
$substr2.UnderLine = $false

# separator
$Column = 1
for ($index = 1; $index -le $result[0].count; $index++) {
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Bottom.Color.SetColor($BorderColorConverted)
  $Column++
}

# Folder paths with security
for ($index = 1; $index -lt $result.count; $index++) {
  $row++
  $Column = 1
  $result[$index] | % {
    $marker = $_.ToString()
    if ($CheckOrphanSecurity -and $Column -gt 1){
      if ($marker -like "*-*"){
        $badsec = $true
        $marker = $marker -replace "-",""
      }
    }
    $worksheet.SetValue($row,$column,$marker)
    $Column++
  }
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Left.Color.SetColor($BorderColorConverted)
  if ($badsec){
    $worksheet.Cells[$cell].Style.Font.Color.SetColor([System.Drawing.Color]::purple)
  }
}
# separator
$Column = 1
for ($index = 1; $index -le $result[0].count; $index++) {
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Bottom.Color.SetColor($BorderColorConverted)
  $Column++
}

# set description and lastlogon header

$Column = $resultgr[0].count + 1
$values = $localize.Desc,$localize.LastLogon
if ($ShowLogonscripts){
  $values += $localize.Logonscript
}
$values | % {
  $worksheet.SetValue($row,$column,$_)
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Bottom.Color.SetColor($BorderColorConverted)
  $worksheet.Cells[$cell].Style.Font.Bold = $true
  $worksheet.Cells[$cell].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
  $Column++
}

#users and groups (with descrition and lastlogon)
for ($index = 2; $index -lt ($resultgr.count); $index++) {
  $row++
  $Column = 1
  $resultgr[$index] | % {
    if ($Column -eq 1){
      $string = $_.ToString()
      $adobject = $SearchNested | ? {$_.name -eq $string} | select -Unique
      if ($adobject.Type -eq "group") { # Set row Layout for groupnames to Bold
        $worksheet.Row($row).Style.Font.Bold = $true
      }else{
        if ($adobject.AccountIsDisabled){ # Set row Layout for disabled accounts to Gray
          $worksheet.Row($row).Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
        }
      }		
    }
    $worksheet.SetValue($row,$column,$_.ToString())
    $Column++
  }
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Left.Color.SetColor($BorderColorConverted)
	
  if ($adobject.Type -eq "user"){
    if ($adobject.Description){
      $worksheet.SetValue($row,$column,$adobject.Description.ToString()) # Append user description
    }
    $Column++
    try{
      $worksheet.SetValue($row,$column,$adobject.LastLogon.ToShortDateString()) # Append last logondate
      $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
      $worksheet.Cells[$cell].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }catch{}
    if ($ShowLogonscripts){
      $Column++
      if ($adobject.Logonscript){
        $worksheet.SetValue($row,$column,$adobject.Logonscript.ToString()) # Append logonscript
      }
    }
  }
}
$worksheet.Column($resultgr[0].count + 1).AutoFit()
$worksheet.Column($resultgr[0].count + 2).AutoFit()
if ($ShowLogonscripts){
  $worksheet.Column($resultgr[0].count + 3).AutoFit()
}

# separator
$Column = 1
for ($index = 1; $index -le $result[0].count; $index++) {
  $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
  $worksheet.Cells[$cell].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
  $worksheet.Cells[$cell].Style.Border.Bottom.Color.SetColor($BorderColorConverted)
  $Column++
}

#indirect Group memberschip
if ($indirectGroups){
  $row++
  $row++
  $Column = 1
  $worksheet.SetValue($row,$column,$localize.GroupContent)
  $indirectGroups | % {
    $row++
    $Column = 1
    $worksheet.SetValue($row,$column,$_.name) 
    $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
    $worksheet.Cells[$cell].Style.Font.Bold = $true
    if ($NoGroupmembershipQuery -notcontains $_.Name){
      $users = Get-ADGroupMember -Recursive $_ | ? {$ExcludedObjects -notcontains $_.Name} | select -ExpandProperty name | sort
      $users | % {
        $row++
        $Column = 1
        $worksheet.SetValue($row,$column,$_)
      }	
    }else{
      $worksheet.Cells[$cell].Style.Font.Color.SetColor([System.Drawing.Color]::darkred)
      $row++
      $Column = 1
      $worksheet.SetValue($row,$column,$localize.GenericGroup) 
      $cell = ConvertTo-ExcelCoordinate -Row $row -Column $column
      $worksheet.Cells[$cell].Style.Font.Bold = $true
      $AlgGroupDetect = $true
    }
    $row++
  }
}

#add "algemene groep" to legende
if ($AlgGroupDetect){
  $substr3 = $legende.RichText.Add("`n`n$($localize.LegendGeneric)")
  $substr3.Color = [System.Drawing.Color]::DarkRed
}

#footer
$Column = 1
$row++
$row++
$row++
$worksheet.SetValue($row,$column,"$($localize.Location) " + $unc)
$row++
$worksheet.SetValue($row,$column,"$($localize.Date) " + $startDate)
$row++
$duration = New-TimeSpan -Start $startDate -End (get-Date)
$worksheet.SetValue($row,$column,"$($localize.Timespan) $([math]::Floor(($duration | select -ExpandProperty TotalMinutes))) $($localize.Minuts) $($duration | select -ExpandProperty Seconds) $($localize.Seconds)")

#Save the file
Write-Log "Saving Excel file"
$package.Save()

Write-log "total duration $([math]::Floor(($duration | select -ExpandProperty TotalMinutes))) Minuts $($duration | select -ExpandProperty Seconds) Seconds"
Write-Log "Finished"

#endregion
[gc]::collect()
exit 0

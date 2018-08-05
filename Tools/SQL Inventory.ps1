<#=================================
# Generated On: 02/04/2014 
# Original By: Microsoft Gallery
# Changed By : Colin Robinson
# Changes    : Multiple Minor Updates
#=================================
#>
<#=================================
# last update: 10/07/2018 
# Changed By : Lars Platzdasch
# Changes    : Minor Updates
#=================================
#>
[CmdletBinding()]


$Filename='SQLBasedInventory-' + (Get-Date -Format 'yyyy-MM-dd-HH-mm')
$FilePath = 'C:\Inventory\sqlserverlist.csv'
$DirectoryToSaveTo = 'C:\Inventory\'


# before we do anything else, are we likely to be able to save the file?
# if the directory doesn't exist, then create it
if (!(Test-Path -path "$DirectoryToSaveTo")) #create it if not existing
  {
  New-Item "$DirectoryToSaveTo" -type directory | out-null
  }

#Create a new Excel object using COM 
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $True
$Excel = $Excel.Workbooks.Add()
$Excel.Worksheets.Add()

$Sheet1 = $Excel.Worksheets.Item(1)
$Sheet2 = $Excel.Worksheets.Item(2)

#Counter variable for rows
$Sheet1Row = 1
$xlOpenXMLWorkbook=[int]51

#Read thru the contents of the SQL_Servers.txt file
$Sheet1.Cells.Item($Sheet1Row,1)  ="InstanceName"
$Sheet1.Cells.Item($Sheet1Row,2)  ="State"
$Sheet1.Cells.Item($Sheet1Row,3)  ="Support_Team"
$Sheet1.Cells.Item($Sheet1Row,4)  ="ComputerName"
$Sheet1.Cells.Item($Sheet1Row,5)  ="NetName"
$Sheet1.Cells.Item($Sheet1Row,6)  ="OS"
$Sheet1.Cells.Item($Sheet1Row,7)  ="OSVersion"
$Sheet1.Cells.Item($Sheet1Row,8)  ="Platform"
$Sheet1.Cells.Item($Sheet1Row,9)  ="Product"
$Sheet1.Cells.Item($Sheet1Row,10)  ="edition"
$Sheet1.Cells.Item($Sheet1Row,11)  ="Version"
$Sheet1.Cells.Item($Sheet1Row,12)  ="VersionString"
$Sheet1.Cells.Item($Sheet1Row,13) ="ProductLevel"
$Sheet1.Cells.Item($Sheet1Row,14) ="DatabaseCount"
$Sheet1.Cells.Item($Sheet1Row,15) ="HasNullSaPassword"
$Sheet1.Cells.Item($Sheet1Row,16) ="IsCaseSensitive"
$Sheet1.Cells.Item($Sheet1Row,17) ="IsFullTextInstalled"
$Sheet1.Cells.Item($Sheet1Row,18) ="Language"
$Sheet1.Cells.Item($Sheet1Row,19) ="LoginMode"
$Sheet1.Cells.Item($Sheet1Row,20) ="Processors"
$Sheet1.Cells.Item($Sheet1Row,21) ="PhysicalMemory"
$Sheet1.Cells.Item($Sheet1Row,22) ="MaxMemory"
$Sheet1.Cells.Item($Sheet1Row,23) ="MinMemory"
$Sheet1.Cells.Item($Sheet1Row,24) ="IsSingleUser"
$Sheet1.Cells.Item($Sheet1Row,25) ="IsClustered"
$Sheet1.Cells.Item($Sheet1Row,26) ="Collation"
$Sheet1.Cells.Item($Sheet1Row,27) ="MasterDBLogPath"
$Sheet1.Cells.Item($Sheet1Row,28) ="MasterDBPath"
$Sheet1.Cells.Item($Sheet1Row,29) ="ErrorLogPath"
$Sheet1.Cells.Item($Sheet1Row,30) ="BackupDirectory"
$Sheet1.Cells.Item($Sheet1Row,31) ="DefaultLog"
$Sheet1.Cells.Item($Sheet1Row,32) ="ResourceLastUpdatetime"
$Sheet1.Cells.Item($Sheet1Row,33) ="AuditLevel"
$Sheet1.Cells.Item($Sheet1Row,34) ="DefaultFile"
$Sheet1.Cells.Item($Sheet1Row,35) ="xp_cmdshell"
$Sheet1.Cells.Item($Sheet1Row,36) ="Domain"
$Sheet1.Cells.Item($Sheet1Row,37) ="IPAddress"



$Sheet1.Name = "Sql Servers"
  for ($col = 1; $col –le 37; $col++)
     {
          $Sheet1.Cells.Item($Sheet1Row,$col).Font.Bold = $True
          $Sheet1.Cells.Item($Sheet1Row,$col).Interior.ColorIndex = 48
          $Sheet1.Cells.Item($Sheet1Row,$col).Font.ColorIndex = 34
     }

    $Sheet1Row++

#Sheet2
$Sheet2Row = 1
$Sheet2.Name = "DataBases"
$Sheet2.Cells.Item($Sheet2Row,1)  ="Support_Team"
$Sheet2.Cells.Item($Sheet2Row,2)  ="ComputerName"
$Sheet2.Cells.Item($Sheet2Row,3)  ="DataBaseName"
$Sheet2.Cells.Item($Sheet2Row,4)  ="DataBaseSize"
$Sheet2.Cells.Item($Sheet2Row,5)  ="PrimaryDataFileLocation"
$Sheet2.Cells.Item($Sheet2Row,6)  ="LogFileLocation"




 for ($col = 1; $col –le 6; $col++)
     {
          $Sheet2.Cells.Item($Sheet2Row,$col).Font.Bold = $True
          $Sheet2.Cells.Item($Sheet2Row,$col).Interior.ColorIndex = 48
          $Sheet2.Cells.Item($Sheet2Row,$col).Font.ColorIndex = 34
     }




$SQLServerList =   [object]$sqlServerList = Import-CSV $filepath 



foreach ($instanceName in $sqlServerList|WHERE {$_.State -eq 'Active'} )
{
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
$server1 = New-Object -Type Microsoft.SqlServer.Management.Smo.Server -ArgumentList $instanceName.Server
$s=$server1.Information.Properties |Select Name, Value 
$st=$server1.Settings.Properties |Select Name, Value
$CP=$server1.Configuration.Properties |Select DisplayName, Description, RunValue, ConfigValue
$dbs=$server1.Databases.count
$InstanceNameServer = $instanceName.Server
$instanceNameState = $instanceName.State
$instanceNameSupportTeam = $instanceName.SupportTeam
$BuildNumber=$s | where {$_.name -eq "BuildNumber"}|select value
$edition=$s | where {$_.name -eq "edition"}|select value
$ErrorLogPath =$s | where {$_.name -eq "ErrorLogPath"}|select value
$HasNullSaPassword =$s | where {$_.name -eq "HasNullSaPassword"}|select value
$IsCaseSensitive =$s | where {$_.name -eq "IsCaseSensitive"}|select value
$Platform =$s | where {$_.name -eq "Platform"}|select value
$IsFullTextInstalled =$s | where {$_.name -eq "IsFullTextInstalled"}|select value
$Language =$s | where {$_.name -eq "Language"}|select value
$MasterDBLogPath =$s | where {$_.name -eq "MasterDBLogPath"}|select value
$MasterDBPath =$s | where {$_.name -eq "MasterDBPath"}|select value
$NetName =$s | where {$_.name -eq "NetName"}|select value
$OSVersion =$s | where {$_.name -eq "OSVersion"}|select value
$PhysicalMemory =$s | where {$_.name -eq "PhysicalMemory"}|select value
$Processors =$s | where {$_.name -eq "Processors"}|select value
$IsSingleUser =$s | where {$_.name -eq "IsSingleUser"}|select value
$Product =$s | where {$_.name -eq "Product"}|select value
$VersionString =$s | where {$_.name -eq "VersionString"}|select value
$Collation =$s | where {$_.name -eq "Collation"}|select value
$IsClustered =$s | where {$_.name -eq "IsClustered"}|select value
$ProductLevel =$s | where {$_.name -eq "ProductLevel"}|select value
$ComputerNamePhysicalNetBIOS =$s | where {$_.name -eq "ComputerNamePhysicalNetBIOS"}|select value
$ResourceLastUpdateDateTime =$s | where {$_.name -eq "ResourceLastUpdateDateTime"}|select value
$AuditLevel =$st | where {$_.name -eq "AuditLevel"}|select value
$BackupDirectory =$st | where {$_.name -eq "BackupDirectory"}|select value
$DefaultFile =$st | where {$_.name -eq "DefaultFile"}|select value
$DefaultLog =$st | where {$_.name -eq "DefaultLog"}|select value
$LoginMode =$st | where {$_.name -eq "LoginMode"}|select value
$min=$CP | where {$_.Displayname -like "*min server memory*"}|select configValue
$max=$CP | where {$_.Displayname -like "*max server memory*"}|select configValue
$xp_cmdshell=$CP | where {$_.Displayname -like "*xp_cmdshell*"}|select configValue
$FQDN=[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
$IPAddress=(Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $instanceName.Server)|Where IPAddress



if ($HasNullSaPassword.value -eq $NULL)
{
	$HasNullSaPassword.value='No'
}
if($DefaultFile.value -eq '')
{
	$DefaultFile.value='NA'
}
if ($VersionString.value -like '8*')
{
	$SQLServer='SQL SERVER 2000'
}
elseif ($VersionString.value -like '9*')
{
	$SQLServer='SQL SERVER 2005'
}
elseif ($VersionString.value -like '10.0*')
{
	$SQLServer='SQL SERVER 2008'
}
elseif ($VersionString.value -like '10.5*')
{
	$SQLServer='SQL SERVER 2008 R2'
}
elseif ($VersionString.value -like '11*')
{
	$SQLServer='SQL SERVER 2012'
}
elseif ($VersionString.value -like '12*')
{
	$SQLServer='SQL SERVER 2014'
}
elseif ($VersionString.value -like '13*')
{
	$SQLServer='SQL SERVER 2016'
}
elseif ($VersionString.value -like '14*')
{
	$SQLServer='SQL SERVER 2017'
}
else
{
	$SQLServer='Invalid'
}


if ($OSVersion.value -like '5.0*')
{
	$OSVer='Windows 2000'
}
elseif ($OSVersion.value -like '5.1*')
{
	$OSVer='Windows XP'
}
elseif ($OSVersion.value -like '5.2*')
{
	$OSVer='Windows Server 2003'
}
elseif ($OSVersion.value -like '6.0*')
{
	$OSVer='Windows Server 2008'
}
elseif ($OSVersion.value -like '6.1*')
{
	$OSVer='Windows Server 2008 R2'
}
elseif ($OSVersion.value -like '6.2*')
{
	$OSVer='Windows Server 2012'
}
elseif ($OSVersion.value -like '6.3*')
{
	$OSVer='Windows Server 2016'
}
else
{
	$OSVer='NA'
}
	    $Sheet1.Cells.Item($Sheet1Row,1)   =$instanceName
        $Sheet1.Cells.Item($Sheet1Row,2)   =$instanceNameState
        $Sheet1.Cells.Item($Sheet1Row,3)   =$instanceNameSupportTeam
        $Sheet1.Cells.Item($Sheet1Row,4)   =$InstanceNameServer
        $Sheet1.Cells.Item($Sheet1Row,5)   =$NetName.value
        $Sheet1.Cells.Item($Sheet1Row,6)   =$OSVer
        $Sheet1.Cells.Item($Sheet1Row,7)   =$OSVersion.value
        $Sheet1.Cells.Item($Sheet1Row,8)   = $Platform.value
        $Sheet1.Cells.Item($Sheet1Row,9)   = $Product.value
        $Sheet1.Cells.Item($Sheet1Row,10)   = $edition.value
        $Sheet1.Cells.Item($Sheet1Row,11)   = $SQLServer
        $Sheet1.Cells.Item($Sheet1Row,12)  = $VersionString.value
        $Sheet1.Cells.Item($Sheet1Row,13)  = $ProductLevel.value
        $Sheet1.Cells.Item($Sheet1Row,14)  = $Dbs
        $Sheet1.Cells.Item($Sheet1Row,15)  = $HasNullSaPassword.value
        $Sheet1.Cells.Item($Sheet1Row,16)  = $IsCaseSensitive.value
        $Sheet1.Cells.Item($Sheet1Row,17)  = $IsFullTextInstalled.value
        $Sheet1.Cells.Item($Sheet1Row,18)  = $Language.value
        $Sheet1.Cells.Item($Sheet1Row,19)  = $LoginMode.value
        $Sheet1.Cells.Item($Sheet1Row,20)  = $Processors.value
        $Sheet1.Cells.Item($Sheet1Row,21)  = $PhysicalMemory.value
        $Sheet1.Cells.Item($Sheet1Row,22)  = $Max.Configvalue
        $Sheet1.Cells.Item($Sheet1Row,23)  = $Min.Configvalue
        $Sheet1.Cells.Item($Sheet1Row,24)  = $IsSingleUser.value
        $Sheet1.Cells.Item($Sheet1Row,25)  = $IsClustered.value
        $Sheet1.Cells.Item($Sheet1Row,26)  = $Collation.value
        $Sheet1.Cells.Item($Sheet1Row,27)  = $MasterDBLogPath.value
        $Sheet1.Cells.Item($Sheet1Row,28)  = $MasterDBPath.value
        $Sheet1.Cells.Item($Sheet1Row,29)  = $ErrorLogPath.value
        $Sheet1.Cells.Item($Sheet1Row,30)  = $BackupDirectory.value
        $Sheet1.Cells.Item($Sheet1Row,31)  = $DefaultLog.value
        $Sheet1.Cells.Item($Sheet1Row,32)  = $ResourceLastUpdateDateTime.value
        $Sheet1.Cells.Item($Sheet1Row,33)  = $AuditLevel.value
        $Sheet1.Cells.Item($Sheet1Row,34) = $DefaultFile.value
        $Sheet1.Cells.Item($Sheet1Row,35) = $xp_cmdshell.Configvalue
        $Sheet1.Cells.Item($Sheet1Row,36) = $FQDN
        $Sheet1.Cells.Item($Sheet1Row,37) = $IPAddress.IPAddress
$Sheet1Row ++





  

    foreach ($db in $server1.databases)
    {
       IF ($Sheet2Row -eq 1)  #wRITE HEADER ROW
        {
          $icol = 4
        foreach ($Property in $db.properties| where {$_.Name -ne 'ActiveConnections' -and $_.Name -ne 'PolicyHealthState' -and $_.Name -ne 'IsManagementDataWarehouse' -and $_.Name -notlike '*Guid'})
            {
            $Sheet2.Cells.Item(1,$icol)  = $Property.Name
            $Sheet2.Cells.Item(1,$icol).Font.Bold = $True
            $Sheet2.Cells.Item(1,$icol).Interior.ColorIndex = 48
            $Sheet2.Cells.Item(1,$icol).Font.ColorIndex = 34
            $icol++
            }
      
        }
        
         If ($db.name -ne 'Master' -and $db.Name -ne 'Model' -and $db.Name -ne  'tempdb' -and $db.Name -ne 'Msdb' )   #Exclude system databases
            { $Sheet2Row++
            $Sheet2.Cells.Item($Sheet2Row,1)  = $instanceNameSupportTeam #"Support_Team"
            $Sheet2.Cells.Item($Sheet2Row,2)  = $InstanceNameServer  #"ComputerName"
            Write-host $InstanceNameServer +','+ $db.Parent
            $Sheet2.Cells.Item($Sheet2Row,3)  = $db.Name  #"DataBaseName"
    
            $icol = 4
            foreach ($Property in $db.properties| where {$_.Name -ne 'ActiveConnections' -and $_.Name -ne 'PolicyHealthState' -and $_.Name -ne 'IsManagementDataWarehouse' -and $_.Name -notlike '*Guid'})
                { 
                write-host $Property.Name
                $Sheet2.Cells.Item($Sheet2Row,$icol)  = $Property
                $icol++
                }
            }
  
    }
     

    
}
  
$filename = "$DirectoryToSaveTo$filename.xlsx"
if (test-path $filename ) { rm $filename } #delete the file if it already exists
$Sheet1.UsedRange.EntireColumn.AutoFit()
$Sheet1.Name = "Sql Servers"
cls
$Excel.SaveAs($filename, $xlOpenXMLWorkbook) #save as an XML Workbook (xslx)
$Excel.Saved = $True
$Excel.Close()





Function sendEmail([string]$emailFrom, [string]$emailTo, [string]$subject,[string]$body,[string]$smtpServer,[string]$filePath)
{
#initate message
$email = New-Object System.Net.Mail.MailMessage 
$email.From = $emailFrom
$email.To.Add($emailTo)
$email.Subject = $subject
$email.Body = $body
# initiate email attachment 
$emailAttach = New-Object System.Net.Mail.Attachment $filePath
$email.Attachments.Add($emailAttach) 
#initiate sending email 
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($email)
}

#Call Function 
#sendEmail -emailFrom $from -emailTo $to "SQL INVENTORY" "SQL INVENTORY DETAILS - COMPLETE DETAILS" -smtpServer $SMTP -filePath $filename

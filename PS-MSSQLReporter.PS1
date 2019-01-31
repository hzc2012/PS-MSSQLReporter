param (
	[String]$Instance = ".",
	[String]$UserName = $null,
	[String]$Password = $null
)

Import-Module ReportHTML -ErrorAction SilentlyContinue


$ReportTitle = "SQL Server Report"
$ReportSavePath = "C:\Users\"
$ReportName = ($(get-date -format "yyyyMMdd")+"MSSQL_Report")
$RightLogo="https://raw.githubusercontent.com/hzc2012/PS-MSSQLReporter/master/example/PS-MSSQLReporter.png"

(Write-Host -NoNewline "`nInstance: " -ForegroundColor Yellow), (Write-Host $Instance -ForegroundColor White)
(Write-Host -NoNewline "Report Save Path: " -ForegroundColor Yellow), (Write-Host $ReportSavePath -ForegroundColor White)
(Write-Host -NoNewline "Report Title: " -ForegroundColor Yellow), (Write-Host $ReportTitle -ForegroundColor White)


Write-Host "Getting data..." -ForegroundColor Green


# Connect to sql server (Local Windows Authentication )
Function GetSqlConnection{
  if ( [string]::IsNullOrWhiteSpace($UserName) -or [string]::IsNullOrWhiteSpace($Password))
  {
	#Write-Host -NoNewline "With SSPI connection to $Instance`n" -ForegroundColor Yellow
    $ConnectionString = "Data Source=$Instance;Initial Catalog=master;Integrated Security=SSPI;"
  } else {
    #Write-Host -NoNewline "With user $UserName connection to $Instance`n" -ForegroundColor Yellow
    $ConnectionString = "Data Source=$Instance;Initial Catalog=master;user id=$UserName;pwd=$Password"
  }
  try{
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    $SqlConnection.Open()
    return $SqlConnection
  }
  catch{
    return $null
  }
}
 
# Execute query statement
Function ExecuteSQL($sql){
  try{
    $SqlConn = GetSqlConnection
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $sql
    $SqlCmd.Connection = $SqlConn
    $Reader= $SqlCmd.ExecuteReader()
    $DataTable = New-Object System.Data.DataTable
    $DataTable.Load($Reader)
  }
  catch {
    Write-Warning $_
  }
  Finally {
    $SqlConn.close()
  }
  return $DataTable
}

# Check if $x can be converted to numeric.
Function IsNumeric ($x) {
  try {
    0 + $x | Out-Null
    return $true
  } catch {
    return $false
  }
}

# Convert byte to kb, mb, gb, tb
Function ConvertSize ($x){
  if (IsNumeric($x)) {
    if ($x -lt 1000){
	  return "{0} Byte" -f [math]::Round($x ,2)
	} 
	elseif ( $x -ge 1000 -and $x -lt 1000000){
	  return "{0} KB" -f [math]::Round($x / 1KB,2)
	} 
	elseif ( $x -ge 1000000 -and $x -lt 1000000000){
	  return "{0} MB" -f [math]::Round($x / 1MB,2)
	} 
	elseif ( $x -ge 1000000000 -and $x -lt 1000000000000){
	  return "{0} GB" -f [math]::Round($x / 1GB,2)
	} 
	else {
	  return "{0} TB" -f [math]::Round($x / 1TB,2)
	}
  } else {
    return $x
  }
}


<################################ Dashboard ##################################>

<################################ BaseInfo ##################################>

# OS information
$OSInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'

$Win1 = Get-CimInstance Win32_ComputerSystem
$Win2 = Get-CimInstance Win32_Operatingsystem

$OSInfoTable.Add([PSCustomObject]@{'Name' = "Domain";'Value' = $Win1.Domain})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "Name";'Value' = $Win1.Name})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "Manufacturer";'Value' = $Win2.Manufacturer})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "OSName";'Value' = $Win2.Name})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "OSArchitecture";'Value' = $Win2.OSArchitecture})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "NumberOfProcessors";'Value' = $Win1.NumberOfProcessors})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "NumberOfLogicalProcessors";'Value' = $Win1.NumberOfLogicalProcessors})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "TotalVirtualMemorySize";'Value' = ConvertSize($Win2.TotalVirtualMemorySize * 1024)})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "TotalVisibleMemorySize";'Value' = ConvertSize($Win2.TotalVisibleMemorySize * 1024)})
$OSInfoTable.Add([PSCustomObject]@{'Name' = "LastBootUpTime";'Value' = $Win2.LastBootUpTime})


# Logical Disk Space
$DiskObj = Get-WmiObject win32_logicaldisk
$DiskCnt = $DiskObj.Count

$DiskInfoPies = foreach ($i in 0..($DiskCnt-1))
{
	$DiskInfoPie = New-Object 'System.Collections.Generic.List[System.Object]'
	$obj = @()
	$obj += [PSCustomObject]@{
		'Name'  = "Free"
		'Count' = [math]::Round( ($DiskObj[$i].FreeSpace / 1024/ 1024/ 1024),2)
	}
	
	$obj += [PSCustomObject]@{
		'Name'  = "Used"
		'Count' = [math]::Round( (($DiskObj[$i].Size - $DiskObj[$i].FreeSpace)/ 1024/ 1024/ 1024) ,2)
	}
	
	$DiskInfoPie.Add($obj)
	$DiskInfoPie
}

$PieDiskLinks = foreach ($i in 0..($DiskCnt-1))
{
	$PieDiskLink = Get-HTMLPieChartObject
	$PieDiskLink.Title = "{0} (GB)"  -f $DiskObj[$i].DeviceID
	$PieDiskLink.Size.Height = (1200 / $DiskCnt)
	$PieDiskLink.Size.width = (1200 / $DiskCnt)
	$PieDiskLink.ChartStyle.ChartType = 'doughnut'
	$PieDiskLink.ChartStyle.ColorSchemeName = "ColorScheme{0}" -f ($i%4)
	$PieDiskLink.ChartStyle.ColorSchemeName = "Generated{0}" -f ($i%8)
	$PieDiskLink.ChartStyle.ColorSchemeName = 'Random'
	$PieDiskLink.DataDefinition.DataNameColumnName = 'Name'
	$PieDiskLink.DataDefinition.DataValueColumnName = 'Count'
	$PieDiskLink
}


# Logical Disk Space (Failed!!)
$DiskWarnObj = $DiskObj | Where-Object { $_.FreeSpace / $_.Size -lt 0.1}
$DiskWarnCnt = $DiskWarnObj.Count

if ( $DiskWarnCnt -gt 0)
{
$DiskWarnInfoPies = foreach ($i in 0..($DiskWarnCnt-1))
{
	$DiskInfoPie = New-Object 'System.Collections.Generic.List[System.Object]'
	$obj = @()
	$obj += [PSCustomObject]@{
		'Name'  = "Free"
		'Count' = [math]::Round( ($DiskWarnObj[$i].FreeSpace / 1024/ 1024/ 1024),2)
	}
	
	$obj += [PSCustomObject]@{
		'Name'  = "Used"
		'Count' = [math]::Round( (($DiskWarnObj[$i].Size - $DiskWarnObj[$i].FreeSpace)/ 1024/ 1024/ 1024) ,2)
	}
	
	$DiskInfoPie.Add($obj)
	$DiskInfoPie
}

$PieDiskWarnLinks = foreach ($i in 0..($DiskWarnCnt-1))
{
	$PieDiskLink = Get-HTMLPieChartObject
	$PieDiskLink.Title = "{0} (GB)"  -f $DiskWarnObj[$i].DeviceID
	$PieDiskLink.Size.Height = 300
	$PieDiskLink.Size.width = 300
	$PieDiskLink.ChartStyle.ChartType = 'doughnut'
	$PieDiskLink.ChartStyle.ColorSchemeName = "ColorScheme{0}" -f ($i%4)
	$PieDiskLink.ChartStyle.ColorSchemeName = "Generated{0}" -f ($i%8)
	$PieDiskLink.ChartStyle.ColorSchemeName = 'Random'
	$PieDiskLink.DataDefinition.DataNameColumnName = 'Name'
	$PieDiskLink.DataDefinition.DataValueColumnName = 'Count'
	$PieDiskLink
}
}

# SQL Server Product Version
$SQLVersionTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SQLVersions = ExecuteSQL "EXEC master.dbo.xp_msver"

foreach ($SQLVersion in $SQLVersions)
{
	$Index = $SQLVersion.Index
	$Name = $SQLVersion.Name
	$Character_Value = $SQLVersion.Character_Value
	
	$obj = [PSCustomObject]@{
		'Index'          = $Index
		'Name'           = $Name
		'CharacterValue' = $Character_Value
	}
	$SQLVersionTable.Add($obj)
}

if (($SQLVersionTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No sql server were found'
	}
	$SQLVersionTable.Add($obj)
}


# Get SQL Server Services
$SQLServiceTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SQLServices = Get-Service -Name *SQL*

foreach ($SQLService in $SQLServices)
{
	$Status = $SQLService.Status
	$Name = $SQLService.Name
	$DisplayName = $SQLService.DisplayName
	
	$obj = [PSCustomObject]@{
		'Status'      = $Status
		'ServiceName' = $Name
		'DisplayName' = $DisplayName
	}
	$SQLServiceTable.Add($obj)
}

if (($SQLServiceTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No sql service were found'
	}
	$SQLServiceTable.Add($obj)
}

# SQL Server Configurations
$SQLConfTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SQLConfs = ExecuteSQL "SELECT configuration_id,name,value,minimum,maximum,value_in_use,description,is_dynamic,is_advanced FROM master.sys.configurations"

foreach ($SQLConf in $SQLConfs)
{
	$configuration_id = $SQLConf.configuration_id
	$name = $SQLConf.name
	$value = $SQLConf.value
	$minimum = $SQLConf.minimum
	$maximum = $SQLConf.maximum
	$value_in_use = $SQLConf.value_in_use
	$description = $SQLConf.description
	$is_dynamic = $SQLConf.is_dynamic
	$is_advanced = $SQLConf.is_advanced
	
	$obj = [PSCustomObject]@{
		'configuration_id' = $configuration_id
		'name'         = $name
		'value'        = $value
		'minimum'      = $minimum
		'maximum'      = $maximum
		'value_in_use' = $value_in_use
		'description'  = $description
		'is_dynamic'   = $is_dynamic
		'is_advanced'  = $is_advanced
	}
	$SQLConfTable.Add($obj)
}

if (($SQLConfTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No sql service were found'
	}
	$SQLConfTable.Add($obj)
}






<################################ Database ##################################>
# database information
$DatabaseInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DatabaseInfos = ExecuteSQL "select database_id,name,compatibility_level,collation_name,user_access_desc,is_read_only
,state_desc,recovery_model_desc,create_date from sys.databases with(nolock)"

foreach ($DatabaseInfo in $DatabaseInfos)
{
	$database_id = $DatabaseInfo.database_id
	$name = $DatabaseInfo.name
	$compatibility_level = $DatabaseInfo.compatibility_level
	$collation_name = $DatabaseInfo.collation_name
	$user_access_desc = $DatabaseInfo.user_access_desc
	$is_read_only = $DatabaseInfo.is_read_only
	$state_desc = $DatabaseInfo.state_desc
	$recovery_model_desc = $DatabaseInfo.recovery_model_desc
	$create_date = $DatabaseInfo.create_date
	
	$obj = [PSCustomObject]@{
		'DatabaseId'         = $database_id
		'DatabaseName'       = $name
		'CompatibilityLevel' = $compatibility_level
		'CollationName'      = $collation_name
		'UserAccessDesc'     = $user_access_desc
		'IsReadOnly'         = $is_read_only
		'StateDesc'          = $state_desc
		'RecoveryModelDesc'  = $recovery_model_desc
		'CreateDate'         = $create_date
	}
	$DatabaseInfoTable.Add($obj)
}

if (($DatabaseInfoTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No database were found'
	}
	$DatabaseInfoTable.Add($obj)
}


# database file information
$DatabaseFileInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DatabaseFileInfos = ExecuteSQL "select database_id,(select b.name from sys.databases b with(nolock) where a.database_id=b.database_id) as database_name
,type_desc,name,state_desc,size,max_size,growth,physical_name,is_percent_growth from sys.master_files a with(nolock)"

foreach ($DatabaseFileInfo in $DatabaseFileInfos)
{
	$database_id = $DatabaseFileInfo.database_id
	$database_name = $DatabaseFileInfo.database_name
	$name = $DatabaseFileInfo.name
	$type_desc = $DatabaseFileInfo.type_desc
	$state_desc = $DatabaseFileInfo.state_desc
	$size = $DatabaseFileInfo.size
	$max_size = $DatabaseFileInfo.max_size
	$growth = $DatabaseFileInfo.growth
	$physical_name = $DatabaseFileInfo.physical_name
	$is_percent_growth = $DatabaseFileInfo.is_percent_growth
	
	$obj = [PSCustomObject]@{
		'DatabaseId'   = $database_id
		'DatabaseName' = $database_name
		'LogicName'    = $name
		'TypeDesc'     = $type_desc
		'StateDesc'    = $state_desc
		'Size'         = ConvertSize($size*8192)
		'MaxSize'      = if($max_size -eq -1){"unlimited"} else{ConvertSize($max_size*8192)}
		'Growth'       = if($is_percent_growth -eq 1){"{0}%" -f $growth} else{ConvertSize($growth*8192)}
		'PhysicalName' = $physical_name
	}
	$DatabaseFileInfoTable.Add($obj)
}

if (($DatabaseFileInfoTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No database file were found'
	}
	$DatabaseFileInfoTable.Add($obj)
}


# Transaction Log
$SQLTranTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SQLTranLogs = ExecuteSQL "dbcc sqlperf(logspace)"

foreach ($SQLTranLog in $SQLTranLogs)
{
	$DatabaseName = $SQLTranLog.Item(0)
	$LogSize = $SQLTranLog.Item(1)
	$LogSpaceUsed = $SQLTranLog.Item(2)
	$Status = $SQLTranLog.Item(3)
	
	$obj = [PSCustomObject]@{
		'DatabaseName'    = $DatabaseName
		'LogSize'         = ConvertSize($LogSize*1024*1024)
		'LogSpaceUsed(%)' = $LogSpaceUsed
		'Status'          = $Status
	}
	$SQLTranTable.Add($obj)
}

if (($SQLTranTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No database were found'
	}
	$SQLTranTable.Add($obj)
}

# Transaction Log (HTML Bar Chart Object)
# $SQLTranLogs = ExecuteSQL "dbcc sqlperf(logspace)"
$SQLTranLogs = $SQLTranLogs | Select-Object -Property @{N='DatabaseName';E={$_."Database Name"}}, @{N='LogSize';E={$_."Log Size (MB)"}}
$SQLTranLogs = $SQLTranLogs | Sort-Object -Property "LogSize" -Descending | Select-Object -First 10

$SQLTranLogBarChar = Get-HTMLBarChartObject
$SQLTranLogBarChar.Title = " MaxSize TOP 10 (MB)"
$SQLTranLogBarChar.Size.Height = 300
$SQLTranLogBarChar.Size.width = 500
$SQLTranLogBarChar.ChartStyle.LegendPosition = 'none'
$SQLTranLogBarChar.ChartStyle.Showlabels= 'true'
$SQLTranLogBarChar.ChartStyle.BorderWidth = "1"
$SQLTranLogBarChar.ChartStyle.Responsive = 'true'
$SQLTranLogBarChar.ChartStyle.ColorSchemeName = 'Random'
$SQLTranLogBarChar.DataDefinition.DataNameColumnName = 'DatabaseName'
$SQLTranLogBarChar.DataDefinition.DataValueColumnName = 'LogSize'


# SQL Error Log Size
$SQLErrorLogTable = New-Object 'System.Collections.Generic.List[System.Object]'
$SQLErrorLogs = ExecuteSQL "EXEC master.dbo.xp_enumerrorlogs 1"

foreach ($SQLErrorLog in $SQLErrorLogs)
{
	$file_id = $SQLErrorLog.Item(0)
	$date = $SQLErrorLog.Item(1)
	$size = $SQLErrorLog.Item(2)
	
	$obj = [PSCustomObject]@{
		'FileId' = $file_id
		'Date'   = $date
		'Size'   = ConvertSize($size)
	}
	$SQLErrorLogTable.Add($obj)
}

if (($SQLErrorLogTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No file were found'
	}
	$SQLErrorLogTable.Add($obj)
}


# Agent Log Size
$AgentErrorLogTable = New-Object 'System.Collections.Generic.List[System.Object]'
$AgentErrorLogs = ExecuteSQL "EXEC xp_enumerrorlogs 2"

foreach ($AgentErrorLog in $AgentErrorLogs)
{
	$file_id = $AgentErrorLog.Item(0)
	$date = $AgentErrorLog.Item(1)
	$size = $AgentErrorLog.Item(2)
	
	$obj = [PSCustomObject]@{
		'FileId' = $file_id
		'Date'   = $date
		'Size'   = ConvertSize($size)
	}
	$AgentErrorLogTable.Add($obj)
}

if (($AgentErrorLogTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No file were found'
	}
	$AgentErrorLogTable.Add($obj)
}


# Login Name
$LoginNameTable = New-Object 'System.Collections.Generic.List[System.Object]'
$LoginNames = ExecuteSQL "SELECT name,create_date,type_desc,modify_date,default_database_name,is_disabled
FROM master.sys.server_principals WITH(NOLOCK)
WHERE type in ('U', 'G', 'S', 'C', 'K') AND principal_id not between 101 and 255 AND name <> N'##MS_AgentSigningCertificate##'"

foreach ($LoginName in $LoginNames)
{
	$name = $LoginName.name
	$create_date = $LoginName.create_date
	$type_desc = $LoginName.type_desc
	$modify_date = $LoginName.modify_date
	$default_database_name = $LoginName.default_database_name
	$is_disabled = $LoginName.is_disabled
	
	$obj = [PSCustomObject]@{
		'name' = $name
		'TypeDesc'   = $type_desc
		'CreateDate'   = $create_date
		'ModifyDate' = $modify_date
		'DefaultDatabaseName'   = $default_database_name
		'IsDisabled'   = $is_disabled
	}
	$LoginNameTable.Add($obj)
}

if (($LoginNameTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No login name were found'
	}
	$LoginNameTable.Add($obj)
}






<################################ Jobs ##################################>
# Jobs base information
$JobsInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$JobsInfos = ExecuteSQL "select a.name,a.enabled,b.name as owneruser,a.description,a.date_created,a.date_modified,c.next_run_date,c.next_run_time
from msdb.dbo.sysjobs a with(nolock) left join master.sys.syslogins b with(nolock) on a.owner_sid=b.sid left join msdb.dbo.sysjobschedules c with(nolock) on a.job_id=c.job_id"

foreach ($JobsInfo in $JobsInfos)
{
	$name = $JobsInfo.name
	$enabled = $JobsInfo.enabled
	$owneruser = $JobsInfo.owneruser
	$description = $JobsInfo.description
	$date_created = $JobsInfo.date_created
	$date_modified = $JobsInfo.date_modified
	$next_run_date = $JobsInfo.next_run_date
	$next_run_time = $JobsInfo.next_run_time
	
	$obj = [PSCustomObject]@{
		'JobNme'       = $name
		'Enabled'      = $enabled
		'OwnerUser'    = $owneruser
		'Description'  = $description
		'DateCreated'  = $date_created
		'DateModified' = $date_modified
		'NextRunDate'  = $next_run_date
		'NextRunTime'  = $next_run_time
	}
	$JobsInfoTable.Add($obj)
}

if (($JobsInfoTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No jobs were found'
	}
	$JobsInfoTable.Add($obj)
}


# Jobs history
$JobsHistoryTable = New-Object 'System.Collections.Generic.List[System.Object]'
$JobsHisInfos = ExecuteSQL "select a.name,a.enabled,b.server,b.run_date,b.run_time,b.run_duration,b.run_status,b.message
from msdb.dbo.sysjobs a with(nolock) inner join msdb.dbo.sysjobhistory b with(nolock) on a.job_id=b.job_id and b.step_id=0"

foreach ($JobsHisInfo in $JobsHisInfos)
{
	$name = $JobsHisInfo.name
	$enabled = $JobsHisInfo.enabled
	$server = $JobsHisInfo.server
	$run_date = $JobsHisInfo.run_date
	$run_time = $JobsHisInfo.run_time
	$run_duration = $JobsHisInfo.run_duration
	$run_status = $JobsHisInfo.run_status
	$message = $JobsHisInfo.message
	
	$obj = [PSCustomObject]@{
		'JobNme'     = $name
		'Enabled'    = $enabled
		'Server'     = $server
		'RunDate'    = $run_date
		'RunTime'    = $run_time
		'RunSeconds' = $run_duration
		'RunStatus' = $run_status
		'Message'    = $message
	}
	$JobsHistoryTable.Add($obj)
}

if (($JobsHistoryTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No jobs history were found'
	}
	$JobsHistoryTable.Add($obj)
}


<################################ Backup & Restore ##################################>
# Backup Database Info
$BackupInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
$BackupInfos = ExecuteSQL "SELECT K.name
,MAX(CASE WHEN type='D' THEN backup_finish_date ELSE NULL END) AS latest_full_backup_time 
,MAX(CASE WHEN type='I' THEN backup_finish_date ELSE NULL END) AS latest_diff_backup_time 
,MAX(CASE WHEN type='L' THEN backup_finish_date ELSE NULL END) AS latest_log_backup_time 
FROM master.sys.databases K with(nolock)
LEFT JOIN ( 
	SELECT database_name,type,MAX(backup_finish_date) AS backup_finish_date  
	FROM msdb.dbo.backupset with(nolock) GROUP BY database_name,type 
) T ON K.name=T.database_name GROUP BY K.name"

foreach ($BackupInfo in $BackupInfos)
{
	$name = $BackupInfo.name
	$latest_full_backup_time = $BackupInfo.latest_full_backup_time
	$latest_diff_backup_time = $BackupInfo.latest_diff_backup_time
	$latest_log_backup_time  = $BackupInfo.latest_log_backup_time
	
	$obj = [PSCustomObject]@{
		'DatabaseName'         = $name
		'LatestFullBackupTime' = $latest_full_backup_time
		'LatestDiffBackupTime' = $latest_diff_backup_time
		'LatestLogBackupTime'  = $latest_log_backup_time
	}
	$BackupInfoTable.Add($obj)
}

if (($BackupInfoTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No bacbkup history were found'
	}
	$BackupInfoTable.Add($obj)
}

# Backup Database and Not backup's
$BackupInfoPie = New-Object 'System.Collections.Generic.List[System.Object]'

$DBAll = ExecuteSQL "select count(name) cnt from master.sys.databases with(nolock)"
$DBBak = ExecuteSQL "select count(distinct database_name) cnt from msdb.dbo.backupset with(nolock)"

$obj1 = [PSCustomObject]@{
	'Name'  = "never backup"
	'Count' = $DBAll.cnt-$DBBak.cnt
}

$BackupInfoPie.Add($obj1)

$obj2 = [PSCustomObject]@{
	'Name'  = "has backup"
	'Count' = $DBBak.cnt
}

$BackupInfoPie.Add($obj2)

$PieBackupLinks = Get-HTMLPieChartObject
$PieBackupLinks.Title = "Database Backup"
$PieBackupLinks.Size.Height = 300
$PieBackupLinks.Size.width = 400
$PieBackupLinks.ChartStyle.ChartType = 'doughnut'
$PieBackupLinks.ChartStyle.ColorSchemeName = "ColorScheme4"
$PieBackupLinks.ChartStyle.ColorSchemeName = "Generated5"
$PieBackupLinks.ChartStyle.ColorSchemeName = 'Random'
$PieBackupLinks.DataDefinition.DataNameColumnName = 'Name'
$PieBackupLinks.DataDefinition.DataValueColumnName = 'Count'


# The latest Backup Information
$BackupDetailTable = New-Object 'System.Collections.Generic.List[System.Object]'
$BackupDetailInfos = ExecuteSQL "select a.database_name 
,(case when a.type='D' then 'Full' when a.type='I' then 'Diff' when a.type='L' then 'Log' else '' end) as type
,CONVERT(varchar(30),a.backup_start_date,120) AS latest_backup_time ,compressed_backup_size as backup_size
,DATEDIFF(MINUTE,backup_start_date,backup_finish_date) as backup_time_Min  
,REVERSE(stuff(REVERSE(physical_device_name),1,CHARINDEX('\',REVERSE(physical_device_name))-1,'')) as backup_path  
from msdb.dbo.backupset a with(nolock)
inner join (select database_name,type,MAX(backup_start_date) as latest_backup_date from msdb.dbo.backupset with(nolock) group by database_name,type )b 
on a.database_name=b.database_name and a.type=b.type and a.backup_start_date=b.latest_backup_date  
left join msdb.dbo.backupmediafamily c with(nolock) on a.media_set_id=c.media_set_id"

foreach ($BackupDetailInfo in $BackupDetailInfos)
{
	$database_name = $BackupDetailInfo.database_name
	$type = $BackupDetailInfo.type
	$latest_backup_time = $BackupDetailInfo.latest_backup_time
	$backup_size  = $BackupDetailInfo.backup_size
	$backup_time_Min = $BackupDetailInfo.backup_time_Min
	$backup_path  = $BackupDetailInfo.backup_path
	
	$obj = [PSCustomObject]@{
		'DatabaseName'          = $database_name
		'BackupType'            = $type
		'LatestBackupTime'      = $latest_backup_time
		'BackupSize'            = ConvertSize($backup_size)
		'BackupTimeCost_Minute' = $backup_time_Min
		'BackupPath'            = $backup_path
	}
	$BackupDetailTable.Add($obj)
}

if (($BackupDetailTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No bacbkup history were found'
	}
	$BackupDetailTable.Add($obj)
}


# The latest Restore Information
$RestoreTable = New-Object 'System.Collections.Generic.List[System.Object]'
$RestoreInfos = ExecuteSQL "select user_name,destination_database_name
,(case when restore_type='D' then 'Full' when restore_type='I' then 'Diff' when restore_type='L' then 'Log' else '' end) as restore_type
,restore_date,replace,recovery,restart,stop_at,file_number,destination_phys_name
from msdb.dbo.restorehistory a with(nolock) inner join msdb.dbo.restorefile b with(nolock) on a.restore_history_id=b.restore_history_id"

foreach ($RestoreInfo in $RestoreInfos)
{
	$user_name = $RestoreInfo.user_name
	$destination_database_name = $RestoreInfo.destination_database_name
	$restore_type  = $RestoreInfo.restore_type
	$restore_date = $RestoreInfo.restore_date
	$replace  = $RestoreInfo.replace
	$recovery = $RestoreInfo.recovery
	$restart  = $RestoreInfo.restart
	$stop_at = $RestoreInfo.stop_at
	$file_number  = $RestoreInfo.file_number
	$destination_phys_name  = $RestoreInfo.destination_phys_name
	
	$obj = [PSCustomObject]@{
		'UserName'         = $user_name
		'DestDatabaseName' = $destination_database_name
		'RestoreType'      = $restore_type
		'RestoreDate'      = $restore_date
		'Replace'          = $replace
		'Recovery'         = $recovery
		'Restart'          = $restart
		'StopAt'           = $stop_at
		'FileNumber'       = $file_number
		'DestPhysName'     = $destination_phys_name
	}
	$RestoreTable.Add($obj)
}

if (($RestoreTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No restore history were found'
	}
	$RestoreTable.Add($obj)
}





<################################ Mirror ##################################>

<################################ Replication ##################################>

<################################ AlwaysOn-AG ##################################>

<################################ AlwaysOn-FCI ##################################>

<################################ Perfmon ##################################>
# Database buffer pool allocation
$DBBufferpoolTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DBBufferpoolInfos = ExecuteSQL "SELECT CASE database_id WHEN 32767 THEN 'ResourceDb' ELSE db_name(database_id) END AS Database_name  
,count(*) AS cached_pages_count ,count(*)*8192 AS cached_space_in_bytes ,sum(free_space_in_bytes) AS free_space_in_bytes  
FROM sys.dm_os_buffer_descriptors(nolock) GROUP BY db_name(database_id) ,database_id"

foreach ($DBBufferpoolInfo in $DBBufferpoolInfos)
{
	$Database_name = $DBBufferpoolInfo.Database_name
	$cached_pages_count = $DBBufferpoolInfo.cached_pages_count
	$cached_space_in_bytes  = $DBBufferpoolInfo.cached_space_in_bytes
	$free_space_in_bytes = $DBBufferpoolInfo.free_space_in_bytes
	
	$obj = [PSCustomObject]@{
		'DatabaseName'     = $Database_name
		'CachedPagesCount' = $cached_pages_count
		'CachedSpaceSize'  = ConvertSize($cached_space_in_bytes)
		'FreeSpaceSize'    = ConvertSize($free_space_in_bytes)
	}
	$DBBufferpoolTable.Add($obj)
}

if (($DBBufferpoolTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No Cached were found'
	}
	$DBBufferpoolTable.Add($obj)
}

# Database dirty page
$DBDirtyPageTable = New-Object 'System.Collections.Generic.List[System.Object]'
$DBDirtyPages = ExecuteSQL "SELECT db_name(database_id) AS DatabaseName,count(page_id) AS DirtyPages,count(page_id)*8192 AS DirtyPages_bytes
FROM sys.dm_os_buffer_descriptors(nolock) WHERE is_modified = 1 GROUP BY database_id"

foreach ($DirtyPage in $DBDirtyPages)
{
	$DatabaseName = $DirtyPage.DatabaseName
	$DirtyPages = $DirtyPage.DirtyPages
	$DirtyPages_bytes  = $DirtyPage.DirtyPages_bytes
	
	$obj = [PSCustomObject]@{
		'DatabaseName'    = $DatabaseName
		'DirtyPages'      = $DirtyPages
		'DirtyPagesSize'  = ConvertSize($DirtyPages_bytes)
	}
	$DBDirtyPageTable.Add($obj)
}

if (($DBDirtyPageTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No Cached were found'
	}
	$DBDirtyPageTable.Add($obj)
}

# Cached Type Size
$CatchTypeTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CatchType01 = ExecuteSQL "select cacheobjtype as Cached_Type,COUNT(*) Number_of_Plans,SUM(size_in_bytes) as size_in_bytes 
from sys.dm_exec_cached_plans group by cacheobjtype"

foreach ($CatchType1 in $CatchType01)
{
	$Cached_Type = $CatchType1.Cached_Type
	$Number_of_Plans = $CatchType1.Number_of_Plans
	$size_in_bytes  = $CatchType1.size_in_bytes
	
	$obj = [PSCustomObject]@{
		'CachedType'    = $Cached_Type
		'NumberOfPlans' = $Number_of_Plans
		'Size'          = ConvertSize($size_in_bytes)
	}
	$CatchTypeTable.Add($obj)
}

if (($CatchTypeTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No Cached were found'
	}
	$CatchTypeTable.Add($obj)
}

# Cached Object Type Size
$CatchObjectTypeTable = New-Object 'System.Collections.Generic.List[System.Object]'
$CatchType02 = ExecuteSQL "select objtype as Cached_Object_Type,COUNT(*) as Number_of_Plans,SUM(size_in_bytes) as size_in_bytes
from sys.dm_exec_cached_plans group by objtype"

foreach ($CatchType2 in $CatchType02)
{
	$Cached_Object_Type = $CatchType2.Cached_Object_Type
	$Number_of_Plans = $CatchType2.Number_of_Plans
	$size_in_bytes  = $CatchType2.size_in_bytes
	
	$obj = [PSCustomObject]@{
		'CachedObjectType' = $Cached_Object_Type
		'NumberOfPlans'    = $Number_of_Plans
		'Size'             = ConvertSize($size_in_bytes)
	}
	$CatchObjectTypeTable.Add($obj)
}

if (($CatchObjectTypeTable).Count -eq 0)
{
	$Obj = [PSCustomObject]@{
		Information = 'Information: No Cached were found'
	}
	$CatchObjectTypeTable.Add($obj)
}


<#####################################################################################>

<#####################################################################################>
Write-Host "Compiling Report..." -ForegroundColor Green

$Options = @('Dashboard','BaseInfo','Database','Jobs','Backup&Restore','Mirror','Replication','AlwaysOn-AG','AlwaysOn-FCI','Perfmon')
$MSSQLReport = New-Object 'System.Collections.Generic.List[System.Object]'
$MSSQLReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString "" -RightLogoString $RightLogo ))
$MSSQLReport.Add($(Get-HTMLTabHeader -TabNames $Options))

<# Dashboard option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[0] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$DatabaseWarnTable = $DatabaseInfoTable | Where-Object {($_.UserAccessDesc -ne "MULTI_USER") -or ($_.IsReadOnly -ne 0) -or ($_.StateDesc -ne "ONLINE")}
if ( $DatabaseWarnTable.Count -gt 0)
{
  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "WARNING : Database" -BackgroundShade 2))
  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $DatabaseWarnTable -Red '$this.DatabaseId -gt 0') -HideFooter))
  $MSSQLReport.Add($(Get-HTMLContentClose))
}
if ( $DiskWarnCnt -gt 0)
{
  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "WARNING : Disk Space less then 10% !" -BackgroundShade 2))
  foreach ($i in 0..($DiskWarnCnt-1))
  {
    $MSSQLReport.Add($(Get-HTMLColumnOpen -ColumnNumber ($i+1) -ColumnCount $DiskWarnCnt))
    $MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieDiskWarnLinks[$i] -DataSet $DiskWarnInfoPies[$i]))
    $MSSQLReport.Add($(Get-HTMLColumnClose))
  }
  $MSSQLReport.Add($(Get-HTMLContentClose))
}
$DisableLoginNameTable = $LoginNameTable | Where-Object {$_.IsDisabled -eq 1}
if ( $DisableLoginNameTable.Count -gt 0)
{
  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "WARNING : Login Name were disabled" -BackgroundShade 2))
  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $DisableLoginNameTable -Yellow '$this.IsDisabled -eq 1') -HideFooter))
  $MSSQLReport.Add($(Get-HTMLContentClose))
}
$FailedJobsTable = $JobsHistoryTable | Where-Object { ($_.RunStatus -eq 0) -and ([datetime]::parseexact($_.RunDate, 'yyyyMMdd',[System.Globalization.CultureInfo]::DefaultThreadCurrentCulture) -gt (Get-date).AddDays(-1))}
if ( @($FailedJobsTable).Count -gt 0)
{
  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "WARNING : Failed jobs in the last one day" -BackgroundShade 2))
  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $FailedJobsTable -Red '$this.RunStatus -eq 0') -HideFooter))
  $MSSQLReport.Add($(Get-HTMLContentClose))
}
$MSSQLReport.Add($(Get-HTMLTabContentClose))


<# BaseInfo option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[1] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "OS Information"))
$MSSQLReport.Add($(Get-HTMLContentTable $OSInfoTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Logical Disk Space"))
foreach ($i in 0..($DiskCnt-1))
{
  $MSSQLReport.Add($(Get-HTMLColumnOpen -ColumnNumber ($i+1) -ColumnCount $DiskCnt))
  $MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieDiskLinks[$i] -DataSet $DiskInfoPies[$i]))
  $MSSQLReport.Add($(Get-HTMLColumnClose))
}
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server Version"))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server Product Version"))
$MSSQLReport.Add($(Get-HTMLContentTable $SQLVersionTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server Service Status"))
$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $SQLServiceTable -Yellow '$this.Status -eq "Stopped"') -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server Configurations"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLConfTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLTabContentClose))


<# Database option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[2] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $DatabaseInfoTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database File Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $DatabaseFileInfoTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Transaction Log"))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Transaction Log Size"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLTranTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Transaction Log Size"))
$MSSQLReport.Add($(Get-HTMLBarChart -ChartObject $SQLTranLogBarChar -DataSet $SQLTranLogs))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Error Log"))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server Error Log Size"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLErrorLogTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Agent Error Log Size"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $AgentErrorLogTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Login Name"))
$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $LoginNameTable -Yellow '$this.IsDisabled -eq 1') -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLTabContentClose))


<# Jobs option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[3] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Jobs Base Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $JobsInfoTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Jobs History Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $JobsHistoryTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLTabContentClose))


<# Backup & Restore option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[4] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database Backup"))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "The Last Backup"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $BackupInfoTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Backup Information"))
$MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieBackupLinks -DataSet $BackupInfoPie))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Backup History Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable ($BackupDetailTable | Sort LatestBackupTime -Descending) -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Restore History Information"))
$MSSQLReport.Add($(Get-HTMLContentDataTable ($RestoreTable | Sort RestoreDate -Descending) -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLTabContentClose))

<# Mirror option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[5] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLTabContentClose))

<# Replication option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[6] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLTabContentClose))

<# AlwaysOn-AG option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[7] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLTabContentClose))

<# AlwaysOn-FCI option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[8] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLTabContentClose))

<# Perfmon option #>
$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[9] -TabHeading ("Report: " + (Get-Date -Format yyyy-MM-dd))))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Buffer allocation" -IsHidden ))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database buffer pool allocation"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $DBBufferpoolTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database dirty page"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $DBDirtyPageTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Cached Type Usage" -IsHidden ))
$MSSQLReport.Add($(Get-HTMLColumn1of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Cached Type Size"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $CatchTypeTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLColumn2of2))
$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Cached Object Type Size"))
$MSSQLReport.Add($(Get-HTMLContentDataTable $CatchObjectTypeTable -HideFooter))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLColumnClose))
$MSSQLReport.Add($(Get-HTMLContentClose))
$MSSQLReport.Add($(Get-HTMLTabContentClose))


# Close HTML Page
$MSSQLReport.Add($(Get-HTMLClosePage))

Write-Host "Showing Report.`n" -ForegroundColor Green

Save-HTMLReport -ShowReport -ReportContent $MSSQLReport -ReportName $ReportName -ReportPath $ReportSavePath

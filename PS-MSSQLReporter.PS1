[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.SMO") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null

##create input form
$z = 20		#Height: box/text 
$left = 20	#Left: box/text
$inputForm               = New-Object System.Windows.Forms.Form
$inputForm.Size          = New-Object System.Drawing.Size(350, 150)
$inputForm.Text          = "Generate SQL Server Html Report"
$inputForm.StartPosition = "CenterScreen"
[System.Windows.Forms.Application]::EnableVisualStyles()

######################  Instance Name ######################
$y = 15
$inputForm.Height += 30
$objLabel          = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($left,  $y)
$objLabel.Size     = New-Object System.Drawing.Size(280, $z)
$objLabel.Text     = "Instance name:"
$inputForm.Controls.Add($objLabel)
$y+=20
$inputForm.Height     +=20
$Instance_Box          = New-Object System.Windows.Forms.ComboBox #TextBox
$Instance_Box.Location = New-Object System.Drawing.Size($left,  $y)
$Instance_Box.Size     = New-Object System.Drawing.Size(290, $z)
$SQLItem = "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server"
If(Test-Path $SQLItem){
	$InstalledIns = (Get-ItemProperty $SQLItem).InstalledInstances | Sort-Object
	$InstalledIns | ForEach {
		[void] $Instance_Box.Items.Add($_)
	}
	$Instance_Box.SelectedIndex=0
}
$Instance_Box.DropDownStyle = 'DropDownList'
$inputForm.Controls.Add($Instance_Box)

######################  Authentication ######################
$y+=30
$inputForm.Height +=20
$objLabel          = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($left,  $y)
$objLabel.Size     = New-Object System.Drawing.Size(280, $z)
$objLabel.Text     = "Authentication:"
$inputForm.Controls.Add($objLabel)
$y+=20
$inputForm.Height +=20
$Auth_Box          = New-Object System.Windows.Forms.ComboBox
$Auth_Box.Location = New-Object System.Drawing.Size($left,  $y)
$Auth_Box.Size     = New-Object System.Drawing.Size(290, $z)
$Auth_Box.Height   = 60
[void] $Auth_Box.Items.Add('SQL Server Authentication')
[void] $Auth_Box.Items.Add('Windows Authentication')
$Auth_Box.SelectedIndex =0
$Auth_Box.DropDownStyle = 'DropDownList'
$inputForm.Controls.Add($Auth_Box)

######################  Username ######################
$y+=30
$inputForm.Height +=20
$objLabel          = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($left,  $y)
$objLabel.Size     = New-Object System.Drawing.Size(280, $z)
$objLabel.Text     = "Username:"
$inputForm.Controls.Add($objLabel)
$y+=20
$inputForm.Height     +=20
$Username_Box          = New-Object System.Windows.Forms.TextBox 
$Username_Box.Location = New-Object System.Drawing.Size($left,  $y)
$Username_Box.Size     = New-Object System.Drawing.Size(290, $z)
$inputForm.Controls.Add($Username_Box)

######################  Password ######################
$y+=30
$inputForm.Height +=20
$objLabel          = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size($left,  $y)
$objLabel.Size     = New-Object System.Drawing.Size(280, $z)
$objLabel.Text     = "Password:"
$inputForm.Controls.Add($objLabel)
$y+=20
$inputForm.Height     +=20
$Password_Box          = New-Object System.Windows.Forms.TextBox 
$Password_Box.Location = New-Object System.Drawing.Size($left,  $y)
$Password_Box.Size     = New-Object System.Drawing.Size(290, $z)
$Password_Box.UseSystemPasswordChar = $true 
$inputForm.Controls.Add($Password_Box)
#$objLabel.Text = ($objLabel.Text -replace '\*','')

######################  Button ######################
$Click_Code={
	$okButton.Enabled = 0
	[int]$LoginType   = $Auth_Box.SelectedIndex		#Authentication:0=SQL;1=Win
	[string]$Instance = $Instance_Box.Text.Trim()	#Instance Name
	[string]$Username = $Username_Box.Text.Trim()	#Username
	[string]$Password = $Password_Box.Text.Trim()	#Password
	
	$StatusBar.Controls.Add($BarProgress)
	Func-Main -T $LoginType -I $Instance -U $Username -P $Password
	$StatusBar.Controls.Remove($BarProgress)
	$StatusBar.Text = "Completed!"
	$Instance_Box.Focus()
	$okButton.Enabled  = 1
}

##create OK button
$okButton          = New-Object System.Windows.Forms.Button
$okButton.Size     = New-Object System.Drawing.Size(75, 23)
$okButton.Text     = "OK"
$okButton.Add_Click($Click_Code)
$inputForm.Controls.Add($okButton)

##create Cancel button
$cancelButton          = New-Object System.Windows.Forms.Button 
$cancelButton.Size     = New-Object System.Drawing.Size(75,23)
$cancelButton.Text     = "Close"
$cancelButton.Add_Click({[Environment]::Exit(0)})
$inputForm.Controls.Add($cancelButton)
$inputForm.CancelButton = $cancelButton

##Set button location
$okButton.Location     = New-Object System.Drawing.Size(80, (30+$y+$z))
$cancelButton.Location = New-Object System.Drawing.Size(180, (30+$y+$z))

######################  Status Bar ######################
$BarPanel             = New-Object System.Windows.Forms.Panel # Panel Label
$BarPanel.Left        = 0
$BarPanel.Top         = (58+$y+$z)
$BarPanel.Width       = 350
$BarPanel.Height      = 22
$inputForm.Controls.Add($BarPanel)

$StatusBar            = New-Object System.Windows.Forms.StatusBar
$StatusBar.Text       = ""
$BarPanel.Controls.Add($StatusBar)

$BarProgress          = New-Object System.Windows.Forms.ProgressBar
$BarProgress.Style    ="Continuous"
$BarProgress.Location = new-object System.Drawing.Size(0,0)
$BarProgress.size     = new-object System.Drawing.Size(350,20)
$BarProgress.Minimum  = 0
$BarProgress.Maximum  = 100
#$StatusBar.Controls.Add($BarProgress)

######################  show form ######################
$inputForm.Topmost       = $False #是否置顶
$inputForm.MinimizeBox   = $False
$inputForm.MaximizeBox   = $False
$inputForm.AutoSizeMode  = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$inputForm.SizeGripStyle = [System.Windows.Forms.SizeGripStyle]::Hide
$inputForm.Add_Shown({
	$inputForm.Activate()
	$Instance_Box.Focus()
})

##################################################################
# Authentication 变化级联 Username & Password
$Auth_Box.Add_SelectedIndexChanged({
	#Write-Host $Auth_Box.SelectedIndex,$Auth_Box.SelectedItem,$Auth_Box.Text
	if($Auth_Box.SelectedIndex -eq 1){
		$Username_Box.Enabled=0
		$Password_Box.Enabled=0
	}else{
		$Username_Box.Enabled=1
		$Password_Box.Enabled=1
	}
})


####################################################################################################################################
#Unknown, String, Unicode, Byte, BigEndianUnicode, UTF8, UTF7, UTF32, Ascii, Default, Oem, BigEndianUTF32”
#$PSDefaultParameterValues['*:Encoding'] = 'Unicode'

Function ImportModule{
	try {
		if (!(Get-Module -ListAvailable -Name ReportHTML)){
			Install-Module -Name ReportHTML -Force
			Import-Module ReportHTML -ErrorAction SilentlyContinue
		} else {
			Import-Module ReportHTML -ErrorAction SilentlyContinue
		}
	} catch {
		Write-Host "Module ReportHTML does not exist!" -ForegroundColor Red
		Write-Host "Stop!" -ForegroundColor Red
		Exit 1
	}

	try {
		if (!(Get-Module -ListAvailable -Name ScheduledTasks)){
			#Install-Module -Name TaskScheduler -Force -AllowClobber
			Install-Module -Name ScheduledTasks -Force
			Import-Module ScheduledTasks -ErrorAction SilentlyContinue
		} else {
			Import-Module ScheduledTasks -ErrorAction SilentlyContinue
		}
	} catch {
		Write-Host "Module ScheduledTasks does not exist!" -ForegroundColor Red
	}
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

# Connect to sql server (Local Windows Authentication )
Function GetSqlConnection{
	Param(
		[Parameter(Position = 0, Mandatory = $true)][Alias ('T')][int]$LoginType,
		[Parameter(Position = 1, Mandatory = $true)][Alias ('I')][String]$Instance,
		[Parameter(Position = 2, Mandatory = $false)][Alias ('U')][String]$UserName,
		[Parameter(Position = 3, Mandatory = $false)][Alias ('P')][String]$Password
	)
	If ( $LoginType -eq 1){
		$ConnectionString = "Data Source=$Instance;Initial Catalog=master;Integrated Security=SSPI;"
	} Else {
		$ConnectionString = "Data Source=$Instance;Initial Catalog=master;user id=$UserName;pwd=$Password"
	}
	Try{
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
		$SqlConnection.Open()
		Return $SqlConnection
	}
	Catch{
		Write-Warning $_
		[System.Windows.Forms.MessageBox]::Show($_,"Error",0)
		Exit 1
	}
}


# Execute query statement
Function ExecuteSQL{
    Param(
        [Parameter(position = 0 , Mandatory = $true)]$SqlConn,
        [Parameter(position = 1 , Mandatory = $true)][string]$SqlString
    )
    Try{
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = $SqlString
        $SqlCmd.Connection = $SqlConn
		$Reader= $SqlCmd.ExecuteReader()
        $DataTable = New-Object System.Data.DataTable
        $DataTable.Load($Reader)
    }
    Catch {
        Write-Warning $_
		Return $Null
    }
    Return $DataTable
}

# Close SQL Connection 
Function CloseSqlConnection{
    Param(
        [Parameter(position = 0 , Mandatory = $true)]$SqlConn
    )
	If ($sqlConn.State -eq 'Open'){
		$SqlConn.Close()
		$SqlConn.Dispose()
	}
}

####################################################################################################################################

Function Func-Main{
	Param(
		[Parameter(Position = 0, Mandatory = $true)][Alias ('T')][int]$LoginType,
		[Parameter(Position = 1, Mandatory = $true)][Alias ('I')][String]$Instance,
		[Parameter(Position = 2, Mandatory = $false)][Alias ('U')][String]$UserName,
		[Parameter(Position = 3, Mandatory = $false)][Alias ('P')][String]$Password
	)
	
	$ReportTitle = "SQL Server Report"
	$ReportSavePath = [Environment]::GetFolderPath("Desktop") #Get-Location
	$ReportName = ( "MSSQLReport_$(hostname)_$($Instance)_" + $(get-date -format "yyyyMMdd"))
	$LeftLogo =""
	$RightLogo="https://raw.githubusercontent.com/hzc2012/PS-MSSQLReporter/master/example/PS-MSSQLReporter.png"
	$LocalHost = $env:computerName
	
	If($Instance -eq "MSSQLSERVER"){
		$Instance = $LocalHost 
	} Else {
		$Instance = "$($LocalHost)\$($Instance)"
	}
	
	(Write-Host -NoNewline "`nInstance: " -ForegroundColor Yellow), (Write-Host $Instance -ForegroundColor White)
	(Write-Host -NoNewline "Report Title: " -ForegroundColor Yellow), (Write-Host $ReportTitle -ForegroundColor White)
	(Write-Host -NoNewline "Report Path: " -ForegroundColor Yellow), (Write-Host $ReportSavePath -ForegroundColor White)
	
	Write-Host "Getting data..." -ForegroundColor Green
	
	#ReportHTML
	$TmpDir = Get-ChildItem $env:TEMP -Recurse -filter "IXP000.TMP" -Directory | % { $_.fullname }
	$PSModulePath = $env:PSModulePath -split ';' | Where-Object { $_ -notmatch "SQL Server" -and $_ -notmatch "Users" }
	try{
		$PSModulePath | ForEach {
			$Destination = Join-Path $_ "ReportHTML"
			If(!(Test-Path $Destination)){
				If($TmpDir){
					If(Test-Path $TmpDir){
						Copy-Item $TmpDir\* -Destination $Destination -Recurse -Force
					}
				}
			}
		}
	}catch {
		Write-Warning $_
		Exit 1
	}
	
	# Install or Import Module
	try{
		Import-Module ReportHTML -ErrorAction SilentlyContinue
	}catch {
		Write-Warning $_
		Exit 1
	}
	
	$SqlConn = GetSqlConnection -T $LoginType -I $Instance -U $Username -P $Password


	<################################################################################################
	# Dashboard
	################################################################################################>
	$Nics = Get-NetAdapter *
	$Win1 = Get-CimInstance Win32_ComputerSystem
	$Win2 = Get-CimInstance Win32_Operatingsystem

	$WarningSystem = New-Object 'System.Collections.Generic.List[System.Object]'
	$WarningAlls = ExecuteSQL $SqlConn "declare @t table(drive varchar(5),sizemb bigint);insert into @t exec master.dbo.xp_fixeddrives;select * from @t;"

	foreach ($WarningAll in $WarningAlls)
	{
		if($WarningAll.sizemb -lt 50000)
		{
			$Warning     = "磁盘 {0} 空间小于 50 GB，为 {1} MB，请预留足够磁盘空间" -f $WarningAll.drive,$WarningAll.sizemb
			$obj = [PSCustomObject]@{ 'Warning' = $Warning }
			$WarningSystem.Add($obj)
		}
	}

	#$Nics = Get-NetAdapter *
	$Nics | ForEach-Object {
		if($_.Speed -le 100000000){
			$Warning     = "网卡 {0} 速度是 {1}，属于慢速网卡" -f $_.Name,$_.LinkSpeed
			$obj = [PSCustomObject]@{ 'Warning' = $Warning }
			$WarningSystem.Add($obj)
		}
	}

	if($Win1.NumberOfLogicalProcessors * $Win1.NumberOfProcessors -lt 4)
	{
		$Warning     = "处理器核心数过少，建议添加CPU"
		$obj = [PSCustomObject]@{ 'Warning' = $Warning }
		$WarningSystem.Add($obj)
	}

	if($Win2.TotalVisibleMemorySize -lt 4000000)
	{
		$Warning     = "服务器内存过少，为 {0} GB，请扩大内存" -f [Math]::Round($Win2.TotalVisibleMemorySize/1024/1024,0)
		$obj = [PSCustomObject]@{ 'Warning' = $Warning }
		$WarningSystem.Add($obj)
	}

	if( $Win2.OSArchitecture -notmatch "[64]")
	{
		$Warning     = "操作系统架构是 {0}，建议改为 64 位系统" -f $Win2.OSArchitecture
		$obj = [PSCustomObject]@{'Warning' = $Warning }
		$WarningSystem.Add($obj)
	}

	Get-Service | Where-Object { ($_.StartType -ne 'Automatic') -and (($_.Name -like 'MSSQL$*') -or ($_.Name -like 'SQLAgent$*') -or ($_.Name -eq 'MSSQLSERVER') -or ($_.Name -eq 'SQLSERVERAGENT'))} | ForEach-Object {
		$Warning     = "服务 {0} 未配置为自动启动" -f $_.DisplayName
		$obj = [PSCustomObject]@{ 'Warning' = $Warning }
		$WarningSystem.Add($obj)
	}

	if (($WarningSystem).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No Warning!'
		}
		$WarningSystem.Add($obj)
	}

	$BarProgress.Value =  5
	###########################################

	$WarningInstance = New-Object 'System.Collections.Generic.List[System.Object]'
	$WarningAlls = ExecuteSQL $SqlConn "select 
	(case when ty='collation' then '发现 '+convert(varchar(10),cnt)+' 个数据库字符集不是 Chinese_PRC_CI_AS'
		when ty='recovery' then '发现 '+convert(varchar(10),cnt)+' 个数据库为简单模式'
		when ty='verify' then '发现 '+convert(varchar(10),cnt)+' 个数据库未启用页验证'
		when ty='productversion' then 'SQL Server 版本过低，请升级至 SQL Server 2012'
		when ty='notbackup' then '发现 '+convert(varchar(10),cnt)+' 个数据库在近 7 天没有备份'
		when ty='TcpPort' then '数据库实例使用了默认端口 1433，已知的端口存在安全隐患，请更改'
		when ty='TcpDynamicPorts' then '数据库实例使用动态端口，建议更改为静态端口'
		when ty='maxdop' then '最大并行度（max degree of parallelism）值为 '+convert(varchar(10),cnt)+'，设置不合理. 建议并行度设置为 '
			+convert(varchar(10),(select (case when cpu_count>8 then 8 when cpu_count between 4 and 8 then 4 else 0 end) from sys.dm_os_sys_info with(nolock)))
		when ty='system_health' then '系统扩展事件 system_health 未启用'
		when ty='is_auto_create_stats_on' then '发现 '+convert(varchar(10),cnt)+' 个数据库未自动创建统计信息'
		when ty='is_auto_update_stats_on' then '发现 '+convert(varchar(10),cnt)+' 个数据库未自动更新统计信息'
		when ty='jobsid' then '发现 '+convert(varchar(10),cnt)+' 个代理作业的所有者已不存在'
		when ty='dbsid' then '发现 '+convert(varchar(10),cnt)+' 个数据库的所有者已不存在'
		when ty='disjobs' then '发现 '+convert(varchar(10),cnt)+' 个代理作业处于禁用状态'
		end) as Warning
	from(
	select 'collation' as ty,count(*) cnt from sys.databases with(nolock) where database_id>4 and collation_name<>'Chinese_PRC_CI_AS' 
	union all
	select 'recovery' as ty,count(*) cnt from sys.databases with(nolock) where database_id>4 and recovery_model=3
	union all
	select 'verify' as ty,count(*) cnt from sys.databases with(nolock) where database_id>4 and page_verify_option=0
	union all
	select 'productversion' as ty,(case when convert(varchar(128),serverproperty('productversion')) like '[8|9]%' then 1 else 0 end) as cnt
	union all
	select 'notbackup' as ty,count(*) cnt from(
	select name from sys.databases with(nolock) where database_id<>2
	except
	select database_name from msdb.dbo.backupset where backup_start_date >= dateadd(d,-7,getdate())
	)a1
	union all
	select 'TcpPort' as ty,count(*) as cnt from sys.dm_server_registry with(nolock)
	where registry_key like '%IPALL' AND value_name = 'TcpPort' AND value_data='1433'
	union all
	select 'TcpDynamicPorts' as ty,count(*) as cnt from sys.dm_server_registry with(nolock)
	where registry_key like '%IPALL' AND value_name = 'TcpDynamicPorts' AND value_data<>''
	union all
	select 'maxdop' as ty,value_in_use as cnt from sys.configurations with(nolock)
	where name='max degree of parallelism'
	and value_in_use > (select cpu_count from sys.dm_os_sys_info with(nolock))
	union all
	SELECT 'system_health' AS ty,count(*) as cnt FROM sys.server_event_sessions WHERE startup_state=0 AND name='system_health'
	union all
	select 'is_auto_create_stats_on' as ty,isnull(sum(case when is_auto_create_stats_on=1 then 1 else 0 end),0) cnt
	from sys.databases with(nolock) where is_auto_create_stats_on=0
	union all
	select 'is_auto_update_stats_on' as ty,isnull(sum(case when is_auto_update_stats_on=1 then 1 else 0 end),0) cnt
	from sys.databases with(nolock) where is_auto_update_stats_on=0
	union all
	select 'jobsid' ty,count(*) cnt from msdb.dbo.sysjobs a with(nolock)
	left join sys.syslogins b with(nolock) on a.owner_sid=b.sid where b.sid is null
	union all
	select 'dbsid' ty,count(*) cnt from sys.databases a with(nolock)
	left join sys.syslogins b with(nolock) on a.owner_sid=b.sid where b.sid is null
	union all
	select 'disjobs' ty,count(*) cnt from msdb.dbo.sysjobs where enabled=0
	)t where cnt>0
	"

	foreach ($WarningAll in $WarningAlls)
	{
		$Warning = $WarningAll.Warning
		$obj = [PSCustomObject]@{'Warning' = $Warning}
		$WarningInstance.Add($obj)
	}

	$BarProgress.Value =  10
	$WarningAlls = ExecuteSQL $SqlConn "select DB_NAME(database_id) as dbname,convert(decimal(18,2),sum(size)*8.0/1024/1024) as size_gb
	from sys.master_files where type = 1 group by database_id having sum(size)*8.0/1024/1024>=1"
	foreach ($WarningAll in $WarningAlls)
	{
		$Warning     = "数据库 {0} 的日志大小为 {1} GB，请及时备份或检查是否正常" -f $WarningAll.dbname,$WarningAll.size_gb
		$obj = [PSCustomObject]@{'Warning'= $Warning }
		$WarningInstance.Add($obj)
	}

	$WarningAlls = ExecuteSQL $SqlConn ";WITH T AS(
	SELECT sum(case when is_disabled=1 then 1 else 0 end) is_disabled
	,sum(case when is_policy_checked=1 then 1 else 0 end) is_policy_checked
	,sum(case when sysadmin=1 then 1 else 0 end) sysadmin
	,(SELECT COUNT(*) FROM sys.sysprocesses with(nolock)
	 WHERE spid>50 AND spid<>@@SPID AND open_tran=0 AND status='sleeping' AND last_batch <= DATEADD(MINUTE,-10,GETDATE()))as FreeSession
	,(SELECT COUNT(*) FROM MASTER.dbo.syscursors WHERE STATUS<>-1) AS Cur
	FROM sys.syslogins a with(nolock) INNER JOIN sys.sql_logins b with(nolock) ON a.sid=b.sid and b.type='S'
	)SELECT '发现 '+convert(varchar(10),is_disabled) +' 个账号处于禁用状态，请确认账号是否有效' as Warning
	FROM T WHERE is_disabled>0
	UNION ALL
	SELECT '发现 '+convert(varchar(10),is_policy_checked) +' 个账号未启用密码策略，可能存在若密码，请启用密码策略并修改为复杂密码'  as Warning
	FROM T WHERE is_policy_checked>0
	UNION ALL
	SELECT '发现 '+convert(varchar(10),sysadmin) +' 个账号有管理员（sysadmin）权限，请确认账号权限范围'  as Warning
	FROM T WHERE sysadmin>0
	UNION ALL
	SELECT '发现 '+convert(varchar(10),FreeSession) +' 个空闲时间超过 10 分钟的会话'  as Warning
	FROM T WHERE FreeSession>0
	UNION ALL
	SELECT '发现 '+convert(varchar(10),Cur) +' 个游标未关闭'  as Warning
	FROM T WHERE FreeSession>0
	"
	foreach ($WarningAll in $WarningAlls)
	{
		$Warning     = $WarningAll.Warning
		$obj = [PSCustomObject]@{ 'Warning' = $Warning }
		$WarningInstance.Add($obj)
	}

	$BarProgress.Value =  15

	$WarningAlls = ExecuteSQL $SqlConn "select name from sys.configurations where value_in_use=1 and name in('xp_cmdshell','Ole Automation Procedures')"
	foreach ($WarningAll in $WarningAlls)
	{
		$Warning     = "参数 {0} 配置不合理，存在安全隐患，建议关闭此参数" -f $WarningAll.name
		$obj = [PSCustomObject]@{'Warning' = $Warning }
		$WarningInstance.Add($obj)
	}

	$WarningAlls = ExecuteSQL $SqlConn "create table #tmp(DBName nvarchar(128) default(DB_NAME()),UserName nvarchar(128),UserSID varbinary(256));
	EXEC sp_MSforeachdb 'use [?];insert into #tmp(UserName,UserSID) exec dbo.sp_change_users_login @Action=''Report'';';
	select count(distinct DBName) as Users from #tmp;"

	foreach ($WarningAll in $WarningAlls)
	{
		if ($WarningAll.Users -ne 0)
		{
			$Warning     = "发现 {0} 个数据库存在孤立用户，请检查确认用户是否有效" -f $WarningAll.Users
			$obj = [PSCustomObject]@{'Warning' = $Warning }
			$WarningInstance.Add($obj)
		}
	}

	$BarProgress.Value =  20

	$WarningAlls = ExecuteSQL $SqlConn "CREATE TABLE #LoginFailed(LogDate DATETIME,ProcessInfo VARCHAR(20),TextInfo VARCHAR(200))
	INSERT INTO #LoginFailed EXEC xp_readerrorlog 0,1,N'Login failed',N'',NULL,NULL,NULL
	SELECT COUNT(*) Users FROM #LoginFailed
	DROP TABLE #LoginFailed"

	foreach ($WarningAll in $WarningAlls)
	{
		if ($WarningAll.Users -ne 0)
		{
			$Warning     = "发现 {0} 次登录失败记录" -f $WarningAll.Users
			$obj = [PSCustomObject]@{'Warning' = $Warning }
			$WarningInstance.Add($obj)

		}
	}

	if (($WarningInstance).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No Warning!'
		}
		$WarningInstance.Add($obj)
	}

	$BarProgress.Value =  25
	###########################################
	$Perfmon = New-Object 'System.Collections.Generic.List[System.Object]'
	$Mons = ExecuteSQL $SqlConn "declare @sql varchar(8000)
	declare @dbname varchar(128)
	create table #tab(dbname varchar(128),cnt int)
	declare cur cursor for select name from sys.databases with(nolock)
	where database_id>4 and state=0 and user_access=0
	open cur
	fetch next from cur into @dbname
	while @@FETCH_STATUS=0
	begin
		set @sql = 'use ['+@dbname+'];
	;with tab as(
	SELECT SCHEMA_NAME(t.schema_id) [schema_name],t.name as [table_name],t1.name as [index_name]
	,t1.type_desc,t1.is_unique,t1.is_primary_key,t1.is_unique_constraint,t1.has_filter,isnull(t1.filter_definition,'''') as filter_definition
	,isnull(STUFF((SELECT '',''+t4.name FROM sys.sysindexkeys t2 with(nolock)
		inner join sys.index_columns t3 with(nolock) ON t2.id=t3.object_id and t2.indid=t3.index_id and t2.colid=t3.column_id
		inner join sys.syscolumns t4 with(nolock) ON t2.id=t4.id and t2.colid=t4.colid
		WHERE t2.id=t1.object_id and t1.index_id=t2.indid  and t2.keyno <> 0 ORDER BY t3.key_ordinal FOR XML PATH('''')),1,1,''''),'''') AS index_cols
	,isnull(STUFF((SELECT '',''+t4.name FROM sys.sysindexkeys t2 with(nolock)
		inner join sys.index_columns t3 with(nolock) ON t2.id=t3.object_id and t2.indid=t3.index_id and t2.colid=t3.column_id
		inner join sys.syscolumns t4 with(nolock) ON t2.id=t4.id and t2.colid=t4.colid
		WHERE t2.id=t1.object_id and t1.index_id=t2.indid  and t2.keyno = 0 ORDER BY t3.key_ordinal FOR XML PATH('''')),1,1,''''),'''')  AS include_cols
	FROM sys.tables as t with(nolock)
	inner join sys.indexes as t1 with(nolock) on (t1.index_id > 0 and t1.is_hypothetical = 0) and (t1.object_id=t.object_id)
	WHERE t1.type in(1,2)
	) insert into #tab(dbname,cnt)
	select '''+@dbname+''' as dbname,count(*) cnt from(
	select [schema_name],table_name,index_cols,include_cols,filter_definition,count(*) idxCnt
	from tab
	group by [schema_name],table_name,index_cols,include_cols,filter_definition
	having count(*)>=2
	)t'
		exec(@sql)
		fetch next from cur into @dbname
	end
	close cur
	deallocate cur
	select * from #tab where cnt>0 order by dbname"

	foreach ($Mon in $Mons)
	{
		$Warning     = "发现数据库 {0} 有 {1} 个重复索引" -f $Mon.dbname,$Mon.cnt
		$Suggestion  = ""
		
		$obj = [PSCustomObject]@{
			'Warning'     = $Warning
			'Suggestion'  = $Suggestion
		}
		$Perfmon.Add($obj)
	}

	$BarProgress.Value =  30
	$Mons = ExecuteSQL $SqlConn "if object_id('tempdb..#tmp') is not null delete from #tab
	declare @sql varchar(8000)
	declare @dbname varchar(128)
	declare cur cursor for select name from sys.databases with(nolock)
	where database_id<>2 and state=0 and user_access=0
	open cur
	fetch next from cur into @dbname
	while @@FETCH_STATUS=0
	begin
		set @sql = 'use ['+@dbname+'];insert into #tab(dbname,cnt) select '''+@dbname+''',count(*) from sys.sysindexes a inner join sys.tables b on a.id=b.object_id where indid in(0,1)and rows>5000000'
		exec(@sql)
		fetch next from cur into @dbname
	end
	close cur
	deallocate cur
	select * from #tab where cnt>0 order by dbname"

	foreach ($Mon in $Mons)
	{
		$Warning     = "发现数据库 {0} 有 {1} 个表数据量超过 500 万行" -f $Mon.dbname,$Mon.cnt
		$Suggestion  = ""
		
		$obj = [PSCustomObject]@{
			'Warning'     = $Warning
			'Suggestion'  = $Suggestion
		}
		$Perfmon.Add($obj)
	}

	if (($Perfmon).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No Warning!'
		}
		$Perfmon.Add($obj)
	}

	$BarProgress.Value =  35
	###########################################
	$MSSQLErrorLog = New-Object 'System.Collections.Generic.List[System.Object]'
	$SQLErrorLogs = ExecuteSQL $SqlConn "DECLARE @StartTime DATE;
	SET @StartTime = DATEADD(DAY,-7,GETDATE());
	EXEC XP_READERRORLOG 1,1,N'Error',NULL,@StartTime,NULL,'ASC'"

	foreach ($SQLErrorLog in $SQLErrorLogs)
	{
		$LogDate     = $SQLErrorLog.Item(0)
		$ProcessInfo = $SQLErrorLog.Item(1)
		$Text        = $SQLErrorLog.Item(2)
		
		$obj = [PSCustomObject]@{
			'LogDate'     = $LogDate
			'ProcessInfo' = $ProcessInfo
			'Text'        = $Text 
		}
		$MSSQLErrorLog.Add($obj)
	}

	$BarProgress.Value =  40
	###########################################
	$Begin = (Get-date).AddDays(-7)

	$WarningEventLog= New-Object 'System.Collections.Generic.List[System.Object]'
	$EventLogs = Get-EventLog -LogName System -Newest 20 -EntryType Error,Warning -After $Begin

	foreach ($EventLog in $EventLogs)
	{
		$Index      = $EventLog.Index
		$Time       = $EventLog.Time
		$EntryType  = $EventLog.EntryType
		$Source     = $EventLog.Source
		$InstanceID = $EventLog.InstanceID
		$Message    = $EventLog.Message
		
		$obj = [PSCustomObject]@{
			'Index'       = $Index
			'Time'        = $Time
			'EntryType'   = $EntryType
			'Source'      = $Source
			'InstanceID'  = $InstanceID
			'Message'     = $Message
		}
		$WarningEventLog.Add($obj)
	}


	$WarningApplicationLog= New-Object 'System.Collections.Generic.List[System.Object]'
	$ApplicationLogs = Get-EventLog -LogName Application -Newest 20 -EntryType Error,Warning -After $Begin

	foreach ($ApplicationLog in $ApplicationLogs)
	{
		$Index      = $ApplicationLog.Index
		$Time       = $ApplicationLog.Time
		$EntryType  = $ApplicationLog.EntryType
		$Source     = $ApplicationLog.Source
		$InstanceID = $ApplicationLog.InstanceID
		$Message    = $ApplicationLog.Message
		
		$obj = [PSCustomObject]@{
			'Index'       = $Index
			'Time'        = $Time
			'EntryType'   = $EntryType
			'Source'      = $Source
			'InstanceID'  = $InstanceID
			'Message'     = $Message
		}
		$WarningApplicationLog.Add($obj)
	}

	#################################################################################################
	$BarProgress.Value =  45
	# OS information
	$OSInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'

	#$Win1 = Get-CimInstance Win32_ComputerSystem
	#$Win2 = Get-CimInstance Win32_Operatingsystem
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Domain";'Value' = $Win1.Domain})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Name";'Value' = $Win1.Name})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Caption";'Value' = $Win2.Caption})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "OS Manufacturer";'Value' = $Win2.Manufacturer})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "OS Language";'Value' = $Win2.OSLanguage})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "OS Architecture";'Value' = $Win2.OSArchitecture})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Current TimeZone";'Value' = $Win1.CurrentTimeZone})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Version";'Value' = $Win2.Version})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "System Type";'Value' = $Win1.SystemType})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Model";'Value' = $Win1.Model})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Manufacturer";'Value' = $Win1.Manufacturer})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Number Of Processors";'Value' = $Win1.NumberOfProcessors})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Number Of Logical Processors";'Value' = $Win1.NumberOfLogicalProcessors})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Total Virtual Memory";'Value' = ("{0} GB" -f [Math]::Round($Win2.TotalVirtualMemorySize/1024/1024,0))})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Free Virtual Memory";'Value' = ("{0} GB" -f [Math]::Round($Win2.FreeVirtualMemory/1024/1024,2))})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Total Physical Memory";'Value' = ("{0} GB" -f [Math]::Round($Win2.TotalVisibleMemorySize/1024/1024,0))})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Free Physical Memory";'Value' = ("{0} GB" -f [Math]::Round($Win2.FreePhysicalMemory/1024/1024,2))})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Number Of NIC";'Value' = $Nics.count})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Install Date";'Value' = $Win2.InstallDate})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Last BootUp Time";'Value' = $Win2.LastBootUpTime})
	$OSInfoTable.Add([PSCustomObject]@{'Name' = "Serial Number";'Value' = $Win2.SerialNumber})


	# Logical Disk Space
	$DiskObj = Get-WmiObject win32_logicaldisk
	$DiskCnt = $DiskObj.Count

	$DiskInfoPies = foreach ($i in 0..($DiskCnt-1))
	{
		$DiskInfoPie = New-Object 'System.Collections.Generic.List[System.Object]'
		$obj = @()
		$obj += [PSCustomObject]@{
			'Name'  = "Free(GB)"
			'Count' = [math]::Round( ($DiskObj[$i].FreeSpace / 1024/ 1024/ 1024),2)
		}
		
		$obj += [PSCustomObject]@{
			'Name'  = "Used(GB)"
			'Count' = [math]::Round( (($DiskObj[$i].Size - $DiskObj[$i].FreeSpace)/ 1024/ 1024/ 1024) ,2)
		}
		
		$DiskInfoPie.Add($obj)
		$DiskInfoPie
	}

	$PieDiskLinks = foreach ($i in 0..($DiskCnt-1))
	{
		$PieDiskLink = Get-HTMLPieChartObject
		$PieDiskLink.Title = "{0}"  -f $DiskObj[$i].DeviceID
		$PieDiskLink.Size.Height = (1200 / $DiskCnt)
		$PieDiskLink.Size.width = (1200 / $DiskCnt)
		$PieDiskLink.ChartStyle.ChartType = 'doughnut'
		$PieDiskLink.ChartStyle.ColorSchemeName = "ColorScheme1" #"ColorScheme{0}" -f ($i%4)
		$PieDiskLink.ChartStyle.ColorSchemeName = "Generated1"  #"Generated{0}" -f ($i%8)
		#$PieDiskLink.ChartStyle.ColorSchemeName = 'Random'
		$PieDiskLink.DataDefinition.DataNameColumnName = 'Name'
		$PieDiskLink.DataDefinition.DataValueColumnName = 'Count'
		$PieDiskLink
	}


	# Logical Disk Space (Failed!!)
	$DiskWarnObj = $DiskObj | Where-Object { $_.FreeSpace / $_.Size -lt 0.5}
	$DiskWarnCnt = $DiskWarnObj.Count

	if ( $DiskWarnCnt -gt 0)
	{
	$DiskWarnInfoPies = foreach ($i in 0..($DiskWarnCnt-1))
	{
		$DiskInfoPie = New-Object 'System.Collections.Generic.List[System.Object]'
		$obj = @()
		$obj += [PSCustomObject]@{
			'Name'  = "Free(GB)"
			'Count' = [math]::Round( ($DiskWarnObj[$i].FreeSpace / 1024/ 1024/ 1024),2)
		}
		
		$obj += [PSCustomObject]@{
			'Name'  = "Used(GB)"
			'Count' = [math]::Round( (($DiskWarnObj[$i].Size - $DiskWarnObj[$i].FreeSpace)/ 1024/ 1024/ 1024) ,2)
		}
		
		$DiskInfoPie.Add($obj)
		$DiskInfoPie
	}

	$PieDiskWarnLinks = foreach ($i in 0..($DiskWarnCnt-1))
	{
		$PieDiskLink = Get-HTMLPieChartObject
		$PieDiskLink.Title = "{0}"  -f $DiskWarnObj[$i].DeviceID
		$PieDiskLink.Size.Height = 300
		$PieDiskLink.Size.width = 300
		$PieDiskLink.ChartStyle.ChartType = 'doughnut'
		$PieDiskLink.ChartStyle.ColorSchemeName = "ColorScheme{0}" -f ($i%4)
		$PieDiskLink.ChartStyle.ColorSchemeName = "Generated{0}" -f ($i%8)
		#$PieDiskLink.ChartStyle.ColorSchemeName = 'Random'
		$PieDiskLink.DataDefinition.DataNameColumnName = 'Name'
		$PieDiskLink.DataDefinition.DataValueColumnName = 'Count'
		$PieDiskLink
	}
	}

	$BarProgress.Value =  50

	# SQL Server Product Version
	$SQLVersionTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$SQLVersions = ExecuteSQL $SqlConn "EXEC master.dbo.xp_msver"

	foreach ($SQLVersion in $SQLVersions)
	{
	#	$Index = $SQLVersion.Index
		$Name = $SQLVersion.Name
		$Character_Value = $SQLVersion.Character_Value
		
		if ( !([string]::IsNullOrWhiteSpace($Character_Value)))
		{
		  $obj = [PSCustomObject]@{
			  'Name'  = $Name
			  'Value' = $Character_Value
		  }
		  $SQLVersionTable.Add($obj)
		}	
	}

	$SQLVersion2s = ExecuteSQL $SqlConn "select 'InstanceName' AS Name,SERVERPROPERTY('ServerName') as Value union all 
	select 'Edition' AS Name,SERVERPROPERTY('Edition')union all select 'ServicePack' AS Name,SERVERPROPERTY('Productlevel')union all 
	select 'IsCluster' AS Name,SERVERPROPERTY('IsClustered')union all 
	select 'SQLServerStartTime' AS Name,convert(varchar(30),sqlserver_start_time,120) from sys.dm_os_sys_info with(nolock)"

	foreach ($SQLVersion2 in $SQLVersion2s)
	{
		$Name = $SQLVersion2.Name
		$Value = $SQLVersion2.Value
		
		$obj = [PSCustomObject]@{
			'Name'  = $Name
			'Value' = $Value
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

	# Installed software
	$InstalledTable = New-Object 'System.Collections.Generic.List[System.Object]'

	Get-WmiObject -Class Win32_Product | Select-Object Name,Version,InstallDate,Description,InstallState -unique | ForEach-Object {
		$Name         = $_.Name
		$Version      = $_.Version
		$InstallDate  = $_.InstallDate
		$Description  = $_.Description
		$InstallState = $_.InstallState
		
		$obj = [PSCustomObject]@{
			'Name'         = $Name
			'Version'      = $Version
			'InstallDate'  = $InstallDate
			'Description'  = $Description
			'InstallState' = $InstallState
		}
		$InstalledTable.Add($obj)
	}

	if (($InstalledTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No sql service were found'
		}
		$InstalledTable.Add($obj)
	}

	# Get windows scheduled task
	$OSTaskTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$WinTasks = Get-ScheduledTask

	foreach ($WinTask in $WinTasks)
	{
		$State    = $WinTask.State
		$TaskPath = $WinTask.TaskPath
		$TaskName = $WinTask.TaskName
		
		$obj = [PSCustomObject]@{
			'State'    = $State
			'TaskPath' = $TaskPath
			'TaskName' = $TaskName
		}
		$OSTaskTable.Add($obj)
	}

	if (($OSTaskTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No scheduled task were found'
		}
		$OSTaskTable.Add($obj)
	}

	# SQL Server Configurations
	$SQLConfTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$SQLConfs = ExecuteSQL $SqlConn "SELECT configuration_id,name,value,minimum,maximum,value_in_use,description,is_dynamic,is_advanced FROM master.sys.configurations"

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


	# Connection Servers
	$ConnServerTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$ConnServers = ExecuteSQL $SqlConn "select name,product,provider,data_source,is_linked,is_system,is_publisher,is_distributor
	,is_subscriber,is_rpc_out_enabled,is_data_access_enabled,is_remote_login_enabled from master.sys.servers"

	foreach ($ConnServer in $ConnServers)
	{
		$name = $ConnServer.name
		$product = $ConnServer.product
		$provider = $ConnServer.provider
		$data_source = $ConnServer.data_source
		$is_linked = $ConnServer.is_linked
		$is_system = $ConnServer.is_system
		$is_publisher = $ConnServer.is_publisher
		$is_distributor = $ConnServer.is_distributor
		$is_subscriber = $ConnServer.is_subscriber
		$is_rpc_out_enabled = $ConnServer.is_rpc_out_enabled
		$is_data_access_enabled = $ConnServer.is_data_access_enabled
		$is_remote_login_enabled = $ConnServer.is_remote_login_enabled
		
		$obj = [PSCustomObject]@{
			'name'           = $name
			'product'        = $product
			'provider'       = $provider
			'data_source'    = $data_source
			'is_linked'      = $is_linked
			'is_system'      = $is_system
			'is_publisher'   = $is_publisher
			'is_distributor' = $is_distributor
			'is_subscriber'  = $is_subscriber
			'rpc_out_enabled'      = $is_rpc_out_enabled
			'data_access_enabled'  = $is_data_access_enabled
			'remote_login_enabled' = $is_remote_login_enabled
		}
		$ConnServerTable.Add($obj)
	}

	if (($ConnServerTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No server were found'
		}
		$ConnServerTable.Add($obj)
	}


	<################################ Instance ##################################>
	$BarProgress.Value =  55
	# database information / WARNING : Database
	$DatabaseInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$DatabaseInfos = ExecuteSQL $SqlConn "select database_id,name,compatibility_level,collation_name,user_access_desc as user_access
	,is_read_only,state_desc as state,recovery_model_desc as recovery_model
	,page_verify_option_desc as page_verify,create_date
	,(case when state<>0 then lower(state_desc)+';' else '' end)
	+(case when is_read_only=1 then 'read only;' else '' end)
	+(case when user_access=1 then 'single user;' else '' end)
	+(case when is_in_standby=1 then 'in standby;' else '' end)
	+(case when recovery_model=3 then 'recovery model is simple;' else '' end)
	+(case when page_verify_option=0 then 'page verify is null;' else '' end) as Warning
	from sys.databases with(nolock)"

	foreach ($DatabaseInfo in $DatabaseInfos)
	{
		$database_id     = $DatabaseInfo.database_id
		$name            = $DatabaseInfo.name
		$compatibility_level = $DatabaseInfo.compatibility_level
		$collation_name  = $DatabaseInfo.collation_name
		$user_access     = $DatabaseInfo.user_access
		$is_read_only    = $DatabaseInfo.is_read_only
		$state           = $DatabaseInfo.state
		$recovery_model  = $DatabaseInfo.recovery_model
		$page_verify     = $DatabaseInfo.page_verify
		$create_date     = $DatabaseInfo.create_date
		$warning         = $DatabaseInfo.warning
		
		$obj = [PSCustomObject]@{
			'DatabaseId'         = $database_id
			'DatabaseName'       = $name
			'CompatibilityLevel' = $compatibility_level
			'CollationName'      = $collation_name
			'UserAccess'         = $user_access
			'IsReadOnly'         = $is_read_only
			'State'              = $state
			'RecoveryModel'      = $recovery_model
			'PageVerify'         = $page_verify
			'CreateDate'         = $create_date
			'Warning'            = $warning
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


	# data file information
	$DatabaseFileInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$DatabaseFileInfos = ExecuteSQL $SqlConn "select database_id,(select b.name from sys.databases b with(nolock) where a.database_id=b.database_id) as database_name
	,type_desc,name,state_desc,size,max_size,growth,physical_name,is_percent_growth from sys.master_files a with(nolock)"

	foreach ($DatabaseFileInfo in $DatabaseFileInfos)
	{
		$database_id       = $DatabaseFileInfo.database_id
		$database_name     = $DatabaseFileInfo.database_name
		$name              = $DatabaseFileInfo.name
		$type_desc         = $DatabaseFileInfo.type_desc
		$state_desc        = $DatabaseFileInfo.state_desc
		$size              = $DatabaseFileInfo.size
		$max_size          = $DatabaseFileInfo.max_size
		$growth            = $DatabaseFileInfo.growth
		$physical_name     = $DatabaseFileInfo.physical_name
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

	# data file read vs write
	$DatabaseReadWriteTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$DatabaseReadWrites = ExecuteSQL $SqlConn "SELECT DB_NAME(fs.database_id) AS DatabaseName
	,SUM(io_stall_read_ms) AS wait_io_read_ms
	,SUM(io_stall_write_ms) AS wait_io_write_ms
	,SUM(num_of_reads) AS num_of_reads
	,SUM(num_of_writes) AS num_of_writes
	,SUM(io_stall_read_ms)/SUM(num_of_reads) AS avg_read_ms
	,SUM(io_stall_write_ms)/SUM(num_of_writes) AS avg_write_ms
	,CONVERT(DECIMAL(10,2),SUM(num_of_reads)*1.0/SUM(num_of_writes)) AS [num_of_read_vs_write]
	FROM sys.dm_io_virtual_file_stats(null,null) AS fs 
	INNER JOIN sys.master_files AS mf WITH (NOLOCK) 
	ON fs.database_id = mf.database_id AND fs.[file_id] = mf.[file_id] 
	GROUP BY fs.database_id
	HAVING SUM(io_stall_read_ms)>0 AND SUM(io_stall_write_ms)>0 AND SUM(num_of_writes)>0 AND SUM(io_stall_write_ms)>0
	ORDER BY DatabaseName"

	foreach ($DatabaseRW in $DatabaseReadWrites)
	{
		$DatabaseName         = $DatabaseRW.DatabaseName
		$wait_io_read_ms      = $DatabaseRW.wait_io_read_ms
		$wait_io_write_ms     = $DatabaseRW.wait_io_write_ms
		$num_of_reads         = $DatabaseRW.num_of_reads
		$num_of_writes        = $DatabaseRW.num_of_writes
		$avg_read_ms          = $DatabaseRW.avg_read_ms
		$avg_write_ms         = $DatabaseRW.avg_write_ms
		$num_of_read_vs_write = $DatabaseRW.num_of_read_vs_write

		$obj = [PSCustomObject]@{
			'DatabaseName'         = $DatabaseName
			'wait_io_read_ms'      = $wait_io_read_ms
			'wait_io_write_ms'     = $wait_io_write_ms
			'num_of_reads'         = $num_of_reads
			'num_of_writes'        = $num_of_writes
			'avg_read_ms'          = $avg_read_ms
			'avg_write_ms'         = $avg_write_ms
			'num_of_read_vs_write' = $num_of_read_vs_write
		}
		$DatabaseReadWriteTable.Add($obj)
	}

	if (($DatabaseReadWriteTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No database file were found'
		}
		$DatabaseReadWriteTable.Add($obj)
	}

	$BarProgress.Value =  60

	# Transaction Log
	$SQLTranTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$SQLTranLogs = ExecuteSQL $SqlConn "dbcc sqlperf(logspace)"

	foreach ($SQLTranLog in $SQLTranLogs)
	{
		$DatabaseName = $SQLTranLog.Item(0)
		$LogSize      = $SQLTranLog.Item(1)
		$LogSpaceUsed = $SQLTranLog.Item(2)
		$Status       = $SQLTranLog.Item(3)
		
		$obj = [PSCustomObject]@{
			'DatabaseName'    = $DatabaseName
			'LogSize_MB'      = [Math]::Round($LogSize,2) #ConvertSize($LogSize*1024*1024)
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
	# $SQLTranLogs = ExecuteSQL $SqlConn "dbcc sqlperf(logspace)"
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
	$SQLErrorLogs = ExecuteSQL $SqlConn "EXEC master.dbo.xp_enumerrorlogs 1"

	foreach ($SQLErrorLog in $SQLErrorLogs)
	{
		$file_id = $SQLErrorLog.Item(0)
		$date    = $SQLErrorLog.Item(1)
		$size    = $SQLErrorLog.Item(2)
		
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

	$BarProgress.Value =  65

	# Agent Log Size
	$AgentErrorLogTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$AgentErrorLogs = ExecuteSQL $SqlConn "EXEC master.dbo.xp_enumerrorlogs 2"

	foreach ($AgentErrorLog in $AgentErrorLogs)
	{
		$file_id = $AgentErrorLog.Item(0)
		$date    = $AgentErrorLog.Item(1)
		$size    = $AgentErrorLog.Item(2)
		
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
	$LoginNames = ExecuteSQL $SqlConn "SELECT name,create_date,type_desc,modify_date,default_database_name,is_disabled
	FROM master.sys.server_principals WITH(NOLOCK)
	WHERE type in ('U', 'G', 'S', 'C', 'K') AND principal_id not between 101 and 255 AND name <> N'##MS_AgentSigningCertificate##'"

	foreach ($LoginName in $LoginNames)
	{
		$name         = $LoginName.name
		$create_date  = $LoginName.create_date
		$type_desc    = $LoginName.type_desc
		$modify_date  = $LoginName.modify_date
		$default_database_name = $LoginName.default_database_name
		$is_disabled  = $LoginName.is_disabled
		
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
	$BarProgress.Value =  70
	# Jobs base information
	$JobsInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$JobsInfos = ExecuteSQL $SqlConn "select a.name,a.enabled,b.name as owneruser,a.description,a.date_created,a.date_modified,c.next_run_date,c.next_run_time
	from msdb.dbo.sysjobs a with(nolock) left join master.sys.syslogins b with(nolock) on a.owner_sid=b.sid
	left join msdb.dbo.sysjobschedules c with(nolock) on a.job_id=c.job_id"

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
	$JobsHisInfos = ExecuteSQL $SqlConn "select a.name,a.enabled,b.server,b.run_date,b.run_time,b.run_duration,b.run_status,b.message
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
			'RunStatus'  = $run_status
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


	# suspect_pages
	$SuspectPagesTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$SuspectPages = ExecuteSQL $SqlConn "select DB_NAME(database_id) as DatabaseName,* from msdb.dbo.suspect_pages"

	foreach ($SuspectPage in $SuspectPages)
	{
		$DatabaseId     = $SuspectPage.database_id
		$DatabaseName   = $SuspectPage.DatabaseName
		$FileId         = $SuspectPage.file_id
		$PageId         = $SuspectPage.page_id
		$EventType      = $SuspectPage.event_type
		$ErrorCount     = $SuspectPage.error_count
		$LastUpdateDate = $SuspectPage.last_update_date


		$obj = [PSCustomObject]@{
			'DatabaseId'     = $DatabaseId
			'DatabaseName'   = $DatabaseName
			'FileId'         = $FileId
			'PageId'         = $PageId
			'EventType'      = $EventType
			'ErrorCount'     = $ErrorCount
			'LastUpdateDate' = $LastUpdateDate
		}
		$SuspectPagesTable.Add($obj)
	}

	if (($SuspectPagesTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No jobs history were found'
		}
		$SuspectPagesTable.Add($obj)
	}

	<################################ Backup & Restore ##################################>
	$BarProgress.Value =  75

	# Backup Database Info
	$BackupInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$BackupInfos = ExecuteSQL $SqlConn "SELECT K.name
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
			'DatabaseName'   = $name
			'FullBackupTime' = $latest_full_backup_time
			'DiffBackupTime' = $latest_diff_backup_time
			'LogBackupTime'  = $latest_log_backup_time
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

	$DBAll = ExecuteSQL $SqlConn "select count(name) cnt from master.sys.databases with(nolock) where database_id<>2"
	$DBBak = ExecuteSQL $SqlConn "select count(name) cnt from master.sys.databases a with(nolock) 
	inner join (select distinct database_name from msdb.dbo.backupset with(nolock))b on a.name=b.database_name
	where a.database_id<>2"

	$obj1 = [PSCustomObject]@{
		'Name'  = "从未备份"
		'Count' = $DBAll.cnt-$DBBak.cnt
	}

	$BackupInfoPie.Add($obj1)

	$obj2 = [PSCustomObject]@{
		'Name'  = "曾有备份"
		'Count' = $DBBak.cnt
	}

	$BackupInfoPie.Add($obj2)

	$PieBackupLinks = Get-HTMLPieChartObject
	$PieBackupLinks.Title = "历史备份情况"
	$PieBackupLinks.Size.Height = 450
	$PieBackupLinks.Size.width = 300
	$PieBackupLinks.ChartStyle.ChartType = 'doughnut'
	$PieBackupLinks.ChartStyle.ColorSchemeName = "ColorScheme4"
	$PieBackupLinks.ChartStyle.ColorSchemeName = "Generated5"
	#$PieBackupLinks.ChartStyle.ColorSchemeName = 'Random'
	$PieBackupLinks.DataDefinition.DataNameColumnName = 'Name'
	$PieBackupLinks.DataDefinition.DataValueColumnName = 'Count'

	$BarProgress.Value =  80

	###############################
	$BackupInfoPie7 = New-Object 'System.Collections.Generic.List[System.Object]'
	$DBBak7 = ExecuteSQL $SqlConn "select count(distinct database_name) cnt from msdb.dbo.backupset with(nolock) where backup_start_date >= dateadd(d,-7,getdate())"

	$obj1 = [PSCustomObject]@{
		'Name'  = "未备份"
		'Count' = $DBAll.cnt-$DBBak7.cnt
	}

	$BackupInfoPie7.Add($obj1)

	$obj2 = [PSCustomObject]@{
		'Name'  = "已备份"
		'Count' = $DBBak7.cnt
	}

	$BackupInfoPie7.Add($obj2)

	$PieBackupLinks7 = Get-HTMLPieChartObject
	$PieBackupLinks7.Title = "最近7天备份情况"
	$PieBackupLinks7.Size.Height = 450
	$PieBackupLinks7.Size.width = 300
	$PieBackupLinks7.ChartStyle.ChartType = 'doughnut'
	$PieBackupLinks7.ChartStyle.ColorSchemeName = "ColorScheme4"
	$PieBackupLinks7.ChartStyle.ColorSchemeName = "Generated5"
	#$PieBackupLinks7.ChartStyle.ColorSchemeName = 'Random'
	$PieBackupLinks7.DataDefinition.DataNameColumnName = 'Name'
	$PieBackupLinks7.DataDefinition.DataValueColumnName = 'Count'


	# The latest Backup Information
	$BackupDetailTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$BackupDetailInfos = ExecuteSQL $SqlConn "select a.database_name 
	,(case when a.type='D' then 'Full' when a.type='I' then 'Diff' when a.type='L' then 'Log' else '' end) as type
	,CONVERT(varchar(30),a.backup_start_date,120) AS latest_backup_time ,compressed_backup_size as backup_size
	,DATEDIFF(MINUTE,backup_start_date,backup_finish_date) as backup_time_Min  
	,REVERSE(stuff(REVERSE(physical_device_name),1,CHARINDEX('\',REVERSE(physical_device_name))-1,'')) as backup_path  
	from msdb.dbo.backupset a with(nolock)
	inner join sys.databases t with(nolock) on a.database_name=t.name
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
			'DatabaseName' = $database_name
			'BackupType'   = $type
			'LatestBackupTime'   = $latest_backup_time
			'Size'         = ConvertSize($backup_size)
			'TimeCost_Min' = $backup_time_Min
			'BackupPath'   = $backup_path
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
	$RestoreInfos = ExecuteSQL $SqlConn "select user_name,destination_database_name
	,(case when restore_type='D' then 'Full' when restore_type='I' then 'Diff' when restore_type='L' then 'Log' else '' end) as restore_type
	,restore_date,replace,recovery,restart,stop_at,file_number,destination_phys_name
	from msdb.dbo.restorehistory a with(nolock) 
	inner join msdb.dbo.restorefile b with(nolock) on a.restore_history_id=b.restore_history_id
	inner join sys.databases t with(nolock) on a.destination_database_name=t.name"

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

	<################################ Perfmon ##################################>
	$BarProgress.Value =  85

	# Database buffer pool allocation
	$DBBufferpoolTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$DBBufferpoolInfos = ExecuteSQL $SqlConn "SELECT CASE database_id WHEN 32767 THEN 'ResourceDb' ELSE db_name(database_id) END AS Database_name  
	,count(*) AS cached_pages_count ,count(*)*convert(BIGINT,8192) AS cached_space_in_bytes ,sum(free_space_in_bytes) AS free_space_in_bytes  
	FROM sys.dm_os_buffer_descriptors(nolock) GROUP BY database_id"

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
	$DBDirtyPages = ExecuteSQL $SqlConn "SELECT db_name(database_id) AS DatabaseName,count(page_id) AS DirtyPages,count(page_id)*8192 AS DirtyPages_bytes
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

	# Thread waiting
	$ThreadWaitTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$ThreadWaits = ExecuteSQL $SqlConn "select wait_type,waiting_tasks_count,wait_time_ms,wait_time_ms/waiting_tasks_count as avg_wait_time_ms
	,signal_wait_time_ms,wait_time_ms-signal_wait_time_ms as run_time_ms,max_wait_time_ms
	from sys.dm_os_wait_stats 
	where waiting_tasks_count>0 and wait_time_ms>0 and max_wait_time_ms>0 and signal_wait_time_ms>0
	order by avg_wait_time_ms desc"

	foreach ($ThreadWait in $ThreadWaits)
	{
		$WaitType          = $ThreadWait.wait_type
		$WaitingTasksCount = $ThreadWait.waiting_tasks_count
		$WaitTime_ms       = $ThreadWait.wait_time_ms
		$AvgWaitTime_ms    = $ThreadWait.avg_wait_time_ms
		$SignalWaitTime_ms = $ThreadWait.signal_wait_time_ms
		$RunTime_ms        = $ThreadWait.run_time_ms
		$MaxWaitTime_ms    = $ThreadWait.max_wait_time_ms
		
		$obj = [PSCustomObject]@{
			'WaitType'          = $WaitType
			'WaitingTasksCount' = $WaitingTasksCount
			'WaitTime_ms'       = $WaitTime_ms
			'AvgWaitTime_ms'    = $AvgWaitTime_ms
			'SignalWaitTime_ms' = $SignalWaitTime_ms
			'RunTime_ms'        = $RunTime_ms
			'MaxWaitTime_ms'    = $MaxWaitTime_ms
		}
		$ThreadWaitTable.Add($obj)
	}

	if (($ThreadWaitTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No thread waitting were found'
		}
		$ThreadWaitTable.Add($obj)
	}

	$BarProgress.Value =  90

	# CPU usage for sql
	$CPUSQLTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$CPUSQLType  = ExecuteSQL $SqlConn "select DB_NAME(dbid) dbname,total_cpu_time,total_execution_count,avg_cpu_time,number_of_statements,[text]
	from (
		select top 50 
		sum(qs.total_worker_time) as total_cpu_time,  
		sum(qs.execution_count) as total_execution_count,
		sum(qs.total_worker_time)/sum(qs.execution_count) as  avg_cpu_time,
		count(*) as  number_of_statements, 
		qs.plan_handle 
		from sys.dm_exec_query_stats qs
		group by qs.plan_handle
		order by avg_cpu_time desc
	) eqs cross apply sys.dm_exec_sql_text(eqs.plan_handle) as est"

	foreach ($CPUSQLType1 in $CPUSQLType)
	{
		$DatabaseName         = $CPUSQLType1.dbname
		$TotalCpuTime         = $CPUSQLType1.total_cpu_time
		$TotalExecutionCount  = $CPUSQLType1.total_execution_count
		$AvgCpuTime           = $CPUSQLType1.avg_cpu_time
		$NumberOfStatements   = $CPUSQLType1.number_of_statements
		$SqlText              = $CPUSQLType1.text
		
		$obj = [PSCustomObject]@{
			'DatabaseName'         = $DatabaseName
			'TotalCpuTime'         = $TotalCpuTime
			'TotalExecutionCount'  = $TotalExecutionCount
			'AvgCpuTime'           = $AvgCpuTime
			'NumberOfStatements'   = $NumberOfStatements
			'SqlText'              = $SqlText
		}
		$CPUSQLTable.Add($obj)
	}

	if (($CPUSQLTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No sql were found'
		}
		$CPUSQLTable.Add($obj)
	}

	# IO usage for sql
	$IOSQLTable = New-Object 'System.Collections.Generic.List[System.Object]'
	$IOSQLType  = ExecuteSQL $SqlConn "select DB_NAME(dbid) dbname,total_logical_reads,total_physical_reads,total_execution_count,avg_logical_reads,number_of_statements,[text]
	from (
		select top 50 
		sum(qs.total_logical_reads) as total_logical_reads,
		sum(qs.total_physical_reads) as total_physical_reads,
		sum(qs.execution_count) as total_execution_count,
		sum(qs.total_logical_reads)/sum(qs.execution_count) as  avg_logical_reads,
		count(*) as  number_of_statements, 
		qs.plan_handle 
		from sys.dm_exec_query_stats qs
		group by qs.plan_handle
		order by avg_logical_reads desc
	) eqs cross apply sys.dm_exec_sql_text(eqs.plan_handle) as est"

	foreach ($IOQLType1 in $IOSQLType)
	{
		$DatabaseName         = $IOQLType1.dbname
		$TotalLogicalReads    = $IOQLType1.total_logical_reads
		$TotalPhysicalReads   = $IOQLType1.total_physical_reads
		$TotalExecutionCount  = $IOQLType1.total_execution_count
		$AvgLogicalReads      = $IOQLType1.avg_logical_reads
		$NumberOfStatements   = $IOQLType1.number_of_statements
		$SqlText              = $IOQLType1.text
		
		$obj = [PSCustomObject]@{
			'DatabaseName'         = $DatabaseName
			'TotalLogicalReads'    = $TotalLogicalReads
			'TotalPhysicalReads'   = $TotalPhysicalReads
			'TotalExecutionCount'  = $TotalExecutionCount
			'AvgLogicalReads'      = $AvgLogicalReads
			'NumberOfStatements'   = $NumberOfStatements
			'SqlText'              = $SqlText
		}
		$IOSQLTable.Add($obj)
	}

	if (($IOSQLTable).Count -eq 0)
	{
		$Obj = [PSCustomObject]@{
			Information = 'Information: No sql were found'
		}
		$IOSQLTable.Add($obj)
	}

	<#####################################################################################
	# HTML Data 
	#####################################################################################>
	$BarProgress.Value =  95
	Write-Host "Compiling Report..." -ForegroundColor Green

	CloseSqlConnection $SqlConn
	$Options = @('问题仪表盘','  系  统  ',' SQLServer ','代理作业','备份恢复','  性  能  ')
	$MSSQLReport = New-Object 'System.Collections.Generic.List[System.Object]'
	$MSSQLReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $LeftLogo -RightLogoString $RightLogo ))
	$MSSQLReport.Add($(Get-HTMLTabHeader -TabNames $Options))


	<# Problem Dashboard #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[0] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Error or Warning"))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "系统"))
	$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $WarningSystem -Yellow "1 -eq 1") -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server"))
	$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $WarningInstance -Yellow "1 -eq 1") -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "性能优化"))
	$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $Perfmon -Yellow "1 -eq 1") -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))


	if ( $MSSQLErrorLog.Count -gt 0)
	{
	  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server 错误日志（7天内错误信息）"))
	  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $MSSQLErrorLog -Red "1 -eq 1") -HideFooter))
	  $MSSQLReport.Add($(Get-HTMLContentClose))
	}

	if ( $WarningEventLog.Count -gt 0)
	{
	  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "系统事件（7天内最新的20个错误或警告事件）"))
	  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $WarningEventLog -Red '$this.EntryType -eq "Error"' -Yellow '$this.EntryType -eq "Warning"') -HideFooter))
	  $MSSQLReport.Add($(Get-HTMLContentClose))
	}

	if ( $WarningApplicationLog.Count -gt 0)
	{
	  $MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "应用程序事件（7天内最新的20个错误或警告事件）"))
	  $MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $WarningApplicationLog -Red '$this.EntryType -eq "Error"' -Yellow '$this.EntryType -eq "Warning"') -HideFooter))
	  $MSSQLReport.Add($(Get-HTMLContentClose))

	}

	$MSSQLReport.Add($(Get-HTMLTabContentClose))


	<# System option #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[1] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "OS Product & SQL Server Product"))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "操作系统信息"))
	$MSSQLReport.Add($(Get-HTMLContentTable $OSInfoTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server 信息"))
	$MSSQLReport.Add($(Get-HTMLContentTable $SQLVersionTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "磁盘使用情况"))
	foreach ($i in 0..($DiskCnt-1))
	{
	  $MSSQLReport.Add($(Get-HTMLColumnOpen -ColumnNumber ($i+1) -ColumnCount $DiskCnt))
	  $MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieDiskLinks[$i] -DataSet $DiskInfoPies[$i]))
	  $MSSQLReport.Add($(Get-HTMLColumnClose))
	}
	$MSSQLReport.Add($(Get-HTMLContentClose))

	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server 服务状态"))
	$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $SQLServiceTable -Green '$this.Status -eq "Running"') -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Windows 已安装软件"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $InstalledTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Windows 任务"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $OSTaskTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLTabContentClose))


	<# Instance option #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[2] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据库信息"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $DatabaseInfoTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据文件信息"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $DatabaseFileInfoTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据库文件读写情况"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $DatabaseReadWriteTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Transaction Log"))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "事务日志使用情况"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLTranTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "事务日志使用对比"))
	$MSSQLReport.Add($(Get-HTMLBarChart -ChartObject $SQLTranLogBarChar -DataSet $SQLTranLogs))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Error Log"))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server 错误日志大小"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLErrorLogTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "SQL Server 代理日志大小"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $AgentErrorLogTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "账号信息"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $LoginNameTable -HideFooter))
	#$MSSQLReport.Add($(Get-HTMLContentTable $(Set-TableRowColor $LoginNameTable -Yellow '$this.IsDisabled -eq 1') -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "配置信息"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $SQLConfTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "链接服务器"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $ConnServerTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLTabContentClose))


	<# Jobs option #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[3] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "作业信息"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $JobsInfoTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "作业执行历史"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $JobsHistoryTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "可能损坏的数据页"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $SuspectPagesTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLTabContentClose))

	<# Backup & Restore option #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[4] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Database Backup"))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "最近备份记录"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $BackupInfoTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据库备份情况占比"))
	$MSSQLReport.Add($(Get-HTMLColumnOpen -ColumnNumber (1) -ColumnCount 2))
	$MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieBackupLinks -DataSet $BackupInfoPie))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumnOpen -ColumnNumber (2) -ColumnCount 2))
	$MSSQLReport.Add($(Get-HTMLPieChart -ChartObject $PieBackupLinks7 -DataSet $BackupInfoPie7))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "备份历史记录"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable ($BackupDetailTable | Sort LatestBackupTime -Descending) -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "还原历史记录"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable ($RestoreTable | Sort RestoreDate -Descending) -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLTabContentClose))


	<# Perfmon option -IsHidden #>
	$MSSQLReport.Add($(Get-HTMLTabContentopen -TabName $Options[5] -TabHeading ("Local Default Instance,Report created on: " + (Get-Date -Format yyyy-MM-dd))))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "Buffer allocation" ))
	$MSSQLReport.Add($(Get-HTMLColumn1of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据库缓存使用情况"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $DBBufferpoolTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLColumn2of2))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "数据库脏页情况"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable $DBDirtyPageTable -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLColumnClose))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "线程等待统计"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable ($ThreadWaitTable | Sort AvgCpuTime -Descending) -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "最耗 CPU/Time 的 SQL 语句(TOP 50)"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable ($CPUSQLTable | Sort AvgCpuTime -Descending) -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLContentOpen -HeaderText "最耗 IO 的 SQL 语句(TOP 50)"))
	$MSSQLReport.Add($(Get-HTMLContentDataTable ($IOSQLTable | Sort AvgLogicalReads -Descending) -HideFooter))
	$MSSQLReport.Add($(Get-HTMLContentClose))
	$MSSQLReport.Add($(Get-HTMLTabContentClose))

	# Close HTML Page
	$MSSQLReport.Add($(Get-HTMLClosePage))

	$BarProgress.Value =  96
	Write-Host "Showing Report.`n" -ForegroundColor Green

	#Save-HTMLReport -ShowReport -ReportContent $MSSQLReport -ReportName $ReportName -ReportPath $ReportSavePath
	Save-HTMLReport -ReportContent $MSSQLReport -ReportName $ReportName -ReportPath $ReportSavePath
	$BarProgress.Value =  98
	$HTMLFile = Join-Path -Path $ReportSavePath -Childpath ($ReportName + ".html")
	
	(Get-Content $HTMLFile -Encoding Default) | Set-Content $HTMLFile -Encoding Unicode
	$BarProgress.Value =  100
	Invoke-Expression $HTMLFile 
}

$inputForm.ShowDialog()

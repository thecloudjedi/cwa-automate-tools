function FetchSQLResults ($Query)	{


    $SQLResult = ."C:\Program Files (x86)\LabTech\mysql\bin\mysql.exe" --database=Labtech --user=root -e $($Query) --batch --raw -N
              		
              if (!$SqlResult) 
		{
		return "Unable to Detect (NULL)";
		}
		else 
		{
		return $SQLResult;
		}  
		
	}


# Application Whitelist #
$productlist = (get-wmiobject -Class Win32_Product)
$productlistfiltered = $productlist|?{$_.name -notlike 'amazon*' -and $_.name -notlike 'labtech*' -and $_.name -notlike 'microsoft*' -and $_.name -notlike 'mysql*' -and $_.name -notlike 'sql*' -and $_.name -notlike 'screenconnect*'}
$whitelist = @('7-Zip 9.20','7-Zip 9.20 (x64 edition)','802.11 USB Wireless LAN Adapter','Adobe Flash Player 11 ActiveX','Adobe Flash Player 12 ActiveX','Adobe Flash Player 16 ActiveX','Adobe Flash Player 17 ActiveX','Amazon SSM Agent','Amazon SSM Agent?','Amazon SSM Agentp','Automate Control Center','AWS Agent','AWS PV Drivers','AWS Tools for Windows','aws-cfn-bootstrap','Citrix Tools for Virtual Machines','ConnectWise Internet Client','ConnectWise Internet Client 64-bit','ConnectWise Manage Outlook Add-in','Crystal Reports 12.0 for LabTech','Crystal Reports 2008 Runtime SP2','Debug Diagnostics 1.2','Debugging Tools for Windows (x86)','Dropbox','Dropbox Setup','EC2ConfigService','EC2ConfigServicep','Google Chrome','hMailServer 5.3.3-B1879','hMailServer 5.4-B1950','hMailServer 5.4.1-B1951','hMailServer 5.6-B2145','hMailServer 5.6.6-B2383','IIS URL Rewrite Module 2','Java 7 Update 25','NVIDIA 3D Vision Controller Driver 276.42','NVIDIA 3D Vision Driver 276.52','NVIDIA Graphics Driver 276.52','NVIDIA nView 136.18','NVIDIA WMI 276.52','Red Hat Paravirtualized Xen Drivers for Windows(R)','Webroot SecureAnywhere','Windows Driver Package - RedHat (rhelscsi) SCSIAda','Windows Resource Kit Tools')


# Installing Word Module #

Install-Module -Name PSWriteWord -Force

Import-Module PSWriteWord


$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-CreateWord1.docx"


### define new document
$WordDocument = New-WordDocument $FilePath -Verbose
### Header ###
Add-WordText -WordDocument $WordDocument -Text "Cloud Server Performance Checklist" -FontFamily "calibri light" -FontSize 18 -Supress $True -Color darkBlue -Bold $True -HeadingType Heading3
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordText -WordDocument $WordDocument -Text "Help for this doc can be found at https://connectwise.sharepoint.com/SvcEdu/AutomateSupport/Pages/Cloud-Server-Performance-Check-Help.aspx" -FontFamily "Calibri" -FontSize 11 -Supress $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
### Issue Description ###
Add-WordText -WordDocument $WordDocument -Text "Issue Description:" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
### Troubleshooting Steps ###
Add-WordText -WordDocument $WordDocument -Text "Troubleshooting Steps Taken:" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
### Worst Performance Times ###
Add-WordText -WordDocument $WordDocument -Text "What time(s) is the partner seeing the worst performance?:" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
### Next Steps ###
Add-WordText -WordDocument $WordDocument -Text "Recommended Next Steps:" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
#Add-WordPageBreak -WordDocument $WordDocument -Verbose

### Partner Info ###
Add-WordText -WordDocument $WordDocument -Text "Partner Info" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $SQL = fetchsqlresults -Query "SELECT CONCAT((SELECT COUNT(*) FROM computers WHERE lastcontact > NOW()-INTERVAL 10 MINUTE),'/',(SELECT COUNT(*) FROM computers),' Online')"
Add-WordText -WordDocument $WordDocument -Text "Agent Count: ", $SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $PS = invoke-restmethod "http://169.254.169.254/latest/meta-data/placement/availability-zone"
Add-WordText -WordDocument $WordDocument -Text "Automate Server Region: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $SQL = FetchSQLResults -Query "SELECT serveraddress FROM templates WHERE NAME = 'default'"
Add-WordText -WordDocument $WordDocument -Text "Automate Server FQDN: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### Server Info ###
Add-WordText -WordDocument $WordDocument -Text "Server Info" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $PS = (gcim Win32_OperatingSystem).caption
Add-WordText -WordDocument $WordDocument -Text "OS Ver: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (Get-WmiObject –class Win32_processor).NumberOfCores
Add-WordText -WordDocument $WordDocument -Text "CPU/Cores: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (get-wmiobject -class "win32_physicalmemory").capacity/1GB
Add-WordText -WordDocument $WordDocument -Text "RAM: ","$PS GB" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $SQL = FetchSQLResults -Query "SELECT CONCAT(majorversion,'.',minorversion) FROM config"
Add-WordText -WordDocument $WordDocument -Text "Automate Version: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordText -WordDocument $WordDocument -Text "Have there been any recent changes to the server? If so list the changes:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line


### Pre-Req Check ###

Add-WordText -WordDocument $WordDocument -Text "Are the prerequisites in place?" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $SQL = FetchSQLResults -Query "SELECT @@version"
Add-WordText -WordDocument $WordDocument -Text "MySQL Version: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = ($productlist|? {$_.name -match 'mysql connector net'}).version
Add-WordText -WordDocument $WordDocument -Text "Current version of MySQL Connector/NET Version: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = ($productlist|? {$_.name -match 'visual C+'}|select name,version|ForEach-Object {"$($_.name) = $($_.version)"}) -join "  |  "
Add-WordText -WordDocument $WordDocument -Text "Current version of Microsoft Visual C++ Redistributable Installed: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (gci "C:\Program Files (x86)\MySQL"|? {$_.name -match 'odbc'}).Name
Add-WordText -WordDocument $WordDocument -Text "Current version of MySQL Connector/ODBC 32 bit Installed: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (gci "C:\Program Files\MySQL"|? {$_.name -match 'odbc'}).Name
Add-WordText -WordDocument $WordDocument -Text "Current version of MySQL Connector/ODBC 64 bit Installed: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black-bold $true
    $PS = ((get-windowsfeature|Where-Object {$_. displayname -like '*asp*'}|Select-Object displayname,installed)|ForEach-Object {"$($_.displayname) = $($_.installed)"}) -join "  |  "
Add-WordText -WordDocument $WordDocument -Text "Is ASP & ASP.Net Installed: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### Installed Software ###
Add-WordText -WordDocument $WordDocument -Text "Installed Software" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $PS = ($productlist|? {$_.name -match 'defend' -or $_.name -match 'secure'}|select name,version|ForEach-Object {"$($_.name) = $($_.version)"}) -join "  |  "
Add-WordText -WordDocument $WordDocument -Text "AV Software Installed: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = ($productlistfiltered|? {$_.name -notin $whitelist}|select name,version|ForEach-Object {"$($_.name) = $($_.version)"}) -join "  |  "
Add-WordText -WordDocument $WordDocument -Text "Other Possible Software Conflicts: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### IIS Application Pool Settings ###
Add-WordText -WordDocument $WordDocument -Text "IIS Application Pool Settings" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordText -WordDocument $WordDocument -Text "Open IIS Manager-> Expand the server on the left hand pane-> Click on Application Pools underneath the server-> " -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordText -WordDocument $WordDocument -Text "Verify that the following settings are set for the Labtech, Labtech WebCC and Labtech Mobile app pools " -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    import-module webadministration
    $PS = (Get-ItemProperty "IIS:\AppPools\CwaRestApi").enable32BitAppOnWin64
Add-WordText -WordDocument $WordDocument -Text "Are 32 bit Apps Enabled: (Should be set to True) : ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (Get-ItemProperty "IIS:\AppPools\CwaRestApi"|select -ExpandProperty cpu).limit
Add-WordText -WordDocument $WordDocument -Text "Limit Interval (minutes) (Shoudl be set to 0): ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (Get-ItemProperty "IIS:\AppPools\CwaRestApi").managedRuntimeVersion
Add-WordText -WordDocument $WordDocument -Text ".Net Framework version (Should be set to v4.0): ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (Get-ItemProperty "IIS:\AppPools\CwaRestApi").managedpipelinemode
Add-WordText -WordDocument $WordDocument -Text "IIs Managed Pipeline Mode (Should be set to Integrated) : ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### Plugins ###
Add-WordText -WordDocument $WordDocument -Text "Plugins" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $SQL = FetchSQLResults -Query "SELECT COUNT(*) FROM plugins WHERE enable=1"
Add-WordText -WordDocument $WordDocument -Text "How Many Plugins are Enabled: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $SQL = FetchSQLResults -Query "SELECT COUNT(*) FROM plugins WHERE enable=1 AND official = 0"
Add-WordText -WordDocument $WordDocument -Text "How many of those plugins are 3rd Party Plugins: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "Have we tried disabling the plugins to see if the issue persists? If so what happened?:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### Database Info ###
Add-WordText -WordDocument $WordDocument -Text "Database Info" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $SQL = FetchSQLResults -Query "SELECT ROUND(SUM(data_length + index_length) / 1024 / 1024 / 1024, 2) AS 'Size' FROM information_schema.TABLES WHERE table_schema = 'labtech' GROUP BY table_schema"
Add-WordText -WordDocument $WordDocument -Text "LT DB size = ",$SQL GB -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $SQL = FetchSQLResults -Query "SELECT GROUP_CONCAT(tablename) FROM (SELECT CONCAT(TABLE_NAME,'= ',ROUND((data_length + index_length) / 1024 / 1024 , 2),' MB') AS 'tablename' FROM information_schema.TABLES WHERE table_schema = 'labtech' ORDER BY (data_length + index_length) DESC LIMIT 10) AS a "
Add-WordText -WordDocument $WordDocument -Text "What are the Largest Tables: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "Is my.ini configured Properly:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $SQL = FetchSQLResults -Query "SELECT GROUP_CONCAT(CONCAT(HOST,',',command,',',TIME,',',state,',',info) SEPARATOR '\n') FROM INFORMATION_SCHEMA.PROCESSLIST WHERE command <> 'sleep' AND TIME > 5"
Add-WordText -WordDocument $WordDocument -Text "What is running in SHOW FULL PROCESSLIST? (> 5 Seconds): ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordText -WordDocument $WordDocument -Text "CheckTables Result:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordText -WordDocument $WordDocument -Text "Disk queue length during Checktables:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
    $PS = (new-timespan -start (get-item 'C:\Program Files\LabTech\Backup\Tablebase\additionalschedules.sql').LastWriteTime -End (get-item 'C:\Program Files\LabTech\Backup\Tablebase\webdashboards.sql').LastWriteTime).Minutes
Add-WordText -WordDocument $WordDocument -Text "Backup Time: ","$PS Minutes" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

### Max Port Setting ###
Add-WordText -WordDocument $WordDocument -Text "MaxUserPort Registry Setting" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $PS = (get-itemproperty "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters").maxuserport
Add-WordText -WordDocument $WordDocument -Text "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\ParametersMaxUserPort Value: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line


### Are the script settings optimized? ###
Add-WordText -WordDocument $WordDocument -Text "Are the script settings optimized?" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $SQL= FetchSQLResults -Query "SELECT COUNT(*) FROM runningscripts WHERE running=1"
Add-WordText -WordDocument $WordDocument -Text "Number of currently running Scripts: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $SQL= FetchSQLResults -Query "SELECT value FROM properties WHERE name = 'MaxRunningSCripts'"
Add-WordText -WordDocument $WordDocument -Text "Max running scripts: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $SQL= FetchSQLResults -Query "SELECT value FROM properties WHERE name = 'ScriptsAtATime'"
Add-WordText -WordDocument $WordDocument -Text "Scripts per batch: ",$SQL -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "What scripts are running during periods of performance issues?:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "Are there any notable long running scripts?:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line


### Resource Utilization ###
Add-WordText -WordDocument $WordDocument -Text "Resource Utilization" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
    $PS = (Get-WmiObject win32_processor).loadpercentage
Add-WordText -WordDocument $WordDocument -Text "Current CPU Usage: ","$PS %" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
#Add-WordText -WordDocument $WordDocument -Text "CPU Usage Last 7 Days:" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black 
    $PS = Get-Ciminstance Win32_OperatingSystem|select freephysicalmemory, totalvisiblememorysize|foreach-object {"$($_.freephysicalmemory) MB/$($_.TotalVisibleMemorySize) MB"}
Add-WordText -WordDocument $WordDocument -Text "Current RAM Usage: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $PS = (Get-WmiObject -Class Win32_perfformatteddata_perfdisk_LogicalDisk |Select-Object name, CurrentDiskQueueLength|ForEach-Object {"$($_.name) = $($_.currentdiskqueuelength)"}) -join "  |  "
Add-WordText -WordDocument $WordDocument -Text "Current Disk Queue Length: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
    $PS = (get-WmiObject win32_logicaldisk|select deviceid,freespace,size|?{$_.deviceid -ne 'L:'}|ForEach-Object {"$($_.deviceid) = $([math]::round($_.freespace/1GB,2)) GB/$([math]::round($_.size/1GB,2)) GB Free"})
Add-WordText -WordDocument $WordDocument -Text "Current Disk Space: ",$PS -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "Current Disk Defragmentation: " -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordText -WordDocument $WordDocument -Text "Run a latency check via cloudping.info from their local environment. Does it show high ping rate in their region?" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black 
Add-WordText -WordDocument $WordDocument -Text "Is another regions ping lower?" -FontFamily "calibri light" -FontSize 10 -Supress $True -Color black  -bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line


### CloudWatch Stats ###

$instanceid = Invoke-RestMethod "http://169.254.169.254/latest/meta-data/instance-id"
$region = (Invoke-RestMethod "http://169.254.169.254/latest/meta-data/placement/availability-zone") -replace ".$"
$offset = (get-date -Format "zz") -replace "0",""
Import-Module awspowershell
Set-DefaultAWSRegion $region


$interval= '900';
$split= '24';  
$time= Get-Date;
$starttime= ($time.AddHours(-$split)).ToUniversalTime();
$endtime= $time.ToUniversalTime();

# The Metrics to pull for EC2 #
$ec2_1 = @("CPUCreditBalance","Count")
$ec2_2 = @("NetworkIn","Bytes")
$ec2_3 = @("StatusCheckFailed","Count")
$ec2_4 = @("CPUUtilization","Percent")
$ec2_5 = @("NetworkOut","Bytes")
$ec2metriclist = @($ec2_1,$ec2_2,$ec2_3,$ec2_4,$ec2_5)

$Dimension1 = New-Object 'Amazon.CloudWatch.Model.Dimension';
$Dimension1.Name = 'InstanceId';
$Dimension1.Value = $instanceid;
$namespace = "AWS/EC2"

foreach($metric in $ec2metriclist)
{
    $metricname = $metric[0]
### The api call that gathers the data ###
    $datapoints=Get-CWMetricStatistics -Namespace $namespace -MetricName $metric[0] -Dimensions $dimension1 -EndTime $endtime -Period $interval -StartTime $starttime -Statistics 'Minimum','Maximum','Average','Sum','SampleCount' -Unit $metric[1]|select -expandproperty datapoints|sort-object timestamp;
if($datapoints) 
    {$datapoints|Add-Member -MemberType NoteProperty -Name InstanceID -Value $instanceid
    $datapoints|Add-Member -MemberType NoteProperty -Name MetricName -Value $metric[0]
    $datapoints|Add-Member -MemberType NoteProperty -Name Interval -Value $interval  
    $datalist = $datapoints|select instanceid, metricname, minimum, maximum, average, unit, timestamp, interval
### Displaying the data. Here is where you can send it to a log ###
    #write-output $datalist

$simpleDataset = [ordered]@{}
$dates = @()
$values = @()

foreach($datapoint in $datalist)
{ $key = (get-date ($datapoint.timestamp).AddHours($offset) -Format g)
  $value = $datapoint.average 
  $simpledataset.add($key,$value)
}

foreach($set in $simpleDataset)
{$dates += $set.keys
 $values += $set.Values
}

Add-WordLineChart -WordDocument $WordDocument -ChartName "$metricname" -Names $dates -Values $values -ChartLegendOverlay $false -nolegend
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
}
}

Add-WordPageBreak -WordDocument $WordDocument -Verbose

### Are there any errors in the logs that may be related to this issue? ###
Add-WordText -WordDocument $WordDocument -Text "Are there any errors in the logs that may be related to this issue?" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line

# SQL Error Log #
Add-WordText -WordDocument $WordDocument -Text "SQL .Error log: " -FontFamily "calibri light" -FontSize 11 -Supress $True -Color darkBlue -Bold $True
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
    $SQL = FetchSQLResults -query "SELECT @@log_error"
    $sqlerrors = gc $SQL -tail 60
foreach($line in $sqlerrors)
   {Add-WordText -WordDocument $WordDocument -Text "$line" -FontFamily "calibri light" -FontSize 8  -Color black -Supress $true}
Add-WordPageBreak -WordDocument $WordDocument -Verbose

# LTAError Log #
Add-WordText -WordDocument $WordDocument -Text "LTAError Log: " -FontFamily "calibri light" -FontSize 11 -Supress $True -Color darkBlue -Bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
    $ltaerrors = gc "C:\program files\labtech\logs\ltaerrors.txt" -tail 60
foreach($line in $ltaerrors)
   {Add-WordText -WordDocument $WordDocument -Text "$line" -FontFamily "calibri light" -FontSize 8  -Color black -Supress $true}
Add-WordPageBreak -WordDocument $WordDocument -Verbose

Add-WordText -WordDocument $WordDocument -Text "Windows Event Logs:" -FontFamily "calibri light" -FontSize 13 -Supress $True -Color darkBlue -Bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordText -WordDocument $WordDocument -Text "Application Logs:" -FontFamily "calibri light" -FontSize 11 -Supress $True -Color darkBlue -Bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
    $applogs = (Get-EventLog -LogName Application -EntryType Error -After (get-date).AddDays(-7))|select timegenerated, source, message
    foreach($log in $applogs) {$time = $log.timegenerated;$Source = $log.source;$msg=$log.Message -replace "`n|`r","|";$line= "$time ----- $source  ----- $msg";
    Add-WordText -WordDocument $WordDocument -Text "$line" -FontFamily "calibri light" -FontSize 8  -Color black -Supress $true 
    } 
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordText -WordDocument $WordDocument -Text "System Logs:" -FontFamily "calibri light" -FontSize 11 -Supress $True -Color darkBlue -Bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
    $syslogs = (Get-EventLog -LogName System -EntryType Error -After (get-date).AddDays(-7))|? {$_.Source -notlike '*schannel*' -and $_.Message -notlike '*printer*'}|select timegenerated, source, message
    foreach($log in $syslogs) {$time = $log.timegenerated;$Source = $log.source;$msg=$log.Message -replace "`n|`r","|";$line= "$time ----- $source  ----- $msg";
    Add-WordText -WordDocument $WordDocument -Text "$line" -FontFamily "calibri light" -FontSize 8  -Color black -Supress $true 
    } 
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
Add-WordText -WordDocument $WordDocument -Text "RMM System Logs:" -FontFamily "calibri light" -FontSize 11 -Supress $True -Color darkBlue -Bold $true
Add-WordParagraph -WordDocument $WordDocument -Supress $True # Adds an empty line
    $rmmlogs = (Get-EventLog -LogName "RMM System" -EntryType Error -After (get-date).AddDays(-7))|? {$_.Message -notlike "*Failed Checkin Procedure, Doing backup.*" -and $_.Message -notlike "*Failed Login, But telling them to ResignUP*"}|select timegenerated, source, message
    foreach($log in $rmmlogs) {$time = $log.timegenerated;$Source = $log.source;$msg=$log.Message -replace "`n|`r","|";$line= "$time ----- $source  ----- $msg";
    Add-WordText -WordDocument $WordDocument -Text "$line" -FontFamily "calibri light" -FontSize 8  -Color black -Supress $true 
    } 
   
$i=0
$count = $WordDocument.Paragraphs.count
do{ 
if($WordDocument.Paragraphs[$i].magictext.formatting[1].fontfamily.name -notlike '*calibri*' -and $WordDocument.Paragraphs[$i].magictext.formatting.count -gt 1){
$worddocument.paragraphs[$i]|Set-WordTextFontFamily -FontFamily "Calibri Light"| Set-WordTextFontSize -FontSize "10"}
$i++;}
until
($i -eq $count) 


### Save document
Save-WordDocument $worddocument -Supress $true -Language "en-US" -Verbose -OpenDocument
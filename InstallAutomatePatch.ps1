    Param
    (
      [Parameter(Mandatory = $true)]
      [String]$PatchDownloadLink
        
    )


#Function Declarations
######################
	Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language" -name Default -Value "0409"
	Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language" -name "InstallLanguage" -Value "0409"

Function Download-Patch
{
    Param
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$DownloadURL,
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$SavePath 
    )
	try
	{
		$DownloadObj = new-object System.Net.WebClient;
		$DownloadObj.DownloadFile($DownloadURL, $SavePath);
	}
	catch
	{
            $Output = $_.exception | Format-List -force | Out-String
            log-message "[*ERROR*] : $Output"
	}
}

Function Log-Message
{
	<#
	.SYNOPSIS
		A function to write ouput messages to a logfile.
	
	.DESCRIPTION
		This function is designed to send timestamped messages to a logfile of your choosing.
		Use it to replace something like write-host for a more long term log.
	
	.PARAMETER StrMessage
		The message being written to the log file.
	
	.EXAMPLE
		PS C:\> log-message -StrMessage 'This is the message being written out to the log.' 
	
	.NOTES
		N/A
#>
	
	Param
	(
		[Parameter(Mandatory = $True, Position = 0)]
		[String]$Message
	)

    
	add-content -path $LogFilePath -value ($Message)
    Write-Output $Message
}

function Get-SQLResult
{
    param 
    (
	    [Parameter(Mandatory = $true, Position = 0)]
	    [string]$Query
	)

	$result = .\mysql.exe --host="localhost" --user="root" --password="$rootpass" --database="LabTech" -e "$query" --batch --raw -N;
	return $result;
}

Function Output-Exception
{
    $Output = $_.exception | Format-List -force | Out-String
    $result = $_.Exception.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($result)
    $UsefulData = $reader.ReadToEnd();

    Write-log "[*ERROR*] : `n$Output `n$Usefuldata "  
}

#Variable Declarations
###########################
$ErrorActionPreference = 'SilentlyContinue'
[String]$LogFilePath = "$Env:windir\Temp\PSUpdateLog.txt"
[String]$PatchSavePath = "$Env:windir\Temp\CurrentPatch.exe"
[String]$PatchResultsPath = "$Env:windir\Temp\LTPatchLog.txt"
[String]$CustomTableName = "lt_patchinformation"
[String]$SQLDir = "C:\Program Files (x86)\Labtech\Mysql\bin\"


#Get Root Pass
###########################
$rootpass = <redacted>


#Remove Possible Leftovers
###########################
Remove-Item -Path $LogFilePath -Force
Remove-Item -Path $PatchSavePath -Force
Remove-Item -Path $PatchResultsPath -Force

#Kill the LTClient Process
###########################
IF(Get-process -Name 'LTClient')
{
    Stop-Process -Name 'LTClient' -Force
    Log-message -Message "The LTClient process has been killed!"
}

#Download the Patch
###########################

Download-Patch -DownloadURL $PatchDownloadLink -SavePath $PatchSavePath
$TestthePath= Test-Path $PatchSavePath
If(-not $TestthePath)
{
    Log-Message "Failed to download the patch."
    Return "Failed to download the patch."
    exit;
}

#Run the Patch
###########################
Set-Location "$Env:Windir\Temp"
dir c:\windows\temp\currentpatch.exe | Unblock-File
$AllArgs = "/ignoreemailtokenforcloud /s /p  360"
Stop-Service ltagent
Stop-Process -name ltagent
Stop-Service LTSCServiceMon
Stop-Process -name scservicemon
Stop-Service LTSCService
Stop-Process -name scservice
Stop-Service LTRedirSvc
Stop-Process -name ltredirsvc
Start-Process -FilePath "$PatchSavePath" -ArgumentList $AllArgs -Wait -WindowStyle Hidden
$LogFileResults = Get-content -Path $PatchResultsPath

$sqlcheck= get-service -name labmysql|select status
if($sqlcheck.status -notmatch 'Running') {write-output 'WARNING: LabMySQL Service is Not Running'; start-service labmysql;}
$ltacheck= get-service -name ltagent|select status;
if($ltacheck.status -notmatch 'Running') {write-output 'WARNING: LTAgent Service is Not Running'; start-service ltagent;}


#Check For the new Table
###########################

set-location "$sqldir";

$TableQuery = @"
SELECT * 
FROM information_schema.tables
WHERE table_schema = 'LabTech' 
    AND table_name = `'$CustomTableName`'
LIMIT 1;
"@

$TableCheck = get-sqlresult -query $TableQuery

If($TableCheck -eq $null)
{
    Log-message "Unable to find $CustomTableName in the database."
    [bool]$TableResult = $False
}

Else
{
    Log-message "Found $CustomTableName in the database."
    [bool]$TableResult = $True
}

If($LogFileResults -match "Automate Server has been successfully updated" -and $TableResult -eq $True)
{
    log-message "Patch was Successful"
    Return "Success"
}

Else
{
    log-message "Patch Failed"
    Return "Failure"
}


#Check for Valid Control Center Installer
###############################

$status = (Get-AuthenticodeSignature "C:\inetpub\wwwroot\LabTech\Updates\ClientUpdate.exe").Status
if($status -ne "Valid") {Invoke-WebRequest -uri "http://labtech-msp.com/release/ClientUpdate.exe" -OutFile "C:\inetpub\wwwroot\LabTech\Updates\ClientUpdate.exe"}
$status = (Get-AuthenticodeSignature "C:\inetpub\wwwroot\LabTech\Updates\ClientUpdate.exe").Status
if($status -ne "Valid") {write-output "###Control Center Update Failed###"}


$status = (Get-AuthenticodeSignature "C:\inetpub\wwwroot\LabTech\Updates\ControlCenterInstaller.exe").Status
if($status -ne "Valid") {Invoke-WebRequest -uri "http://labtech-msp.com/release/ControlCenterInstaller.exe" -OutFile "C:\inetpub\wwwroot\LabTech\Updates\ControlCenterInstaller.exe"}
$status = (Get-AuthenticodeSignature "C:\inetpub\wwwroot\LabTech\Updates\ControlCenterInstaller.exe").Status
if($status -ne "Valid") {write-output "###Control Center Update Failed###"}

  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force
  $wshell = New-Object -com WScript.Shell
  Add-Type -AssemblyName Microsoft.VisualBasic
  Add-Type -AssemblyName System.Windows.Forms
  $ping = New-Object System.Net.NetworkInformation.Ping

  $thiscmdid=  (Get-Process cmd |Sort-Object StartTime -ea SilentlyContinue|Where-Object{$_.MainWindowHandle -ne 0} |Select-Object -last 1).id

function Test-Admin {
param([switch]$Elevated)
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process $PsHome\powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

 function Set-WindowState {
	<#
	.LINK
	https://gist.github.com/Nora-Ballard/11240204
	#>

	[CmdletBinding(DefaultParameterSetName = 'InputObject')]
	param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
		[Object[]] $InputObject,

		[Parameter(Position = 1)]
		[ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE',
					 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED',
					 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
		[string] $State = 'SHOW'
	)

	Begin {
		$WindowStates = @{
			'FORCEMINIMIZE'		= 11
			'HIDE'				= 0
			'MAXIMIZE'			= 3
			'MINIMIZE'			= 6
			'RESTORE'			= 9
			'SHOW'				= 5
			'SHOWDEFAULT'		= 10
			'SHOWMAXIMIZED'		= 3
			'SHOWMINIMIZED'		= 2
			'SHOWMINNOACTIVE'	= 7
			'SHOWNA'			= 8
			'SHOWNOACTIVATE'	= 4
			'SHOWNORMAL'		= 1
		}

		$Win32ShowWindowAsync = Add-Type -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
'@ -Name "Win32ShowWindowAsync" -Namespace Win32Functions -PassThru

		if (!$global:MainWindowHandles) {
			$global:MainWindowHandles = @{ }
		}
	}

	Process {
		foreach ($process in $InputObject) {
			if ($process.MainWindowHandle -eq 0) {
				if ($global:MainWindowHandles.ContainsKey($process.Id)) {
					$handle = $global:MainWindowHandles[$process.Id]
				} else {
					Write-Error "Main Window handle is '0'"
					continue
				}
			} else {
				$handle = $process.MainWindowHandle
				$global:MainWindowHandles[$process.Id] = $handle
			}

			$Win32ShowWindowAsync::ShowWindowAsync($handle, $WindowStates[$State]) | Out-Null
			Write-Verbose ("Set Window State '{1} on '{0}'" -f $MainWindowHandle, $State)
		}
	}
}

### remove temp folder ##

 if(test-path  C:\temp_aitest\){
    remove-item C:\temp_aitest\ -Recurse -Force -ea SilentlyContinue 
    }

 if(test-path  C:\logs\){
    remove-item  C:\logs\ -Recurse -Force -ea SilentlyContinue 
    }

####### Close DellCommandUpdate####

$programname="DellCommandUpdate"

if(get-process -Name $programname -ea SilentlyContinue){
Stop-Process -Name $programname
}


#######'running with full privileges'


if($PSScriptRoot.length -eq 0){
$scriptRoot="C:\testing_AI"
}
else{
$scriptRoot=$PSScriptRoot
}


###########  new logs_timemap.csv ########

$logtimemap="$scriptRoot\logs\logs_timemap.csv"

do{
Start-Sleep -s 1
$readok=Get-ChildItem $logtimemap
$checkcontent=get-content $logtimemap
}until($readok -and $checkcontent.Length -gt 0)

if(test-path $logtimemap){

$current=import-csv  $logtimemap

 $checkfinish=($current|Where-Object{$_.Actions -eq "finalization"}).Results
 if($checkfinish){
  $ansnew=read-host "A New Auto-test Start? (Enter:yes,N:no) "
  if(-not($ansnew -like "*n*")){
$day=get-date -format "yyyyMMdd_HHmm"
$path="$scriptRoot\logs\logs_timemap.csv" 
$path2="$scriptRoot\logs\logs_timemap2.csv" 
$path3="$scriptRoot\logs\logs_timemap_backup_$day.csv" 
$title=get-content $path|Select-Object -First 1
Set-Content -path $path2  -value $title
Move-Item $path $path3 -Force
Move-Item $path2  $path -Force
}
 else{
 exit
 }

 }
 }

#$flsettings=import-csv $scriptRoot\settings\flowsettings.csv

Do {

$flsettings=import-csv $scriptRoot\settings\flowsettings.csv

$current=import-csv $scriptRoot\logs\logs_timemap.csv

 $donetc=($current|Select-Object -last 1).TC
 $donestep=($current|Select-Object -last 1).Step_No

$skiptc=($current|Where-Object{$_."TC" -eq $donetc -and  $_."must" -match "y" -and ($_."results" -match "NG" -or $_."results" -match "FAIL")}).TC
 
# $wait_check=test-path C:\testing_AI\logs\wait.txt 

######  check next #########

if(($current.TC).count -eq 0 ){
$flsetting=$flsettings[0]
}

else{

  $next=$flsettings.IndexOf(($flsettings|Where-Object{$_."TC" -eq $donetc -and  $_."Step_No" -eq  $donestep}))

  if($next+1 -eq $flsettings.Count){
  $next=-1
  }

 $flsetting=$flsettings[$next+1]

  #if($wait_check -eq $false){  $flsetting=$flsettings[$next+1]}
   # if($wait_check -eq $true){ $flsetting=$flsettings[$next]}

  }

$tcnumber=$flsetting."TC"
$tcstep=$flsetting."Step_No"
$action=$flsetting."programs"
$para1=$flsetting."para1"
$para2=$flsetting."para2"
$para3=$flsetting."para3"
$para4=$flsetting."para4"
$para5=$flsetting."para5"

new-item -path $scriptRoot\currentjob\TC.txt -value "$tcnumber,$tcstep,$action" -Force |Out-Null

if($tcnumber -ne $donetc -and ($current.TC).count -ne 0 -and $donetc.Length -gt 0){
	
	# rollback if dns not auto
	$actiondns="dns_settings"
    Get-Module -name $actiondns |remove-module
    $mdpath=(get-childitem -path $scriptRoot -r -file |Where-Object{$_.name -match "^$actiondns\b" -and $_.name -match "psm1"}).fullname
    Import-Module $mdpath -WarningAction SilentlyContinue -Global
  
	&$actiondns -para2 "nonlog"

	#region check net connecting
	$connet11=1
	$connet12=1
	$connet21=1
	$connet22=1
	try{
	($ping.Send("192.168.2.249", 1000)).Status
	}catch{
	$connet11=0
	}
	try{
	Invoke-WebRequest -Uri "www.msn.com" -UseBasicParsing 
	}catch{
	$connet12=0
	}
	try {		
	($ping.Send("www.google.com", 1000)).Status 
	}
	catch {
	$connet21=0
	}
	try {		
	Invoke-WebRequest -Uri "www.msn.com" -UseBasicParsing 
	}
	catch {
	$connet22=0
	}

 if(($connet11 -or $connet12) -ne 1 -and ($connet21 -or $connet22) -ne 0){
    Write-Output "connecting status not right. need to set the default network"
    $actionconnect="net_connecting"
    Get-Module -name $actionconnect |remove-module
    $mdpath=(get-childitem -path $scriptRoot -r -file |Where-Object{$_.name -match "^$actionconnect\b" -and $_.name -match "psm1"}).fullname
    Import-Module $mdpath -WarningAction SilentlyContinue -Global
	
    $actiondisconnect="net_disconnecting"
    Get-Module -name  $actiondisconnect |remove-module
    $mdpath=(get-childitem -path $scriptRoot -r -file |Where-Object{$_.name -match "^ $actiondisconnect\b" -and $_.name -match "psm1"}).fullname
    Import-Module $mdpath -WarningAction SilentlyContinue -Global
	
	&$actionconnect -para2 nonlog
	&$actiondisconnect -para1 internet -para2 nonlog

  }
  #endregion

  #region close apps
   $apps_needclose=@("3DMark","gpu-z","nvwdmcpl")
   foreach ($app in $apps_needclose){
   $app_pid=(Get-Process $app -ErrorAction SilentlyContinue).id
   if($app_pid){
   Write-Output "close $app"
   (get-process -name $app).CloseMainWindow()
   $app_pid=(Get-Process $app -ErrorAction SilentlyContinue).id
   if($app_pid){
	stop-process -id $app_pid -Force
   }
   }
   }
  #endregion

}

## revise last TC logs filename with timestamps ##
if($tcnumber -ne $donetc -and ($current.TC).count -ne 0 -and $donetc -ne "fin" -and $donetc.Length -gt 0){
$renamefiles=Get-ChildItem $scriptRoot\logs\$donetc\* -file -Exclude "*.exe" -ErrorAction SilentlyContinue
foreach($renamefile in $renamefiles){
if(!($renamefile.name -match "^\d{6}_${6}")){
$datewrite=get-date((Get-ChildItem $renamefile.fullname).LastWriteTime) -format "yyMMdd_HHmmss"
$newname=$datewrite+"_"+$renamefile.Name
rename-item $renamefile.fullname -NewName $newname ## rename as same date format
}
}
}


   # stop-Transcript
    try{Stop-Transcript }catch {write-host ""}
    
   $scriptlog2="$scriptRoot\logs\scriptlogs\scriptlogs2.txt"
   if(!(test-path $scriptlog2)){New-Item -Path $scriptlog2 -Force |Out-Null}
   $scriptloga=(Get-ChildItem "$scriptRoot\logs\scriptlogs\PowerShell_transcript*.txt"|Sort-Object lastwritetime|Select-Object -last 1).fullname

## log content shrink ##       
if( $scriptloga.Length -gt 0){
$rec=""
$newscript= (get-content  $scriptloga) |ForEach-Object{
if($_ -match "Start time" -or $_ -match "\*{1,}" -or $_ -match "End time"){
$rec="yes"
}
if($_ -match "Windows PowerShell transcript end" -or $_ -match "Username\:"){
$rec=""
}

if($rec -match "yes" -and ($_ -notmatch "transcript start" -or $_ -notmatch "transcript end")){
$_
}
}
    $newscript|add-Content $scriptlog2 -Force
    
    remove-item $scriptloga -Force

    }

    ## starting log ##
    
 if(!(test-path $scriptRoot\logs\scriptlogs)){new-item -ItemType directory -path $scriptRoot\logs\scriptlogs\ -Force |Out-Null}

    Start-Transcript -OutputDirectory $scriptRoot\logs\scriptlogs\ |out-null  ## save logs

## if TC need skip ##

  if($tcnumber -eq $skiptc){

write-host "SKIP $($tcnumber)_step$($tcstep)"

Get-Module -name "outlog"|remove-module
$mdpath=(Get-ChildItem -path "C:\testing_AI\modules\"  -r -file |Where-Object{$_.name -match "outlog" -and $_.name -match "psm1"}).fullname
Import-Module $mdpath -WarningAction SilentlyContinue -Global

#write-host "Do $action!"
outlog $action "SKIP" $tcnumber $tcstep "-"

  }
  else{
#write-host "Do $action!"
Get-Module -name $action|remove-module
$mdpath=(Get-ChildItem -path $scriptRoot\modules\* -r -file |Where-Object{$_.name -match "^$action\b" -and $_.name -match "psm1"}).fullname
Import-Module $mdpath -WarningAction SilentlyContinue -Global

if(-not($tcnumber -match "test")){
##
 Get-Process -id $thiscmdid  | Set-WindowState -State MINIMIZE
}
##>

write-host " $($action) -para1 ""$($para1)"" -para2 ""$($para2)"" -para3 ""$($para3)"" -para4 ""$($para4)"" -para5 ""$($para5)"""

&$action -para1 $para1 -para2 $para2 -para3 $para3 -para4 $para4 -para5 $para5

 #$Error >> $scriptRoot\logs\logserr.txt 
 
 }
######  check finish #########

start-sleep -s 2

$current=import-csv $scriptRoot\logs\logs_timemap.csv

 $checkfinish=($current|Where-Object{$_.Actions -eq "finalization"}).Results

 
} until ($checkfinish -eq "OK")

 
 # create a new .NET type
$signature = @"
[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@
Add-Type -MemberDefinition $signature -Name MyType -Namespace MyNamespace

Get-Process |   Where-Object { $_.MainWindowHandle -ne 0  } |
  ForEach-Object { 
  $handle = $_.MainWindowHandle

  # minimize window
  $null = [MyNamespace.MyType]::ShowWindowAsync($handle, 2)
}

    #stop-Transcript |out-null
    
   $scriptlog2="$scriptRoot\logs\scriptlogs\scriptlogs2.txt"
   if(!(test-path $scriptlog2)){New-Item -Path $scriptlog2 -Force |Out-Null}

    $scriptloga=(Get-ChildItem "$scriptRoot\logs\scriptlogs\PowerShell_transcript*.txt"|Sort-Object lastwritetime|Select-Object -last 1).fullname
       

$rec=""
$newscript= (get-content  $scriptloga) |ForEach-Object{
if($_ -match "Start time" -or $_ -match "\*{1,}" -or $_ -match "End time"){
$rec="yes"
}
if($_ -match "Windows PowerShell transcript end" -or $_ -match "Username\:"){
$rec=""
}

if($rec -match "yes" -and ($_ -notmatch "transcript start" -or $_ -notmatch "transcript end")){
$_
}
}
    $newscript|add-Content $scriptlog2 -Force

    remove-item $scriptloga -Force

[System.Windows.Forms.MessageBox]::Show($this,"TC Auto Test Complete!")



$logfs=Get-ChildItem C:\testing_AI\logs\ -directory |Where-Object{$_.Name -notlike "*script*" -and $_.Name -notmatch "_\d{6}_\d{4}\b"}

$datenow=get-date -format "yyMMdd_HHmm"
foreach($logf in $logfs){
$newname=($logf.name)+"_"+$datenow
Rename-Item -Path $logf.fullname -NewName $newname
}

#Run Log Html
C:\testing_AI\settings\loghtmlcheck.ps1

C:\testing_AI\logs\loglist.html

#explorer 'C:\testing_AI\logs\'

#create shrotcut to desktop
$shortcut = $wshell.CreateShortcut("C:\Users\$($env:USERNAME)\Desktop\loglist.lnk")
$shortcut.TargetPath = "C:\testing_AI\logs\loglist.html"
$shortcut.WorkingDirectory = "C:\testing_AI\logs\"
$shortcut.Description = "start"
$shortcut.IconLocation = "C:\testing_AI\logs\loglist.html"
#$shortcut.Arguments = "YourArgumentsHere"
$shortcut.Save()

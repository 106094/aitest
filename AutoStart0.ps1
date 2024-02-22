
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
    $wshell=New-Object -ComObject wscript.shell
    $shell=New-Object -ComObject shell.application
      Add-Type -AssemblyName Microsoft.VisualBasic
      Add-Type -AssemblyName System.Windows.Forms
      $ping = New-Object System.Net.NetworkInformation.Ping

#region check/settings

 ## ping id
 ## credential of dell server settins
 ## copy AutoStart.ps1 from 192.168.2.249

#endregion


#region form

function formaform([string]$para1){

Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Sample Form"
$Form.AutoSize = $True

# Create a font for the label
$Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Regular)
$Form.Font = $Font

# Create a label
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "Waiting for connecting checks.       "
$Label.AutoSize = $True
$label.ForeColor="blue"

# Calculate the position to center the label both horizontally and vertically
$labelWidth = $Label.Width
$labelHeight = $Label.Height
$horizontalOffset = ($Form.ClientSize.Width - $labelWidth) / 2
$verticalOffset = ($Form.ClientSize.Height - $labelHeight) / 2
$Label.Location = New-Object System.Drawing.Point($horizontalOffset, $verticalOffset)

# Add the label to the form
$Form.Controls.Add($Label)

# Show the form
$Form.Show()


function connectpass{
start-sleep -Milliseconds 300
$Label.Text = "Waiting for connecting checks.      "
$Form.Update()
start-sleep -Milliseconds 300
$Label.Text = "Waiting for connecting checks..     "
$Form.Update()
start-sleep -Milliseconds 300
$Label.Text = "Waiting for connecting checks...     "
$Form.Update()
start-sleep -Milliseconds 300
$Label.Text = "connecting check ok      "
$label.ForeColor = "green"
$Form.Update()
start-sleep -s 2
}

function reconnect{
$a=0
do{
start-sleep -Milliseconds 300
$a++
$Label.Text = "Waiting for connecting checks.      "
$Form.Update()
start-sleep -Milliseconds 300
$Label.Text = "Waiting for connecting checks..     "
$Form.Update()
start-sleep -Milliseconds 300
$Label.Text = "Waiting for connecting checks...     "
$Form.Update()
}until($a -gt 10)

}

# Wait for 3 seconds and update the label text
if($para1 -match "pass"){
connectpass
}
if($para1 -match "reconnect"){
reconnect
}
$Form.Close()
}

#endregion

#region admin
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

#endregion

 ### close file explore windows

 $shell.Windows() |Where-Object{$_.name -eq "File Explorer"}| ForEach-Object { $_.Quit() }
 
#region ping 

### test connecting

 $mess=""
 
 $testping1= ($ping.Send("192.168.2.249", 1000)).Status
 # $testping1= ping 192.168.2.249 /n 3
  #if( ($testping1 -match "unreachable" -or $testping1 -match "Request timed out" -or $testping1 -match "failed")){$mess=$mess+@("Need collecting to 192.168.2.249 server")}
  if( !($testping1 -match "Success")){$mess=$mess+@("Need connect to 192.168.2.249 server")}

  if(!(test-path "C:\testing_AI\")){

  $setting=Get-ItemPropertyValue -Path HKLM:SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity -Name Enabled -ErrorAction SilentlyContinue

 if( $setting -ne 1){
   [Microsoft.VisualBasic.Interaction]::MsgBox("等會兒需手動 turn On Memory Integrity`n 並依照Windows指示Reboot `n (Reboot後會再繼續自動執行排定流程)",'OKOnly,SystemModal,Information', 'check')
  }

   $testping2= ($ping.Send("172.16.21.249", 1000)).Status
   #$testping2= ping 172.16.21.249 /n 3 ## swtool.allion.com
   if( !($testping2 -match "Success")){$mess=$mess+@("Need connect Allion.test (172.16.21.249)")}
  #if( ($testping2 -match "unreachable" -or $testping2 -match "Request timed out" -or $testping2 -match "failed")){$mess=$mess+@("Need collecting to Allion.test (172.16.21.249)")}
 
  if($mess.length -eq 0){
   formaform -para1 pass
   }  

  if($mess.length -ne 0){

write-host "fail to connect"
write-host "enable net settings, please wait..."

#regioni dns set to default

$getinfo=ipconfig
$checkline=$getinfo -match "allion.test"

if($checkline){

$linenu=$getinfo.IndexOf($checkline)
($getinfo|Select-Object -Skip $linenu|Select-Object -First 3|Select-Object -last 1) -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}" |Out-Null
$ipout=$matches[0]
if($ipout){

Get-NetIPAddress|%{

if($_.IPAddress -match $ipout -and $_.AddressFamily -eq "IPv4"){
$adtname=$_.InterfaceAlias 
}
}

## rollback dns settins
Set-DnsClientServerAddress -InterfaceAlias $adtname -ResetServerAddresses
Clear-DnsClientCache
}

}

#endregion

 $netnames=((Get-NetAdapter)|Where-Object{$_.status -ne "UP" -and $_.name -notmatch "Bluetooth"}).name
  foreach($netname in $netnames){
  Enable-NetAdapter -Name $netname -Confirm:$false
  #Disable-NetAdapter -Name $netname -Confirm:$false
  }
   formaform -para1 reconnect
 
  ##check again ## 

 do{
   
 try{ ipconfig /renew}
 catch{
 write-host "ipconfig renew fail"
 }

 $mess=""
  #$testping1= ping 192.168.2.249 /n 3
  #if( ($testping1 -match "unreachable" -or $testping1 -match "Request timed out" -or $testping1 -match "failed")){$mess=$mess+@("Need collecting to 192.168.2.249 server")}
    $testping1= ($ping.Send("192.168.2.249", 1000)).Status
    if( !($testping1 -match "Success")){$mess=$mess+@("Need collecting to 192.168.2.249 server")}

  if(!(test-path "C:\testing_AI\")){
   #$testping2= ping 172.16.21.249 /n 3 ## swtool.allion.com
    #if( ($testping2 -match "unreachable" -or $testping2 -match "Request timed out" -or $testping2 -match "failed")){$mess=$mess+@("Need collecting to Allion.test (172.16.21.249)")}
     $testping2= ($ping.Send("172.16.21.249", 1000)).Status
       if( !($testping2 -match "Success")){$mess=$mess+@("Need collecting to Allion.test (172.16.21.249)")}
   }
  if($mess.length -ne 0){
  $mess=$mess.trim()|out-string
  $result2 = [System.Windows.Forms.MessageBox]::Show("$mess `n Continue without internet? `n`n About: Exit `n Retry: Need to connect to internet. (click Retry after link the cable) `n Ignore: No need to connect to internet " , "Info" , 2)
 
 if($result2 -eq "abort" ){
 stop-process -name cmd -Force
  exit
 }
 
 }
  }until($mess.length -eq 0 -or $result2 -eq "ignore" )
  
  write-host "net connecting check ok, continue..."

  formaform -para1 pass
  

  }

}


# disabling net device is conflict with dellgui check DM -> move to autostart.ps1
<#
  if(test-path "C:\testing_AI\" -and !($env:computername -match "cycling") ){
  ## disconnect internet but pass cycling tool
  
$netdisc="net_disconnecting"
Get-Module -name $netdisc|remove-module
$mdpath=(gci -path C:\testing_AI\modules\ -r -file |?{$_.name -match "^$netdisc\b" -and $_.name -match "psm1"}).fullname
Import-Module $mdpath -WarningAction SilentlyContinue -Global

&$netdisc -para1 "internet" -para2 "nolog"

  }
  #>

#endregion

#region  add local server web credential ##

$targetpath="192.168.2.249"
$usernm="pctest"
$passwd="pctest"
$checkcd= (cmdkey /list) -match $targetpath
if(!$checkcd){
#$secpasswd = ConvertTo-SecureString $passwd -AsPlainText -Force
#$credential = New-Object System.Management.Automation.PSCredential ($usernm, $secpasswd)
$command_set=cmdkey /add:$targetpath /user:$usernm /pass:$passwd
Start-Sleep -s 5
$checkcd= (cmdkey /list) -match $targetpath
}

if($checkcd){
 write-host  $checkcd
 }
else{
[System.Windows.Forms.MessageBox]::Show($this, "Fail to setup credential for 192.168.2.249, please check!") 
exit
}

#endregion

#region copy files

if($PSScriptRoot.length -eq 0){
$scriptRoot="D:\!_Dell_AITest_Update"
}
else{
$scriptRoot=$PSScriptRoot
}

Copy-Item \\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\_Dell_AITest_Start\AutoStart.ps1  -Destination $scriptRoot -Force
Copy-Item \\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\_Dell_AITest_Start\tc_flowsettings.csv  -Destination $scriptRoot -Force
Copy-Item \\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\_Dell_AITest_Start\dellgui.ps1  -Destination $scriptRoot -Force

start-sleep -s 2

#gci \\192.168.2.249\srvprj\Inventec\Dell\Matagorda\temp\_Dell_AITest_Start\* -Exclude @("flowsettings.csv","AutoStart.bat","AutoStart0.ps1")|Copy-Item -Destination $scriptRoot -Force

#endregion

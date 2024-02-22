#region check/settings
 
 ## check os entering status
 ## disable wu
 ## netconnection check -> AutoStart0.ps1
 ## credential of dell server settins  -> AutoStart0.ps1
 ## powersetting
 ## turn off notification
 ## display resolution check and set up to 1920*1080
 ## cmd and windows terminal settings for windows 11
 ## rename system
 ## check system information->dellgui
 ## check memory integrity　->dellgui
 ## run dellgui
 ## disconnect internet but pass cycling tool
 ## create shortcut to dell server at desktop
 ## Reboot
 ## copy AI Tool files
 ## setting remote control
 ## setting WinRM service

#endregion

## check os entering status
while(!$osstatus){
start-sleep -s 1
try{
    $testpath="$env:USERPROFILE/desktop/test.log"
    new-item $testpath -Force |Out-Null
    $writesomething=(get-date -format "yyyyMD_HHmmss"|Out-String)
    add-content -path $testpath -Value $writesomething
    $readcontent=get-content $testpath |Out-String
    if($readcontent.trim() -eq $writesomething.trim()){
        $osstatus=$true
    }
}catch{
    $osstatus=$false
}

}
remove-item $testpath -Force

$starttime=Get-Date
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
    $wshell=New-Object -ComObject wscript.shell
    $shell=New-Object -ComObject shell.application
      Add-Type -AssemblyName Microsoft.VisualBasic
      Add-Type -AssemblyName System.Windows.Forms

#close oobe
stop-process -name WWAHost -Force -ErrorAction SilentlyContinue
stop-process -name WebExperienceHostApp -Force -ErrorAction SilentlyContinue

#region general functions

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

 # create a new .NET type for  close  app windows
$signature = @"
[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@
Add-Type -MemberDefinition $signature -Name MyType -Namespace MyNamespace


Function Set-ScreenResolution { 
 
<# 
    .Synopsis 
        Sets the Screen Resolution of the primary monitor 
    .Description 
        Uses Pinvoke and ChangeDisplaySettings Win32API to make the change 
    .Example 
        Set-ScreenResolution -Width 1024 -Height 768         
    #> 
param ( 
[Parameter(Mandatory=$true, 
           Position = 0)] 
[int] 
$Width, 
 
[Parameter(Mandatory=$true, 
           Position = 1)] 
[int] 
$Height 
) 
 
$pinvokeCode = @" 
 
using System; 
using System.Runtime.InteropServices; 
 
namespace Resolution 
{ 
 
    [StructLayout(LayoutKind.Sequential)] 
    public struct DEVMODE1 
    { 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmDeviceName; 
        public short dmSpecVersion; 
        public short dmDriverVersion; 
        public short dmSize; 
        public short dmDriverExtra; 
        public int dmFields; 
 
        public short dmOrientation; 
        public short dmPaperSize; 
        public short dmPaperLength; 
        public short dmPaperWidth; 
 
        public short dmScale; 
        public short dmCopies; 
        public short dmDefaultSource; 
        public short dmPrintQuality; 
        public short dmColor; 
        public short dmDuplex; 
        public short dmYResolution; 
        public short dmTTOption; 
        public short dmCollate; 
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] 
        public string dmFormName; 
        public short dmLogPixels; 
        public short dmBitsPerPel; 
        public int dmPelsWidth; 
        public int dmPelsHeight; 
 
        public int dmDisplayFlags; 
        public int dmDisplayFrequency; 
 
        public int dmICMMethod; 
        public int dmICMIntent; 
        public int dmMediaType; 
        public int dmDitherType; 
        public int dmReserved1; 
        public int dmReserved2; 
 
        public int dmPanningWidth; 
        public int dmPanningHeight; 
    }; 
 
 
 
    class User_32 
    { 
        [DllImport("user32.dll")] 
        public static extern int EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE1 devMode); 
        [DllImport("user32.dll")] 
        public static extern int ChangeDisplaySettings(ref DEVMODE1 devMode, int flags); 
 
        public const int ENUM_CURRENT_SETTINGS = -1; 
        public const int CDS_UPDATEREGISTRY = 0x01; 
        public const int CDS_TEST = 0x02; 
        public const int DISP_CHANGE_SUCCESSFUL = 0; 
        public const int DISP_CHANGE_RESTART = 1; 
        public const int DISP_CHANGE_FAILED = -1; 
    } 
 
 
 
    public class PrmaryScreenResolution 
    { 
        static public string ChangeResolution(int width, int height) 
        { 
 
            DEVMODE1 dm = GetDevMode1(); 
 
            if (0 != User_32.EnumDisplaySettings(null, User_32.ENUM_CURRENT_SETTINGS, ref dm)) 
            { 
 
                dm.dmPelsWidth = width; 
                dm.dmPelsHeight = height; 
 
                int iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_TEST); 
 
                if (iRet == User_32.DISP_CHANGE_FAILED) 
                { 
                    return "Unable To Process Your Request. Sorry For This Inconvenience."; 
                } 
                else 
                { 
                    iRet = User_32.ChangeDisplaySettings(ref dm, User_32.CDS_UPDATEREGISTRY); 
                    switch (iRet) 
                    { 
                        case User_32.DISP_CHANGE_SUCCESSFUL: 
                            { 
                                return "Success"; 
                            } 
                        case User_32.DISP_CHANGE_RESTART: 
                            { 
                                return "You Need To Reboot For The Change To Happen.\n If You Feel Any Problem After Rebooting Your Machine\nThen Try To Change Resolution In Safe Mode."; 
                            } 
                        default: 
                            { 
                                return "Failed To Change The Resolution"; 
                            } 
                    } 
 
                } 
 
 
            } 
            else 
            { 
                return "Failed To Change The Resolution."; 
            } 
        } 
 
        private static DEVMODE1 GetDevMode1() 
        { 
            DEVMODE1 dm = new DEVMODE1(); 
            dm.dmDeviceName = new String(new char[32]); 
            dm.dmFormName = new String(new char[32]); 
            dm.dmSize = (short)Marshal.SizeOf(dm); 
            return dm; 
        } 
    } 
} 
 
"@ 
 
Add-Type $pinvokeCode -ErrorAction SilentlyContinue 
[Resolution.PrmaryScreenResolution]::ChangeResolution($width,$height) 
} 

function outwords{

param(
    [string]$line,
    [int]$timeset
    
    )
    
$i=0
if($line.Length -gt 0){
do{
write-host $line[$i]  -NoNewline
$i++
start-sleep -millisecond $timeset
}until($i -eq $line.length)
}
}

function ShowWindow( $hwnd ){

    $constants = @{
        ShowWindowCommands = @{ #from https://pinvoke.net/default.aspx/Enums/ShowWindowCommands.html
            Hide = 0; #completely hides the window
            Normal = 1; #if min/max'd, restores to original size and position
            ShowMinimized = 2; #activates and minimizes window
            Maximize = 3; #activates and maximizes window
            ShowMaximized = 3; #activates and maximizes window
            ShowNoActivate = 4; #shows a window in its most recent size and position without activating it
            Show = 5; #activates the window and displays it in its current size and position
            Minimize = 6; #minimizes and activates the next top-level window
            ShowMinNoActive = 7; #minimizes and activates no windows"
            ShowNA = 8; #shows a window in its current size and position without activating it
            Restore = 9; #activates and displays window. if min/max'd, restores to original size and position
            ShowDefault = 10; #sets the window to its default show state
            ForceMinimize = 11; #Windows 2000/XP-only feature. minimize window even if thread is hung
        };
        WindowLongParam = { #from https://www.pinvoke.net/default.aspx/Constants/GWL%20-%20GetWindowLong.html
            SetWndProc = -4; #sets new address for procedure (can't be changed unless the window belongs to the same thread)
            SetHndInst = -6; #sets a new application instance handle
            SetHndParent = -8 #unsepecified
            SetId = -12 #sets a new identifier of the child window
            SetStyle = -16 #sets a new window style
            SetExtStyle = -20 #sets a new extended window style
            SetUserData = -21 #sets the user data associated with the window
            #there's a few more of these that are positive for the dialog box procedure
        };
        WindowsStyles = {
            #see https://learn.microsoft.com/en-us/windows/win32/winmsg/window-styles
            # and
            #  http://pinvoke.net/default.aspx/Constants.Window%20styles
            # for more... (there's a lot)
            Minimize = 0x20000000; #hexadecimal of 536870912
        }
    }

    $sigs = '[DllImport("user32.dll", EntryPoint="GetWindowLong")]
    public static extern IntPtr GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("User32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, int lpdwProcessId);
    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();'

    $type = Add-Type -MemberDefinition $sigs -Name WindowAPI5 -IgnoreWarnings -PassThru
    
    $type::ShowWindow( $hwnd, $constants.ShowWindowCommands.Minimize )
    $type::ShowWindow( $hwnd, $constants.ShowWindowCommands.Restore )

    [int] $currentlyFocusedWindowProcessId = $type::GetWindowThreadProcessId($type::GetForegroundWindow(), 0)
    [int] $appThread = [System.AppDomain]::GetCurrentThreadId()

    if( $currentlyFocusedWindowProcessId -ne $appThread ){
    
        $type::AttachThreadInput( $currentlyFocusedWindowProcessId, $appThread, $true )
        $type::BringWindowToTop( $hwnd )
        $type::ShowWindow( $hwnd, $constants.ShowWindowCommands.Show )
        $type::AttachThreadInput( $currentlyFocusedWindowProcessId, $appThread, $false )

    } else {
        $type::BringWindowToTop( $hwnd )
        $type::ShowWindow( $hwnd, $constants.ShowWindowCommands.Show )
    }

}

#endregion

function initial_copytool{

    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force;
    $wshell=New-Object -ComObject wscript.shell
      Add-Type -AssemblyName Microsoft.VisualBasic
      Add-Type -AssemblyName System.Windows.Forms

### create netdisk for file copying##      
            
 function netdisk_connect([string]$webpath,[string]$username,[string]$passwd,[string]$diskid){

net use $webpath /delete
net use $webpath /user:$username $passwd /PERSISTENT:yes
 net use $webpath /SAVECRED 

 if($diskid.length -ne 0){
  $diskpath=$diskid+":"
  $checkdisk=net use
   if($checkdisk -match $diskpath){net use $diskpath /delete}
    net use $diskpath $webpath
}

}

netdisk_connect -webpath \\192.168.2.249\srvprj\Inventec\Dell -username pctest -passwd pctest -diskid Y

### remove buff ###

$dest="C:\testing_AI"
$shortcutadd=$false

if(!(test-path $dest)){
$shortcutadd=$true
}
else{
  if(test-path "$dest\AutoRun_.bat"){rename-Item "$dest\AutoRun_.bat" "$dest\AutoRun.bat" -Force -ea SilentlyContinue}
  if(test-path "$dest\logs\wait.txt"){Remove-Item "$dest\logs\wait.txt" -Force -ea SilentlyContinue}
  start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_Run" -f' 
$path="$($dest)\logs\logs_timemap.csv" 
 
  if((test-path $path) -eq $true){

########### force  new logs_timemap.csv ########

$logtimemap="C:\testing_AI\logs\logs_timemap.csv"
$current=import-csv  $logtimemap

$day=get-date -format "yyyyMMdd_HHmm"
$path="C:\testing_AI\logs\logs_timemap.csv" 
$path2="C:\testing_AI\logs\logs_timemap2.csv" 
$path3="C:\testing_AI\logs\logs_timemap_backup_$day.csv" 
$title=get-content $path|select -First 1
Set-Content -path $path2  -value $title
Move-Item $path $path3 -Force
Move-Item $path2  $path -Force


}
 }

## define and copy　files

  $autofilesraw = @{
  
  "mainzip" = @{
       rank=1
       filename = "testing_AI.zip"
       copyto = ""
       update = "no"
    }

  "modules" = @{
       rank=2
       filename = "modules\"
       copyto = ""
       update = "yes" 
    }

   "selenium" = @{
       rank=3
       filename = "selenium\"
       copyto = "modules\"
       update = "yes" 
    }

  "BI_config"= @{
       rank=4
       filename= "BI_config\*"
       copyto = "modules\BITools\config\"
       update = "yes" 
    }

    "nv_Controlpanel" = @{
       rank=5
       filename = "nv_Controlpanel\"
       copyto = "settings\" 
       update = "yes" 
    }

    "pcai" = @{
        rank=14
        filename = "pcai\Main\"
        copyto = "modules\PC_AI_Tool*\"
        update = "yes" 
    }

    "driverinstall" = @{
        rank=7
        filename = "driver\"
        copyto = "modules\"
        update = "yes" 
    }

    "flex" = @{
       rank=9
       filename = "extra_tools\flex.zip"
       copyto = "modules\BITools\flex\"
       update = "no" 
    }


      "PC_AI_Tool" = @{
       rank=13
       filename = "extra_tools\PC_AI_Tool_*.zip"
       copyto = "modules\"
       update = "no" 
    }

      "InvCol" = @{
       rank=15
       filename = "extra_tools\InvCol.zip"
        copyto = "modules\InvCol"
        update = "no" 
    }

     "yuvplayer" = @{
       rank=16
       filename = "extra_tools\yuvplayer.zip"
        copyto = "modules\yuvplayer"
        update = "no" 
    }

     "pswindowsupdate" = @{
       rank=17
       filename = "extra_tools\pswindowsupdate.zip"
        copyto = "modules\pswindowsupdate"
        update = "no" 
    }

    "py" = @{
       rank=18
       filename = "py\*"
       copyto = "modules\py"
       update = "yes" 
    }

      "settings" = @{
       rank=19
       filename = "settings\*"
       copyto = "settings\"
       update = "yes" 
    }

    "Dominion" = @{
       rank=20
       filename = "extra_tools\Dominion.zip"
       copyto = "modules\Dominion"
       update = "no" 
    }

      "DDU" = @{
       rank=21
       filename = "extra_tools\DDU.zip"
       copyto = "modules\DDU"
       update = "no" 
    }
      "GPUMon" = @{
       rank=22
       filename = "extra_tools\GPUMon.zip"
       copyto = "modules\GPUMon"
       update = "no" 
    }
      "Auto-Click" = @{
       rank=23
       filename = "extra_tools\Auto-Click.zip"
       copyto = "modules\Auto-Click"
       update = "no" 
    }
      "NXTest_Freeware_Windows" = @{
       rank=24
       filename = "extra_tools\NXTest_Freeware_Windows.zip"
       copyto = "modules\NXTest_Freeware_Windows"
       update = "no" 
    }
<##
   "PCAgent" = @{
       rank=21
       filename = "extra_tools\PCAgent*.zip"
       copyto = "modules\"
       update = "no" 
    }
    
    "3dmark" = @{
       rank=10
       filename = "extra_tools\3dmark.zip"
       copyto = "modules\BITools\3dmark\"
       update = "no" 
    }

    "cloudegate" = @{
       rank=10
       filename = "extra_tools\cloudegate.zip"
       copyto = "modules\BITools\cloudegate\"
       update = "no" 
    }

    #>

}

$autofiles=$autofilesraw.GetEnumerator()|sort {$_.value.rank}|%{@{$_.Key = $_.Value}} ### hashtable sorting ###

#$autopath="\\192.168.2.24\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool" 
$autopath="Y:\Matagorda\07.Tool\_AutoTool"

$copyfiles=($autofiles.Keys)

foreach ($copyfile in $copyfiles){

$fileinfo=$autofiles.$copyfile
$filename=$autopath+"\"+$fileinfo.filename
$copytopath=$dest+"\"+$fileinfo.copyto
$updateflag=$fileinfo.update

if(((Get-ChildItem $filename).FullName).length -gt 0){

if($shortcutadd -eq $true -or $updateflag -match "yes"){

if($filename -match "\.zip"){   

    if(!(test-path $copytopath)){new-item -ItemType directory $copytopath |Out-Null}
   
    write-host " Item: $copyfile, unzip $filename  to $copytopath"
    $filename=(Get-ChildItem $filename).FullName
    $shell.NameSpace($copytopath).copyhere($shell.NameSpace($filename).Items(),16)

}

else{
    if($copytopath.Substring($copytopath.Length-2,2) -match "\*\\"){$copytopath=(Get-Item $copytopath).FullName}
    write-host " Item: $copyfile, copy $filename  to $copytopath"
    copy-item $filename -Destination $copytopath -Recurse -Force
    
}

}

}

}

## create shortcut

if($shortcutadd -eq $true){

New-Item -ItemType SymbolicLink -Path $env:USERPROFILE\desktop\ -Name "logs" -Value "C:\testing_AI\logs\" -force -ErrorAction SilentlyContinue | out-null

#New-Item -ItemType SymbolicLink -Path $env:USERPROFILE\desktop\ -Name "testing_AI_link" -Value "C:\testing_AI\" -force -ErrorAction SilentlyContinue | out-null
#New-Item -ItemType SymbolicLink -Path $env:USERPROFILE\desktop\ -Name "AutoRun.bat" -Value "C:\testing_AI\StartAutoRun.bat" -force  -ErrorAction SilentlyContinue| out-null

New-Item -ItemType SymbolicLink -Path $env:USERPROFILE\desktop\ -Name "STOP.bat" -Value "C:\testing_AI\StopAutoRun.bat" -force -ErrorAction SilentlyContinue | out-null
New-Item -ItemType SymbolicLink -Path $env:USERPROFILE\desktop\ -Name "dell_Handy.bat" -Value "C:\testing_AI\dell_Handy.bat" -force -ErrorAction SilentlyContinue | out-null


}


## close all app windows 

Get-Process |   Where-Object { $_.MainWindowHandle -ne 0  } |
  ForEach-Object { 
  $handle = $_.MainWindowHandle

  # minimize window
  $null = [MyNamespace.MyType]::ShowWindowAsync($handle, 2)
}

$spenttime= (New-TimeSpan -start $starttime -end (Get-Date)).TotalMinutes
$spenttime2=[math]::Round($spenttime,1)

$index= "Copy Completed with $spenttime2 minutes." 
$results="copy OK"

write-host $index

#[System.Windows.Forms.MessageBox]::Show($this, " Copy Completed with $spenttime2 minutes. {0}{0} Please start AI testing with 【AutoRun.bat】 at deskeop {0}{0} (you may remove USB disk now) " -f [environment]::NewLine)    


}

 ## disable wu
 
# set the Windows Update service to "disabled"
sc.exe config wuauserv start=disabled

# display the status of the service
#sc.exe query wuauserv

# stop the service, in case it is running
sc.exe stop wuauserv
#stop Background Intelligent Transfer Service (BITS) 
sc.exe stop BITS
# double check it's REALLY disabled - Start value should be 0x4
$checkreg=REG.exe QUERY HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\wuauserv /v Start 

if($checkreg -match "0x4"){
 write-host " disable windows update OK"
 }
 else{
  write-host " disable windows update NG"
  }


 ### close file explore windows

 $shell.Windows() |Where-Object{$_.name -eq "File Explorer"}| ForEach-Object { $_.Quit() }

 ## close all app windows except cmd/windows terminal
 
 $cmdpid=(get-process -name cmd -ea SilentlyContinue|sort starttime|select -last 1).Id
 $wtpid=(get-process -name WindowsTerminal  -ea SilentlyContinue|sort starttime|select -last 1).Id

if($cmdpid){
Get-Process |   Where-Object { $_.MainWindowHandle -ne 0 -and $_.id -ne $cmdpid -and $_.id -ne $wtpid  } |
  ForEach-Object { 
  $handle = $_.MainWindowHandle

  # minimize window
  $null = [MyNamespace.MyType]::ShowWindowAsync($handle, 2)
}
}


### define path ##

$dest="C:\testing_AI\"
$idracf="C:\temp_aitest\tester.txt"


## for  new or update copy to temp, exclude real-initial after reboot###

### copy initial needed files ###

if(!(test-path C:\temp_aitest\)){
New-Item -ItemType directory C:\temp_aitest\ -Force |Out-Null
gci \\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\_Dell_AITest_Start\AutoStart* |Copy-Item -Destination C:\temp_aitest -Force
}

## record logs

start-Transcript -append -path C:\temp_aitest\scriptlogs.txt |Out-Null

### for 1st time initial ##


if(!(test-path  $dest) -and !(test-path $idracf)){
              

### message ###

$line1 = "Good Day! This is Allion Dell-Inventech TC Auto Testing tool."
#$line2 ="May I have your Name or Employee ID:"
#$line_time2 ="Time setting is OK"
#$line_idrac  ="Now we setup iDRAC link information"
#$line_idrac2 ="iDRAC Infomation setting is OK"
#$line_system  ="Now we check the system environment"
#$line_system2  ="The system environment check is OK"
$line_display  ="First we check the display if support 1920*1080 above"
$line_display2  ="The display check ok"
#$line_copy = "Please wait for a while for data copying ... "
$line_reboot = "System is going to reboot in 10s ... "

#region cmd and windows termial settings for windows 11 ###

 $Version = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'

if($Version.CurrentBuildNumber -ge 22000){

start-process wt.exe -WindowStyle Minimized
start-sleep -s 5
Stop-Process -Name  WindowsTerminal -Force
start-sleep -s 1

## wt json file ##

$wtsettings=get-content "$env:USERPROFILE\AppData\Local\Packages\Microsoft.WindowsTerminal*\localState\settings.json"

function ConvertTo-Hashtable {
    [CmdletBinding()]
    [OutputType('hashtable')]
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process {
        ## Return null if the input is null. This can happen when calling the function
        ## recursively and a property is null
        if ($null -eq $InputObject) {
            return $null
        }

        ## Check if the input is an array or collection. If so, we also need to convert
        ## those types into hash tables as well. This function will convert all child
        ## objects into hash tables (if applicable)
        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string]) {
            $collection = @(
                foreach ($object in $InputObject) {
                    ConvertTo-Hashtable -InputObject $object
                }
            )

            ## Return the array but don't enumerate it because the object may be pretty complex
            Write-Output -NoEnumerate $collection
        } elseif ($InputObject -is [psobject]) { ## If the object has properties that need enumeration
            ## Convert it to its own hash table and return it
            $hash = @{}
            foreach ($property in $InputObject.PSObject.Properties) {
                $hash[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
            }
            $hash
        } else {
            ## If the object isn't an array, collection, or other object, it's already a hash table
            ## So just return it.
            $InputObject
        }
    }
}

$wshash=$wtsettings | ConvertFrom-Json | ConvertTo-HashTable
$guidcmd=($wshash.profiles.list|?{$_.name -match "command"}).guid

$wtsettings|%{

if($_ -match "defaultProfile"){
$_= """defaultProfile"": ""$guidcmd"","
}
$_

}|set-content "$env:USERPROFILE\AppData\Local\Packages\Microsoft.WindowsTerminal*\localState\settings.json"

###　default app settings
## https://support.microsoft.com/en-us/windows/command-prompt-and-windows-powershell-for-windows-11-6453ce98-da91-476f-8651-5c14d5777c20

$RegPath = 'HKCU:\Console\%%Startup'
if(!(test-path $RegPath)){New-Item -Path $RegPath}
$RegKey = 'DelegationConsole'
$RegValue = '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
set-ItemProperty -Path $RegPath -Name $RegKey -Value $RegValue -Force | Out-Null

$RegKey2 = 'DelegationTerminal'
$RegValue2 = '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
set-ItemProperty -Path $RegPath -Name $RegKey2 -Value $RegValue2 -Force | Out-Null
}

#endregion

#region show cmd window ##

 $id=(Get-Process cmd -ErrorAction SilentlyContinue|Where-Object{($_.MainWindowTitle).length -gt 0}).Id

 if($id){
[Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
}
#endregion 

Clear-Host

#region greeting ##
& outwords $line1 30
write-host ""
write-host ""
#endregion 

#region display check and setup

& outwords $line_display 20
Write-Host ""
Write-Host ""

## check max resolution ## ## $maxx $maxy2

#$horlist=wmic /namespace:\\ROOT\WMI path WmiMonitorListedSupportedSourceModes get MonitorSourceModes /format:list |?{$_ -match "HorizontalActivePixels"}
#$verlist=wmic /namespace:\\ROOT\WMI path WmiMonitorListedSupportedSourceModes get MonitorSourceModes /format:list |?{$_ -match "VerticalActivePixels"}

$displaySettings = Get-WmiObject -Namespace root\cimv2 -Class Win32_DesktopMonitor
$extendedMode = $displaySettings.Count -gt 1

if ($extendedMode) {

write-host "Extent display if multi-display"
displayswitch.exe  /extend
Start-Sleep -s 10

}

$horlist=((Get-WmiObject -N "root\wmi" -Class WmiMonitorListedSupportedSourceModes |Where-Object{$_.active -eq "Ture"}).MonitorSourceModes).HorizontalActivePixels
$verlist=((Get-WmiObject -N "root\wmi" -Class WmiMonitorListedSupportedSourceModes |Where-Object{$_.active -eq "Ture"}).MonitorSourceModes).VerticalActivePixels

$maxx=0
$maxy=0
$rnk=0
foreach($verlist1 in $verlist){
$check1=$verlist1 -match "\d{3,}"
$maxy1=$Matches[0]

if([int]$maxy1 -ge [int]$maxy){
  $maxy=[int]$maxy1

  $check2=$horlist[$rnk] -match "\d{3,}"
  $maxx1=$Matches[0]
  if([int]$maxx1 -ge [int]$maxx){
   $maxx=[int]$maxx1
    $maxy2=$maxy
   $maxres= "$maxx|$maxy2"
   }
   
  }

  $rnk++
}

if($maxx -lt 1920 -or $maxy2 -lt 1080){
 $ans1= [Microsoft.VisualBasic.Interaction]::MsgBox("Disply not support 1920*1080 above, if continue? ",'YesNo,SystemModal,Information', 'check')
 if( $ans1 -eq "No"){exit}
}

& outwords $line_display2 20
Write-Host ""
Write-Host ""
#endregion display check and setup
      

###########  Check Time/Revise ########
  
  $id="Taipei"
  $time1=get-date -Format "yy/MM/dd HH:mm"
  $desid=(Get-TimeZone *).id -match $id
  $timezone=(Get-TimeZone -ListAvailable|?{$_.id -match "$id" }).id
  Set-TimeZone -id $timezone |Out-Null
  ## sync server time ##
  $settime_syncserver =  net time \\192.168.2.249 /set /yes
  $currentime=get-date -format "yyyy/MM/dd HH:mm:ss"
  
Write-Host "Check System settings and choose the flow to run"
$rungui= .\dellgui.ps1

# if close gui

if(!(test-path $idracf)){

exit

}

#region change powersetting incase enter S4

$powerlog1="$env:userprofile\desktop\powercfg_before.txt"
$powerlog2="$env:userprofile\desktop\powercfg_after.txt"

#collect original powerset
$poid= (powercfg /getactivescheme).split(" ")|Where-Object {$_ -match "-"}
$psettings=powercfg /q $poid

$l=0
$psettings|ForEach-Object{
$l++
if($_ -match "after" ){$_line=$psettings.IndexOf($_)}
$nowlinegap= $l - $_line

if( $nowlinegap -ge 0 -and $nowlinegap -le 10 -and ($_ -match "index" -or$_ -match "after" )){
 $checks=$checks+@($_)
}
}

set-content -path $powerlog1 -value  $checks -Force

#change powerset

powercfg /change monitor-timeout-ac 0
powercfg -change -standby-timeout-ac 0
powercfg /x -hibernate-timeout-ac 0
powercfg /x -disk-timeout-ac 0

if($acdc.Length -gt 0){
powercfg /change monitor-timeout-dc 0
powercfg -change -standby-timeout-dc 0
powercfg /x -hibernate-timeout-dc 0
powercfg /x -disk-timeout-dc 0
}

powercfg.exe /SETACTIVE SCHEME_CURRENT

#collect after change powerset
$checks.clear()
$poid= (powercfg /getactivescheme).split(" ")|Where-Object {$_ -match "-"}
$psettings=powercfg /q $poid

$l=0
$psettings|%{
$l++
if($_ -match "after" ){$_line=$psettings.IndexOf($_)}
$nowlinegap= $l - $_line

if( $nowlinegap -ge 0 -and $nowlinegap -le 10 -and ($_ -match "index" -or$_ -match "after" )){
 $checks=$checks+@($_)
}
}

set-content -path $powerlog2 -value  $checks -Force

#endregion change powersetting incase enter S4

#region turn off notification
New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\PushNotifications" -Name "ToastEnabled" -Value 0 -PropertyType DWORD -Force
#endregion

#region create a shortcuts ##
$SourceFilePath = "\\192.168.2.249\srvprj\Inventec\Dell"
$ShortcutPath = "$env:userprofile\Desktop\dell.lnk"
$WScriptObj = New-Object -ComObject ("WScript.Shell")
$shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
$shortcut.TargetPath = $SourceFilePath
$shortcut.WindowStyle = 1
$shortcut.Save()
#endregion create a shortcuts ##


#region new name the system by username ##
  $newname="$env:USERNAME"
  $computer = Get-WmiObject -Class Win32_ComputerSystem
  $newaname=$computer.rename($newname)
#endregion new name the system by username ##

 $ipv4ip= (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.PrefixOrigin -eq 'dhcp' -and $_.IPAddress -match "192\.168\.2\." }| select -First 1).IPAddress
 #set-content C:\temp_aitest\idrac.txt -value "$idracip,$idracusern,$idracpasswd"
 $idracip="na"
 if(test-path C:\temp_aitest\idrac.txt){
  $idracip=((get-content  C:\temp_aitest\idrac.txt).split(","))[0]
 } 
 $idname=get-content "C:\temp_aitest\tester.txt"
 $timestart=get-date -Format "yyyy/M/dd HH:mm:ss"
 $addstartlog="\\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool_Monitor\start_records.txt"

 add-content -path $addstartlog -value "$timestart|$idname|$env:USERNAME|$idracip|$ipv4ip" -Force   

  $setting=Get-ItemPropertyValue -Path HKLM:SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity -Name Enabled -ErrorAction SilentlyContinue

  <##this message moved to autostart0#
   if( $setting -ne 1){
   [Microsoft.VisualBasic.Interaction]::MsgBox(" 請手動 turn On Memory Integrity`n 並依照Windows指示Reboot `n (Reboot後會再繼續自動執行排定流程)",'OKOnly,SystemModal,Information', 'check')
  }
  ##>

#region check if need setup memory integrity ###

#region setup task schedule at logon ##

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_Run" -f' 
start-sleep -s 5

$timeset=[double]1
$TimeSpan = New-TimeSpan -Minutes $timeset
$action = New-ScheduledTaskAction -Execute "C:\temp_aitest\AutoStart.bat"
$trigger = New-JobTrigger -AtLogOn -RandomDelay $TimeSpan #00:05:00
$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$user=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$STPrin= New-ScheduledTaskPrincipal   -User $user  -RunLevel Highest

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_Run" -Principal $STPrin

   start-sleep -s 5

   $taskready =(Get-ScheduledTask | Where-Object {$_.TaskName -like "Auto_Run" } ).State
   if($taskready -eq "Ready"){ $results  ="Task schedule set up OK"} else{$results  ="Task schedule set up NG"}
   $results

#endregion task schedule

if( $setting -ne 1){

 # Write-Host ""
 #& outwords "Setup Memory integrity need manual action, please change core isolation settings to ""ON"" and follow step to reboot"  20
 #Write-Host ""
 # & outwords  "And you are free to go after reboot the system. See you next time. "  20

  stop-Transcript |Out-Null
    #$Error > C:\temp_aitest\logserr.txt

  explorer.exe windowsdefender://coreisolation

start-sleep -s 10

 $id=((Get-Process *)|?{$_.MainWindowTitle -match "Windows Security"}).Id
 start-sleep -s 5
 [Microsoft.VisualBasic.interaction]::AppActivate($id)|out-null
  
 exit

 }

else{
   Write-Host ""
 & outwords "Initial checks complete! Auto tool will take charge of following tests since now on."  20
 Write-Host ""
  & outwords  "You are free to go. See you next time. "  20
   Write-Host ""
  & outwords  $line_reboot  20
  

  #region  dns settings change to 127.0.0.1

$getinfo=ipconfig
$checkline=$getinfo -match "allion.test"
$dnsip="127.0.0.1"

if($checkline){
$linenu=$getinfo.IndexOf($checkline)
($getinfo|Select-Object -Skip $linenu|Select-Object -First 3|Select-Object -last 1) -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}" |Out-Null
$ipout=$matches[0]

Get-NetIPAddress|ForEach-Object{

if($_.IPAddress -match $ipout -and $_.AddressFamily -eq "IPv4"){
$adtname=$_.InterfaceAlias 
}
}

Get-DnsClientServerAddress -InterfaceAlias $adtname|Where-Object{$_.AddressFamily -eq "2"}

if($dnsip.Length -gt 2){
## change dns
Set-DnsClientServerAddress -InterfaceAlias $adtname -ServerAddresses $dnsip
Clear-DnsClientCache
start-sleep -s 30
$dnsresults=Get-DnsClientServerAddress -InterfaceAlias $adtname|Where-Object{$_.AddressFamily -eq "2"}|Out-String

if($dnsresults -match $dnsip){
write-host "DNS changed to $($dnsip)"
}
else{
write-host "Fail to change DNS to $($dnsip)"
}

} 


}

#endregion


  stop-Transcript |Out-Null
    #$Error > C:\temp_aitest\logserr.txt

<#region handle \ReportingEvents.log

   if(test-path "C:\Windows\SoftwareDistribution\ReportingEvents.log"){
        $timenow=get-date -format "yyMMdd_HHmmss"
        if(!(test-path "C:\temp_aitest\ReportingEvents\backups\")){
        new-item -ItemType directory "C:\temp_aitest\ReportingEvents\backups\"|Out-Null
        }
        $bkreportingeventlog="C:\temp_aitest\ReportingEvents\backups\"+"ReportingEvents_$($timenow).log"
        Copy-Item "C:\Windows\SoftwareDistribution\ReportingEvents.log" -Destination $bkreportingeventlog -Force
        }
        remove-item C:\Windows\SoftwareDistribution\* -Recurse -Force -ErrorAction SilentlyContinue
#>

  shutdown /r /t 0

  }
      
#endregion 


}


if((!(test-path $dest) -and (test-path $idracf)) -or ((test-path  $dest) -and !(test-path $idracf))){  ## the 1st circumstance is real-initial after reboot and the 2nd circumstance is for update 

if(!(test-path $dest) -and (test-path $idracf)){

#region setting remote control

Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -name "fDenyTSConnections" -value 0
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -name "fSingleSessionPerUser" -value 0
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services" -Name "Shadow" -Value 2 -Type "DWORD"
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp" -Name "Shadow" -Value 2 -PropertyType DWord　-ErrorAction SilentlyContinue
# Get the Security Identifier (SID) for the current user
$currentUser = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, $env:USERNAME)
$currentUserSid = $currentUser.Translate([System.Security.Principal.SecurityIdentifier]).Value
$remoteControlSDDL = "O:BAG:BAD:(A;;0xf0007;;;$currentUserSid)"
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp" -Name "SecurityDescriptor" -Value $remoteControlSDDL
gpupdate /force

Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -name "LimitBlankPasswordUse" -value 0
Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
Enable-NetFirewallRule -DisplayName "File and Printer Sharing (SMB-In)"
Enable-NetFirewallRule -DisplayName "Remote Desktop - Shadow (TCP-In)"
New-NetFirewallRule -DisplayName "Session Shadowing Ports" -Direction Inbound -Protocol TCP -LocalPort 139,445,49152-65535 -Action Allow

#endregion

}

do{
Start-Sleep -s 10
$testping= ping 192.168.2.249 /n 3
write-host "connecting to 192.168.2.249 ..."
}until ($testping -match "Reply from")

if(!(test-path "C:\temp_aitest\flowsettings.csv")){
$gettime=(get-date)
Write-Host "Please choose run script for TC."
$rungui= .\dellgui.ps1
$checktime=(gci "C:\testing_AI\settings\flowsettings.csv").LastWriteTime
if($checktime -gt $gettime){
Write-Host "Choose complete, save this flow."
Write-Host ""
}
else{
exit
}
}

#region ## disconnect internet but pass cycling tool

 if( (test-path "C:\testing_AI\") -and !($env:computername -match "cycling") ){

$netdisc="net_disconnecting"
Get-Module -name $netdisc|remove-module
$mdpath=(gci -path C:\testing_AI\modules\ -r -file |?{$_.name -match "^$netdisc\b" -and $_.name -match "psm1"}).fullname
Import-Module $mdpath -WarningAction SilentlyContinue -Global

&$netdisc -para1 "internet" -para2 "nolog"

  }

  #endregion

initial_copytool

if(test-path "C:\temp_aitest\flowsettings.csv" ){move-Item "C:\temp_aitest\flowsettings.csv"  -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
if(test-path "C:\temp_aitest\idrac.txt"){move-Item "C:\temp_aitest\idrac.txt" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
if(test-path "C:\temp_aitest\tester.txt"){move-Item "C:\temp_aitest\tester.txt" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
 
$updatefolder="\\192.168.2.249\srvprj\Inventec\Dell\Matagorda\07.Tool\_AutoTool\_Dell_AITest_Start"
if(test-path "$updatefolder\tc_flowsettings.csv"){copy-Item "$updatefolder\tc_flowsettings.csv" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
if(test-path "$updatefolder\agile_account.txt"){copy-Item "$updatefolder\agile_account.txt" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
if(test-path "$updatefolder\gui_functions.csv"){copy-Item "$updatefolder\gui_functions.csv" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}
if(test-path "$updatefolder\mainflow.ps1"){copy-Item "$updatefolder\mainflow.ps1" -Destination $dest\ -Force -ErrorAction SilentlyContinue}
if(test-path "$updatefolder\dellgui.ps1"){copy-Item "$updatefolder\dellgui.ps1" -Destination $dest\ -Force -ErrorAction SilentlyContinue}
if(test-path "$updatefolder\loghtmlcheck.ps1"){copy-Item "$updatefolder\loghtmlcheck.ps1" -Destination $dest\settings\ -Force -ErrorAction SilentlyContinue}

#if(test-path "C:\temp_aitest\scriptlogs.txt"){move-Item "C:\temp_aitest\scriptlogs.txt" -Destination $dest\logs\ -Force}
dir $dest -Recurse| Unblock-File
#dir $dest\*\* | Unblock-File
#dir $dest\*\*\* | Unblock-File

#region setting WinRM

$checkrm=(Get-Service -Name "*WinRM*" | Select-Object status).status
if($checkrm -ne "running"){
winrm qc -force
#Enable-PSRemoting -Force
Enable-PSRemoting -SkipNetworkProfileCheck -Force
}

Set-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)" -RemoteAddress Any
Enable-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)"

#endregion

## scehdule auto_run after 1 min ##

$timeset=[double]1

start-process cmd -ArgumentList '/c schtasks /delete /TN "Auto_Run" -f' 
start-sleep -s 5

$action = New-ScheduledTaskAction -Execute "C:\testing_AI\AutoRun.bat"
$etime=(Get-Date).AddMinutes($timeset)
$trigger = New-ScheduledTaskTrigger -Once -At $etime 
$Stset = New-ScheduledTaskSettingsSet -Priority 0 -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$user=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$STPrin= New-ScheduledTaskPrincipal   -User $user  -RunLevel Highest

Register-ScheduledTask -Action $action -Trigger $trigger -Settings $Stset -Force -TaskName "Auto_Run" -Principal $STPrin

   start-sleep -s 10

   $taskready =(Get-ScheduledTask | Where-Object {$_.TaskName -like "Auto_Run" } ).State
    if($taskready -eq "Ready"){ $results  ="Task schedule set up OK"} else{$results  ="Task schedule set up NG"}
    $taskready
  
  #region  rollback dns settings
  
  $getinfo=ipconfig
$checkline=$getinfo -match "allion.test"

if($checkline){
$linenu=$getinfo.IndexOf($checkline)
($getinfo|Select-Object -Skip $linenu|Select-Object -First 3|Select-Object -last 1) -match "\d{1,}\.\d{1,}\.\d{1,}\.\d{1,}" |Out-Null
$ipout=$matches[0]

Get-NetIPAddress|%{

if($_.IPAddress -match $ipout -and $_.AddressFamily -eq "IPv4"){
$adtname=$_.InterfaceAlias 
}
}


## rollback dns settins
Set-DnsClientServerAddress -InterfaceAlias $adtname -ResetServerAddresses
Clear-DnsClientCache
start-sleep -s 30
$dnsresults=Get-DnsClientServerAddress -InterfaceAlias $adtname|Where-Object{$_.AddressFamily -eq "2"}|Out-String

if($dnsresults -match "168"){
write-host "DNS rollback ok"
}
else{
write-host "Fail to rollback DNS settings"
}


}
     #endregion    


  stop-Transcript |out-null

if(!(test-path $dest\logs\scriptlogs)){new-item -ItemType directory -path $dest\logs\scriptlogs\ -Force |Out-Null}
if(test-path "C:\temp_aitest\scriptlogs.txt"){move-Item "C:\temp_aitest\scriptlogs.txt" -Destination $dest\logs\scriptlogs\ -Force}
 
    }
 
 



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#$wshell=New-Object -ComObject wscript.shell
#$shell=New-Object -ComObject shell.application

# Create a Ping object
$ping = New-Object System.Net.NetworkInformation.Ping

#testernamefle
$testername_path="C:\temp_aitest\tester.txt"
$testername_path2="C:\testing_AI\settings\tester.txt"
$testernm="(required)"

 if(test-path $testername_path2){
 $testernm=get-content $testername_path2
 }

#yellowbang
$ye=Get-WmiObject Win32_PnPEntity|Where-Object{$_.ConfigManagerErrorCode -ne 0}|Select-Object Name,Description, DeviceID, Manufacturer
if($ye.DeviceID.Count -gt 0){
Start-Process devmgmt.msc
}

#check idrac if exist
# set-content C:\temp_aitest\idrac.txt -value "$idracip,$idracusern,$idracpasswd"
 $idracpath1="C:\temp_aitest\idrac.txt"
 $idracpath2="C:\testing_AI\settings\idrac.txt"

  $idrac_ip="192.168.2."
   $idrac_name="root"
    $idrac_pwd="calvin"

 if(test-path $idracpath2){
 $idrac_content=(get-content $idracpath2).split(",")
  $idrac_ip=$idrac_content[0]
   $idrac_name=$idrac_content[1]
    $idrac_pwd=$idrac_content[2]
      $pingresult1 = $ping.Send($idrac_ip, 1000)
        $pingcheck1 = $pingresult1.Status
     }
     
     if( (test-path "C:\testing_AI\settings\") -and !(test-path $idracpath2)){
      $idrac_ip="bypassed"
     
     }

$tcflowpath = (Get-ChildItem  -path .\ -Filter tc_flowsettings.csv -Recurse).FullName
if($tcflowpath.Length -eq 0){
$tcflowpath = (Get-ChildItem  -path C:\testing_AI\settings\ -Filter tc_flowsettings.csv -Recurse).FullName
}

$con = import-csv $tcflowpath
$totalRows = $con.Length
$TC = @()
for($i=1;$i -le $totalRows;$i++){
    if(($con[$i].TC -ne "ini") -and ($con[$i].TC -ne "fin") -and ($con[$i].TC -ne "")){
        $TC += $con[$i].TC
    }    
}

$TClist = $TC | Sort-Object -Unique

function doFunctest1($func, $tabname,$indexg) { 
write-host "$func, $tabname,$indexg" 
}

 #region # add buttons 
function addbuttons ($i) {
     #$taba
    
     if($butc){
        $buttonArr = New-Object System.Windows.Forms.Button[]($butc)
     }
     

     if($taba -match "TC"){$x0=20;$y0=60}
     else{$x0=20;$y0=30}
 
     $j=$x1=$y1=0

    foreach($btt in $bts){

      #$tab[$i].Controls.Remove($buttonArr)

      if($j -ne 0){
      $x1=[int][Math]::Floor( $j / 7 )
      $y1=( $j % 7 )
      } 


     $x1p=$x0 + [System.Convert]::ToInt32($x1) * 150
     $y1p=$y0 + [System.Convert]::ToInt32($y1) * 50
    
     $button = New-Object System.Windows.Forms.Button
     $button.Location = New-Object System.Drawing.Point($x1p, $y1p)
     
     $button.Text=$btt."Button"
     
     $button.Size = New-Object System.Drawing.Size(120, 40)
     $buttonArr[$j] = $button

     $buttonArr[$j].Add_Click({
     $Form.WindowState = 'Minimized'
     #$indexg=$buttonArr.IndexOf($this)
     $indexg = $args[0].Parent.Controls.IndexOf($args[0])
     $index =($args[0]).text  # Get the index of the clicked button
     $tabName = ($args[0]).Parent.Text
     doFunctest1 $index  $tabName $indexg 
     doFunc $index  $tabName    # Call the function and pass the index as a parameter
     #doFunctest1 $index  $tabName $indexg
    })
     
     #Add Tip info
     $Tip = $btt.button_index
     $TipTemp = New-Object System.Windows.Forms.ToolTip
     $TipTemp.SetToolTip($buttonArr[$j], $Tip)

     ## add query button in TC tab ##
     $tab[$i].Controls.Add($buttonArr[$j])
         
     $j++
     #$k++
    
     }
   
}

function doFunctest($que) { 

$bts=$btlists|?{($_.flag).split("`n") -match "^TC" -and ($_.flag).split("`n")  -match  $que}
$btscon=($bts.Button).Count
write-host "$que $btscon" 
    foreach($btt in $bts){
     $funcname=$btt."Function"
    write-host "$funcname" 
    }
$buttonArr = New-Object System.Windows.Forms.Button[](($bts.Button).Count)


}

function doFunc2 ($que,$i) { 

#write-host "$que,$i"
$taba="TC"

if($btlists){
    $bts=$btlists|?{($_.flag).split("`n") -match "^TC" -and ($_.flag).split("`n")  -match  $que}
}


addbuttons $i
  
  # $Form.Controls.Add($TabControl)
    $form.Refresh() | Out-Null

}

function doFunc($func, $tabname) { 

      $btt=$btlists|?{$_.Button -eq $func -and $_.flag -match "$tabname"}|select -First 1

      $funcname=$btt."Function"
      $para1=$btt."para1"
      $para2=$btt."para2"
      $para3=$btt."para3"
      $para4=$btt."para4"
      $para5=$btt."para5"
      
    write-host "run $funcname -para1 $para1 -para2 $para2 -para3 $para3 -para4 $para4 -para5 $para5"

        Get-Module -name $funcname|remove-module
        $mdpath=(Get-ChildItem -path C:\testing_AI\modules\* -r -file |Where-Object{$_.name -match "^$funcname\b" -and $_.name -match "psm1"}).fullname
        Import-Module $mdpath -WarningAction SilentlyContinue -Global
        & $funcname -para1 $para1 -para2 $para2 -para3 $para3 -para4 $para4 -para5 $para5
       # [System.Windows.Forms.MessageBox]::Show($this,"Function Run Complete!")
        
}

#endregion

##btlists,catch csv data for button info##
##butc,button count##
##button type##


# Create the Main form

$Form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 985, 720
    Text          = "AI Test Tool @Allion"
    Topmost       = $true
}

# Create the tab control
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$TabControl.SelectedIndex = 2

if($path2){
    $path2 = (Get-ChildItem  -path .\ -Filter gui_functions.csv -Recurse) 
    $btlists=import-csv $path2.PSPath
    $butc=(($btlists.flag).split("`n")).count
    $types=($btlists.flag).split("`n")|Sort-Object|Get-Unique|Where-Object{$_ -notmatch "^TC"}
    $tabs = $types+@("TC")
}else{
    $tabs = @("TC")
}

$tab = New-Object System.Windows.Forms.TabPage[]($tabs.Count)
$Label = New-Object System.Windows.Forms.Label[]($tabs.Count)
$i=$k=0

#dynamic control create

foreach($taba in $tabs){
    $tab[$i] = New-Object System.Windows.Forms.TabPage
    $tab[$i].Text = $taba
    $TabControl.TabPages.Add($tab[$i])
    $Label[$i] = New-Object System.Windows.Forms.Label
    #$Label[$i].Text = $taba
    $Label[$i].AutoSize = $true
    $tab[$i].Controls.Add($Label[$i])
    
    if($path2){
        if( $taba -match "TC") {
            $bts=$btlists|Where-Object{($_.flag).split("`n") -match "^$taba"}}
        else{ 
            $bts=$btlists|Where-Object{($_.flag).split("`n") -match "^$taba\b"}
        }
    }
    
    addbuttons $i
        
    $i++
}

$InputBox = New-Object System.Windows.Forms.TextBox
$InputBox.Location = New-Object System.Drawing.Point(50, 250)
$InputBox.Size = New-Object System.Drawing.Size(150,30)
$InputBox.Text = "請輸入搜尋關鍵字..."
$InputBox.ForeColor = "gray"

$InputBox.Add_Click({
    $InputBox.Text = ""
    $InputBox.ResetForeColor()
})

$InputBox.Add_TextChanged({

    $index = $TabControl.TabPages.IndexOf($args[0].Parent)
    $InputText = $InputBox.Text
    $tabName = ($args[0]).Parent.Text
    $tabid=$tabs.indexof($tabName)
    doFunc2 $InputText $index
    #$InputBox.Text = ""


    if($InputText -eq ""){
        $originalOptions = $TClist | Where-Object { $_ -notin $enabledList.Items }
        
        $disabledList.Items.Clear()
        $disabledList.Items.AddRange($originalOptions)
    }else{
        $filteredOptions = $TClist | Where-Object { $_ -like "*$InputText*" -and $_ -notin $enabledList.Items }

        $disabledList.Items.Clear()

        if($filteredOptions -ne $null){
            $disabledList.Items.AddRange($filteredOptions)
        }       
    }

    if($disabledList.Items -ne $null){
        $searchTemp.Text = '搜尋關鍵字"' + $InputBox.Text + '",' + "已搜尋到" + $disabledList.Items.Count +"筆"
    }else{
        $searchTemp.Text = "該關鍵字沒有任何資料"
    }
})


$searchTemp = New-Object System.Windows.Forms.Label
$searchTemp.Location = New-Object System.Drawing.Point(210, 250)
$searchTemp.Text = ""
$searchTemp.AutoSize = $true
$searchTemp.Font = New-Object System.Drawing.Font($searchTemp.Font.FontFamily, 10)
# Add controls to the form
$tab[$i-1].Controls.Add($InputBox)
#$tab[$i-1].Controls.Add($SendButton)
$tab[$i-1].Controls.Add($searchTemp)

#----------------------------------
    
#### disabled listbox
    $disabledList = New-Object system.Windows.Forms.ListBox
    $disabledList.location = New-Object System.Drawing.Point(10,275)
    $disabledList.width    = 395
    $disabledList.height   = 360
    $disabledList.SelectionMode = 'MultiExtended'
    $TClist | % { $disabledList.Items.Add($_) | Out-Null }
    $tab[$i-1].Controls.Add($disabledList)

#### enabled listbox
    $enabledList = New-Object system.Windows.Forms.ListBox
    $enabledList.location = New-Object System.Drawing.Point(540,275)
    $enabledList.width    = 395
    $enabledList.height   = 360
    $enabledList.SelectionMode = 'MultiExtended'
    $tab[$i-1].Controls.Add($enabledList)


  #region  #form  name info

    $nameinfo = New-Object System.Windows.Forms.Label
    $nameinfo.Location = New-Object System.Drawing.Point(5,15)
    $nameinfo.AutoSize = $True
    $nameinfo.Text = "Your Name or ID:"
    $nameinfo.ForeColor="Blue"
    $nameinfo.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 14)
    $tab[$i-1].Controls.Add( $nameinfo)

    $nameinfo2 = New-Object System.Windows.Forms.Label
    $nameinfo2.Location = New-Object System.Drawing.Point(400,15)
    $nameinfo2.AutoSize = $True  
    $nameinfo2.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 14)
    $nameinfo2.ForeColor="Red"
    $tab[$i-1].Controls.Add( $nameinfo2)
    

    $nameinput = New-Object System.Windows.Forms.TextBox
    $nameinput.Location = New-Object System.Drawing.Point(160,15)
    $nameinput.size= New-Object System.Drawing.Size(200,60)
    $nameinput.Text = $testernm
    if($testernm -match "required"){$nameinput.ForeColor = "gray"}
    $nameinput.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $nameinput.TextAlign = "center"
    $tab[$i-1].Controls.Add($nameinput)

    $nameinput.Add_Click({
    $nameinput.Text = ""
    $nameinput.ResetForeColor()
})
 
 #endregion

    
#region ## get system info

$syslines=@()
 $timenow= (get-date -Format "yyyy-M-d HH:mm") + " / " +(Get-TimeZone).id
  $timenow=$timenow.ToString()
    $syslines=$syslines+@("Current System Time:$($timenow)")
    
#$BIOS=systeminfo | findstr /I /c:bios
$BIOS=(Get-CimInstance -ClassName Win32_BIOS).SMBIOSBIOSVersion
  $syslines= $syslines+@("BIOS:$($BIOS)")

  $cpus=@() 
Get-CimInstance -ClassName Win32_Processor|%{
 $cpu=$_.DeviceID +" : "+ $_.Name
 $cpus=  $cpus+@($cpu) 

  }
   $syslines= $syslines+@(($cpus|Out-String).trim())

$CompObject =  Get-WmiObject -Class WIN32_OperatingSystem
$RAM = (($CompObject.TotalVisibleMemorySize - $CompObject.FreePhysicalMemory)/1024)
$RAM2 = [math]::Round(($CompObject.TotalVisibleMemorySize - $CompObject.FreePhysicalMemory)/1024/1024)

$RAM = (Get-WmiObject -class "cim_physicalmemory" | Measure-Object -Property Capacity -Sum).Sum /1024/1024
$RAM2=$RAM/1024

$syslines= $syslines+@("System RAM : $($RAM2) GB ($($RAM) MB)")

$disks=  (Get-WmiObject Win32_PnPSignedDriver|Where-Object{$_.DeviceClass -match "disk"}).FriendlyName

$syslines= $syslines+@("DiskInfo:"+ [string]::Join("|",$disks))
    
$name=(Get-WmiObject Win32_OperatingSystem).caption
 $bit=(Get-WmiObject Win32_OperatingSystem).OSArchitecture
 $Versiona=(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name DisplayVersion).DisplayVersion
  $Version = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\'
   $OScheck=" $name, $bit, $Versiona (OS Build $($Version.CurrentBuildNumber).$($Version.UBR))"
 
 $syslines= $syslines+@("OS Version: $($OScheck)")
  
  $comutername=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name
     
 $syslines= $syslines+@("Computer Name\User Name: $($comutername)")
 $syslines=$syslines|Out-String
   #endregion
    
   
#region  #form  system info

$sysPanel = New-Object System.Windows.Forms.Panel
$sysPanel.Left = 20
$sysPanel.Top = 80
$sysPanel.Width = 600
$sysPanel.Height = 130
#$sysPanel.BackColor = '255, 255, 255'
$sysPanel.AutoScroll=$true
$sysPanel.BorderStyle="Fixed3d"
 $tab[$i-1].Controls.Add($sysPanel)

    $checkboxsys = New-Object System.Windows.Forms.CheckBox
    $checkboxsys.Location = New-Object System.Drawing.Point(10,60)
    $checkboxsys.Size = New-Object System.Drawing.Size(200,20)
    $checkboxsys.text = "System Info Check OK"
    $tab[$i-1].Controls.Add($checkboxsys)

    $sysinfo = New-Object System.Windows.Forms.Label
    $sysinfo.Location = New-Object System.Drawing.Point(5,2)
    $sysinfo.AutoSize = $True
    $sysinfo.Text = $syslines
    $sysinfo.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
    $sysPanel.Controls.Add( $sysinfo)

    $checkboxsys2 = New-Object System.Windows.Forms.CheckBox
    $checkboxsys2.Location = New-Object System.Drawing.Point(10,220)
    $checkboxsys2.Size = New-Object System.Drawing.Size(150,20)
    $checkboxsys2.text = "Yellow Bang Check"
    $tab[$i-1].Controls.Add($checkboxsys2)

    if($ye.DeviceID.Count -eq 0){
    $checkboxsys2.Checked=$true
    $checkboxsys2.Enabled=$false
    }
        
    $sysinfo2 = New-Object System.Windows.Forms.Label
    $sysinfo2.Location = New-Object System.Drawing.Point(170,220)
    $sysinfo2.AutoSize = $True    
    $sysinfo2.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
   

    if($ye.DeviceID.Count -eq 0){
    $sysinfo2.Text = "No Yellow Bang"
    }
    if($ye.DeviceID.Count -ne 0){
    $sysinfo2.Text = "With Yellow Bang, need check"
    $sysinfo2.ForeColor="Red"
    }
     $tab[$i-1].Controls.Add( $sysinfo2)

#endregion


#region  #form  idrac info

    $checkboxidrac = New-Object System.Windows.Forms.CheckBox
    $checkboxidrac.Location = New-Object System.Drawing.Point(650,60)
    $checkboxidrac.Size = New-Object System.Drawing.Size(200,20)
    $checkboxidrac.text = "iDRAC Info Check OK"
    $tab[$i-1].Controls.Add($checkboxidrac)

    
    $checkboxidrac2 = New-Object System.Windows.Forms.CheckBox
    $checkboxidrac2.Location = New-Object System.Drawing.Point(890,60)
    $checkboxidrac2.Size = New-Object System.Drawing.Size(100,20)
    $checkboxidrac2.text = "Bypass"
    $tab[$i-1].Controls.Add($checkboxidrac2)

$idracPanel = New-Object System.Windows.Forms.Panel
$idracPanel.Left = 650
$idracPanel.Top = 80
$idracPanel.Width = 300
$idracPanel.Height = 150
#$sysPanel.BackColor = '255, 255, 255'
#$idracPanel.AutoScroll=$true
$idracPanel.BorderStyle="Fixed3d"
 $tab[$i-1].Controls.Add($idracPanel)
 
    $idracinfo = New-Object System.Windows.Forms.Label
    $idracinfo.Location = New-Object System.Drawing.Point(5,2)
    $idracinfo.AutoSize = $True
    $idracinfo.Text = "iDRAC IP:"
    $idracinfo.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $idracPanel.Controls.Add($idracinfo)

    $idracinfo2 = New-Object System.Windows.Forms.Label
    $idracinfo2.Location = New-Object System.Drawing.Point(5,42)
    $idracinfo2.AutoSize = $True
    $idracinfo2.Text = "iDRAC username:"
    $idracinfo2.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $idracPanel.Controls.Add($idracinfo2)
    
    $idracinfo3 = New-Object System.Windows.Forms.Label
    $idracinfo3.Location = New-Object System.Drawing.Point(5,82)
    $idracinfo3.AutoSize = $True
    $idracinfo3.Text = "iDRAC password:"
    $idracinfo3.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $idracPanel.Controls.Add($idracinfo3)
        
    $idracinput = New-Object System.Windows.Forms.TextBox
    $idracinput.Location = New-Object System.Drawing.Point(30,2)
    $idracinput.size= New-Object System.Drawing.Size(200,30)
    $idracinput.Text = $idrac_ip
    $idracinput.TextAlign = "Center"
    $idracPanel.Controls.Add($idracinput)
    
    $idracinput2 = New-Object System.Windows.Forms.TextBox
    $idracinput2.Location = New-Object System.Drawing.Point(80,42)
    $idracinput2.size= New-Object System.Drawing.Size(200,30)
    $idracinput2.Text =  $idrac_name
    $idracinput2.TextAlign = "center"
    $idracPanel.Controls.Add($idracinput2)
    
    $idracinput3 = New-Object System.Windows.Forms.TextBox
    $idracinput3.Location = New-Object System.Drawing.Point(80,82)
    $idracinput3.size= New-Object System.Drawing.Size(200,30)
    $idracinput3.Text = $idrac_pwd
    $idracinput3.TextAlign =  "center"
    $idracPanel.Controls.Add($idracinput3)


    $idracinfo4 = New-Object System.Windows.Forms.Label
    $idracinfo4.Location = New-Object System.Drawing.Point(20,110)
    $idracinfo4.AutoSize = $True
    $idracinfo4.Text = ""
    $idracinfo4.ForeColor= "green"
    $idracinfo4.Font= [System.Drawing.Font]::new("Microsoft Sans Serif", 16)
    $idracPanel.Controls.Add($idracinfo4)
    
     if ($pingcheck1 -eq [System.Net.NetworkInformation.IPStatus]::Success) {
        $checkboxidrac.Checked=$true
        $idracinput.Enabled=$false
        $idracinput2.enabled=$false
        $idracinput3.enabled=$false
        $idracinfo4.Text = "$idracip Ping Passed"
        $idracinfo4.ForeColor= "green"
               }

               if($idrac_ip -match "bypassed"){
                $checkboxidrac2.Checked=$true
                $idracinput.Enabled=$false
                $idracinput2.enabled=$false
                $idracinput3.enabled=$false
                $idracinfo4.Text = "Bypass iDRAC settings"
                $idracinfo4.ForeColor= "Gray"               
               }

    
#endregion

#region TC selection region

#region ## ini checkbox
    $checkboxini = New-Object System.Windows.Forms.CheckBox
    $checkboxini.Location = New-Object System.Drawing.Point(10,250)
    $checkboxini.text = "ini"
    $tab[$i-1].Controls.Add($checkboxini)
      if(!(test-path C:\testing_AI)){
        $checkboxini.Checked = $true
        $enabledList.Items.Add("ini")
      }
    # Get Checkbox Event
    $checkboxini.add_CheckedChanged({
        # check CheckBox selected
      
        if ($this.Checked) {
            # CheckBox text add to top of listbox
            $enabledList.Items.Insert(0, $this.Text)
        }else{
            $enabledList.Items.Remove($this.Text)
        }
    }) 

#endregion

#region### add button

    $addButton = New-Object system.Windows.Forms.Button
    $addButton.location = New-Object System.Drawing.Point(440,325) 
    $addButton.text     = "-->"
    $addButton.width    = 70
    $addButton.height   = 30
    $tab[$i-1].Controls.Add($addButton)

    $addButton.Add_Click(
        {
            @( $disabledList.SelectedItems ) | ForEach-Object {
                $enabledList.Items.Add($_)
                $disabledList.Items.Remove($_)
                
                #fin fixed position
                #$enabledList.Items.Remove("fin")
                #$enabledList.Items.Add("fin")
            }
        }.GetNewClosure()
    )

#endregion

#region ### remove button
    $removeButton = New-Object system.Windows.Forms.Button
    $removeButton.location = New-Object System.Drawing.Point(440,370)
    $removeButton.text     = "<--"
    $removeButton.width    = 70
    $removeButton.height   = 30
    $removeButton.Add_Click(
        {
            @( $enabledList.SelectedItems ) | ForEach-Object {
                if(($enabledList.SelectedItem -ne "fin") -and ($enabledList.SelectedItem -ne "ini")){
                    $disabledList.Items.Add($_)
                    $enabledList.Items.Remove($_)
                }               
            }
        }.GetNewClosure()
    )
    $tab[$i-1].Controls.Add($removeButton)

 #endregion
 
 
 #region ###movestepUp
    $movestepUp = New-Object system.Windows.Forms.Button
    $movestepUp.location = New-Object System.Drawing.Point(660,240)
    $movestepUp.text     = "↑"
    $movestepUp.width    = 50
    $movestepUp.height   = 30
    $movestepUp.Add_Click(
        {
            @( $enabledList.SelectedItems ) | ForEach-Object {
                if(($enabledList.Items[$enabledList.SelectedIndex - 1] -ne "ini") -and ($enabledList.SelectedItem -ne "ini")){
                    $itemToMove = $enabledList.SelectedItem                   

                    $newIndex = $enabledList.SelectedIndex - 1 
                    $enabledList.Items.RemoveAt($enabledList.SelectedIndex)
                    if ($newIndex -lt 0) {
                        $newIndex = 0  # 限制不超過最小索引值
                    }
                    $enabledList.Items.Insert($newIndex, $itemToMove)
                    $enabledList.SelectedIndex = $newIndex
                }               
            }
        }.GetNewClosure()
    )
    $tab[$i-1].Controls.Add($movestepUp)

 #endregion

#region ###movestepDown
    $movestepDown = New-Object system.Windows.Forms.Button
    $movestepDown.location = New-Object System.Drawing.Point(750,240)
    $movestepDown.text     = "↓"
    $movestepDown.width    = 50
    $movestepDown.height   = 30
    $movestepDown.Add_Click(
        {
            @( $enabledList.SelectedItems ) | ForEach-Object {
                if(($enabledList.SelectedIndex + 1) -lt $enabledList.Items.Count -and ($enabledList.SelectedItem -ne "ini")){
                #if(($enabledList.Items[$enabledList.SelectedIndex + 1] -ne "fin")){
                    $itemToMove = $enabledList.SelectedItem                   

                    $newIndex = $enabledList.SelectedIndex + 1 
                    $enabledList.Items.RemoveAt($enabledList.SelectedIndex)
                    #if ($enabledList.Items.Count -gt $newIndex) {
                    #    $newIndex = 0  # 限制不超過最小索引值
                    #}
                    $enabledList.Items.Insert($newIndex, $itemToMove)
                    $enabledList.SelectedIndex = $newIndex
                }               
            }
        }.GetNewClosure()
    )
    $tab[$i-1].Controls.Add($movestepDown)
#endregion

#region ###OkButton
    $OkButton          = New-Object system.Windows.Forms.Button
    $OkButton.location = New-Object System.Drawing.Point(440,415)
    $OkButton.text     = "Run"
    $OkButton.width    = 70
    $OkButton.height   = 30
    $OkButton.Font     = 'Microsoft Sans Serif,10'
    $OkButton.Enabled = $false
    $global:clickflag = 0
    $OkButton.Add_Click({
        if($enabledList.Items -ne $null){
            $global:clickflag = 1
            OkClick
        }else{
            [System.Windows.Forms.MessageBox]::Show("請選擇要執行的TC腳本", "提示", "OK", "Information")
        }
    })
    
    $tab[$i-1].Controls.Add($OkButton)

 #endregion   


#region #CleanButton
    $CleanButton          = New-Object system.Windows.Forms.Button
    $CleanButton.location = New-Object System.Drawing.Point(440,415)
    $CleanButton.text     = "Clean"
    $CleanButton.width    = 70
    $CleanButton.height   = 30
    $CleanButton.Font     = 'Microsoft Sans Serif,10'
    
    $CleanButton.Add_Click({
        $InputBox.Text = "請輸入搜尋關鍵字..."
        $InputBox.ForeColor = "gray"
        $disabledList.Items.Clear()
        $disabledList.Items.AddRange($TClist)
        $enabledList.Items.Clear()
        $checkboxini.Checked = $false
        $searchTemp.Text = ""
        #$enabledList.Items.Add("fin")
       
    })
    $tab[$i-1].Controls.Add($CleanButton)

   #endregion   
  
#endregion


#region Event handler for the $nameinput checkbox CheckedChanged event

 $nameinput.Add_TextChanged({

     $nameid=$nameinput.text
    # write-host "$nameid"
      $nameinfo2.Text = ""
  if($nameid.Length -eq 0 -or $nameid -match "required"){
    $nameinfo2.Text = "Must Input"
    $OkButton.Enabled = $false
  }
  else{
  $nameinfo2.Text = ""
     }

  $nameid=$nameinput.text
  
       if ($checkboxsys.Checked -and ($checkboxidrac.Checked -or $checkboxidrac2.Checked) -and $nameid.Length -ne 0  -and !($nameid -match "required") -and $checkboxsys2.Checked) {
        $OkButton.Enabled = $true
    } else {
     
        $OkButton.Enabled = $false
    }

})   

#endregion

  # Event handler for the $checkboxidrac checkbox CheckedChanged event

 
 $checkboxidrac2.add_CheckedChanged({
   if ($checkboxidrac2.Checked) {
        $checkboxidrac.Enabled=$false
        $idracinput.Enabled=$false
        $idracinput2.enabled=$false
        $idracinput3.enabled=$false
        $idracinfo4.Text = "Bypass iDRAC settings"
        $idracinfo4.ForeColor= "Gray"
   }
   else{
       $checkboxidrac.Enabled=$true
       $idracinput.Enabled=$true
       $idracinput2.enabled=$true
       $idracinput3.enabled=$true
        $idracinfo4.Text = ""
   
   }

          $nameid=$nameinput.text


     if ($checkboxsys.Checked -and ($checkboxidrac.Checked -or $checkboxidrac2.Checked) -and $nameid.Length -ne 0  -and !($nameid -match "required") -and $checkboxsys2.Checked) {
        $OkButton.Enabled = $true
    } else {
     
        $OkButton.Enabled = $false
    }

 })

 
# Create an empty variable to hold the ping result
$pingcheck = $null
 
  $checkboxidrac.add_CheckedChanged({
  if ($checkboxidrac.Checked) {
     # ping ip
      $idracip=$idracinput.Text
      if($idracip -eq "192.168.2."){
           $checkboxidrac.Checked = $false
            $idracinput.Enabled=$true
            $idracinput2.enabled=$true
            $idracinput3.enabled=$true
            $idracinfo4.Text = "$idracip Ping Failed"
            $idracinfo4.ForeColor= "red"
      }
      else{
        $pingresult = $ping.Send($idracip, 1000)
        $pingcheck = $pingresult.Status
        #Write-Host "IP is $idracip"
        #Write-Host "$pingcheck"
        if ($pingcheck -ne [System.Net.NetworkInformation.IPStatus]::Success) {
            $checkboxidrac.Checked = $false
            $idracinput.Enabled=$true
            $idracinput2.enabled=$true
            $idracinput3.enabled=$true
            $idracinfo4.Text = "$idracip Ping Failed"
            $idracinfo4.ForeColor= "red"

        }
        else{
        $idracinput.Enabled=$false
        $idracinput2.enabled=$false
        $idracinput3.enabled=$false
        $idracinfo4.Text = "$idracip Ping Passed"
        $idracinfo4.ForeColor= "green"
               }
  }

  }
  else{
            $checkboxidrac.Checked = $false
            $idracinput.Enabled=$true
            $idracinput2.enabled=$true
            $idracinput3.enabled=$true
            $idracinfo4.Text = ""
            
  }

       $nameid=$nameinput.text


     if ($checkboxsys.Checked -and ($checkboxidrac.Checked -or $checkboxidrac2.Checked) -and $nameid.Length -ne 0  -and !($nameid -match "required") -and $checkboxsys2.Checked) {
        $OkButton.Enabled = $true
    } else {
     
        $OkButton.Enabled = $false
    }
})  




#region Event handler for the $checkboxsys2  CheckedChanged event
$checkboxsys2.add_CheckedChanged({

 if ($checkboxsys2.Checked){
 $checkdvm=get-process *|?{$_.MainWindowTitle -match "Device Manager"}
 if( $checkdvm){
 stop-process -name mmc -ErrorAction SilentlyContinue
 }
 
}

  $nameid=$nameinput.text

     if ($checkboxsys.Checked -and ($checkboxidrac.Checked -or $checkboxidrac2.Checked ) -and $nameid.Length -ne 0  -and !($nameid -match "required") -and $checkboxsys2.Checked) {
        $OkButton.Enabled = $true
    } else {
     
        $OkButton.Enabled = $false
    }


})  

#endregion


#region Event handler for the $checkboxsys  CheckedChanged event

$checkboxsys.add_CheckedChanged({

  $nameid=$nameinput.text

 if ($checkboxsys.Checked -and ($checkboxidrac.Checked -or $checkboxidrac2.Checked ) -and $nameid.Length -ne 0  -and !($nameid -match "required") -and $checkboxsys2.Checked) {
        $OkButton.Enabled = $true
    } else {
       $OkButton.Enabled = $false
    }
    
})   

#endregion

function OkClick {
    $Form.Close()
}


$Form.Controls.Add($TabControl)
# Show the form

$Form.ShowDialog() | Out-Null
$Form.Show| Out-Null

if($global:clickflag -eq 1){

    $enabledList.Items.Add("fin") | Out-Null

    $con2=$con | Where-Object{$_.TC -in $enabledList.Items}
    #$con2 =$con2 | Where-Object { $_.TC -in $enabledList.Items -and $_.status -ne "" } | Sort-Object { $enabledList.Items.IndexOf($_.TC) },{[int]$_.Step_No} | Select-Object -Property TC,Step_No,programs,para1,para2,para3,para4,para5
    $con2 =$con2 | Where-Object { $_.TC -in $enabledList.Items} | Sort-Object { $enabledList.Items.IndexOf($_.TC) },{[int]$_.Step_No} | Select-Object -Property TC,TC_step,Step_No,programs,must,para1,para2,para3,para4,para5

    
    if($con2){
        #Set-Content -Path "C:\testing_AI\settings\flowsettings.csv" -Value $con2      
        
        if(Test-Path C:\testing_AI\){
            $path3 = "C:\testing_AI\settings\flowsettings.csv"
            $con2| export-csv $path3 -Encoding UTF8 -NoTypeInformation
            #C:\testing_AI\AutoRun*
        }else{
            $path3 = "C:\temp_aitest\flowsettings.csv"
            $con2| export-csv $path3 -Encoding UTF8 -NoTypeInformation
        }  
       if(Test-Path $PSScriptRoot\flowsettings.csv){
       Copy-Item $PSScriptRoot\flowsettings.csv -Destination  $path3 -Force
       }
    }


#region write tester/idrac info
 $nameid=$nameinput.text
 $idracip=$idracinput.text
 $idracusern=$idracinput2.text
 $idracpasswd=$idracinput3.text

     if(test-path $idracpath2){
      if(!($checkboxidrac2.Checked)){
       set-content $idracpath2 -value "$idracip,$idracusern,$idracpasswd"
       }
       set-content $testername_path2 -value $nameid

     }
     else{
      if(!($checkboxidrac2.Checked)){
      set-content $idracpath1 -value "$idracip,$idracusern,$idracpasswd"
      }
      set-content $testername_path -value $nameid
     }
     }
#endregion
##   BUILD BUDDY         ##
##   6/26/2020           ##
##   peter cleaves       ##
###########################
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Description."
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Description."
$cancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Description."
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
$installpath = "\\gvnyfs01v\Shared\pcleaves\script"
$installpath2 = "\\gvnyfs01v\Shared\pcleaves\optional"
$loggedinuser = (get-wmiobject -class win32_process -filter 'Name="explorer.exe"').getowner().user
Set-Location C:
Write-Host "`n"
Write-Host "`n"

##############################
##############################
#STEP 1 - UPDATES
#Prompt user choice to run check for and run any reboots, reboots if needed
$title = Write-Host "STEP1 - UPDATES" -ForegroundColor Green 
$message = Write-Host "Do you want to check for updates and install them?" -ForegroundColor Green
$result = $host.ui.PromptForChoice($title, $message, $options, 1)
switch ($result) {
  0{
    Write-Host "Yes - SEARCHING..." -ForegroundColor Green
    Write-Host "`n"
    install-module pswindowsupdate
    Get-WUInstall -AcceptAll -AutoReboot -Install
  }1{
    Write-Host "No - SKIPPING" -ForegroundColor Gray
    continue
  }2{
  Write-Host "Cancel"
  exit
  }
}
Write-Host "`n"
Write-Host "`n"

##############################
##############################
#STEP 2 -CORE SOFTWARE
Write-Host "STEP 2 - CHECKING FOR CORE SOFTWARE (PROMPT FOR INSTALLS)" -ForegroundColor Green
#Search for all software categorized as MSI or Programs and filter results down to just name
start-sleep -s 1
$SEARCH = Get-Package -providername msi,programs | select name | sort name
#baseline is a list of the software that SHOULD be on the system
$baseline = get-content $installpath\baseline.txt
#Check if each of the progams in the baseline file match the results of the Search. 
#If any are not installed, prompt the user if they want to install passes to a batch file of the same name to run the installer.
foreach ($program in $baseline)
{
    if($Search -imatch $program)
    {
    Write-Host "$Program is installed" -ForegroundColor Green
    }
    else
    {
    Write-Host "$Program IS NOT INSTALLED - DO YOU WANT TO INSTALL NOW? " -ForegroundColor Magenta
    $result = $host.ui.PromptForChoice($title, $message, $options, 1)
    switch ($result) {
      0{
        Write-Host "Yes - SEARCHING..." -ForegroundColor Green
        Write-Host "INSTALLING.." -ForegroundColor Green
        & "$installpath\$program.bat" -verbose
        Write-Host "DONE" -ForegroundColor Green
      }1{
        Write-Host "No - SKIPPING"-ForegroundColor Yellow
        continue
      }2{
      Write-Host "Cancel"
      exit
      }
    }
    }
 }

 ##Have to handle Teams differently because its there is no package in get-package for some reason, I think because its an extension of Office
 $errorActionPreference = 'silentlycontinue'
 $Teamsinstalled = Gci "C:\users\$loggedinuser\appdata\roaming\microsoft\teams"
 $errorActionPreference = 'continue'
 if ($Teamsinstalled -eq $null){
 Write-Host "Installing Teams..." -ForegroundColor Green
 &$installpath\teams.bat
 }
 else{
 Write-Host "Teams is installed" -ForegroundColor Green
 }
 
 start-sleep -s 1

##############################
##############################
 #STEP 3 - INFORMATIONAL SOFTWARE
 ###Non-core but frequent software, not installing - just checking
Write-Host "`n"
Write-Host "`n"
Write-Host "STEP 3 - CHECKING FOR OPTIONAL SOFTWARE VIEW AT A GLANCE (NO INSTALLS)" -ForegroundColor Green
 start-sleep -s 1
 $informational = Get-Content $installpath\informational.txt | sort
 foreach ($Program in $informational){
    if($Search -imatch $Program)
    {
    Write-Host "$Program is installed" -ForegroundColor Yellow
    }
    else
    {
    Write-Host "$Program is NOT installed" -ForegroundColor Gray
    }
    }
  
Write-Host "`n"
Write-Host "`n"

##############################
##############################
# STEP 4 - Hardware info and System Updates (HP/Lenovo) 
#Map logged in user registry (works even when PS is launched as administrator. 
$user = New-Object System.Security.Principal.NTAccount($loggedinuser)
$sid = $user.Translate([System.Security.Principal.SecurityIdentifier]).Value
New-PSDRive HKU Registry HKEY_USERS | out-null
#Detect System Type
start-sleep -s 1
Write-Host "STEP 4 - CHECKING HARDWARE INFO..." -ForegroundColor Green
start-sleep -s 1
$manufacturer = (Get-CimInstance -classname Win32_ComputerSystem).Manufacturer
Write-Host "Hardware Manufacturer is Detected as $manufacturer" -ForegroundColor Green
### HP DETECTED
if ($manufacturer -eq 'HP' -and (Get-Package "*hp support assistant").status -eq "Installed"){
Write-Host "HP System Update Already Installed" -ForegroundColor Green} 
if ($manufacturer -eq 'HP' -and (Get-Package "*hp support assistant") -eq $null) {
Write-Host "HP System Update Not Detected... Do you want to Install HP SYstem Update?" -ForegroundColor Magenta
$result = $host.ui.PromptForChoice($title, $message, $options, 1)
    switch ($result) {
      0{
        Write-Host "Yes - SEARCHING..." -ForegroundColor Green
        Write-Host "INSTALLING.." -ForegroundColor Green
        & "$installpath2\$manufacturer.bat"
        Write-Host "DONE" -ForegroundColor Green
      }1{
        Write-Host "No - SKIPPING" -ForegroundColor Yellow
        continue
      }2{
      Write-Host "Cancel"
      exit
      }
    }
}
#Change IE homepage to GCM homepage for HP PCs
if ($manufacturer -eq "HP"){
$browserpath = "HKU:\$sid\SOFTWARE\MICROSOFT\INTERNET EXPLORER\MAIN\"
$browsername = 'start page'
$browservalue = "http://gcm/"
Write-Host "IE home page set to $browservalue ..." -ForegroundColor Green
Set-ItemProperty -Path $browserpath -Name $browsername -Value $browservalue
}

##Get Hardware information
Write-Host (Get-wmiobject -Class win32_Processor | select -first 1).Name -ForegroundColor Green
Write-Host ((Get-CimInstance Win32_PhysicalMemory | measure -Property Capacity -Sum).Sum/1gb)  "GB of RAM Installed" -ForegroundColor Green
Write-host ("{0:N2}" -f ((Get-Volume | where{$_.DriveLetter -eq "C"}).SizeRemaining/1gb))"/"("{0:N2}" -f ((Get-Volume | where{$_.DriveLetter -eq "C"}).Size/1gb)) "GB Free Disk space" -ForegroundColor green 
Write-Host "`n"
##############################
##############################
###STEP 5 checking and setting minor settings
Write-Host STEP 5 - MINOR CONFIG CHANGES... -ForegroundColor Green
start-sleep -s 1



#Copy Desktop Shorcuts
Write-Host "Copying desktop shortcuts..." -ForegroundColor Green
Copy-Item "\\gvnyit01v.corp.glenviewcapital.com\install\win10\desktop shortcuts\*" C:\Users\$loggedinuser\Desktop\
#Bitlocker check and enable
$errorActionPreference = 'silentlycontinue'
$tpm = (get-tpm).TpmPresent
$errorActionPreference = 'continue'
$bitlockerstatus = (Get-BitLockerVolume C:).VolumeStatus
if ($bitlockerstatus -eq "FullyDecrypted" -and $tpm -eq "True"){
Write-host "Not Encrypted...Encrypting" -ForegroundColor Magenta
Enable-BitLocker -MountPoint C: -UsedSpaceOnly -SkipHardwareTest -RecoveryPasswordProtector
}
else{
Write-Host "Bitlocker already configured (Or VDI)" -ForegroundColor Green
}
#Add logged in user to remote dekstop group
$rdusers = (Get-LocalGroupMember -group "Remote Desktop Users")
if ($rdusers -eq $null){
Add-LocalGroupMember -Group "Remote Desktop Users"-Member "$loggedinuser"
Write-Host "Adding $loggedinuser to Remote Desktop Users Group... " -ForegroundColor Magenta
Get-LocalGroupMember -Group "Remote Desktop Users"
}
else{
Write-Host "Remote Desktop Users Already Configured" -ForegroundColor Green
}
#google default IE search engine
cd HKU:\
if ((Get-ItemProperty -Path "HKU:\$Sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\" -Name "DefaultScope").defaultscope -eq "{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}"){
Write-host "Google search already set to default IE search provider" -ForegroundColor Green
}
else{
Write-Host "Google search not IE default - FIXING.." -ForegroundColor Magenta
Start-Sleep -s 2
cd "HKLM:\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\"
Remove-Item "HKLM:\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\*"
$searchpath = "HKLM:\SOFTWARE\Microsoft\Internet Explorer\SearchScopes"
$searchname = 'DefaultScope'
$searchvalue = "{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}"
Set-ItemProperty -Path $searchpath -Name $searchname -Value $searchvalue
New-Item -Path $searchpath\$searchvalue
New-ItemProperty -Path $searchpath\$searchvalue -Name 'DisplayName' -Value 'Google'
#New-ItemProperty -Path $searchpath\$searchvalue -Name 'OSDFileURL' -Value 'https://www.microsoft.com/cms/api/am/binary/RWilsM'
#New-ItemProperty -Path $searchpath\$searchvalue -Name 'FaviconPath' -Value "%USERPROFILE%\\AppData\\LocalLow\\Microsoft\\Internet Explorer\\Services\\Search_{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}.ico"
New-ItemProperty -Path $searchpath\$searchvalue -Name 'FaviconURL' -Value 'https://www.google.com/favicon.ico'
New-ItemProperty -Path $searchpath\$searchvalue -Name 'ShowSearchSuggestions' -PropertyType DWord -Value '1'
New-ItemProperty -Path $searchpath\$searchvalue -Name 'URL' -Value "https://www.google.com/search?q={searchTerms}&sourceid=ie7&rls=com.microsoft:{language}:{referrer:source}&ie={inputEncoding?}&oe={outputEncoding?}"
New-ItemProperty -Path $searchpath\$searchvalue -Name 'SuggestionsURL' -Value "http://clients5.google.com/complete/search?q={searchTerms}&client=ie8&mw={ie:maxWidth}&sh={ie:sectionHeight}&rh={ie:rowHeight}&inputencoding={inputEncoding}&outputencoding={outputEncoding}"
cd "HKU:\$sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\"
Remove-Item "HKU:\$sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\*"
Set-Location hku:
Set-ItemProperty -Path "HKU:\$sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes" -Name 'DefaultScope' -Value $searchvalue
New-Item -Path "HKU:\$sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}"
Set-Location "HKU:\$sid\SOFTWARE\Microsoft\Internet Explorer\SearchScopes\{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}"
New-ItemProperty -Path .\ -Name 'DisplayName' -Value 'Google'
#New-ItemProperty -Path .\ -Name 'OSDFileURL' -Value 'https://www.microsoft.com/cms/api/am/binary/RWilsM'
#New-ItemProperty -Path .\ -Name 'FaviconPath' -Value "%USERPROFILE%\\AppData\\LocalLow\\Microsoft\\Internet Explorer\\Services\\Search_{78993B0B-5FA4-4B65-9D5A-FAA4CBA2666F}.ico"
New-ItemProperty -Path .\ -Name 'FaviconURL' -Value 'https://www.google.com/favicon.ico'
New-ItemProperty -Path .\ -Name 'ShowSearchSuggestions' -PropertyType DWord -Value '1'
New-ItemProperty .\ -Name 'URL' -Value "https://www.google.com/search?q={searchTerms}&sourceid=ie7&rls=com.microsoft:{language}:{referrer:source}&ie={inputEncoding?}&oe={outputEncoding?}"
New-ItemProperty .\ -Name 'SuggestionsURL' -Value "http://clients5.google.com/complete/search?q={searchTerms}&client=ie8&mw={ie:maxWidth}&sh={ie:sectionHeight}&rh={ie:rowHeight}&inputencoding={inputEncoding}&outputencoding={outputEncoding}"
}
#local admin check
$localadmins = Get-LocalGroupMember -Group "Administrators" | where{$_.name -like "*$loggedinuser*"}
if ($localadmins -eq $null){
Write-Host "Current Logged in User is NOT a local administrator" -ForegroundColor Green
}
else{
Write-Host "CURRENT LOGGED IN USER IS A LOCAL ADMIN" -ForegroundColor Magenta
Write-Host "Do you want to remove this user as a Local Admin" -ForegroundColor Magenta
Get-LocalGroupMember -Group "Administrators" | Out-Host
$result = $host.ui.PromptForChoice($title, $message, $options, 1)
switch ($result) {
  0{
    Write-Host "YES.. REMOVING ADMIN RIGHTS From $loggedinuser" -ForegroundColor Green
    Remove-LocalGroupMember -Group "Administrators"-Member "$loggedinuser"
  }1{
    Write-Host "No - SKIPPING" -ForegroundColor Gray
    continue
  }2{
  Write-Host "Cancel"
  exit
  }
}
}
#bloomberg excel add-in
$errorActionPreference = 'silentlycontinue'
$BloombergInstalled = (get-package *Bloomberg*)
$errorActionPreference = 'continue'
if ($BloombergInstalled -ne $null){
$AddinPath = "C:\blp\API\Office Tools\BloombergUI.xla"
try {
    $Excel = New-Object -ComObject Excel.Application
} catch {
    Write-Error 'Microsoft Excel does not appear to be installed.'
}
try {
    $ExcelAddins = $Excel.Addins
    # The Add() method of the AddIns interface will fail if we don't have a workbook!
    $ExcelWorkbook = $Excel.Workbooks.Add()
    $Addin = Get-ChildItem -Path ($AddinPath)
    $AddinInstalled = ($ExcelAddins | ? { $_.Name -eq $Addin.Name}).installed
    if ($AddinInstalled -ne "True") 
    {
        
        $Addin.IsReadOnly = $true
        $NewAddin = $ExcelAddins.Add($Addin.FullName, $false)
        $NewAddin.Installed = $true
        Write-Host ('Add-in "' + $Addin.BaseName + '" successfully installed!') -ForegroundColor Green
    } else {
        Write-Host ('Add-in "' + $Addin.BaseName + '" already installed!') -ForegroundColor Green
           }
    } finally {
    $Excel.Quit()
              }
}
else{
Write-Host "Bloomberg not installed , skipping Excel Add-in" -ForegroundColor Green
}
#PDF DEFAULT (All) --SKIP--
#homepage to google (lenovo)
#Bomgar (lenovo)
#Power Settings (lenovo)
#LENOVO TVSU (Lenovo Only)
#Cisco VPN (Lenovo Only)

Function PCEnable-Remoting{
    param ( 
        $computername = 'localhost'
       
    )
    
    Invoke-WmiMethod -Class win32_process -ComputerName $computername -name create -argument 'winrm.cmd quickconfig -quiet'
    Write-Host "TARGETING $computername" -ForegroundColor Green
    if ($computername -ne 'localhost')
    {Enter-PSSession $computername}
    
    }
    

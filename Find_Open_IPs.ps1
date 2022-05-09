#Look up IP range for available IPs
#Peter Cleaves
$Network = "10.32.1."
$RunTime = Get-Date
$PingTest = ForEach ($Oct in (1..255))
{   $IP = "$Network$Oct"
    $Ping = Get-CimInstance Win32_PingStatus -Filter "Address = '$IP' AND ResolveAddressNames = TRUE"
    [PSCustomObject]@{
        Responded = [bool](-not $ping.StatusCode)
        ReverseDNS = $Ping.ProtocolAddressResolved
        RunTime = $RunTime
        IP = $Oct
    }
}
$PingTest | Out-GridView -Title "Results of PingTest at $RunTime"

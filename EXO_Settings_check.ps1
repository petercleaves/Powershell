# glenview capital management
# powershell script to check status of mailrules.ps1
# author: peter cleaves
# date: august 3, 2021
#add connect-exchangeonline cmdlet into powershell session
Import-Module ExchangeOnlineManagement
#using password
Connect-ExchangeOnline -CertificateFilePath "C:\Scripts\MailRulesCheck\MyCertificate.pfx" -CertificatePassword (ConvertTo-SecureString -String '#######' -AsPlainText -Force) -AppID “#######-#####-#####-#######-########” -Organization “########.OnMicrosoft.com”
Get-CASMailbox | select -first 1 | fl *
$EWSCheck = Get-CASMailbox | Where-Object {$_.EWSenabled -ne "True"} 
$OWACheck = Get-CASMailbox | Where-Object {$_.OWAenabled -eq "True"}
#$POPCheck = Get-CASMailbox | Where-Object {$_.POPEnabled -eq "True"}
#$IMAPCheck = Get-CASMailbox | Where-Object {$_.IMAPEnabled -eq "True"}
$MAPICheck = Get-CASMailbox | Where-Object {$_.MAPIEnabled -ne "True"}
#$ActiveSyncCheck = Get-CASMailbox | Where-Object {$_.ActiveSyncEnabled -eq "True"}
$mailMessage = New-Object system.net.mail.mailmessage
$computer = Get-WmiObject win32_computersystem
$smtpclient = new-object system.Net.Mail.SmtpClient
$smtpclient.host = "gvnysmtp"
$mailmessage.from = "ExchangeOnlineSettings@glenviewcapital.com"
$mailmessage.to.add("IT@glenviewcapital.com")
$mailmessage.subject = "Exchange Online insecure access settings detected "
$mailmessage.replyto = "pcleaves@glenviewcapital.com"

if ($OWACheck -ne $null)
{
$Names = $OWACheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "OWA Enabled for the following users:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }


}

<#
if ($PopCheck -ne $null)
{
$Names = $PopCheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "`n`n`n POP Enabled for the following users:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }


}
#>
<#if ($IMAPCheck -ne $null)
{
$Names = $Imapcheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "`n`n`n IMAP Enabled for the following users:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }


}
#>
if ($MAPICheck -ne $null)
{
$Names = $MAPICheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "`n`n`n MAPI DISABLED for the following users:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }


}
<#
if ($ActiveSyncCheck -ne $null)
{
$Names = $ActiveSyncCheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "`n`n`n ActiveSync Enabled for the following users:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }


}
#>
if ($EWScheck -ne $null)
{
$Names = $EWScheck | sort name | Select-Object -ExpandProperty name
$mailmessage.body += "`n`n`n EWS DISABLED for the following users (needed for Free/Busy calendar):`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n`n"
    }


}

$groupo365 = Get-ADGroupMember -Identity grp_o365 | Select-Object -ExpandProperty name | sort
$o365mailboxes = Get-Mailbox | Select-Object -ExpandProperty name | sort

$o365check = Compare-Object ($groupo365) ($o365mailboxes) | Select-Object -ExpandProperty inputobject | sort

if ($o365check -ne $Null)
{
$mailMessage.Body += "grp_o365 and Exchange Online MISMATCH:`n`n"
    foreach ($check in $o365check)
    {
    $mailmessage.Body += $check + "`n"
    }
}

$names = Get-EXOMailbox -Properties Name, RetentionPolicy | Where-Object {$_.RetentionPolicy -notlike "*Glenview*"} | Select-Object -ExpandProperty DisplayName | sort
$mailmessage.body += "`n`n`n Glenview 90 Days Email retention policy NOT set:`n`n" 
    foreach ($name in $names)
    {
    $mailmessage.body += $Name + "`n"
    }




$mailmessage.body += "`n`n*****This process runs from " + $computer.name + " and is located at " + $MyInvocation.MyCommand.Definition + "*****"
$smtpclient.Send($mailmessage)

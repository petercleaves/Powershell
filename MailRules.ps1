# glenview capital management
# powershell script to enable and disable mail rules in exchange online
# author: peter cleaves
# date: july 22, 2021
#add connect-exchangeonline cmdlet into powershell session
Import-Module ExchangeOnlineManagement

#using password (automatic password connection removed for public version)
Connect-ExchangeOnline 

#Flip mail rules enabled <> disabled
$targets = Get-TransportRule | Where-Object {$_.Identity -like "*On-Call*" -or $_.Identity -like "*pagerduty*"} 
foreach ($target in $targets) 
{
      if($target.State -eq 'Disabled') {$target | Enable-TransportRule}
        else {$target | Disable-TransportRule -Confirm:$false}    
}
$transportstatus = Get-TransportRule | Where-Object {$_.Identity -like "*On-Call*" -or $_.Identity -like "*pagerduty*" -or $_.Identity -like "*wake*"} 


#flip out of office bounceback enabled <> disabled
$OOOstatus = Get-MailboxAutoReplyConfiguration oncall
if ($OOOstatus.AutoReplyState -eq 'Disabled')
{Set-MailboxAutoReplyConfiguration -Identity oncall -AutoReplyState enabled -ExternalAudience all}
else {Set-MailboxAutoReplyConfiguration -Identity oncall -AutoReplyState disabled}
$OOOstatus = Get-MailboxAutoReplyConfiguration oncall
#send the email
$mailMessage = New-Object system.net.mail.mailmessage
$computer = Get-WmiObject win32_computersystem
$smtpclient = new-object system.Net.Mail.SmtpClient
$smtpclient.host = "gvnysmtp"
$mailmessage.from = "MailRules@glenviewcapital.com"
$mailmessage.to.add("pcleaves@glenviewcapital.com")
$mailmessage.subject = "On Call Autoreply " + ($transportstatus.State | select -first 1)
$mailmessage.replyto = "it@glenviewcapital.com"

$mailmessage.body += "`n`n Oncall@glenview OOO Bounceback: " + $OOOstatus.AutoReplyState
$mailmessage.body += "`n`n Oncall Transport Rule: " + ($transportstatus | Where-Object {$_.Identity -like "*On-Call*"}).state
$mailmessage.body += "`n`n IT rotation Pagerduty Transport Rule: " + ($transportstatus | Where-Object {$_.Identity -like "*pagerduty*"}).state
$mailmessage.body += "`n`n 'Wake' for traders Pagerduty Transport Rule: " + ($transportstatus | Where-Object {$_.Identity -like "*wake*"}).state



$mailmessage.body += "`n`n*****This process runs from " + $computer.name + " and is located at " + $MyInvocation.MyCommand.Definition + "*****"
$smtpclient.Send($mailmessage)

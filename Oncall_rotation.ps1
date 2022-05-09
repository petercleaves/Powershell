# glenview capital management
# powershell script to change oncall rotation
# author: peter cleaves
# date: july 27, 2021
#add connect-exchangeonline cmdlet into powershell session
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline
#What's your username? 
$peter = New-Object System.Management.Automation.Host.ChoiceDescription "&Peter","Description."
$Employee = New-Object System.Management.Automation.Host.ChoiceDescription "&Employee","Description."
$Employee= New-Object System.Management.Automation.Host.ChoiceDescription "&Employee","Description."
#$Employee= New-Object System.Management.Automation.Host.ChoiceDescription "&Employee","Description."
$cancel= New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel","Description."
#configure the OOO replies
$Employee_html = @'
<html>
                                   <body>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt">For<span>&nbsp;</span><b>time sensitive<span>&nbsp;</span></b>issues or requests 
                                   outside of our regular coverage hours (7:30am - 7pm Mon - Fri), please call the following person who is on-call this 
                                   week:</span>
                                   <div style="margin:0px; font-size:12pt"><br>
                                   </div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Employee</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Cell: 123-456-7890</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)"><br>
                                   </span></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451">We aim to quickly respond to every request and we will respond 
                                   to this request whether you call us or not. However, outside of regular coverage hours, we want to make sure you have our 
                                   number in
                                    case of a time sensitive issue. The rest of our contact information is below. Thanks a lot.</font></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451"><br>
                                   </font></div>
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">-Glenview IT</span></span><br>
                                   </div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><br>
                                   </span></span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><span 
                                   style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">Employee</span></span>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 123-456-7890</font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451"><br>
                                   </font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Peter Cleaves</font></div>
                                   <span style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 703-943-7989</font></span><br>
                                   </span></span></div>
                                   </body>
                                   </html>
'@

$Employee_html = @'
<html>
                                   <body>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt">For<span>&nbsp;</span><b>time sensitive<span>&nbsp;</span></b>issues or requests 
                                   outside of our regular coverage hours (7:30am - 7pm Mon - Fri), please call the following person who is on-call this 
                                   week:</span>
                                   <div style="margin:0px; font-size:12pt"><br>
                                   </div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Employee</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Cell: 123-456-7890</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)"><br>
                                   </span></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451">We aim to quickly respond to every request and we will respond 
                                   to this request whether you call us or not. However, outside of regular coverage hours, we want to make sure you have our 
                                   number in
                                    case of a time sensitive issue. The rest of our contact information is below. Thanks a lot.</font></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451"><br>
                                   </font></div>
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">-Glenview IT</span></span><br>
                                   </div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><br>
                                   </span></span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><span 
                                   style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">Employee
                                    Employee</span></span>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 123-456-7890</font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451"><br>
                                   </font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Peter Cleaves</font></div>
                                   <span style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 123-456-7890</font></span><br>
                                   </span></span></div>
                                   </body>
                                   </html>
'@
<# $Employee_html = @'
<html>
                                   <body>
                                   <div></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,0)">This is a test of the new on-call rule</span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,0)"><br>
                                   </span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255)">Please Call the following person who is on-call this week:</span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255)"><br>
                                   </span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(81,167,249); font-size:18pt">Employee</span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(81,167,249); font-size:18pt">Cell: 123-456-7890</span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(81,167,249); font-size:18pt"><br>
                                   </span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(0,36,81); font-size:12pt">We will reply to your email, please call if an emergency</span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(0,36,81); font-size:12pt"><br>
                                   </span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="background-color:rgb(255,255,255); color:rgb(0,36,81); font-size:12pt">-IT</span></div>
                                   </body>
                                   </html>
'@
#>
$peter_html = @'
<html>
                                   <body>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt">For<span>&nbsp;</span><b>time sensitive<span>&nbsp;</span></b>issues or requests 
                                   outside of our regular coverage hours (7:30am - 7pm Mon - Fri), please call the following person who is on-call this 
                                   week:</span>
                                   <div style="margin:0px; font-size:12pt"><br>
                                   </div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Peter Cleaves</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)">Cell: 703-943-7989</span></div>
                                   <div style="margin:0px; font-size:12pt"><span style="margin:0px; font-size:18pt; color:rgb(81,167,249); 
                                   background-color:rgb(255,255,255)"><br>
                                   </span></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451">We aim to quickly respond to every request and we will respond 
                                   to this request whether you call us or not. However, outside of regular coverage hours, we want to make sure you have our 
                                   number in
                                    case of a time sensitive issue. The rest of our contact information is below. Thanks a lot.</font></div>
                                   <div style="margin:0px; font-size:12pt"><font color="#002451"><br>
                                   </font></div>
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">-Glenview IT</span></span><br>
                                   </div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><br>
                                   </span></span></div>
                                   <div style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)">
                                   <span style="margin:0px; font-size:12pt"><span style="margin:0px; color:rgb(0,36,81); background-color:rgb(255,255,255)"><span 
                                   style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><span style="margin:0px; color:rgb(0,36,81); 
                                   background-color:rgb(255,255,255)">Employee
                                    Employee</span></span>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 123-456-7890</font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451"><br>
                                   </font></div>
                                   <div style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Employee</font></div>
                                   <span style="margin:0px; font-size:12pt; color:rgb(0,0,0)"><font color="#002451">Cell: 123-456-7890</font></span><br>
                                   </span></span></div>
                                   </body>
                                   </html>
'@




$options = [System.Management.Automation.Host.ChoiceDescription[]]($Peter, $Employee, $Employee, $cancel)
$result = $host.ui.PromptForChoice($title, $message, $options, 3)

switch ($result) {
    0{
      Write-Host "You chose Peter's on call" -ForegroundColor Green
      Get-TransportRule | Where-Object {$_.Identity -like "*pagerduty*"}  | Set-TransportRule -BlindCopyTo "peter@gvhelp.pagerduty.com"
      Get-MailboxAutoReplyConfiguration -Identity oncall | Set-MailboxAutoReplyConfiguration -InternalMessage $peter_html -ExternalMessage $peter_html
      $username = "Peter"
    }1{
      Write-Host "You chose Employee's on call" -ForegroundColor Green
      Get-TransportRule | Where-Object {$_.Identity -like "*pagerduty*"}  | Set-TransportRule -BlindCopyTo "employee@gvhelp.pagerduty.com"
      Get-MailboxAutoReplyConfiguration -Identity oncall | Set-MailboxAutoReplyConfiguration -InternalMessage $Employee_html -ExternalMessage $Employee_html
      $username = "Employee"
    }2{
      Write-Host "You chose Employee's on call" -ForegroundColor Green
      Get-TransportRule | Where-Object {$_.Identity -like "*pagerduty*"}  | Set-TransportRule -BlindCopyTo "employee@gvhelp.pagerduty.com"
      Get-MailboxAutoReplyConfiguration -Identity oncall | Set-MailboxAutoReplyConfiguration -InternalMessage $Employee_html -ExternalMessage $Employee_html
      $username = "Employee"
<#     }3{
      Write-Host "You chose Employee's on call" -ForegroundColor Green
      Get-TransportRule | Where-Object {$_.Identity -like "*pagerduty*"}  | Set-TransportRule -BlindCopyTo "employee@gvhelp.pagerduty.com"
      Get-MailboxAutoReplyConfiguration -Identity oncall | Set-MailboxAutoReplyConfiguration -InternalMessage $Employee_html 
      $username = "Employee"
#>
    }3{
      Write-Host "Cancel"
      exit
    }
  }
 #send the email
$mailMessage = New-Object system.net.mail.mailmessage
$computer = Get-WmiObject win32_computersystem
$smtpclient = new-object system.Net.Mail.SmtpClient
$smtpclient.host = "gvnysmtp"
$mailmessage.from = "OncallRotationScript@glenviewcapital.com"
$mailmessage.to.add("it@glenviewcapital.com")
$mailmessage.subject = "Oncall Rotation script changed to " + $username
$mailmessage.replyto = "it@glenviewcapital.com"
$mailmessage.body += "Oncall rotation script ran by " + $username +" , pagerduty email and OOO reply have been updated"




$mailmessage.body += "`n`n*****This process runs from " + $computer.name + " and is located at " + $MyInvocation.MyCommand.Definition + "*****"
$smtpclient.Send($mailmessage)

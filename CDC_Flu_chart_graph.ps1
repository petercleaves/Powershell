# Glenview Capital Management
# Powershell script to check CDC website for updated data
# Author: Peter Cleaves
# Date: February 18, 2020
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$week = ((get-date -UFormat %V) - 1).ToString().PadLeft(2,'0')
$lastweek = (((get-date -UFormat %V) - 2).ToString().PadLeft(2,'0'))
$yearweek = (get-date -UFormat %Y) + $week
if ((test-path($yearweek + "table")) -and (test-path($yearweek + "image"))) { exit }
$image = "ILI" + $lastweek + "_small.gif"
#https://www.cdc.gov/flu/weekly/weeklyarchives2020-2021/images/ILI08_small.gif
$imageURL = "https://www.cdc.gov/flu/weekly/weeklyarchives2020-2021/images/" + $image 
$table = "senAllregt" + $lastweek + ".html"
$tableURL = "https://www.cdc.gov/flu/weekly/weeklyarchives2020-2021/data/" + $table
$dir = (Split-Path -Parent $MyInvocation.MyCommand.Definition) + "\"
$bFlag = $false
$smtp = "gvnysmtp"
$from = "CDC_Flu_Data@glenviewcapital.com"
#$to = @("")
$to = @("")
#$to = @("pcleaves@glenviewcapital.com")
$cc = "@glenviewcapital.com"
$subject = " Flu Week " + ((get-date -UFormat %V) - 1) + "-" + (get-date -UFormat %Y)
$body = "<html><body style=`"font-family:calibri`">" 
    $mailparams = @{to = $to; from = $from; cc = $cc; smtpserver = $smtp; bodyashtml = $true}
if (-NOT (test-path($yearweek + "image"))) {
    try { Invoke-WebRequest -Uri $imageURL -outfile  ($dir + $image) }
    catch { exit }
    if (Test-Path ($dir + $image)) { 
        $attachment = ($dir + $image)
        $mailparams += @{attachments = $attachment}
        $body += "<img src='cid:" + $image + "'><br>"
        $bFlag = $true
        new-item ($yearweek + "image")
    }
}
if (-NOT (test-path($yearweek + "table"))) {
    try { 
        $data = Invoke-WebRequest -Uri $tableURL
        $newtablename = "FluTable"
        $newtable = New-Object system.Data.DataTable "$newtablename"
        $col1 = New-Object system.Data.DataColumn "Week",([string])
        $col2 = New-Object system.Data.DataColumn "% Weighted ILI",([string])
        $newtable.columns.add($col1)
        $newtable.columns.add($col2)
        $tabledata = $data.ParsedHtml.getElementsByTagName("table")[0]
        $rows = @($tabledata.rows)
        $rowcount = @($tabledata.rows).count
        for ($i = 1; $i -le $rowcount; $i++) {
            $newrow = $newtable.NewRow()
            $cells = @($rows[$i].cells)
            $newrow."Week" = ($cells[0].InnerText)
            $newrow."% Weighted ILI" = $cells[9].InnerText
            $newtable.rows.add($newrow)
            if ($i -eq ($rowcount - 2)) { $previousILI = $cells[9].InnerText }
            if ($i -eq ($rowcount - 1)) { $currentILI = $cells[9].InnerText }
        }
        $body = "ILI % is at <b>" + $currentILI + "</b> in week " + ((get-date -UFormat %V) - 1) + " vs <b>" + $previousILI + "</b> last week.<br><br>" + $body
        $subject += ": ILI $currentILI`%"
        $body += $newtable | convertto-html -property "Week", "% Weighted ILI" -head "<style>th,td {text-align:center;}</style>"
        $bFlag = $true
        new-item ($yearweek + "table")
    }
    catch { 
        remove-item ($yearweek + "image")
        exit
    }
}
$body += "<br>*****This process runs from " + $computer.name + " and is located at " + $MyInvocation.MyCommand.Definition + "*****"
if ($bFlag) {
    $mailparams += @{subject = $subject}
    $mailparams += @{body = $body}
    Send-MailMessage @mailparams
    if (test-path($dir + $image)) { remove-item $image }
}

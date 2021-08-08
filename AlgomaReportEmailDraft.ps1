$date = (Get-Date).ToShortDateString()
$timeNow = (Get-Date).TimeOfDay
$noon = (Get-Date -Hour 12 -Minute 0 -Second 0).TimeOfDay

if ($timeNow -lt $noon){
    $greeting = "morning"
} else{
    $greeting = "afternoon"
}

$sReceipientAddr = @("jeff.greco@algoma.com", "joshua.jansen@algoma.com", "kimbal.beckett@algoma.com")
$sCCAddr = "nascofreight@qsl.com"
$sMsgSubject = "Algoma Daily Report"
$sMsgBody = "Good afternoon,`n`nFind attached the Algoma daily report.`n`nBest,`n`n"
$file = "C:\Users\kyle.conrad\Documents\Kyle Conrad\VBA Scripts\Algoma\Algoma Report.xlsx"

$oOutlook = New-Object -ComObject Outlook.Application

$oMailMsg = $oOutlook.CreateItem(0)

foreach ($addr in $sReceipientAddr){
    $null = $oMailMsg.Recipients.Add($addr)
}

$null = $oMailMsg.Recipients.Add($sCCAddr).Type = 2
$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = $sMsgBody
$oMailMsg.Attachments.Add($file)
$oMailMsg.Save()

$inspector = $oMailMsg.GetInspector
$inspector.Display()

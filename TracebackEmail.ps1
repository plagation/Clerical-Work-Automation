param([string] $traceback = "")

$time = (Get-Date).DateTime

$sReceipientAddr = "kyle.conrad@qsl.com"
$sMsgSubject = "Traceback Error of Exception"
$sMsgBody = "The following error occurred at $time under username $env:USERNAME :`n`n$traceback"
$oOutlook = New-Object -ComObject Outlook.Application

$oMailMsg = $oOutlook.CreateItem(0)

$null = $oMailMsg.Recipients.Add($sReceipientAddr)
$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = $sMsgBody
$oMailMsg.Save()
$oMailMsg.Send()
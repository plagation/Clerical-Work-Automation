param([string]$order = "")

$timenow = (Get-Date).TimeOfDay
$noon = (Get-Date -Hour 12 -Minute 0 -Second 0).TimeOfDay

if ($timenow -lt $noon){
    $greeting = "morning"
} else{
    $greeting = "afternoon"
}

$sReceipientAddr = "adurocher@boscus.com"
$sCCAddr = @("inventory@boscus.com", "nascofreight@qsl.com")
$sMsgSubject = "BOL for Order $order"
$Msgbody = "Good $greeting,`n`nFind attached the BOL for order $order.`n`nCheers,`n`n"

$oOutlook = New-Object -ComObject Outlook.Application

$oMailMsg = $oOutlook.CreateItem(0)
$null = $oMailMsg.Recipients.Add($sReceipientAddr)

foreach ($addr in $sCCAddr){
   $null = $oMailMsg.Recipients.Add($addr).Type = 2 
}

$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = $Msgbody
$oMailMsg.Save()

$inspector = $oMailMsg.GetInspector
$inspector.Display()
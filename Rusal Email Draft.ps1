param([string]$order = "RAC-015208", [string]$load = "3")

$timeNow = (Get-Date).TimeOfDay
$noon = (Get-Date -Hour 12 -Minute 0 -Second 0).TimeOfDay

if ($timeNow -lt $noon){
    $greeting = "morning"
} else{
    $greeting = "afternoon"
}

$sReceipientAddr = "shippingadvice@rusalamerica.com"
$sCCAddr = "nascofreight@qsl.com"
$sMsgSubject = "BOL for $order Load $load"

$oOutlook = New-Object -ComObject Outlook.Application 

$oMailMsg = $oOutlook.CreateItem(0)

$null = $oMailMsg.Recipients.Add($sReceipientAddr)  
$null = $oMailMsg.Recipients.Add($sCCAddr).Type = 2
$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = "Good $greeting,`n`nFind attached the BOL for $order load $load.`n`nCheers,`n`n"
$oMailMsg.Save()

$inspector = $oMailMsg.GetInspector
$inspector.Display()


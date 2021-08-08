param([string]$order = "BOS-0215153", [string]$weight = "80460")

$timeNow = (Get-Date).TimeOfDay
$noon = (Get-Date -Hour 12 -Minute 0 -Second 0).TimeOfDay

if ($timeNow -lt $noon){
    $greeting = "morning"
} else{
    $greeting = "afternoon"
}

$sReceipientAddr = "adurocher@boscus.com"
$sCCAddr = @("nascofreight@qsl.com", "inventory@boscus.com")
$sMsgSubject = "Driver Overweight - Advise Requested"

$oOutlook = New-Object -ComObject Outlook.Application 

$oMailMsg = $oOutlook.CreateItem(0)

$null = $oMailMsg.Recipients.Add($sReceipientAddr)  

foreach ($addr in $sCCAddr){
    $null = $oMailMsg.Recipients.Add($addr).Type = 2
}

$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = "Good $greeting Annie,`n`nDriver is here to pick up order $order and is overweight with a weight of $weight lbs. Please advise.`n`nCheers,`n`n"
$oMailMsg.Save()

$inspector = $oMailMsg.GetInspector
$inspector.Display()
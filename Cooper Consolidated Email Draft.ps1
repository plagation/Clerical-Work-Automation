$date = (Get-Date).ToShortDateString()
$timeNow = (Get-Date).TimeOfDay
$noon = (Get-Date -Hour 12 -Minute 0 -Second 0).TimeOfDay

if ($timeNow -lt $noon){
    $greeting = "morning"
} else{
    $greeting = "afternoon"
}

$sReceipientAddr = @("Tim.Oberkrom@cgb.com", "Chad.Sutter@cooperconsolidated.com", "Nick.Cruise@cgb.com", "Logistics3@cgb.com", "Brittany.Oalmann@cgb.com")
$sCCAddr = "nascofreight@qsl.com"
$sMsgSubject = "BOLs for $date"

$oOutlook = New-Object -ComObject Outlook.Application 

$oMailMsg = $oOutlook.CreateItem(0)

 
foreach ($addr in $sReceipientAddr){
    $null = $oMailMsg.Recipients.Add($addr) 
}

$null = $oMailMsg.Recipients.Add($sCCAddr).Type = 2
$oMailMsg.Subject = $sMsgSubject
$oMailMsg.Body = "Good $greeting,`n`nFind attached the BOLs for $date.`n`nCheers,`n`n"
$oMailMsg.Save()

$inspector = $oMailMsg.GetInspector
$inspector.Display()


#$ScriptPath = "C:\Users\kyle.conrad\Documents\Kyle Conrad\PowerShell\Cooper Consolidated Email Draft.ps1"
#$Trigger= New-ScheduledTaskTrigger -Daily -At 03:30pm
#$Action= New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-executionpolicy bypass -noprofile -file $ScriptPath" 
#Register-ScheduledTask -TaskName "MyTask" -Trigger $Trigger -User $env:username -Action $Action
$clientArr = @("RUSAL")
foreach ($client in $clientArr){
    $dirPath = "Z:\SCALE OFFICE\$client\Current Releases"
    Get-ChildItem -Path $dirPath -Name | ForEach {
        if ($client -eq "RUSAL"){
            $null = $_ -Match 'RAC-\d{6}' 
        }
        Rename-Item -Path "$dirPath\$_" -NewName ($Matches[0].toString() + ".pdf")
    }
}
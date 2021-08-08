<#
Get-ExecutionPolicy
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Get-Process #shows processes currently running on comp
Get-Service #shows list of sevices with their status
Get-Content C:\Windows\System32\drivers\etc\hosts #shows content of file you specify
#>
#Set-Location "C:\Users\kyle.conrad\Documents\PowerShell\" #sets the current work directory

param([string]$fileName = "")

function Convert-PDFtoText{
    param([Parameter(Mandatory=$true)][string]$file)
    Add-Type -Path "$PSScriptRoot\itextsharp.dll"
    $pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
    for ($page =1; $page -le $pdf.NumberOfPages; $page++){
        $text=[iTextSharp.text.pdf.parser.PDfTextExtractor]::GetTextFromPage($pdf,$page)
        Write-Output $text
    }
    $pdf.Close()
}

function Get-MatchPage{
    param([Parameter(Mandatory = $true)][string]$match)
    for ($page = 0; $page -lt $text.Length; $page++){
        If ($text[$page] -Match $match){
            break
        }
    }
    Write-Output $page
}

$text = convert-PDFtoText $fileName

$releaseInfo = @($text[0].Substring($text[0].IndexOf("BOS"),11))
$releaseInfo += $text[0].Substring($text[0].IndexOf("PO:")+3,12)


$page = Get-MatchPage "RELEASE"

$temp = $text[$page].Substring($text[$page].Indexof("RELEASE")-70,70)
$temp = $temp.SubString($temp.IndexOf("`n")+1).replace("  ", "")
$releaseInfo += $temp.Substring(0,$temp.IndexOf("`n"))

$page = Get-MatchPage "PQT/BDL"

$releaseInfo += $text[$page].SubString($text[$page].IndexOf("PQT/BDL")-10,10).replace(" ", "")

Remove-Item "$PSScriptRoot\output.txt"

for ($i =0; $i -lt $releaseInfo.Length; $i++){
    if ($i -eq 0){
        $releaseInfo[$i] | Out-File -FilePath "$PSScriptRoot\output.txt"
    } else{
        $releaseInfo[$i] | Add-Content -Path "$PSScriptRoot\output.txt"
    }
}

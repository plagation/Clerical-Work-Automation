param([string] $filename = "C:\Users\kyle.conrad\Downloads\LC-064631-UNITED STATES STEEL CORP.pdf")

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

$pageText = Convert-PDFtoText $filename
$text = ""

foreach ($page in $pageText){
    $text += $page
}

$textLines = $text.split("`n")

$bolNum = $textLines[0]

Select-String 'Lot\s:\sMKT.[A-Z]{2}.\d{4}' -InputObject $text -AllMatches | Foreach {$lot = $_.Matches.Value.Replace("Lot : ", "")}

Select-String '\n[A-Z]{2}\n' -InputObject $text -AllMatches | Foreach {$clerk = $_.Matches.Value.Replace("`n", "")}

Select-String 'NET\s\d{4,5}' -InputObject $text -AllMatches | ForEach {$weight = $_.Matches.Value.Replace("NET ", "")}

$date = $textLines[1].split(" ")[0]

for ($i = 0; $i -lt $textLines.Length; $i++){
    if ($textLines[$i] -Match 'Transportation Carrier Carrier'){
        $transportation = $textLines[$i+1].split(" ")[0]
        break
    }
}

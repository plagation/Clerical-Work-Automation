Add-Type -AssemblyName System.Windows.Forms

$workingPath = $PSScriptRoot
$PSScriptRoot
Import-Module "$($workingPath)\WebDriver.dll"
Import-Module "$($workingPath)\WebDriver.Support.dll"

function waitForElement($locator, $timeInSeconds,[switch]$byClass,[switch]$byName,[switch]$byXPath){
    $webDriverWait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($Script:driver, (New-TimeSpan -Seconds $timeInSeconds))
    try{
        if($byClass){
            $null = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::ClassName($locator)))
        }
        elseif($byName){
            $null = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Name($locator)))
        }
        elseif($byXPath){
            $null = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::XPath($locator)))
        }
        else{
            $null = $webDriverWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementIsVisible( [OpenQA.Selenium.by]::Id($locator)))
        }
        return $true
    }
    catch{
        return "Wait for $locator timed out"
    }
}

function waitToClick($locator){
    waitForElement "$locator" 10 -byXPath
    $Script:driver.FindElementByXPath("$locator").Click()
}

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

$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
InitialDirectory = [Environment]::GetFolderPath('Desktop')
Filter = 'PDFs (*.pdf)|*.pdf'}

$fileBrowser.Multiselect = $true
$fileBrowser.ShowDialog() | Out-Null
$fileBrowser.FileNames

$recieverArr = @(
    @("4400 COUNTY ROAD 59","HEIDTMAN STEEL PRODUCTS`n4400 COUNTY ROAD 59`nBUTLER, IN`nUSA 46721"),
    @("2500 EUCLID AVENUE","ESMARK STEEL GROUP-MIDWEST, LLC`nC/O SUN STEEL COMPANY`n2500 EUCLID AVENUE`nCHICAGO HEIGHTS, IL`nUSA 60411"),
    @("4435 SOUTH WESTERN BLVD","WHEATLAND TUBE , LLC`nA DIVISION OF ZEKELMAN IND, INC`n4435 SOUTH WESTERN BLVD`nCHICAGO, IL`nUSA 60609"),
    @("46368-1383 6755 WATERWAY","VIKING MATERIALS INC.`nC/O FERALLOY MIDWEST`n46368-1383 6755 WATERWAY`nPORTAGE, IN`nUSA"),
    @("700 CENTRAL AVENUE","ESMARK STEEL GROUP-MIDWEST, LLC`nC/O CHICAGO STEEL AND IRON LLC`n700 CENTRAL AVENUE`nUNIVERSITY PARK, IL`nUSA 60454"),
    @("4407 RAILROAD AVENUE","HEIDTMAN STEEL PRODUCTS`n4407 RAILROAD AVENUE`nEAST CHICAGO, IN`nUSA 46312"),
    @("11305 FRANKLIN AVENUE","VIKING MATERIALS INC`n11305 FRANKLIN AVENUE`nFRANKLIN PARK, IL`nUSA 60131"),
    @("330 EAST JOE ORR RD BUILDING C","JDM STEEL SERVICE LTD`n330 EAST JOE ORR RD BUILDING C`nCHICAGO HEIGHTS, IL`nUSA 60411"),
    @("7201 S 78TH STREET","SIGNODE`n7201 S 78TH STREET`nBRIDGEVIEW, IL`nUSA 60455"),
    @("701 LOOP ROAD","ASI C/O VOSS CLARK - INDIANA`nVOSS CLARK`n701 LOOP ROAD`nJEFFERSONVILLE, IN`nUSA 47130-8428")
)

$orderMap = @{
    "HEIDTMAN STEEL PRODUCTS`n4400 COUNTY ROAD 59`nBUTLER, IN`nUSA 46721" = "Heidtman Butler IN - 48365"
    "ESMARK STEEL GROUP-MIDWEST, LLC`nC/O SUN STEEL COMPANY`n2500 EUCLID AVENUE`nCHICAGO HEIGHTS, IL`nUSA 60411" = "Esmark c/o Sun Steel - 14016" 
    "WHEATLAND TUBE , LLC`nA DIVISION OF ZEKELMAN IND, INC`n4435 SOUTH WESTERN BLVD`nCHICAGO, IL`nUSA 60609" = "WHEATLAND TUBE, LLC - 74629"
    "VIKING MATERIALS INC.`nC/O FERALLOY MIDWEST`n46368-1383 6755 WATERWAY`nPORTAGE, IN`nUSA" = "VIKING MATERIALS INC. C/O FERALLOY MIDWEST - 57504"
    "ESMARK STEEL GROUP-MIDWEST, LLC`nC/O CHICAGO STEEL AND IRON LLC`n700 CENTRAL AVENUE`nUNIVERSITY PARK, IL`nUSA 60454" = "Esmark c/o CSI University Park - 19843"
    "HEIDTMAN STEEL PRODUCTS`n4407 RAILROAD AVENUE`nEAST CHICAGO, IN`nUSA 46312" = "Heidtman East Chicago, IN - 41941"
    "VIKING MATERIALS INC`n11305 FRANKLIN AVENUE`nFRANKLIN PARK, IL`nUSA 60131" = "VIKING MATERIALS INC - FRANKLIN PARK - 56427"
    "JDM STEEL SERVICE LTD`n330 EAST JOE ORR RD BUILDING C`nCHICAGO HEIGHTS, IL`nUSA 60411" = "JDM Steel Chicago Heights, IL - 65714"
    "SIGNODE`n7201 S 78TH STREET`nBRIDGEVIEW, IL`nUSA 60455" = "Signode - 84261"
    "ASI C/O VOSS CLARK - INDIANA`nVOSS CLARK`n701 LOOP ROAD`nJEFFERSONVILLE, IN`nUSA 47130-8428" = "ASI C/O Voss Clark - 20043"
}

$typeMap = @{
    "CR STEEL SHEET" = "Cold Rolled Steel Sheet"
    "HR STEEL SHEET" = "Hot Rolled Steel Sheet"
    "HR FLOOR PLATE" = "HR FLOOR PLATE"
}

$clerk = ""
foreach ($name in $env:USERNAME.Split(".")){
    $clerk += $name.SubString(0,1).toUpper()
}

#Start web driver instance

$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
$driver.Manage().Window.Maximize()
$driver.Navigate().GoToUrl("https://tos.qsl.com/")
waitForElement '//*[@id="username"]' 5 -byXPath
$driver.FindElementByXPath('//*[@id="username"]').SendKeys("$env:USERNAME@qsl.com")
$driver.FindElementByXPath('/html/body/div/div[2]/main/section/div/div/div/form/div[2]/button').Click()

waitForElement 'tc3-menu' 45

$driver.Manage().Window.Minimize()


#loop through each selected file
foreach ($file in $fileBrowser.FileNames){
    $pdf = Convert-PDFtoText -file "$file"
    $railCar = $bol = $offloadDate = ""
    
    if ($pdf.length -gt 10){
        $pdfStr = ""
        for ($i = 0; $i -lt $pdf.length; $i++){
            $pdfStr += $pdf[$i]
        }
        $pdf = @($pdfStr) 
    }

    #Check if file format is supported
    if ($pdf[0] -Match 'ADVANCED SHIPPING NOTIFICATION'){
    
        $null = $pdf[0].Substring($pdf[0].IndexOf("USD FUNDS")+11, 20) -match  '[A-Z]{2,4}\s[0-9]+'
        $railcar = $Matches[0]
        $null = $pdf[0].Substring($pdf[0].IndexOf($railCar) + $railCar.Length, 15) -match '[0-9]+'
        $bol = $Matches[0]

        do {
            $offloadDate = Read-Host "Enter railcar $railcar offload date"
            if (!($offloadDate -match '\d{2}/\d{2}/\d{4}')){
                Write-Host "Invalid date format! Enter date as MM/DD/YYYY!"
            }
        } while (!($offloadDate -match '\d{2}/\d{2}/\d{4}'))

        $remarks = "$railcar`nOffloaded on $offloadDate"

        $row = 3
        $xl = New-Object -ComObject Excel.Application
        $wb = $xl.WorkBooks.Open("$PSScriptRoot\import-manifest.xlsx")
        $ws = $wb.WorkSheets.Item(1)
        $xl.Visible = $false

        do {
            $ws.Cells.Item(3,1).EntireRow.Delete()
        } while ($ws.Cells.Item(3,1).text.length -ne 0)

        #loop through each page in the file
        for ($i = 0; $i -lt $pdf.length; $i++){
            $poNum = $itemNum = $weight = $reciever = $heat = $mark = $width = $thickness = $coilType = ""

            #Get PO Number
            $pdf[$i] = $pdf[$i].Substring($pdf[$i].IndexOf($bol) + 2*$bol.Length + 2, $pdf[$i].Length-$pdf[$i].IndexOf($bol)-2*$bol.Length-2)
            $poNum = $pdf[$i].Substring(0,$pdf[$i].IndexOf("`n"))
            
            #Get reciever
            if ($poNum -eq 2202011){
                $reciever = $recieverArr[4][1]
            } else {
                for ($j = 0; $j -lt $recieverArr.length;$j++){
                    if ($pdf[$i] -match $recieverArr[$j][0]){
                        $reciever = $recieverArr[$j][1]        
                        break
                    }
                }
            }

            $pdf[$i] = $pdf[$i].Substring($pdf[$i].IndexOf("DESCRIPTION"),$pdf[$I].length-$pdf[$i].IndexOf("DESCRIPTION"))

            #get coil heats
            $heats = @()
            Select-String '[0-9]{4}\w{2}\s[0-9]{2}' -InputObject $pdf[$i] -AllMatches | Foreach {$heats += $_.Matches.Value}
   
            #get coil marks
            $marks = @()
            Select-String '[A-Z]{3}[0-9]{4}[A-Z]{0,1}' -InputObject $pdf[$i] -AllMatches | ForEach {$marks += $_.Matches.Value}

            #get coil weights
            $weights = @()
            Select-String '\s[0-9]{2},[0-9]{3}' -InputObject $pdf[$i] -AllMatches | Foreach {$weights += $_.Matches.Value.Trim().Replace(",","")}        
            if ($heats.length -ne $weights.Length){
                $coils = @()
                for ($j = 0; $j -lt $weights.length-1; $j++){
                    $coils += $weights[$j]
                }
                $weights = $coils
            }
            
            #get coil item numbers
            $itemNums = @()
            Select-String '[NO]{2}.:[A-Z0-9/-]+' -InputObject $pdf[$i] -AllMatches | ForEach {$itemNums += $_.Matches.Value.Trim().Replace("NO.:","")}

            #get coil thicknesses and widths
            $measurements = @()
            $parsedMeasurements = @()
            $thicknesses = @()
            $widths = @()

            Select-String '(\sx\s){0,1}(:\s){0,1}\d{1,2}.\d{3,4}"' -InputObject $pdf[$i] -AllMatches | ForEach {$parsedMeasurements += $_.Matches.Value.Replace(": ","").Replace(" x ", "").Replace('"', "")}

            for ($j = 0; $j -lt $parsedMeasurements.Length; $j++){
                if (($j + 1) % 2 -eq 0){
                    $widths += $parsedMeasurements[$j]
                } else {
                    $thicknesses += $parsedMeasurements[$j]
                }
            }

            #Get coil typing
            $coilTypes = @()
            for ($j = 0; $j -lt $itemNums.Length; $j++){
                $pdf[$i] = $pdf[$i].Substring($pdf[$i].indexOf($itemNums[$j]), $pdf[$i].Length - $pdf[$i].indexOf($itemNums[$j]))
                $coilTypes += $pdf[$i].Split("`n")[1].Trim()
            }

            #iterate through each coil on current page adding to the site
            $ws.Columns.Item(17).NumberFormat = '@'
            for($j = 0; $j -lt $weights.Length; $j++){
                $ws.Cells.Item($row,1)= "Algoma 2021"
                $ws.Cells.Item($row,2) = $orderMap.Get_Item("$reciever")
                $ws.Cells.Item($row,3) = "Steel Coils USA"
                $ws.Cells.Item($row,4) = $typeMap.Get_Item($coilTypes[$j])
                $ws.Cells.Item($row,5) = "Loose"
                $ws.Cells.Item($row,6) = 1
                $ws.Cells.Item($row,7) = "Can be containerized"
                $ws.Cells.Item($row,9) = 1
                $ws.Cells.Item($row,10) = "Mark"
                $ws.Cells.Item($row,11) = $marks[$j]
                $ws.Cells.Item($row,12) = "HeatNumber"
                $ws.Cells.Item($row,13) = $heats[$j]
                $ws.Cells.Item($row,14) = "Scope"
                $ws.Cells.Item($row,15) = "$bol / $poNum / " + $itemNums[$j]
                $ws.Cells.Item($row,16) = "Other"
                $ws.Cells.Item($row,17) = $offloadDate
                $ws.Cells.Item($row,18) = "Thickness"
                $ws.Cells.Item($row,19) = $thicknesses[$j]
                $ws.Cells.Item($row,20) = "in"
                $ws.Cells.Item($row,21) = "Width"
                $ws.Cells.Item($row,22) = $widths[$j]
                $ws.Cells.Item($row,23) = "in"
                $ws.Cells.Item($row,32) = $weights[$j]
                $ws.Cells.Item($row,33) = "lb"
                $ws.Cells.Item($row,34) = "Rail Building"
                $row++
            }

            $wb.Save()     
            
            if ($weights.Length -eq 1) {
                $remarks += "`n`n1 coil for:`n$reciever"
            } else {
                $remarks += "`n`n" + $weights.Length + " coils for:`n$reciever"
           }
            #After adding coils modify special instructions with reciever and coils for reciever
        }
        $remarks += "`n$clerk"

        #Create new reception
        
        $driver.Manage().Window.Maximize()

        $driver.Navigate().GoToUrl("https://tos.qsl.com/client-inventories/receptions-of-materials")

        waitToClick '//*[@id="viewport"]/article/section/section[1]/div[2]/div[2]/button'
        waitToClick '/html/body/div[3]/div/form/div/div[2]/label[3]'
        $driver.FindElementByXPath('/html/body/div[3]/div/form/header/menu/button[2]').Click()
       
        waitForElement '//*[@id="react-select-5-input"]' 10 -byXPath
        $driver.FindElementByXPath('//*[@id="react-select-5-input"]').SendKeys("CN" + [OpenQA.Selenium.Keys]::Enter)
        $driver.FindElementByXPath('//*[@id="driverName"]').SendKeys($clerk)
        $driver.FindElementByXPath('//*[@id="carrierBill"]').SendKeys("Add In")
        $driver.FindElementByXPath('//*[@id="transportationNumber"]').SendKeys($railCar)
        $driver.FindElementByXPath('//*[@id="specialInstructions"]').SendKeys($remarks)
        $driver.FindElementByXPath('//*[@id="react-select-13-input"]').SendKeys("Algoma" + [OpenQA.Selenium.Keys]::Enter)
        Start-Sleep -Milliseconds 500

        $driver.FindElementByXPath('//*[@id="viewport"]/article/section/form/section[2]/div/div[2]/button[2]').Click()
        Start-Sleep -Seconds 1

        $driver.Navigate().GoToUrl($driver.Url.Substring(0,$driver.Url.IndexOf("general")) + "incoming-items")
        waitForElement '//*[@id="viewport"]/article/section/section[3]/div[2]/div[4]/button' 10 -byXPath

        $driver.FindElementByXPath('//*[@id="viewport"]/article/section/section[3]/div[2]/div[1]').Click()
        Start-Sleep -Milliseconds 250

        $driver.FindElementByCssSelector("input[type='file']").SendKeys("$PSScriptRoot/import-manifest.xlsx")

        $wb.Close()
        
    }else {
        "File Format Not Supported!"
        $driver.quit()
        exit 1
    }
}

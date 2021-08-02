param([string] $bol = "")

Start-Process -FilePath "C:\Users\$env:USERNAME\Downloads\LR-$bol.pdf" -Verb Print -PassThru | %{Sleep 10;$_} | Kill
Remove-Item -Path "C:\Users\$env:USERNAME\Downloads\LR-$bol.pdf"
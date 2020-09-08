
$Ex_Start_Excel = { $exl = New-Object -ComObject Excel.Application }
$Ex_Open_Book = {
    param($book)
    if ( Test-Path $book ) {
        $bk = $exl.Workbooks.Open($book)
    }
    else {
        $bk = $exl.Workbooks.Add()
        $bk.SaveAs($book)
    }
}

$Ex_Start_Excel_With_Book = { param($book); . $Ex_Start_Excel ; . $Ex_Open_Book -book $book }

$Ex_Close_Book = { $bk.Close(); $bk = $null }
$Ex_Quit_Excel = { $exl.Quit(); $exl = $null }
$Ex_Quit_Excel_With_Book = { . $Ex_Close_Book ; . $Ex_Quit_Excel }

Get-ChildItem $PSScriptRoot | ?{ $_.Extension -eq ".ps1" } | %{ . $_.FullName }
Export-ModuleMember -Function '*'
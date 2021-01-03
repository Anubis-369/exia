Function Add-ExDescription {
    <#
    .SYNOPSIS
    Listオブジェクトにタイトルや備考を追加する。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトから書き込みを行う。

    .EXAMPLE
    Add-Desctiption -book $bk -Title <Title> -Description <Description>

    .PARAMETER Book
    書き込みをするBook。以下のように事前にオブジェクトを作成しておく。
    $exl = $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workbooks.Open("Excelファイルのパス")

    .PARAMETER Table
    TitleやDescriptionを挿入するListObjectの名前。

    .PARAMETER Title
    タイトルに入れる文字列。空白の場合、タイトルは挿入されない。

    .PARAMETER Description
    備考として入れる文字列。空白の場合、備考は挿入されない。

    .PARAMETER Label
    タイトルや備考のセルに付けられる名前。空白の場合、名前は付けられない。

    #>
    param(
    [Parameter(Mandatory = $true)]$Book,
    [Parameter(Mandatory = $true)][string]$Table,
    [string]$Title="",
    [string]$Discription="",
    [string]$Label=""
    )

    $Table_Data = $Book.Worksheets | % { $_.ListObjects } | ? { $_.Name -eq $Table }
    if ($Table_Data.count -eq 0) { Write-Warning "$Table not Exist!"; return }

    $pt = $Table_Data.HeaderRowRange(1).Address()
    if($Title -ne ""){
        $Table_Data.HeaderRowRange(1).EntireRow.Insert()
        $Table_Data.Parent.Range($pt) | %{
            $_.Value2 = $Title
            if ($Label -ne "") { $_.Range($pt).Name = ("{0}_Title" -f $Label) }
        }
    }

    if($Discription -ne ""){
        $Table_Data.HeaderRowRange(1).EntireRow.Insert()
        $Table_Data.Parent.Range($pt) | %{
            $_.Resize(1,$Table_Data.HeaderRowRange.count).Merge()
            $_.HorizontalAlignment = -4131
            $_.VerticalAlignment   = -4162
            $_.Value2 = $Title
            if ($Label -ne "") { $_.Range($pt).Name = ("{0}_Description" -f $Label) }
        }
    }
}


Function Write-ExPSOToList {
    <#
    .SYNOPSIS
    PSObjectの内容をCOM経由で開いているExcelのファイルにList形式で書き込む。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトから書き込みを行う。

    .EXAMPLE
    <PSObject> | Write-ExPSOToList -Sheet $ws -Address B2 -Title Name

    .EXAMPLE
    Write-ExPSOToList -PSObject <PSObject> -Sheet $ws -Address B2 -Title Name

    .PARAMETER Sheet
    書き込みを行うExcel上のシートのオブジェクトを渡す。
    シートのオブジェクトは以下のように生成する。
    $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workbooks.Open("Excelファイルのパス")
    $ws = $bk.Worksheets(1)

    .PARAMETER Address
    書き込みを行うアドレスをExcel上のセルのアドレスで指定。

    .PARAMETER PSObject
    書き込みを行うPSObject。パイプからも入力できる。

    .PARAMETER Title
    書き込んだ Listのタイトルに指定するメンバーの名前。

    .PARAMETER Label
    書き込んだセルに付けられるセルの名前。タイトル部分に <Label>_Title_<Number> の名前が振られ、
    データ部分には、　<Label>_<Number> の名前が付けられる。

    .PARAMETER Members
    データを書き込むPSObjectのメンバー。

    .PARAMETER Headers
    Excelに書き込まれるときの見出し。

    .PARAMETER Format
    データを書き込んだセルに設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        $Sheet,
        [string]$Address = "A1",
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [psobject[]]$PSObject,
        [string]$Title = "",
        [string]$Label = "",
        [int]$tmargin=0,
        [int]$vmargin=0,
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )

    begin {
        $pt = $Sheet.Range("$Address")
        $count = 1
    }

    Process {
        if ($Members.count -eq 0) { $Members = $PSObject | get-member -Type NoteProperty | % Name }
        if ($Title -ne "" ) {$Members = $Members | ? { $_ -ne $Title }}
        if ($Header.count -eq 0) { $Header = $Members }
        
        $PSObject | % {
            $ps = $_
            if ($Title -ne "") {
                $pt.Value2 = $_ | % $Title
                if ( $Label -ne "" ) { $pt.Name = "{0}_Title_{1}" -f $Label, $count }
                $pt.Font.Bold = $True
                [void]$pt.Offset(1, 0).EntireRow.Insert()
                $pt = $pt.Offset(1, 0)
                $pt.Font.Bold = $false
            }

            for($i=0;$i -lt $Members.count;$i++) {
                $pt.Value2 = ($Header[$i] + ":")
                $pt.Offset(0, (1 + $vmargin)) | % {
                    $_.Value2 = [string]($ps | % $Members[$i])
                    $_.WrapText = $False
                    if (($Format.count -ne 0 ) -and ($Format.Keys -contains $Members[$_])) {
                        $_.NumberFormatLocal = $Format[$Members[$i]]
                    }
                }
                $pt.Resize(1, (2 + $tmargin + $vmargin )).Borders(9) | % { $_.LineStyle = 1; $_.Weight = -4138 }
                [void]$pt.Offset(1, 0).EntireRow.Insert()
                $pt = $pt.Offset(1, 0)
            }

            if ( $Label -ne "" ) {
                $pt.Offset((1 - $Members.count), 0).Resize(($Members.count - 1), 2).Name = `
                    "{0}_{1}" -f $Label, $count
            }

            $count += 1
            [void]$pt.Offset(1, 0).EntireRow.Insert()
            $pt = $pt.Offset(1, 0)
        }
    }
    End {
        return $pt
    }
}

Function Write-ExPSOToTable {
    <#
    .SYNOPSIS
    PSObjectの内容をCOM経由で開いているExcelのファイルにテーブル形式で書き込む。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトから書き込みを行う。

    .EXAMPLE
    <PSObject> | Write-ExPSOToTable -Sheet $ws -Address B2

    .EXAMPLE
    Write-ExPSOToTable -PSObject <PSObject> -Sheet $ws -Address B2

    .PARAMETER Sheet
    書き込みを行うExcel上のシートのオブジェクトを渡す。
    シートのオブジェクトは以下のように生成する。
    $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workbooks.Open("Excelファイルのパス")
    $ws = $bk.Worksheets(1)

    .PARAMETER Address
    書き込みを行うアドレスをExcel上のセルのアドレスで指定。

    .PARAMETER PSObject
    書き込みを行うPSObject。パイプからも入力できる。

    .PARAMETER Label
    書き込んだリスト形式の表に付けられる名前。空白にした場合、ExccelのListが設定されない。

    .PARAMETER Members
    データを書き込むPSObjectのメンバー。

    .PARAMETER Headers
    Excelに書き込まれるときの見出し。

    .PARAMETER Format
    データを書き込んだセルに設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        $Sheet,
        [string]$Address = "A1",
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [psobject[]]$PSObject,
        [string]$Label = "",
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )
    begin {
        $pt = $Sheet.Range("$Address")
        $Data_Count = 1
        $first_flag = $true
    }
    process {
        if ( $first_flag -eq $true ) {
            if ($Members.count -eq 0) { $Members = $PSObject | get-member -Type NoteProperty | % Name }
            if ($Header.count -eq 0) { $Header = $Members }

            $header_Cell = $pt.Resize(1, $Header.count)
            $header_Cell.Value2 = $Header
            $lo = $Sheet.ListObjects.Add(1, $header_Cell, 0, 1)
            $first_flag = $false
        }

        $PSObject | %{
            $pso = $_
            $Members | % { $val = @() } { $Val += ($pso | % $_) } {}
            $row = $lo.ListRows.Add($Data_Count)
            @(1..$val.count) | % {
                $row.Range($_).Value2 = [string]$val[($_ - 1)]
            }
            $Data_Count += 1
        }
    }

    End {
        if ($Format.Count -ne 0) {
            $Format.Keys | % {
                $lo.ListColumns | ? { $_.Name -eq $_ } | `
                    % { $_.Range.NumberFormatLocal = $Format[$_] }
            }
        }

        if ( $Label -ne "") {
            $lo.Name = $Label
        }
        else {
            $lo.unlist()
        }
        return $pt.Offset($Data_Count + 1, 0)
    }
}

Function Write-ExPSOToXlsxTable {
    <#
    .SYNOPSIS
    PSObjectの内容をExcelを起動してファイルにテーブル形式で書き込む。

    .Descript
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    Excelの起動処理、終了処理をコマンド内で完結させているため、前処理と後処理なしで書き込みができる。

    .EXAMPLE
    <PSObject> | Write-ExPSOToXlsxTable -File <File Path> -Sheet <Worksheet Name> -Address B2

    .EXAMPLE
    Write-ExPSOToXlsxTable -File <File Path> -PSObject <PSObject> -Sheet <Worksheet Name> -Address B2

    .PARAMETER File
    書き込みを行うファイルのパス

    .PARAMETER Address
    書き込みを行うアドレスをExcel上のセルのアドレスで指定。

    .PARAMETER Sheet
    書き込みを行うExcel上のシートの名前。

    .PARAMETER Label
    書き込んだリスト形式の表に付けられる名前。空白にした場合、ExccelのListが設定されない。

    .PARAMETER Members
    データを書き込むPSObjectのメンバー。

    .PARAMETER Headers
    Excelに書き込まれるときの見出し。

    .PARAMETER Format
    データを書き込んだセルに設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        [string]$File,
        [string]$Sheet = "",
        [string]$Address = "A1",
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [psobject[]]$PSObject,
        [string]$Label = "",
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )

    if ( $input.count -eq 0 ) { $pso = $PSObject } else { $pso = $input }
    . $Ex_Start_Excel_With_Book -book $File

    if ($Sheet -eq "") {
        $ws = $bk.Worksheets(1)
    }
    else {
        $ws = $bk.Worksheets($Sheet)
    }
    $para = @{
        Sheet     = $ws;
        Address   = $address;
        Tablename = $Label;
        Members   = $Members;
        Header    = $Header;
        Format    = $Format
    }
    $pso | Write-ExPSOToTable @para
    $bk.save()
    . $Ex_Quit_Excel_With_Book
}

Function Write-ExGPOToTable {
    <#
    .SYNOPSIS
    PSObjectの内容をグルーピングしてCOM経由で開いているExcelのファイルにテーブル形式で書き込む。

    .Description
    Group-Objectでグルーピングして、Nameをタイトルにして、Groupの内容をListとして書き出す。
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトから書き込みを行う。

    .EXAMPLE
    <PSObject> | Write-ExGPOToTable -Property <Member> -Sheet $ws -Address B2

    .EXAMPLE
    Write-Write-ExGPOToTable -$Property <Member> -PSObject <PSObject> -Sheet $ws -Address B2

    .PARAMETER Sheet
    書き込みを行うExcel上のシートのオブジェクトを渡す。
    シートのオブジェクトは以下のように生成する。
    $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workbooks.Open("Excelファイルのパス")
    $ws = $bk.Worksheets(1)

    .PARAMETER Property
    グルーピングするメンバーの名前。

    .PARAMETER Address
    書き込みを行うアドレスをExcel上のセルのアドレスで指定。

    .PARAMETER PSObject
    書き込みを行うPSObject。パイプからも入力できる。

    .PARAMETER Title
    グルーピングしたメンバーを入力したセルの名前。

    .PARAMETER Label
    書き込んだリスト形式の表に付けられる名前。空白にした場合、ExccelのListが設定されない。

    .PARAMETER Members
    データを書き込むPSObjectのメンバー。

    .PARAMETER Headers
    Excelに書き込まれるときの見出し。

    .PARAMETER Format
    データを書き込んだセルに設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        [Parameter(Mandatory = $true)]
        $Sheet,
        [Parameter( Mandatory = $true)]
        [string]$Property,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [psobject[]]$PSObject,
        [string]$Address = "A1",
        [string]$Label = "",
        [string[]]$Names = @(),
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )

    if ($input.count -eq 0) {
        $gp = $PSObject | Group-Object $Property
    } else {
        $gp = $input | Group-Object $Property
    }
    
    if ($Names.count -eq 0) { $Names = $gp | % Name }
    $Members = $Members | ?{$_ -ne $Property }

    $pt = $Sheet.range($Address)
    $count = 1
    
    $Names | % {
        $n = $_
        $gp | ? { $_.Name -eq $n } | % {
            $pt.Value2 = $n
            [void]$pt.Offset(1, 0).EntireRow.Insert()
            $pt = $pt.Offset(1, 0)

            if ( $Label -ne "" ) {
                $pt.Name = "{0}_{1}_Title" -f $Title, $count
                $Child_Label = "{0}_{1}_List" -f $Label, $count
            }

            $para = @{
                Sheet   = $Sheet;
                Address = $pt.Address();
                Label   = $Child_Label;
                Members = $Members;
                Header  = $Header;
                Format  = $Format
            }
            
            $pt = $_.Group | Write-ExPSOToTable @para
            $count += 1
        }
    }
    return $pt
}

Function Write-ExValToXlsx {
    <#
    .SYNOPSIS
    名前付きのセルに値を書き込む

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクト生成しておき、
    $exl.Workbooks.Open($book) からあらかじめ開いたブックから書き込みを行う。

    .EXAMPLE
    Write-ExValToXlsx -book $bk -Name <Cell Name> -Value

    .EXAMPLE
    Write-Write-ExGPOToTable -$Property <Member> -PSObject <PSObject> -Sheet $ws -Address B2

    .PARAMETER Book
    書き込みをするBook。以下のように事前にオブジェクトを作成しておく。
    $exl = $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workbooks.Open("Excelファイルのパス")

    .PARAMETER Name
    書き込みをするセルについている名前

    .PARAMETER Value
    書き込みする値

    .PARAMETER Format
    データを書き込んだセルに設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        [Parameter(Mandatory = $true)]$Book,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Value,
        [string]$Format = ""
    )
    $All_Address = $Book.Names | ? { $Name -eq $_.Name } | % Value
    $GP = ([regex]"=(?<sheet>.+)!(?<address>.*)").Matches($All_Address) | % Groups
    $Sheet = $GP | ? { $_.Name -eq "sheet" } | % Value
    $Address = $GP | ? { $_.Name -eq "address" } | % Value
    if ($Format -ne "") {
        $Book.Worksheets($Sheet).Range($Address).NumberFormatLocal = $Format
    }
    $Book.Worksheets($Sheet).Range($Address).Cells.Value2 = $Value
}


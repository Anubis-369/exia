Function Read-ExTableToPSO {
    <#
    .SYNOPSIS
    Excelファイル上のListのデータを読み込んでPSObjectに書き出す。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトからファイルの読み込みを行う。

    .EXAMPLE
    Read-ExTableToPSO -Book $bk -Table <Table Name> 

    .PARAMETER Book
    書き込みをするBook。以下のように事前にオブジェクトを作成しておく。
    $exl = $exl = New-Object -ComObject Excel.Application
    $bk = $exl.Workboo　　ks.Open("Excelファイルのパス")

    .PARAMETER Table
    読み込みを行うテーブルの名前。

    .PARAMETER Members
    データを読み込んだ時につけるPSObjectのメンバー名。

    .PARAMETER Headers
    ExcelのListObjectの読み込みを行うデータの見出し。

    .PARAMETER Format
    データを読み込むときに一時的に設定するフォーマット。
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        $Book,
        [string]$Table,
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )
    $Table_Data = $Book.Worksheets | % { $_.ListObjects } | ? { $_.Name -eq $Table }
    if ($Table_Data.count -eq 0) { Write-Warning "$Table not Exist!"; return }

    if ($Format.Count -ne 0) {
        $Format.Keys | % {
            $Table_Data.ListColumns | ? { $_.Name -eq $_ } | `
                % { $_.Range.NumberFormatLocal = $Format[$_] }
        }
    }

    if ($header.count -eq 0) { $Header = $Table_Data.HeaderRowRange | % text }
    if ($Members.count -eq 0) { $Members = $Header }

    $Data_Array = @()

    $Header | % {
        if ( ($Table_Data.ListColumns | % Name) -contains $_) {
            $Data_Array += $Table_Data.ListColumns($_)
        }
        else { Write-Warning "Header $_ is Wrong!"; return }
    }

    @(2..($Table_Data.ListRows.count + 1)) | % {
        $data_count = $_
        $elm = New-Object -TypeName psobject
        if ($Members.Count -ge $Header.count ) {
            $count = $Header.count - 1
        }
        else {
            $count = $Members.Count - 1
        }

        @(0..$count) | % {
            $elm | Add-Member -Type NoteProperty `
                -Name $Members[$_] `
                -Value $Data_Array[$_].range[$data_count].text
        }
        $elm
    }
}

Function Read-ExXlsxTableToPSO {
    <#
    .SYNOPSIS
    Excelを起動してファイル上のListのデータを読み込んでPSObjectに書き出す。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    Excelの起動処理、終了処理をコマンド内で完結させているため、前処理と後処理なしで書き込みができる。

    .EXAMPLE
    Read-ExXlsxTableToPSO -Book <File Path> -Table <Table Name> 

    .PARAMETER Book
    データの読み込みを行うファイルのパス

    .PARAMETER Table
    読み込みを行うテーブルの名前。

    .PARAMETER Members
    データを読み込んだ時につけるPSObjectのメンバー名。

    .PARAMETER Headers
    ExcelのListObjectの読み込みを行うデータの見出し。

    .PARAMETER Format
    データを読み込むときに一時的に設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$BookPath,
        [string]$Table,
        [string[]]$Members = @(),
        [string[]]$Header = @(),
        [hashtable]$Format = @{}
    )

    . $Ex_Start_Excel_With_Book -book $BookPath

    $para = @{
        Book    = $bk;
        Table   = $Table;
        Members = $Members;
        Header  = $Header;
        Format  = $Format
    }

    Read-ExTableToPSO @para

    . $Ex_Quit_Excel_With_Book
}

Function Read-ExXLSXToVal {
    <#
    .SYNOPSIS
    Excelファイル上の名前付きセルから値を抜き出す。

    .Description
    このコマンドは、ExcelがインストールしてあるWindowsのシステムで使用できる。
    New-Object -ComObject Excel.Applicationで生成したExcelのCOMオブジェクトからファイルの読み込みを行う。

    .EXAMPLE
    $exl = $exl = New-Object -ComObject Excel.Application
    Read-ExXLSXToVal -Excel $exl -Book <File Path> -Name <Cell Name>

    .PARAMETER Excel
    ExcelのCOMオブジェクト

    .PARAMETER Book
    データの読み込みを行うファイルの名前。複数指定、パイプラインからの入力可。

    .PARAMETER Name
    読み込みを行うセルの名前。複数指定可。

    .PARAMETER Format
    データを読み込むときに一時的に設定するフォーマット
    $Format = @{リストのカラム名 = フォーマット;リストのカラム名 = フォーマット;...}
　　 の形式で指定する。

    #>
    param(
        $Excel,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string]$Book,
        [string[]]$Name = @(),
        [hashtable]$Format
    )
    begin {}
    Process {
        $Book | %{
            $bk = $Excel.Workbooks.Open($_)
            $Book_Name = $bk.Name
            if ($Name.count -eq 0) { $list = $bk.Names } `
            else { $list = $bk.Names | ? { $Name -contains $_.Name }}

            $list | % {
                $GP = ([regex]"=(?<sheet>.+)!(?<address>.*)").Matches($_.Value) | % Groups
                $Sheet = $GP | ? { $_.Name -eq "sheet" } | % Value
                $Address = $GP | ? { $_.Name -eq "address" } | % Value
                if ($Format.Keys -contains $_.Name) {
                    $bk.Worksheets($Sheet).Range($Address).NumberFormatLocal = $Format[$_.Name]
                }
                $Value = $bk.Worksheets($Sheet).Range($Address).Cells | % Text
                
                New-Object PSObject -Property @{
                    Filename = $Book_Name;
                    FilePath = $Book;
                    Sheet    = $Sheet;
                    Address  = $Address;
                    Name     = $_.Name;
                    Value    = $Value
                } | Select-Object Filename, FilePath, Sheet, Address, Name, Value
            }
            $bk.close()
            $bk = $null
        }
    }
    end {}
}

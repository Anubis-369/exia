# Exia
PowershellのPSObjectをExcelのファイルに書き込を行ったり、Excelファイルの内容を読み込んでPSObjectを生成するためのモジュールです。ComObjectでExcel.Applicationを作成してExcelのファイルの読み込みや書き込みを行うので、ExcelがインストーるされたWindowsのシステムで使用してください。

このツールの特徴は、Excelのテーブル機能を利用することです。Excelファイルからデータを読み込むとき、Excelで作成したテーブルの名前を指定します。書き込みを行う時は、データをリスト形式で書き込みます。特定のセルに値を入力する際も、セルに付けた名前を指定して値を入力します。セルのアドレスを直接指定すると、スクリプトが複雑化しやすいこと、あとでスクリプトを読んだときに意味が分からなくなり易いという理由で、リストの名前やセルの名前を指定するようにしました。

# 使用方法
ExiaのモジュールをダウンロードしてImport-Moduleで読み込むと、以下のコマンドが使用できるようになる。

```
> Import-Module (Exiaのフォルダのパス)
> Get-Command -Module exia

CommandType     Name                                               Version    Source
-----------     ----                                               -------    ------
Function        Add-ExDescription                                  0.0        Exia
Function        Read-ExTableToPSO                                  0.0        Exia
Function        Read-ExXlsxTableToPSO                              0.0        Exia
Function        Read-ExXLSXToVal                                   0.0        Exia
Function        Write-ExGPOToTable                                 0.0        Exia
Function        Write-ExPSOToList                                  0.0        Exia
Function        Write-ExPSOToTable                                 0.0        Exia
Function        Write-ExPSOToXlsxTable                             0.0        Exia
Function        Write-ExValToXlsx                                  0.0        Exia
```

使い方はヘルプから参照できます。

```
> get-help Write-ExPSOToList

名前
    Write-ExPSOToList

概要
    PSObjectの内容をCOM経由で開いているExcelのファイルにList形式で書き込む。


構文
    Write-ExPSOToList [[-Sheet] <Object>] [[-Address] <String>] [-PSObject] <PSObject[]> [[-Title] <String>] [[-Label]
    <String>] [[-tmargin] <Int32>] [[-vmargin] <Int32>] [[-Members] <String[]>] [[-Header] <String[]>] [[-Format] <Hash
    table>] [<CommonParameters>]
・・・・
```


# いきなり使える系のコマンド
この二つのコマンドレットは、事前にComObjectを生成しなくても、Excelファイルから読み書きができる。ただし、内部でComobjectの生成と破棄を行っているのでオーバーヘッドが大きい。

- Read-ExXlsxTableToPSO
- Write-ExPSOToXlsxTable

Excelファイルからテーブルのデータを読み込む。
```
Read-ExXlsxTableToPSO -Book <File Path>-Table <Table Name>
```

Excelファイルからテーブルのデータを読み込む。File Pathのファイルの Worksheet Name のシートの B2 から、PSObjectの内容を書き込む。
```

Write-ExPSOToXlsxTable -File <File Path> -PSObject <PSObject> -Sheet <Worksheet Name> -Address B2
```
パイプラインを通して書き込みを行うこともできる。
```
<PSObject> | Write-ExPSOToXlsxTable -File <File Path> -Sheet <Worksheet Name> -Address B2
```

# Exia
PowershellのPSObjectをExcelのファイルに書き込を行ったり、Excelファイルの内容を読み込んでPSObjectを生成するためのモジュールです。ComObjectでExcel.Applicationを作成してExcelのファイルの読み込みや書き込みを行うので、ExcelがインストーるされたWindowsのシステムで使用してください。

このツールの特徴は、Excelのテーブル機能を利用することです。Excelファイルからデータを読み込むとき、Excelで作成したテーブルの名前を指定します。書き込みを行う時は、データをリスト形式で書き込みます。特定のセルに値を入力する際も、セルに付けた名前を指定して値を入力します。セルのアドレスを直接指定すると、スクリプトが複雑化しやすいこと、あとでスクリプトを読んだときに意味が分からなくなり易いという理由で、リストの名前やセルの名前を指定するようにしました。

# 使用方法
ExiaのモジュールをダウンロードしてImport-Moduleで読み込むと、以下のコマンドが使用できるようになります。

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


# いきなり使える系のコマンドレット
この二つのコマンドレットは、事前にComObjectを生成しなくても、Excelファイルから読み書きができます。Exiaの使い方をつかむのに、このコマンドレットを試してみてみるといい。ただし、内部でComobjectの生成と破棄を行っているのでオーバーヘッドが大きいです。

- Read-ExXlsxTableToPSO
- Write-ExPSOToXlsxTable

以下のようにコマンドレットを実行して、Excelファイルからテーブルのデータを読み込みます。
```
Read-ExXlsxTableToPSO -Book <File Path>-Table <Table Name>
```

以下のようにコマンドレットを実行して、Excelファイルからテーブルのデータを読み込みます。File Pathのファイルの Worksheet Name のシートの B2 から、PSObjectの内容を書き込む。
```
Write-ExPSOToXlsxTable -File <File Path> -PSObject <PSObject> -Sheet <Worksheet Name> -Address B2
```
パイプラインを通して書き込みを行うこともできます。
```
<PSObject> | Write-ExPSOToXlsxTable -File <File Path> -Sheet <Worksheet Name> -Address B2
```

# 基本的な使い方
## 1. ComObjectの生成
コマンドレットを使用する前に、以下のコマンドを実行して事前にComobjectやExcelファイルを読み込んだオブジェクトを生成しておきます。
```
$exl = New-Object -ComObject Excel.Application     # この操作で、$exl にExcel.Application のComObjectが生成される。
$bk = $exl.Workbooks.Open("Excelファイルのパス")　　# $bk に＄exl に読み込まれたExcelファイルのデータが入る。
$ws = $bk.Worksheets(1)                            # $bk の一番目のWorksheetのデータが入る。
```

## コマンドレットの実行

## オブジェクトの破棄
作業が終わったら、生成したオブジェクトの破棄を行います。これを行わないと、プロセスが起動し続けてメモリ上にデータが残ります。
```
$bk.Close()
$bk = $null
$exl.Quit()
$exl = $null
```
もし、この操作をする前にコンソールを閉じてしまったら、タスクマネージャからプロセスを落とします。

# Exia
PowershellのPSObjectをExcelのファイルに書き込むツール。

## 基本的な使い方
ExiaはExcelにComオブジェクトを経由して操作を行う。以下の操作で、ComオブジェクトでExcel.Application を作成し、Bookを開く。

New-Object -com　Excel.Application

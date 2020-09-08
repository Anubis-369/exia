$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module $here\..\Exia

Describe "Exia" {
    It "モジュールがロードできる" {
        Get-Module -Name Exia | Should Be $true
    }

    $Commands = @(
        "Add-ExDescription",
        "Read-ExTableToPSO",
        "Read-ExXlsxTableToPSO",
        "Read-ExXLSXToVal",
        "Write-ExGPOToTable",
        "Write-ExPSOToList",
        "Write-ExPSOToTable",
        "Write-ExPSOToXlsxTable",
        "Write-ExValToXlsx"
    )

    $Result = Get-Command -Module Exia | % Name

    $Commands | %{
        it ("{0} がロードされている" -f $_ )　{
            $Result -contains $_ | Should Be $true
        }
    }
}
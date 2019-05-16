. .\Shared.ps1
# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$public\$sut"

Describe "New-Foo" {

    $obj = New-MockObject -Type 'CrystalDecisions.CrystalReports.Engine.ReportDocument'
    # $obj.Title='sfdsf'
    $obj.FilePath='path/to/report/report.rpt'
    
    It "does something useful" {
        $true | Should -Be $false
    }

}

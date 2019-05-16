<#
.SYNOPSIS
    Saves the report.

.DESCRIPTION
    lorem ipsum

.PARAMETER ReportDocument
    A CrystalDecisions.CrystalReports.Engine.ReportDocument created by using Open-Report.

.PARAMETER Replace
    Replace the current file.

.EXAMPLE
    Open-Report $rpt | Out-Report

.INPUTS
    CrystalDecisions.CrystalReports.Engine.ReportDocument

.OUTPUTS
    CrystalDecisions.CrystalReports.Engine.ReportDocument

.NOTES

.LINK

#>
function Out-Report {

    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)]
        [CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument,

        [switch]$Replace
    )

    begin {}
    process {

        $FilePath = $reportDocument.FilePath

        if ( -not $Replace )
        {
            $stamp = '{0:yyyy.MM.dd-HH.mm.ss}' -f (get-date)
            $FilePath = $FilePath -replace '.rpt', ".$stamp.rpt"
        }
        
        Write-Debug "Savings as $FilePath"

        $reportDocument.SaveAs($FilePath)

		try {
		    $reportDocument.Close()
		}
		catch {
		    Write-Error $_.Exception
		}
		finally {
		    $reportDocument.Dispose()
        }
    
    } # /process
    end {}

}

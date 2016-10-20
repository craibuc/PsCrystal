<#
.SYNOPSIS
    Refresh the report and its data.

.DESCRIPTION
.NOTES
    This command may be unnecessary: Export-Report automatically generates report
.LINK
.EXAMPLE
.INPUTTYPE
.RETURNVALUE
.COMPONENT
.ROLE
.FUNCTIONALITY
.PARAMETER
#>

function Invoke-Report {

	[cmdletbinding()]
    param (
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument
    )

	begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}

    process {

        try {
            $reportDocument.Refresh
        }
        catch [Exception] {
            write-host $_.Exception
        }
        return $reportDocument
    }

    end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}


}
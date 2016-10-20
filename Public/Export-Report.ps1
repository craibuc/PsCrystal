<#
.SYNOPSIS
    Refresh the report and export it to the specified format

.DESCRIPTION
.NOTES
.LINK
.EXAMPLE
.INPUTTYPE
.RETURNVALUE
.COMPONENT
.ROLE
.FUNCTIONALITY
.PARAMETER
#>

function Export-Report {

	[cmdletbinding()]
    param (
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument,
        [CrystalDecisions.Shared.ExportFormatType]$format=[CrystalDecisions.Shared.ExportFormatType]::PortableDocFormat,
        [string]$path
    )

    begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}

    process {

        try {
            Write-Verbose "Exporting $path ..."

            #$reportDocument.ExportToDisk([CrystalDecisions.Shared.ExportFormatType]::PortableDocFormat, $path)
            $reportDocument.ExportToDisk($format, $path)

        }
        catch [Exception] {
            write-host $_.Exception
        }
        return $reportDocument

    }

    end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}

}
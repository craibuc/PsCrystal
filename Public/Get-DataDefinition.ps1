<#
.SYNOPSIS
    Extracts the tables, fields, linking, and commands from a RPT file.
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
function Get-DataDefinition {

    [cmdletbinding()]
    param (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelinebyPropertyName=$true,ValueFromPipeline=$true)]
        [Alias('FullName')]
        [string[]]$files
    )

    begin {
        Write-Debug "$($MyInvocation.MyCommand.Name)::Begin"

		# $assemblyPath='C:\Program Files (x86)\SAP BusinessObjects\SAP BusinessObjects Enterprise XI 4.0\win32_x86\dotnet\iPoint'
		$assemblyPath = $MyInvocation.MyCommand.Module.PrivateData.AssemblyPath
		Write-Debug "AP: $assemblyPath"

		Add-Type -Path "$assemblyPath\CrystalDecisions.CrystalReports.Engine.dll"

        $reportDocument = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    }

    process {
        Write-Debug "$($MyInvocation.MyCommand.Name)::Process"

        foreach ($file In $files){
        	$resolvedPath = Resolve-Path $file
            Write-Verbose "Opening $resolvedPath"
            try {
                $reportDocument.Load($resolvedPath)

				$report = [PsCustomObject]@{
					Title="$($reportDocument.SummaryInfo.ReportTitle)";
					Author="$($reportDocument.SummaryInfo.ReportAuthor)";
					Uri=$null;
					UncPath="$(Get-UncPath -Path $reportDocument.FilePath)"
					FilePath="$($reportDocument.FilePath)";
					FileInfo=$(Get-Item $reportDocument.FilePath);
					Queries=@();
					Links=@();
					Fields=@()
				}
				$Uri = New-Object System.Uri($report.UncPath)
				$report.Uri = $Uri.AbsoluteUri

				$report.Links += Get-Links $reportDocument.Database
				$report.Fields += Get-Fields $reportDocument.Database
				$report.Queries += Get-Queries $reportDocument.ReportClientDocument.DataDefController.Database

				foreach ($subreport In $reportDocument.Subreports)
				{
					$report.Links += Get-Links $subreport.Database $subreport.Name
					$report.Fields += Get-Fields $subreport.Database $subreport.Name
					#
					$subreportClientDocument = $reportDocument.ReportClientDocument.SubreportController.GetSubreport($subreport.Name)
					$report.Queries += Get-Queries $subreportClientDocument.DataDefController.Database $subreport.Name
				}

				# return PsCustomObject to pipeline
				$report
            }
            catch [CrystalDecisions.Shared.CrystalReportsException] {
                switch ($_.Exception.InnerException.ErrorCode) {
                    0x80004005 {
                        #The system cannot find the path specified [0x80004005]
                        # throw [System.Io.FileNotFoundException] ("File not found: {0}" -f $path)
                        Write-Error  ("File not found: {0}" -f $path)
                    }
                    0x8000020D {
                        # Unable to load report [0x8000020D]
                        # throw [System.FormatException] ("Invalid file: {0}" -f $path)
                        Write-Error  ("Invalid file: {0}" -f $path)
                    }
                    default {
                        throw [System.Exception] ($error)
                    }
                }
            }
            catch {
                Write-Error $_.Exception
            }
            finally {
            	if ($reportDocument) {
            		$reportDocument.Close()
            	}
            }
        }

    } # /process

    end { Write-Debug "$($MyInvocation.MyCommand.Name)::End" }

}

<#
.SYNOPSIS
    Extracts the table links from a RPT file.
#>
function Get-Links
{
    param(
        [CrystalDecisions.CrystalReports.Engine.Database]$database,
        [string]$subreport
    )

    $database.Links | Foreach {

        for ($i=0; $i -le  $_.SourceFields.Count-1; $i++) {

            [PsCustomObject]@{
                Subreport = $subreport;
                Table = $_.SourceTable.Location;
                TableAlias = if ($_.SourceTable.Location -ne $_.SourceTable.Name) { $_.SourceTable.Name };
                Field = $_.SourceFields[$i].Name;
                LinkType = $_.JoinType;
                LinkedTable = $_.DestinationTable.Location;
                LinkedTableAlias = if ($_.DestinationTable.Location -ne $_.DestinationTable.Name) { $_.DestinationTable.Name };
                LinkedField = $_.DestinationFields[$i].Name;
            }

        } # /for

    } # /foreach

} # /get-links

<#
.SYNOPSIS
    Extracts the tables and fields from a RPT file.
#>
function Get-Fields
{
    param(
        [CrystalDecisions.CrystalReports.Engine.Database]$database,
        [string]$subreport
    )

    $database.Tables.Fields | Where-Object {$_.UseCount -gt 0 } | Foreach {

        [PsCustomObject]@{
            Subreport = $subreport;
            Table = $_.TableName;
            Field = $_.Name;
        }

    } # /foreach

} # /Get-Fields

<#
.SYNOPSIS
    Extracts the queries from a RPT file.
#>
function Get-Queries
{
    param(
        [CrystalDecisions.ReportAppServer.DataDefModel.Database]$database,
        [string]$subreport
    )
    write-Verbose "Get-Queries: $subreport"

    $database.Tables | Foreach {

        if ( $_ -is [CrystalDecisions.ReportAppServer.DataDefModel.CommandTable]) {
            [PsCustomObject]@{
                Subreport = $subreport;
                Query = $_.CommandText
            }
        } # /if

    } # /foreach

} # /Get-Queries

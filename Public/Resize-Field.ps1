<#
.SYNOPSIS
    Resizes all fields in the details section and associated header field to the desired size.

.DESCRIPTION
    Resizes all fields in the details section and associated header field to the desired size.

.PARAMETER ReportDocument
    A CrystalDecisions.CrystalReports.Engine.ReportDocument created by using Open-Report.

.PARAMETER FieldWidth
    The desired width of each field in inches, defaults to 720 twips (0.5 inches).

.EXAMPLE
    Open-Report $rpt | Resize-Field

.INPUTS
    CrystalDecisions.CrystalReports.Engine.ReportDocument

.OUTPUTS
    CrystalDecisions.CrystalReports.Engine.ReportDocument

.NOTES
    One twip is 1/1440 inch or 1/567 cm.  Ergo, there are 1440 twips/inch and 567 twips/centimeter.

.LINK
* https://en.wikipedia.org/wiki/Twip

#>
function Resize-Field
{

	[cmdletbinding()]
    param (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
        [CrystalDecisions.CrystalReports.Engine.ReportDocument]$ReportDocument,

        [int]$FieldWidth = 720
    )

    begin
    { 
        Write-Debug "FieldWidth: $FieldWidth" 
        $TPI = 1440
        $TPC = 567
    }

    process
    {
        Write-Verbose "Processing $($ReportDocument.FilePath)"

        # why are there mulitple reportdefinitions?
        $ReportDefinition = $ReportDocument.ReportDefinition[0]

        $ReportDefinition.Areas.Sections | Where-Object { $_.Kind -eq 'Detail' } | ForEach-Object {

            # field counter
            $Counter = 0

            # calculate desired width
            $Width = $FieldWidth
            
            # process each ReportObject in the (Detail) Section
            $_.ReportObjects | ForEach-Object {

                $Left = $Counter * $Width
                Write-Debug "Left: $($Left/$TPI)in ($($Left/$TPC)cm)"
    
                # capture current field
                $Field = $_
                Write-Verbose "Before: $($Field.Name) Left: $($Field.Left/$TPI)in ($($Field.Left/$TPC)cm) Width: $($Field.Width/$TPI)in ($($Field.Width/$TPC)cm)"

                # find associated text field (header)
                $Header = $ReportDefinition.ReportObjects | Where-Object { $_.Kind -eq 'FieldHeadingObject' -and $_.FieldObjectName -eq $Field.Name} 

                # resize field and header
                $Field.Width = $Width
                $Header.Width = $Width

                # reposition field and header
                $Field.Left = $Left
                $Header.Left = $Left

                Write-Verbose "After: $($Field.Name) Left: $($Field.Left/$TPI)in ($($Field.Left/$TPC)cm) Width: $($Field.Width/$TPI)in ($($Field.Width/$TPC)cm)"

                # increment counter
                $Counter += 1

            } # / For-Each ReportObjects

        } # / For-Each Section

        # return the modified report to the pipeine
        Write-Output $ReportDocument

    } # /process

    end {}

} # / Resize-Field

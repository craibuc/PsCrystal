Import-Module PsCrystal -Force

$folder = "C:\Documents and Settings\All Users\Reports"

foreach($file in Get-ChildItem $folder) {

	try {
	
		$rptPath = $folder + "\" + $file
		
		# open report
		$rpt = Open-Report $rptPath | Out-Null
		
		# other stuff
		$rpt.SummaryInfo.ReportComments += $("`r`nModified: {0}" -f (Get-Date))
		
		# set DB credentials
		
		# set parameter values
		
		# generated PDF
		
		#save report
		Save-Report $rpt -path $rptPath

	}
	catch {
		#handle exception
	}
	finally {
		if ($rpt) {
			# close report to release resources
			Close-Report $rpt
		}
	}

}
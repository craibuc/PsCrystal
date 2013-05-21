
[cmdletbinding()]
param ()

begin {
    Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"
    import-module PsCrystal -force
	[reflection.assembly]::LoadWithPartialName('CrystalDecisions.ReportAppServer.DataDefModel') | Out-Null

	}

process {

    $rpt = $null
	
    try {

        $rptFile = Convert-Path ".\SQLCommand.12.rpt"
        
        # TODO: test for presence of file
        
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing $rptFile"
        
        $rpt = Open-Report -path $rptFile -verbose #| Out-null

		# TODO: prompt for credentials based on values in report

#		$server = Read-Host -Prompt "Server"
#		$account = Read-Host -Prompt "Account"
#		$securePassword = Read-Host -AsSecureString -Prompt "Password"
#		$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
#		$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

#        $rpt = Set-Credentials $rpt -svr $server -acct $account -pwd $password

		$RAS = $rpt.ReportAppServer
		$reportClient = $rpt.ReportClientDocument
		
		#http://devlibrary.businessobjects.com/businessobjectsxi/en/en/RAS_SDK/rassdk_net_doc/doc/rassdk_net_doc/html/CrystalDecisions.ReportAppServer.DataDefModel~CrystalDecisions.ReportAppServer.DataDefModel.CommandTableClass_members.html

        # use the RowsetController to get the SQL query
        # Note: If a report has parameters they must be supplied before getting the SQL query.
        $database = $reportClient.DataDefController.Database
		Write-Host "SQL: " $database.Tables[0].CommandText
                  
    }
    catch {
        write-host $_.Exception.Message
    }
    finally {
		# clear password from memory
#		[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
		# close report
        if ($rpt) { Close-Report $rpt -verbose }
    }
}

end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}
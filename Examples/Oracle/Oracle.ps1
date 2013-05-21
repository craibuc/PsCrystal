
[cmdletbinding()]
param ()

begin {
    Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"
    import-module PsCrystal -force
    }

process {

    $rpt = $null
	
	try {

        $rptPath = Convert-Path ".\Oracle.12.rpt"
        
        # TODO: test for presence of file
        
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing $rptFile"
        
        $rpt = Open-Report -path $rptFile -verbose #| Out-null
		
		# TODO: prompt for credentials based on values in report

		$server = Read-Host -Prompt "Server"
		$account = Read-Host -Prompt "Account"
		$securePassword = Read-Host -AsSecureString -Prompt "Password"
		$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
		$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

        $rpt = Set-Credentials $rpt -svr $server -acct $account -pwd $password        
        $rpt = Export-Report $rpt -path ($rptPath -replace ".rpt", ("." + (Get-Date -f "yyyyMMdd-HHmmss") + ".pdf"))  -verbose

    }
    catch {
        write-host $_.Exception.Message
    }
    finally {
		# clear password from memory
		[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
		# close report
        if ($rpt) { Close-Report $rpt -verbose }
    }
}

end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}
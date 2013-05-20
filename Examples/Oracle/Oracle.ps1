
[cmdletbinding()]
param ()

begin {
    Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"
    import-module PsCrystal -force
    }

process {

    $rpt = $null
	
	$server = Read-Host -Prompt "Server"
	$user = Read-Host -Prompt "Account"
	$securePassword = Read-Host -AsSecureString -Prompt "Password"
	
    try {

        $rptFile = Convert-Path ".\Oracle.12.rpt"
        
        # TODO: test for presence of file
        
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing $rptFile"
        
        $rpt = Open-Report -path $rptFile -verbose #| Out-null
        
		$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
			[Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword))
			
        $rpt = Set-Credentials $rpt -svr $server -acct $account -pwd $password
        
        $rpt = Export-Report $rpt -path ($rptFile -replace '.rpt', '.pdf')  -verbose
                        
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
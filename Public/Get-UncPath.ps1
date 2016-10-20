<#
.SYNOPSIS
    Convert a mapped drive (e.g. D:\) to its UNC representation (e.g. \\server\path\to\folder)
.DESCRIPTION
.NOTES
Source:	http://superuser.com/a/29972/10768
.LINK
.EXAMPLE
.INPUTTYPE
.RETURNVALUE
.COMPONENT
.ROLE
.FUNCTIONALITY
.PARAMETER
#>
function Get-UncPath {

    [cmdletbinding()]
    param (
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelinebyPropertyName=$true,ValueFromPipeline=$true)]
        [Alias('FullName')]
        [string[]]$path
    )
    PROCESS {
    	foreach ($p in $path) {
    		Write-Verbose $p

    		# get the FileInfo object specified by $path
    		$item = Get-Item $p

    		# folder or file
    		switch ($item.PsIsContainer) {
    			# folder
    			$true	{
    				$currentDrive = Split-Path -Qualifier $item.FullName
    				Write-Debug $currentDrive
    			}
    			# file
    			$false	{
    				$currentDrive = Split-Path -Qualifier $item.DirectoryName
    				Write-Debug $currentDrive
    			}
    		}

			$logicalDisk = Gwmi Win32_LogicalDisk -filter "DriveType = 4 AND DeviceID = '$currentDrive'"
			Write-Debug $logicalDisk.ProviderName

			$uncPath = $path.Replace($currentDrive, $logicalDisk.ProviderName)
			$uncPath

    	} # /foreach
    } # /PROCESS
}

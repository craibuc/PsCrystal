Import-Module PsCrystal -Force

# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
# $sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
# . "$here\$sut"

Describe "Close-Report.ps1" {

	$moduleDirectory = Split-Path -Path (Get-Module -ListAvailable PsCrystal).Path

	Context "File supplied via pipeline" {

		$file = "$moduleDirectory\Examples\SQLCommand\SQLCommand.12.rpt"

		It -skip "Closes the report, returning null" {
			$actual = Get-ChildItem $file | Open-Report | Close-Report -Verbose
			$actual | Should BeNullOrEmpty
		}

	}

}

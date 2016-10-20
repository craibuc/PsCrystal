Import-Module PsCrystal -Force
# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
# $sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
# . "$here\$sut"

Describe "Get-DataDefinition" {

	#$moduleDirectory = Split-Path -Path (Get-Module -ListAvailable PsCrystal).Path
	$artifactDirectory = "~\Desktop\REPORTS"

	Context "Single file supplied via positional parameter" {

		$file = "$artifactDirectory\Oracle\Oracle.Linked.12.rpt"

		It "Creates a PsCustomObject" {
			$actual = Get-DataDefinition $file -Verbose
			$actual.GetType() | Should Be System.Management.Automation.PsCustomObject
		}

		It "Creates a PsCustomObject with contains an array of links" {
			$actual = Get-DataDefinition $file -Verbose
			$actual.Links.GetType() | Should Be System.Object[]
			$actual.Links.Count | Should BeGreaterThan 0
		}

		It "Creates a PsCustomObject with contains an array of fields" {
			$actual = Get-DataDefinition $file -Verbose
			$actual.Fields.GetType() | Should Be System.Object[]
			$actual.Fields.Count | Should BeGreaterThan 0
		}

		It "Creates a PsCustomObject with contains an array of queries" {
			$actual = Get-DataDefinition $file -Verbose
			$actual.Queries.GetType() | Should Be System.Object[]
			$actual.Queries.Count | Should BeGreaterThan 0
		}

	}

}

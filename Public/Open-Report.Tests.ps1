Import-Module PsCrystal -Force

# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
# $sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
# . "$here\$sut"

Describe "Open-Report" {

	$moduleDirectory = Split-Path -Path (Get-Module -ListAvailable PsCrystal).Path

	Context "Single file supplied via positional parameter" {

		$file = "$moduleDirectory\Examples\SQLCommand\SQLCommand.12.rpt"

		It -skip "Opens the report, returning a CrystalDecisions.CrystalReports.Engine.ReportDocument" {

			$actual = Open-Report $file -Verbose
			$actual.GetType() | Should Be CrystalDecisions.CrystalReports.Engine.ReportDocument
		}

	}

	Context "Single file supplied via pipeline" {

		$file = "$moduleDirectory\Examples\SQLCommand\SQLCommand.12.rpt"

		It -skip "Opens the report, returning a CrystalDecisions.CrystalReports.Engine.ReportDocument" {
			$actual = Get-ChildItem $file | Open-Report -Verbose
			$actual.GetType() | Should Be CrystalDecisions.CrystalReports.Engine.ReportDocument
		}

	}

	Context "Multiple files supplied via pipeline" {

		$directory = "$moduleDirectory\Examples\SQLCommand"

		It -skip "Opens the report, returning a CrystalDecisions.CrystalReports.Engine.ReportDocument" {
			$actual = @(Get-ChildItem $directory *.rpt | Open-Report -Verbose)
			# $actual | Should Be CrystalDecisions.CrystalReports.Engine.ReportDocument
			# Write-Host $actual.GetType().Name
			$actual.GetType() | Should Be System.Object[]
		}

	}

	Context "Invalid path is supplied" {

		$file = "$moduleDirectory\Foo\Bar.rpt"

		It -skip "Generates a non-terminating error" {
			{ Open-Report $file -Verbose} | Should Throw
		}

	}

	Context "Invalid path is encountered in the midst of valid paths" {

		$files=@()
		$files += "$moduleDirectory\Examples\SQLCommand\SQLCommand.12.rpt"
		$files += "$moduleDirectory\Foo\Bar.rpt"
		$files += "$moduleDirectory\Examples\SQLCommand\SQLCommand.Parameters.12.rpt"

		It -skip "Generates a non-terminating error, but continues processing" {
			$actual = @( $files | Open-Report -Verbose )
			$actual.Count | Should Be 2
		}

	}

}

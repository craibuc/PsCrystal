$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Get-UncPath" {

	Context "Given a mapped drive" {

		$folder = "N:\IT\Private\Clarity\EdgeReporting"
		$expected = "\\enterprise.stanfordmed.org\depts\IT\Private\Clarity\EdgeReporting"

	    It "Returns a UNC path" {
	    	$actual = Get-UncPath -Path $folder -Verbose
	        $actual | Should Be $expected
	    }
	}

	Context "Given a file on mapped drive" {

		$file = "N:\IT\Private\Clarity\EdgeReporting\BOE Migration - EXAMPLE.xlsx"
		$expected = "\\enterprise.stanfordmed.org\depts\IT\Private\Clarity\EdgeReporting\BOE Migration - EXAMPLE.xlsx"

	    It "Returns a UNC path" {
	    	$actual = Get-UncPath -Path $file -Verbose
	        $actual | Should Be $expected
	    }
	}

	Context "Given items supplied via a pipeline" {

		$items = @()
		$items += "N:\IT\Private\Clarity\EdgeReporting"
		$items += "N:\IT\Private\Clarity\EdgeReporting\BOE Migration - EXAMPLE.xlsx"

		$expected = @()
		$expected += "\\enterprise.stanfordmed.org\depts\IT\Private\Clarity\EdgeReporting"
		$expected += "\\enterprise.stanfordmed.org\depts\IT\Private\Clarity\EdgeReporting\BOE Migration - EXAMPLE.xlsx"

	    It "Returns a UNC path" {
	    	$actual = $items | Get-UncPath -Verbose
	        $actual | Should Be $expected
	    }
	}

}

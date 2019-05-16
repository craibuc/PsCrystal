$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "ConvertTo-PlainText" {

	$expected = 'pa55w0rd'
	$securePassword = ConvertTo-SecureString -String $expected -AsPlainText -Force

    It "Convert a secure string to a plain-text string" {
    	$actual = ConvertTo-PlainText $securePassword -Verbose
        $actual | Should Be $expected
    }
}

<#
.SUMMARY
Convert an encrypted string ([System.Security.SecureString]) to its plain-text equivalent ([System.String])

.DESCRIPTION
Convert an encrypted string ([System.Security.SecureString]) to its plain-text equivalent ([System.String])

.PARAMETER $SecureString
The encrypted value to be converted.

.EXAMPLE
PS > $password = ConvertTo-SecureString 'pa55w0rd' -AsPlainText -Force
PS > ConvertTo-PlainText $password
'pa55word'

.INPUTTYPE
[System.Security.SecureString]

.RETURNVALUE
[System.String]
#>
function ConvertTo-PlainText {

	[CmdletBinding()]
	param(
		[Parameter(Position = 0, Mandatory = $true)]
		[SecureString]$SecureString
	)

	$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
	[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

}

# import-module CrystalReports\Engine -force

Add-Type -TypeDefinition @"
   [System.Flags] 
   public enum FlagsEnum {
      None = 0,
      SummaryInfo = 1,
      ReportOptions = 2,
      ParameterFields = 4,
      All = SummaryInfo | ReportOptions | ParameterFields
   }
"@

function Get-Definition {

	[cmdletbinding()]
    param ( 
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument,
        [FlagsEnum]$options
    )
    
	begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}

	process {
	
	    $def = @{}
	    
	    try {
	    
	        if ( $options -band [FlagsEnum]::SummaryInfo ) { $def += Get-SummaryInfo $rpt } 
	        if ( $options -band [FlagsEnum]::ReportOptions ) { $def += Get-ReportOptions $rpt } 
	        if ( $options -band [FlagsEnum]::ParameterFields ) { $def += Get-ParameterFields $rpt } 

	    }
	    catch [Exception] {
	        write-host $_.Exception
	    }
        return $def

	}
	
	end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}

	<#
	.SYNOPSIS
	    Crystal Report represented as a PowerShell Hashtable
	.PARAMETER reportDocument
	    the report docuemnt
	.PARAMETER options
	    the objects to be converted
	.OUTPUTS
		Hastable
	#>
}

function Get-SummaryInfo {

	[cmdletbinding()]
    param ( 
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument
    )

	begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}
	
	process {
	
	    $si = @{"SummaryInfo"=@{
	            "ReportAuthor"=$reportDocument.SummaryInfo.ReportAuthor;
	            "ReportComments"=$reportDocument.SummaryInfo.ReportComments
	            "KeywordsInReport"=$reportDocument.SummaryInfo.KeywordsInReport;
	            "ReportSubject"=$reportDocument.SummaryInfo.ReportSubject;
	            "ReportTitle"=$reportDocument.SummaryInfo.ReportTitle
	            }
	        }  
	               
	    return $si

    }
	
	end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}
	
	<#
	.SYNOPSIS
	    SummaryInfo as a PowerShell Hashtable
	.PARAMETER reportDocument
	    the report docuemnt
	.OUTPUTS
		Hastable
	#>

}

function Get-ReportOptions {

	[cmdletbinding()]
    param ( 
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument
    )

	begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}

	process {
	
	    $ro = @{"ReportOptions"=@{
	            "EnableSaveDataWithReport"=$reportDocument.ReportOptions.EnableSaveDataWithReport;
	            "EnableSavePreviewPicture"=$reportDocument.ReportOptions.EnableSavePreviewPicture
	            }
	        } 
	                   
	    return $ro
   
   	}
	
	end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}
	
	<#
	.SYNOPSIS
	    ReportOptions as a PowerShell Hashtable
	.PARAMETER reportDocument
	    the report docuemnt
	.OUTPUTS
		Hastable
	#>
	
}

function Get-ParameterFields {

	[cmdletbinding()]
    param ( 
        [Parameter(ValueFromPipeline=$true,Position=0,Mandatory=$true)][CrystalDecisions.CrystalReports.Engine.ReportDocument]$reportDocument
    )

	begin {Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin"}

	process {
	
	    try {
	    
	        $PFs = @{}
	        
	        foreach ($parameterField in $reportDocument.dataDefinition.parameterFields) {
	            
	#            write-host $parameterField.Name

	#            write-host $("EditMask: {0}" -f $parameterField.EditMask)
	#            write-host $("EnableAllowEditingDefaultValue: {0}" -f $parameterField.EnableAllowEditingDefaultValue)
	#            write-host $("EnableAllowMultipleValue: {0}" -f $parameterField.EnableAllowMultipleValue)
	#            write-host $("EnableNullValue: {0}" -f $parameterField.EnableNullValue)
	#            write-host $("HasCurrentValue: {0}" -f $parameterField.HasCurrentValue)
	#            write-host $("Kind: {0}" -f $parameterField.Kind)
	#            write-host $("PromptingType: {0}" -f $parameterField.PromptingType) 
	#            write-host $("ParameterFieldName: {0}" -f $parameterField.ParameterFieldName)
	#            write-host $("ParameterFieldUsage: {0}" -f $parameterField.ParameterFieldUsage.toString())
	#            write-host $("ParameterType: {0}" -f $parameterField.ParameterType.toString())
	#            write-host $("ParameterValueKind: {0}" -f $parameterField.ParameterValueKind)
	#            write-host $("ParameterValueType: {0}" -f $parameterField.ParameterValueType)
	#            write-host $("PromptText: {0}" -f $parameterField.PromptText)
	#            write-host $("PromptingType: {0}" -f $parameterField.PromptingType) 
	#            write-host $("ReportParameterType: {0}" -f $parameterField.ReportParameterType)            
	#            write-host $("ValueType: {0}" -f $parameterField.ValueType)
	            
	            #write-host "CVs:" -nonewline #$parameterField.CurrentValues.count
	            #foreach ($parameterValue in $parameterField.CurrentValues) {
	            #    write-host $parameterValue.value "," -nonewline
	            #}
	            
	#            switch ($parameterField.ParameterValueKind) {
	#                "BooleanParameter" {write-host "It's a boolean"}
	#                "CurrencyParameter" {write-host "It's a currency"}
	#                "DateParameter" {write-host "It's a date"}
	#                "DateTimeParameter" {write-host "It's a date/time"}
	#                "StringParameter" {write-host "It's a string"}
	#                "NumberParameter" {write-host "It's a number"}
	#                "TimeParameter" {write-host "It's a time"}
	#                default {write-host "something else"}
	#            }
	            
	            $DVs = @()
	            foreach ($parameterValue in $parameterField.DefaultValues) {
	                $DVs += $parameterValue.value
	            }
	             
	            $PF = @{
	                $parameterField.Name=@{
	                    "DiscreteOrRangeKind"=$parameterField.DiscreteOrRangeKind.toString();
	                    "EnableAllowMultipleValue"=$parameterField.EnableAllowMultipleValue;
	                    "DefaultValues"=($DVs -join ',');
	                    "CurrentValues"=@();
	                    "ParameterValueKind"=$parameterField.ParameterValueKind;
	                    "PromptText"=$parameterField.PromptText
	                }
	            }
	                  
	            $PFs +=$PF 
	                      
	        } #foreach
	        
	        
	    }
	    catch [Exception] {
	        write-host $_.Exception
	    }
        return $PFs
    }
	
	end {Write-Verbose "$($MyInvocation.MyCommand.Name)::End"}
	
	<#
	.SYNOPSIS
	    ParameterFields as a PowerShell Hashtable
	.PARAMETER reportDocument
	    the report docuemnt
	.OUTPUTS
		Hastable
	#>

}

Export-ModuleMember Get-Definition
#Set-Alias rd Get-Definition
#Export-ModuleMember -Alias rd

Export-ModuleMember Get-SummaryInfo
#Set-Alias si Get-SummaryInfo
#Export-ModuleMember -Alias si

Export-ModuleMember Get-ReportOptions
#Set-Alias ro Get-ReportOptions
#Export-ModuleMember -Alias ro

Export-ModuleMember Get-ParameterFields
#Set-Alias pf Get-ParameterFields
#Export-ModuleMember -Alias pf

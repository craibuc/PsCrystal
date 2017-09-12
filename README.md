# PsCrystal

PowerShell modules to help automate Crystal Reports tasks.

# Usage

## Open-Report

```powershell
# open the report
PS > Open-Report \path\to\report.rpt
```

## Close-Report

```powershell
# open the report, then close it
PS > Open-Report \path\to\report.rpt | Close-Report
```
# Examples

## Extract the unique list of tables in a report
```powershell
PS> Get-ChildItem '\path\to\reports' *.rpt -Recurse | Open-Report -Verbose | % {

    $_.Database.Tables | Select-Object location -Expand location
    $_.Subreports | % { $_.Database.Tables | Select-Object location -Expand location }

} | Select-Object location -Unique | Sort-Object Location
```

## Extract SQL from the reports' command objects

```powershell
# extract the SQL queries from the RPT files in the specified directory
PS> Get-ChildItem '\path\to\reports' *.rpt | Open-Report -Verbose | % {
    # save for later
    $report = $_

    $_.ReportClientDocument.DataDefController.Database.Tables | 
        Where-Object { $_.ClassName -Eq 'CrystalReports.CommandTable' } |
        Select-Object @{Name='Title';Expression={$report.SummaryInfo.Title}}, @{Name='FilePath';Expression={$report.FilePath}}, @{Name='Subreport';Expression={$null}}, @{Name='CommandAlias';Expression={$_.Alias}}, @{Name='Query'; Expression={$_.CommandText}}

    $_.Subreports | % { 
            # save for later
            $subrepot = $_
            $_.ReportClientDocument.DataDefController.Database.Tables
        } |
        Where-Object { $_.ClassName -Eq 'CrystalReports.CommandTable' } |
        Select-Object @{Name='Title';Expression={$report.SummaryInfo.Title}}, @{Name='FilePath';Expression={$report.FilePath}}, @{Name='Subreport'; Expression={$subreport.Name}}, @{Name='CommandAlias';Expression={$_.Alias}}, @{Name='Query'; Expression={$_.CommandText}}
}
```

## Extract database connection information

```powershell
Get-ChildItem '\path\to\reports' *.rpt -Recurse | Open-Report -Verbose | % {

    $report = $_

    $_.Database.Tables | % {
        [PsCustomObject]@{
            FilePath=$report.FilePath
            FileName=(Split-Path $report.FilePath -Leaf)
            Subreport=$null
            TableName=$_.Name
            ServerName=$_.LogOnInfo.ConnectionInfo.ServerName
            UserID=$_.LogOnInfo.ConnectionInfo.UserID
        }
    }

    $_.Subreports | % { 
        
        $subreport = $_

        $_.Database.Tables | % {
            [PsCustomObject]@{
                FilePath=$report.FilePath
                FileName=(Split-Path $report.FilePath -Leaf)
                Subreport=$subreport.Name
                TableName=$_.Name
                ServerName=$_.LogOnInfo.ConnectionInfo.ServerName
                UserID=$_.LogOnInfo.ConnectionInfo.UserID
            }
        }
        
    }

} | Sort-Object FileName,TableName | Select-Object FileName,Subreport,TableName,ServerName,UserID | Format-Table
```


## Extract the tables and field from the reports

```powershell
# extract the fields from the RPT files in the specified directory and save them in a single CSV file
Get-ChildItem 'path\to\reports' *.rpt -Recurse | % {
    $report = $_

    # TODO

    $_.Subreports | % {
        $subreport = $_
	
	# TODO

    } # /Subreports
    
} | ConvertTo-Csv -NoTypeInformation | Out-File ~\Desktop\reports.fields.csv
```

## Test reports to determine if their tables or command-objects' queries match a pattern

```powershell

$patterns = @()
$patterns += [PsCustomObject]@{Title='Custom View (V_*)';Pattern='\bV_'}
$patterns += [PsCustomObject]@{Title='Procedure (SP_*)';Pattern='\bSP_'}
Write-Host "Patterns: $patterns"

Get-ChildItem 'path\to\reports' *.rpt -Recurse | % {

    Write-Host $_.FullName
    $report = $_
    
    Open-Report $_.FullName | % {

        $_.ReportClientDocument.Database.Tables | % {

            Write-Host "`t$($_.ClassName)/$($_.Name)"

            $text = if ( $_.ClassName -Eq 'CrystalReports.CommandTable' ) {  $_.CommandText } else { $_.Name }

            $matches = $patterns | % { if ($text -match $_.Pattern ) { $_.Title } }

            if ( $matches ) { 
                Write-Host "`t`tMatches: $( $matches -Join ';' )"
                [PsCustomObject]@{Title=$report.Name;Subreport=$null;ClassName=$_.ClassName;Name=$_.Name;Matches=$( $matches -Join ';' )} 
            }


        } # /Tables

        if ( $_.Subreports ) {

            $_.Subreports | % {

                $subreport = $_

                Write-Host ("`t`tsubreport: {0}" -f $subreport.Name)

                $_.Database.Tables | % {

                    Write-Host "`t`t`t$($_.ClassName)/$($_.Name)"

                    $text = if ( $_.ClassName -Eq 'CrystalReports.CommandTable' ) {  $_.CommandText } else { $_.Name }
                    $matches = $patterns | % { if ($text -match $_.Pattern ) { $_.Title } }

                    if ( $matches ) { 
                        Write-Host "`t`t`t`tMatches: $( $matches -Join ';' )"
                        [PsCustomObject]@{Title=$report.Name;Subreport=$subreport.Name;ClassName=$_.ClassName;Name=$_.Name;Matches=$( $matches -Join ';' )} 
                    }

                } # /Tables

            } # /Subreports

        }

    }

} | Format-Table
```

## Extract the table linking from the reports

```powershell
# extract the linking from the RPT files in the specified directory and save them in a single CSV file
PS ~\Desktop\Reports> Get-ChildItem . *.rpt -Recurse | Get-DataDefinition | 
		Select title, filepath -ExpandProperty links | 
		ConvertTo-Csv -NoTypeInformation | 
		Out-File .\reports.links.csv
```
# Installation

- Install PowerShell
- Set your execution policy to 'RemoteSigned':
```powershell
# test its current value
PS > Get-ExecutionPolicy
None

# set to 'RemoteSigned'; requires administrative rights
PS > Set-ExecutionPolicy 'RemoteSigned'
```

## Git
Assuming that you have Git installed:

```powershell
# change to your profile's PowerShell directory
PS> cd ~\Documents\WindowsPowerShell

# create a Modules folder
PS> mkdir Modules

# change to this directory
PS> cd Modules

# clone the project to Modules directory (it will create the PsCrystal folder and contents)
PS> git clone git@github.com:craibuc/PsCrystal.git

# import the module into current PowerShell session
PS> Import-Module PsCrystal

```

## Manually

```powershell
TODO
```

# Personnel

## Author

 - [Craig Buchanan](https://github.com/craibuc)

## Contributors

 - [Steve Romanow](https://github.com/slestak)

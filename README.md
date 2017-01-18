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
~~~~powershell
PS> Get-ChildItem '\path\to\reports' *.rpt -Recurse | Open-Report -Verbose | % {

    $_.Database.Tables | Select-Object location -Expand location
    $_.Subreports | % { $_.Database.Tables | Select-Object location -Expand location }

} | Select-Object location -Unique | Sort-Object Location
~~~

## Extract SQL from the reports' command object

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
## Extract the tables and field from the reports

```powershell
# extract the fields from the RPT files in the specified directory and save them in a single CSV file
PS ~\Desktop\Reports> Get-ChildItem . *.rpt -Recurse | Get-DataDefinition | 
		Select title, filepath -ExpandProperty fields | 
		ConvertTo-Csv -NoTypeInformation | 
		Out-File .\reports.fields.csv
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

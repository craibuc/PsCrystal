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

## Get-DataDefinition

### Extract SQL from the reports' command object

```powershell
# extract the SQL queries from the RPT files in the specified directory and save them in a SQL file named like the RPT file
PS ~\Desktop\Reports> Get-ChildItem . *.rpt -Recurse | Get-DataDefinition | 
		Foreach { $fp= $_.FilePath; $_.Queries | 
			Foreach { $_.Query | 
				Out-File ($fp -Replace '.rpt', '.sql')
			}
		}
```
### Extract the tables and field from the reports

```powershell
# extract the fields from the RPT files in the specified directory and save them in a single CSV file
PS ~\Desktop\Reports> Get-ChildItem . *.rpt -Recurse | Get-DataDefinition | 
		Select title, filepath -ExpandProperty fields | 
		ConvertTo-Csv -NoTypeInformation | 
		Out-File .\reports.fields.csv
```
### Extract the table linking from the reports

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

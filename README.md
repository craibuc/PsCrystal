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

# Personnel

## Author

 - [Craig Buchanan](https://github.com/craibuc)

## Contributors

 - [Steve Romanow](https://github.com/slestak)

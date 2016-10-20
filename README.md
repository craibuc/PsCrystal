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

```powershell
# extract the SQL queries from the RPT files in the specified directory and save them in a SQL file named like the RPT file
PS ~\Desktop> Get-ChildItem .\Reports *.rpt -Recurse | Get-DataDefinition | 
		Foreach { $fp= $_.FilePath; $_.Queries | 
			Foreach { $_.Query | 
				Out-File ($fp -Replace '.rpt', '.sql')
			}
		}
```

```powershell
# extract the linking from the RPT files in the specified directory and save them in a single CSV file
PS ~\Desktop> Get-ChildItem .\reports *.rpt -Recurse | Get-DataDefinition | Select title, filepath -ExpandProperty links | ConvertTo-Csv -NoTypeInformation | Out-File ~\Desktop\reports.csv
```

# Roadmap

 - Examples
 - Documentation
 - Unit tests (using Pester)
 - Add to PsGet

# Personnel

## Author

 - [Craig Buchanan](https://github.com/craibuc)

## Contributors

 - [Steve Romanow](https://github.com/slestak)

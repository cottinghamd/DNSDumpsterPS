<#	
	.NOTES
	===========================================================================
	 Created by:   	David Cottingham  	
	===========================================================================
	.DESCRIPTION
		This is an unofficial Powershell script to perform lookups against DNSSumpster.com and collate results.
		This script currently requires Microsoft Excel to be installed on the machine that the script is run on
		Usage: Import a domain list CSV, with the header of 'Sites'
		The script will then output results to the scripts current working directory. Note: You must have permissions to write to the script working directory.
#>


$WorkingDir = Get-Location

If (Test-Path -path "$WorkingDir\domainlist.csv" -ErrorAction SilentlyContinue)
{
	$SitestoScan = Import-CSV -Path "$WorkingDir\domainlist.csv"
}
else
{
	$SiteCSV = Read-Host "Please type the full path to the CSV containing the sites you wish to scan. e.g. C:\domainlist.csv (Note: This CSV must have a header row called Sites)"
	$SitestoScan = Import-CSV -Path $SiteCSV
}

If ($SitestoScan.Sites -eq $null)
{
	Write-Output "The CSV is not valid or has been incorrectly formatted. Please ensure the CSV has a header row of Sites and each site you want to scan on a new line in the file"
	Pause
	break
}

$SitestoScan = Import-CSV -Path "$WorkingDir\domainlist.csv"
$myObject = New-Object System.Object
$SiteNotFound = @()

#Function taken from https://podlisk.wordpress.com/2011/11/20/import-excel-spreadsheet-into-powershell/, Ideally in the future I would like to remove the excel dependency, under a time crunch
function Import-Excel
{
	param (
		[string]$FileName,
		[string]$WorksheetName,
		[bool]$DisplayProgress = $true
	)
	
	if ($FileName -eq "")
	{
		throw "Please provide path to the Excel file"
		Exit
	}
	
	if (-not (Test-Path $FileName))
	{
		throw "Path '$FileName' does not exist."
		exit
	}
	
	$FileName = Resolve-Path $FileName
	$excel = New-Object -com "Excel.Application"
	$excel.Visible = $false
	$workbook = $excel.workbooks.open($FileName)
	
	if (-not $WorksheetName)
	{
		Write-Warning "Defaulting to the first worksheet in workbook."
		$sheet = $workbook.ActiveSheet
	}
	else
	{
		$sheet = $workbook.Sheets.Item($WorksheetName)
	}
	
	if (-not $sheet)
	{
		throw "Unable to open worksheet $WorksheetName"
		exit
	}
	
	$sheetName = $sheet.Name
	$columns = $sheet.UsedRange.Columns.Count
	$lines = $sheet.UsedRange.Rows.Count
	
	Write-Warning "Worksheet $sheetName contains $columns columns and $lines lines of data"
	
	$fields = @()
	
	for ($column = 1; $column -le $columns; $column++)
	{
		$fieldName = $sheet.Cells.Item.Invoke(1, $column).Value2
		if ($fieldName -eq $null)
		{
			$fieldName = "Column" + $column.ToString()
		}
		$fields += $fieldName
	}
	
	$line = 2
	
	
	for ($line = 2; $line -le $lines; $line++)
	{
		$values = New-Object object[] $columns
		for ($column = 1; $column -le $columns; $column++)
		{
			$values[$column - 1] = $sheet.Cells.Item.Invoke($line, $column).Value2
		}
		
		$row = New-Object psobject
		$fields | foreach-object -begin { $i = 0 } -process {
			$row | Add-Member -MemberType noteproperty -Name $fields[$i] -Value $values[$i]; $i++
		}
		$row
		$percents = [math]::round((($line/$lines) * 100), 0)
		if ($DisplayProgress)
		{
			Write-Progress -Activity:"Importing from Excel file $FileName" -Status:"Imported $line of total $lines lines ($percents%)" -PercentComplete:$percents
		}
	}
	$workbook.Close()
	$excel.Quit()
}

#Setup DNS Dumpster Page Load and get CSRF Token
try { $login = Invoke-WebRequest -Uri 'https://dnsdumpster.com' -SessionVariable session }
catch { $PageError = $_.Exception }

If ($PageAuthError -ne $null)
{
	Write-Output "Error Thrown, Aborting: $($PageError.Message)"
	break
}
else
{
	write-host "Successful Web Connection"
}

#Scan all sites in array
$SitestoScan | ForEach-Object{

	#Setup CSRF Tokens, Form Body and Headers to Pass
	$csrf = $login.InputFields[0].value
	$login = @{ csrfmiddlewaretoken = $csrf; targetip = $_.Sites }
	$header = @{ Referer = 'https://dnsdumpster.com/'; }
	
	#Send site scan request to DNSDumpster
	try { $ScanResults = Invoke-WebRequest -Uri 'https://dnsdumpster.com' -Body $login -Method Post -WebSession $session -ContentType 'application/x-www-form-urlencoded' -Headers $header }
	catch { $PageError2 = $_.Exception }
	
	If ($PageError2 -ne $null)
	{
		Write-Output "Error scanning site $SiteToScan $($PageError2.Message)"
		break
	}
	else
	{
		write-host "Successfully Got Results for" $_.Sites
	}
	
	If ($ScanResults -match "There was an error getting results")
	{
		Write-Host "The domain" $_.Sites "does not exist, or there was an error scanning"
		$Nodomain = $_.Sites
		$Site = New-Object System.Object
		$Site | Add-Member -type NoteProperty -name Hostname -value "$Nodomain does not exist"
		$SiteNotFound += $Site
	}
	else
	{
		Write-Host "The domain" $_.Sites "exists"
		#Parse out links to get XLSX Content
		$Results = $ScanResults.Links
		$Results = $Results -match ".xlsx"
		$XLSX = $Results.href
		$XLSXFilename = Split-Path $XLSX -leaf
		Write-Host "XLSX to Download is $XLSX"
		Invoke-WebRequest -Uri $XLSX -OutFile "$WorkingDir\$XLSXFilename"
		Write-Host "Saved Results in $WorkingDir"
		$ParsedResults += Import-Excel -FileName:"$WorkingDir\$XLSXFilename" -WorksheetName:"All Hosts"
	}
}
$ParsedResults | Export-CSV -Path "$WorkingDir\SiteScan.csv"
$SiteNotFound | Export-CSV -Path "$WorkingDir\SiteNotFound.csv"

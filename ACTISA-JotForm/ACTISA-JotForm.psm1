<#
	.NOTES
	===========================================================================
	 Created on:   	26/08/2019 11:16 AM
	 Created by:   	Aaron Duncan
	 Organization: 	
	 Filename:     	ACTISA-JotForm.psm1
	-------------------------------------------------------------------------
	 Module Name: ACTISA-JotForm
	===========================================================================

	.SYNOPSIS
		Collection of generic routines used in processing ACTISA jotforms.
#>

<#
	.SYNOPSIS
		Get Submission Entries from Google Sheets.
	
	.DESCRIPTION
		Get Submission Entries from Google Sheets.
	
	.PARAMETER url
		A description of the url parameter.
	
	.EXAMPLE
		PS C:\> Get-SubmissionEntries
	
	.NOTES
		Additional information about the function.
#>
function Get-SubmissionEntries
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$url
	)
	
	$entries = $null
	
	# Since people love to include filenames with umlauts etc, we can't use any
	# powershell routines that try to parse the encoding
	# (because they get them wrong)
	#
	# this includes:
	#   - using the ".Content" member from "Invoke-WebRequest"
	#   - using "Get-Content"
	#
	# Instead, we'll pass the output from Invoke-WebRequest directly to a temporary file
	Write-Host "Get-SubmissionEntries-local(): url = '$url'"
	if ($url.Contains("?output="))
	{
		$format = $url.Split('=')[-1]
		$tmpfile = New-TemporaryFile
		Invoke-WebRequest -Uri $url -OutFile $tmpfile
		$entries = Import-SpreadsheetFile -filename $tmpfile -format $format
	}
	else
	{
		Write-Error "Invalid Spreadsheet URL provided (this should end in '?output=...')"
	}
	
	return $entries
}

<#
	.SYNOPSIS
		Import spreadsheet according to specified format
	
	.DESCRIPTION
		A detailed description of the Import-Spreadsheet function.
	
	.PARAMETER filename
		A description of the filename parameter.
	
	.PARAMETER format
		A description of the format parameter.
	
	.EXAMPLE
		PS C:\> Import-Spreadsheet
	
	.NOTES
		Additional information about the function.
#>
function Import-SpreadsheetFile
{
	[CmdletBinding()]
	param
	(
		[Alias('path')]
		[string]$filename,
		[ValidateSet('csv', 'tsv', 'xlsx')]
		[ValidateNotNullOrEmpty()]
		[Alias('type')]
		[string]$format
	)
	
	$filepath = (Resolve-Path $filename).Path
	
	if ($format -eq 'xlsx')
	{
		$entries = Import-XLSX $filepath
	}
	elseif ($format -eq 'tsv')
	{
		$entries = Import-Csv -Delimiter "`t" $filepath
	}
	elseif ($format -eq 'csv')
	{
		$entries = Import-Csv $filepath
	}
	else
	{
		Write-Error "Invalid Spreadsheet format: $format"
		$entries = $null
	}
	return $entries
}

<#
	.SYNOPSIS
		Import Excel Spreadsheet
	
	.DESCRIPTION
		A detailed description of the Import-XLSX function.
	
	.PARAMETER filename
		A description of the filename parameter.
	
	.EXAMPLE
				PS C:\> Import-XLSX -filename 'Value1'
	
	.NOTES
		Additional information about the function.
#>
function Import-XLSX
{
	[CmdletBinding()]
	[OutputType([pscustomobject[]])]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$filename
	)
	
	$filepath = (Resolve-Path $filename).Path
	
	# Create an excel object using the Com interface
	$objExcel = New-Object -ComObject Excel.Application
	
	# Disable the 'visible' property so the document won't open in excel
	$objExcel.Visible = $false
	
	# Open the Excel file and save it in $WorkBook
	$WorkBook = $objExcel.WorkBooks.Open($filepath)
	
	$WorkSheet = $WorkBook.ActiveSheet
	
	$numRows = $WorkSheet.UsedRange.Rows.Count
	$numCols = $WorkSheet.UsedRange.Columns.Count
	
	$entries = @()
	$headerRow = $WorkSheet.Rows.Item(1)
	for ($r = 2; $r -le $numRows; $r++)
	{
		$entryRow = [ordered]@{ }
		$row = $WorkSheet.Rows.Item($r)
		for ($c = 1; $c -le $numCols; $c++)
		{
			$label = $headerRow.Columns.Item($c).Text
			$value = $row.Columns.Item($c).Text
			$entryRow.Add($label, $value)
		}
		$entries += [PSCustomObject]$entryRow
	}
	
	# close excel
	$WorkBook.Close()
	$objExcel.Quit()
	
	# release the COM objects
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet)
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook)
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
	[System.GC]::Collect()
	
	return $entries
}

<#
	.SYNOPSIS
		Download File from the Web
	
	.DESCRIPTION
		A detailed description of the Get-WebFile function.
	
	.PARAMETER url
		A description of the url parameter.
	
	.PARAMETER destination
		A description of the destination parameter.
	
	.EXAMPLE
		PS C:\> Get-WebFile
	
	.NOTES
		Additional information about the function.
#>
function Get-WebFile
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$url,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$destination
	)
	
	try
	{
		#$progressPreference = 'silentlyContinue'
		
		# powershell defaults to SSL3/TLS1.0 but jotform (rightly) only accepts TLS1.2+
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		
		# encode the url for safety
		$index = $url.LastIndexOf('/') + 1
		$enc_url = $url.Substring(0, $index) + [uri]::EscapeDataString($url.Substring($index))
		
		Invoke-WebRequest -Uri $enc_url -OutFile $destination
	}
	catch
	{
		"Download failed: '$url'"
		"Encoded url: '$encoded'"
		"Destination: '$destination'"
		"ERROR Message:" + $_.Exception.Message
	}
}

<#
	.SYNOPSIS
		Apply Mailmerge
	
	.DESCRIPTION
		A detailed description of the Invoke-MailMerge function.
	
	.PARAMETER template
		A description of the template parameter.
	
	.PARAMETER datasource
		A description of the datasource parameter.
	
	.PARAMETER destination
		A description of the destination parameter.
	
	.EXAMPLE
		PS C:\> Invoke-MailMerge
	
	.NOTES
		Additional information about the function.
#>
function Invoke-MailMerge
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$template,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$datasource,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$destination
	)
	
	if ([String]::IsNullOrEmpty($template))
	{
		Write-Warning "ERROR: empty MailMerge template file supplied to Invoke-MailMerge()"
	}
	elseif ((Test-Path -Path $template -ErrorAction SilentlyContinue) -eq $false)
	{
		Write-Warning "ERROR: missing MailMerge template file '$template'"
	}
	elseif ((Test-Path -Path $datasource -ErrorAction SilentlyContinue) -eq $false)
	{
		Write-Warning "ERROR: missing MailMerge datasource file '$datasource'"
	}
	else
	{
		if ((Test-Path -Path $destination -ErrorAction SilentlyContinue) -eq $false)
		{
			New-Item $destination -Type Directory | Out-Null
		}
		
		$word = New-Object -ComObject Word.Application
		$word.Visible = $false
		
		$Doc = $word.Documents.Open($template)
		$Doc.MailMerge.OpenDataSource($datasource)
		$Doc.MailMerge.Execute()
		
		
		if (Test-Path -Path $destination -Type Container)
		{
			#
			# the destination is a folder, so generate the destination name to be the 
			# template name with a pdf extension (located in the destination folder)
			#
			$extension = ".pdf"
			
			$filename = [System.IO.Path]::GetFileNameWithoutExtension($template) + $extension
			$destfile = [System.IO.Path]::Combine($destination, $filename)
		}
		else
		{
			#
			# the destination file has been specified, so determine if it is a word doc or pdf
			#
			$extension = [System.IO.Path]::GetExtension($template)
			$destfile = $destination
		}
		
		if ($extension -match ".pdf")
		{
			#$word.ActiveDocument.ExportAsFixedFormat($destfile, 17)
			$PDF = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
			$word.ActiveDocument.SaveAs($destfile, [ref]$PDF)
		}
		else
		{
			$word.ActiveDocument.SaveAs($destfile)
		}
		
		$Close = [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges
		
		$word.ActiveDocument.Close([ref]$Close)
		$Doc.Close([ref]$Close)
		
		$word.Quit()
	}
}

<#
	.SYNOPSIS
		Get Music File Duration
	
	.DESCRIPTION
		A detailed description of the Get-MusicFileDuration function.
	
	.PARAMETER filename
		A description of the filename parameter.
	
	.EXAMPLE
		PS C:\> Get-MusicFileDuration
	
	.NOTES
		Additional information about the function.
#>
function Get-MusicFileDuration
{
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$filename
	)
	
	Write-Host "Getting music file duration for '$filename'"
	
	$duration = $null
	if (Test-Path -Path "$filename" -PathType Leaf -ErrorAction SilentlyContinue)
	{
		$shell = New-Object -COMObject Shell.Application
		
		$filepath = (Resolve-Path $filename).Path
		$folder = Split-Path $filepath
		$file = Split-Path $filepath -Leaf
		Write-Host "FILE: $file"
		$shellfolder = $shell.Namespace($folder)
		$shellfile = $shellfolder.ParseName($file)
		
		# find the index of "Length" (limited to 1000 extended attributes to stop infinite looping)
		$index = 0
		for ($index = 0; $index -lt 1000; $index++)
		{
			$details = $shellfolder.GetDetailsOf($shellfolder.Items, $index)
			if ([string]::IsNullOrEmpty($details))
			{
				Write-Warning "failed to get music duration"
				$duration = 'notfound'
				break
			}
			elseif ($details -eq 'Length')
			{
				Write-Host "details = Length"
				$length_details = $shellfolder.GetDetailsOf($shellfile, $index)
				if (-not [string]::IsNullOrEmpty($length_details))
				{
					$length  = $length_details.split(':')
					$minutes = $length[1].ToString().TrimStart('0')
					$seconds = $length[2].ToString().TrimStart('0')
					
					"MINUTES: $minutes SECONDS: $seconds"
					if ([int]$minutes -lt 10) { $minutes = '0' + $minutes }
					if ([int]$seconds -lt 10) { $seconds = '0' + $seconds }
					
					$duration = "${minutes}m${seconds}s"
				}
				else
				{
					Write-Host "failed to split details: $($shellfolder.GetDetailsOf($shellfile, $index))"
					$duration = 'notfound'
				}
				break
			}
		}
	}
	else
	{
		Write-Warning "Failed to find file: $filename"
	}
	
	return $duration
}

<#
	.SYNOPSIS
		Capitalizes names according to normal "people name" conventions
	
	.DESCRIPTION
		A detailed description of the ConvertTo-CapitalizedName function.
	
	.PARAMETER name
		A description of the name parameter.
	
	.EXAMPLE
		PS C:\> ConvertTo-CapitalizedName
	
	.NOTES
		Additional information about the function.
#>
function ConvertTo-CapitalizedName
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[string]$name
	)
	
	(Get-Culture).textinfo.totitlecase($name.ToLower().Replace("'", "_")).Replace("_", "'").Replace("De ", "de ") -replace " +", " "
}

<#
	.SYNOPSIS
		Provides User Folder Selection Dialog
	
	.DESCRIPTION
		A detailed description of the Find-Folders function.
	
	.PARAMETER title
		A description of the title parameter.
	
	.PARAMETER default
		A description of the default parameter.
	
	.EXAMPLE
		PS C:\> Find-Folders
	
	.NOTES
		Additional information about the function.
#>
function Find-Folders
{
	param
	(
		[Parameter(Mandatory = $false)]
		[ValidateNotNull()]
		[string]$title = 'Select the Competition directory',
		[ValidateNotNull()]
		[string]$default = 'C:\'
	)
	
	#[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
	Add-Type -AssemblyName System.Windows.Forms
	
	[System.Windows.Forms.Application]::EnableVisualStyles()
	
	$browse = New-Object System.Windows.Forms.FolderBrowserDialog
	$browse.SelectedPath = $default
	$browse.RootFolder = [System.Environment+SpecialFolder]::Desktop
	$browse.ShowNewFolderButton = $true
	$browse.Description = $title
	
	$result = $browse.ShowDialog((New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }))
	if ($result -eq [Windows.Forms.DialogResult]::OK)
	{
		$browse.SelectedPath
		$browse.Dispose()
	}
	else
	{
		exit
	}
}

<#
	.SYNOPSIS
		Provides Selection Dialog for MailMerge Template
	
	.DESCRIPTION
		A detailed description of the Find-Template function.
	
	.PARAMETER message
		A description of the message parameter.
	
	.PARAMETER initial_dir
		A description of the initial_dir parameter.
	
	.PARAMETER default
		A description of the default parameter.
	
	.EXAMPLE
		PS C:\> Find-Template
	
	.NOTES
		Additional information about the function.
#>
function Find-Template
{
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[string]$message,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$initial_dir,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$default
	)
	
	$openFileDialog = New-Object windows.forms.openfiledialog
	$openFileDialog.initialDirectory = $initial_dir
	$openFileDialog.title = $message
	$openFileDialog.filter = "Word Files (*.doc, *.docx)|*.doc;*.docx"
	$openFileDialog.FileName = $default
	$openFileDialog.ShowHelp = $True
	$result = $openFileDialog.ShowDialog()
	
	# in ISE you may have to alt-tab or minimize ISE to see dialog box  
	if ($result -eq "OK")
	{
		$OpenFileDialog.filename
	}
	else
	{
		Write-Host "Template Selection Cancelled!" -ForegroundColor Yellow
	}
}

<#
	.SYNOPSIS
		Provides Selection Dialog for MailMerge DataSource file.
	
	.DESCRIPTION
		A detailed description of the Find-DataSource function.
	
	.PARAMETER message
		A description of the message parameter.
	
	.PARAMETER initial_dir
		A description of the initial_dir parameter.
	
	.PARAMETER default
		A description of the default parameter.
	
	.EXAMPLE
		PS C:\> Find-Template
	
	.NOTES
		Additional information about the function.
#>
function Find-DataSource
{
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[string]$message,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$initial_dir,
		[Parameter(Mandatory = $true)]
		[string]$default
	)
	
	[string]$filename = $null
	
	$openFileDialog = New-Object windows.forms.openfiledialog
	$openFileDialog.initialDirectory = $initial_dir
	$openFileDialog.title = $message
	$openFileDialog.filter = "*.csv;*.txt"
	$openFileDialog.FileName = $default
	$openFileDialog.ShowHelp = $True
	$result = $openFileDialog.ShowDialog()
	
	# in ISE you may have to alt-tab or minimize ISE to see dialog box  
	if ($result -eq "OK")
	{
		$filename = $OpenFileDialog.filename
	}
	else
	{
		Write-Host "Import Settings File Cancelled!" -ForegroundColor Yellow
	}
	
	return $filename
}

<#
	.SYNOPSIS
		Adds two music durations and returns the result
	
	.DESCRIPTION
		A detailed description of the Add-MusicDurations function.
	
	.PARAMETER duration1
		First duration, in the form: "<number_of_minutes>m<number_of_seconds>s"
	
	.PARAMETER duration2
		First duration, in the form: "<number_of_minutes>m<number_of_seconds>s"
	
	.EXAMPLE
		PS C:\> Add-MusicDurations "1m20s" "0m23s"
	
	.NOTES
		Additional information about the function.
#>
function Add-MusicDurations
{
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$duration1,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$duration2
	)
	
	$minutes = 0
	$seconds = 0
	
	if ($duration1 -match "(\d+)m(\d+)s")
	{
		$minutes += $Matches[1]
		$seconds += $Matches[2]
	}
	
	if ($duration2 -match "(\d+)m(\d+)s")
	{
		$minutes += $Matches[1]
		$seconds += $Matches[2]
	}
	
	while ($seconds -ge 60)
	{
		$minutes++
		$seconds -= 60
	}
	
	$seconds = "{0:D2}" -f $seconds
	"${minutes}m${seconds}s"
}

<#
	.SYNOPSIS
		Creates a new XLS/CSV Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-Spreadsheet function.
	
	.PARAMETER name
		Name to be assigned to the tab for EXCEL spreadsheet.  This parameter is not used for CSV formatted spreadsheets.
	
	.PARAMETER path
		Full filepath of the spreadsheet to be created.
	
	.PARAMETER headers
		An array of labels to be used on the top line of the spreadsheet and used as headers.  This row, in the case of EXCEL spreadsheets, will have a filter applied to it.
	
	.PARAMETER rows
		This is a
	
	.PARAMETER format
		A description of the format parameter.
	
	.EXAMPLE
		PS C:\> New-Spreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-Spreadsheet
{
	param
	(
		[string]$name,
		[Parameter(Mandatory = $true)]
		[string]$path,
		[string[]]$headers,
		[hashtable[]]$rows,
		[string]$format = 'csv'
	)
	if ($format -eq 'xls')
	{
		if (!(Test-Path HKLM:SOFTWARE\Classes\Excel.Application))
		{
			Write-Host "Excel Application not installed on this computer.  Switching to CSV."
			$path = [System.IO.Path]::ChangeExtension($path, "csv")
			$format = 'csv'
		}
	}
	
	if ($format.StartsWith('xls'))
	{
		try
		{
			$excel = New-Object -ComObject excel.application
			$excel.DisplayAlerts = $false
			$excel.Visible = $false
			
			#Add a workbook
			$workbook = $excel.Workbooks.Add()
			
			$sheet = $workbook.Worksheets.Item(1)
			$sheet.Activate() | Out-Null
			
			$sheet.Name = $name
			
			# Create the header line
			for ($c = 1; $c -le $headers.Count; $c++)
			{
				$sheet.Cells.Item(1, $c) = $headers[$c - 1]
				$sheet.Cells.Item(1, $c).Interior.ColorIndex = 15
				$sheet.Cells.Item(1, $c).Font.Bold = $True
			}
			
			for ($r = 1; $r -le $rows.Count; $r++)
			{
				$row = $rows[$r - 1]
				if ($row -ne $null)
				{
					$values = $row['values']
					$border = $row['border']
					for ($c = 1; $c -le $values.Count; $c++)
					{
						$cell = $sheet.Cells($r + 1, $c)
						$cell.value = [String]$values[$c - 1]
						if ($border)
						{
							$cell.Borders.LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
						}
					}
				}
			}
			
			
			# format the first row to an autofilter
			$range = $sheet.Range("A1", "$([char]($headers.Count + 64))1")
			$range.AutoFilter() | Out-Null
			$range.Borders.LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlContinuous
			$range.Borders([Microsoft.Office.Interop.Excel.XlBordersIndex]::xlEdgeBottom).LineStyle = [Microsoft.Office.Interop.Excel.XlLineStyle]::xlDouble
			
			$sheet.Columns("A:$([char]($headers.Count + 64))").AutoFit() | Out-Null
			if (Test-Path $path) { Remove-Item -Path $path }
			$workbook.SaveAs($path)
			$excel.Workbooks.Close()
			$excel.Quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
		}
		catch
		{
			Write-Warning "Creating XLS Spreadsheet '$name' failed.  Switching to CSV format"
			$path = [System.IO.Path]::ChangeExtension($path, "csv")
			Write-Host "New filepath = $path"
			$format = 'csv'
		}
	}
	
	if ($format -eq 'csv')
	{
		if (Test-Path $path) { Remove-Item -Path $path }
		
		Add-Content -Path $path -Value ($headers -join ',')
		foreach ($row in $rows)
		{
			Add-Content -Path $path -Value ($row['values'] -join ',')
		}
	}
}

<#
.Synopsis
Retrieves extension attributes from files or folder

.DESCRIPTION
Uses the dynamically generated parameter -ExtensionAttribute to select one or multiple extension attributes and display the attribute(s) along with the FullName attribute

.NOTES   
Name: Get-ExtensionAttribute.ps1
Author: Jaap Brasser
Version: 1.0
DateCreated: 2015-03-30
DateUpdated: 2015-03-30
Blog: http://www.jaapbrasser.com

.LINK
http://www.jaapbrasser.com

.PARAMETER FullName
The path to the file or folder of which the attributes should be retrieved. Can take input from pipeline and multiple values are accepted.

.PARAMETER ExtensionAttribute
Additional values to be loaded from the registry. Can contain a string or an array of string that will be attempted to retrieve from the registry for each program entry

.EXAMPLE   
. .\Get-ExtensionAttribute.ps1
    
Description 
-----------     
This command dot sources the script to ensure the Get-ExtensionAttribute function is available in your current PowerShell session

.EXAMPLE
Get-ExtensionAttribute -FullName C:\Music -ExtensionAttribute Size,Length,Bitrate

Description
-----------
Retrieves the Size,Length,Bitrate and FullName of the contents of the C:\Music folder, non recursively

.EXAMPLE
Get-ExtensionAttribute -FullName C:\Music\Song2.mp3,C:\Music\Song.mp3 -ExtensionAttribute Size,Length,Bitrate

Description
-----------
Retrieves the Size,Length,Bitrate and FullName of Song.mp3 and Song2.mp3 in the C:\Music folder

.EXAMPLE
Get-ChildItem -Recurse C:\Video | Get-ExtensionAttribute -ExtensionAttribute Size,Length,Bitrate,Totalbitrate

Description
-----------
Uses the Get-ChildItem cmdlet to provide input to the Get-ExtensionAttribute function and retrieves selected attributes for the C:\Videos folder recursively

.EXAMPLE
Get-ChildItem -Recurse C:\Music | Select-Object FullName,Length,@{Name = 'Bitrate' ; Expression = { Get-ExtensionAttribute -FullName $_.FullName -ExtensionAttribute Bitrate | Select-Object -ExpandProperty Bitrate } }

Description
-----------
Combines the output from Get-ChildItem with the Get-ExtensionAttribute function, selecting the FullName and Length properties from Get-ChildItem with the ExtensionAttribute Bitrate
#>
function Get-ExtensionAttribute
{
	[CmdletBinding()]
	Param (
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[string[]]$FullName
	)
	DynamicParam
	{
		$Attributes = new-object System.Management.Automation.ParameterAttribute
		$Attributes.ParameterSetName = "__AllParameterSets"
		$Attributes.Mandatory = $false
		$AttributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
		$AttributeCollection.Add($Attributes)
		$Values = @($Com = (New-Object -ComObject Shell.Application).NameSpace('C:\'); 1 .. 400 | ForEach-Object { $com.GetDetailsOf($com.Items, $_) } | Where-Object { $_ } | ForEach-Object { $_ -replace '\s' })
		$AttributeValues = New-Object System.Management.Automation.ValidateSetAttribute($Values)
		$AttributeCollection.Add($AttributeValues)
		$DynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ExtensionAttribute", [string[]], $AttributeCollection)
		$ParamDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
		$ParamDictionary.Add("ExtensionAttribute", $DynParam1)
		$ParamDictionary
	}
	
	begin
	{
		$ShellObject = New-Object -ComObject Shell.Application
		$DefaultName = $ShellObject.NameSpace('C:\')
		$ExtList = 0 .. 400 | ForEach-Object {
			($DefaultName.GetDetailsOf($DefaultName.Items, $_)).ToUpper().Replace(' ', '')
		}
	}
	
	process
	{
		foreach ($Object in $FullName)
		{
			# Check if there is a fullname attribute, in case pipeline from Get-ChildItem is used
			if ($Object.FullName)
			{
				$Object = $Object.FullName
			}
			
			# Check if the path is a single file or a folder
			if (-not (Test-Path -Path $Object -PathType Container))
			{
				$CurrentNameSpace = $ShellObject.NameSpace($(Split-Path -Path $Object))
				$CurrentNameSpace.Items() | Where-Object {
					$_.Path -eq $Object
				} | ForEach-Object {
					$HashProperties = @{
						FullName    = $_.Path
					}
					foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute)
					{
						$HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
					}
					New-Object -TypeName PSCustomObject -Property $HashProperties
				}
			}
			elseif (-not $input)
			{
				$CurrentNameSpace = $ShellObject.NameSpace($Object)
				$CurrentNameSpace.Items() | ForEach-Object {
					$HashProperties = @{
						FullName    = $_.Path
					}
					foreach ($Attribute in $MyInvocation.BoundParameters.ExtensionAttribute)
					{
						$HashProperties.$($Attribute) = $CurrentNameSpace.GetDetailsOf($_, $($ExtList.IndexOf($Attribute.ToUpper())))
					}
					New-Object -TypeName PSCustomObject -Property $HashProperties
				}
			}
		}
	}
	
	end
	{
		Remove-Variable -Force -Name DefaultName
		Remove-Variable -Force -Name CurrentNameSpace
		Remove-Variable -Force -Name ShellObject
	}
}

Export-ModuleMember -Function Get-SubmissionEntries,
					Get-WebFile,
					Import-SpreadsheetFile,
					Invoke-MailMerge,
					Get-MusicFileDuration,
					ConvertTo-CapitalizedName,
					Find-Folders,
					Find-Template,
					Find-DataSource,
					Add-MusicDurations,
					New-Spreadsheet,
					Get-ExtensionAttribute

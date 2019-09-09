
[Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
[Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.Word") | Out-Null

$env:PSModulePath = $PSScriptRoot + ';' + $env:PSModulePath
Import-Module ACTISA-JotForm

#================================================================================
#---------------------      GLOBAL VARIABLES/OBJECTS      -----------------------
#================================================================================

#
# create a hashtable here for the skating warmup/performance times etc
#
$MAX_WARMUP_GROUP_SIZE = 8

$abbreviations = @{
	"Free Skate"	  = "FS";
	"Aussie Skate"    = "AS";
	"Free Dance"	  = "FD";
	"Short Program"   = "SP";
	"Free Program"    = "FP";
	"Advanced Novice" = "AdvNov";
}


#
# To link to the google spreadsheet:
#     - open sheet
#     - select "File > Publish to the web ..." menuitem
#     - leave the 1st dropdown as "Entire Document"
#     - change the 2nd dropdown to "Tab-separated values (.tsv)
#     - click Publish
#     - Click OK
#     - copy link

# the '2019 Gwen Peterson Theatre on Ice Registration' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6oYCuRxKVL2G7rnlWAErA7luYwS53H7wbX7XaU2NDRKwv8NkSdGt80xjhMQKPS-8B7Hj2wuVP6wH3/pub?output=tsv'

$template_folder = 'C:\Users\aaron\Google Drive\Skating\Skating Templates';

$Competition = "GPTOI $(Get-Date -Format yyyy)";
$Engraving_Title = "Gwen Peterson Theatre On Ice $(Get-Date -Format yyyy)"

$certificate_template = "Certificate - $Competition.docx"

$comp_folder = "D:\ACTISA_COMP - $($Competition)"

if (!(Test-Path -Path $comp_folder))
{
	New-Item -ItemType Directory -Force -Path $comp_folder | Out-Null
}

#================================================================================
#-------------------------          FUNCTIONS          --------------------------
#================================================================================

<#
	.SYNOPSIS
		Renames music file according to filename conventions, and publishes it to the appropriate location in the Music folder
	
	.DESCRIPTION
		A detailed description of the Publish-MusicFile function.
	
	.PARAMETER filename
		A description of the filename parameter.
	
	.PARAMETER division
		A description of the division parameter.
	
	.PARAMETER teamname
		A description of the teamname parameter.
	
	.PARAMETER destination
		A description of the destination parameter.
	
	.EXAMPLE
				PS C:\> Publish-MusicFile
	
	.NOTES
		Additional information about the function.
#>
function Publish-MusicFile
{
	param
	(
		$filename,
		$division,
		$teamname,
		$destination
	)
	
	#
	# Get the duration of the song from the metadata
	#
	try
	{
		$music_duration = Get-MusicFileDuration -filename $filename
	}
	catch
	{
		$music_duration = $null
		Write-Warning "Invalid Music File: $filename"
	}
	
	#
	# determine the destination music folder
	#
	$music_subdir = $division
	$music_dest = [System.IO.Path]::Combine($destination, $music_subdir)
	
	#
	# Calculate the new music filename
	#
	
	$new_music_file = "TOI_${division}"
	
	foreach ($key in $abbreviations.Keys)
	{
		$new_music_file = $new_music_file -replace $key, $abbreviations.Item($key)
	}
	
	$new_music_file = $new_music_file -replace " ; ", "_" -replace " ", ""
	
	$new_music_file += "_${teamname}"
	
	if ([String]::IsNullOrEmpty($music_duration))
	{
		$new_music_file = "BADFILE_${new_music_file}"
	}
	elseif ($music_duration -match "notfound")
	{
		$new_music_file = "NOTFOUND_${new_music_file}"
	}
	else
	{
		$new_music_file += "_${music_duration}"
	}
	
	$new_music_file += $extension
	
	$new_music_path = [System.IO.Path]::Combine($music_dest, $new_music_file)
	
	if ((Test-Path $music_dest -ErrorAction SilentlyContinue) -eq $false)
	{
		$music_dest = [System.IO.Path]::Combine($destination, $music_subdir)
		New-Item $music_dest -Type Directory | Out-Null
	}
	
	if ((Test-Path $filename -PathType Leaf -ErrorAction SilentlyContinue) -eq $false)
	{
		if ((Test-Path $new_music_path -PathType Leaf -ErrorAction SilentlyContinue) -eq $false)
		{
			New-Item -Path $new_music_path -ItemType File | Out-Null
		}
	}
	else
	{
		Copy-Item -Path $filename $new_music_path -Force -ErrorAction SilentlyContinue
		if ($? -eq $false)
		{
			Write-Warning "Failed to copy '$filename' -> '$new_music_path'"
		}
		else
		{
			"Source Music File: $filename"
			"Music Destination: $new_music_path"
			""
		}
	}
}

<#
	.SYNOPSIS
		Process All Music Files for a Submission Entry
	
	.DESCRIPTION
		Calls Publish-SkaterMusicFile() for each music file for a submission, which can vary from 1-3 files depending on the divison.
	
	.PARAMETER entry
		A description of the entry parameter.
	
	.PARAMETER music_folder
		A description of the music_folder parameter.
	
	.PARAMETER submissionFullPath
		A description of the submissionFullPath parameter.
	
	.EXAMPLE
				PS C:\> Publish-EntryMusicFiles
	
	.NOTES
		Additional information about the function.
#>
function Publish-EntryMusicFiles
{
	param
	(
		$entry,
		$music_folder,
		$submissionFullPath
	)
	
	$submission_id = $entry.'Submission ID'
	$division = $entry.Division
	$gender = $entry.'Skater 1 Gender'
	$music_url = $entry.'Music File'
	$team_name = $entry.'Team Name:'
	
	$submission_folder =
	(Resolve-Path "$([System.IO.Path]::Combine($submissionFullPath, "${submission_id}_"))*").Path
	
	if ([string]::IsNullOrEmpty($submission_folder))
	{
		$submission_folder = [System.IO.Path]::Combine($submissionFullPath, "${submission_id}_-")
	}
	
	#Write-Host "Submission Fullpath: $submissionFullPath"
	Write-Host "Team Name: $team_name"
	
	$music_file = [System.Web.HttpUtility]::UrlDecode($music_url.Split("/")[-1]);
	$music_fullpath = [System.IO.Path]::Combine($submission_folder, $music_file)
	$extension = [System.IO.Path]::GetExtension($music_file)
	
	if ((Test-Path -Path $submission_folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $submission_folder -Type Directory | Out-Null
	}
	
	if ((Test-Path -Path $music_fullpath -ErrorAction SilentlyContinue) -eq $false)
	{
		# music file is missing, so download it
		Get-WebFile -url $music_url -destination $music_fullpath
	}
	
	Publish-MusicFile -filename $music_fullpath -division $division -teamname $team_name -destination $music_folder
}

<#
	.SYNOPSIS
		Generate List of Skater Certificates
	
	.DESCRIPTION
		A detailed description of the New-CertificateList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.EXAMPLE
				PS C:\> New-CertificateList
	
	.NOTES
		Additional information about the function.
#>
function New-CertificateList
{
	param
	(
		$entries,
		$folder
	)
	
	Write-Host "Generating Certificates"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$template = Find-Template -message "Select Certificate Template" -initial_dir $template_folder -default $certificate_template
	
	$input_csv = [System.IO.Path]::Combine($folder, "certificate_inputs.csv")
	
	#Write-Host "Input CSV file: $input_csv"
	
	$results = @()
	
	$entries | ForEach-Object {
		
		$division = $_.Division
		
		$names = $_."Team Members:"
		foreach ($member in $names.Split("`r"))
		{
			$name = ConvertTo-CapitalizedName -name $member.trim()
			
			if (-not [String]::IsNullOrWhiteSpace($name))
			{
				#Write-Host "Generating certificate for $name ($division)"
				
				$results += New-Object -TypeName PSObject -Property @{
					"Name"	   = $name
					"Division" = $division
				}
			}
		}
	}
	
	Write-Host "Number of Certificates = $($results.Count)"
	$results | Select-Object "Name", "Division" | export-csv -path $input_csv -Force -NoTypeInformation
	
	Invoke-MailMerge -template $template -datasource $input_csv -destination $folder
}

<#
	.SYNOPSIS
		Generate Skating Schedule
	
	.DESCRIPTION
		A detailed description of the New-SkatingSchedule function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.EXAMPLE
				PS C:\> New-SkatingSchedule
	
	.NOTES
		Additional information about the function.
#>
function New-SkatingSchedule
{
	param
	(
		$entries,
		$folder
	)
	
	Write-Host "Generating Skating Schedule"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$schedule1 = [System.IO.Path]::Combine($folder, "Schedule1.csv")
	$schedule2 = [System.IO.Path]::Combine($folder, "Schedule2.csv")
	
	if (Test-Path $schedule1) { Remove-Item -Path $schedule1 }
	if (Test-Path $schedule2) { Remove-Item -Path $schedule2 }
	
	$divhash = @{ }
	foreach ($entry in $entries)
	{
		$category = "TOI"
		$division = $entry.'Division'
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	# $divhash.GetEnumerator() | Sort Name
	
	foreach ($div in $divhash.Keys | Sort)
	{
		# NOTE: need to put this in a hash based on category and division
		# Perhaps something like have a category number, which is overridden by division
		$warmup = 0
		$performance_time = 12 # this is not right - but a hack until a proper set of values is implemented
		
		Add-Content -Path $schedule1 -Value "Category: $category,Division: $div"
		
		$count = $divhash.Item($div).Count
		$num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
		Add-Content -path $schedule1 -Value "Num Entries: $count"
		"Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes" -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule1
		"Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time) | Add-Content -Path $schedule1
		
		Add-Content -path $schedule1 -Value "Team,State,Coach Name,Music Title"
		
		$divhash.Item($div) | ForEach-Object {
			$num_segments = $_.'Number of Segments in Music'
			if ($_.'Music Details (Segment 1)' -match 'Title: (.*) Artist:')
			{
				$music_title = $Matches[1]
			}
			else
			{
				$music_title = '<NOT PROVIDED>'
			}
			
			$coach_name = $_.'Primary Coach Name:'
			if ($coach_name -eq 'Other')
			{
				$coach_name = $_.'Primary Coach Name: 2'
			}
			
			"{0},{1},{2},{3},{4},{5}" -f $_.'Team Name:', $_.'Team State:', $coach_name, $music_title.Trim(), $_.'Coach E-mail', $_.'Other Coach Names' | Add-Content -path $schedule1
		}
		Add-Content -Path $schedule1 -Value ""
	}
}

<#
	.SYNOPSIS
		Get List of Skater Names for an Entry
	
	.DESCRIPTION
		A detailed description of the Get-EntrySkaterNames function.
	
	.PARAMETER entry
		A description of the entry parameter.
	
	.EXAMPLE
				PS C:\> Get-EntrySkaterNames
	
	.NOTES
		Additional information about the function.
#>
function Get-EntrySkaterNames
{
	param
	(
		$entry
	)
	
	$members = @()
	
	$names = $entry.'Team Members:'
	foreach ($name in $names.Split("`r"))
	{
		if (-not [String]::IsNullOrWhiteSpace($name))
		{
			$members += ConvertTo-CapitalizedName -name $name
		}
	}
	
	return $members
}

<#
	.SYNOPSIS
		Get number of skaters listed in an Entry
	
	.DESCRIPTION
		A detailed description of the Get-EntryNumberOfSkaters function.
	
	.PARAMETER entry
		A description of the entry parameter.
	
	.EXAMPLE
		PS C:\> Get-EntryNumberOfSkaters
	
	.NOTES
		Additional information about the function.
#>
function Get-EntryNumberOfSkaters
{
	[OutputType([int])]
	param
	(
		$entry
	)
	
	return (Get-EntrySkaterNames -entry $entry).Count
}

<#
	.SYNOPSIS
		Get the maximum team size from a list of entries
	
	.DESCRIPTION
		A detailed description of the Get-MaxEntryNumberOfSkaters function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.EXAMPLE
				PS C:\> Get-MaxEntryNumberOfSkaters
	
	.NOTES
		Additional information about the function.
#>
function Get-MaxEntryNumberOfSkaters
{
	[OutputType([int])]
	param
	(
		$entries
	)
	
	$max_size = 1
	foreach ($entry in $entries)
	{
		$size = Get-EntryNumberOfSkaters -entry $entry
		$max_size = [Math]::Max($size, $max_size)
	}
	return $max_size
}

<#
	.SYNOPSIS
		Create Division Counts Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-DivisionCountsSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (csv,xlsx)
	
	.EXAMPLE
				PS C:\> New-DivisionCountsSpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-DivisionCountsSpreadsheet
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Division Counts Spreadsheet"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$divcounts_csv = [System.IO.Path]::Combine($folder, "division_medal_counts.${format}")
	
	if (Test-Path $divcounts_csv) { Remove-Item -Path $divcounts_csv }
	
	# group entries by division
	$divhash = @{ }
	foreach ($entry in $entries)
	{
		$division = $entry.'Division'
		
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	$rows = @()
	
	$medal_count = @{ gold = 0; silver = 0; bronze = 0 }
	
	foreach ($division in $divhash.Keys | Sort)
	{
		# get the number of entries in the division
		$count = $divhash.Item($division).Count
		
		# get the largest team size
		$max_size = Get-MaxEntryNumberOfSkaters -entries $divhash.Item($division)
		#Write-Host "Max team size of $division is $max_size"
		
		$medal_gold = ""
		$medal_silver = ""
		$medal_bronze = ""
		if ($count -ge 1) { $medal_gold = $max_size; $medal_count.gold += $max_size }
		if ($count -ge 2) { $medal_silver = $max_size; $medal_count.silver += $max_size }
		if ($count -ge 3) { $medal_bronze = $max_size; $medal_count.bronze += $max_size }
		
		$rows += (@{ 'border' = $true; 'values' = @($division, $count, $medal_bronze, $medal_silver, $medal_gold) })
	}
	
	$rows += (@{
			border = $false;
			values = @('', 'TOTAL:', $medal_count.bronze, $medal_count.silver, $medal_count.gold)
		})
	
	$headers = @('Division', '# Entries', 'Bronze Medal', 'Silver Medal', 'Gold Medal')
	New-SpreadSheet -name "Division Counts" -path $filepath -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create Medal engraving schedules
	
	.DESCRIPTION
		A detailed description of the New-EngravingSchedule function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.EXAMPLE
		PS C:\> New-EngravingSchedule
	
	.NOTES
		Additional information about the function.
#>
function New-EngravingSchedule
{
	param
	(
		$entries,
		$folder
	)
	
	Write-Host "Generating Engraving Schedule"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$trophyPath = [System.IO.Path]::Combine($folder, "${Competition} - MEDALS.docx")
	
	if (Test-Path $trophyPath) { Remove-Item -Path $trophyPath }
	
	$divhash = @{ }
	foreach ($entry in $entries)
	{
		$division = $entry.'Division'
		
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	#    $divhash.GetEnumerator() | Sort Name
	
	$cellAlignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
	
	# Create Trophy Table
	$TrophyWord = New-Object -ComObject Word.Application
	$TrophyWord.Visible = $False
	$TrophyDoc = $TrophyWord.Documents.Add()
	$margin = 36 # 1.26 cm
	$TrophyDoc.PageSetup.LeftMargin = $margin
	$TrophyDoc.PageSetup.RightMargin = $margin
	$TrophyDoc.PageSetup.TopMargin = $margin
	$TrophyDoc.PageSetup.BottomMargin = $margin
	$TrophyDoc.PageSetup.Orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape
	$TrophyTable = $TrophyDoc.Tables.Add($TrophyWord.Selection.Range(), 2, 4)
	$TrophyTable.Range.Style = "No Spacing"
	$TrophyTable.Range.paragraphFormat.Alignment = $cellAlignment
	$TrophyTable.Borders.Enable = $True
	$TrophyTable.Rows(2).Cells(1).Range.Text = "Division"
	$TrophyTable.Rows(2).Cells(1).Range.Bold = $True
	#$TrophyTable.Rows(2).Cells(1).Range.paragraphFormat.Alignment = $cellAlignment
	
	# Create the title row on each table
	$Row = $TrophyTable.Rows(1)
	$Row.Cells.Merge()
	$Row.Cells(1).Range.Text = "$Competition engraving schedule - TROPHY"
	$Row.Cells(1).Range.Bold = $true
	#$Row.Cells(1).Range.paragraphFormat.Alignment = $cellAlignment
	
	foreach ($div in $divhash.Keys | Sort)
	{
		#$category  = $div.Split("-",2)[0].trim()
		#$division  = $div.Split("-",2)[1].trim()
		$division = $div
		
		$divEntries = $divhash.Item($div)
		$count = $divEntries.Count
		$maxTeamSize = Get-MaxEntryNumberOfSkaters -entries $divEntries
		
		$Row = $TrophyTable.Rows.Add()
		$Row.Cells(1).Range.Text = "$division`r(x$maxTeamSize)"
		$Row.Cells(1).Range.Bold = $True
		$Row.Cells(1).Range.Font.Spacing = 1.0
		#$Row.Cells(1).Range.paragraphFormat.Alignment = $cellAlignment
		
		if ($count -ge 1) { $Row.Cells(2).Range.Text = "${Engraving_Title}`r$division`r1st Place" }
		if ($count -ge 2) { $Row.Cells(3).Range.Text = "${Engraving_Title}`r$division`r2nd Place" }
		if ($count -ge 3) { $Row.Cells(4).Range.Text = "${Engraving_Title}`r$division`r3rd Place" }
	}
	
	# Save the documents
	$TrophyDoc.SaveAs($trophyPath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
	$TrophyDoc.close()
	$TrophyWord.quit()
	
	# Cleanup the memory
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyDoc) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyWord) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyTable) | Out-Null
	Remove-Variable TrophyDoc, TrophyWord, Row, TrophyTable
	[gc]::collect()
	[gc]::WaitForPendingFinalizers()
}

<#
	.SYNOPSIS
		Generate List of Team E-mails
	
	.DESCRIPTION
		A detailed description of the New-TeamEmailList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		A description of the format parameter.
	
	.EXAMPLE
				PS C:\> New-TeamEmailList
	
	.NOTES
		Additional information about the function.
#>
function New-TeamEmailList
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Team Email List"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filename = [System.IO.Path]::Combine($folder, "team_email_list.${format}")
	if (Test-Path $filename) { Remove-Item -Path $filename }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{
				'border' = $true;
				'values' = @($entry.'Team Name:', $entry.'Team Manager:', $entry.'Team Manager Email:')
			})
	}
	if ($rows.Count -gt 0)
	{
		$headers = @('Team Name', 'Team E-Mail')
		Write-Host "$($rows.Count) teams found."
		New-Spreadsheet -name "Team E-Mail List" -path $filename -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no teams found."
	}
}

<#
	.SYNOPSIS
		Generate Coach E-mail List
	
	.DESCRIPTION
		A detailed description of the New-CoachEmailList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (csv,xlsx)
	
	.EXAMPLE
				PS C:\> New-CoachEmailList
	
	.NOTES
		Additional information about the function.
#>
function New-CoachEmailList
{
	param
	(
		$entries,
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Coach Email List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filename = [System.IO.Path]::Combine($folder, "coach_email_list.${format}")
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{
				'border' = $true;
				'values' = @($entry.'Primary Coach Name:', $entry.'Primary Coach E-mail:')
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @('Coach Name', 'Coach E-Mail')
		New-Spreadsheet -name "Coach E-Mail List" -path $filename -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no coaches found."
	}
}

<#
	.SYNOPSIS
		Generate Skater Registration CheckList
	
	.DESCRIPTION
		A detailed description of the New-RegistrationList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (csv,xlsx)
	
	.EXAMPLE
				PS C:\> New-RegistrationList
	
	.NOTES
		Additional information about the function.
#>
function New-RegistrationList
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Registration List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "registration_list.${format}")
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$skaters = @()
	foreach ($entry in $entries)
	{
		$skaters += Get-EntrySkaterNames -entry $entry
	}
	
	$rows = @()
	foreach ($skater in ($skaters | Sort-Object | Get-Unique))
	{
		$rows += (@{
				'border' = $true;
				'values' = @($skater, '', '')
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = "Skater Name, Water, Fruit"
		New-SpreadSheet -name "Registration" -path $filepath -headers $headers -rows $rows -format $format
	}
}

<#
	.SYNOPSIS
		Generate Volunteer Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-VolunteerSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (csv,xlsx)
	
	.EXAMPLE
				PS C:\> New-VolunteerSpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-VolunteerSpreadsheet
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Volunteer Spreadsheet"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "volunteer.${format}");
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		if (-not [String]::IsNullOrEmpty($entry.'I am able to assist with the following tasks:'))
		{
			$division = $entry.'Division'
			
			$rows += (@{
					'border' = $true;
					'values' = @(
						$entry.'Team Name:',
						$entry.'Team Manager:',
						$entry.'Division',
						$entry.'Volunteer Name',
						$entry.'Volunteer E-mail',
						$entry.'Volunteer Contact Mobile',
						$entry.'I am able to assist with the following tasks:')
				})
		}
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Team Name", "Team Manager", "Division", "Volunteer Name", "Volunteer E-mail", "Volunteer Phone", "Roles")
		New-SpreadSheet -name "Volunteers" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no volunteers found."
	}
}

<#
	.SYNOPSIS
		Generate Payment Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-PaymentSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (XLSX, CSV)
	
	.EXAMPLE
				PS C:\> New-PaymentSpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-PaymentSpreadsheet
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Payment Spreadsheet ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "payments.csv");
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{
				'border' = $true;
				'values' = @(
					$entry.'Team Name:',
					$entry.'Team State:',
					$entry.'Division',
					$entry.'Team Manager:',
					$entry.'Team Manager Email:',
					$entry.'Payment due (AUD)',
					$entry.'Direct Debit Receipt')
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Team Name", "State", "Division", "Team Manager", "Team Manager Email", "Payment Due (AUD)", "Direct Debit Receipt")
		New-SpreadSheet -name "Payments" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no payments found."
	}
}

<#
	.SYNOPSIS
		A brief description of the New-CoachSkatersList function.
	
	.DESCRIPTION
		A detailed description of the New-CoachSkatersList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (XLSX, CSV)
	
	.EXAMPLE
				PS C:\> New-CoachSkatersList
	
	.NOTES
		Additional information about the function.
#>
function New-CoachSkatersList
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Coach/Skaters List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "coach_skaters.${format}")
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$hash = @{ }
	foreach ($entry in $entries)
	{
		$coach_name = $entry.'Primary Coach Name:'
		$coach_email = $entry.'Primary Coach E-mail:'
		
		if (!$hash.ContainsKey($coach_name))
		{
			$hash[$coach_name] = @()
		}
		$hash[$coach_name] += $entry.'Skater 1 Name' + " ($($entry.Division))"
		if (![String]::IsNullOrEmpty($entry.'Skater 1 Name'))
		{
			$hash[$coach_name] += $entry.'Skater 2 Name' + " ($($entry.Division))"
		}
	}
	
	$rows = @()
	foreach ($coach_name in $hash.Keys | Sort-Object)
	{
		foreach ($skater in $hash[$coach_name])
		{
			$rows += (@{ 'border' = $true; 'values' = @($coach_name, $skater) })
		}
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @('Coach Name', 'Skater')
		New-Spreadsheet -name "Coach Skaters" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no coaches found."
	}
}

<#
	.SYNOPSIS
		Create Photo Permission List
	
	.DESCRIPTION
		A detailed description of the New-PhotoPermissionList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (XLSX, CSV)
	
	.EXAMPLE
		PS C:\> New-PhotoPermissionList
	
	.NOTES
		Additional information about the function.
#>
function New-PhotoPermissionList
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Photo Permission List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "photo_permissions.${format}");
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{
				'border' = $true;
				'values' = @($entry.'Division',
					$entry.'Skater 1 Name',
					$entry.'Skater 2 Name',
					$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
				)
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Division", "Skater 1 Name", "Skater 2 Name", "ACTISA granted permission to use photos")
		New-Spreadsheet -name "Photo Permissions" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no entries found."
	}
}

<#
	.SYNOPSIS
		A brief description of the New-TeamDescriptionSpreadsheet function.
	
	.DESCRIPTION
		A detailed description of the New-TeamDescriptionSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (XLSX, CSV)
	
	.EXAMPLE
		PS C:\> New-TeamDescriptionSpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-TeamDescriptionSpreadsheet
{
	param
	(
		$entries,
		[string]$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Team Descriptions"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "team_descriptions.${format}");
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{
				'border' = $true;
				'values' = @(
					$entry.'Team Name:',
					$entry.'Team State:',
					$entry.'Division',
					$entry.'Team Program Written Description (maximum 50 words):')
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Team Name", "State", "Division", "Description")
		New-SpreadSheet -name "Payments" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no teams found."
	}
}

<#
	.SYNOPSIS
		A brief description of the New-TeamMemberSpreadsheet function.
	
	.DESCRIPTION
		A detailed description of the New-TeamMemberSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		Format of Spreadsheet (XLSX, CSV)
	
	.EXAMPLE
				PS C:\> New-TeamMemberSpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-TeamMemberSpreadsheet
{
	param
	(
		$entries,
		$folder,
		[string]$format = 'csv'
	)
	
	Write-Host "Generating Team Members Spreadsheet"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "team_members.${format}")
	
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$team_name = $entry.'Team Name:'
		$division = $entry.Division
		$skaters = Get-EntrySkaterNames -entry $entry
		
		$rows += (@{ 'border' = $true; 'values' = @($team_name, $division, $skaters.Count, ($skaters -join "`r`n")) })
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @('Team Name', 'Division', 'Number of Members', 'Members')
		New-Spreadsheet -name "Team Members" -path $filepath -headers $headers -rows $rows -format $format
	}
	else
	{
		Write-Host " - no teams found."
	}
}

#================================================================================
#------------------------          MAIN CONTROL          ------------------------
#================================================================================

# prompt the user to specify location
$comp_folder = Find-Folders -title "Select the Competition folder (default=$comp_folder)" -default $comp_folder
$template_folder = Find-Folders -title "Select the MailMerge Template folder (default=$template_folder)" -default $template_folder

Push-Location $comp_folder

foreach ($f in ('Submissions', 'Music', 'Certificates', 'Schedule'))
{
	if ((Test-Path $f -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $f -ItemType Directory | Out-Null
	}
}

Pop-Location

$submissionFullPath = [System.IO.Path]::Combine($comp_folder, "Submissions")
$music_folder = [System.IO.Path]::Combine($comp_folder, "Music")
$certificate_folder = [System.IO.Path]::Combine($comp_folder, "Certificates")
$schedule_folder = [System.IO.Path]::Combine($comp_folder, "Schedule")

Write-Host "Competition Folder: $comp_folder"
write-host "Music Folder: $music_folder"

$entries = &Get-SubmissionEntries

foreach ($entry in $entries)
{
	Publish-EntryMusicFiles -entry $entry -submissionFullPath $submissionFullPath -music_folder $music_folder
}

Write-Host "Number of entries = $($entries.Count)`n" -ForegroundColor Yellow

New-TeamMemberSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-CertificateList -entries $entries -folder $certificate_folder
New-SkatingSchedule -entries $entries -folder $schedule_folder
New-DivisionCountsSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-EngravingSchedule -entries $entries -folder $comp_folder
New-RegistrationList -entries $entries -folder $comp_folder -format 'xlsx'
New-VolunteerSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-PaymentSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-TeamDescriptionSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-TeamEmailList -entries $entries -folder $comp_folder -format 'xlsx'
New-CoachSkatersList -entries $entries -folder $comp_folder -format 'xlsx'
New-PhotoPermissionList -entries $entries -folder $comp_folder -format 'xlsx'
#Read-Host -Prompt "Press Enter to exit
}
<#
	.SYNOPSIS
		Processes JotForm submissions for Reg Park Artistic (RPA) Competition
	
	.DESCRIPTION
		Processes JotForm submissions for Reg Park Artistic (RPA) Competition
	
	.PARAMETER prompt
		if the user should be prompted for folder/file locations

	.EXAMPLE
			process_rpa_submissions.ps1
			process_rpa_submissions.ps1 -prompt $true
			process_rpa_submissions.ps1 -prompt $false
	
	.NOTES
		===========================================================================
		Created on:   	22/08/2019 10:59 AM
		Created by:   	Aaron Duncan
		Organization: 	ACTISA
		Filename:     	process_rpa_submissions.ps1
		===========================================================================
#>
param
(
	[bool]
	$prompt = $true
)

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

$year = Get-Date -Format yyyy

# the '2019 Reg Park Artistic Registration' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTDWmR2TY3KSykfVKjx76ngEXgg23vtFSyxui8GzYm894b5_QlE_e2kpQqRUcHsYha_22jOURV271ps/pub?output=csv'

$template_folder = 'D:\Skating Templates';

$Competition = "RPA $year";

$RP_Engraving_Title = "Reg Park Artistic $year"
$AS_Engraving_Title = "Manzano Aussie Skate Artistic $year"

$RP_CertificateTemplate = "Certificate - RPA ${year}.docx"
$AS_CertificateTemplate = "Certificate - Manzano ${year}.docx"

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
		Place the music file in the correct location, with the correctly formatted name.
	
	.PARAMETER filename
		A description of the filename parameter.
	
	.PARAMETER category
		A description of the category parameter.
	
	.PARAMETER division
		A description of the division parameter.
	
	.PARAMETER gender
		A description of the gender parameter.
	
	.PARAMETER skatername
		A description of the skatername parameter.
	
	.PARAMETER destination
		A description of the destination parameter.
	
	.PARAMETER program
		A description of the program parameter.
	
	.EXAMPLE
				PS C:\> Publish-SkaterMusicFile
	
	.NOTES
		Additional information about the function.
#>
function Publish-SkaterMusicFile
{
	param
	(
		$filename,
		$category,
		$division,
		$gender,
		$entrantName,
		$destination,
		$program
	)
	
	Write-Host "Publishing music files for $entrantName (division: $division)."
	
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
		Write-Warning " - invalid Music File: $filename"
	}
	
	#
	# determine the destination music folder
	#
	$music_subdir = "$category - $division"
	
	if ($gender -ne $null)
	{
		if ($gender.Equals("Female"))
		{
			$music_subdir += " Ladies"
		}
		else
		{
			$music_subdir += " Mens"
		}
	}
	
	$music_dest = [System.IO.Path]::Combine($destination, $music_subdir)
	
	#
	# Calculate the new music filename
	#
	
	$new_music_file = $division -replace "\(.*\)", ""
	
	if ($category.StartsWith("Aussie Skate"))
	{
		$new_music_file = "AS_${new_music_file}"
	}
	elseif ($category -match "Adult|Dance")
	{
		$extra = $category -replace "\(.*\)", ""
		$new_music_file = "${extra}${new_music_file}"
	}
	
	foreach ($key in $abbreviations.Keys)
	{
		$new_music_file = $new_music_file -replace $key, $abbreviations.Item($key)
	}
	
	$new_music_file = $new_music_file -replace " ; ", "_" -replace " ", ""
	
	if ($category -match 'Singles' -and $division -match 'Advanced Novice|Junior|Senior')
	{
		if ($gender.Equals("Female"))
		{
			$new_music_file += "Ladies"
		}
		else
		{
			$new_music_file += "Men"
		}
	}
	
	if ((! $category.StartsWith("Aussie")) -and (! [String]::IsNullOrEmpty($program)))
	{
		$new_music_file += "_${program}"
	}
	
	$new_music_file += "_$entrantName"
	
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
			Write-Warning " - failed to copy '$filename' -> '$new_music_path'"
		}
		else
		{
			#"Source Music File: $filename"
			#"Music Destination: $new_music_path"
			#""
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
	$catdiv = $entry.'Division:'
	if ([String]::IsNullOrEmpty($catdiv))
	{
		Write-Warning "Failed to retrieve Category/Division for $($entry.'Skater 1 Name')"
	}
	else
	{
		[String[]]$list = $catdiv.Split('-', 2)
		if ($list.Count -lt 1)
		{
			Write-Warning "Failed to extract category from '$catdiv'"
		}
		elseif ($list.Count -lt 2)
		{
			Write-Warning "Extracted Category, but failed to extract Division from '$catdiv'"
		}
		else
		{
			[string]$category = $list[0].trim()
			[string]$division = $list[1].trim()
			
			[string]$gender = $null
			[string]$music_url = $entry.'Music File'
			
			if ($division.Contains('Group'))
			{
				$name = $entry.'Group Name'
			}
			elseif ($division -match 'Couple')
			{
				$name = $entry.'Skater 1 Name' + "_" + $entry.'Skater 2 Name';
			}
			else
			{
				$name = ConvertTo-CapitalizedName -name $entry.'Skater 1 Name'
				$gender = $entry.'Skater 1 Gender'
			}
			
			$submission_folder =
			(Resolve-Path "$([System.IO.Path]::Combine($submissionFullPath, "${submission_id}_${name}"))*").Path
			
			if ([string]::IsNullOrEmpty($submission_folder))
			{
				$submission_folder = [System.IO.Path]::Combine($submissionFullPath, "${submission_id}_${name}-")
			}
			
			# half ice divisions don't have music file uploads
			if (!$category.StartsWith("Aussie Skate (Half"))
			{
				$music_file = [System.Web.HttpUtility]::UrlDecode($music_url.Split("/")[-1]);
				$extension = [System.IO.Path]::GetExtension($music_file)
				$music_fullpath = [System.IO.Path]::Combine($submission_folder, $music_file) -replace '([][[])', ''
				
				if ((Test-Path -Path $submission_folder -ErrorAction SilentlyContinue) -eq $false)
				{
					New-Item $submission_folder -Type Directory | Out-Null
				}
				
				if ((Test-Path -Path $music_fullpath -ErrorAction SilentlyContinue) -eq $false)
				{
					# music file is missing, so download it
					Get-WebFile -url $entry.'Music File' -destination $music_fullpath
				}
				
				if ($category -match 'Dance')
				{
					Publish-SkaterMusicFile -filename $music_fullpath -category $category -division $division -entrantName $name -gender $gender -destination $music_folder -program "FD"
				}
				else
				{
					Publish-SkaterMusicFile -filename $music_fullpath -category $category -division $division -entrantName $name -gender $gender -destination $music_folder -program "FS"
				}
			}
		}
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
	
	if (-not [String]::IsNullOrWhiteSpace($entry.'Skater 1 Name'))
	{
		$members += $entry.'Skater 1 Name'
	}
	
	if (-not [String]::IsNullOrWhiteSpace($entry.'Skater 2 Name'))
	{
		$members += $entry.'Skater 2 Name'
	}
	
	if ($entry)
	{
		if ($entry.'Division:'.Contains("Group"))
		{
			for ($i = 1; $i -le 24; $i++)
			{
				$first = $entry."Type a question >> Skater $i >> First Name"
				$last = $entry."Type a question >> Skater $i >> Last Name"
				if (-not ([String]::IsNullOrEmpty($first) -or [String]::IsNullOrEmpty($first)))
				{
					$members += ConvertTo-CapitalizedName -name "$first $last"
				}
			}
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
		Generate the list of skater certificates
	
	.DESCRIPTION
		Constructs a list of skaters/divisions, and uses mailmerge to create a pdf list of certificates using word templates (which this function interactively prompts the user for).
	
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
		[string]$folder
	)
	
	Write-Host "Generating Certificates"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	if ($prompt)
	{
		$RP_template = Find-Template -message "Select Reg Park Artistic Certificate Template" -initial_dir $template_folder -default $RP_CertificateTemplate
		$AS_template = Find-Template -message "Select Manzano Aussie Skate Artistic Certificate Template" -initial_dir $template_folder -default $AS_CertificateTemplate
	}
	else
	{
		$RP_template = Resolve-Path -Path "${template_folder}/${RP_CertificateTemplate}"
		$AS_template = Resolve-Path -Path "${template_folder}/${AS_CertificateTemplate}"
	}
	
	$RP_input_csv = [System.IO.Path]::Combine($folder, "RP_certificate_inputs.csv")
	$AS_input_csv = [System.IO.Path]::Combine($folder, "AS_certificate_inputs.csv")
	
	#Write-Host "Input CSV file: $input_csv"
	
	$AS_results = @()
	$RP_results = @()
	
	$entries | ForEach-Object {
		$category = $_.'Division:'.Split("-", 2)[0].trim()
		$division = $_.'Division:'.Split("-", 2)[1].trim()
		
		if ($division.Contains("Aussie Skate"))
		{
			foreach ($name in (Get-EntrySkaterNames -entry $_))
			{
				#Write-Host "Generating certificate for $name ($division)"
				$AS_results += New-Object -TypeName PSObject -Property @{
					"Name" = ConvertTo-CapitalizedName -name $name
					"Division" = $division
				}
			}
		}
		else
		{
			foreach ($name in (Get-EntrySkaterNames -entry $_))
			{
				#Write-Host "Generating certificate for $name ($division)"
				$RP_results += New-Object -TypeName PSObject -Property @{
					"Name" = ConvertTo-CapitalizedName -name $name
					"Division" = $division
				}
			}
		}
	}
	
	#Write-Host "Number of entries = $($entries.Count)"
	
	$AS_results | Select-Object "Name", "Division" | export-csv -path $AS_input_csv -Force -NoTypeInformation
	$RP_results | Select-Object "Name", "Division" | export-csv -path $RP_input_csv -Force -NoTypeInformation
	
	Write-Host " - merging Reg Park Artistic certificates"
	Invoke-MailMerge -template $RP_template -datasource $RP_input_csv -destination $folder
	Write-Host " - merging Manzano Artistic certificates"
	Invoke-MailMerge -template $AS_template -datasource $AS_input_csv -destination $folder
}

<#
	.SYNOPSIS
		Generate Skating Schedule Templates
	
	.DESCRIPTION
		Generates initial spreadsheets to be used to construct Skating Schedules.
		Two separate spreadsheets are constructed.
		All performances will go in the first spreadsheet, unless there is a requirement for two performances to occur on separate days; in this case, the second performance will be placed in the second spreadsheet.
	
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
		$division = $entry.'Division:'
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	#$divhash.GetEnumerator() | Sort Name
	
	foreach ($div in $divhash.Keys | Sort)
	{
		$category = $div.Split("-", 2)[0].trim()
		$division = $div.Split("-", 2)[1].trim()
		
		# NOTE: need to put this in a hash based on category and division
		# Perhaps something like have a category number, which is overridden by division
		if ($category.StartsWith("XXX"))
		{
			$warmup = 3
			$performance_time = 2
		}
		else
		{
			$warmup = 3
			$performance_time = 4
		}
		
		#Write-Host "Category: $category, Division: $division"
		Add-Content -Path $schedule1 -Value "Category: $category, Division: $division"
		
		$count = $divhash.Item($div).Count
		$num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
		Add-Content -path $schedule1 -Value "Num Entries: $count"
		"Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes" -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule1
		"Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time) | Add-Content -Path $schedule1
		
		if ($division.Contains("Group"))
		{
			Add-Content -path $schedule1 -Value "Group Name,, State, Coach Name, Other Coach Names, Music Title"
		}
		else
		{
			Add-Content -path $schedule1 -Value "First Name, Last Name, State, Coach Name, Other Coach Names, Music Title"
		}
		
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
			
			if ($division.Contains("Group"))
			{
				"`"{0}`",, `"{1}`", `"{2}`", `"{3}`", `"{4}`"" -f $_.'Group Name', $_.'Skater 1 State/Territory', $_.'Primary Coach Name:', $_.'Other Coach Names', $music_title.Trim() | Add-Content -path $schedule1
			}
			else
			{
				"`"{0}`", `"{1}`", `"{2}`", `"{3}`", `"{4}`", `"{5}`"" -f $_.'First Name', $_.'Last Name', $_.'Skater 1 State/Territory', $_.'Primary Coach Name:', $_.'Other Coach Names', $music_title.Trim() | Add-Content -path $schedule1
			}
		}
		Add-Content -Path $schedule1 -Value ""
	}
}

<#
	.SYNOPSIS
		Create Count Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-DivisionCountsSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		A description of the format parameter.
	
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
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Division Counts Spreadsheet ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filepath = [System.IO.Path]::Combine($folder, "division_counts.${format}")
	if (Test-Path $filepath) { Remove-Item -Path $filepath }
	
	$gender = @{ Female = "Ladies"; Male = "Men" }
	
	$divhash = @{ }
	foreach ($entry in $entries)
	{
		$division = $entry.'Division:'
		
		if ([String]::IsNullOrEmpty($division))
		{
			Write-Warning " - failed to retrieve division from entry."
		}
		else
		{
			if ($division -notmatch "Aussie Skate" -and $division -notmatch "Dance")
			{
				try
				{
					$division += " " + $gender[$entry.'Skater 1 Gender']
				}
				catch
				{
					Write-Warning " - failed to get gender for $($entry.'Skater 1 Name')"
				}
			}
			
			if (!$divhash.ContainsKey($division))
			{
				$divhash[$division] = @()
			}
			$divhash[$division] += $entry
		}
	}
	
	$rows = @()
	
	$trophy_count = @{ gold = 0; silver = 0; bronze = 0 }
	
	foreach ($div in $divhash.Keys | Sort-Object)
	{
		[String[]]$list = $div.Split('-', 2)
		if ($list.Count -lt 1)
		{
			Write-Warning "Failed to extract category from '$catdiv'"
		}
		elseif ($list.Count -lt 2)
		{
			Write-Warning "Extracted Category, but failed to extract Division from '$catdiv'"
		}
		else
		{
			$category = $list[0].trim()
			$division = $list[1].trim()
			
			if ($category -match "(Adult | Dance)")
			{
				#$division = "$category $division"
			}
			
			if ($category -match 'Couple')
			{
				$numSkaters = 2
			}
			else
			{
				$numSkaters = 1
			}
			
			$count = $divhash.Item($div).Count
			$trophy_gold = ""
			$trophy_silver = ""
			$trophy_bronze = ""
			
			if ($count -ge 1) { $trophy_gold = $numSkaters; $trophy_count.gold += $numSkaters }
			if ($count -ge 2) { $trophy_silver = $numSkaters; $trophy_count.silver += $numSkaters }
			if ($count -ge 3) { $trophy_bronze = $numSkaters; $trophy_count.bronze += $numSkaters }
			
			$rows += (@{ 'border' = $true; 'values' = @($category, $division, $count, $trophy_bronze, $trophy_silver, $trophy_gold) })
		}
	}
	
	$rows += @(@{
			border = $false;
			values = @('', '', 'TOTAL:', $trophy_count.bronze, $trophy_count.silver, $trophy_count.gold)
		},
		@{ border = $false; values = @('') },
		@{ border = $false; values = @('', 'Manzano Artistic Trophy') },
		@{ border = $false; values = @('', 'Reg Park Artistic Trophy') }
	)
	
	$headers = @('Category', 'Division', '# Entries', 'Bronze Trophy', 'Silver Trophy', 'Gold Trophy')
	New-SpreadSheet -name "Division Counts" -path $filepath -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create Trophy/Medal engraving schedules
	
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
	
	$trophyPath = [System.IO.Path]::Combine($folder, "${Competition} - TROPHIES.docx")
	
	if (Test-Path $trophyPath) { Remove-Item -Path $trophyPath }
	
	$divhash = @{ }
	foreach ($entry in $entries)
	{
		$division = $entry.'Division:'
		
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	$cellAlignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
	
	# Create Trophy Table
	$TrophyWord = New-Object -ComObject Word.Application
	$TrophyWord.Visible = $False
	$TrophyDoc = $TrophyWord.Documents.Add()
	$margin = 24 # 1.26 cm
	$TrophyDoc.PageSetup.LeftMargin = $margin
	$TrophyDoc.PageSetup.RightMargin = $margin
	$TrophyDoc.PageSetup.TopMargin = $margin
	$TrophyDoc.PageSetup.BottomMargin = $margin
	$TrophyDoc.PageSetup.Orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape
	$Selection = $TrophyWord.Selection
	$TrophyTable = $Selection.Tables.Add($Selection.Range, 2, 4,
		[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior,
		[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent)
	$TrophyTable.Range.Style = "No Spacing"
	$TrophyTable.Range.paragraphFormat.Alignment = $cellAlignment
	$TrophyTable.Borders.Enable = $True
	$TrophyTable.Rows(2).Cells(1).Range.Text = "Division"
	$TrophyTable.Rows(2).Cells(1).Range.Bold = $True
	#$TrophyTable.Rows(2).Cells(1).PreferredWidthType = 3
	#$TrophyTable.Rows(2).Cells(1).Width = 16
	
	# Create the title row on each table
	$Row = $TrophyTable.Rows(1)
	$Row.Cells.Merge()
	$Row.Cells(1).Range.Text = "$Competition engraving schedule - TROPHY"
	$Row.Cells(1).Range.Bold = $true
	#$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
	
	foreach ($div in $divhash.Keys | Sort)
	{
		$category = $div.Split("-", 2)[0].trim()
		$division = $div.Split("-", 2)[1].trim()
		
		if ($division.Contains("Aussie Skate"))
		{
			$Engraving_Title = $AS_Engraving_Title
		}
		else
		{
			$Engraving_Title = $RP_Engraving_Title
		}
		
		$divEntries = $divhash.Item($div)
		$count = $divEntries.Count
		$maxTeamSize = Get-MaxEntryNumberOfSkaters -entries $divEntries
		
		$Row = $TrophyTable.Rows.Add()
		#$Row.Cells.PreferredWidthType = 2
		#$Row.Cells.Width = 28
		
		if ($maxTeamSize -gt 1)
		{
			$Row.Cells(1).Range.Text = "$division`r`r($maxTeamSize of each trophy)"
		}
		else
		{
			$Row.Cells(1).Range.Text = $division
		}
		#$Row.Cells(1).PreferredWidthType = 2
		#$Row.Cells(1).Width = 28
		
		$Row.Cells(1).Range.Bold = $True
		$Row.Cells(1).Range.Font.Spacing = 1.0
		#$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
		
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
		Create list of skater email addresses
	
	.DESCRIPTION
		A detailed description of the New-SkaterEmailList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.EXAMPLE
				PS C:\> New-SkaterEmailList
	
	.NOTES
		Additional information about the function.
#>
function New-SkaterEmailList
{
	param
	(
		$entries,
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Skater Email List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filename = [System.IO.Path]::Combine($folder, "skater_email_list.${format}")
	
	$list = @{ }
	foreach ($entry in $entries)
	{
		$name = $entry.'Skater 1 Name'
		$email = $entry.'Skater 1 Contact E-mail:'
		if (-not $list.ContainsKey($name))
		{
			$list.Add($name, $email)
		}
	}
	
	$headers = @('Name', 'E-Mail')
	$rows = @()
	foreach ($name in $list.Keys)
	{
		$rows += (@{ 'border' = $true; 'values' = @($name, $list[$name]) })
	}
	New-SpreadSheet -name "Skater Email List" -path $filename -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create list of coach email addresses
	
	.DESCRIPTION
		A detailed description of the New-CoachEmailList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
	
	$emails = @{}
	foreach ($entry in $entries)
	{
		$emails[$entry.'Primary Coach E-mail:'] = $entry.'Primary Coach Name:'
	}
	
	$headers = @('Coach Name', 'Coach E-Mail')
	$rows = @()
	$emails.GetEnumerator() | ForEach-Object {
		$rows += (@{ 'border' = $true; 'values' = @($_.value, $_.key) })
	}
	
	New-Spreadsheet -name "Coach E-Mail List" -path $filename -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create Registration List
	
	.DESCRIPTION
		This creates a list of skaters, for use at the front door to "check off" skaters as they arrive.
		This currently prints out an option for fruit/water but could be modified to include things like "has been given a goodie bag" etc.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Skater Registration List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$filename = [System.IO.Path]::Combine($folder, "registration_list.${format}")
	
	$hash = @{ }
	foreach ($entry in $entries)
	{
		$surname = ConvertTo-CapitalizedName -name $entry.'Last Name'
		if (!$hash.ContainsKey($surname))
		{
			$hash[$surname] = @{ }
		}
		$firstname = ConvertTo-CapitalizedName -name $entry.'First Name'
		
		if (-not $hash[$surname].ContainsKey($firstname))
		{
			$hash[$surname].Add($firstname, $true)
		}
		
		if ($entry.Division -match 'Dance')
		{
			$surname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name: (Last Name)'
			if (!$hash.ContainsKey($surname))
			{
				$hash[$surname] = @{ }
			}
			$firstname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name: (First Name)'
			if (-not $hash[$surname].ContainsKey($firstname))
			{
				$hash[$surname].Add($firstname, $true)
			}
		}
	}
	
	$headers = @('Skater Surname', 'Skater First Name', 'Water', 'Fruit')
	$rows = @()
	foreach ($surname in $hash.Keys | Sort)
	{
		$name = $hash.Item($surname)
		foreach ($firstname in $name.Keys | sort)
		{
			$rows += (@{ 'border' = $true; 'values' = @($surname, $firstname, '', '') })
		}
	}
	New-SpreadSheet -name "Skater Registration List" -path $filename -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create Volunteer Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-VolunteerSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Volunteer Spreadsheet ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$spreadsheet = [System.IO.Path]::Combine($folder, "volunteer.${format}");
	
	$headers = @("Name", "Division", "Volunteer Name", "Volunteer E-mail", "Volunteer Phone", "Availability", "Roles", "Other Notes")
	$rows = @()
	foreach ($entry in $entries)
	{
		if (-not [String]::IsNullOrEmpty($entry.'I am able to assist with the following tasks:'))
		{
			$catdiv = $entry.'Division:'
			if ([String]::IsNullOrEmpty($catdiv))
			{
				Write-Warning "Failed to retrieve Category/Division for $($entry.'Skater 1 Name')"
			}
			else
			{
				[String[]]$list = $catdiv.Split('-', 2)
				if ($list.Count -lt 1)
				{
					Write-Warning "Failed to extract category from '$catdiv'"
				}
				elseif ($list.Count -lt 2)
				{
					Write-Warning "Extracted Category, but failed to extract Division from '$catdiv'"
				}
				else
				{
					$category = $list[0].trim()
					$division = $list[1].trim()
					
					if ($category -eq "Adult")
					{
						$division = "Adult ${division}"
					}
					
					$rows += (@{
							'border' = $true;
							'values' = @($entry.'Skater 1 Name',
								$division,
								$entry.'Volunteer Name',
								$entry.'Volunteer E-mail',
								$entry.'Volunteer Contact Mobile',
								$entry.'Availability:',
								$entry.'I am able to assist with the following tasks:',
								$entry.'Other Notes:')
						})
				}
			}
		}
	}
	New-Spreadsheet -name "Volunteers" -path $spreadsheet -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create Payment Spreadsheet
	
	.DESCRIPTION
		A detailed description of the New-PaymentSpreadsheet function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
	
	$filepath = [System.IO.Path]::Combine($folder, "payments.${format}");
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$catdiv = $entry.'Division:'
		if ([String]::IsNullOrEmpty($catdiv))
		{
			Write-Warning "Failed to retrieve Category/Division for $($entry.'Skater 1 Name')"
		}
		else
		{
			[String[]]$list = $catdiv.Split('-', 2)
			if ($list.Count -lt 1)
			{
				Write-Warning "Failed to extract category from '$catdiv'"
			}
			elseif ($list.Count -lt 2)
			{
				Write-Warning "Extracted Category, but failed to extract Division from '$catdiv'"
			}
			else
			{
				$category = $list[0].trim()
				$division = $list[1].trim()
				
				if ($category -eq "Adult")
				{
					$division = "Adult ${division}"
				}
				
				$parent_firstname = $entry.'Parent/Guardian Name: (First Name)'
				if ([String]::IsNullOrEmpty($parent_firstname))
				{
					$parentname = ""
				}
				else
				{
					$parent_lastname = $entry.'Parent/Guardian Name: (Last Name)'
					$parentname = "${parent_firstname} ${parent_lastname}"
				}
				
				if ($division.Contains("Group"))
				{
					$skater_group_name = $entry.'Group Name'
				}
				else
				{
					$skater_group_name = $entry.'Skater 1 Name'
				}
				$rows += (@{
						'border' = $true;
						'values' = @(
							$skater_group_name,
							$division,
							$parentName,
							$entry.'Payment due (AUD)',
							$entry.'Direct Debit Receipt')
					})
			}
		}
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Skater/Group Name", "Division", "Parent/Guardian (if applicable)", "Payment Due (AUD)", "Direct Debit Receipt")
		New-SpreadSheet -name "Payments" -path $filepath -headers $headers -rows $rows -format $format
	}
}

<#
	.SYNOPSIS
		Create List of Coaches, and their skaters
	
	.DESCRIPTION
		A detailed description of the New-CoachSkatersList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
	
	$outfile = [System.IO.Path]::Combine($folder, "coach_skaters.${format}")
	
	if (Test-Path $outfile) { Remove-Item -Path $outfile }
	
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
	
	$headers = @('Coach Name', 'Skater')
	$rows = @()
	
	foreach ($coach_name in $hash.Keys | Sort)
	{
		foreach ($skater in $hash[$coach_name])
		{
			$rows += (@{ 'border' = $true; 'values' = @($coach_name, $skater) })
		}
	}
	New-Spreadsheet -name "Coach Skaters" -path $outfile -headers $headers -rows $rows -format $format
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
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Photo Permission List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$outfile = [System.IO.Path]::Combine($folder, "photo_permissions.${format}");
	
	$headers = @("Category", "Division", "Skater 1 Name", "Skater 2 Name", "ACTISA granted permission to use photos")
	$rows = @()
	foreach ($entry in $entries)
	{
		$catdiv = $entry.Division.Split("; ")
		$category = $catdiv[0].trim()
		$division = $catdiv[1].trim()
		$rows += (@{
				'border' = $true;
				'values' = @($category,
					$division,
					$entry.'Skater 1 Name',
					$entry.'Skater 2 Name',
					$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
				)
			})
	}
	New-Spreadsheet -name "Photo Permissions" -path $outfile -headers $headers -rows $rows -format $format
}

<#
	.SYNOPSIS
		Create List of Coaches, and their skaters
	
	.DESCRIPTION
		A detailed description of the New-CoachSkatersList function.
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
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
	
	Write-Host "Generating Coach Skater/Group List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$outfile = [System.IO.Path]::Combine($folder, "coach_skaters_groups.${format}")
	if (Test-Path $outfile) { Remove-Item -Path $outfile }
	
	$hash = @{ }
	foreach ($entry in $entries)
	{
		$coach_name = $entry.'Primary Coach Name:'
		$coach_email = $entry.'Primary Coach E-mail:'
		$division = $entry.'Division:'
		
		if ($division -ne $null)
		{
			if (!$hash.ContainsKey($coach_name))
			{
				$hash[$coach_name] = @()
			}
			if ($division.Contains('Group'))
			{
				$hash[$coach_name] += @{ name = $entry.'Group Name'; division = $division }
			}
			else
			{
				$hash[$coach_name] += @{ name = $entry.'Skater 1 Name'; division = $division }
				if (![String]::IsNullOrEmpty($entry.'Skater 2 Name'))
				{
					$hash[$coach_name] += @{ name = $entry.'Skater 2 Name'; division = $division }
				}
			}
		}
	}
	
	$headers = @('Coach Name', 'Skater', 'Division')
	$rows = @()
	
	foreach ($coach_name in $hash.Keys | Sort-Object)
	{
		foreach ($entrant in $hash[$coach_name])
		{
			$rows += (@{ 'border' = $true; 'values' = @($coach_name, $entrant['name'], $entrant['division']) })
		}
	}
	New-Spreadsheet -name "Coach Skaters" -path $outfile -headers $headers -rows $rows -format $format
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
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating Photo Permission List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$outfile = [System.IO.Path]::Combine($folder, "photo_permissions.${format}");
	if (Test-Path $outfile) { Remove-Item -Path $outfile }
	
	$rows = @()
	foreach ($entry in $entries)
	{
		$catdiv = $entry.'Division:'
		if ($catdiv -eq $null)
		{
			Write-Host " - failed to split division with the '-' character"
		}
		else
		{
			[string[]]$list = $catdiv.Split('-', 2)
			if ($list.Count -lt 1)
			{
				Write-Warning "Failed to extract category from '$catdiv'"
			}
			elseif ($list.Count -lt 2)
			{
				Write-Warning "Extracted Category, but failed to extract Division from '$catdiv'"
			}
			else
			{
				$category = $list[0].trim()
				$division = $list[1].trim()
				
				if ($division.Contains('Group'))
				{
					$rows += (@{
							'border' = $true;
							'values' = @($category,
								$division,
								$entry.'Group Name',
								$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
							)
						})
				}
				elseif ($division.Contains('Couple'))
				{
					$rows += (
						@{
							'border' = $true;
							'values' = @($category,
								$division,
								$entry.'Skater 1 Name',
								$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
							)
						},
						@{
							'border' = $true;
							'values' = @($category,
								$division,
								$entry.'Skater 2 Name',
								$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
							)
						}
					)
				}
				else
				{
					$rows += (@{
							'border' = $true;
							'values' = @($category,
								$division,
								$entry.'Skater 1 Name',
								$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
							)
						})
					}
				}
			}
		}
		if ($rows.Count -gt 0)
	{
		$headers = @("Category", "Division", "Skater/Group Name", "ACTISA granted permission to use photos")
		New-Spreadsheet -name "Photo Permissions" -path $outfile -headers $headers -rows $rows -format $format
	}
}

function New-ProofOfAgeAndMemberships
{
	[CmdletBinding()]
	param
	(
		$entries,
		$folder,
		$format = 'csv'
	)
	
	Write-Host "Generating POA/Memberships List ($format)"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$outfile = [System.IO.Path]::Combine($folder, "poa_and_memberships.${format}")
	
	if (Test-Path $outfile) { Remove-Item -Path $outfile -ErrorAction SilentlyContinue }
	
	$list = @{ }
	foreach ($entry in $entries)
	{
		$catdiv = $entry.Division.Split(";")
		$category = $catdiv[0].trim()
		$division = $catdiv[1].trim()
		
		$name = $entry.'Skater 1 Name'
		if (-not $list.ContainsKey($name))
		{
			Write-Host "Name1: '$name'"
			$list.Add($name,
				@(
					$name,
					$entry.'Skater 1 State/Territory:',
					$category,
					$division,
					$entry.'Primary Coach Name:',
					$entry.'Other Coach Names:',
					$entry.'Skater 1 Membership Number:'
					$entry.'Skater 1 Proof Of Age (POA):'
				))
		}
		
		$name = $entry.'Skater 2 Name'
		if (![string]::IsNullOrEmpty($name))
		{
			if (-not $list.ContainsKey($name))
			{
				Write-Host "Name2: '$name'"
				$list.Add($name,
					@(
						$name,
						$entry.'Skater 2 State/Territory:',
						$category,
						$division,
						$entry.'Primary Coach Name:',
						$entry.'Other Coach Names:',
						$entry.'Skater 2 Membership Number:'
						$entry.'Skater 2 Proof Of Age (POA):'
					))
			}
		}
		
	}
	
	$headers = @('Name', 'State', 'Category', 'Division', 'Primary Coach', 'Other Coaches', 'Membership #', 'POA')
	$rows = @()
	foreach ($entry in $list.Values)
	{
		$rows += (@{ 'border' = $true; 'values' = $entry })
	}
	New-SpreadSheet -name "POA and Membership Numbers" -path $outfile -headers $headers -rows $rows -format $format
}

#================================================================================
#------------------------          MAIN CONTROL          ------------------------
#================================================================================

if ($prompt)
{
	# prompt the user to specify location
	$comp_folder = Find-Folders -title "Select the Competition folder" -default $comp_folder
	$template_folder = Find-Folders -title "Select the MailMerge Template folder" -default $template_folder
}
else
{
	if (!(Test-Path -Path $comp_folder -ErrorAction SilentlyContinue))
	{
		New-Item -ItemType Directory -Force -Path $comp_folder | Out-Null
	}
	if (!(Test-Path -Path $template_folder -ErrorAction SilentlyContinue))
	{
		New-Item -ItemType Directory -Force -Path $template_folder | Out-Null
	}
}

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

$entries = @()
foreach ($entry in (Get-SubmissionEntries -url $google_sheet_url))
{
	# strip out entries with no submission ID
	if (-not [String]::IsNullOrWhiteSpace($entry.'Submission ID'))
	{
		$entries += $entry
	}
}

foreach ($entry in $entries)
{
	Publish-EntryMusicFiles -entry $entry -submissionFullPath $submissionFullPath -music_folder $music_folder
}

Write-Host "Number of entries = $($entries.Count)`n" -ForegroundColor Yellow

New-CertificateList -entries $entries -folder $certificate_folder
New-SkatingSchedule -entries $entries -folder $schedule_folder
New-DivisionCountsSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-EngravingSchedule -entries $entries -folder $comp_folder
New-RegistrationList -entries $entries -folder $comp_folder -format 'xlsx'
New-VolunteerSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-PaymentSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-SkaterEmailList -entries $entries -folder $comp_folder -format 'xlsx'
New-CoachEmailList -entries $entries -folder $comp_folder -format 'xlsx'
New-CoachSkatersList -entries $entries -folder $comp_folder -format 'xlsx'
New-PhotoPermissionList -entries $entries -folder $comp_folder -format 'xlsx'
New-ProofOfAgeAndMemberships -entries $entries -folder $comp_folder -format 'xlsx'

#Read-Host -Prompt "Press Enter to exit"

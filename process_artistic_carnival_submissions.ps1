<#
	.SYNOPSIS
		Processes Jotform for Artistic Carnival Competition
	
	.DESCRIPTION
		Processes Jotform for Artistic Carnival Competition.
	
	.EXAMPLE
		PS C:\> .\process_artistic_carnival_submissions.ps1
	
	.NOTES
		===========================================================================
		Created on:   	22/08/2019 10:59 AM
		Created by:   	Aaron Duncan
		Organization: 	ACTISA
		Filename:     	process_artistic_carnival_submissions.ps1
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

# the '2019 Carnival of Artistic Skating and TOI' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRS41PTQyGeHZJhPpZmGvlViCWgTu8dHsgvIOWo7lVa3WcOUfx4gK91uOi5g_K5e7_fdJ8Sl9PdYIuN/pub?output=tsv'

$template_folder = 'D:\Skating Templates';

$Competition = "Carnival of Artistic Skating and TOI $(Get-Date -Format yyyy)"

$certificate_template = "Certificate - ${Competition}.docx"

$comp_folder = "D:\ACTISA_COMP - $Competition";

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
		$skatername,
		$destination,
		$program
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
	else
	{
		Write-Warning "gender is null: category is $category"
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
	
	if ($category -match 'Couple')
	{
		$name = $entry.'Skater 1 Name' + " / " + $entry.'Skater 2 Name';
		$member_num = $entry.'Skater 1 Membership Number' + " / " + $entry.'Skater 2 Membership Number';
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
	
	$new_music_file += "_${skatername}"
	
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
	$div_field = $entry.'Division:';
	$category = $div_field.Split("-", 2)[0].trim()
	$division = $div_field.Split("-", 2)[1].trim()
	$gender = $entry.'Skater 1 Gender'
	$music_url = $entry.'Music File'
	$music_sp_url = $entry.'SP Music File'
	
	if ($category -match 'Couple')
	{
		$name = $entry.'Skater 1 Name' + "_" + $entry.'Skater 2 Name';
	}
	else
	{
		$name = ConvertTo-CapitalizedName -name $entry.'Skater 1 Name'
	}
	
	$submission_folder =
	(Resolve-Path "$([System.IO.Path]::Combine($submissionFullPath, "${submission_id}_${name}"))*").Path
	
	if ([string]::IsNullOrEmpty($submission_folder))
	{
		$submission_folder = [System.IO.Path]::Combine($submissionFullPath, "${submission_id}_${name}-")
	}
	
	Write-Host "Name: $name"
	
	# half ice divisions don't have music file uploads
	if (!$category.StartsWith("Aussie Skate (Half"))
	{
		$music_file = [System.Web.HttpUtility]::UrlDecode($music_url.Split("/")[-1]);
		$music_fullpath = [System.IO.Path]::Combine($submission_folder, $music_file)
		$extension = [System.IO.Path]::GetExtension($music_file)
		
		if ((Test-Path -Path $submission_folder -ErrorAction SilentlyContinue) -eq $false)
		{
			New-Item $submission_folder -Type Directory | Out-Null
		}
		#Write-Host "music url: $music_url"
		#Write-Host "music fullpath: $music_fullpath"
		
		if ((Test-Path -Path $music_fullpath -ErrorAction SilentlyContinue) -eq $false)
		{
			# music file is missing, so download it
			Get-WebFile -url $entry.'Music File' -destination $music_fullpath
		}
		
		if ($division -match 'Advanced Novice|Junior|Senior')
		{
			Publish-SkaterMusicFile -filename $music_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FP"
		}
		elseif ($category -match 'Dance')
		{
			Publish-SkaterMusicFile -filename $music_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FD"
		}
		else
		{
			Publish-SkaterMusicFile -filename $music_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FS"
		}
	}
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
		$folder
	)
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	if ($prompt)
	{
		$template = Find-Template -message "Select Certificate Template" -initial_dir $template_folder -default $certificate_template
	}
	else
	{
		$template = Resolve-Path -Path "${template_folder}/${certificate_template}"
	}
	$input_csv = [System.IO.Path]::Combine($folder, "certificate_inputs.csv")
	
	Write-Host "Input CSV file: $input_csv"
	
	$results = @()
	
	$entries | ForEach-Object {
		$category = $_.'Division:'.Split("-", 2)[0].trim()
		$division = $_.'Division:'.Split("-", 2)[1].trim()
		
		Write-Host "Generating certificate for $name ($division)"
		if ($category -match "Adult|Dance")
		{
			if ($category.Contains('('))
			{
				$cat_parts = $category.split('()')
				$division = "$(${cat_parts}[1]) $(${cat_parts}[0]) - $division"
			}
			else
			{
				$division = "${category} ${division}"
			}
		}
		
		
		foreach ($name in $_.'Skater 1 Name', $_.'Skater 2 Name')
		{
			if (-not [String]::IsNullOrEmpty($name))
			{
				$CapitalisedName = ConvertTo-CapitalizedName -name $name
				Write-Host "Name: $CapitalisedName"
				
				$results += New-Object -TypeName PSObject -Property @{
					"Name"	   = $CapitalisedName
					"Division" = $division
				}
			}
		}
		
		$members = @('Captain')
		for ($i = 1; $i -lt 25; $i++)
		{
			$members += "Skater $i"
		}
		
		foreach ($member in $members)
		{
			$firstname = $_."Team Members Details: >> $member >> First Name"
			$lastname = $_."Team Members Details: >> $member >> Last Name"
			if (-not [String]::IsNullOrEmpty($firstname))
			{
				$CapitalisedName = ConvertTo-CapitalizedName -name "$firstname $lastname"
				Write-Host "Name: $CapitalisedName"
				
				$results += New-Object -TypeName PSObject -Property @{
					"Name"	   = $CapitalisedName
					"Division" = $division
				}
			}
		}
	}
	
	Write-Host "Number of entries = $($entries.Count)"
	
	$results | Select-Object "Name", "Division" | Sort-Object -Property "Division", "Name" | Export-Csv -path $input_csv -Force -NoTypeInformation
	
	Invoke-MailMerge -template $template -datasource $input_csv -destination $folder
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
	
	# $divhash.GetEnumerator() | Sort Name
	
	foreach ($div in $divhash.Keys | Sort-Object)
	{
		$category = $div.Split("-", 2)[0].trim()
		$division = $div.Split("-", 2)[1].trim()
		
		# NOTE: need to put this in a hash based on category and division
		# Perhaps something like have a category number, which is overridden by division
		if ($category.StartsWith("Aussie Skate (Half"))
		{
			$warmup = 3
			$performance_time = 2
		}
		else
		{
			$warmup = 6
			$performance_time = 4 # this is not right - but a hack until a proper set of values is implemented
		}
		
		Add-Content -Path $schedule1 -Value "Category: $category,Division: $division"
		
		$count = $divhash.Item($div).Count
		$num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
		Add-Content -path $schedule1 -Value "Num Entries: $count"
		"Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes" -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule1
		"Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time) | Add-Content -Path $schedule1
		
		Add-Content -path $schedule1 -Value "Last Name,First Name,State,Coach Name,Other Coach Names,Music Title,Gender"
		
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
			
			"{0},{1},{2},{3},`"{4}`",`"{5}`",{6}" -f $_.'Last Name', $_.'First Name', $_.'Skater 1 State/Territory', $coach_name, $_.'Other Coach Names', $music_title.Trim(), $_.'Skater 1 Gender' | Add-Content -path $schedule1
		}
		Add-Content -Path $schedule1 -Value ""
		
		if ($category -match 'Singles' -and $division -match 'Advanced Novice|Junior|Senior')
		{
			# need to generate additional schedule for short programs
			Add-Content -Path $schedule2 -Value "Category: $category,Division: $division"
			$count = $divhash.Item($div).Count
			$num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
			Add-Content -path $schedule2 -Value "Num Entries: $count"
			"Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes" -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule2
			"Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time) | Add-Content -Path $schedule2
			
			Add-Content -path $schedule2 -Value "Last Name,First Name,Age,Coach Name,Music Title,Coach E-mail,Other Coach Names"
			
			$divhash.Item($div) | ForEach-Object {
				$num_segments = $_.'Number of Segments in Music (1-4)'
				if ($_.'SP Music Details (Segment 1)' -match 'Title: (.*) Artist:')
				{
					$music_title = $Matches[1]
				}
				else
				{
					$music_title = '<NOT PROVIDED>'
				}
				
				"{0},{1},{2},{3},{4},{5}" -f $_.'Last Name', $_.'First Name', $_.'Skater 1 Age', $_.'Coach Name', $music_title.Trim(), $_.'Coach E-mail', $_.'Other Coach Names' | Add-Content -path $schedule2
			}
			Add-Content -Path $schedule2 -Value ""
		}
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
		$division = $entry.'Division'
		
		if ($division -notmatch "Aussie Skate" -and $division -notmatch "Dance")
		{
			try
			{
				$division += " " + $gender[$entry.'Skater 1 Gender:']
			}
			catch
			{
				Write-Warning "failed to get gender for $($entry.'Skater 1 Name')"
			}
		}
		
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	$rows = @()
	
	$trophy_count = @{ gold = 0; silver = 0; bronze = 0 }
	$medal_count = @{ gold = 0; silver = 0; bronze = 0 }
	
	foreach ($div in $divhash.Keys | Sort-Object)
	{
		$category = $div.Split(";")[0].trim()
		$division = $div.Split(";")[1].trim()
		
		if ($category -match "Adult|Dance")
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
	
	$rows += (@{
			border = $false;
			values = @('', '', 'TOTAL:', $trophy_count.bronze, $trophy_count.silver, $trophy_count.gold)
		})
	
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
	
	$medalPath = [System.IO.Path]::Combine($folder, "${Competition} - MEDALS.docx")
	$trophyPath = [System.IO.Path]::Combine($folder, "${Competition} - TROPHIES.docx")
	
	if (Test-Path $medalPath) { Remove-Item -Path $medalPath }
	if (Test-Path $trophyPath) { Remove-Item -Path $trophyPath }
	
	$gender = @{ Female = "Ladies"; Male = "Men" }
	
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
	
	# Create Trophy Table
	$TrophyWord = New-Object -ComObject Word.Application
	$TrophyWord.Visible = $False
	$TrophyDoc = $TrophyWord.Documents.Add()
	$TrophyTable = $TrophyDoc.Tables.Add($TrophyWord.Selection.Range(), 2, 4)
	$TrophyTable.Range.Style = "No Spacing"
	$TrophyTable.Borders.Enable = $True
	$TrophyTable.Rows(2).Cells(1).Range.Text = "Division"
	$TrophyTable.Rows(2).Cells(1).Range.Bold = $True
	
	foreach ($div in $divhash.Keys | Sort-Object)
	{
		$category = $div.Split("-", 2)[0].trim()
		$division = $div.Split("-", 2)[1].trim()
		
		$count = $divhash.Item($div).Count
		$Row = $TrophyTable.Rows.Add()
		
		$Row.Cells(1).Range.Text = $division
		$Row.Cells(1).Range.Bold = $True
		$Row.Cells(1).Range.Font.Spacing = 1.0
		$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
		
		if ($count -ge 1) { $Row.Cells(2).Range.Text = "$Competition`r$division`r1st Place" }
		if ($count -ge 2) { $Row.Cells(3).Range.Text = "$Competition`r$division`r2nd Place" }
		if ($count -ge 3) { $Row.Cells(4).Range.Text = "$Competition`r$division`r3rd Place" }
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
	
	$headers = @('Coach Name', 'Coach E-Mail')
	$rows = @()
	foreach ($entry in $entries)
	{
		$rows += (@{ 'border' = $true; 'values' = @($entry.'Primary Coach Name:', $entry.'Primary Coach E-mail:') })
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
	foreach ($surname in $hash.Keys | Sort-Object)
	{
		$name = $hash.Item($surname)
		foreach ($firstname in $name.Keys | Sort-Object)
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
			$category = $entry.Division.Split(";")[0].trim()
			$division = $entry.Division.Split(";")[1].trim()
			
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
		[string]
		$folder,
		[string]
		$format = 'csv'
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
		$category = $entry.Division.Split(";")[0].trim()
		$division = $entry.Division.Split(";")[1].trim()
		
		if ($category -eq "Adult")
		{
			$division = "Adult ${division}"
		}
		
		$parentName = $entry.'Parent/Guardian Name: (First Name)' + ' ' + $entry.'Parent/Guardian Name: (Last Name)'
		$rows += (@{
				'border' = $true;
				'values' = @(
					$entry.'Skater 1 Name',
					$division,
					$parentName,
					$entry.'Payment due (AUD)',
					$entry.'Direct Debit Receipt')
			})
	}
	
	if ($rows.Count -gt 0)
	{
		$headers = @("Skater Name", "Division", "Parent/Guardian (if applicable)", "Payment Due (AUD)", "Direct Debit Receipt")
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
		[string]
		$folder,
		[string]
		$format = 'csv'
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
	
	foreach ($coach_name in $hash.Keys | Sort-Object)
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
		$catdiv = $entry.Division.Split(";")
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
		Generate Skater Membership/POA Spreadsheet
	
	.DESCRIPTION
		Generate a list of all skaters, with their state/coach/membership number/POA
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.PARAMETER format
		A description of the format parameter.
	
	.EXAMPLE
				PS C:\> New-MembershipAndPOASpreadsheet
	
	.NOTES
		Additional information about the function.
#>
function New-MembershipAndPOASpreadsheet
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
New-RegistrationList -entries $entries -folder $comp_folder -format 'xlsx'
New-EngravingSchedule -entries $entries -folder $comp_folder
New-VolunteerSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-PaymentSpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'
New-SkaterEmailList -entries $entries -folder $comp_folder -format 'xlsx'
New-CoachEmailList -entries $entries -folder $comp_folder -format 'xlsx'
New-CoachSkatersList -entries $entries -folder $comp_folder -format 'xlsx'
New-PhotoPermissionList -entries $entries -folder $comp_folder -format 'xlsx'
New-MembershipAndPOASpreadsheet -entries $entries -folder $comp_folder -format 'xlsx'

#Read-Host -Prompt "Press Enter to exit"

﻿<#
	.SYNOPSIS
		Processes JotForm submissions for Autumn Trophy
	
	.DESCRIPTION
		Processes JotForm submissions for Autumn Trophy
	
	.PARAMETER prompt
		if the user should be prompted for folder/file locations

	.EXAMPLE
			process_autumn_trophy_submissions.ps1
			process_autumn_trophy_submissions.ps1 -prompt $true
			process_autumn_trophy_submissions.ps1 -prompt $false
	
	.NOTES
		===========================================================================
		Created on:   	22/08/2019 10:59 AM
		Created by:   	Aaron Duncan
		Organization: 	ACTISA
		Filename:     	process_autumn_trophy_submissions.ps1
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
    "Free Skate"      = "FS";
    "Aussie Skate"    = "AS";
    "Free Dance"      = "FD";
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

# the '2020 Autumn Trophy' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR4Nw5LFzH7r5vFdZU-u39kXfpwxguYQPDgvzcENeiAWmEec7b2T616VjhQYchPxh7N4D5xmPGDh0BJ/pub?output=xlsx'

$template_folder = 'D:\Jotform MailMerge Templates';

$Competition = "Autumn Trophy $(Get-Date -Format yyyy)";

$certificate_template = "Certificate - $Competition.docx"

if (Test-Path D:)
{
	$comp_folder = "D:\ACTISA_COMP - $($Competition)"
}
else
{
	$comp_folder = "C:\ACTISA_COMP - $($Competition)"
}

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
	$extension = [System.IO.Path]::GetExtension($filename)
	try
	{
		$music_duration = Get-MusicFileDuration -filename $filename
		if ($music_duration -eq 'notfound')
		{
			# we failed to get the metadata from the file, so try an alternate file extension
			if ($extension -eq 'mp3')
			{
				$newExt = 'm4a'
			}
			else
			{
				$newExt = 'mp3'
			}
			
			# make a copy of the file with the new file extension
			$newFilename = [System.IO.Path]::ChangeExtension($filename, $newExt)
			Copy-Item $filename $newFilename
			
			# get the duration of the song from the new file
			$music_duration = Get-MusicFileDuration -filename $newFilename
			if ($music_duration -ne 'notfound')
			{
				# we got a duration this time, so reference the new file (with the valid extension)
				$filename = $newFilename
				$extension = $newExt
			}
		}
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
	
	if ($category -notmatch "Aussie|Couple")
	{
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
	}
	
	if ($category -match "Adult|Dance")
	{
		$music_subdir = "$category $music_subdir"
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
	
	#$entry | Format-Table -Force *
	$submission_id = $entry.'Submission ID'
	$div_field = $entry.'Division'
	$category = $div_field.Split(";")[0].trim()
	$division = $div_field.Split(";")[1].trim()
	$gender = $entry.'Skater 1 Gender:'
	$music_fs_url = $entry.'FS/FD Music File'
	$music_sp_url = $entry.'SP Music File'
	$music_pd2_url = $entry.'PD2 Music File'
	
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
		# 1. all have a FS
		# 2. AdvNov/Junior/Senior have a SP
		# 3. Dance has a PD2 (except Junior/Senior)
		
		$music_fs_file = [System.Web.HttpUtility]::UrlDecode($music_fs_url.Split("/")[-1]);
		$music_fs_fullpath = [System.IO.Path]::Combine($submission_folder, $music_fs_file)
		
		$extension = [System.IO.Path]::GetExtension($music_fs_file)
		
		if ((Test-Path -Path $submission_folder -ErrorAction SilentlyContinue) -eq $false)
		{
			New-Item $submission_folder -Type Directory | Out-Null
		}
		
		if ((Test-Path -Path $music_fs_fullpath -ErrorAction SilentlyContinue) -eq $false)
		{
			# music file is missing, so download it
			Get-WebFile -url $music_fs_url -destination $music_fs_fullpath
		}
		
		if ($division -match 'Advanced Novice|Junior|Senior')
		{
			#Write-Host "Splitting music_sp_url: '${music_sp_url}'"
			$music_sp_file = [System.Web.HttpUtility]::UrlDecode($music_sp_url.Split("/")[-1]);
			$music_sp_fullpath = [System.IO.Path]::Combine($submission_folder, $music_sp_file)
			
			if ((Test-Path -Path $music_sp_fullpath -ErrorAction SilentlyContinue) -eq $false)
			{
				# music file is missing, so download it
				Get-WebFile -url $music_sp_url -destination $music_sp_fullpath
			}
			
			if ($category -match 'Dance')
			{
				Publish-SkaterMusicFile -filename $music_fs_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FD"
				Publish-SkaterMusicFile -filename $music_sp_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "SD"
			}
			else
			{
				Publish-SkaterMusicFile -filename $music_fs_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FS"
				Publish-SkaterMusicFile -filename $music_sp_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "SP"
			}
			
			if ([String]::IsNullOrEmpty($music_sp_file))
			{
				Write-Warning "WARNING: No SD/SP/PD1 Music file provided for $name in category '$category' division '$division'"
				Write-Host "Music file: $music_sp_file"
			}
			else
			{
				if ((Test-Path -Path $music_sp_fullpath -ErrorAction SilentlyContinue) -eq $false)
				{
					# music file is missing, so download it
					Get-WebFile -url $entry.'SD/SP/PD1 Music File' -destination $music_sp_fullpath
				}
			}
			
			# process the file even if it is nonexistent - the result will be a "NOTFOUND" file in the Music folder (so we get alerted)
			#Publish-SkaterMusicFile -filename $music_sp_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "SP"
		}
		elseif ($category -match 'Dance')
		{
			$music_pd1_file = [System.Web.HttpUtility]::UrlDecode($music_sp_url.Split("/")[-1]);
			$music_pd1_fullpath = [System.IO.Path]::Combine($submission_folder, $music_pd1_file)
			$music_pd2_file = [System.Web.HttpUtility]::UrlDecode($music_pd2_url.Split("/")[-1]);
			$music_pd2_fullpath = [System.IO.Path]::Combine($submission_folder, $music_pd2_file)
			
			if ((Test-Path -Path $music_pd1_fullpath -ErrorAction SilentlyContinue) -eq $false)
			{
				# music file is missing, so download it
				Get-WebFile -url $music_sp_url -destination $music_pd1_fullpath
			}
			
			if ((Test-Path -Path $music_pd2_fullpath -ErrorAction SilentlyContinue) -eq $false)
			{
				# music file is missing, so download it
				Get-WebFile -url $music_pd2_url -destination $music_pd2_fullpath
			}
			
			Publish-SkaterMusicFile -filename $music_fs_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FD"
			Publish-SkaterMusicFile -filename $music_pd1_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "PD1"
			Publish-SkaterMusicFile -filename $music_pd2_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "PD2"
		}
		else
		{
			Publish-SkaterMusicFile -filename $music_fs_fullpath -category $category -division $division -skatername $name -gender $gender -destination $music_folder -program "FS"
		}
	}
}

<#
	.SYNOPSIS
		Generate PPC Forms
	
	.DESCRIPTION
		Generates the PPC Forms from the submission data
	
	.PARAMETER entries
		A description of the entries parameter.
	
	.PARAMETER folder
		A description of the folder parameter.
	
	.EXAMPLE
				PS C:\> New-PPCForms
	
	.NOTES
		Additional information about the function.
#>
function New-PPCForms
{
	param
	(
		$entries,
		$folder
	)
	
	Write-Host "Generating PPC Forms"
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$sp_ppc_file = [System.IO.Path]::Combine($folder, "short_ppc_entries.csv");
	$fp_ppc_file = [System.IO.Path]::Combine($folder, "long_ppc_entries.csv");
	
	$sp_template = [System.IO.Path]::Combine($template_folder, "Short Program PPC.doc");
	$fp_template = [System.IO.Path]::Combine($template_folder, "Free Skate PPC.doc");
	
	$sp_results = @()
	$fp_results = @()
	foreach ($entry in $entries)
	{
		$category = $entry.Division.Split(";")[0].trim()
		$division = $entry.Division.Split(";")[1].trim()
		
		if ($category -notmatch 'Aussie Skate')
		{
			if ($category -match 'Adult|Dance')
			{
				$division = "${category} ${division}"
			}
			
			if ($category -match 'Couple')
			{
				$name = $entry.'Skater 1 Name' + " / " + $entry.'Skater 2 Name';
				$member_num = $entry.'Skater 1 Membership Number:' + " / " + $entry.'Skater 2 Membership Number:';
			}
			else
			{
				$name = $entry.'Skater 1 Name';
				$member_num = $entry.'Skater 1 Membership Number:';
			}
			
			if ($category -match 'Singles' -and $division -match 'Advanced Novice|Junior|Senior')
			{
				$sp_results += New-Object -TypeName PSObject -Property @{
					"Name"   = $name
					"State"  = $entry.'Skater 1 State/Territory:'
					"Membership Number" = $member_num
					"Division" = $division
					"Element 1" = $entry.'Element 1'
					"Element 2" = $entry.'Element 2'
					"Element 3" = $entry.'Element 3'
					"Element 4" = $entry.'Element 4'
					"Element 5" = $entry.'Element 5'
					"Element 6" = $entry.'Element 6'
					"Element 7" = $entry.'Element 7'
					"Element 8" = $entry.'Element 8'
				}
			}
			
			$fp_results += New-Object -TypeName PSObject -Property @{
				"Name"   = $name
				"State"  = $entry.'Skater 1 State/Territory:'
				"Membership Number" = $member_num
				"Division" = $division
				"Element 1" = $entry.'Element 1 2'
				"Element 2" = $entry.'Element 2 2'
				"Element 3" = $entry.'Element 3 2'
				"Element 4" = $entry.'Element 4 2'
				"Element 5" = $entry.'Element 5 2'
				"Element 6" = $entry.'Element 6 2'
				"Element 7" = $entry.'Element 7 2'
				"Element 8" = $entry.'Element 8 2'
				"Element 9" = $entry.'Element 9'
				"Element 10" = $entry.'Element 10'
				"Element 11" = $entry.'Element 11'
				"Element 12" = $entry.'Element 12'
				"Element 13" = $entry.'Element 13'
			}
		}
	}
	
	if ($sp_results.Count -gt 0)
	{
		$sp_results |
		Select-Object "Name", "State", "Membership Number", "Division",
					  "Element 1", "Element 2", "Element 3", "Element 4",
					  "Element 5", "Element 6", "Element 7", "Element 8" |
		export-csv -path $sp_ppc_file -Force -NoTypeInformation
	}
	
	if ($fp_results.Count -gt 0)
	{
		$fp_results |
		Select-Object "Name", "State", "Membership Number", "Division",
					  "Element 1", "Element 2", "Element 3", "Element 4", "Element 5",
					  "Element 6", "Element 7", "Element 8", "Element 9", "Element 10",
					  "Element 11", "Element 12", "Element 13" |
		export-csv -path $fp_ppc_file -Force -NoTypeInformation
	}
	
	if (Test-Path -Path $sp_ppc_file -ErrorAction SilentlyContinue)
	{
		Invoke-MailMerge -template $sp_template -datasource $sp_ppc_file -destination $folder
	}
	else
	{
		Write-Warning "There are no Short Program PPC submissions"
	}
	
	if (Test-Path -Path $fp_ppc_file -ErrorAction SilentlyContinue)
	{
		Invoke-MailMerge -template $fp_template -datasource $fp_ppc_file -destination $folder
	}
	else
	{
		Write-Warning "There are no Free Skate PPC submissions"
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
	
	Write-Host "Generating Certificates."
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$input_csv = [System.IO.Path]::Combine($folder, "autumn_trophy_inputs.csv")
	
	$results = @()
	$entries | ForEach-Object {
		$name = ConvertTo-CapitalizedName -name $_.'Skater 1 Name'
		$category = $_.Division.Split(";")[0].trim()
		$division = $_.Division.Split(";")[1].trim()
		
		if ($category -match "Adult|Dance")
		{
			$division = "${category} ${division}"
		}
		
		if ($category -match "Dance")
		{
			$name2 = ConvertTo-CapitalizedName -name $_.'Skater 2 Name'
			Write-Host "    - $name / $name2 ($division)"
		}
		else
		{
			Write-Host "    - $name ($division)"
		}
		
		$results += New-Object -TypeName PSObject -Property @{
			"Name"	    = $name
			"Firstname" = $_.'First Name'
			"Surname"   = $_.'Last Name'
			"Division"  = $division
		}
		
		if ($category -match "Dance")
		{
			$name = ConvertTo-CapitalizedName -name $_.'Skater 2 Name'
			$results += New-Object -TypeName PSObject -Property @{
				"Name"	    = $name
				"Firstname" = $_.'Skater 2 Name: (First Name)'
				"Surname"   = $_.'Skater 2 Name: (Last Name)'
				"Division"  = $division
			}
		}
	}
	
	if ($results.Count -gt 0)
	{
		Write-Host " - generating $($results.Count) Certificates."
		$results | Sort-Object -Property "Division", "Surname", "Firstname" | Select-Object "Name", "Division" | export-csv -path $input_csv -Force -NoTypeInformation
		if ($prompt)
		{
			$comp_template = Find-Template -message "Select Certificate Template" -initial_dir $template_folder -default $certificate_template
		}
		else
		{
			$comp_template = Resolve-Path -Path "${template_folder}/${certificate_template}"
		}
		#Write-Host "COMP TEMPLATE: ${comp_template}"
		Invoke-MailMerge -template $comp_template -datasource $input_csv -destination $folder
	}
	else
	{
		Write-Host " - no Competition entries found..."
	}
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
		$division = $entry.'Division'
		if (!$divhash.ContainsKey($division))
		{
			$divhash[$division] = @()
		}
		$divhash[$division] += $entry
	}
	
	#$divhash.GetEnumerator() | Sort Name
	
	foreach ($div in $divhash.Keys | Sort)
	{
		$category = $div.Split(";")[0].trim()
		$division = $div.Split(";")[1].trim()
		
		# NOTE: need to put this in a hash based on category and division
		# Perhaps something like have a category number, which is overridden by division
		if ($category.StartsWith("Aussie Skate (Half"))
		{
			$warmup = 3
			$performance_time = 2
		}
		elseif ($category.StartsWith("Aussie Skate (Full"))
		{
			$warmup = 4
			$performance_time = 3
		}
		elseif ($category -eq 'Adult')
		{
			$warmup = 6
			$performance_time = 4
			$division = "Adult $division"
		}
		elseif ($category -eq 'Singles' -and $division -eq 'Preliminary')
		{
			$warmup = 6
			$performance_time = 3
		}
		elseif ($category -eq 'Singles' -and $division -eq 'Elementary')
		{
			$warmup = 6
			$performance_time = 4
		}
		else
		{
			$warmup = 6
			$performance_time = 4
		}
		#Write-Host "Category: $category,Division: $division"
		Add-Content -Path $schedule1 -Value "Category: $category,Division: $division"
		
		$count = $divhash.Item($div).Count
		$num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
		Add-Content -path $schedule1 -Value "Num Entries: $count"
		"Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes" -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule1
		"Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time) | Add-Content -Path $schedule1
		
		Add-Content -path $schedule1 -Value "Last Name,First Name,Age,Coach Name,Music Title,Coach E-mail,Other Coach Names"
		
		$divhash.Item($div) | ForEach-Object {
			$num_segments = $_.'Number of Segments in Music'
			if ($_.'FS/FD Music Details (Segment 1)' -match 'Title: (.*) Artist:')
			{
				$music_title = $Matches[1]
			}
			else
			{
				$music_title = '<NOT PROVIDED>'
			}
			
			"{0},{1},{2},{3},{4},{5}" -f $_.'Last Name', $_.'First Name', $_.'Skater 1 Age', $_.'Primary Coach Name:', $music_title.Trim(), $_.'Primary Coach E-mail:', $_.'Other Coach Names' | Add-Content -path $schedule1
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
		$medal_gold = ""
		$medal_silver = ""
		$medal_bronze = ""
		$trophy_gold = ""
		$trophy_silver = ""
		$trophy_bronze = ""
		
		if ($category.StartsWith("Aussie"))
		{
			if ($count -ge 1) { $medal_gold = $numSkaters; $medal_count.gold += $numSkaters }
			if ($count -ge 2) { $medal_silver = $numSkaters; $medal_count.silver += $numSkaters }
			if ($count -ge 3) { $medal_bronze = $numSkaters; $medal_count.bronze += $numSkaters }
		}
		else
		{
			if ($count -ge 1) { $trophy_gold = $numSkaters; $trophy_count.gold += $numSkaters }
			if ($count -ge 2) { $trophy_silver = $numSkaters; $trophy_count.silver += $numSkaters }
			if ($count -ge 3) { $trophy_bronze = $numSkaters; $trophy_count.bronze += $numSkaters }
		}
		
		$rows += (@{ 'border' = $true; 'values' = @($category, $division, $count, $medal_bronze, $medal_silver, $medal_gold, $trophy_bronze, $trophy_silver, $trophy_gold) })
	}
	
	$rows += (@{
			border  = $false;
			values  = @('', '', 'TOTAL:', $medal_count.bronze, $medal_count.silver, $medal_count.gold, $trophy_count.bronze, $trophy_count.silver, $trophy_count.gold)
		})
	
	$headers = @('Category', 'Division', '# Entries', 'Bronze Medal', 'Silver Medal', 'Gold Medal', 'Bronze Trophy', 'Silver Trophy', 'Gold Trophy')
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
	
	# Create Medals Table
	$MedalWord = New-Object -ComObject Word.Application
	$MedalWord.Visible = $False
	$MedalDoc = $MedalWord.Documents.Add()
	$MedalTable = $MedalDoc.Tables.Add($MedalWord.Selection.Range(), 2, 4)
	$MedalTable.Range.Style = "No Spacing"
	$MedalTable.Borders.Enable = $True
	$MedalTable.Rows(2).Cells(1).Range.Text = "Division"
	$MedalTable.Rows(2).Cells(1).Range.Bold = $True
	
	# Create Trophy Table
	$TrophyWord = New-Object -ComObject Word.Application
	$TrophyWord.Visible = $False
	$TrophyDoc = $TrophyWord.Documents.Add()
	$TrophyTable = $TrophyDoc.Tables.Add($TrophyWord.Selection.Range(), 2, 4)
	$TrophyTable.Range.Style = "No Spacing"
	$TrophyTable.Borders.Enable = $True
	$TrophyTable.Rows(2).Cells(1).Range.Text = "Division"
	$TrophyTable.Rows(2).Cells(1).Range.Bold = $True
	
	# Create the title row on each table
	$Row = $MedalTable.Rows(1)
	$Row.Cells.Merge()
	$Row.Cells(1).Range.Text = "$Competition engraving schedule - MEDALS"
	$Row.Cells(1).Range.Bold = $true
	$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
	$Row = $TrophyTable.Rows(1)
	$Row.Cells.Merge()
	$Row.Cells(1).Range.Text = "$Competition engraving schedule - TROPHY"
	$Row.Cells(1).Range.Bold = $true
	$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
	
	foreach ($div in $divhash.Keys | Sort-Object)
	{
		$category = $div.Split(";")[0].trim()
		$division = $div.Split(";")[1].trim()
		
		if ($category -match "Adult|Dance")
		{
			#$division = "$category $division"
		}
		
		$count = $divhash.Item($div).Count
		if ($category.StartsWith("Aussie"))
		{
			$Row = $MedalTable.Rows.Add()
		}
		else
		{
			$Row = $TrophyTable.Rows.Add()
		}
		
		$Row.Cells(1).Range.Text = $division
		$Row.Cells(1).Range.Bold = $True
		$Row.Cells(1).Range.Font.Spacing = 1.0
		$Row.Cells(1).Range.paragraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
		
		if ($count -ge 1) { $Row.Cells(2).Range.Text = "$Competition`r$division`r1st Place" }
		if ($count -ge 2) { $Row.Cells(3).Range.Text = "$Competition`r$division`r2nd Place" }
		if ($count -ge 3) { $Row.Cells(4).Range.Text = "$Competition`r$division`r3rd Place" }
	}
	
	# Save the documents
	$MedalDoc.SaveAs($medalPath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
	$TrophyDoc.SaveAs($trophyPath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
	$MedalDoc.close()
	$TrophyDoc.close()
	$MedalWord.quit()
	$TrophyWord.quit()
	
	# Cleanup the memory
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalDoc) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyDoc) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalWord) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyWord) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalTable) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyTable) | Out-Null
	Remove-Variable MedalDoc, TrophyDoc, MedalWord, TrophyWord, Row, MedalTable, TrophyTable
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
					'border'  = $true;
					'values'  = @($entry.'Skater 1 Name',
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
		$category = $entry.Division.Split(";")[0].trim()
		$division = $entry.Division.Split(";")[1].trim()
		
		if ($category -eq "Adult")
		{
			$division = "Adult ${division}"
		}
		
		$parentName = $entry.'Parent/Guardian Name: (First Name)' + ' ' + $entry.'Parent/Guardian Name: (Last Name)'
		$rows += (@{
				'border'  = $true;
				'values'  = @(
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
	Write-Host "FOLDER: '${folder}'"
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
				'border'  = $true;
				'values'  = @($category,
					$division,
					$entry.'Skater 1 Name',
					$entry.'Skater 2 Name',
					$entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
				)
			})
	}
	New-Spreadsheet -name "Photo Permissions" -path $outfile -headers $headers -rows $rows -format $format
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
			Write-Host "    - $name"
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
				Write-Host "    - $name"
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
	$comp_folder = Find-Folders -title "Select the Competition folder (default = $comp_folder)" -default $comp_folder
	$template_folder = Find-Folders -title "Select the MailMerge Template folder (default = $template_folder)" -default $template_folder
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

foreach ($f in ('Submissions', 'Music', 'PPC', 'Certificates', 'Schedule'))
{
    if ((Test-Path $f -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $f -ItemType Directory | Out-Null
    }
}

Pop-Location

$submissionFullPath = [System.IO.Path]::Combine($comp_folder, "Submissions")
$music_folder       = [System.IO.Path]::Combine($comp_folder, "Music")
$ppc_folder         = [System.IO.Path]::Combine($comp_folder, "PPC")
$certificate_folder = [System.IO.Path]::Combine($comp_folder, "Certificates")
$schedule_folder    = [System.IO.Path]::Combine($comp_folder, "Schedule")

Write-Host "Competition Folder: $comp_folder"
write-host "Music Folder: $music_folder"

$entries = @()
Get-SubmissionEntries -url $google_sheet_url | ForEach-Object { if ($_ -ne 0) { $entries += $_ } }

Write-Host "Number of entries = $($entries.Count)`n" -ForegroundColor Yellow

foreach ($entry in $entries)
{
	Publish-EntryMusicFiles -entry $entry -submissionFullPath $submissionFullPath -music_folder $music_folder
}

New-PPCForms -entries $entries -folder $ppc_folder
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

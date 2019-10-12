
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

# the '2019 Christmas Gala' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRPrJY_zj6i-Ir-qfZARYjloDENPnYha6kwFjeJdkOEXZ12w6tD9kP3P46pQYi59Hm7ANjg_e7WdpPg/pub?output=tsv'

$template_folder = 'D:\Skating Templates';

$Competition = "Christmas Gala $(Get-Date -Format yyyy)";

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
	
	.PARAMETER skatername
		A description of the skatername parameter.
	
	.PARAMETER destination
		A description of the destination parameter.
	
	.PARAMETER program
		A description of the program parameter.
	
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
	$music_subdir = $division
	
	$music_dest = [System.IO.Path]::Combine($destination, $music_subdir)
	
	#
	# Calculate the new music filename
	#
	
	$new_music_file = $division -replace "\(.*\)", ""
	
	foreach ($key in $abbreviations.Keys)
	{
		$new_music_file = $new_music_file -replace $key, $abbreviations.Item($key)
	}
	
	$new_music_file = $new_music_file -replace " ; ", "_" -replace " ", ""
	$new_music_file += "_${program}_${skatername}"
	
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
		Calls Publish-SkaterMusicFile() for each music file for a submission.
	
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
	$music_url = $entry.'Music File:'
	
	$name = ConvertTo-CapitalizedName -name $entry.'Skater 1 Name'
	
	$submission_folder = [System.IO.Path]::Combine($submissionFullPath, "${submission_id}")
	
	# half ice divisions don't have music file uploads
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
	
	Publish-MusicFile -filename $music_fullpath -skatername $name -gender $gender -destination $music_folder -program "Gala"
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
	
	.PARAMETER format
		A description of the format parameter.
	
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
	
	if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
	{
		New-Item $folder -Type Directory | Out-Null
	}
	
	$list_csv = [System.IO.Path]::Combine($folder, "registration_list.csv")
	
	$hash = @{ }
	foreach ($entry in $entries)
	{
		$surname = ConvertTo-CapitalizedName -name $entry.'Last Name'
		if (!$hash.ContainsKey($surname))
		{
			$hash[$surname] = @{ }
		}
		$firstname = ConvertTo-CapitalizedName -name $entry.'First Name'
		$hash[$surname].Add($firstname, $true)
		
		if ($entry.Division -match 'Dance')
		{
			$surname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name (Last Name)'
			if (!$hash.ContainsKey($surname))
			{
				$hash[$surname] = @{ }
			}
			$firstname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name (First Name)'
			$hash[$surname].Add($firstname, $true)
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
$music_folder       = [System.IO.Path]::Combine($comp_folder, "Music")
$certificate_folder = [System.IO.Path]::Combine($comp_folder, "Certificates")
$schedule_folder    = [System.IO.Path]::Combine($comp_folder, "Schedule")

Write-Host "Competition Folder: $comp_folder"
write-host "Music Folder: $music_folder"

$entries = Get-SubmissionEntries -url $google_sheet_url

foreach ($entry in $entries)
{
    Publish-EntryMusicFiles -entry $entry -submissionFullPath $submissionFullPath -music_folder $music_folder
}

Write-Host "Number of entries = $($entries.Count)`n" -ForegroundColor Yellow

#generate_certificates          -entries $entries -folder $certificate_folder 
#generate_skating_schedule      -entries $entries -folder $schedule_folder
#generate_division_counts       -entries $entries -folder $eventFolder
New-RegistrationList  -entries $entries -folder $eventFolder -format 'xlsx'
New-VolunteerSpreadsheet -entries $entries -folder $eventFolder -format 'xlsx'
New-PaymentSpreadsheet -entries $entries -folder $eventFolder -format 'xlsx'

#Read-Host -Prompt "Press Enter to exit"

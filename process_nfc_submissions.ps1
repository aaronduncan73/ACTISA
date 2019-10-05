
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

# the '2019 National Federation Challenge Registration' form on the ACTISA account
$google_sheet_url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQuVTDAdKSd79QusmGXZ5DQVVEDdpD_czH2eY8vSlYmPdNkXYTyQY65yoPRAvFoiBudP8w9qiiHuXz3/pub?output=tsv'
$template_folder = 'C:\Users\aaron\Google Drive\Skating\Skating Templates';

$Competition = "NFC $(Get-Date -Format yyyy)";
$Engraving_Title = "National Federation Challenge $(Get-Date -Format yyyy)"

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
		[string]$url,
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
		[string]$template,
		[string]$datasource,
		[string]$destination
	)
	
	if ([String]::IsNullOrEmpty($template))
	{
		Write-Warning "ERROR: empty MailMerge template file supplied to do_mail_merge()"
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
	param
	(
		[string]$filename
	)
	
	Write-Host "Getting music file duration for '$filename'"
	
	if (Test-Path -Path "$filename" -PathType Leaf -ErrorAction SilentlyContinue)
	{
		$shell = New-Object -COMObject Shell.Application
		
		$folder = Split-Path $filename
		$file = Split-Path $filename -Leaf
		
		$shellfolder = $shell.Namespace($folder)
		$shellfile = $shellfolder.ParseName($file)
		
		# find the index of "Length"
		for ($index = 0; -not $lenidx; $index++)
		{
			$details = $shellfolder.GetDetailsOf($shellfolder.Items, $index)
			if ([string]::IsNullOrEmpty($details))
			{
				Write-Host "failed to get music duration"
				"notfound"
				break
			}
			else
			{
				if ($shellfolder.GetDetailsOf($shellfolder.Items, $index) -eq 'Length')
				{
					$lenidx = $index;
				}
			}
		}
		
		$duration = $shellfolder.GetDetailsOf($shellfile, $lenidx).split(":");
		$minutes = $duration[1].ToString().TrimStart('0')
		$seconds = $duration[2].ToString().TrimStart('0')
		
		if ($minutes -eq '') { $minutes = '0' }
		if ($seconds -eq '') { $seconds = '0' }
		
		"${minutes}m${seconds}s"
	}
	else
	{
		Write-Warning "Failed to find file: $filename"
		"notfound"
	}
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
		[string]$title = "Select the Competition directory",
		[string]$default = "C:\"
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
		[string]$message,
		[string]$initial_dir,
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
		[string]$duration1,
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

#--------------------------------------------------------------------------------
# FUNCTION:
#  process_music_file
#
# DESCRIPTION:
#  Place the music file in the correct location, with the correctly formatted name.
#--------------------------------------------------------------------------------
function process_music_file($filename, $category, $division, $gender, $skatername, $destination, $program)
{
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

    foreach ( $key in $abbreviations.Keys )
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

    $new_music_path = [System.IO.Path]::Combine($music_dest,   $new_music_file)

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

#--------------------------------------------------------------------------------
# FUNCTION:
#  process_music_files
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function process_music_files($entry, $music_folder, $submissionFullPath)
{
    $submission_id = $entry.'Submission ID'
    $div_field     = $entry.'Division'
    $category      = $div_field.Split(";")[0].trim()
    $division      = $div_field.Split(";")[1].trim()
    $gender        = $entry.'Skater 1 Gender'
    $music_fs_url  = $entry.'FS Music File:'
    $music_pd1_url = $entry.'PD1 Music File:'
    $music_pd2_url = $entry.'PD2 Music File:'

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
    #Write-Host "Submission folder: $submission_folder"

    # half ice divisions don't have music file uploads
    if (!$category.StartsWith("Aussie Skate (Half"))
    {
        # 1. all have a FS
        # 2. AdvNov/Junior/Senior have a SP
        # 3. Dance has a PD2 (except Junior/Senior)

        $music_fs_file      = [System.Web.HttpUtility]::UrlDecode($music_fs_url.Split("/")[-1]);
        $music_pd1_file     = [System.Web.HttpUtility]::UrlDecode($music_pd1_url.Split("/")[-1]);

        $music_fs_fullpath  = [System.IO.Path]::Combine($submission_folder, $music_fs_file) -replace '([][[])', ''
        $music_pd1_fullpath = [System.IO.Path]::Combine($submission_folder, $music_pd1_file) -replace '([][[])', ''

        $extension          = [System.IO.Path]::GetExtension($music_fs_file)

        if ((Test-Path -Path $submission_folder -ErrorAction SilentlyContinue) -eq $false)
        {
            New-Item $submission_folder -Type Directory | Out-Null
        }

        # ensure that the FS music file is downloaded
        if ([String]::IsNullOrEmpty($music_fs_file))
        {
            Write-Warning "WARNING: No FS/FD Music file provided for $name in category '$category' division '$division'"
        }
        else
        {
            if ((Test-Path -Path $music_fs_fullpath -ErrorAction SilentlyContinue) -eq $false)
            {
                # music file is missing, so download it
                Write-Host "Music File: '$music_fs_url'"
                Get-WebFile -url $music_fs_url -destination $music_fs_fullpath
            }
        }

        if ($category -match 'Dance')
        {
            process_music_file `
                    -filename $music_fs_fullpath `
                    -category $category `
                    -division $division `
                    -skatername $name `
                    -gender $gender `
                    -destination $music_folder `
                    -program 'FD'

            process_music_file `
                -filename $music_pd1_fullpath `
                -category $category `
                -division $division `
                -skatername $name `
                -gender $gender `
                -destination $music_folder `
                -program "PD1"

            # ensure the PD2 file is downloaded
            if ([String]::IsNullOrEmpty($music_pd2_url))
            {
                Write-Warning "WARNING: Failed to retrieve PD2 musicfile URL for $name in category '$category' division '$division'"
            }
            else
            {
                $music_pd2_file = [System.Web.HttpUtility]::UrlDecode($music_pd2_url.Split("/")[-1]);
                if ([String]::IsNullOrEmpty($music_pd2_file))
                {
                    Write-Warning "WARNING: No PD2 Music file provided for $name in category '$category' division '$division'"
                }
                else
                {
                    $music_pd2_fullpath = [System.IO.Path]::Combine($submission_folder, $music_pd2_file)
                    if ((Test-Path -Path $music_pd2_fullpath -ErrorAction SilentlyContinue) -eq $false)
                    {
                        # music file is missing, so download it
                        Write-Host "Music File: '$music_pd2_url'"
                        Get-WebFile -url $music_pd2_url -destination $music_pd2_fullpath
                    }
                }

                process_music_file `
                    -filename $music_pd2_fullpath `
                    -category $category `
                    -division $division `
                    -skatername $name `
                    -gender $gender `
                    -destination $music_folder `
                    -program "PD2"
            }
        }
        else
        {
            process_music_file `
                    -filename $music_fs_fullpath `
                    -category $category `
                    -division $division `
                    -skatername $name `
                    -gender $gender `
                    -destination $music_folder `
                    -program 'FS'
        }
    }
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_ppc_forms
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_ppc_forms($entries,$folder)
{
    Write-Host "Generating PPC Forms"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $fp_ppc_file = [System.IO.Path]::Combine($folder, "long_ppc_entries.csv");
    $fp_template = [System.IO.Path]::Combine($template_folder, "Free Skate PPC.doc");

    $fp_results = @()
    foreach ($entry in $entries)
    {
        $category  = $entry.Division.Split(";")[0].trim()
        $division  = $entry.Division.Split(";")[1].trim()

        if ($category -notmatch 'Aussie Skate')
        {
            if ($category -match 'Adult|Dance')
            {
                $division = "${category} ${division}"
            }

            if ($category -match 'Couple')
            {
                $name = $entry.'Skater 1 Name' + " / " + $entry.'Skater 2 Name';
                $member_num = $entry.'Skater 1 Membership Number' + " / " + $entry.'Skater 2 Membership Number';
            }
            else
            {
                $name = $entry.'Skater 1 Name';
                $member_num = $entry.'Skater 1 Membership Number';
            }

            $fp_results += New-Object -TypeName PSObject -Property @{
                "Name"              = $name
                "State"             = $entry.'Skater 1 State/Territory'
                "Membership Number" = $member_num
                "Division"          = $division
                "Element 1"         = $entry.'Element 1'
                "Element 2"         = $entry.'Element 2'
                "Element 3"         = $entry.'Element 3'
                "Element 4"         = $entry.'Element 4'
                "Element 5"         = $entry.'Element 5'
                "Element 6"         = $entry.'Element 6'
                "Element 7"         = $entry.'Element 7'
                "Element 8"         = $entry.'Element 8'
                "Element 9"         = $entry.'Element 9'
                "Element 10"        = $entry.'Element 10'
                "Element 11"        = $entry.'Element 11'
                "Element 12"        = $entry.'Element 12'
                "Element 13"        = $entry.'Element 13'
            }
        }
    }

    if ($fp_results.Count -gt 0)
    {
        $fp_results |
        Select-Object "Name", "State", "Membership Number", "Division",
                      "Element 1", "Element 2", "Element 3", "Element 4", "Element 5",
                      "Element 6", "Element 7", "Element 8", "Element 9", "Element 10",
                      "Element 11","Element 12", "Element 13" | 
        export-csv -path $fp_ppc_file -Force -NoTypeInformation
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

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_certificates
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_certificates($entries, $folder)
{
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
        $name     = ConvertTo-CapitalizedName -name $_.'Skater 1 Name'
        $category = $_.'Division'.Split(";")[0].trim()
        $division = $_.'Division'.Split(";")[1].trim()

        #Write-Host "Name: $name"

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

        $results += New-Object -TypeName PSObject -Property @{
            "Name"     = $name
            "Division" = $division
        }

        if ($category -match "Couple")
        {
            $name = ConvertTo-CapitalizedName -name $_.'Skater 2 Name'
            $results += New-Object -TypeName PSObject -Property @{
                "Name"     = $name
                "Division" = $division
            }
        }
    }

    #Write-Host "Number of entries = $($entries.Count)"
    $results | Select-Object "Name", "Division"
    $results | Select-Object "Name", "Division" | export-csv -path $input_csv -Force -NoTypeInformation -Encoding UTF8

    Invoke-MailMerge -template $template -datasource $input_csv -destination $folder
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_skating_schedule
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_skating_schedule($entries, $folder)
{
    Write-Host "Generating Skating Schedule"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $schedule1 = [System.IO.Path]::Combine($folder, "Schedule1.csv")
    $schedule2 = [System.IO.Path]::Combine($folder, "Schedule2.csv")

    if (Test-Path $schedule1) { Remove-Item -Path $schedule1 }
    if (Test-Path $schedule2) { Remove-Item -Path $schedule2 }
    
    $divhash = @{}
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
        $category  = $div.Split(";")[0].trim()
        $division  = $div.Split(";")[1].trim()

        if ($category.StartsWith("Aussie Skate (Half"))
        {
            $warmup = 3
            $performance_time = 2
        }
        elseif ($category.StartsWith("Aussie Skate (Full"))
        {
            $warmup = 3
            $performance_time = 3
        }
        elseif ($category -eq 'Adult')
        {
            $warmup = 4
            $performance_time = 4
            $division = "Adult $division"
        }
        elseif ($category -eq 'Singles' -and $division -eq 'Preliminary')
        {
            $warmup = 4
            $performance_time = 3            
        }
        elseif ($category -eq 'Singles' -and $division -eq 'Elementary')
        {
            $warmup = 4
            $performance_time = 4                
        }
        else
        {
            $warmup = 5
            $performance_time = 4
        }

        #Write-Host "Category: $category,Division: $division"
        Add-Content -Path $schedule1 -Value "Category: $category,Division: $division"

        $count = $divhash.Item($div).Count
        $num_warmup_groups = [Math]::Ceiling($count/$MAX_WARMUP_GROUP_SIZE)
        Add-Content -path $schedule1 -Value "Num Entries: $count"
        "Warmup Time: ($num_warmup_groups x $warmup) = {0} minutes"      -f ($num_warmup_groups * $warmup) | Add-Content -Path $schedule1 
        "Performance time = ($count x $performance_time) = {0} minutes " -f ($count * $performance_time)   | Add-Content -Path $schedule1 

        Add-Content -path $schedule1 -Value "First Name,Last Name,State,Coach Name,Other Coach Names,Music Title"

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

            "`"{0}`",`"{1}`",`"{2}`",`"{3}`",`"{4}`",`"{5}`"" -f $_.'First Name', $_.'Last Name', $_.'Skater 1 State/Territory', $_.'Primary Coach Name:',  $_.'Other Coach Names:', $music_title.Trim() | Add-Content -path $schedule1
        }
        Add-Content -Path $schedule1 -Value ""
    }

}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_division_counts
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_division_counts($entries, $folder)
{
    Write-Host "Generating Division Counts Spreadsheet"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $divcounts_csv = [System.IO.Path]::Combine($folder, "division_medal_and_trophy_counts.csv")

    if (Test-Path $divcounts_csv) { Remove-Item -Path $divcounts_csv }
    
    $gender = @{ Female = "Ladies"; Male = "Men" }

    $divhash = @{}
    foreach ($entry in $entries)
    {
        $division = $entry.'Division'

        if ($division -notmatch "Aussie Skate" -and $division -notmatch "Dance")
        {
            try
            {
                $division += " " + $gender[$entry.'Skater 1 Gender']
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

#    $divhash.GetEnumerator() | Sort Name

    Add-Content -Path $divcounts_csv -Value "Category,Division,# Skaters,Bronze Medal,Silver Medal,Gold Medal,Bronze Trophy,Silver Trophy,Gold Trophy"

    $trophy_count = @{ gold = 0; silver = 0; bronze = 0 }
    $medal_count  = @{ gold = 0; silver = 0; bronze = 0 }

    foreach ($div in $divhash.Keys | Sort)
    {
        $category  = $div.Split(";")[0].trim()
        $division  = $div.Split(";")[1].trim()

        $count = $divhash.Item($div).Count
        $medal_gold = ""
        $medal_silver = ""
        $medal_bronze = ""
        $trophy_gold = ""
        $trophy_silver = ""
        $trophy_bronze = ""

        if ($category.StartsWith("Aussie"))
        {
            if ($count -ge 1) { $medal_gold   = "x"; $medal_count.gold++   }
            if ($count -ge 2) { $medal_silver = "x"; $medal_count.silver++ }
            if ($count -ge 3) { $medal_bronze = "x"; $medal_count.bronze++ }
        }
        else
        {
            if ($count -ge 1) { $trophy_gold   = "x"; $trophy_count.gold++   }
            if ($count -ge 2) { $trophy_silver = "x"; $trophy_count.silver++ }
            if ($count -ge 3) { $trophy_bronze = "x"; $trophy_count.bronze++ }
        }
        "{0},{1},{2},{3},{4},{5},{6},{7},{8}" -f $category, $division, $count, $medal_bronze, $medal_silver, $medal_gold, $trophy_bronze, $trophy_silver, $trophy_gold | Add-Content -Path $divcounts_csv
    }

    ",,TOTAL:,{0},{1},{2},{3},{4},{5}" -f $medal_count.bronze, $medal_count.silver, $medal_count.gold, $trophy_count.bronze, $trophy_count.silver, $trophy_count.gold | Add-Content -Path $divcounts_csv
}

#--------------------------------------------------------------------------------
#  FUNCTION:  generate_engraving_schedule
#--------------------------------------------------------------------------------
function generate_engraving_schedule($entries, $folder)
{
    Write-Host "Generating Engraving Schedule"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $medalPath = [System.IO.Path]::Combine($folder, "${Competition} - MEDALS.docx")
    $trophyPath = [System.IO.Path]::Combine($folder, "${Competition} - TROPHIES.docx")

    if (Test-Path $medalPath)  { Remove-Item -Path $medalPath  }
    if (Test-Path $trophyPath) { Remove-Item -Path $trophyPath }
    
    $gender = @{ Female = "Ladies"; Male = "Men" }

    $divhash = @{}
    foreach ($entry in $entries)
    {
        $division = $entry.'Division'

        if ($division -notmatch "Aussie Skate" -and $division -notmatch "Dance")
        {
            try
            {
                $division += " " + $gender[$entry.'Skater 1 Gender']
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
    $MedalWord         = New-Object -ComObject Word.Application
    $MedalWord.Visible = $False
    $MedalDoc     = $MedalWord.Documents.Add()
    $MedalTable = $MedalDoc.Tables.Add($MedalWord.Selection.Range(),2,4)
    $MedalTable.Range.Style = "No Spacing"
    $MedalTable.Borders.Enable = $True
    $MedalTable.Rows(2).Cells(1).Range.Text = "Division"
    $MedalTable.Rows(2).Cells(1).Range.Bold = $True

    # Create Trophy Table
    $TrophyWord   = New-Object -ComObject Word.Application
    $TrophyWord.Visible = $False
    $TrophyDoc   = $TrophyWord.Documents.Add()
    $TrophyTable = $TrophyDoc.Tables.Add($TrophyWord.Selection.Range(),2,4)
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

    foreach ($div in $divhash.Keys | Sort)
    {
        $category  = $div.Split(";")[0].trim()
        $division  = $div.Split(";")[1].trim()

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

        if ($count -ge 1) { $Row.Cells(2).Range.Text = "${Engraving_Title}`r$division`r1st Place" }
        if ($count -ge 2) { $Row.Cells(3).Range.Text = "${Engraving_Title}`r$division`r2nd Place" }
        if ($count -ge 3) { $Row.Cells(4).Range.Text = "${Engraving_Title}`r$division`r3rd Place" }
    }

    # Save the documents
    $MedalDoc.SaveAs($medalPath,   [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
    $TrophyDoc.SaveAs($trophyPath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
    $MedalDoc.close()
    $TrophyDoc.close()
    $MedalWord.quit()
    $TrophyWord.quit()

    # Cleanup the memory
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalDoc)    | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyDoc)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalWord)   | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyWord)  | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($MedalTable)  | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TrophyTable) | Out-Null
    Remove-Variable MedalDoc, TrophyDoc, MedalWord, TrophyWord, Row, MedalTable, TrophyTable
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_skater_email_list
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_skater_email_list($entries, $folder)
{
    Write-Host "Generating Skater Email List"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $filename = [System.IO.Path]::Combine($folder, "skater_email_list.txt")

    $list = @()
    foreach ($entry in $entries)
    {
        $list += $entry.'Skater 1 Contact E-mail'
    }

    if (Test-Path $filename) { Remove-Item -Path $filename }

    $list | Get-Unique | Out-File $filename
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_coach_email_list
#
# DESCRIPTION:
##--------------------------------------------------------------------------------
function generate_coach_email_list($entries, $folder)
{
    Write-Host "Generating Coach Email List"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $filename = [System.IO.Path]::Combine($folder, "coach_email_list.txt")

    $list = @()
    foreach ($entry in $entries)
    {
        $list += $entry.'Primary Coach E-mail:'
    }

    if (Test-Path $filename) { Remove-Item -Path $filename }

    $list | Get-Unique | Out-File $filename
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_registration_list
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_registration_list($entries, $folder)
{
    Write-Host "Generating Registration List"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $list_csv = [System.IO.Path]::Combine($folder, "registration_list.csv")

    $hash = @{}
    foreach ($entry in $entries)
    {
        $surname = ConvertTo-CapitalizedName -name $entry.'Last Name'
        if (!$hash.ContainsKey($surname))
        {
            $hash[$surname] = @{}
        }
        $firstname = ConvertTo-CapitalizedName -name $entry.'First Name'
        if (!$hash[$surname].ContainsKey($firstname))
        {
            $hash[$surname].Add($firstname, $true)
        }

        if ($entry.Division -match 'Dance')
        {
            $surname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name (Last Name)'
            if (!$hash.ContainsKey($surname))
            {
                $hash[$surname] = @{}
            }
            $firstname = ConvertTo-CapitalizedName -name $entry.'Skater 2 Name (First Name)'
            if (!$hash[$surname].ContainsKey($firstname))
            {
                $hash[$surname].Add($firstname, $true)
            }
        }
    }

    if (Test-Path $list_csv) { Remove-Item -Path $list_csv }

    Add-Content -Path $list_csv -Value "Skater Surname, Skater First Name, Water, Fruit"
    foreach ($surname in $hash.Keys | Sort)
    {
        $name = $hash.Item($surname)
        foreach ($firstname in $name.Keys | sort)
        {
            "{0},{1}" -f $surname, $firstname | Add-Content -Path $list_csv
        }
    }
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_volunteer_spreadsheet
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_volunteer_spreadsheet($entries, $folder)
{
    Write-Host "Generating Volunteer Spreadsheet"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $spreadsheet = [System.IO.Path]::Combine($folder, "NFC_volunteer.csv");

    $results = @()
    foreach ($entry in $entries)
    {
        if (-not [String]::IsNullOrEmpty($entry.'I am able to assist with the following tasks:'))
        {
            $category  = $entry.Division.Split(";")[0].trim()
            $division  = $entry.Division.Split(";")[1].trim()

            if ($category -eq "Adult")
            {
                $division = "Adult ${division}"
            }

            $results += New-Object -TypeName PSObject -Property @{
                    "Name"              = $entry.'Skater 1 Name'
                    "Division"          = $division
                    "Volunteer Name"    = $entry.'Volunteer Name'
                    "Volunteer E-mail"  = $entry.'Volunteer E-mail'
                    "Volunteer Phone"   = $entry.'Volunteer Contact Mobile'
                    "Availability"      = $entry.'Availability:'
                    "Roles"             = $entry.'I am able to assist with the following tasks:'
                    "Other Notes"       = $entry.'Other Notes:'
            }
        }
    }

    if ($results.Count -gt 0)
    {
        $results | export-csv -path $spreadsheet -Force -NoTypeInformation
    }
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_payment_spreadsheet
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_payment_spreadsheet($entries, $folder)
{
    Write-Host "Generating Payment Spreadsheet"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $spreadsheet = [System.IO.Path]::Combine($folder, "payments.csv");

    $results = @()
    foreach ($entry in $entries)
    {
        $category  = $entry.Division.Split(";")[0].trim()
        $division  = $entry.Division.Split(";")[1].trim()

        if ($category -eq "Adult")
        {
            $division = "Adult ${division}"
        }

        $results += New-Object -TypeName PSObject -Property @{
                "Skater Name"                     = $entry.'Skater 1 Name'
                "Division"                        = $division
                "Parent/Guardian (if applicable)" = $entry.'Parent/Guardian Name: (First Name)' + ' ' + $entry.'Parent/Guardian Name: (Last Name)'
                "Payment Due (AUD)"               = $entry.'Payment due (AUD)'
                "Direct Debit Receipt"            = $entry.'Direct Debit Receipt'
        }
    }

    if ($results.Count -gt 0)
    {
        $results | Select-Object "Skater Name", "Division", "Parent/Guardian (if applicable)", "Payment Due (AUD)", "Direct Debit Receipt" | export-csv -path $spreadsheet -Force -NoTypeInformation
    }
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_coach_skaters_list
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_coach_skaters_list($entries, $folder)
{
    Write-Host "Generating Coach/Skaters List"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $outfile = [System.IO.Path]::Combine($folder, "coach_skaters.txt")

    if (Test-Path $outfile) { Remove-Item -Path $outfile }

    $hash = @{}
    foreach ($entry in $entries)
    {
        $coach_name  = $entry.'Primary Coach Name:'
        $coach_email = $entry.'Primary Coach E-mail:'

        #Write-Host "---------------------------"
        #Write-Host "Skater: $($entry.'Skater 1 Name')"
        #Write-Host "Division: $($entry.'Division')"
        #Write-Host "Coach Name: $coach_name"
        #Write-Host "Coach Email: $coach_email"

        if (!$hash.ContainsKey($coach_name))
        {
            $hash[$coach_name] = @()
        }
        $hash[$coach_name] += $entry.'Skater 1 Name' + " ($($entry.Division))"
    }

    Add-Content -Path $outfile -Value "Coach Name: Skaters"

    foreach ($c in $hash.Keys | Sort)
    {
        "{0}: {1}" -f $c, ($hash[$c] -join ", ") | Add-Content -Path $outfile
    }
}

#--------------------------------------------------------------------------------
# FUNCTION:
#  generate_photo_permission_list
#
# DESCRIPTION:
#
#--------------------------------------------------------------------------------
function generate_photo_permission_list($entries, $folder)
{
    Write-Host "Generating Photo Permission List"

    if ((Test-Path -Path $folder -ErrorAction SilentlyContinue) -eq $false)
    {
        New-Item $folder -Type Directory | Out-Null
    }

    $spreadsheet = [System.IO.Path]::Combine($folder, "photo_permissions.csv");

    $results = @()
    foreach ($entry in $entries)
    {
        $category = $entry.Division.Split(";")[0].trim()
        $division = $entry.Division.Split(";")[1].trim()
        $results += New-Object -TypeName PSObject -Property @{
                "Category" = $category
                "Division" = $division
                "Skater 1 Name" = $entry.'Skater 1 Name'
                "Skater 2 Name" = $entry.'Skater 2 Name'
                "ACTISA granted permission to use photos" = $entry.'I give permission for the Australian Capital Territory Ice Skating Association (ACTISA) to take photographs of myself/my child, and use the footage for promotional purposes on the official ACTISA website and social media.'
        }
    }

    if ($results.Count -gt 0)
    {
        $results | Select-Object "Category", "Division", "Skater 1 Name", "Skater 2 Name", "ACTISA granted permission to use photos" | export-csv -path $spreadsheet -Force -NoTypeInformation
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

# prompt the user to specify location
$comp_folder     = Find-Folders -title "Select the Competition folder (default=$comp_folder)" -default $comp_folder
$template_folder = Find-Folders -title "Select the MailMerge Template folder (default=$template_folder)" -default $template_folder

Push-Location $comp_folder

foreach ($f in ('Submissions','Music','PPC','Certificates','Schedule'))
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

$entries = Get-SubmissionEntries -url $google_sheet_url

foreach ($entry in $entries)
{
    process_music_files -entry $entry -submissionFullPath $submissionFullPath -music_folder $music_folder
}

Write-Host "Number of entries = $($entries.Count)`n" -ForegroundColor Yellow

generate_ppc_forms             -entries $entries -folder $ppc_folder
generate_certificates          -entries $entries -folder $certificate_folder
generate_skating_schedule      -entries $entries -folder $schedule_folder
generate_division_counts       -entries $entries -folder $comp_folder
generate_engraving_schedule    -entries $entries -folder $comp_folder
generate_registration_list     -entries $entries -folder $comp_folder
generate_volunteer_spreadsheet -entries $entries -folder $comp_folder
generate_payment_spreadsheet   -entries $entries -folder $comp_folder
generate_skater_email_list     -entries $entries -folder $comp_folder
generate_coach_email_list      -entries $entries -folder $comp_folder
generate_coach_skaters_list    -entries $entries -folder $comp_folder
generate_photo_permission_list -entries $entries -folder $comp_folder
New-ProofOfAgeAndMemberships -entries $entries -folder $comp_folder -format 'xlsx'

#Read-Host -Prompt "Press Enter to exit"
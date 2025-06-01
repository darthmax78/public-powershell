<#
    .SYNOPSIS
    Read all Movie Files from a specific Location and catalog them as a csv file, after
    analyzing them with the Get-MediaInfo Module.
        
    .DESCRIPTION
    An specific Location is scanned for all Movie Files. The File Name, File Size, 
    File Path and the Date of Creation are stored in a csv file. The csv file is 
    imported into an Excel File. It also checks, if the Get-MediaInfo Module is installed.
    If not, it will prompt the user to install it and provide a link to the GitHub repository.
    If the Module is installed, it will continue to scan the specified movie location for all movie files
    
    .REQUIREMENTS
    The script requires the following modules:
    - Export-Csv
    - Get-MediaInfo
    
    .NOTES
    *) v 0.1 Initial Version. Basic Function - First functions.
    *) v 0.2 Added Check for Get-MediaInfo Module.
    *) v 0.5 Added Progress Bar and improved error handling.
    *) v 0.6 Sorting the Movie Files; Optimizing the csv Output; Reading the File Name correctly.
    *) v 0.7 Further Sorting and Coloring in the Output Table; Added Resolution Detection and Encoding Date.
    *) v 0.8 Smaller Corrections in the Output Table; Fixing for Title / Year / Edition Detection; Frame Rate Detection.
    *) v 0.9 Added Audio Channel Detection; Smaller Corrections in the Output Table; Added csv Export.
    *) v 1.0 Added further Audio Analyis;

    .AUTHOR
    Magnus Witzik
#>

# Basic Function, which loads all necessary variables and checks if the Get-MediaInfo Module is installed.
function check_variables
{
    Clear-Host
    Write-Host "Checking Variables..." -ForegroundColor Cyan
    $global:movie_location          = "\\colonial-one.opti-net.at\Filme\Movies\"
    $global:movie_csv_file          = "\\colonial-one.opti-net.at\Skripte\Auswertungen"
    $global:date_scanning           = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")  

    # Checking if the Get-MediaInfo Module is installed
    try
    {
        Get-MediaInfo -ErrorAction Stop | Out-Null
    }
    catch
    {
        if ( $error[0].Exception.Message -like "*Get-MediaInfo*" )
        {
            Write-Host "The Module Get-MediaInfo is not installed! Please install it first!" -ForegroundColor Red
            $url    = "https://github.com/stax76/Get-MediaInfo"
            Set-Clipboard -Value $url
            Write-Host "The URL to the Module is copied to the Clipboard: " -ForegroundColor Yellow -NoNewline; Write-Host $url -ForegroundColor Cyan
            exit
        }
    }

    # if the Module is installed, then we can continue
    Write-Host "Get-MediaInfo Module is installed." -ForegroundColor Green    
    $global:all_movies              = Get-ChildItem -Path $global:movie_location -Recurse -File -Include *.mkv, *.mp4, *.avi, *.mov, *.wmv, *.flv, *.webm | Sort-Object -Property Name
}

# The Main Function, which analyzes the movies and stores the information in a csv file.
function analyzing_movies
{
    if ( $global:all_movies.Count -eq 0 )
    {
        Write-Host "No Movies found in the specified location." -ForegroundColor Yellow
        return
    }
    elseif ( $global:all_movies.Count -gt 0 )
    {
        Write-Host "Found $($global:all_movies.Count) Movies in the specified location." -ForegroundColor Green
        $counter                = 0
        $global:movie_table     = @()

        $global:all_movies | ForEach-Object `
        {
            $counter++
            $counter_percent    = [math]::Round(($counter / $global:all_movies.Count) * 100, 2)
            Write-Progress -Activity "Analyzing Movies" -Status "Processing $($_.Name) ($counter_percent%)" -PercentComplete $counter_percent
            $movie_file_info    = $_
            $movie_media        = Get-MediaInfo -Path $movie_file_info.FullName -ErrorAction SilentlyContinue
            $bild_weite         = $movie_media.Width
            $bild_hoehe         = $movie_media.Height
            $auflösung          = [STRING]$bild_weite+"x"+[STRING]$bild_hoehe

            # Write the Resolution into a more readable format (like PAL,HD 720, HD 1080, UHD 4K)
            if ( ($bild_weite -in 500..600) -and ($bild_hoehe -in 300..450) )
            {
                $standard       = "PAL"
            }
            elseif ( ($bild_weite -in 700..720) -and ($bild_hoehe -in 320..580) )
            {
                $standard       = "DVD"
            }
            elseif ( ($bild_weite -in 960..1280) -and ($bild_hoehe -in 400..720) )
            {
                $standard       = "HD 720"
            }
            elseif ( ($bild_weite -in 1320..1960) -and ($bild_hoehe -in 790..1100) )
            {
                $standard       = "HD 1080"
            }
            elseif ( ($bild_weite -in 2500..4096) -and ($bild_hoehe -in 1600..2180) )
            {
                $standard       = "UHD 4K"
            }
            else { }

            # Defining the Movie Year and Edition Information out of the File Name
            if ( $movie_file_info.BaseName -match "edition" )
            {
                $edition        = ($movie_file_info.BaseName).Split('{')[1] -Replace('}','') -Replace ('Edition-','')
            }
            else { }

            if ( $movie_file_info.BaseName -match "(\d{4})" )
            {
                $year           = (($movie_file_info.BaseName).Split('(')[1]).Split('{')[0] -replace '[()]', ''
            }
            else
            {
                $year           = "Unbekannt"
            }

            # Erfasst die Audio Auswertung
            $audio_channels         = ($movie_media.AudioCodec.Split('/')).Count
            if ( $audio_channels -eq 1 )
            {
                $audio_format       = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Format/String'
                $audio_channels     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Channel(s)/String'
                $audio_language     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Language/String'
                
                # Renaming, if the Channel Language is in Short Form
                if ( $audio_language -eq "de|ger" )
                {
                    $audio_language = "Deutsch"
                }
                elseif ( $audio_language -eq "en|eng" )
                {
                    $audio_language = "Englisch"
                }
                else { }

                $audio_channel_1    = $audio_language + " - " + $audio_channels + " - " + $audio_format
            }
            elseif ( $audio_channels -gt 2 )
            {
                $audio_format       = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Format/String'
                $audio_channels     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Channel(s)/String'
                $audio_language     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 2 -Parameter 'Language/String'
                
                # Renaming, if the Channel Language is in Short Form
                if ( $audio_language -eq "de|ger" )
                {
                    $audio_language = "Deutsch"
                }
                elseif ( $audio_language -eq "en|eng" )
                {
                    $audio_language = "Englisch"
                }
                else { }

                $audio_channel_1    = $audio_language + " - " + $audio_channels + " - " + $audio_format

                # Analyzing the second Audio Channel, if it exists
                $audio_format       = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 3 -Parameter 'Format/String'
                $audio_channels     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 3 -Parameter 'Channel(s)/String'
                $audio_language     = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Audio -Index 3 -Parameter 'Language/String'
                
                # Renaming, if the Channel Language is in Short Form
                if ( $audio_language -eq "de|ger" )
                {
                    $audio_language = "Deutsch"
                }
                elseif ( $audio_language -eq "en|eng" )
                {
                    $audio_language = "Englisch"
                }
                else { }

                $audio_channel_2    = $audio_language + " - " + $audio_channels + " - " + $audio_format
            }
            else 
            {
                $audio_channel_1    = "Unbekannt"
            }

            # Convert the Encoding Date into a more readable format
            $date_created       = ((Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Encoded_Date").Split("/")[0]) -replace ("[A-Za-z]","")

            $movie_info = [PSCustomObject]@{
            "Film Titel"        = ($movie_file_info.BaseName).Split('(')[0]
            "Film Edition"      = $edition
            "Jahr"              = $year
            "Laufzeit"          = [TIMESPAN]::FromMilliseconds((Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Duration"))
            "Standard"          = $standard
            "Auflösung"         = $auflösung
            "Frame Rate"        = $movie_media.FrameRate
            "Format"            = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Video -Parameter "Format"                  
            "Audio Spuren"      = $audio_channels
            "Audio Spur 1"      = $audio_channel_1
            "Audio Spur 2"      = $audio_channel_2
            "Film Erstellt"     = $date_created 
            "Programm"          = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Encoded_Application/String"
            "Film Größe"        = human_readable -Bytes ($movie_file_info.Length)
            }

            $global:movie_table += $movie_info
            $movie_info | Export-Csv -Path "$global:movie_csv_file\Movie-Report_$global:date_scanning.csv" -Append -NoTypeInformation -Encoding utf8BOM -Delimiter ";"
        }
    }
}

function show_movie_report
{
    $global:movie_table | Format-Table -Property * -AutoSize | Out-String -Stream | ForEach-Object `
    {
        $studiostatecolor = `
            if ($_ -match "HD 1080") {@{'ForegroundColor' = 'Blue'}}
            elseif ($_ -match "UHD 4K") {@{'ForegroundColor' = 'Magenta'}}
            elseif ($_ -match "DVD") {@{'ForegroundColor'='Yellow'}}
            elseif ($_ -match "HD 720") {@{'ForegroundColor'='Red'}}
            elseif ($_ -match "PAL|SD") { @{ 'ForegroundColor' = 'Red'} }
            else {@{}}
            Write-Host @studiostatecolor $_
    }
}

check_variables
analyzing_movies
show_movie_report
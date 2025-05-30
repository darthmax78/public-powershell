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

    .AUTHOR
    Magnus Witzik
#>

# Basic Function, which loads all necessary variables and checks if the Get-MediaInfo Module is installed.
function check_variables
{
    Clear-Host
    Write-Host "Checking Variables..." -ForegroundColor Cyan
    $global:movie_location          = "\\colonial-one.opti-net.at\Filme\Movies\"
    $global:movie_csv_file          = "\\colonial-one.opti-net.at\Skripte\Auswertungen\"

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
            if ( $bild_weite -in 500..600 )
            {
                $auflösung      = "PAL"+$auflösung
            }
            elseif ( $bild_weite -in 700..800 )
            {
                $auflösung      = "HD 720"+$auflösung
            }
            elseif ( ($bild_weite -in 1320..1960) -and ($bild_hoehe -in 1000..1100) )
            {
                $auflösung      = "HD 1080"+$auflösung
            }
            else { }

            # Convert the Encoding Date into a more readable format
            $date_created       = ((Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Encoded_Date").Split("/")[0]) -replace ("[A-Z]{3}","")

            if ( $movie_media )
            {
                $movie_info = [PSCustomObject]@{
                    "Film Titel"        = ($movie_file_info.BaseName).Split('(')[0]
                    "Film Edition"      = $edition
                    "Jahr"              = ($movie_file_info.BaseName).Split('(')[1] -replace '[()]', '' -replace '[^\d]', ''
                    "Laufzeit"          = [TIMESPAN]::FromMilliseconds((Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Duration"))
                    "Auflösung"         = $auflösung
                    "Format"            = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind Video -Parameter "Format"                  
                    "Film Erstellt"     = $date_created 
                    "Programm"          = Get-MediaInfoValue -Path $movie_file_info.FullName -Kind General -Parameter "Encoded_Application/String"
                    "Film Größe"        = human_readable -Bytes ($movie_file_info.Length)
                    
                    AudioCodec       = $movie_media.AudioCodec
                    Resolution       = $movie_media.VideoResolution
                }

                $global:movie_table += $movie_info
            }
            else
            {
                Write-Host "Could not analyze movie: $($_.Name)" -ForegroundColor Red
            }
        }
    }
}

check_variables
analyzing_movies
$global:movie_table | Format-Table -AutoSize
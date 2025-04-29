function human_readable
{
    param([DOUBLE]$Bytes)
    function convert
    {
        if ( ($Bytes -ge 1kb) -and ($Bytes -lt 1mb) ) { "{0:n3} KB" -f ($Bytes/1Kb) }
        elseif ( ($Bytes -ge 1mb) -and ($Bytes -lt 1gb) ) { "{0:n3} MB" -f ($Bytes/1Mb) }
        elseif ( ($Bytes -ge 1gb) -and ($Bytes -lt 1tb) ) { "{0:n3} GB" -f ($Bytes/1Gb) }
        elseif ( ($Bytes -ge 1tb) -and ($Bytes -lt 1pb) ) { "{0:n3} TB" -f ($Bytes/1Tb) }
        elseif ( ($Bytes -ge 1pb) ) { "{0:n3} PB" -f ($Bytes/1Pb) }
        else { "{0} Bytes" -f $Bytes }
    }

    function convert_negativ
    {
        if ( ($Bytes -ge 1kb) -and ($Bytes -lt 1mb) ) { "-{0:n3} KB" -f ($Bytes/1Kb) }
        elseif ( ($Bytes -ge 1mb) -and ($Bytes -lt 1gb) ) { "-{0:n3} MB" -f ($Bytes/1Mb) }
        elseif ( ($Bytes -ge 1gb) -and ($Bytes -lt 1tb) ) { "-{0:n3} GB" -f ($Bytes/1Gb) }
        elseif ( ($Bytes -ge 1tb) -and ($Bytes -lt 1pb) ) { "-{0:n3} TB" -f ($Bytes/1Tb) }
        elseif ( ($Bytes -ge 1pb) ) { "-{0:n3} PB" -f ($Bytes/1Pb) }
        else { "-{0} Bytes" -f $Bytes }
    }

    if ( $Bytes -lt '0') 
    {
        [STRING]$Bytes  = $Bytes
        $Bytes          = $Bytes.Replace('-','')
        [DOUBLE]$Bytes   = $Bytes

        convert_negativ
    }
    else
    {
        convert
    }
}
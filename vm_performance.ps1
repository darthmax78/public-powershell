<#
.SYNOPSIS
    Dieses Programm dient zur Erfassung von Performance Daten aller virtuellen Maschinen einer VMware Umgebung.
.DESCRIPTION
    In erster Linie dient dieses Programm dazu, damit Performance Daten für Parameter von CPU (Ready Time; Wait-States; Usage), RAM (Consumed; Active; Ballooned; Swapped) und Datenspeicher (I/O-Latenz; IOPS; Durchsatz) gesammelt werden können.
.NOTES
    v. 0.1  Die Anmelde Funktion wird implementiert, sowie die Abfrage aller VMs und Cluster.
            Alle VMs werden in einer Variablen gesammelt
            Alle Cluster, mit Ausnahme des Folder für "Sort Out" Hardware, werden in einer globalen Variable gesammelt

    v. 0.2  Auswertung für Cluster wird implementiert ( Grundsätzliche Daten, samt Überbuchung )
            Progress Bar wird für den Cluster eingebaut ( mit Restzeit Anzeige )
            Sort Out für Cluster ohne Hosts eingebaut
            Funktion für Geo-Redundant getagte Cluster eingebaut

    v. 0.3  Erste Schritte für die VM Auswertung werden implementiert

.AUTHOR
    Magnus Witzik

.REQUIREMENTS
    VMware PowerCLI Modul
    Verbindung zu einem vCenter Server
    Powershell 7+

#>
Clear-Host

function get_variable
{
    $global:all_vms         = Get-VM | Where-Object { ($_.PowerState -match "PoweredOn") -and ($_.Name -notmatch "\AvCLS|_backup")} | Sort-Object Name
    # $all_vms | FT -AutoSize -Property Name, @{Name="Guest OS"; E={ $_.Guest.OSFullName}}, CreateDate, PowerState, NumCPU, MemoryGB, ProvisionedSpaceGB, UsedSpaceGB
    $global:all_clusters    = Get-Cluster | Where-Object { $_.ParentFolder -notmatch "Gen8" } | Sort-Object Name
    $global:all_clusters | ForEach-Object `
    {
        $host_count         = ($_ | Get-VMHost).Count
        if ( $host_count -eq 0 )
        {
            $name = $_.Name
            Write-Host "Cluster $($_.Name) wird übersprungen, da keine Hosts vorhanden sind." -ForegroundColor Yellow
            $global:all_clusters = $global:all_clusters | Where-Object { $_.Name -notmatch $name }
        }
        else { }
    }
}

function get_cluster_data
{
    $global:report_cluster  = @()
    $counter                = 0
    $estimated_total        = 1
    $global:all_clusters | Foreach-Object `
    {
        $counter++
        $percent            = (($counter) / $global:all_clusters.Count * 100)
        $start_time         = Get-Date
        Write-Progress -Activity "Cluster Daten werden gesammelt" -Status "Cluster: $($_.Name) $percent%" -PercentComplete $percent -SecondsRemaining $estimated_total
        $cluster_vm         = $_ | Get-VM | Where-Object { $_.Name -notmatch "\AvCLS" } | Sort-Object Name
        $cluster_host       = $_ | Get-VMHost | Where-Object { $_.ConnectionState -eq "Connected" } | Sort-Object Name
        $cluster_entry      = "" | Select-Object Cluster, VMs, "Physical CPU", "Virtual CPU", "CPU Overcommit", "Geo-Redundancy", "Physical RAM (GB)", "Virtual RAM (GB)", "RAM Overcommit"
        $cluster_entry.Cluster          = $_.Name
        $cluster_entry.VMs              = $cluster_vm.Count
        if ( $_ | Get-TagAssignment | Where-Object { $_.Tag -match "Geo-Redundant"}) { $cluster_entry."Geo-Redundancy" = "Ja" } else { $cluster_entry."Geo-Redundancy" = "Nein" }
        if ( $cluster_entry."Geo-Redundancy" -match "Ja" )
        {
            $cluster_entry."Physical CPU"       = [INT]($cluster_host | Measure-Object -Property NumCpu -Sum).Sum / 2
            $cluster_entry."Physical RAM (GB)"  = [INT]($cluster_host | Measure-Object -Property MemoryTotalGB -Sum).Sum / 2
        }
        else
        {
            $cluster_entry."Physical CPU"       = [INT]($cluster_host | Measure-Object -Property NumCpu -Sum).Sum
            $cluster_entry."Physical RAM (GB)"  = [INT]($cluster_host | Measure-Object -Property MemoryTotalGB -Sum).Sum
        }                
        $cluster_entry."Virtual CPU"            = [INT]($cluster_vm | Measure-Object -Property NumCpu -Sum).Sum
        $cluster_entry."Virtual RAM (GB)"       = [INT]($cluster_vm | Measure-Object -Property MemoryGB -Sum).Sum
        $cluster_entry."CPU Overcommit"         = "{0:N2}" -f (($cluster_entry."Virtual CPU" / $cluster_entry."Physical CPU"))
        $cluster_entry."RAM Overcommit"         = "{0:N2}" -f (($cluster_entry."Virtual RAM (GB)" / $cluster_entry."Physical RAM (GB)"))

        $global:report_cluster += $cluster_entry
        $end_time = Get-Date
        $runtime            = (New-TimeSpan -Start $start_time -End $end_time).TotalSeconds
        $estimated_total    = (($global:all_clusters.Count-$counter) * $runtime)
    }
}

function get_vm_data
{

}

get_variable
get_cluster_data
$global:report_cluster | FT -AutoSize -Property *
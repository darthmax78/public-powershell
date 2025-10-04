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
            Alle Metadaten der VMs werden gesammelt
            CPU und RAM Performance wird gesammelt
            RAM Performance finished
            Disk Latenz eingepflegt
            Datastore IOPS sind eingepflegt
            CPU Werte sind eingepflegt
    
    v. 0.4  Test ob die Auswertung mittels Abfrage Optimierung beschleunigt werden kann
            Korrektur der Abfrage aufgrund der Sammlung im Bucket
            Laufzeit Auswertung hinzugefügt
            Zeitstempel und globale Variable für den Export des Reports hinzugefügt
            Create Date der VM hinzugefügt
            Disk Durchsatz hinzugefügt

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
    # Variable für den Export des Reports
    $global:path_export     = "W:\Reports\VM_Performance"
    $global:date_export     = Get-Date -Format "yyyy-MM-dd_HH-mm"

    # Do not change anything below this line
    $global:all_vms         = Get-VM | Where-Object { ($_.PowerState -match "PoweredOn") -and ($_.Name -notmatch "\AvCLS|_backup")} | Sort-Object Name
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
    Write-Progress -Activity "Cluster Daten werden gesammelt" -Completed
}

function get_vm_data
{
    $global:report_vm       = @()
    $counter                = 0
    $estimated_total        = 1
    $global:all_vms | Foreach-Object `
    {
        $counter++
        $percent            = (($counter) / $global:all_vms.Count * 100)
        $start_time         = Get-Date
        Write-Progress -Activity "Daten aller VMs werden gesammelt" -Status "VM: $($_.Name) $percent%" -PercentComplete $percent -SecondsRemaining $estimated_total
        $vm_entry           = "" | Select-Object Name, "Guest OS", "Created", Cluster, Host, PowerState, NumCPU, MemoryGB, ProvisionedSpaceGB, UsedSpaceGB, "CPU Ready (%)", "CPU Co-Stop (ms)", "Cluster CPU Overbookin", "RAM Consumed (GB)", "RAM Active (GB)", "RAM Ballooned (GB)", "RAM Swapped (GB)", "Cluster RAM Overbookin", "Disk Write Latency (ms)", "Disk Read Latency (ms)", "Disk Read IOPS", "Disk Write IOPS", "Disk Read Throughput (MB/s)", "Disk Write Throughput (MB/s)"
        $vm_host            = $_.VMHost
        $vm_stats           = $_ | Get-Stat -Stat cpu.ready.summation, cpu.costop.summation, mem.consumed.average, mem.active.average, mem.vmmemctl.average, mem.swapped.average, virtualDisk.totalWriteLatency.average, virtualDisk.totalReadLatency.average, virtualDisk.numberReadAveraged.average, virtualDisk.numberWriteAveraged.average, virtualDisk.read.average, virtualDisk.write.average -Realtime

        $vm_entry.Name                  = $_.Name
        $vm_entry."Guest OS"            = $_.Guest.OSFullName
        $vm_entry."Created"             = $_.CreateDate
        $vm_entry.Cluster               = $vm_host.Parent
        $vm_entry.Host                  = $vm_host
        $vm_entry.PowerState            = $_.PowerState
        $vm_entry.NumCPU                = $_.NumCPU
        $vm_entry.MemoryGB              = $_.MemoryGB
        $vm_entry.ProvisionedSpaceGB    = [MATH]::ROUND($_.ProvisionedSpaceGB,2)
        $vm_entry.UsedSpaceGB           = [MATH]::ROUND($_.UsedSpaceGB,2)

        $vm_entry."CPU Ready (%)"               = [INT64]($vm_stats | Where-Object { ($_.MetricId -like 'cpu.ready.summation') -and ($_.Instance -like '') } | Measure-Object -Property Value -Sum).Sum
        $vm_entry."CPU Ready (%)"               = ($vm_entry."CPU Ready (%)"/$vm_entry.NumCPU/(20*1000))*100
        $vm_entry."CPU Co-Stop (ms)"            = [INT]($vm_stats | Where-Object { ($_.MetricId -like 'cpu.costop.summation') -and ($_.Instance -like '') } | Measure-Object -Property Value -Sum).Sum
        $vm_entry."Cluster CPU Overbookin"      = $global:report_cluster | Where-Object { $_.Cluster -match $vm_entry.Cluster } | Select-Object -ExpandProperty "CPU Overcommit"
        $vm_entry."Cluster RAM Overbookin"      = $global:report_cluster | Where-Object { $_.Cluster -match $vm_entry.Cluster } | Select-Object -ExpandProperty "RAM Overcommit"
        $vm_entry."RAM Consumed (GB)"           = [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'mem.consumed.average') } | Measure-Object -Property Value -Average).Average/1024)/1024,2)
        $vm_entry."RAM Active (GB)"             = [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'mem.active.average') } | Measure-Object -Property Value -Average).Average/1024)/1024,2)
        $vm_entry."RAM Ballooned (GB)"          = [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'mem.vmmemctl.average') } | Measure-Object -Property Value -Average).Average/1024)/1024,2)
        $vm_entry."RAM Swapped (GB)"            = [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'mem.swapped.average') } | Measure-Object -Property Value -Average).Average/1024)/1024,2)
        $vm_entry."Disk Write Latency (ms)"     = [MATH]::ROUND(($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.totalWriteLatency.average') } | Measure-Object -Property Value -Average).Average,2)
        $vm_entry."Disk Read Latency (ms)"      = [MATH]::ROUND(($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.totalReadLatency.average') } | Measure-Object -Property Value -Average).Average,2)
        $vm_entry."Disk Read IOPS"              = [MATH]::ROUND(($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.numberReadAveraged.average') } | Measure-Object -Property Value -Average).Average,2)
        $vm_entry."Disk Write IOPS"             = [MATH]::ROUND(($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.numberWriteAveraged.average') } | Measure-Object -Property Value -Average).Average,2)
        $vm_entry."Disk Read Throughput (MB/s)" = [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.read.average') } | Measure-Object -Property Value -Average).Average/1024),2)
        $vm_entry."Disk Write Throughput (MB/s)"= [MATH]::ROUND((($vm_stats | Where-Object { ($_.MetricId -like 'virtualDisk.write.average') } | Measure-Object -Property Value -Average).Average/1024),2)
        
        $global:report_vm   += $vm_entry
        $end_time = Get-Date
        $runtime            = (New-TimeSpan -Start $start_time -End $end_time).TotalSeconds
        $estimated_total    = (($global:all_vms.Count-$counter) * $runtime)
    }
    Write-Progress -Activity "Daten aller VMs werden gesammelt" -Completed
}   

$start_script   = Get-Date
get_variable
get_cluster_data
get_vm_data
# $global:report_cluster | Out-GridView
$global:report_vm | Out-GridView
$global:report_vm | Export-Csv -Path "$global:path_export\VM_Performance_$global:date_export.csv" -NoTypeInformation -Delimiter ";" -Encoding utf8BOM

$end_script     = Get-Date
$runtime_script = New-TimeSpan -Start $start_script -End $end_script
Write-Host "Script Laufzeit: $($runtime_script.TotalSeconds) Sekunden" -ForegroundColor Green
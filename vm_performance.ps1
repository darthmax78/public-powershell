<#
.SYNOPSIS
    Dieses Programm dient zur Erfassung von Performance Daten aller virtuellen Maschinen einer VMware Umgebung.
.DESCRIPTION
    In erster Linie dient dieses Programm dazu, damit Performance Daten für Parameter von CPU (Ready Time; Wait-States; Usage), RAM (Consumed; Active; Ballooned; Swapped) und Datenspeicher (I/O-Latenz; IOPS; Durchsatz) gesammelt werden können.
.NOTES
    v. 0.1  Die Anmelde Funktion wird implementiert, sowie die Abfrage aller VMs und Cluster.
.COMPONENT
    VMware PowerCLI
    Powershell 7+
.AUTHOR
    Magnus Witzik
#>


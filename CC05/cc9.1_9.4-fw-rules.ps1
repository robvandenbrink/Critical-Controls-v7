$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,IPV4Address
$fwinfo = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    $tempval = new-object psobject
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    $tempval = Get-NetFirewallProfile -PolicyStore activestore | select name, enabled, defaultinboundaction, DefaultOutboundAction
    $tempval | add-member -membertype noteproperty -name HostName -value $pc.dnshostname
    $tempval | add-member -membertype noteproperty -name OperatingSystem -value $pc.OperatingSystem
    $tempval | add-member -membertype noteproperty -name IpAddress -value $pc.IPV4Address
    $fwinfo += $tempval
}
}
$fwinfo | export-csv -path ./fwstate.csv




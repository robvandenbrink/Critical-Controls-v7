$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,IPV4Address
$fwinfo = @()
$fwrules = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    $tempval = new-object psobject
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    $Computer = $pc.DNSHostname
    # get state
    $tempval = Get-NetFirewallProfile -cimsession $Computer -PolicyStore activestore | select name, enabled, defaultinboundaction, DefaultOutboundAction
    # get rules
    $tempval2 = get-NetFirewallRule -cimsession $Computer 
    # add workstation fields
    $tempval2 | add-member -membertype noteproperty -name HostName -value $Computer
    $tempval2 | add-member -membertype noteproperty -name OperatingSystem -value $pc.OperatingSystem
    $tempval2 | add-member -membertype noteproperty -name IpAddress -value $pc.IPV4Address
    if ($tempval.count -gt 0) {
        $tempval | add-member -membertype noteproperty -name HostName -value $pc.dnshostname
        $tempval | add-member -membertype noteproperty -name OperatingSystem -value $Computer
        $tempval | add-member -membertype noteproperty -name IpAddress -value $pc.IPV4Address
        }
           
    $fwinfo += $tempval
    $fwrules += $tempval2
}
}
$fwinfo | export-csv -path ./fwstate.csv
$fwrules | export-csv -path ./fwrules.csv




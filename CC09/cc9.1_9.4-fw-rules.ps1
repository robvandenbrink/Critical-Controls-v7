$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,IPV4Address
$fwinfo = @()
$fwrules = @()
# might need to supply domain credentials to authentication the CIM Session
# either input them manually here, or supply them in a stored credential file
# $creds = Get-Credential
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    $tempval = new-object psobject
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    # depending on the situation, you may need to add "-credential $creds" to the CIMSession, 
    # where $creds holds the required credentials.  This can be inputted manually (see above)
    # or from a stored credential file
    # test without first in your environment, may not be required
    $TGTComputer = New-CIMSession -Computername $pc.DNSHostname
    # get state
    $tempval = Get-NetFirewallProfile -cimsession $TGTComputer -PolicyStore activestore | select name, enabled, defaultinboundaction, DefaultOutboundAction
    # get rules
    $tempval2 = get-NetFirewallRule -cimsession $TGTComputer 
    # add workstation fields
    $tempval2 | add-member -membertype noteproperty -name HostName -value $pc.DNSHostName
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




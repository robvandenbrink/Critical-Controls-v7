$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,LastLogonDate,IPV4Address
$AVApps = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    $Computer = $pc.DNSHostName
    $wmiQuery = "SELECT * FROM AntiVirusProduct"
    $AntiVirusProduct = Get-WmiObject -ComputerName $computer -Namespace "root\SecurityCenter2" -Query $wmiQuery  @psboundparameters 

    $AVapps += $AntiVirusProduct.displayName
}
}
$avapps | export-csv -path ./avapps.csv


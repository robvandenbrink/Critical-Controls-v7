$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,LastLogonDate,IPV4Address
$AVApps = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    $Computer = $pc.DNSHostName
    $AntiVirusProduct = Get-WmiObject -Namespace "root\SecurityCenter2" -Query $wmiQuery  @psboundparameters # -ErrorVariable myError -ErrorAction 'SilentlyContinue' -Computer $Computer

    $AVapps += $AntiVirusProduct
}
}
$avapps | export-csv -path ./avapps.csv


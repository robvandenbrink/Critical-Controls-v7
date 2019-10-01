$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,LastLogonDate,IPV4Address
$patchinfo = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
    $appsonpc = Get-WmiObject -Class Win32_Product -computername . | select pscomputername, vendor, name, version, installdate
    $domainapps += $appsonpc
}
}
$domainapps | export-csv -path ./domainapps.csv


$i = 1
$targets =get-adcomputer -filter * -Property DNSHostName
$vallist = @()

foreach ($targethost in $targets) {
  write-host $i  $targethost.DNSHostName
  if (Test-Connection -ComputerName $targethost.DNSHostName -count 2 -Quiet) {
    $SVClist += Get-WmiObject Win32_service -Computer $targethost.DNSHostName | select-object systemname, displayname, startname, state
    ++$i 
    }
  }
$SVClist | export-csv allservices.csv


$goodservices = @("LocalSystem","LocalService","NetworkService","NT AUTHORITY\LocalService","NT AUTHORITY\NetworkService")
$temp1 = $SVClist | where-object {$goodservices -notcontains $_.startname}
$badservices = $b | where-object { $_.startname.length -gt 0 }
$badservices | export-csv badservices.csv

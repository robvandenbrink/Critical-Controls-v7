$i = 1
$targets =get-adcomputer -filter * -Property DNSHostName
$vallist = @()

foreach ($targethost in $targets) {
  write-host $i  $targethost.DNSHostName
  if (Test-Connection -ComputerName $targethost.DNSHostName -count 2 -Quiet) {
    $vallist += Get-WmiObject Win32_service -Computer $targethost.DNSHostName | select-object systemname, displayname, startname, state
    ++$i 
    }
  }
$vallist | export-csv services.csv

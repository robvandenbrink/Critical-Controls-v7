import-module ActiveDirectory

function get-localadmin { 
  param ($strcomputer) 
  $admins = Gwmi win32_groupuser –computer $strcomputer  
  $admins = $admins |? {$_.groupcomponent –like '*"Administrators"'} 
  $admins |% { 
    $_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul 
    $matches[1].trim('"') + “\” + $matches[2].trim('"') 
  } 
}

$i = 1
$vallist = @()
$targets = Get-ADComputer -Filter * -Property DNSHostName
foreach ($targethost in $targets) {
  write-host $i  $targethost.DNSHostName
  if (Test-Connection -ComputerName $targethost.DNSHostName -count 2 -Quiet) {
    $admins = get-localadmin $targethost.DNSHostName
    foreach ($a in $admins) {
      $val = new-object psobject
      $val | add-member -membertype NoteProperty -name Hostname -value $targethost.name
      $val | add-member -membertype NoteProperty -name AdminID -value $a
      $vallist += $val
      }
  ++$i
  }
}
$vallist | export-csv -append localadminusers.csv

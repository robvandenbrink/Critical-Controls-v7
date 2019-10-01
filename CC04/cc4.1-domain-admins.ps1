$b = @()
$a = $()
'Domain Admins', 'Administrators', 'Enterprise Admins', 'Schema Admins', 'Server Operators', 'Backup Operators' | ForEach-Object {
    $groupName = $_
    $a = Get-ADGroupMember -Identity $_ -Recursive | Get-ADUser | Select-Object Name, samaccountname, DisplayName, @{n='GroupName';e={ $groupName }} 
    $b += $a
}
$b | export-csv alladmins.csv
$U_TO_S = @()
$events = @()

$days = -1
$filter = @{
    Logname = 'Security'
    ID = 4624
    StartTime =  [datetime]::now.AddDays($days)
    EndTime = [datetime]::now
}

# Get your ad information
$DomainName = (Get-ADDomain).DNSRoot
# Get all DC's in the Domain
$AllDCs = Get-ADDomainController -Filter * -Server $DomainName | Select-Object Hostname,Ipv4address,isglobalcatalog,site,forest,operatingsystem

foreach($DC in $AllDCs) {
    if (test-connection $DC.hostname -quiet) {
        # collect the events
        write-host "Collecting from server" $dc.hostname
        $events += Get-WinEvent -FilterHashtable $filter -ComputerName $DC.Hostname
        }
    }
 
# filter to network logins only (Logon Type 3), userid and ip address only
$b = $events | where { $_.properties[8].value -eq 3 } | `
     select-object @{Name ="user"; expression= {$_.properties[5].value}}, `
     @{name="ip"; expression={$_.properties[18].value} }

# filter out workstation logins (ends in $) and any other logins that won't apply (in this case "ends in ADFS")
# as we are collecting the station IP's, adding an OS filter to remove anything that includes "Server" might also be useful
# filter out duplicate username and ip combos
$c = $b | where { $_.user -notmatch 'ADFS$' } | where { $_.user -notmatch '\$$' } | sort-object -property user,ip -unique

# collect all user contact info from AD
# this assumes that these fields are populated
$userinfo = Get-ADUser -filter * -properties Mobile, TelephoneNumber | select samaccountname, name, telephonenumber, mobile, emailaddress | sort samaccountname

# combine our data into one "users to stations to contact info" variable
# any non-ip stn fields will error out - for instance "-".  This is not a problem
foreach ( $logevent in $c ) {
            $u = $userinfo | where { $_.samaccountname -eq $logevent.user }
            $tempobj = [pscustomobject]@{
                user = $logevent.user
                stn = [system.net.dns]::gethostbyaddress($logevent.ip).hostname
                mobile = $u.mobile
                telephone = $u.telephonenumber
                emailaddress = $u.emailaddress
                }
            $U_to_S += $tempobj
            }

# We can easily filter here to
# in this case we are filtering out Terminal Servers (or whatever) by hostname, remove duplicate entries
$user_to_stn_db = $U_TO_S | where { $_.stn -notmatch 'TS' } | sort-object -property stn, user -unique | sort-object -property user

$user_to_stn_db | Out-GridView

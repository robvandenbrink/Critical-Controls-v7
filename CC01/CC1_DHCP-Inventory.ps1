$subnets = @()
$dhcpinven = @()
$rawleases = @()
$inventoryexception = @()
$inventoryupdate = @()

#populate OUI table
$ouis = import-csv -Delimiter "`t" -path c:\Scripts\inventory\oui.txt

# collect DC list (the DHCP Servers for most organizations)
$DomainName = (Get-ADDomain).DNSRoot
$AllDCs = Get-ADDomainController -Filter * -Server $DomainName | Select-Object Hostname

foreach ($DC in $AllDCs) {
    if (test-connection $DC.hostname -quiet) {
        # collect the events
        write-host "Collecting DHCP Scopes and Leases from server" $dc.hostname
        $scopes = Get-DhcpServerv4Scope -computername $DC.hostname
        $rawleases += $scopes | foreach { get-dhcpserverv4lease -computername $DC.hostname $_.ScopeId -allleases }
        }
    }

# filter out duplicate leases due to DHCP redundancy
$leases = $rawleases | Sort-Object -property IPAddress -unique

write-host "Collecting Sites and Subnets data"

# collect the sites to subnets table.  Remove subnet masks
$RawSubnets = Get-ADReplicationSubnet -filter * -Properties * | Select Name, Site, Location, Description

ForEach($sub in $RawSubnets) {
    $SiteName = ""
    If ($Sub.Site -ne $null) {$SiteName = ($sub.Site.Split(',')[0]).split("=")[1]}
    $tempobj = [pscustomobject]@{
        subnet = $sub.name.split("/")[0]
        site = $SiteName
    }
    $subnets += $tempobj
}

write-host "Processing Lease Information"

foreach ($lease in $leases) {

    # convert everything to strings, create a temp var to work on
    $tempobj = [pscustomobject]@{
        IPAddress = $lease.ipaddress.ipaddresstostring
        subnet = $lease.ScopeID.IpAddressToString
        MAC = $lease.clientid
        HostName = $lease.HostName
        LeaseExpiryTime = "" 
        AddressState = $lease.AddressState
        Site =  "NOT DEFINED"
        OUIMfg = "NOT DEFINED"
        }

    # this accounts for Static Reservations, which have no expiry times
    if ($lease.LeaseExpiryTime) { 
        $tempobj.LeaseExpiryTime = $lease.leaseexpirytime.tostring("MMM-dd-yyyy") 
    }

    # assign a site to each lease
    foreach ($sub in $subnets) {
        if ($tempobj.subnet -eq $sub.subnet) {$dhcp
            $tempobj.site = $sub.site
            }
        }

    # assign a "Manufacturer" based on OUI to each lease
    $OUI = ($tempobj.MAC -replace "-").ToUpper().substring(0,6)
    $OUIRecord = ($OUIS -match $OUI)
    $tempobj.OUIMfg = $OUIRecord.VendorString


    $dhcpinven += $tempobj

    #progress indicator
    if (($dhcpinven.length % 70) -eq 0 ) {write-host "."} else { write-host -NoNewLine "." }

    }

# to populate first-time inventory at this point (this is already done):
# $dhcpinven | select hostname,mac | sort-object -property Hostname | Export-Csv -Path ./inventory.csv

# dhcpinven is now the full list of on-network hosts, as seen by dhcp 
# static reservations are of course in the list, and statically addressed hosts of course are not
# compare this list to the current "real" inventory to alert on new hosts

# first, open the current inventory file:
$inventory = import-csv -Path c:\Scripts\inventory\inventory.csv

# Now loop through dhcp inventory to find exceptions - look for new MAC addresses:
foreach ( $h in $dhcpinven ) {
    if ($inventory.MAC -notcontains $h.MAC) {
        $inventoryexception += $h
        }
    }
$inventoryupdate = $inventoryexception | select Hostname, MAC

# Output to file
# Review inventory-exception.csv, as it has all the dhcp data
# for updates, pick and choose out of inventory-update.csv, and add approved hosts into inventory.csv

$inventoryexception | export-csv -path "c:\Scripts\inventory\inventory-exception.csv"
$inventoryupdate | export-csv -path "c:\Scripts\inventory\inventory-update.csv"


# add alerts - email perhaps?
# Send-MailMessage -from "InventoryService@kindredcu.com" -to "rvandenbrink@coherentsecurity.com" `
#        -SmtpServer "serverip or fqdn here" `
#        -subject "DHCP Exception Report" -Attachments "./dhcp-exceptions.csv" 
$email =@{
From = "username@domain.com"
To = "username@domain.com"
#To = $to.Split(';')
Subject = "DHCP Exception Report"
Body = "This is the DHCP Exception Report.  Merge in desired records from the inventory-update.csv file to update the inventory (as known by DHCP)"
SMTPServer = "server.ip.or.fqdn"
Attachments = "c:\scripts\inventory\inventory-update.csv"
}

if ($inventory.update.count -gt 0) {send-mailmessage @email}
$inventoryexception | out-gridview
# pause so that gridview does not close with powershell
pause
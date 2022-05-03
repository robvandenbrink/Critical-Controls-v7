$pcs = get-adcomputer -filter * -property Name,OperatingSystem,Operatingsystemversion,LastLogonDate,IPV4Address
$patchinfo = @()
$i=0
$count = $pcs.count
foreach ($pc in $pcs) {
    $i+=1
    # keep total progress count 
    write-host "Host" $i "of" $count "is being checked"
    if (Test-Connection -ComputerName $pc.DNSHostName -count 2 -Quiet) {
        # echo the host being assessed (only live hosts hit this print)
        write-host $pc.dnshostname "is up, and is being assessed"
        $tempval = new-object psobject
        $hfs = get-hotfix -computername $pc.dnshostname | sort -descending InstalledOn
        foreach ($hf in $hfs) {
                $kbnum = $hf.hotfixid
                $WebResponse = Invoke-WebRequest "https://www.catalog.update.microsoft.com/Search.aspx?q=$kbnum"
                $returntable = $WebResponse.ParsedHtml.body.getElementsByTagName("table") | Where {$_.className -match "resultsBorder"}
                # write-host $returntable.outertext # uncomment if debugging
                if ($returntable.outertext -like "*Cumulative*")  {
                     $lasthf = $hf
                     break
                     }
                }
        $tempval | add-member -membertype noteproperty -name Name -value $pc.dnshostname
        $tempval | add-member -membertype noteproperty -name PatchDate -value $lasthf.installedon
        $tempval | add-member -membertype noteproperty -name OperatingSystem -value $pc.OperatingSystem
        $tempval | add-member -membertype noteproperty -name OperatingSystemVersion -value $pc.OperatingSystemVersion
        $tempval | add-member -membertype noteproperty -name IpAddress -value $pc.IPV4Address
        $tempval | add-member -membertype noteproperty -name LastLogonDate -value $pc.LastLogonDate
        $patchinfo += $tempval
        }
    }
$patchinfo | export-csv -path ./patchdate.csv





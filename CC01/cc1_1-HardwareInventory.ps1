$pcs = get-adcomputer -filter * -property Name,dnshostname,OperatingSystem,Operatingsystemversion,LastLogonDate,IPV4Address
$inventory = @()
$i=0
foreach ($pc in $pcs) {
    $i+=1
    write-host $i
    $Computer = $pc.DNSHostName
    if (Test-Connection -ComputerName $Computer -count 2 -Quiet) {
    $computerhw  = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer | select Name,Manufacturer, Model, SystemSKUNumber, TotalPhysicalMemory   
    $computerBIOS = gwmi win32_bios -ComputerName $Computer 
    $Serial_AssetTag = gwmi -ComputerName $Computer Win32_SystemEnclosure | Select-Object SerialNumber, SMBiosAssetTag
    $computerCPU = gwmi win32_processor -ComputerName $Computer | select DeviceID,Name
    $computertotalRAM = (gwmi Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
    $computerDisks = gwmi -Class Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $Computer | select DeviceID,VolumeName,Size,FreeSpace
    $computerOS = gwmi Win32_OperatingSystem -ComputerName $Computer | select Version, Caption
    $computerGPU = Get-WmiObject -ComputerName $Computer Win32_VideoController | Select description,driverversion,CurrentHorizontalResolution,CurrentVerticalResolution

    $tempval = $computerhw
    $tempval | add-member -membertype noteproperty -name SerialNum -value $computerBIOS.SerialNumber
    $tempval | add-member -membertype noteproperty -name AssetTag -value $Serial_AssetTag.SMBiosAssetTag
    $tempval | add-member -membertype noteproperty -name BIOSVersion -value $computerBIOS.Name
    $tempval | add-member -membertype noteproperty -name CPUID -value $computerCPU[0].Name
    $tempval | add-member -membertype noteproperty -name MEM -value $computertotalRAM
    $tempval | add-member -membertype noteproperty -name Disk0 -value $computerDisks[0].DeviceID
    $tempval | add-member -membertype noteproperty -name Disk0Size -value $computerDisks[0].Size
    $tempval | add-member -membertype noteproperty -name Disk0FreeSpace -value $computerDisks[0].FreeSpace
    $tempval | add-member -membertype noteproperty -name OSVersion -value $computerOS.Version
    $tempval | add-member -membertype noteproperty -name OSCaption -value $computerOS.Caption
    $tempval | add-member -membertype noteproperty -name GPUDesc -value $computerGPU.Description
    $tempval | add-member -membertype noteproperty -name GPUDriver -value $computerGPU.DriverVersion
    $tempval | add-member -membertype noteproperty -name GPUHorRes -value $computerGPU.CurrentHorizontalResolution
    $tempval | add-member -membertype noteproperty -name GPUVertRes -value $computerGPU.CurrentVerticalResolution
    # $tempval

    $inventory += $tempval
}
}
$inventory | export-csv -path ./domainhwinventory.csv

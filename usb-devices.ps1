# gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}

# $Devices = gwmi win32_diskdrive | ?{$_.interfacetype -eq "USB"} | %{gwmi -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=`"$($_.DeviceID.replace('\','\\'))`"} WHERE AssocClass = Win32_DiskDriveToDiskPartition"} |  %{gwmi -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=`"$($_.DeviceID)`"} WHERE AssocClass = Win32_LogicalDiskToPartition"} | %{$_.deviceid}

$Devices = @(Get-WmiObject -Query "Select * From Win32_LogicalDisk" | ? { $_.driveType -eq 2 })
ForEach ($Device in $Devices){
    gwmi win32_volume | Where-Object {$_.DriveLetter -eq ($Device.DeviceID)} | Select-Object DriveLetter

}

#$first_device = gwmi win32_volume | Where-Object {$_.DriveLetter -eq ($Device.DeviceID)} | Select-Object DriveLetter
#write-host $first_device | Select-Object DriveLetter
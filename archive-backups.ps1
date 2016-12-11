#requirements
# 1.- Start Windows PowerShell with the "Run as Administrator" option. Only members of the Administrators group on the computer can change the execution policy.
# 2.- Enable running unsigned scripts by entering:
#   set-executionpolicy remotesigned

# set-executionpolicy unrestricted
# set-executionpolicy restricted
# set-executionpolicy remotesigned

#write-Host "moving files from C:\COPIABD to C:\COPIABD_ARCHIVED"
write-Host "archiving files"

$path = "C:\COPIABD"
$archpath = "C:\COPIABD_ARCHIVED"
$days = "30"

write-progress -activity "Archiving Data" -status "Progress:"

Get-Childitem -Path $path |
Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-$days)} |
ForEach {
  $date = Get-Date
  $filename = $_.fullname
  try {
   Move-Item $_.FullName -destination $archpath -force -ErrorAction:SilentlyContinue
   "$date - Moved $filename to $archpath successfully" | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
  catch {
   "$date - Error moving $filename:  $_ " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
}


write-Host "operation complete, check log at c:\COPIABD_ARCHIVED\archiving-copiadb.log"

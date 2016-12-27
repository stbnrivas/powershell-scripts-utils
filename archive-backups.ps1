#requirements
# 1.- Start Windows PowerShell with the "Run as Administrator" option. Only members of the Administrators group on the computer can change the execution policy.
# 2.- Enable running unsigned scripts by entering:
#   set-executionpolicy remotesigned

# set-executionpolicy unrestricted
# set-executionpolicy restricted
# set-executionpolicy remotesigned

#write-Host "moving files from C:\COPIABD to C:\COPIABD_ARCHIVED"
write-Host "archiving files"
$current_date = get-date
$path = "C:\COPIABD"
$archive_path = "C:\COPIABD_ARCHIVED"
$days = "30"
"=== " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
"$current_date ARCHIVING DATA OLDER THAN A MONTH " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
write-progress -activity "Archiving Data" -status "Progress:"

Get-Childitem -Path $path |
Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-$days)} |
ForEach {
  $date = Get-Date
  $filename = $_.Fullname
  try {
   Move-Item $_.Fullname -destination $archive_path -force -ErrorAction:SilentlyContinue
   "$date - Moved $filename to $archive_path successfully" | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
  catch {
   "$date - Error moving $filename:  $_ " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
}
write-Host "operation archiving, check log at c:\COPIABD_ARCHIVED\archiving-copiadb.log"


#######################################

write-Host "deleting files"
"$current_date DELETION DATA OLDER THAN A YEAR " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log

$a_year_ago = "365"
Get-Childitem -Path $archive_path |
Where-Object {$_.LastWriteTime -lt (get-date).AddDays(-$a_year_ago)} |
ForEach {
  $date = Get-Date
  $filename = $_.Fullname
  try {
   Remove-Item -Path $filename -Exclude *.log
   "$date - delete $filename" | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
  catch {
   "$date - Error deleting $filename " | add-content c:\COPIABD_ARCHIVED\archiving-copiadb.log
  }
}
write-Host "operation deletion, check log at c:\COPIABD_ARCHIVED\archiving-copiadb.log"

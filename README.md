# powershell-scripts-utils
some scripts with salt pepper and other spices.

## requeriments

some scripts requires Windows Management Framework 5.1, specifically that use a GUI needs when is used in windows 7

https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7-KB3191566-x86.zip
https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win7AndW2K8R2-KB3191566-x64.zip



## how to create an clickable icon that open the scripts:

1. new direct access
2. select location of script
3. give a name
4. accept dialog
5. properties over direct access
6. on destination embrace the path with "path"
7. prefix with %SystemRoot%\syswo64\WindowsPowerShell\v1.0\powershell.exe "path"
  for instance %SystemRoot%\syswo64\WindowsPowerShell\v1.0\powershell.exe "C:\Users\user\scripts.ps1"

#Get-ExecutionPolicy -List

#Set-Executionpolicy remotesigned
#Set-ExecutionPolicy Unrestricted
#Set-ExecutionPolicy Restricted

#$Host.Version
#(Get-Host).Version


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$user = $env:UserName


$text = ""
$pathDefault = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\SCE")

$pathAndrea = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\Correo"),     
                ("Dropbox (Asinex)"),
                ("C:\Users\andrea2\AppData\Local\Microsoft\Outlook\Outlook.pst"),
                ("C:\Copia ordenador viejo"),
                ("C:\Aplicaciones BS Online")

$pathEsperanza = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\Correo")


$pathMonique = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\Correo")


$pathPilar = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\Correo")


$pathSamuel = ("Downloads"),
                ("Desktop"),
                ("Favorites"),
                ("Documents"),
                ("Links"),
                ("C:\Correo")

$pathFran = ("Downloads"),
            ("Desktop"),
            ("Favorites"),
            ("Documents"),
            ("Links"),
            ("C:\AEAT"),
            ("C:\Correo"),
            ("C:\ProgramData\ArticSoluciones"),
            ("C:\Users\fran\AppData\Roaming\xArticSoluciones")            



$textOfFoldersIntoBackups = ""
$arrayPaths = ""
#$user = "andrea2"

If ($user -eq "andrea2"){
    $arrayPaths = $pathAndrea
}  ElseIf ($user -eq "esperanza"){
    $arrayPaths = $pathEsperanza
}  ElseIf ($user -eq "monique"){    
    $arrayPaths = $pathMoniqu
}  ElseIf ($user -eq "pilar"){    
    $arrayPaths = $pathPilar
}  ElseIf ($user -eq "samuel"){    
    $arrayPaths = $pathSamuel
}  ElseIf ($user -eq "fran"){    
    $arrayPaths = $pathFran
}  ElseIf ($user -eq "stbn"){    
    $arrayPaths = $pathDefault
} Else {
    [System.Windows.MessageBox]::Show('Backup can not find user. Please add to script new user configuration')
    exit
}


foreach ($path in $arrayPaths){
    #$textOfFoldersIntoBackups = "$textOfFoldersIntoBackups &#x0a; $path "
    if (split-path $path -IsAbsolute){
        $textOfFoldersIntoBackups = "$textOfFoldersIntoBackups &#x0a; $path "                     
        #Write-Host "$path is absolute"
    } else {                    
        $textOfFoldersIntoBackups = "$textOfFoldersIntoBackups &#x0a; C:\Users\$user\$path " 
        #Write-Host "C:\Users\$user\$path"
    }   
}






function main {
    Param()

$inputXML = @"
<Window x:Class="calibradoraWPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:calibradoraWPF"
        mc:Ignorable="d"
        Title="Backup Tool Automation username: $user" SizeToContent="WidthAndHeight" MinWidth="430">    
        <Grid>
            <Grid.ColumnDefinitions>                        
                    <ColumnDefinition Width="1*" />
                    <ColumnDefinition Width="1*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="50" />
                    <RowDefinition Height="50" />
            </Grid.RowDefinitions>
            <StackPanel Margin="5 5 5 5">
                <TextBlock FontWeight="Bold" TextAlignment="Left" Text="ORIGIN:"></TextBlock>
                <TextBlock TextAlignment="Left" Text="$textOfFoldersIntoBackups"></TextBlock>
            </StackPanel>
            <StackPanel Grid.Column="1">
                <TextBlock FontWeight="Bold" TextAlignment="Left" Text="DESTINATION:"></TextBlock>
                <ComboBox Name="comboOutput"  Margin="1 20 20 1" SelectedIndex="0">
                    <ComboBoxItem>R:\</ComboBoxItem> 
                    <ComboBoxItem>D:\BACKUPS\$user</ComboBoxItem>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Row="1" Margin="5 5 5 5">
                <Label FontWeight="Bold">Shutdown after backup</Label>
                <CheckBox Name="cbShutdown" IsChecked="True">Yes</CheckBox>                
            </StackPanel>
            <Button Grid.Row="2" Grid.ColumnSpan="2" Name="btOpenLogsFolder" Content="Open Logs Folder" Margin="3 5 3 3"></Button>
            <Button Grid.Row="3" Grid.ColumnSpan="2" Name="btBackup" Content="Start Backup" Margin="3 5 3 3"></Button>
        </Grid>

</Window>
"@ 


    [xml]$XAML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window' 
    
    #Read XAML 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
    try {
        $Form=[Windows.Markup.XamlReader]::Load( $reader )
    } catch {    
        Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."        
    }   
    
    #Create variables to control form elements as objects in PowerShell
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -Scope Global  
        # Write-Host "WPF$($_.Name)"
    } 


    $WPFbtOpenLogsFolder.Add_Click{
        ii "C:\BACKUPS-LOGS\"
    }
         
        
    $WPFbtBackup.Add_Click{        
        $prefix_path_user = "C:\Users\$user"
        $prefix_destination = $WPFcomboOutput.Text

        Write-Host "Backup: " $prefix_path_user "->" $prefix_destination        

        # PARAMS FOR ROBOCOPY:
        # /MIR               mirror mode equivalent /E /PURGE
        # /R:N               max number of retries
        # /W:N               wait N seconds to retry
        # /Z                 copies files in "restart mode", so partially copied files can be continued after an interruption.
        # /LOG:<LogFile>     log replace file
        # /LOG+:<LogFile>    log append contain
        # /XJD               excludes"junction points" for directories, symbolic links

        # robocopy "C:\Users\fran\$path" "D:\BACKUPS\$user\$path" /MIR /R:5 /W:5 /Z /LOG+:C:\BACKUPS\LOGS\$((Get-Date).ToString("yyyy-MM-dd")).log

        If ( (Test-Path $prefix_path_user) -and (Test-Path $prefix_destination) ){

            foreach ($path in $arrayPaths) {	            
                if (split-path $path -IsAbsolute){
                    $pathWithoutUnit = $path.Remove(0,3)                    
                    write-host "robocopy '$path' '$prefix_destination\$pathWithoutUnit' /MIR /R:5 /W:5 /Z /LOG+:'C:\BACKUPS-LOGS\$((Get-Date).ToString("yyyy-MM-dd")).log'"
                    robocopy "$path" "$prefix_destination\$pathWithoutUnit" /MIR /R:5 /W:5 /Z /LOG+:"C:\BACKUPS-LOGS\$((Get-Date).ToString("yyyy-MM-dd")).log"
                    
                } else {                    
                    write-host "robocopy 'C:\Users\$user\$path' '$prefix_destination\$path' /MIR /R:5 /W:5 /Z /LOG+:'C:\BACKUPS\LOGS\$((Get-Date).ToString("yyyy-MM-dd")).log'"
                    robocopy "C:\Users\$user\$path" "$prefix_destination\$path" /MIR /R:5 /W:5 /Z /LOG+:"C:\BACKUPS-LOGS\$((Get-Date).ToString("yyyy-MM-dd")).log"
                }
            }    

        } else {
            [System.Windows.MessageBox]::Show('Please, check origin and destination folders existence.')
        }
        
        If ($WPFcbShutdown.IsChecked) {            
            [System.Windows.MessageBox]::Show('backup successfully, i will shutdown the machine.')
            # shutdown /t 60 /s /c
        }        
    
    }

    $form.ShowDialog() | Out-Null
    
}

        


#Call function to open the form
main
#

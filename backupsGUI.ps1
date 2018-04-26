#Get-ExecutionPolicy -List

#Set-Executionpolicy remotesigned
#Set-ExecutionPolicy Unrestricted
#Set-ExecutionPolicy Restricted

#$Host.Version
#(Get-Host).Version


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$user = $env:UserName


$folders_of_default = @"
&#x0a;
 C:\Users\$user\Downloads&#x0a;
 C:\Users\$user\Desktop&#x0a;
 C:\Users\$user\Favorites&#x0a;
 C:\Users\$user\Documents&#x0a;
 C:\Users\$user\Links&#x0a;
"@
$path_of_default = ("C:\Users\$user\Downloads"),
                ("C:\Users\$user\Desktop"),
                ("C:\Users\$user\Favorites"),
                ("C:\Users\$user\Documents"),
                ("C:\Users\$user\Links")

$folders = ""
if ( $user -eq "stbn" ){ 
$textOfFolders = $folders_of_default
$paths = $path_of_default

}



function mainBackupGui {
    Param()

$inputXML = @"
<Window x:Class="calibradoraWPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:calibradoraWPF"
        mc:Ignorable="d"
        Title="Backup Tool Automation" Width="500" Height="300">    
        <Grid>
            <Grid.ColumnDefinitions>                        
                    <ColumnDefinition Width="1*" />
                    <ColumnDefinition Width="1*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="1*" />
                    <RowDefinition Height="50" />
            </Grid.RowDefinitions>
            <StackPanel Margin="5 5 5 5">
                <TextBlock FontWeight="Bold" TextAlignment="Left" Text="ORIGIN:"></TextBlock>
                <TextBlock TextAlignment="Left" Text="$textoffolders"></TextBlock>
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
            <Button Grid.Row="2" Grid.ColumnSpan="2" Name="btBackup" Content="Start Backup" Margin="3 5 3 3"></Button>
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

         
        
    $WPFbtBackup.Add_Click{        
        $root_origin = "C:\Users\$user"
        $root_destination = $WPFcomboOutput.Text

        Write-Host $root_origin
        Write-Host $root_destination

        # PARAMS FOR ROBOCOPY:
        # /MIR               mirror mode equivalent /E /PURGE
        # /R:N               max number of retries
        # /W:N               wait N seconds to retry
        # /Z                 copies files in "restart mode", so partially copied files can be continued after an interruption.
        # /LOG:<LogFile>     log replace file
        # /LOG+:<LogFile>    log append contain
        # /XJD               excludes"junction points" for directories, symbolic links

        If ( (Test-Path $root_origin) -and (Test-Path $root_destination) ){

            If ( $user -eq "user1") {
                        
                # robocopy "$user\Downloads" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\Downloads" /E /MIR
                #robocopy "$user\Desktop" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\Desktop" /E /MIR
                #robocopy "$user\fran\Favorites" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\Favorites" /E /MIR
                #robocopy "$user\fran\Documents" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\Documents" /E /MIR
                #robocopy "$user\Links" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\Links" /E /MIR

                # PROBLEMS WITH CIRCULAR LINKS fixed at window 10
                #robocopy "$user\AppData" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\AppData" /E /MIR /XD /zb
                #robocopy "$user\AppData" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\AppData" /MIR /XD

                #robocopy "C:\ProgramData\ArticSoluciones" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\ProgramData\ArticSoluciones" /E /MIR
                #robocopy "$user\AppData\Roaming\xArticSoluciones" "D:\Backups\$((Get-Date).ToString("yyyy-MM-dd"))\AppData\Roaming\xArticSoluciones" /E  /MIR /XJ
            }

            If ($user -eq "user2"){
            
            
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
mainBackupGui
#

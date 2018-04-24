#Get-ExecutionPolicy -List

#Set-Executionpolicy remotesigned
#Set-ExecutionPolicy Unrestricted
#Set-ExecutionPolicy Restricted

#$Host.Version
#(Get-Host).Version


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')





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
        Title="Backup Tool Automation" Height="420" Width="600">    
    <Grid>
        <StackPanel>
            <TextBlock Name="textUser"  Margin="1 1 1 1" Padding="1 1 1 1" Text="User"></TextBlock>
            <ComboBox Name="comboUser" Margin="1 1 1 1" SelectedIndex="0">
                <ComboBoxItem>Default</ComboBoxItem>
                <ComboBoxItem>Customize for User</ComboBoxItem>
            </ComboBox>
            <TextBlock Name="textUserSelected" TextAlignment="Center" Text="Default profiles selected" Margin="0 15"></TextBlock>

            <TextBlock Name="textInput"  Margin="1 10 1 1" Padding="1 1 1 1" Text="Origin Folder"></TextBlock>
            <ComboBox Name="comboInput" Margin="1 1 1 1" SelectedIndex="0">
                <ComboBoxItem>Full Copy</ComboBoxItem>
            </ComboBox>
            <Button Name="buttonCheckInputFolder" Content="Check Origin" Margin="1 5 1 2"/>
            <TextBlock Name="textCheckInputFolder" TextAlignment="Center" Text="Destination not checked" Margin="0 15"></TextBlock>

            
            <TextBlock Name="textOutput" Text="Destination Folder"  Margin="1 1 1 1"></TextBlock>
            <ComboBox Name="comboOutput"  Margin="1 1 1 1" SelectedIndex="0">
                <ComboBoxItem>R:\</ComboBoxItem> 
                <ComboBoxItem>D:\</ComboBoxItem>
            </ComboBox>
            <Button Name="buttonCheckOutputFolders" Content="Check Destination" Margin="1 5 1 2"/>
            <TextBlock Name="textCheckOutputFolder" TextAlignment="Center" Text="Destination not checked" Margin="0 15"></TextBlock>
            
            <Button Name="buttonStartBackup" Content="Start Backup" Margin="1 10 1 2" Padding="0 10 0 10"/>
            
        </StackPanel>
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
        #Write-Host "WPF$($_.Name)"
    } 



    #WPFcomboUser
    #WPFtextUserSelected

    #WPFbuttonCheckInputFolder    
    #WPFcomboInput
    #WPFbuttonCheckInputFolder
    #WPFtextCheckInputFolder

    #WPFtextOutput
    #WPFcomboOutput
    #WPFbuttonCheckOutputFolders
    #WPFtextCheckOutputFolder

    #WPFbuttonStartBackup
    


    $form.ShowDialog() | Out-Null
    
}

        


#Call function to open the form
mainBackupGui
#

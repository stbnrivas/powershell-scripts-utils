#Get-ExecutionPolicy -List

#Set-Executionpolicy remotesigned
#Set-ExecutionPolicy Unrestricted
#Set-ExecutionPolicy Restricted

#$Host.Version
#(Get-Host).Version


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')




#######################################################
# Convert 
#######################################################

function Convert-MDB( $mdbPath, $mdbName, $txtPath, $txtName) {
    #[cmdletbinding()]
    #Param($mdbPath,$mdbName,$txtPath,$txtName)
    #Param([string]$mdbName,[string]$mdbPath,[string]$txtName,[string]$txtPath)
    
#    $mdbPath = $args[0]
#    $mdbName = $args[1]
#    $txtPath = $args[2]
#    $txtName = $args[3]    
    

    #$mdbName = "EN1800404.MDB"
    #$mdbPath = "C:\Users\stbn\Desktop\garsan\"
    #$txtName = "EN1800404.TXT"
    #$txtPath = "C:\Users\stbn\Desktop\garsan\"

    Write-Host "converting" $mdbPath $mdbName "->" $txtPath $txtName

    if ( Test-Path "$mdbPath$mdbName" ){

        if (Test-Path "$txtPath$txtName"){
          Remove-Item "$txtPath$txtName"
        }

        $connection = New-Object -ComObject ADODB.Connection
        $recordset = New-Object -ComObject ADODB.Recordset
        $connection.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$mdbPath$mdbName")
        #  “The ‘Microsoft.Jet.OLEDB.4.0’ provider is not registered on the local machine.”
        $adOpenStatic = 3
        $adLockOptmistic = 3
                
        New-Item "$txtPath$txtName" -ItemType file        
        "FRUTAS GARSAN S.L." | Add-Content "$txtPath$txtName"
        "" | Add-Content "$txtPath$txtName"

        $recordset.Open("select DataOraInizioLavorazione from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $valueInit = $recordset.Fields.Item("DataOraInizioLavorazione").value 
        $recordset.Close()
        $recordset.Open("select DataOraFineLavorazione from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $valueEnd = $recordset.Fields.Item("DataOraFineLavorazione").value 
        $recordset.Close()
        "FRUTO SELECCIONADO desde $valueInit hasta $valueEnd " | Add-Content "$txtPath$txtName"
        $deliveryNoteDate = $valueEnd.ToShortDateString()
        "Fecha de albaran: $deliveryNoteDate" | Add-Content "$txtPath$txtName"
        "Partida: $($mdbName.Substring(0,$mdbName.Length-4))" | Add-Content "$txtPath$txtName"

        $recordset.Open("select NumImballiTotale from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value = $recordset.Fields.Item("NumImballiTotale").value 
        $recordset.Close()
        "N. Bins: $value" | Add-Content "$txtPath$txtName"

        $recordset.Open("select Commento1 from ScarCom",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value = $recordset.Fields.Item("Commento1").value 
        $recordset.Close()
        "Comentario: $value " | Add-Content "$txtPath$txtName"

        $recordset.Open("select PgmLavorazione from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value = $recordset.Fields.Item("PgmLavorazione").value 
        $recordset.Close()
        "Programa: $value" | Add-Content "$txtPath$txtName"

        "Codigo productor: " | Add-Content "$txtPath$txtName"

        $recordset.Open("select CodiceVarieta from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value = $recordset.Fields.Item("CodiceVarieta").value 
        $recordset.Close()
        "Variedad: $value" | Add-Content "$txtPath$txtName"

        $recordset.Open("select PesoTotaleInEntrata from Testata",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $fullWeight = $recordset.Fields.Item("PesoTotaleInEntrata").value 
        $recordset.Close()
        "Peso partida: $fullWeight" | Add-Content "$txtPath$txtName"

        "" | Add-Content "$txtPath$txtName"

        $recordset.Open("select NomeQualita1 from ScarCom",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value = $recordset.Fields.Item("NomeQualita1").value 
        $recordset.Close()
        "Calidad: $value" | Add-Content "$txtPath$txtName"

        "" | Add-Content "$txtPath$txtName"

        "N Clase Numero Peso Porcent." | Add-Content "$txtPath$txtName"
        "" | Add-Content "$txtPath$txtName"
        $recordset.Open("select Indice, Nome, NumFrutti, PesoFrutti from ContClasse",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        do {
            $i1 = $recordset.Fields.Item("Indice").value
            $i2 = $recordset.Fields.Item("Nome").value
            $i3 = $recordset.Fields.Item("NumFrutti").value    
            $i4 = [convert]::ToDecimal(($recordset.Fields.Item("PesoFrutti").value))
            $i4 = [math]::Round($i4 / 1000,1)
    
            $i5 = [convert]::ToDecimal($recordset.Fields.Item("PesoFrutti").value)        
            #$i5 = ($i4*97.0)/$fullWeight # QUE? ... LOS DE LA CALIBRADORA NO DIVIDEN ENTRE 100
            $i5 = ($i4*100.0)/$fullWeight 
            $i5 = [math]::Round($i5,1)

            "$i1 $i2 $i3 $i4 Kg $i5 " | Add-Content "$txtPath$txtName"
            $recordset.MoveNext()
        } until ($recordset.EOF -eq $true)
        $recordset.Close()


        "" | Add-Content "$txtPath$txtName"
        $recordset.Open("select NumFrutti, PesoFrutti from ContQualita",$connection,$adOpenStatic,$adLockOptmistic)
        $recordset.MoveFirst()
        $value1 = $recordset.Fields.Item("NumFrutti").value 
        $value2 = [convert]::ToDecimal($recordset.Fields.Item("PesoFrutti").value)
        $value2 = [math]::Round($value2 / 1000,1)
        $recordset.Close()
        "Total calidad: $value1 $value2 Kg 100.0" | Add-Content "$txtPath$txtName"
        "" | Add-Content "$txtPath$txtName"
        "Total trabajado: $value2 Kg 100.0" | Add-Content "$txtPath$txtName"
    
        $connection.Close()
    } else { write-host "mdb path failed" }
}



#######################################################
# GUI creation and invocation
#######################################################
function Invoke-GUI {
    Param()

    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$inputXML = @"
<Window x:Class="calibradoraWPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:calibradoraWPF"
        mc:Ignorable="d"
        Title="Conversor Documentos Calibradora" Height="300" Width="600">    
    <Grid>
        <StackPanel>
            <TextBlock Name="textInput"  Margin="1 1 1 1" Padding="1 1 1 1" Text="Carpeta Origen"></TextBlock>
            <ComboBox Name="comboInput" Margin="1 1 1 1" SelectedIndex="0">
                <ComboBoxItem>\\10e137\c$\SCE\PartiteCali1\</ComboBoxItem>
                <ComboBoxItem>C:\SCE\PartiteCali1\</ComboBoxItem>
                <ComboBoxItem>C:\SCE\inputs\</ComboBoxItem>
            </ComboBox>
            <TextBlock Name="textOutput" Text="Carpeta Destino"  Margin="1 1 1 1"></TextBlock>
            <ComboBox Name="comboOutput"  Margin="1 1 1 1" SelectedIndex="0">
                <ComboBoxItem>C:\UNITEC\</ComboBoxItem> 
                <ComboBoxItem>N:\CALIBRADORA\UNITEC\ORIGEN\</ComboBoxItem>            
                <ComboBoxItem>C:\SCE\PartiteCali1\outputs\</ComboBoxItem>
                <ComboBoxItem>C:\Documents and Settings\Administrador\Escritorio\clasificaciones 2018\</ComboBoxItem>
            </ComboBox>
            <DatePicker Name="datePicker" Margin="1 1 1 1"></DatePicker>
            <TextBlock Name="textCountsSelected" TextAlignment="Center" Text="0 elementos seleccionados" Margin="0 15"></TextBlock>
            <Button Name="buttonCheckFolders" Content="Comprobar y Crear Directorios" Margin="1 5 1 2"/>
            <ComboBox Name="comboFormat"  Margin="1 1 1 1" SelectedIndex="0" IsEditable="False" IsReadOnly="True">
                <ComboBoxItem>Formato Salida: Informe Modo Texto</ComboBoxItem>                
            </ComboBox>
            <Button Name="buttonConvert" Content="Convertir" Margin="1 5 1 2"/>
            
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

    
    #WPFtextInput
    #WPFcomboInput
    #WPFtextOutput
    #WPFcomboOutput

    $WPFdatePicker.Add_SelectedDateChanged{        
        $Date = Get-Date $WPFdatePicker.Text
        $input_path = $WPFcomboInput.text
        $filesSelected = ( Get-ChildItem $input_path -Recurse | where { $_.LastWriteTime -gt $Date -and $_.LastWriteTime -lt $Date.AddDays(1) } | Measure-Object ).Count 
        $WPFtextCountsSelected.Text = $filesSelected.ToString() + " elementos seleccionados"        
    }

    #WPFtextCountsSelected
    $WPFbuttonCheckFolders.Add_Click{
        $input = "C:\SCE\PartiteCali1\"        
        $output = "$input\outputs"
        if(!(Test-Path -Path $input )){
            New-Item -ItemType directory -Path $input
        }
        if(!(Test-Path -Path $output )){
            New-Item -ItemType directory -Path $output
        }
    }
        
    $WPFbuttonConvert.Add_Click{
        $input_path = $WPFcomboInput.text
        $output_path = $WPFcomboOutput.Text
        $Date = Get-Date $WPFdatePicker.Text
        #Write-Host "converting from $input_path to $output_path"
        Get-ChildItem $input_path -Recurse | where { $_.LastWriteTime -gt $Date -and $_.LastWriteTime -lt $Date.AddDays(1) } | ForEach-Object {   
            Convert-MDB -mdbPath  $WPFcomboInput.text -mdbName $_.Name -txtPath $WPFcomboOutput.Text -txtName "$($_.Basename).TXT"
        }
         
    }
        
    #Show form
    $form.ShowDialog() | Out-Null
    
}

        


#Call function to open the form
Invoke-GUI
#

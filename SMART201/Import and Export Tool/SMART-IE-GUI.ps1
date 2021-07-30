########################################################################################
# SMA Runbook Toolkit (SMART) Import/Export GUI
# Version 1.2
# Published on Building Clouds blog : http://aka.ms/buildingclouds
# by the Windows Server and System Center CAT team
# Please send feedback to brunosa@microsoft.com
########################################################################################
# Special thanks to Jim Britt for the versions 1.0 and 1.1 of this tool.
# Features are unchanged in v1.2, except for the new GUI to simplify use and setup.
# v1.1 : http://blogs.technet.com/b/privatecloud/archive/2014/03/10/automation-sma-runbook-toolkit-smart-for-runbook-import-export-updated.aspx
# v1.0 : http://blogs.technet.com/b/privatecloud/archive/2013/10/23/automation-service-management-automation-sma-runbook-toolkit-spotlight-smart-for-runbook-import-and-export.aspx
########################################################################################

    param (
    [String]$WSEndPoint = "https://localhost",
    [String]$WSPort = "9090",
    [Bool]$AllowDownloadSMARTImportExport = $true
    )

$ToolVersion = "1.2"

########################################################################################
#Functions
########################################################################################

Function ImportRunbooksWithSMART
{
    param
    (
        [string]$ImportDirectory,
        [string]$WebServiceEndpoint,
        [string]$WebServicePort,
        [string]$RunbookState,
        [boolean]$ImportAssets,
        [Boolean]$overwrite        
    )

    # Builk Import from Directory with XML Files
    #$ScriptDirectory = (Get-Location -PSProvider FileSystem).Path + "\SMART for Runbook Import and Export"
    $ScriptDirectory = (Get-Location -PSProvider FileSystem).Path
    $OriginalLocation = (Get-Location -PSProvider FileSystem).Path
    $Files = Get-ChildItem $ImportDirectory -File

    # We setup a while loop to ensure we process as Import/Publish/Edit/Publish (runs twice) to validate parent / child dependencies on Runbooks
    $i = 0
    $Count = 1
    while($i -le 1)
    {
        foreach ($File in $Files)
        {
            Write-Progress -Activity "Importing files..." -PercentComplete (100*$Count/(2*$Files.Count))
            # Only Process XML and PS1 in working directory
            if($File.extension -iin ".XML",".PS1")
            {
                # Get Runbook Name
                $RunbookToPublish = $File.BaseName
                # Modify parameters below according to your needs
                    $parms = @{'WebServiceEndpoint'=$WebServiceEndpoint;
                       'Port'=$WebServicePort;
		               'ImportDirectory'=$ImportDirectory;
                       'FileName'=$File.Name;
                       'overwrite'=$overWrite;
                       'RunbookState'=$RunbookState;
                       'ImportAssets'=$ImportAssets}
                    cd $ScriptDirectory
                    .\Import-SMARunbookfromXMLorPS1.ps1 @parms
                # Publish the Runbook into SMA that we've just imported.               
                $PublishedRunbook = Publish-SMARunbook -Name $RunbookToPublish -WebServiceEndpoint $WebServiceEndpoint -Port $WebServicePort
            }
            $Count++
        }
        # Looping once to ensure Parent and Child have dependencies built properly in SMA
        $i++
    }
    Write-Progress -Activity "Importing files..." -Completed
    cd $OriginalLocation
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Completed importing SMART Import and Export Runbooks into SMA"
}

Function ExportRunbooksWithSMART
{

    param (
        [PSObject[]]$ExportRunbooksArray
    )

            #$ScriptDirectory = (Get-Location -PSProvider FileSystem).Path + "\SMART for Runbook Import and Export"
            $ScriptDirectory = (Get-Location -PSProvider FileSystem).Path
            $OriginalLocation = (Get-Location -PSProvider FileSystem).Path

            $count = 1
            cd $ScriptDirectory
            foreach ($Runbook in $ExportRunbooksArray)
                {
                Write-Progress -Activity "Exporting files..." -PercentComplete (100*$Count/($ExportRunbooksArray.Count))
                $parms = @{'WebServiceEndpoint'=$Global:SMAWSEndPoint;
                       'Port'= $Global:SMAWSPort;
                       'RunbookName' = $Runbook.RunbookName;
		               'ExportDirectory'=$FORM.FindName('ExportFolderLocationTextBox').Text;
                        'EnableScriptOutput'=$true;
                        'ExportPS1'=$true;
                       'ExportAssets'=($FORM.FindName('ExportFolderAssetsCB').IsChecked)}
                 cd $ScriptDirectory
                 .\Export-SMARunbookToXML.ps1 @parms
                $Count++
                }      
            Write-Progress -Activity "Exporting files..." -Completed
            cd $OriginalLocation
}

function CacheFullRunbookList()
{

    $J = Start-Job -ScriptBlock {
            param ($WS,$Port)
            get-smarunbook -WebServiceEndpoint $WS -Port $Port | select RunbookName, RunbookID, Tags, Description
            } -ArgumentList $Global:SMAWSEndPoint, $Global:SMAWSPort
    $Loop = 1
    while ( ($J.state -ne "Completed") -And ($Loop -lt 40) )
            {
            Start-Sleep(3)
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job is still running..."
            $Loop = $Loop+1
            }
    If ($J.State -eq "Completed")
            {
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job completed. If there was an error it should be displayed below."
            $Global:FullRunbookList  = Receive-Job -Job $J -Keep
            }
}

function Popup2()
{

param (
    [String]$Message,
    [Boolean]$ClosedExternally,
    [Boolean]$ShowDialog,
    [int]$NbLines
)
        $MB = New-Object System.Windows.Window
        $MB.Width = 500
        $MB.Height = 55*$NbLines
        $MB.WindowStyle = "None"

        $MBLabel = New-Object System.Windows.Controls.Label
        $MBLabel.HorizontalAlignment = "Center"
        $MBLabel.FontFamily = "Verdana"
        $MBLabel.Content = $Message

        $MBStackPanel = New-Object System.Windows.Controls.StackPanel
        $MBStackPanel.VerticalAlignment = "Top"

        If ($ClosedExternally -eq $False){
            $MBButton = New-Object System.Windows.Controls.Button
            $MBButton.Content = "Close"
            $MBButton.HorizontalAlignment = "Center"
            $MBButton.Width="100"
            $MBButton.Add_Click({
            $this.parent.parent.Close()
            })
            $result=$MBStackPanel.Children.Add($MBLabel)
            $result=$MBStackPanel.Children.Add($MBButton)
            $MB.Content = $MBStackPanel
        }
        else
        {
        $MB.Content = $MBLabel
        }
        
        If ($ShowDialog){$MB.ShowDialog()} else {$MB.Show()}
        #write-host ("POPUP : " + $Message.Replace("`r`n"," "))
        Return $MB
        
}


function ListTags()
{


        $Global:ProgressCount = 0

        $FORM.FindName('ImportLabel').IsEnabled = $False
        $FORM.FindName('ImportFolderTextBoxName').IsEnabled = $False
        $FORM.FindName('ImportFolderBrowseButton').IsEnabled = $False
        $FORM.FindName('ImportFolderCB').IsEnabled = $False
        $FORM.FindName('ImportFolderAssetsCB').IsEnabled = $False
        $FORM.FindName('ImportFolderRunbookStateLabel').IsEnabled = $False
        $FORM.FindName('ImportFolderRunbookState').IsEnabled = $False
        $FORM.FindName('ImportFolderButton').IsEnabled = $False
        $FORM.FindName('ExportLabel').IsEnabled = $False
        $FORM.FindName('ExportTagsLabel').IsEnabled = $False
        $FORM.FindName('ExportTagsComboBox').IsEnabled = $False
        $FORM.FindName('ExportFolderAssetsCB').IsEnabled = $False
        $FORM.FindName('ExportFolderLocationLabel').IsEnabled = $False
        $FORM.FindName('ExportFolderLocationTextBox').IsEnabled = $False
        $FORM.FindName('ExportFolderButton').IsEnabled = $False
        $FORM.FindName('ExportFolderAllButton').IsEnabled = $False

        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Trying to connect to web service $Global:SMAWSEndPoint"
        Write-Progress -Activity "Connecting to Web Service on $Global:SMAWSEndPoint and retrieving Tags..."
        $eap = $ErrorActionPreference = "SilentlyContinue"
        $TempVar = (get-smavariable -WebServiceEndpoint $Global:SMAWSEndPoint -Port $Global:SMAWSPort | ? name -like "*")
        if (!$?) {
            $ErrorActionPreference =$eap
            $MB = popup2 -Message ("Runbook hierarchy cannot be displayed.`r`nConnection to web service " + $Global:SMAWSEndPoint + " could not be opened.`r`nPlease configure or check the server name on the next screen and try again.") -ClosedExternally $False -NbLines 2 -ShowDialog $True
            }
            else{  
            $ErrorActionPreference =$eap
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Connected to web service, starting job to retrieve Tags..."

            $J = Start-Job -ScriptBlock {
                param ($WS,$Port)
                get-smarunbook -WebServiceEndpoint $WS -Port $Port | select RunbookName, RunbookId, Tags | Group-Object Tags
                } -ArgumentList $Global:SMAWSEndPoint, $Global:SMAWSPort
            $Loop = 1
            while ( ($J.state -ne "Completed") -And ($Loop -lt 40) )
                {
                Start-Sleep(3)
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job is still running..."
                $Loop = $Loop+1
                }

            If ($J.State -eq "Completed")
                {
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Job completed. If there was an error it should be displayed below."
                $data  = Receive-Job -Job $J -Keep
                $Global:UnfilteredTagArray = @()
                foreach ($TagGroup in $data)
                    {
                    foreach ($Runbook in $TagGroup.Group)
                        {
                        $Global:UnfilteredTagArray += ($TagGroup.Name -split (","))
                        }
                    $Global:ProgressCount = $Global:ProgressCount + ($TagGroup.Count/$NumberOfRunbooks)*100
                    }
                $Global:UnfilteredTagArray = $Global:UnfilteredTagArray | Select-Object -Unique
                $Global:TagArray = @()
                foreach ($NewTag in $Global:UnfilteredTagArray)
                        {
                        If ($NewTag -eq "") {$NewTag = "[No Tags]"}
                        $tmpObject = select-object -inputobject "" TagChecked, TagName
                        $tmpObject.TagChecked = $false
                        $tmpObject.TagName = $NewTag
                        $Global:TagArray += $tmpObject
                        }
                write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Tags parsing finished..."
                Write-Progress -Activity "Connecting to web service $Global:SMAWSEndPoint and retrieving Tags..." -Complete
                $FORM.FindName('ExportTagsComboBox').ItemsSource = $Global:TagArray
                $FORM.FindName('ImportLabel').IsEnabled = $True
                $FORM.FindName('ImportFolderTextBoxName').IsEnabled = $True
                $FORM.FindName('ImportFolderBrowseButton').IsEnabled = $True
                $FORM.FindName('ImportFolderCB').IsEnabled = $True
                $FORM.FindName('ImportFolderAssetsCB').IsEnabled = $True
                $FORM.FindName('ImportFolderRunbookStateLabel').IsEnabled = $True
                $FORM.FindName('ImportFolderRunbookState').IsEnabled = $True
                $FORM.FindName('ImportFolderButton').IsEnabled = $True
                $FORM.FindName('ExportLabel').IsEnabled = $True
                $FORM.FindName('ExportTagsLabel').IsEnabled = $True
                $FORM.FindName('ExportTagsComboBox').IsEnabled = $True
                $FORM.FindName('ExportFolderAssetsCB').IsEnabled = $True
                $FORM.FindName('ExportFolderLocationLabel').IsEnabled = $True
                $FORM.FindName('ExportFolderLocationTextBox').IsEnabled = $True
                $FORM.FindName('ExportFolderButton').IsEnabled = $True
                $FORM.FindName('ExportFolderAllButton').IsEnabled = $True
                }
            }
            
}



function LoadConfig()
 {

  param (
    [String]$ConfigFileLocation

  )

        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Loading Configuration from file $ConfigFileLocation..."
        $XmlReader= New-Object System.Xml.XmlTextReader($ConfigFileLocation)
        While ($XmlReader.Read()){
                              If ($XmlReader.NodeType -eq [System.Xml.XmlNodeType]::Element){
                                   switch ($XmlReader.Name){
                                        "WSEndPoint" {$Global:SMAWSEndpoint=$XmlReader.ReadString()}
                                        "WSPort"{$Global:SMAWSPort=$XmlReader.ReadString()}
                                    }
                               }
        }
        $XmlReader.Close()
 }

 function SaveConfig()
 {

 param (
    [String]$ConfigFileLocation,
    [String]$WSEndPoint,
    [String]$WSPort
 )

 write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] Saving Configuration to file $ConfigFileLocation..."

 $XMLData = 
@”
<DefaultConfiguration>
<WSEndPoint>$WSEndPoint</WSEndPoint>
<WSPort>$WSPort</WSPort>
</DefaultConfiguration>
“@

$XMLData | Out-File $ConfigFileLocation -Force

 }


########################################################################################
#Make sure we run elevated, or relaunch as admin
########################################################################################


$CurrentScriptDirectory = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\"))
Set-Location $CurrentScriptDirectory


    #Thanks to http://gallery.technet.microsoft.com/scriptcenter/63fd1c0d-da57-4fb4-9645-ea52fc4f1dfb
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $ScriptToLaunch = $PSCommandPath.Substring(0,$PSCommandPath.LastIndexOf("\")) + "\SMART-IE-GUI.ps1"
                $arg = "-file `"$($ScriptToLaunch)`"" 
                write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] WARNING : This script should run with administrative rights - Relaunching the script in elevated mode in 3 seconds..."
                start-sleep 3
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'

            } 
            catch 
            { 
                write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] Error : Failed to restart script with administrative rights - please make sure this script is launched elevated."  
                break               
            } 
            exit
        }
        else
        {
        write-host -ForegroundColor gray "["(date -format "HH:mm:ss")"] We are running in elevated mode, we can proceed with launching the tool."
        }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

##########################################################################################
# GUI definition
# Events handling buttons and clicks are also added here
########################################################################################## 

# Form and GUI definition
[XML]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="SMART for Import and Export GUI" Height="450" Width="520">
        
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="30"/>  
                            <RowDefinition Height="30"/>                                 
                            <RowDefinition Height="30"/>                       
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="50"/>
                            <RowDefinition Height="30"/>      
                            <RowDefinition Height="30"/> 
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="50"/>
                            <RowDefinition Height="50"/>
                        </Grid.RowDefinitions>

                        <Label Content="SMA Endpoint :" VerticalAlignment="Center" Grid.Row="0"></Label>
                        <TextBox Text="localhost" Name="DBServer" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center"  Margin="-100,0,0,0" Width="200"></TextBox>
                        <TextBox Text="9090" Name="DBPort" Grid.Row="0" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="160,0,0,0" Width="35"></TextBox>
                        <Button IsDefault="true" Content="Connect" Name="UpdateWSEndpoint" HorizontalAlignment="Right" VerticalAlignment="Center" Width="100" Margin="0,0,10,0" Grid.Row="0"/>

                        <Label FontWeight="Bold" Name="ImportLabel" IsEnabled="False" Content="Import" HorizontalAlignment="Center" VerticalAlignment="Center" Width="200" Margin="135,0,0,0" Grid.Row="1"></Label>
                        <TextBox Name="ImportFolderTextBoxName" IsEnabled="False" Text="" HorizontalAlignment="Left" VerticalAlignment="Center" Width="300" Margin="7,0,0,0" Grid.Row="2"></TextBox>
                        <Button Content="Browse..." Name="ImportFolderBrowseButton" IsEnabled="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="100" Margin="330,0,0,0" Grid.Row="2"/>
                        <CheckBox Content="Overwrite?" IsEnabled="False" Name="ImportFolderCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="7,0,0,0" Grid.Row="3"></CheckBox>
                        <CheckBox Content="Import Assets?" IsEnabled="False" IsChecked="True" Name="ImportFolderAssetsCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="100,0,0,0" Grid.Row="3"></CheckBox>
                        <Label IsEnabled="False" Name="ImportFolderRunbookStateLabel" Content="Import as" HorizontalAlignment="Left" VerticalAlignment="Center" Width="200" Margin="250,0,0,0" Grid.Row="3"></Label>
                        <ComboBox Name="ImportFolderRunbookState" IsEditable="True" IsEnabled="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="100"  Margin="330,0,0,0" Grid.Row="3"></ComboBox>
                        <Button Content="Import Runbooks" Name="ImportFolderButton" IsEnabled="False" HorizontalAlignment="Center" VerticalAlignment="Center" Height="40" Width="130" Margin="0,0,0,0" Grid.Row="4"/>

                        <Label FontWeight="Bold" IsEnabled="False" Content="Export" Name="ExportLabel" HorizontalAlignment="Center" VerticalAlignment="Center" Width="200" Margin="135,0,0,0" Grid.Row="6"></Label>
                        <Label IsEnabled="False" Name="ExportTagsLabel" Content="Select tag(s)" HorizontalAlignment="Left" VerticalAlignment="Center" Width="200" Margin="7,0,0,0" Grid.Row="7"></Label>
                        <ComboBox Name="ExportTagsComboBox" IsEnabled="False" HorizontalAlignment="Left" VerticalAlignment="Center" Width="200" Margin="100,0,0,0" Grid.Row="7">
                                <ComboBox.ItemTemplate>
                                    <DataTemplate>
                                        <StackPanel Orientation="Horizontal">
                                            <CheckBox Margin="5" IsChecked="{Binding TagChecked}"/>
                                            <TextBlock Margin="5" Text="{Binding TagName}"/>
                                        </StackPanel>
                                    </DataTemplate>
                            </ComboBox.ItemTemplate>
                        </ComboBox>
                        <CheckBox Content="Export Assets?" IsEnabled="False" IsChecked="True" Name="ExportFolderAssetsCB" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="320,0,0,0" Grid.Row="7"></CheckBox>

                        <Label IsEnabled="False" Name="ExportFolderLocationLabel" Content="Output folder" HorizontalAlignment="Left" VerticalAlignment="Center" Width="200" Margin="7,0,0,0" Grid.Row="8"></Label>
                        <TextBox Name="ExportFolderLocationTextBox" IsEnabled="False" Text="" HorizontalAlignment="Left" VerticalAlignment="Center" Width="300" Margin="100,0,0,0" Grid.Row="8"></TextBox>

                        <Button Content="Export Runbooks" IsEnabled="False" Name="ExportFolderButton" HorizontalAlignment="Center" VerticalAlignment="Center" Height="40" Width="130" Margin="0,0,0,0" Grid.Row="9"/>
                        <Button Content="Export All Runbooks" IsEnabled="False" Name="ExportFolderAllButton" HorizontalAlignment="Center" VerticalAlignment="Center" Height="40" Width="130" Margin="0,0,0,0" Grid.Row="10"/>

                    </Grid>

</Window>

'@

$Reader = (New-Object System.XML.XMLNodeReader $XAML)
$FORM = [Windows.Markup.XAMLReader]::Load($Reader)
$DBServerTB = $FORM.FindName('DBServer')
$DBPortTB = $FORM.FindName('DBPort')

$FORM.FindName('UpdateWSEndpoint').Add_Click({
        $Global:SMAWSEndPoint = $FORM.FindName('DBServer').Text
        $Global:SMAWSPort = $FORM.FindName('DBPort').Text
        ListTags
})

$FORM.FindName('ImportFolderBrowseButton').Add_Click({
    
    $NewShell = New-Object -comObject Shell.Application   
    $PickedFolder = $NewShell.BrowseForFolder(0, "Pick a folder with the PS1 files you wish to import", 0, 0)  
    if ($PickedFolder -ne $null) {  
        $PS1Folder = $PickedFolder.self.Path
        $NumberOfRunbooksInPickedfolder = (Get-Item "$PS1Folder\*.ps1").Count
        $FORM.FindName('ImportFolderTextBoxName').Text = $PS1Folder
    }  
    else
    {
    write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] No Folder selected by user, returning to main window..."
    }
})

$FORM.FindName('ImportFolderButton').Add_Click({
    If ($FORM.FindName('ImportFolderTextBoxName').Text -ne "")
        {
        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Importing Runbooks from folder" $FORM.FindName('ImportFolderTextBoxName').Text ": Starting..."
        ImportRunbooksWithSMART -WebServiceEndpoint $Global:SMAWSEndPoint -WebServicePort $Global:SMAWSPort -ImportDirectory $FORM.FindName('ImportFolderTextBoxName').Text -RunbookState $FORM.FindName('ImportFolderRunbookState').Text -ImportAssets $FORM.FindName('ImportFolderAssetsCB').IsChecked -overwrite $FORM.FindName('ImportFolderCB').IsChecked
        ListTags
        $FORM.FindName('ExportTagsComboBox').ItemsSource = $Global:TagArray
        write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Importing Runbooks from folder" $FORM.FindName('ImportFolderTextBoxName').Text ": Done!"
        popup2 -Message ("Runbook Import completed from folder " + $FORM.FindName('ImportFolderTextBoxName').Text) -ClosedExternally $False -NbLines 2 -ShowDialog $True
        }
        else
        {
        write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] WARNING : No Folder was selected for import in the tool."
        popup2 -Message "No Folder was selected for import in the tool." -ClosedExternally $False -NbLines 2 -ShowDialog $True
        }
})

$FORM.FindName('ExportFolderButton').Add_Click({
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exporting Runbooks: Starting..."
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Computing list of runbooks to export (Step 1 of 2)"
            CacheFullRunbookList

            $ExportRunbooksArray = @()
            $NoTagsRunbooksRequested = $false
            $CheckedTags = ($FORM.FindName('ExportTagsComboBox').Items | ? TagChecked -eq $True).TagName
            If ($CheckedTags)
                {
                If ($CheckedTags -contains "[No Tags]") {$NoTagsRunbooksRequested = $true}
                [regex] $CheckedTagsRegEx = ‘(‘ + (($CheckedTags |foreach {[regex]::escape($_)}) –join “|”) + ‘)’
                foreach ($Runbook in ($Global:FullRunbookList | Where-Object {$_.Tags -match $CheckedTagsRegEx})) 
                    {
                    $ExportRunbooksArray += New-Object PSObject -Property @{ 
                        RunbookName = $Runbook.RunbookName
                        RunbookID = $Runbook.RunbookID
                        } 
                    }
                }
            #We also  we may have to take into account runbooks without any tags
            If ($NoTagsRunbooksRequested)
                {
                foreach ($Runbook in ($Global:FullRunbookList | Where-Object {$_.Tags -eq $null})) 
                    {
                    $ExportRunbooksArray += New-Object PSObject -Property @{ 
                        RunbookName = $Runbook.RunbookName
                        RunbookID = $Runbook.RunbookID
                        } 
                    }
                }
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Exporting" $ExportRunbooksArray.Count "Runbooks to folder" $FORM.FindName('ExportFolderLocationTextBox').Text "(Step 2 of 2)"
            ExportRunbooksWithSMART -ExportRunbooksArray $ExportRunbooksArray
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exporting Runbooks: Done!"
            popup2 -Message ("Runbook Export completed in folder " + $FORM.FindName('ExportFolderLocationTextBox').Text) -ClosedExternally $False -NbLines 2 -ShowDialog $True
})

$FORM.FindName('ExportFolderAllButton').Add_Click({
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exporting Runbooks: Starting..."
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"]  - Computing list of runbooks to export (Step 1 of 2)"
            CacheFullRunbookList

            $ExportRunbooksArray = @()
            foreach ($Runbook in $Global:FullRunbookList) 
                {
                $ExportRunbooksArray += New-Object PSObject -Property @{ 
                    RunbookName = $Runbook.RunbookName
                    RunbookID = $Runbook.RunbookID
                    } 
                }
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] - Exporting" $ExportRunbooksArray.Count "Runbooks to folder" $FORM.FindName('ExportFolderLocationTextBox').Text "(Step 2 of 2)"
            ExportRunbooksWithSMART -ExportRunbooksArray $ExportRunbooksArray
            write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exporting Runbooks: Done!"
            popup2 -Message ("Runbook Export completed in folder " + $FORM.FindName('ExportFolderLocationTextBox').Text) -ClosedExternally $False -NbLines 2 -ShowDialog $True

})

$FORM.FindName('ImportFolderRunbookState').Items.Add("Published") | out-null
$FORM.FindName('ImportFolderRunbookState').Items.Add("Draft") | out-null
$FORM.FindName('ImportFolderRunbookState').Text = "Published"
$FORM.FindName('ExportFolderLocationTextBox').Text = (Get-Location -PSProvider FileSystem).Path + "\Export\"

##########################################################################################
# Actual start of the main process
##########################################################################################

cls
write-host -ForegroundColor green "["(date -format "HH:mm:ss")"] SMART Import/Export GUI $ToolVersion"
# We default to the web service endpoint and port values from the script, as well as initialize Visio variables
$Global:SMAWSEndPoint = $WSEndPoint
$Global:SMAWSPort = $WSPort
# Let's try to see if there is a config file and, if yes, update the global variables
$ConfigFileLocation = (Get-Location -PSProvider FileSystem).Path + "\Config-SMART-IE-GUI.xml"
If (Test-Path -Path $ConfigFileLocation){LoadConfig -ConfigFileLocation $ConfigFileLocation}
$DBServerTB.Text = $Global:SMAWSEndPoint
$DBPortTB.Text = $Global:SMAWSPort
#Fill out the GUI and display the form
ListTags
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Displaying GUI..."
$FORM.ShowDialog() | Out-Null
# Everything after  this line is executed when exiting the tool
SaveConfig  -ConfigFileLocation $ConfigFileLocation -WSEndPoint $Global:SMAWSEndPoint -WSPort $Global:SMAWSPort
write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] Exiting..."


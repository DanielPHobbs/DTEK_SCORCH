##########################################################################################
# SMA Runbook Toolkit - Documentation and Conversion Helper - Image Export Script
# Version : 2.0
# Windows Server and System Center Customer, Architecture and Technologies (CAT) team
# Please send feedback to brunosa@microsoft.com
# Parameters (all optional) :
#              -OrchestratorExtensionsDir : Path to the Orchestrator extensions
#               This is typically on the Orchestrator management server

##########################################################################################
    
    param (
    [String]$Global:OrchestratorExtensionsDir
    )

#$OrchestratorExtensionsDir = "\\ORCH01\c$\Program Files (x86)\Common Files\Microsoft System Center 2012\Orchestrator\Extensions"
If ($Global:OrchestratorExtensionsDir -eq "")
    {$Global:OrchestratorExtensionsDir = "\\MyOrchServer\c$\Program Files (x86)\Common Files\Microsoft System Center 2012\Orchestrator\Extensions"}

##########################################################################################
# Loading dependent functions
##########################################################################################

$code=@' 
[DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern IntPtr LoadLibrary(
    string lpFileName
);
[DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern bool FreeLibrary(
    IntPtr hModule
);

[DllImport("Kernel32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern uint GetLastError();

[DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern IntPtr LoadBitmap(
    IntPtr hInstance,
    String lpBitmapName
);

[DllImport("gdi32.dll", SetLastError=true, CharSet=CharSet.Auto)]
public static extern bool DeleteObject(
    IntPtr hObject
);
'@ 

Add-Type -memberDefinition $code -name API -namespace SMARTWin32Functions
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

##########################################################################################
# Image Export GUI and functions
########################################################################################## 

[XML]$XAMLImageExport = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        ResizeMode="NoResize"
        Title="Orchestrator Extensions Path" Height="140" Width="630">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="40"/>
            <RowDefinition Height="30"/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>

            <Label HorizontalAlignment="Left" Grid.Row="0" Width="580">  
                <TextBlock Text="Please confirm the path to Orchestrator extensions. The path will be checked when clicking on 'Export Images', and details will be logged in the main PowerShell window" TextWrapping= "Wrap"></TextBlock>
            </Label>
            <TextBox Text="\\MyOrchestratorServer\c$\Program Files (x86)\Common Files\Microsoft System Center 2012\Orchestrator\Extensions" Name="OrchestratorExtensionsDir" Grid.Row="1" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="5,0,0,0" Width="600"></TextBox>
            <Button Content="Export Images" Name="ExportImages" HorizontalAlignment="Left" VerticalAlignment="Center" Width="100" Margin="350,0,0,0" Grid.Row="2"/>
            <Button Content="Cancel" Name="CancelExport" HorizontalAlignment="Left" VerticalAlignment="Center" Width="100" Margin="500,0,0,0" Grid.Row="2"/>

    </Grid>
</Window>

'@
$ReaderImageExport = (New-Object System.XML.XMLNodeReader $XAMLImageExport)
$FORMImageExport = [Windows.Markup.XAMLReader]::Load($ReaderImageExport)

$FORMImageExport.FindName('ExportImages').Add_Click({
    If ((Test-Path $FORMImageExport.FindName('OrchestratorExtensionsDir').Text) -eq $True)   
        {
        $Global:OrchestratorExtensionsDir = $FORMImageExport.FindName('OrchestratorExtensionsDir').Text
        $FORMImageExport.Close()
        }
        else
        {Write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] The Extensions path could not be reached. Please double check it is valid and reachable, and try again. You can also hit 'Cancel Export' and exit the script."}
    })

$FORMImageExport.FindName('CancelExport').Add_Click({
    $FORMImageExport.Close()
    Write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] User chose to cancel specifying the extensions path, exiting the script..."
    exit
    })

##########################################################################################
# Validate Prerequisites
##########################################################################################

Write-host -ForegroundColor green "["(date -format "HH:mm:ss")"] SMART Documentation and Conversion Helper 2.0 - Export Images Script"

If ($env:Processor_Architecture -eq "x86")
        {Write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] This script is running in PowerShell 32-bit, we can continue execution"}
  else
        {Write-host -ForegroundColor red "["(date -format "HH:mm:ss")"] This script should be running in PowerShell 32-bit. It was started with Powershell 64-bit, exiting in 10 seconds...";start-sleep 10;exit;}

If ((Test-Path $Global:OrchestratorExtensionsDir) -eq $false)   
        {
        $FORMImageExport.FindName('OrchestratorExtensionsDir').Text = $Global:OrchestratorExtensionsDir
        $FORMImageExport.ShowDialog() | Out-Null
        }

  
##########################################################################################
# Actual start of the main process
##########################################################################################
   

foreach ($ExtensionFile in (Get-Item "$OrchestratorExtensionsDir\*.xml"))
        {
        Write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] Working with file $ExtensionFile"
        $Reader = New-Object System.XML.XmlTextReader($ExtensionFile)
        $result = $Reader.ReadToFollowing("ResourceFile")
        $TempResource = $Reader.ReadString()
        If ($TempResource.Substring($TempResource.Length-4, 4) -eq ".dll")
                {
                Write-host -ForegroundColor yellow "["(date -format "HH:mm:ss")"] There are resources for this file in DLL $TempResource"
                $result = $Reader.ReadToFollowing("ObjectType")
                #we do this a second time to skip the category
                $result = $Reader.ReadToFollowing("ObjectType")
                $TmpObjectType = $Reader.ReadString()
                while ($TmpObjectType)
                    {
                    $result = $Reader.ReadToFollowing("Bitmap32")
                    $TmpImage32 = $Reader.ReadString()
                    Write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting $TmpImage32 Starting..."
                    $hModule = [SMARTWin32Functions.API]::LoadLibrary("$OrchestratorExtensionsDir\$TempResource") 
                    $hBitmap = [SMARTWin32Functions.API]::LoadBitmap($hModule, $TmpImage32)
                    #[SMARTWin32Functions.API]::GetLastError()
                    $bmp = [System.Drawing.Image]::FromHbitmap($hBitmap)
                    $result = [SMARTWin32Functions.API]::DeleteObject($hBitmap)
                    $result = [SMARTWin32Functions.API]::FreeLibrary($hModule)
                    $bmp.Save(((Get-Location -PSProvider FileSystem).ProviderPath) + "\$TmpObjectType.jpg", "Jpeg")
                    Write-host -ForegroundColor white "["(date -format "HH:mm:ss")"] -- Exporting $TmpImage32 Done!"
                    $result = $Reader.ReadToFollowing("ObjectType")
                    $TmpObjectType = $Reader.ReadString()
                    }
                }
        $Reader.Close()
        }
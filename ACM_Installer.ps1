

Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'

[Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)

#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="800" Title="(ACM) Azure Certificate Manager / Key Manager Plus Edition - Installer" Height="400" Name="Window_Main" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
<Grid>
 <Border BorderBrush="Black" BorderThickness="1" HorizontalAlignment="Left" Height="56" VerticalAlignment="Top" Width="787" Margin="6,8,0,0" Background="#6c60a5">
<TextBlock TextWrapping="Wrap" Text="Fill in the Storage account Name and Key. Fill in the Email account information (uses only 365 accounts with App Password). Installer will create local file and the storage account file structure. All Files containing sensitive information will be encrypted locally using a randomized key created during the install process." Margin="8,3,10,1" Foreground="#ffffff" FontSize="12"/>
</Border>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="96" Width="769" TextWrapping="Wrap" Margin="7,258,0,0" BorderBrush="#000000" Background="#eeecf3" Text="{Binding Output}" Name="lo64cwmsdcbuf"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Output" Margin="8,241,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Storage Account Name" Margin="10,67,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Storage Account Key" Margin="11,120,0,0"/>


<Button Content="Install" HorizontalAlignment="Left" VerticalAlignment="Top" Width="75" Margin="641,157,0,0" Foreground="#194d4a" BorderBrush="#194d4a" Height="30" FontSize="15" FontWeight="DemiBold" Background="#ffffff" Name="Install_Button"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="28" Width="321" TextWrapping="Wrap" Margin="10,84,0,0" Text="{Binding Storage_ACCT_Name}" Padding="4,8,0,0" Name="lo64cwmsu38x4"/>



<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Once the fields are filled in, click Install" Margin="583,109,0,0"/>
<TextBlock HorizontalAlignment="Left" VerticalAlignment="Top" TextWrapping="Wrap" Text="Storage Account File Share Name" Margin="11,178,0,0"/>
<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="30" Width="324" TextWrapping="Wrap" Margin="10.84375,196.984375,0,0" Text="{Binding Storage_ACCT_FS}" Padding="4,8,0,0" Name="lo64cwmt9j7ie"/>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="156" Width="2" Margin="563.844,76.9844,0,0"/>
<Rectangle HorizontalAlignment="Left" VerticalAlignment="Top" Fill="#FFF4F4F5" Stroke="Black" Height="2" Width="227" Margin="564.844,230.984,0,0"/>


<TextBox HorizontalAlignment="Left" VerticalAlignment="Top" Height="31" Width="549" TextWrapping="Wrap" Margin="11,138,0,0" Name="lo64cwmttvvb8" Text="{Binding Storage_ACCT_Key}" Padding="4,8,0,0"/>
</Grid></Window>
"@

#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#


#region Install
#Install button function. Starts the Softinstalls and configurations
function install_button {
    Async {
        $outputText = "Installation Started..."
        $State.Output = $outputText
        $installTimeStart = Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"
        softInstallS1
    }
}

#Software Install Step1 - Create local directories and Encryption Key
function softInstallS1 {
        $outputText = $outputText + "`nCreating local directories..."
        $State.Output = $outputText
            New-Item -Path "C:\Program Files\" -Name "ACM" -ItemType "directory"
            New-Item -Path "C:\Program Files\ACM\" -Name "Logs" -ItemType "directory"
            New-Item -Path "C:\Program Files\ACM\" -Name "Config" -ItemType "directory"
            New-Item -Path "C:\Program Files\ACM\" -Name "Creds" -ItemType "directory"
            New-Item -Path "C:\Program Files\ACM\" -Name "Software" -ItemType "directory"
            
        $outputText = $outputText + "`nCreating Enc Key..."
        $State.Output = $outputText
            $EncryptionKeyBytes = New-Object Byte[] 32
            [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($EncryptionKeyBytes)
            $EncryptionKeyBytes | Out-File "C:\Program Files\ACM\Config\ACM_Enc.key"
        softInstallS2
}

#Software Install Step2 - Connect to storage account, setup folder structure and set the encrypted local JSON file
function softInstallS2 {
        $outputText = $outputText + "`nTrying to connect to storage..."
        $State.Output = $outputText
        $storName = $State.Storage_ACCT_Name
        $storeKey = $State.Storage_ACCT_Key
        $storeFS = $State.Storage_ACCT_FS
            $DriveLtr = (68..90 | %{$L=[char]$_; if ((gdr).Name -notContains $L) {$L}})[0]
            cmd.exe /C "cmdkey /add:`"$storName.file.core.windows.net`" /user:`"localhost\$storName`" /pass:`"$storeKey`""
        try {    
            New-PSDrive -Name $DriveLtr -PSProvider FileSystem -Root "\\$storName.file.core.windows.net\$storeFS"
            $outputText = $outputText + "`nConnected to Storage..."
            $State.Output = $outputText
            $outputText = $outputText + "`nCreating Directory in Storage..."
            $State.Output = $outputText
            $DriveGT = $DriveLtr + ':'
            cd $DriveGT
            New-Item -Name "Delete" -ItemType "directory"
            New-Item -Name "Employee" -ItemType "directory"
            New-Item -Name "Logs" -ItemType "directory"
            New-Item -Name "Search" -ItemType "directory"
            New-Item -Name "Update" -ItemType "directory"
            $outputText = $outputText + "`nDirectories Created..."
            $State.Output = $outputText
            $outputText = $outputText + "`nCreating local Enc file from storage information..."
            $State.Output = $outputText
            softInstallS3
        } catch {    
            $outputText = "ERROR - There was an error connecting to the storage account, please fix connection issue and then restart the install."
            $State.Output = $outputText
}
}

#Software Install Step3 - Write the Storage account information to the local server using the enc key
function softInstallS3 {
        $ACMKeyData = Get-Content "C:\Program Files\ACM\Config\ACM_Enc.key"
        $StorageName = $State.Storage_ACCT_Name | ConvertTo-SecureString -AsPlainText -Force
        $StorageName | ConvertFrom-SecureString -Key $ACMKeyData | Out-File "C:\Program Files\ACM\Creds\StorageName.txt"
        $outputText = $outputText + "`nStorage Name Encrypted file created and set..."
        $State.Output = $outputText
        
        $ACMKeyData = Get-Content "C:\Program Files\ACM\Config\ACM_Enc.key"
        $StorageKey = $State.Storage_ACCT_Key | ConvertTo-SecureString -AsPlainText -Force
        $StorageKey | ConvertFrom-SecureString -Key $ACMKeyData | Out-File "C:\Program Files\ACM\Creds\StorageKey.txt"
        $outputText = $outputText + "`nStorage Key Encrypted file created and set..."
        $State.Output = $outputText
        
        $ACMKeyData = Get-Content "C:\Program Files\ACM\Config\ACM_Enc.key"
        $StorageFSN = $State.Storage_ACCT_FS | ConvertTo-SecureString -AsPlainText -Force
        $StorageFSN | ConvertFrom-SecureString -Key $ACMKeyData | Out-File "C:\Program Files\ACM\Creds\StorageFSN.txt"
        $outputText = $outputText + "`nStorage File Share Name Encrypted file created and set..."
        $State.Output = $outputText
        softInstallS4
}

#Software Install Step4 - Download Software to server from github account
function softInstallS4 {
    
        $outputText = $outputText + "`nDownloading Software..."
        $State.Output = $outputText
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Cert_Updates.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Cert_Updates.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Delete_Old.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Delete_Old.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Search.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Search.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Self_Signed.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Self_Signed.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Configurator.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Configurator.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Updater.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Updater.ps1"
        $outputText = $outputText + "`nSoftware Downloaded..."
        $State.Output = $outputText
        softInstallS5
}

function softInstallS5 {
        $outputText = $outputText + "`nCreating Shortcuts and setting to run as admin..."
        $State.Output = $outputText
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\ACM_Configurator.lnk")
        $Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Shortcut.Arguments = "-noexit -ExecutionPolicy Bypass -File ""C:\Program Files\ACM\Software\ACM_Configurator.ps1"""
        $Shortcut.RelativePath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $shortcut.IconLocation = "shell32.dll,21"
        $Shortcut.Save()
        
        $bytes = [System.IO.File]::ReadAllBytes("C:\Users\Public\Desktop\ACM_Configurator.lnk")
        $bytes[0x15] = $bytes[0x15] -bor 0x20 #set byte 21 (0x15) bit 6 (0x20) ON
        [System.IO.File]::WriteAllBytes("C:\Users\Public\Desktop\ACM_Configurator.lnk", $bytes)
        
        $outputText = $outputText + "`nShortcuts Created..."
        $State.Output = $outputText
        softInstallS6
}

function softInstallS6 {
    
        $outputText = $outputText + "`nStarting Install Log..."
        $State.Output = $outputText
        $installTimeStop = Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"
        $UserName = $env:UserName
        $ComputerName = $env:computername
            $InstallLogFile = "C:\Program Files\ACM\Logs\install.log"
            $InstallLogOutput = "","","Install Date and Time",$installTimeStart,"","Install Stop Time",$installTimeStop,"","Computer Name:",$ComputerName,"","Installed By:",$UserName,"","Folder_Creation_Errors",$folderCreateErrors,"" #,"Download_File_Errors",$downloadFilesErrors,"","Create_Shortcut_Errors",$shortcutsCreated,
            $InstallLogOutput  | Out-File -FilePath $InstallLogFile -append
         $outputText = $outputText + "Install Log written..."
        $State.Output = $outputText
        Start-Sleep -seconds 3
        $outputText = "Install complete." 
        $outputText = $outputText + "`nPlease close the Installer with the X in the upper right hand corner. You will find the Configurator Shortcut on your desktop. Click that to get started."
        $State.Output = $outputText
}







#endregion 
#region section

#endregion 


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }


$Install_Button.Add_Click({install_button $this $_})

$State = [PSCustomObject]@{}


Function Set-Binding {
    Param($Target,$Property,$Index,$Name,$UpdateSourceTrigger)
 
    $Binding = New-Object System.Windows.Data.Binding
    $Binding.Path = "["+$Index+"]"
    $Binding.Mode = [System.Windows.Data.BindingMode]::TwoWay
    if($UpdateSourceTrigger -ne $null){$Binding.UpdateSourceTrigger = $UpdateSourceTrigger}


    [void]$Target.SetBinding($Property,$Binding)
}

function FillDataContext($props){

    For ($i=0; $i -lt $props.Length; $i++) {
   
   $prop = $props[$i]
   $DataContext.Add($DataObject."$prop")
   
    $getter = [scriptblock]::Create("Write-Output `$DataContext['$i'] -noenumerate")
    $setter = [scriptblock]::Create("param(`$val) return `$DataContext['$i']=`$val")
    $State | Add-Member -Name $prop -MemberType ScriptProperty -Value  $getter -SecondValue $setter
               
       }
   }



$DataObject =  ConvertFrom-Json @"

{
     "Output" : "",
     "Storage_ACCT_Name" : "",
     "Storage_ACCT_Key" : "",
     "Storage_ACCT_FS" : ""
     
}

"@

$DataContext = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
FillDataContext @("Output","Storage_ACCT_Name","Storage_ACCT_Key","Storage_ACCT_FS") 

$Window.DataContext = $DataContext
Set-Binding -Target $lo64cwmsdcbuf -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 0 -Name "Output"  
Set-Binding -Target $lo64cwmsu38x4 -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 1 -Name "Storage_ACCT_Name"  
Set-Binding -Target $lo64cwmt9j7ie -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 3 -Name "Storage_ACCT_FS"  
Set-Binding -Target $lo64cwmttvvb8 -Property $([System.Windows.Controls.TextBox]::TextProperty) -Index 2 -Name "Storage_ACCT_Key"  




$Global:SyncHash = [HashTable]::Synchronized(@{})
$SyncHash.Window = $Window
$Jobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$initialSessionState = [initialsessionstate]::CreateDefault()

Function Start-RunspaceTask
{
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True,Position=0)][ScriptBlock]$ScriptBlock,
          [Parameter(Mandatory=$True,Position=1)][PSObject[]]$ProxyVars)
            
    $Runspace = [RunspaceFactory]::CreateRunspace($InitialSessionState)
    $Runspace.ApartmentState = 'STA'
    $Runspace.ThreadOptions  = 'ReuseThread'
    $Runspace.Open()
    ForEach($Var in $ProxyVars){$Runspace.SessionStateProxy.SetVariable($Var.Name, $Var.Variable)}
    $Thread = [PowerShell]::Create('NewRunspace')
    $Thread.AddScript($ScriptBlock) | Out-Null
    $Thread.Runspace = $Runspace
    [Void]$Jobs.Add([PSObject]@{ PowerShell = $Thread ; Runspace = $Thread.BeginInvoke() })
}

$JobCleanupScript = {
    Do
    {    
        ForEach($Job in $Jobs)
        {            
            If($Job.Runspace.IsCompleted)
            {
                [Void]$Job.Powershell.EndInvoke($Job.Runspace)
                $Job.PowerShell.Runspace.Close()
                $Job.PowerShell.Runspace.Dispose()
                $Job.Powershell.Dispose()
                
                $Jobs.Remove($Job)
            }
        }

        Start-Sleep -Seconds 1
    }
    While ($SyncHash.CleanupJobs)
}

Get-ChildItem Function: | Where-Object {$_.name -notlike "*:*"} |  select name -ExpandProperty name |
ForEach-Object {       
    $Definition = Get-Content "function:$_" -ErrorAction Stop
    $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
    $InitialSessionState.Commands.Add($SessionStateFunction)
}


$Window.Add_Closed({
    Write-Verbose 'Halt runspace cleanup job processing'
    $SyncHash.CleanupJobs = $False
})

$SyncHash.CleanupJobs = $True
function Async($scriptBlock){ Start-RunspaceTask $scriptBlock @([PSObject]@{ Name='DataContext' ; Variable=$DataContext},[PSObject]@{Name="State"; Variable=$State},[PSObject]@{Name = "SyncHash";Variable = $SyncHash})}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name='Jobs' ; Variable=$Jobs },[PSObject]@{Name = "SyncHash";Variable = $SyncHash})



$Window.ShowDialog()

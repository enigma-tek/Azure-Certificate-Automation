###----Functions----###
Function Close_Configurator {
    Clear-Host

    Write-Host  -foregroundColor Magenta "`n `n      ============ ACM Update Utility ============ `n"
    Write-Host  -foregroundColor Yellow "    This utility will close the Configurator utility and then download the new versions"
    Write-Host  -foregroundColor Yellow "    of the Configurator Utility and the Orchestrator. Once that is complete it will update" 
    Write-Host  -foregroundColor Yellow "    the versions file, re-launch the Configurator and then exit. There is a log file "
    Write-Host  -foregroundColor Yellow "    created under "C:\Program Files\Enigma-Tek\ACM\Logs\Updater\""
  
        Write-Host -ForegroundColor Green "`n `n Closing Configurator..."
            try {
                $FormPID = Get-Content "C:\Program Files\Enigma-Tek\ACM\Temp\tempPid.tmp" -Verbose
                taskkill /pid $FormPID
        Write-Host -ForegroundColor Green "Configurator Closed"
        Write-Host -ForegroundColor Green "Starting Download Function..."

                Start-Sleep -Seconds 4
                ASPC_Software_Download
            } catch {
                Write-Host -ForegroundColor Red "There was an issue closing the Configurator. Closing the Updater. Please review log file"
                Updater_Exit_Script
            }
}

Function ACM_Software_Download {

    Clear-Host

    Write-Host  -foregroundColor Magenta "`n `n      ============ ACM Update Utility ============ `n"
    Write-Host  -foregroundColor Yellow "    This utility will close the Configurator utility and then download the new versions"
    Write-Host  -foregroundColor Yellow "    of the Configurator Utility and the Orchestrator. Once that is complete it will update" 
    Write-Host  -foregroundColor Yellow "    the versions file, re-launch the Configurator and then exit. There is a log file "
    Write-Host  -foregroundColor Yellow "    created under "C:\Program Files\Enigma-Tek\ACM\Logs\Updater\""

        Write-Host -ForegroundColor Green "`n `n Removing old files..."
            try {
            Remove-Item -Path "C:\Program Files\Enigma-Tek\ACM\Software\ASPC_Engine.ps1" -Verbose
            Remove-Item -Path "C:\Program Files\Enigma-Tek\ACM\Software\ASPC_Configurator.ps1" -Verbose
            Remove-Item -Path "C:\Program Files\Enigma-Tek\ACM\Configs\Version\versions.json" -Verbose
            Remove-Item -Path "C:\Program Files\Enigma-Tek\ACM\Temp\tempPid.tmp" -Verbose
            Write-Host -ForegroundColor Green "Files removed"

            Write-Host -ForegroundColor Green "Downloading Orchestrator..."
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Orchestrator" -OutFile "C:\Program Files\ACM\Software\ACM_Orchestrator.ps1"
            Write-Host -ForegroundColor Green "Orchestrator Downloaded..."

            Write-Host -ForegroundColor Green "Downloading Configurator..."
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/ACM_Configurator.ps1" -OutFile "C:\Program Files\ACM\Software\ACM_Configurator.ps1"
            Write-Host -ForegroundColor Green "Configurator Downloaded..."

            Write-Host -ForegroundColor Green "Updating Versions File..."
            Invoke-WebRequest -Uri "https://raw.githubusercontent.com/enigma-tek/AzureCertificateManager_Pub/main/ACM_Files/versions.json" -OutFile "C:\Program Files\ACM\Config\Version\versions.json"
            Write-Host -ForegroundColor Green "Versions File Downloaded..."
            Start-Sleep -Seconds 4
            Updater_Exit_Script
            
            
            } catch {
             Write-Host -ForegroundColor Red "There was an issue removing the old files. Closing the Updater. Please review log file"
             Updater_Exit_Script
            }
}

Function  Updater_Exit_Script {

    Clear-Host

     Write-Host  -foregroundColor Magenta "`n `n      ============ ACM Update Utility ============ `n"
    Write-Host  -foregroundColor Yellow "    This utility will close the Configurator utility and then download the new versions"
    Write-Host  -foregroundColor Yellow "    of the Configurator Utility and the Orchestrator. Once that is complete it will update" 
    Write-Host  -foregroundColor Yellow "    the versions file, re-launch the Configurator and then exit. There is a log file "
    Write-Host  -foregroundColor Yellow "    created under "C:\Program Files\Enigma-Tek\ACM\Logs\Updater\""

        Write-Host -ForegroundColor Green "`nExit Script function called"
        Write-Host -ForegroundColor Green "Opening Configurator.."
        $powershellPath = "$env:windir\system32\windowspowershell\v1.0\powershell.exe"
        $scriptPath = "C:\Program Files\Enigma-Tek\ACM\Software\ACM_Configurator.ps1"
        Start-Process $powershellPath -ArgumentList "-ExecutionPolicy Bypass ""& '$scriptPath'"""


        $scriptStopDateTime = Get-Date -Format HH.mm-MM.dd.yyyy
        Write-Host "Script stopped at " $scriptStopDateTime
        Stop-Transcript

        Write-Host -ForegroundColor Red "Script exiting in 5 seconds..."
        Start-Sleep -Seconds 5
        Exit
}


###----SCRIPT STARTS HERE---###
#Set the Title
$Title = "ACM Update Utility"
$host.UI.RawUI.WindowTitle = $Title
#Set the windows size
Function Set-WindowSize {
Param([int]$ConWidth=$host.ui.rawui.windowsize.width,
      [int]$ConHeight=$host.ui.rawui.windowsize.heigth)
    $size=New-Object System.Management.Automation.Host.Size($ConWidth,$ConHeight) 
    $host.ui.rawui.WindowSize=$size   
}
Set-WindowSize 199 80 *>$null
#Start transcript (log file)
$scriptStartDateTime = Get-Date -Format HH.mm-MM.dd.yyyy
Start-Transcript -Path "C:\Program Files\Enigma-Tek\ACM\Logs\Updater\ACM_Updater_Log_$scriptStartDateTime.log"
Write-Host "Starting script at " $scriptStartDateTime
#Calling first function
Close_Configurator
startDisplay

param (
    [switch]$InstallExchange
)
$SourcePath = 'https://raw.githubusercontent.com/michaelfolin/O365MasterClass/master/'
$LabFilesSource = "{0}{1}.zip" -f $SourcePath,$env:COMPUTERNAME
# Install mandatory features on all machines
$AdminUsername = 'corp\sysadmin'
$AdminPassword = (ConvertTo-SecureString -AsPlainText 'AA11AAss' -Force)
$Credentials = (New-Object System.Management.Automation.PSCredential -ArgumentList $AdminUsername, $AdminPassword)

Add-WindowsFeature Telnet-Client
Get-NetFirewallRule -DisplayName "File*" | Enable-NetFirewallRule 
Get-NetFirewallRule -DisplayName "*RPC*" | Enable-NetFirewallRule
New-Item -ItemType Directory -Path C:\Temp -Force
New-Item -ItemType Directory -Path C:\LabFiles -Force
function Expand-ZIPFile {
param (
    $File, 
    $Destination
)
     $shell = New-Object -ComObject Shell.Application
     $zip = $shell.NameSpace($file)
     foreach($item in $zip.items()) {
        $shell.Namespace($destination).copyhere($item)
     }
 }

function Invoke-Win10VDIOpt {
        <#
    .SYNOPSIS
        This script configures Windows 10 with minimal configuration for VDI.
    .DESCRIPTION
        This script configures Windows 10 with minimal configuration for VDI.
    
        // ============== 
        // General Advice 
        // ============== 

        Before finalizing the image perform the following tasks: 
        - Ensure no unwanted startup files by using autoruns.exe from SysInternals 
        - Run the Disk Cleanup tool as administrator and delete all temporary files and system restore points
        - Run disk defrag and consolidate free space: defrag c: /v /x
        - Reboot the machine 6 times and wait 120 seconds after logging on before performing the next reboot (boot prefetch training)
        - Run disk defrag and optimize boot files: defrag c: /v /b
        - If using a dynamic virtual disk, use the vendor's utilities to perform a "shrink" operation

        // ************* 
        // *  CAUTION  * 
        // ************* 

        THIS SCRIPT MAKES CONSIDERABLE CHANGES TO THE DEFAULT CONFIGURATION OF WINDOWS.

        Please review this script THOROUGHLY before applying to your virtual machine, and disable changes below as necessary to suit your current
        environment.

        This script is provided AS-IS - usage of this source assumes that you are at the very least familiar with PowerShell, and the tools used
        to create and debug this script.

        In other words, if you break it, you get to keep the pieces.
    .PARAMETER NoWarn
        Removes the warning prompts at the beginning and end of the script - do this only when you're sure everything works properly!
    .EXAMPLE
        .\ConfigWin10asVDI.ps1 -NoWarn $true
    .NOTES
        Author:       Carl Luberti
        Last Update:  12th November 2015
        Version:      1.0.2
    .LOG
        1.0.1 - modified sc command to sc.exe to prevent PS from invoking set-content
        1.0.2 - modified Universal Application section to avoid issues with CopyProfile, updated onedrive removal, updated for TH2
    #>


    # Parse Params:
    [CmdletBinding()]
    Param(
        [Parameter(
            Position=0,
            Mandatory=$False,
            HelpMessage="True or False, do you want to see the warning prompts"
            )] 
            [bool] $NoWarn = $true
        )


    # Throw caution (to the wind?) - show if NoWarn param is not passed, or passed as $false:
    If ($NoWarn -eq $False)
    {
        Write-Host "THIS SCRIPT MAKES CONSIDERABLE CHANGES TO THE DEFAULT CONFIGURATION OF WINDOWS." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please review this script THOROUGHLY before applying to your virtual machine, and disable changes below as necessary to suit your current environment." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "This script is provided AS-IS - usage of this source assumes that you are at the very least familiar with PowerShell, and the tools used to create and debug this script." -ForegroundColor Yellow
        Write-Host ""
        Write-Host ""
        Write-Host "In other words, if you break it, you get to keep the pieces." -ForegroundColor Magenta
        Write-Host ""
        Write-Host ""
    }


    $ProgressPreference = "SilentlyContinue"
    $ErrorActionPreference = "SilentlyContinue"


    # Validate Windows 10 Enterprise:
    $Edition = Get-WindowsEdition -Online
    If ($Edition.Edition -ne "Enterprise")
    {
        Write-Host "This is not an Enterprise SKU of Windows 10, exiting." -ForegroundColor Red
        Write-Host ""
        Exit
    }


    # Configure Constants:
    $BranchCache = "False"
    $Cortana = "False"
    $DiagService = "False"
    $EAPService = "False"
    $EFS = "False"
    $FileHistoryService = "False"
    $iSCSI = "False"
    $MachPass = "True"
    $MSSignInService = "True"
    $OneDrive = "True"
    $PeerCache = "False"
    $Search = "True"
    $SMB1 = "False"
    $SMBPerf = "False"
    $Themes = "True"
    $Touch = "False"

    $StartApps = "False"
    $AllStartApps = "True"

    $Install_NetFX3 = "False"
    $NetFX3_Source = "D:\Sources\SxS"

    $RDPEnable = 1
    $RDPFirewallOpen = 1
    $NLAEnable = 0


    # Set up additional registry drives:
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null


    # Get list of Provisioned Start Screen Apps
    $Apps = Get-ProvisionedAppxPackage -Online


    # // ============
    # // Begin Config
    # // ============


    # Set VM to High Perf scheme:
    Write-Host "Setting VM to High Performance Power Scheme..." -ForegroundColor Green
    Write-Host ""
    POWERCFG -SetActive '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'


    #Install NetFX3
    If ($Install_NetFX3 -eq "True")
    {
        Write-Host "Installing .NET 3.5..." -ForegroundColor Green
        dism /online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:$NetFX3_Source /NoRestart
        Write-Host ""
        Write-Host ""
    }


    # Remove (Almost All) Inbox Universal Apps:
    If ($StartApps -eq "False")
    {
        Write-Host "Removing (most) built-in Universal Apps..." -ForegroundColor Yellow
        Write-Host ""
    
        Write-Host "Removing Candy Crush App..." -ForegroundColor Green
        Get-AppxPackage -AllUsers | Where-Object {$_.Name -like "king.com*"} | Remove-AppxPackage
        Write-Host "Removing Twitter App..." -ForegroundColor Green
        Get-AppxPackage -AllUsers | Where-Object {$_.Name -like "*Twitter"} | Remove-AppxPackage
    
        ForEach ($App in $Apps)
        {
            # News / Sports / Weather
            If ($App.DisplayName -eq "Microsoft.BingFinance")
            {
                Write-Host "Removing Finance App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.BingNews")
            {
                Write-Host "Removing News App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.BingSports")
            {
                Write-Host "Removing Sports App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.BingWeather")
            {
                Write-Host "Removing Weather App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            # Help / "Get" Apps
            If ($App.DisplayName -eq "Microsoft.Getstarted")
            {
                Write-Host "Removing Get Started App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.SkypeApp")
            {
                Write-Host "Removing Get Skype App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.MicrosoftOfficeHub")
            {
                Write-Host "Removing Get Office App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            # Games / XBox apps
            If ($App.DisplayName -eq "Microsoft.XboxApp")
            {
                Write-Host "Removing XBox App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.ZuneMusic")
            {
                Write-Host "Removing Groove Music App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.ZuneVideo")
            {
                Write-Host "Removing Movies & TV App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.MicrosoftSolitaireCollection")
            {
                Write-Host "Removing Microsoft Solitaire Collection App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            # Others
            If ($App.DisplayName -eq "Microsoft.3DBuilder")
            {
                Write-Host "Removing 3D Builder App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.People")
            {
                Write-Host "Removing People App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.Windows.Photos")
            {
                Write-Host "Removing Photos App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.WindowsAlarms")
            {
                Write-Host "Removing Alarms App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            <#
            If ($App.DisplayName -eq "Microsoft.WindowsCalculator")
            {
                Write-Host "Removing Calculator Store App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }
            #>

            If ($App.DisplayName -eq "Microsoft.WindowsCamera")
            {
                Write-Host "Removing Camera App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.WindowsMaps")
            {
                Write-Host "Removing Maps App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.WindowsPhone")
            {
                Write-Host "Removing Phone Companion App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }

            If ($App.DisplayName -eq "Microsoft.WindowsSoundRecorder")
            {
                Write-Host "Removing Voice Recorder App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }
        
            If ($App.DisplayName -eq "Microsoft.Office.Sway")
            {
                Write-Host "Removing Office Sway App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }
        
            If ($App.DisplayName -eq "Microsoft.Messaging")
            {
                Write-Host "Removing Messaging App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }
        
            If ($App.DisplayName -eq "Microsoft.ConnectivityStore")
            {
                Write-Host "Removing Connectivity Store helper App..." -ForegroundColor Yellow
                Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                Remove-AppxPackage -Package $App.PackageName | Out-Null
            }
        }

        Start-Sleep -Seconds 5
        Write-Host ""
        Write-Host ""

        # Remove (the rest of the) Inbox Universal Apps:
        If ($AllStartApps -eq "False")
        {
            Write-Host "Removing (the rest of the) built-in Universal Apps..." -ForegroundColor Magenta
            Write-Host ""
            ForEach ($App in $Apps)
            {
                If ($App.DisplayName -eq "Microsoft.Office.OneNote")
                {
                    Write-Host "Removing OneNote App..." -ForegroundColor Magenta
                    Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                    Remove-AppxPackage -Package $App.PackageName | Out-Null
                }

                If ($App.DisplayName -eq "Microsoft.windowscommunicationsapps")
                {
                    Write-Host "Removing People, Mail, and Calendar Apps support..." -ForegroundColor Magenta
                    Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                    Remove-AppxPackage -Package $App.PackageName | Out-Null
                }
            
                If ($App.DisplayName -eq "Microsoft.CommsPhone")
                {
                    Write-Host "Removing CommsPhone helper App..." -ForegroundColor Yellow
                    Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                    Remove-AppxPackage -Package $App.PackageName | Out-Null
                }

                If ($App.DisplayName -eq "Microsoft.WindowsStore")
                {
                    Write-Host "Removing Store App..." -ForegroundColor Red
                    Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName | Out-Null
                    Remove-AppxPackage -Package $App.PackageName | Out-Null
                }
            }
            Start-Sleep -Seconds 5
            Write-Host ""
            Write-Host ""
        }
    }


    # Disable Cortana:
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\' -Name 'Windows Search' | Out-Null
    If ($Cortana -eq "False")
    {
        Write-Host "Disabling Cortana..." -ForegroundColor Yellow
        Write-Host ""
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name 'AllowCortana' -PropertyType DWORD -Value '0' | Out-Null
    }


    # Remove OneDrive:
    If ($OneDrive -eq "False")
    {
        # Remove OneDrive (not guaranteed to be permanent - see https://support.office.com/en-US/article/Turn-off-or-uninstall-OneDrive-f32a17ce-3336-40fe-9c38-6efb09f944b0):
        Write-Host "Removing OneDrive..." -ForegroundColor Yellow
        C:\Windows\SysWOW64\OneDriveSetup.exe /uninstall
        Start-Sleep -Seconds 30
        New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\' -Name 'Skydrive' | Out-Null
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Skydrive' -Name 'DisableFileSync' -PropertyType DWORD -Value '1' | Out-Null
        New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Skydrive' -Name 'DisableLibrariesDefaultSaveToSkyDrive' -PropertyType DWORD -Value '1' | Out-Null 
        Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}' -Recurse
        Remove-Item -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}' -Recurse
        Set-ItemProperty -Path 'HKCR:\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Name 'System.IsPinnedToNameSpaceTree' -Value '0'
        Set-ItemProperty -Path 'HKCR:\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Name 'System.IsPinnedToNameSpaceTree' -Value '0' 
    }


    # Set PeerCaching to Disabled (0) or Local Network PCs only (1):
    If ($PeerCache -eq "False")
    {
        Write-Host "Disabling PeerCaching..." -ForegroundColor Yellow
        Write-Host ""
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config' -Name 'DODownloadMode' -Value '0'
    }
    Else
    {
        Write-Host "Configuring PeerCaching..." -ForegroundColor Cyan
        Write-Host ""
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Config' -Name 'DODownloadMode' -Value '1'
    }


    # Disable Services:
    Write-Host "Configuring Services..." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Disabling AllJoyn Router Service..." -ForegroundColor Cyan
    Set-Service AJRouter -StartupType Disabled

    Write-Host "Disabling Application Layer Gateway Service..." -ForegroundColor Cyan
    Set-Service ALG -StartupType Disabled


    Write-Host "Disabling Bitlocker Drive Encryption Service..." -ForegroundColor Cyan
    Set-Service BDESVC -StartupType Disabled

    Write-Host "Disabling Block Level Backup Engine Service..." -ForegroundColor Cyan
    Set-Service wbengine -StartupType Disabled

    Write-Host "Disabling Bluetooth Handsfree Service..." -ForegroundColor Cyan
    Set-Service BthHFSrv -StartupType Disabled

    Write-Host "Disabling Bluetooth Support Service..." -ForegroundColor Cyan
    Set-Service bthserv -StartupType Disabled

    If ($BranchCache -eq "False")
    {
        Write-Host "Disabling BranchCache Service..." -ForegroundColor Yellow
        Set-Service PeerDistSvc -StartupType Disabled
    }

    Write-Host "Disabling Computer Browser Service..." -ForegroundColor Cyan
    Set-Service Browser -StartupType Disabled

    Write-Host "Disabling Device Association Service..." -ForegroundColor Cyan
    Set-Service DeviceAssociationService -StartupType Disabled

    Write-Host "Disabling Device Setup Manager Service..." -ForegroundColor Cyan
    Set-Service DsmSvc -StartupType Disabled

    Write-Host "Disabling Diagnostic Policy Service..." -ForegroundColor Cyan
    Set-Service DPS -StartupType Disabled

    Write-Host "Disabling Diagnostic Service Host Service..." -ForegroundColor Cyan
    Set-Service WdiServiceHost -StartupType Disabled

    Write-Host "Disabling Diagnostic System Host Service..." -ForegroundColor Cyan
    Set-Service WdiSystemHost -StartupType Disabled

    If ($DiagService -eq "False")
    {
        Write-Host "Disabling Diagnostics Tracking Service..." -ForegroundColor Yellow
        Set-Service DiagTrack -StartupType Disabled
    }

    If ($EFS -eq "False")
    {
        Write-Host "Disabling Encrypting File System Service..." -ForegroundColor Yellow
        Set-Service EFS -StartupType Disabled
    }

    If ($EAPService -eq "False")
    {
        Write-Host "Disabling Extensible Authentication Protocol Service..." -ForegroundColor Yellow
        Set-Service Eaphost -StartupType Disabled
    }

    Write-Host "Disabling Fax Service..." -ForegroundColor Cyan
    Set-Service Fax -StartupType Disabled

    Write-Host "Disabling Function Discovery Resource Publication Service..." -ForegroundColor Cyan
    Set-Service FDResPub -StartupType Disabled

    If ($FileHistoryService -eq "False")
    {
        Write-Host "Disabling File History Service..." -ForegroundColor Yellow
        Set-Service fhsvc -StartupType Disabled
    }

    Write-Host "Disabling Geolocation Service..." -ForegroundColor Cyan
    Set-Service lfsvc -StartupType Disabled

    Write-Host "Disabling Home Group Listener Service..." -ForegroundColor Cyan
    Set-Service HomeGroupListener -StartupType Disabled

    Write-Host "Disabling Home Group Provider Service..." -ForegroundColor Cyan
    Set-Service HomeGroupProvider -StartupType Disabled

    Write-Host "Disabling Home Group Provider Service..." -ForegroundColor Cyan
    Set-Service HomeGroupProvider -StartupType Disabled

    Write-Host "Disabling Internet Connection Sharing (ICS) Service..." -ForegroundColor Cyan
    Set-Service SharedAccess -StartupType Disabled

    If ($MSSignInService -eq "False")
    {
        Write-Host "Disabling Microsoft Account Sign-in Assistant Service..." -ForegroundColor Yellow
        Set-Service wlidsvc -StartupType Disabled
    }

    If ($iSCSI -eq "False")
    {
        Write-Host "Disabling Microsoft iSCSI Initiator Service..." -ForegroundColor Yellow
        Set-Service MSiSCSI -StartupType Disabled
    }

    Write-Host "Disabling Microsoft Software Shadow Copy Provider Service..." -ForegroundColor Cyan
    Set-Service swprv -StartupType Disabled

    Write-Host "Disabling Microsoft Storage Spaces SMP Service..." -ForegroundColor Cyan
    Set-Service swprv -StartupType Disabled

    Write-Host "Disabling Offline Files Service..." -ForegroundColor Cyan
    Set-Service CscService -StartupType Disabled

    Write-Host "Disabling Optimize drives Service..." -ForegroundColor Cyan
    Set-Service defragsvc -StartupType Disabled

    Write-Host "Disabling Program Compatibility Assistant Service..." -ForegroundColor Cyan
    Set-Service PcaSvc -StartupType Disabled

    Write-Host "Disabling Quality Windows Audio Video Experience Service..." -ForegroundColor Cyan
    Set-Service QWAVE -StartupType Disabled

    Write-Host "Disabling Retail Demo Service..." -ForegroundColor Cyan
    Set-Service RetailDemo -StartupType Disabled

    Write-Host "Disabling Secure Socket Tunneling Protocol Service..." -ForegroundColor Cyan
    Set-Service SstpSvc -StartupType Disabled

    Write-Host "Disabling Sensor Data Service..." -ForegroundColor Cyan
    Set-Service SensorDataService -StartupType Disabled

    Write-Host "Disabling Sensor Monitoring Service..." -ForegroundColor Cyan
    Set-Service SensrSvc -StartupType Disabled

    Write-Host "Disabling Sensor Service..." -ForegroundColor Cyan
    Set-Service SensorService -StartupType Disabled

    Write-Host "Disabling Shell Hardware Detection Service..." -ForegroundColor Cyan
    Set-Service ShellHWDetection -StartupType Disabled

    Write-Host "Disabling SNMP Trap Service..." -ForegroundColor Cyan
    Set-Service SNMPTRAP -StartupType Disabled

    Write-Host "Disabling Spot Verifier Service..." -ForegroundColor Cyan
    Set-Service svsvc -StartupType Disabled

    Write-Host "Disabling SSDP Discovery Service..." -ForegroundColor Cyan
    Set-Service SSDPSRV -StartupType Disabled

    Write-Host "Disabling Still Image Acquisition Events Service..." -ForegroundColor Cyan
    Set-Service WiaRpc -StartupType Disabled

    Write-Host "Disabling Telephony Service..." -ForegroundColor Cyan
    Set-Service TapiSrv -StartupType Disabled

    If ($Themes -eq "False")
    {
        Write-Host "Disabling Themes Service..." -ForegroundColor Yellow
        Set-Service Themes -StartupType Disabled
    }

    If ($Touch -eq "False")
    {
        Write-Host "Disabling Touch Keyboard and Handwriting Panel Service..." -ForegroundColor Yellow
        Set-Service TabletInputService -StartupType Disabled
    }

    Write-Host "Disabling UPnP Device Host Service..." -ForegroundColor Cyan
    Set-Service upnphost -StartupType Disabled

    Write-Host "Disabling Volume Shadow Copy Service..." -ForegroundColor Cyan
    Set-Service VSS -StartupType Disabled

    Write-Host "Disabling Windows Color System Service..." -ForegroundColor Cyan
    Set-Service WcsPlugInService -StartupType Disabled

    Write-Host "Disabling Windows Connect Now - Config Registrar Service..." -ForegroundColor Cyan
    Set-Service wcncsvc -StartupType Disabled

    Write-Host "Disabling Windows Error Reporting Service..." -ForegroundColor Cyan
    Set-Service WerSvc -StartupType Disabled

    Write-Host "Disabling Windows Image Acquisition (WIA) Service..." -ForegroundColor Cyan
    Set-Service stisvc -StartupType Disabled

    Write-Host "Disabling Windows Media Player Network Sharing Service..." -ForegroundColor Cyan
    Set-Service WMPNetworkSvc -StartupType Disabled

    Write-Host "Disabling Windows Mobile Hotspot Service..." -ForegroundColor Cyan
    Set-Service icssvc -StartupType Disabled

    If ($Search -eq "False")
    {
        Write-Host "Disabling Windows Search Service..." -ForegroundColor Yellow
        Set-Service WSearch -StartupType Disabled
    }

    Write-Host "Disabling WLAN AutoConfig Service..." -ForegroundColor Cyan
    Set-Service WlanSvc -StartupType Disabled

    Write-Host "Disabling WWAN AutoConfig Service..." -ForegroundColor Cyan
    Set-Service WwanSvc -StartupType Disabled

    Write-Host "Disabling Xbox Live Auth Manager Service..." -ForegroundColor Cyan
    Set-Service XblAuthManager -StartupType Disabled

    Write-Host "Disabling Xbox Live Game Save Service..." -ForegroundColor Cyan
    Set-Service XblGameSave -StartupType Disabled

    Write-Host "Disabling Xbox Live Networking Service Service..." -ForegroundColor Cyan
    Set-Service XboxNetApiSvc -StartupType Disabled
    Write-Host ""


    # Reconfigure / Change Services:
    Write-Host "Configuring Network List Service to start Automatic..." -ForegroundColor Green
    Write-Host ""
    Set-Service netprofm -StartupType Automatic
    Write-Host ""

    Write-Host "Configuring Windows Update Service to run in standalone svchost..." -ForegroundColor Cyan
    Write-Host ""
    sc.exe config wuauserv type= own
    Write-Host ""


    # Disable Scheduled Tasks:
    Write-Host "Disabling Scheduled Tasks..." -ForegroundColor Cyan
    Write-Host ""
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Application Experience\ProgramDataUpdater" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Application Experience\StartupAppTask" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Autochk\Proxy" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Bluetooth\UninstallDeviceTask" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Diagnosis\Scheduled" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticResolver" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Maintenance\WinSAT" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Maps\MapsToastTask" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Maps\MapsUpdateTask" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\MemoryDiagnostic\ProcessMemoryDiagnosticEvents" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\MemoryDiagnostic\RunFullMemoryDiagnostic" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Mobile Broadband Accounts\MNO Metadata Parser" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Power Efficiency Diagnostics\AnalyzeSystem" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Ras\MobilityManager" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Registry\RegIdleBackup" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\RetailDemo\CleanupOfflineContent" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Shell\FamilySafetyMonitor" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Shell\FamilySafetyRefresh" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\SystemRestore\SR" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\UPnP\UPnPHostConfig" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\WDI\ResolutionHost" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\Windows Media Sharing\UpdateLibrary" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\WOF\WIM-Hash-Management" | Out-Null
    Disable-ScheduledTask -TaskName "\Microsoft\Windows\WOF\WIM-Hash-Validation" | Out-Null


    # Disable Hard Disk Timeouts:
    Write-Host "Disabling Hard Disk Timeouts..." -ForegroundColor Yellow
    Write-Host ""
    POWERCFG /SETACVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0
    POWERCFG /SETDCVALUEINDEX 381b4222-f694-41f0-9685-ff5bb260df2e 0012ee47-9041-4b5d-9b77-535fba8b1442 6738e2c4-e8a5-4a42-b16a-e040e769756e 0


    # Disable Hibernate
    Write-Host "Disabling Hibernate..." -ForegroundColor Green
    Write-Host ""
    POWERCFG -h off


    # Disable Large Send Offload
    Write-Host "Disabling TCP Large Send Offload..." -ForegroundColor Green
    Write-Host ""
    New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters -Name 'DisableTaskOffload' -PropertyType DWORD -Value '1' | Out-Null


    # Disable System Restore
    Write-Host "Disabling System Restore..." -ForegroundColor Green
    Write-Host ""
    Disable-ComputerRestore -Drive "C:\"


    # Disable NTFS Last Access Timestamps
    Write-Host "Disabling NTFS Last Access Timestamps..." -ForegroundColor Yellow
    Write-Host ""
    FSUTIL behavior set disablelastaccess 1 | Out-Null

    If ($MachPass -eq "False")
    {
        # Disable Machine Account Password Changes
        Write-Host "Disabling Machine Account Password Changes..." -ForegroundColor Yellow
        Write-Host ""
        Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters' -Name 'DisablePasswordChange' -Value '1'
    }


    # Disable Memory Dumps
    Write-Host "Disabling Memory Dump Creation..." -ForegroundColor Green
    Write-Host ""
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' -Name 'CrashDumpEnabled' -Value '1'
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' -Name 'LogEvent' -Value '0'
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' -Name 'SendAlert' -Value '0'
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' -Name 'AutoReboot' -Value '1'


    # Increase Service Startup Timeout:
    Write-Host "Increasing Service Startup Timeout To 180 Seconds..." -ForegroundColor Yellow
    Write-Host ""
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'ServicesPipeTimeout' -Value '180000'


    # Increase Disk I/O Timeout to 200 Seconds:
    Write-Host "Increasing Disk I/O Timeout to 200 Seconds..." -ForegroundColor Green
    Write-Host ""
    Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Disk' -Name 'TimeOutValue' -Value '200'


    # Disable IE First Run Wizard:
    Write-Host "Disabling IE First Run Wizard..." -ForegroundColor Green
    Write-Host ""
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft' -Name 'Internet Explorer' | Out-Null
    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer' -Name 'Main' | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main' -Name DisableFirstRunCustomize -PropertyType DWORD -Value '1' | Out-Null


    # Disable New Network Dialog:
    Write-Host "Disabling New Network Dialog..." -ForegroundColor Green
    Write-Host ""
    New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Network' -Name 'NewNetworkWindowOff' | Out-Null


    If ($SMB1 -eq "False")
    {
        # Disable SMB1:
        Write-Host "Disabling SMB1 Support..." -ForegroundColor Yellow
        dism /online /Disable-Feature /FeatureName:SMB1Protocol /NoRestart
        Write-Host ""
        Write-Host ""
    }


    If ($SMBPerf -eq "True")
    {
        # SMB Modifications for performance:
        Write-Host "Changing SMB Parameters..."
        Write-Host ""
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'DisableBandwidthThrottling' -PropertyType DWORD -Value '1' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'DisableLargeMtu' -PropertyType DWORD -Value '0' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'FileInfoCacheEntriesMax' -PropertyType DWORD -Value '8000' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'DirectoryCacheEntriesMax' -PropertyType DWORD -Value '1000' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'FileNotFoundcacheEntriesMax' -PropertyType DWORD -Value '1' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters' -Name 'MaxCmds' -PropertyType DWORD -Value '8000' | Out-Null
        New-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters' -Name 'EnableWsd' -PropertyType DWORD -Value '0' | Out-Null
    }


    # Remove Previous Versions:
    Write-Host "Removing Previous Versions Capability..." -ForegroundColor Yellow
    Write-Host ""
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\\Microsoft\Windows\CurrentVersion\Explorer' -Name 'NoPreviousVersionsPage' -Value '1'


    # Change Explorer Default View:
    Write-Host "Configuring Windows Explorer..." -ForegroundColor Green
    Write-Host ""
    New-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name 'LaunchTo' -PropertyType DWORD -Value '1' | Out-Null


    # Configure Search Options:
    Write-Host "Configuring Search Options..." -ForegroundColor Green
    Write-Host ""
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name 'AllowSearchToUseLocation' -PropertyType DWORD -Value '0' | Out-Null
    New-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search' -Name 'ConnectedSearchUseWeb' -PropertyType DWORD -Value '0' | Out-Null
    New-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search' -Name 'SearchboxTaskbarMode' -PropertyType DWORD -Value '1' | Out-Null


    # Use Solid Background Color:
    Write-Host "Configuring Winlogon..." -ForegroundColor Green
    Write-Host ""
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name 'DisableLogonBackgroundImage' -Value '1'


    # DisableTransparency:
    Write-Host "Removing Transparency Effects..." -ForegroundColor Green
    Write-Host ""
    Set-ItemProperty -Path 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name 'EnableTransparency' -Value '0'


    # Configure WMI:
    Write-Host "Modifying WMI Configuration..." -ForegroundColor Green
    Write-Host ""
    $oWMI=get-wmiobject -Namespace root -Class __ProviderHostQuotaConfiguration
    $oWMI.MemoryPerHost=768*1024*1024
    $oWMI.MemoryAllHosts=1536*1024*1024
    $oWMI.put()
    Set-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\Winmgmt -Name 'Group' -Value 'COM Infrastructure'
    winmgmt /standalonehost
    Write-Host ""


    # Enable RDP:
    $RDP = Get-WmiObject -Class Win32_TerminalServiceSetting -Namespace root\CIMV2\TerminalServices -Authentication PacketPrivacy
    $Result = $RDP.SetAllowTSConnections($RDPEnable,$RDPFirewallOpen)
    if ($Result.ReturnValue -eq 0){
       Write-Host "Remote Connection settings changed sucessfully" -ForegroundColor Cyan
    } else {
       Write-Host ("Failed to change Remote Connections setting(s), return code "+$Result.ReturnValue) -ForegroundColor Red
       exit
    }
    # NLA (Network Level Authentication)
    $NLA = Get-WmiObject -Class Win32_TSGeneralSetting -Namespace root\CIMV2\TerminalServices -Authentication PacketPrivacy
    $NLA.SetUserAuthenticationRequired($NLAEnable) | Out-Null
    $NLA = Get-WmiObject -Class Win32_TSGeneralSetting -Namespace root\CIMV2\TerminalServices -Authentication PacketPrivacy
    if ($NLA.UserAuthenticationRequired -eq $NLAEnable){
       Write-Host "NLA setting changed sucessfully" -ForegroundColor Cyan
    } else {
       Write-Host "Failed to change NLA setting" -ForegroundColor Red
       exit
    }
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""


    # Did this break?:
    If ($NoWarn -eq $False)
    {
        Write-Host "This script has completed." -ForegroundColor Green
        Write-Host ""
        Write-Host "Please review output in your console for any indications of failures, and resolve as necessary." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Remember, this script is provided AS-IS - review the changes made against the expected workload of this VDI VM to validate things work properly in your environment." -ForegroundColor Magenta
        Write-Host ""
        Write-Host "Good luck! (reboot required)" -ForegroundColor White
    }



}

function Disable-InternetExplorerESC {
    $AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
    $UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
    Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0
    Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0
}

Disable-InternetExplorerESC
Set-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3' `
                 -Name '1803' -Value 0 
Set-ItemProperty -Path 'HKLM:\Software\Microsoft\Internet Explorer\Main' `
                 -Name 'Start Page' -Value 'about:blank'

function Write-Log {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param ( 
        [ValidateSet('ERROR', 'INFO')]
        [string]$LogType = 'INFO',
        $LogString
    )
    #Generate log file name based on LogType and Month
    $LogFile = 'C:\Temp\InstallationLog-{0}.log' -f (Get-Date -Format yyyyMM)
    #Construct log output
    $InputData = "{0} {1}:$LogString" -f $LogType,(Get-Date -Format u)
    #Log to file
    Add-Content -Path $LogFile -Value $InputData -Force
    Write-Output -Message $InputData
}

Start-BitsTransfer -Source $LabFilesSource -Destination c:\Temp

Expand-ZIPFile -File "c:\temp\$($env:COMPUTERNAME).zip" -Destination C:\LabFiles 

 switch ($env:COMPUTERNAME) {
    #Domain Controller
    'ACME-DC01' {
       
        #Create OU Structure
        Add-WindowsFeature RSAT-ADDS,RSAT-DNS-Server
        $OU = Get-ADOrganizationalUnit -Identity "OU=Users,OU=ACME,DC=corp,DC=acme,DC=com" -ErrorAction Ignore -WarningAction Ignore
        if (-not($OU)) {
            New-ADOrganizationalUnit -Name "ACME" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "Users" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "Groups" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "ServiceAccounts" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "Computers" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "Servers" -ProtectedFromAccidentalDeletion $false
            New-ADOrganizationalUnit -Path "OU=ACME,DC=corp,DC=acme,DC=com" -Name "Contacts" -ProtectedFromAccidentalDeletion $false
        }
        #region create test users
        $SourceFile = "{0}FirstLastEurope.csv" -f $SourcePath
        Invoke-WebRequest -Uri $SourceFile -OutFile C:\temp\FirstLastEurope.csv -UseBasicParsing
        $Names = Import-CSV C:\Temp\FirstLastEurope.csv | Select-Object -First 50
        
        $Session = New-PSSession -ConnectionUri http://acme-ex01/powershell -ConfigurationName Microsoft.Exchange -Authentication Kerberos -Credential $credentials
        Import-PSSession $Session -AllowClobber 
        $NumUsers = 50
        #Define variables
        $OU = "OU=Users,OU=ACME,DC=corp,DC=acme,DC=com"
        $Names = Import-CSV C:\temp\FirstLastEurope.csv | Select-Object -first $NumUsers
        $Password = 'Pa$$w0rd'
        $UPNSuffix = (Get-ADDomain).DnsRoot

        #Import required module ActiveDirectory
        try{
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        catch{
            throw "Module GroupPolicy not Installed"
        }

        foreach ($name in $names) {
            
            #Generate username and check for duplicates
            $firstname = $name.firstname
            $lastname = $name.lastname 

            $username = $name.firstname.Substring(0,3).tolower() + $name.lastname.Substring(0,3).tolower()
            $exit = 0
            $count = 1
            do
            { 
                try { 
                    $userexists = Get-AdUser -Identity $username
                    $username = $firstname.Substring(0,3).tolower() + $lastname.Substring(0,3).tolower() + $count++
                }
                catch {
                    $exit = 1
                }
            }
            while ($exit -eq 0)

            #Set Displayname and UserPrincipalNBame
            $displayname = "$firstname $lastname ($username)"
            if ($username -eq "alpast") {
                $upn = "{0}.{1}@{2}" -f $firstname,$lastname,(Get-ADForest).upnsuffixes[0]
            } else {
                $upn = "$username@$upnsuffix"
            }
            #Create the user
            Write-Host "Creating user $username in $ou"
            New-ADUser –Name $displayname –DisplayName $displayname `
                 –SamAccountName $username -UserPrincipalName $upn `
                 -GivenName $firstname -Surname $lastname -description "Test User" `
                 -Path $ou –Enabled $true –ChangePasswordAtLogon $false -Department $Department `
                 -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) 

        }
        Get-User -OrganizationalUnit "OU=Users,OU=ACME,DC=corp,DC=acme,DC=com" | Enable-Mailbox 
        $WrongChars = ".","|","#","å","ä"
        $Names = Import-CSV c:\temp\FirstLastEurope.csv | Select-Object -Last $WrongChars.Length  
        $i = 0
        foreach ($name in $names) {
            
            #Generate username and check for duplicates
            $firstname = $name.firstname
            $lastname = $name.lastname 

            $username = $name.firstname.Substring(0,3).tolower() + $name.lastname.Substring(0,3).tolower()
            $exit = 0
            $count = 1
            do
            { 
                try { 
                    $userexists = Get-AdUser -Identity $username
                    $username = $firstname.Substring(0,3).tolower() + $lastname.Substring(0,3).tolower() + $count++
                }
                catch {
                    $exit = 1
                }
            }
            while ($exit -eq 0)

            #Set Displayname and UserPrincipalNBame
            $displayname = "$firstname $lastname ($username)"
    
            $upn = "$username{0}@$upnsuffix" -f $WrongChars[$i]
      
            #Create the user
            Write-Host "Creating user $username in $ou"
            New-ADUser –Name $displayname –DisplayName $displayname `
                 –SamAccountName $username -UserPrincipalName $upn `
                 -GivenName $firstname -Surname $lastname -description "Test User" `
                 -Path $ou –Enabled $true –ChangePasswordAtLogon $false -Department $Department `
                 -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -force) 
            $i++

        }

        Get-ADUser -SearchBase "OU=Users,OU=ACME,DC=corp,DC=acme,DC=com" -filter * | ForEach-Object -Process {
            $AdditionalMail = "smtp:{0}@migration.target" -f $_.samaccountname
            Set-ADUser -Identity $_.samaccountname -Add @{proxyaddresses=$AdditionalMail}
        }
        1..5 | ForEach-Object {
            New-Mailbox -Room -Name "HQ Conference Room $_" -OrganizationalUnit $OU
            Set-CalendarProcessing -Identity "HQ Conference Room $_" -AutomateProcessing AutoAccept
        }
        "Sales","Finance","IT Department" | ForEach-Object {
            New-Mailbox -Shared -Name "SM-$_" -OrganizationalUnit $OU
        }
        Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize 30 | Enable-Mailbox -Archive
        New-ADObject -Type contact -DisplayName "Antony Barnwell (Consultant)" -OtherAttributes @{mail="Antony.Barnwell@corp.acme.com";proxyaddresses="SMTP:Antony.Barnwell@corp.acme.com"} -Name "Antony Barnwell (Consultant)" -path "OU=Contacts,OU=ACME,DC=corp,DC=acme,DC=com"
        New-ADObject -Type contact -DisplayName "Aurora Beers (Consultant)" -OtherAttributes @{mail="Aurora.Beers@corp.acme.com";proxyaddresses="SMTP:Aurora.Beers@corp.acme.com"} -Name "Aurora Beers (Consultant)" -path "OU=Contacts,OU=ACME,DC=corp,DC=acme,DC=com"

        #endregion create test users
        
    } 
    'ACME-EX01' {
             #Create folders    
             New-Item -ItemType Directory -Path C:\Temp\Exchange -Force
             New-Item -ItemType Directory -Path C:\Temp\ExchangePreReq -Force
             Set-Service NetTcpPortSharing -StartupType Automatic
             #Download Exchange Media
         
             #$ExchangeMedia = Get-ChildItem "C:\temp\Exchange2013-x64-cu10.exe"
             $ExchangeMedia = Get-ChildItem "C:\temp\Exchange2013-x64-cu11.exe"
             #if ($ExchangeMedia.Length -ne "1739259000") {
             if ($ExchangeMedia.Length -ne "1739511520") {
                Write-Log -LogType INFO -LogString "Starting download of Exchange Media"
                #(New-Object System.Net.WebClient).DownloadFile("https://download.microsoft.com/download/1/D/1/1D15B640-E2BB-4184-BFC5-83BC26ADD689/Exchange2013-x64-cu10.exe", "C:\temp\Exchange2013-x64-cu10.exe")
                Start-BitsTransfer https://download.microsoft.com/download/A/A/B/AAB18934-BC8F-429D-8912-6A98CBC96B07/Exchange2013-x64-cu11.exe c:\temp
                Write-Log -LogType INFO -LogString "Successfully downloaded Exchange Media"
                $ExchangeMedia = Get-ChildItem "C:\temp\Exchange2013-x64-cu11.exe"
                Write-Log -LogType INFO -LogString "Starting to extract Exchange Media"
                Start-Process -FilePath $ExchangeMedia.FullName -ArgumentList "/extract:C:\temp\Exchange" 
             } 

             do {
                Write-Output "Wating for Exchange media to get extracted..."
                Start-Sleep 15 
             } until ((Get-ChildItem c:\temp\exchange | Measure-Object -Property Length -Sum).sum -eq "26603199" ) #26622071
             Write-Log -LogType INFO -LogString "Successfully extracted Exchange Media"
             #Download and install Prereqs
             Write-Log -LogType INFO -LogString "Starting to install Windows Features"
             Add-WindowsFeature AS-HTTP-Activation, Desktop-Experience, NET-Framework-45-Features, RPC-over-HTTP-proxy, RSAT-Clustering, RSAT-Clustering-CmdInterface, RSAT-Clustering-Mgmt, RSAT-Clustering-PowerShell, Web-Mgmt-Console, WAS-Process-Model, Web-Asp-Net45, Web-Basic-Auth, Web-Client-Auth, Web-Digest-Auth, Web-Dir-Browsing, Web-Dyn-Compression, Web-Http-Errors, Web-Http-Logging, Web-Http-Redirect, Web-Http-Tracing, Web-ISAPI-Ext, Web-ISAPI-Filter, Web-Lgcy-Mgmt-Console, Web-Metabase, Web-Mgmt-Console, Web-Mgmt-Service, Web-Net-Ext45, Web-Request-Monitor, Web-Server, Web-Stat-Compression, Web-Static-Content, Web-Windows-Auth, Web-WMI, Windows-Identity-Foundation,RSAT-ADDS
             Write-Log -LogType INFO -LogString "Finished installing windows features"
             if ((Get-ChildItem C:\temp\ExchangePreReq | Measure-Object -Property Length -Sum).Sum -ne "259534152") {
                 $PreReqs = "http://download.microsoft.com/download/2/C/4/2C47A5C1-A1F3-4843-B9FE-84C0032C61EC/UcmaRuntimeSetup.exe","http://download.microsoft.com/download/A/A/3/AA345161-18B8-45AE-8DC8-DA6387264CB9/filterpack2010sp1-kb2460041-x64-fullfile-en-us.exe","http://download.microsoft.com/download/0/A/2/0A28BBFA-CBFA-4C03-A739-30CCA5E21659/FilterPack64bit.exe"
                 foreach ($PreReq in $PreReqs) {
                    $Destination = "C:\Temp\ExchangePreReq\{0}" -f $PreReq.split("/")[-1]
                    (New-Object System.Net.WebClient).DownloadFile("$PreReq", "$destination")
                    Write-Log -LogType INFO -LogString "Successfully downloaded $Destination from $PreReq"
                 }
             }
         
    }
    {$_ -match 'ADFS'} { 
        Enable-PSRemoting -Force
        Write-Log -LogType INFO -LogString "Attempting to install ADFS"
        Add-WindowsFeature ADFS-Federation -IncludeManagementTools -IncludeAllSubFeature          
        Write-Log -LogType INFO -LogString "Successfully installed ADFS"
    }
    {$_ -match 'WAP'} {
        Enable-PSRemoting -Force
        Write-Log -LogType INFO -LogString "Attempting to install WAP"
        Add-WindowsFeature DNS -IncludeManagementTools
        Add-WindowsFeature Web-application-proxy -IncludeManagementTools         
        Write-Log -LogType INFO -LogString "Successfully installed WAP"
    }
    {$_ -match 'CL'} {
        Get-Service audiosrv | Set-Service -StartupType Disabled
        Get-Service audiosrv | Stop-Service -Force
        Invoke-Win10VDIOpt
        $localGroupName = "Remote Desktop Users"
        $domainGroupName = "Domain Users"
        $DomainName = "corp.acme.com"
        $vname = $env:COMPUTERNAME
        try { 
            $adsi = [ADSI]"WinNT://$vname/$localGroupName,group" 
            $adsi.add("WinNT://$DomainName/$domainGroupName,group")  
        } catch {
        }
        if ($env:COMPUTERNAME -eq "ACME-CL01") {
            Write-Log -LogType INFO -LogString "Starting to Download Office 365 ProPlus"
            $Source = "{0}Office365ProPlus.zip" -f $SourcePath
            Start-BitsTransfer -Source $source -Destination C:\Temp
            Write-Log -LogType INFO -LogString "Successfully downloaded Office 365 ProPlus"
            Write-Log -LogType INFO -LogString "Starting to extract and Install Office 365 ProPlus"
            Expand-ZIPFile C:\temp\Office365ProPlus.zip -Destination C:\Temp
            $Installation = Start-Process -FilePath C:\temp\Office365ProPlus\setup.exe -ArgumentList "/configure C:\temp\Office365ProPlus\Intune.xml" -Wait -PassThru
            Write-Log -LogType INFO -LogString "Successfully installed Office 365 ProPlus with ExitCode $($Installation.ExitCode)"
            Write-Log -LogType INFO -LogString "Starting to download and install SharePoint Designer"
            $SPOFiles = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=35491","https://www.microsoft.com/en-us/download/confirmation.aspx?id=42009"
            $Destination = "C:\Temp"
            foreach ($DLAddress in $SPOFiles) {
                $File =  ((Invoke-WebRequest -Uri $DLAddress -UseBasicParsing).links | 
                            Where-Object -Property href -Match  -Value "msi$|exe$|docx$|bin$|zip$").href | Select-Object -Unique | Select-String "64"
                $InstallFile = "{0}\{1}" -f $destination,$file.tostring().Split("/")[-1]
                if (-not(Get-ChildItem $InstallFile)) {
                    Start-BitsTransfer -Source $File -Destination $Destination 
                }
                #Write-Log -LogType INFO -LogString "Starting to install $installfile"
                #if ($InstallFile -like "sharepoint*") {
                #    $Installation = Start-Process -FilePath $installfile -ArgumentList "/extract:C:\temp\spdesigner /quiet" -Wait -PassThru
                #    Start-BitsTransfer -Source "https://365lab.blob.core.windows.net/scripts/SPDesigner.MSP" -Destination "c:\temp\spdesigner\updates"
                #    $Installation = Start-Process -FilePath c:\temp\spdesigner\setup.exe -Wait -PassThru
                #    Write-Log -LogType INFO -LogString "Successfully installed $installfile with ExitCode $($Installation.ExitCode)"
                #}
                #$Installation = Start-Process -FilePath $installfile -ArgumentList "/quiet /norestart" -Wait -PassThru
                #Write-Log -LogType INFO -LogString "Successfully installed $installfile with ExitCode $($Installation.ExitCode)"
            }
        } 
        
    }

 }
 

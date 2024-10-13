<#
.SYNOPSIS
    ConfigMgr Client Health is a tool that validates and automatically fixes errors on Windows computers managed by Microsoft Configuration Manager.
.EXAMPLE
    .\ConfigMgrClientHealth.ps1 -Config .\Config.Xml
.EXAMPLE
    \\cm01.rodland.lab\ClientHealth$\ConfigMgrClientHealth.ps1 -Config \\cm01.rodland.lab\ClientHealth$\Config.Xml -Webservice https://cm01.rodland.lab/ConfigMgrClientHealth
.PARAMETER Config
    A single Parameter specifying the path to the configuration XML file.
.PARAMETER Webservice
    A single Parameter specIfying the URI to the ConfigMgr Client Health Webservice.
.DESCRIPTION
    ConfigMgr Client Health detects and fixes following errors:
        * ConfigMgr client is not installed
        * ConfigMgr client is assigned the correct site code
        * ConfigMgr client is upgraded to current version if not at specified minimum version
        * ConfigMgr client not able to forward state messages to management point
        * ConfigMgr client stuck in provisioning mode
        * ConfigMgr client maximum log file size
        * ConfigMgr client cache size
        * Corrupt WMI
        * Services for ConfigMgr client is not running or disabled
        * Other services can be specified to start and run and specific state
        * Hardware inventory is running at correct schedule
        * Group Policy fails to update Registry.pol
        * Pending reboot blocking updates from installing
        * ConfigMgr Client Update Handler is working correctly with Registry.pol
        * Windows Update Agent not working correctly, causing client not to receive patches
.NOTES
    You should run this with at least local administrator rights. It is recommended to run this script under the SYSTEM context.
.NOTES
    DO NOT GIVE USERS WRITE ACCESS TO THIS FILE. LOCK IT DOWN!
.NOTES
    Author: Anders Rødland
    Blog: https://www.andersrodland.com
    Twitter: @AndersRodland
.LINK
    Full documentation: https://www.andersrodland.com/configmgr-client-health/
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
Param(
    [Parameter(HelpMessage = 'Path to XML Configuration File')]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [ValidatePattern('.xml$')]
    [string]$Config,
    [Parameter(HelpMessage = 'URI to ConfigMgr Client Health Webservice')]
    [string]$Webservice
)

Begin {
    # ConfigMgr Client Health Version
    $Version = '0.8.3'
    $PowerShellVersion = [int]$PSVersionTable.PSVersion.Major
    $global:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

    #If no config file was passed in, use the default.
    If ((!$PSBoundParameters.ContainsKey('Config')) -and (!$PSBoundParameters.ContainsKey('Webservice'))) {
        $Config = Join-Path ($global:ScriptPath) "Config.xml"
        Write-Verbose "No config provided, defaulting to $Config"
    }

    Write-Verbose "Script version: $Version"
    Write-Verbose "PowerShell version: $PowerShellVersion"

    Function Test-XML {
        <#
        .SYNOPSIS
            Test the validity of an XML file
        #>
        [CmdletBinding()]
        Param ([Parameter(mandatory = $true)][ValidateNotNullorEmpty()][string]$xmlFilePath)
        # Check the file exists
        If (!(Test-Path -Path $xmlFilePath)) {
            Throw "$xmlFilePath is not valid. Please provide a valid path to the .xml config file" 
        }
        # Check for Load or Parse errors when loading the XML file
        $xml = New-Object System.Xml.XmlDocument
        Try {
            $xml.Load((Get-ChildItem -Path $xmlFilePath).FullName)
            Return $true
        } Catch [System.Xml.XmlException] {
            Write-Error "$xmlFilePath : $($_.toString())"
            Write-Error "Configuration file $Config is NOT valid XML. Script will not execute."
            Return $false
        }
    }

    # Read configuration from XML file
    If ($config) {
        If (Test-Path $Config) {
            # Test If valid XML
            If ((Test-XML -xmlFilePath $Config) -ne $true ) {
                Exit 1 
            }

            # Load XML file into variable
            Try {
                $Xml = [xml](Get-Content -Path $Config) 
            } Catch {
                $ErrorMessage = $_.Exception.Message
                $Text = "Error, could not read $Config. Check file location and share/ntfs permissions. Is XML config file damaged?"
                $Text += "`nError message: $ErrorMessage"
                Write-Error $Text
                Exit 1
            }
        } Else {
            $Text = "Error, could not access $Config. Check file location and share/ntfs permissions. Did you misspell the name?"
            Write-Error $Text
            Exit 1
        }
    }


    # Import Modules
    # Import BitsTransfer Module (Does not work on PowerShell Core (6), disable check If module failes to import.)
    If (Get-Module -ListAvailable -Name BitsTransfer) {
        Try {
            Import-Module BitsTransfer -ErrorAction stop
            $BitsCheckEnabled = $true
        } Catch {
            $BitsCheckEnabled = $false
        }
    } Else {
        $BitsCheckEnabled = $false
    }

    #region functions
    Function Get-DateTime {
        $format = (Get-XMLConfigLoggingTimeFormat).ToLower()

        # UTC Time
        If ($format -like "utc") {
            $obj = ([DateTime]::UtcNow).ToString("yyyy-MM-dd HH:mm:ss") 
        }
        # ClientLocal
        Else {
            $obj = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") 
        }

        Write-Output $obj
    }

    # Converts a DateTime object to UTC time.
    Function Get-UTCTime {
        Param([Parameter(Mandatory = $true)][DateTime]$DateTime)
        $obj = $DateTime.ToUniversalTime()
        Write-Output $obj
    }

    Function Get-Hostname {
        <#
        If ($PowerShellVersion -ge 6) { $Obj = (Get-CimInstance Win32_ComputerSystem).Name }
        Else { $Obj = (Get-WmiObject Win32_ComputerSystem).Name }
        #>
        $obj = $env:COMPUTERNAME
        Write-Output $Obj
    }

    # Update-WebService use ClientHealth Webservice to update database. RESTful API.
    Function Update-Webservice {
        Param([Parameter(Mandatory = $true)][String]$URI, $Log)

        $Hostname = Get-Hostname
        $Obj = $Log | ConvertTo-Json
        $URI = $URI + "/Clients"
        $ContentType = "application/json"

        # Detect If we use PUT or POST
        Try {
            Invoke-RestMethod -Uri "$URI/$Hostname" | Out-Null
            $Method = "PUT"
            $URI = $URI + "/$Hostname"
        } Catch {
            $Method = "POST" 
        }

        Try {
            Invoke-RestMethod -Method $Method -Uri $URI -Body $Obj -ContentType $ContentType | Out-Null 
        } Catch {
            $ExceptionMessage = $_.Exception.Message
            Write-Host "Error Invoking RestMethod $Method on URI $URI. Failed to update database using webservice. Exception: $ExceptionMessage"

        }
    }

    # Retrieve configuration from SQL using webserivce
    Function Get-ConfigFromWebservice {
        Param(
            [Parameter(Mandatory = $true)][String]$URI,
            [Parameter(Mandatory = $false)][String]$ProfileID
        )

        $URI = $URI + "/ConfigurationProfile"
        #Write-Host "ProfileID = $ProfileID"
        If ($ProfileID -ge 0) {
            $URI = $URI + "/$ProfileID" 
        }

        Write-Verbose "Retrieving configuration from webservice. URI: $URI"
        Try {
            $Obj = Invoke-RestMethod -Uri $URI
        } Catch {
            Write-Host "Error retrieving configuration from webservice $URI. Exception: $ExceptionMessage" -ForegroundColor Red
            Exit 1
        }

        Write-Output $Obj
    }

    Function Get-ConfigClientInstallPropertiesFromWebService {
        Param(
            [Parameter(Mandatory = $true)][String]$URI,
            [Parameter(Mandatory = $true)][String]$ProfileID
        )

        $URI = $URI + "/ClientInstallProperties"

        Write-Verbose "Retrieving client install properties from webservice"
        Try {
            $CIP = Invoke-RestMethod -Uri $URI
        } Catch {
            Write-Host "Error retrieving client install properties from webservice $URI. Exception: $ExceptionMessage" -ForegroundColor Red
            Exit 1
        }

        $string = $CIP | Where-Object { $_.profileId -eq $ProfileID } | Select-Object -ExpandProperty cmd
        $obj = ""

        foreach ($i in $string) {
            $obj += $i + " "
        }

        # Remove the trailing space from the last Parameter caused by the foreach loop
        $obj = $obj.Substring(0, $obj.Length - 1)
        Write-Output $Obj
    }

    Function Get-ConfigServicesFromWebservice {
        Param(
            [Parameter(Mandatory = $true)][String]$URI,
            [Parameter(Mandatory = $true)][String]$ProfileID
        )

        $URI = $URI + "/ConfigurationProfileServices"

        Write-Verbose "Retrieving client install properties from webservice"
        Try {
            $CS = Invoke-RestMethod -Uri $URI
        } Catch {
            Write-Host "Error retrieving client install properties from webservice $URI. Exception: $ExceptionMessage" -ForegroundColor Red
            Exit 1
        }

        $obj = $CS | Where-Object { $_.profileId -eq $ProfileID } | Select-Object Name, StartupType, State, Uptime



        Write-Output $Obj
    }

    Function Get-LogFileName {
        #$OS = Get-WmiObject -class Win32_OperatingSystem
        #$OSName = Get-OperatingSystem
        $logshare = Get-XMLConfigLoggingShare
        #$obj = "$logshare\$OSName\$env:computername.log"
        $obj = "$logshare\$env:computername.log"
        Write-Output $obj
    }

    Function Get-ServiceUpTime {
        Param([Parameter(Mandatory = $true)]$Name)

        Try {
            $ServiceDisplayName = (Get-Service $Name).DisplayName 
        } Catch {
            Write-Warning "The '$($Name)' service could not be found."
            Return
        }

        #First Try and get the service start time based on the last start event message in the system log.
        Try {
            [datetime]$ServiceStartTime = (Get-EventLog -LogName System -Source "Service Control Manager" -EnTryType Information -Message "*$($ServiceDisplayName)*running*" -Newest 1).TimeGenerated
            Return (New-TimeSpan -Start $ServiceStartTime -End (Get-Date)).Days
        } Catch {
            Write-Verbose "Could not get the uptime time for the '$($Name)' service from the event log.  Relying on the process instead."
        }

        #If the event log doesn't contain a start event then use the start time of the service's process.  Since processes can be shared this is less reliable.
        Try {
            If ($PowerShellVersion -ge 6) {
                $ServiceProcessID = (Get-CimInstance Win32_Service -Filter "Name='$($Name)'").ProcessID 
            } Else {
                $ServiceProcessID = (Get-WmiObject -Class Win32_Service -Filter "Name='$($Name)'").ProcessID 
            }

            [datetime]$ServiceStartTime = (Get-Process -Id $ServiceProcessID).StartTime
            Return (New-TimeSpan -Start $ServiceStartTime -End (Get-Date)).Days

        } Catch {
            Write-Warning "Could not get the uptime time for the '$($Name)' service.  Returning max value."
            Return [int]::MaxValue
        }
    }

    #Loop backwards through a Configuration Manager log file looking for the latest matching message after the start time.
    Function Search-CMLogFile {
        Param(
            [Parameter(Mandatory = $true)]$LogFile,
            [Parameter(Mandatory = $true)][String[]]$SearchStrings,
            [datetime]$StartTime = [datetime]::MinValue
        )

        #Get the log data.
        $LogData = Get-Content $LogFile

        #Loop backwards through the log file.
        :loop for ($i = ($LogData.Count - 1); $i -ge 0; $i--) {

            #Parse the log line into its parts.
            Try {
                $LogData[$i] -match '\<\!\[LOG\[(?<Message>.*)?\]LOG\]\!\>\<time=\"(?<Time>.+)(?<TZAdjust>[+|-])(?<TZOffset>\d{2,3})\"\s+date=\"(?<Date>.+)?\"\s+component=\"(?<Component>.+)?\"\s+context="(?<Context>.*)?\"\s+type=\"(?<Type>\d)?\"\s+thread=\"(?<TID>\d+)?\"\s+file=\"(?<Reference>.+)?\"\>' | Out-Null
                $LogTime = [datetime]::ParseExact($("$($matches.date) $($matches.time)"), "MM-dd-yyyy HH:mm:ss.fff", $null)
                $LogMessage = $matches.message
            } Catch {
                Write-Warning "Could not parse the line $($i) in '$($LogFile)': $($LogData[$i])"
                continue
            }

            #If we have gone beyond the start time then stop searching.
            If ($LogTime -lt $StartTime) {
                Write-Verbose "No log lines in $($LogFile) matched $($SearchStrings) before $($StartTime)."
                break loop
            }

            #Loop through each search string looking for a match.
            ForEach ($String in $SearchStrings) {
                If ($LogMessage -match $String) {
                    Write-Output $LogData[$i]
                    break loop
                }
            }
        }

        #Looped through log file without finding a match.
        #Return
    }

    Function Test-LocalLogging {
        $clientpath = Get-LocalFilesPath
        If ((Test-Path -Path $clientpath) -eq $False) {
            New-Item -Path $clientpath -ItemType Directory -Force | Out-Null 
        }
    }

    Function Out-LogFile {
        Param([Parameter(Mandatory = $false)][xml]$Xml, $Text, $Mode,
            [Parameter(Mandatory = $false)][ValidateSet(1, 2, 3, 'Information', 'Warning', 'Error')]$Severity = 1)

        switch ($Severity) {
            'Information' {
                $Severity = 1 
            }
            'Warning' {
                $Severity = 2 
            }
            'Error' {
                $Severity = 3 
            }
        }

        If ($Mode -like "Local") {
            Test-LocalLogging
            $clientpath = Get-LocalFilesPath
            $Logfile = "$clientpath\ClientHealth.log"
        } Else {
            $Logfile = Get-LogFileName 
        }

        If ($mode -like "ClientInstall" ) { 
            $Text = "ConfigMgr Client installation failed. Agent not detected 10 minutes after triggering installation." 
            $Severity = 3
        }

        foreach ($item in $Text) {
            $item = '<![LOG[' + $item + ']LOG]!>'
            $time = 'time="' + (Get-Date -Format HH:mm:ss.fff) + '+000"' #Should actually be the bias
            $date = 'date="' + (Get-Date -Format MM-dd-yyyy) + '"'
            $component = 'component="ConfigMgrClientHealth"'
            $context = 'context=""'
            $type = 'type="' + $Severity + '"'  #Severity 1=Information, 2=Warning, 3=Error
            $thread = 'thread="' + $PID + '"'
            $file = 'file=""'

            $logblock = ($time, $date, $component, $context, $type, $thread, $file) -join ' '
            $logblock = '<' + $logblock + '>'

            $item + $logblock | Out-File -Encoding utf8 -Append $logFile
        }
        # $obj | Out-File -Encoding utf8 -Append $logFile
    }

    Function Get-OperatingSystem {
        If ($PowerShellVersion -ge 6) {
            $OS = Get-CimInstance Win32_OperatingSystem 
        } Else {
            $OS = Get-WmiObject Win32_OperatingSystem 
        }


        # Handles dIfferent OS languages
        $OSArchitecture = ($OS.OSArchitecture -replace ('([^0-9])(\.*)', '')) + '-Bit'
        switch -Wildcard ($OS.Caption) {
            "*Embedded*" {
                $OSName = "Windows 7 " + $OSArchitecture 
            }
            "*Windows 7*" {
                $OSName = "Windows 7 " + $OSArchitecture 
            }
            "*Windows 8.1*" {
                $OSName = "Windows 8.1 " + $OSArchitecture 
            }
            "*Windows 10*" {
                $OSName = "Windows 10 " + $OSArchitecture 
            }
            "*Server 2008*" {
                If ($OS.Caption -like "*R2*") {
                    $OSName = "Windows Server 2008 R2 " + $OSArchitecture 
                } Else {
                    $OSName = "Windows Server 2008 " + $OSArchitecture 
                }
            }
            "*Server 2012*" {
                If ($OS.Caption -like "*R2*") {
                    $OSName = "Windows Server 2012 R2 " + $OSArchitecture 
                } Else {
                    $OSName = "Windows Server 2012 " + $OSArchitecture 
                }
            }
            "*Server 2016*" {
                $OSName = "Windows Server 2016 " + $OSArchitecture 
            }
            "*Server 2019*" {
                $OSName = "Windows Server 2019 " + $OSArchitecture 
            }
        }
        Write-Output $OSName
    }

    Function Get-MissingUpdates {
        $UpdateShare = Get-XMLConfigUpdatesShare
        $OSName = Get-OperatingSystem

        $build = $null
        If ($OSName -like "*Windows 10*") {
            $build = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber
            switch ($build) {
                10240 {
                    $OSName = $OSName + " 1507" 
                }
                10586 {
                    $OSName = $OSName + " 1511" 
                }
                14393 {
                    $OSName = $OSName + " 1607" 
                }
                15063 {
                    $OSName = $OSName + " 1703" 
                }
                16299 {
                    $OSName = $OSName + " 1709" 
                }
                17134 {
                    $OSName = $OSName + " 1803" 
                }
                17763 {
                    $OSName = $OSName + " 1809" 
                }
                default {
                    $OSName = $OSName + " Insider Preview" 
                }
            }
        }

        $Updates = $UpdateShare + "\" + $OSName + "\"
        $obj = New-Object PSObject @{}
        If ((Test-Path $Updates) -eq $true) {
            $regex = "\b(?!(KB)+(\d+)\b)\w+"
            $hotfixes = (Get-ChildItem $Updates | Select-Object -ExpandProperty Name)
            If ($PowerShellVersion -ge 6) {
                $installedUpdates = (Get-CimInstance -ClassName Win32_QuickFixEngineering).HotFixID 
            } Else {
                $installedUpdates = Get-HotFix | Select-Object -ExpandProperty HotFixID 
            }

            foreach ($hotfix in $hotfixes) {
                $kb = $hotfix -replace $regex -replace "\." -replace "-"
                If ($installedUpdates -like $kb) {
                } Else {
                    $obj.Add('Hotfix', $hotfix) 
                }
            }
        }
        Write-Output $obj
    }

    Function Get-RegistryValue {
        Param (
            [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Path,
            [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Name
        )

        Return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue).$Name
    }

    Function Set-RegistryValue {
        Param (
            [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Path,
            [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Name,
            [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$Value,
            [ValidateSet("String", "ExpandString", "Binary", "DWord", "MultiString", "Qword")]$ProperyType = "String"
        )

        #Make sure the key exists
        If (!(Test-Path $Path)) {
            New-Item $Path -Force | Out-Null
        }

        New-ItemProperty -Force -Path $Path -Name $Name -Value $Value -PropertyType $ProperyType | Out-Null
    }

    Function Get-Sitecode {
        Try {
            <#
            If ($PowerShellVersion -ge 6) { $obj = (Invoke-CimMethod -Namespace "ROOT\ccm" -ClassName SMS_Client -MethodName GetAssignedSite).sSiteCode }
            Else { $obj = $([WmiClass]"ROOT\ccm:SMS_Client").getassignedsite() | Select-Object -Expandproperty sSiteCode }
            #>
            $sms = New-Object -ComObject 'Microsoft.SMS.Client'
            $obj = $sms.GetAssignedSite()
        } Catch {
            $obj = '...' 
        } finally {
            Write-Output $obj 
        }
    }

    Function Get-ClientVersion {
        Try {
            If ($PowerShellVersion -ge 6) {
                $obj = (Get-CimInstance -Namespace root/ccm SMS_Client).ClientVersion 
            } Else {
                $obj = (Get-WmiObject -Namespace root/ccm SMS_Client).ClientVersion 
            }
        } Catch {
            $obj = $false 
        } finally {
            Write-Output $obj 
        }
    }

    Function Get-ClientCache {
        Try {
            $obj = (New-Object -ComObject UIResource.UIResourceMgr).GetCacheInfo().TotalSize
            #If ($PowerShellVersion -ge 6) { $obj = (Get-CimInstance -Namespace "ROOT\CCM\SoftMgmtAgent" -Class CacheConfig -ErrorAction SilentlyContinue).Size }
            #Else { $obj = (Get-WmiObject -Namespace "ROOT\CCM\SoftMgmtAgent" -Class CacheConfig -ErrorAction SilentlyContinue).Size }
        } Catch {
            $obj = 0 
        } finally {
            If ($null -eq $obj) {
                $obj = 0 
            }
            Write-Output $obj
        }
    }

    Function Get-ClientMaxLogSize {
        Try {
            $obj = [Math]::Round(((Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global').LogMaxSize) / 1000) 
        } Catch {
            $obj = 0 
        } finally {
            Write-Output $obj 
        }
    }


    Function Get-ClientMaxLogHistory {
        Try {
            $obj = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global').LogMaxHistory 
        } Catch {
            $obj = 0 
        } finally {
            Write-Output $obj 
        }
    }


    Function Get-Domain {
        Try {
            If ($PowerShellVersion -ge 6) {
                $obj = (Get-CimInstance Win32_ComputerSystem).Domain 
            } Else {
                $obj = (Get-WmiObject Win32_ComputerSystem).Domain 
            }
        } Catch {
            $obj = $false 
        } finally {
            Write-Output $obj 
        }
    }

    Function Get-CCMLogDirectory {
        $obj = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global').LogDirectory
        If ($null -eq $obj) {
            $obj = "$env:SystemDrive\windows\ccm\Logs" 
        }
        Write-Output $obj
    }

    Function Get-CCMDirectory {
        $logdir = Get-CCMLogDirectory
        $obj = $logdir.replace("\Logs", "")
        Write-Output $obj
    }

    <#
    .SYNOPSIS
    Function to test If local database files are missing from the ConfigMgr client.

    .DESCRIPTION
    Function to test If local database files are missing from the ConfigMgr client. Will tag client for reinstall If less than 7. Returns $True If compliant or $False If non-compliant

    .EXAMPLE
    An example

    .NOTES
    Returns $True If compliant or $False If non-compliant. Non.compliant computers require remediation and will be tagged for ConfigMgr client reinstall.
    #>
    Function Test-CcmSDF {
        $ccmdir = Get-CCMDirectory
        $files = @(Get-ChildItem "$ccmdir\*.sdf" -ErrorAction SilentlyContinue)
        If ($files.Count -lt 7) {
            $obj = $false 
        } Else {
            $obj = $true 
        }
        Write-Output $obj
    }

    Function Test-CcmSQLCELog {
        $logdir = Get-CCMLogDirectory
        $ccmdir = Get-CCMDirectory
        $logFile = "$logdir\CcmSQLCE.log"
        $logLevel = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global').logLevel

        If ( (Test-Path -Path $logFile) -and ($logLevel -ne 0) ) {
            # Not in debug mode, and CcmSQLCE.log exists. This could be bad.
            $LastWriteTime = (Get-ChildItem $logFile).LastWriteTime
            $CreationTime = (Get-ChildItem $logFile).CreationTime
            $FileDate = Get-Date($LastWriteTime)
            $FileCreated = Get-Date($CreationTime)

            $now = Get-Date
            If ( (($now - $FileDate).Days -lt 7) -and ((($now - $FileCreated).Days) -gt 7) ) {
                $Text = "CM client not in debug mode, and CcmSQLCE.log exists. This is very bad. Cleaning up local SDF files and reinstalling CM client"
                Write-Host $Text -ForegroundColor Red
                # Delete *.SDF Files
                $Service = Get-Service -Name ccmexec
                $Service.Stop()

                $seconds = 0
                Do {
                    Start-Sleep -Seconds 1
                    $seconds++
                } While ( ($Service.Status -ne "Stopped") -and ($seconds -le 60) )

                # Do another test to make sure CcmExec service really is stopped
                If ($Service.Status -ne "Stopped") {
                    Stop-Service -Name ccmexec -Force 
                }

                Write-Verbose "Waiting 10 seconds to allow file locking issues to clear up"
                Start-Sleep -Seconds 10

                Try {
                    $files = Get-ChildItem "$ccmdir\*.sdf"
                    $files | Remove-Item -Force -ErrorAction Stop
                    Remove-Item -Path $logFile -Force -ErrorAction Stop
                } Catch {
                    Write-Verbose "Obviously that wasn't enough time"
                    Start-Sleep -Seconds 30
                    # We Try again
                    $files = Get-ChildItem "$ccmdir\*.sdf"
                    $files | Remove-Item -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path $logFile -Force -ErrorAction SilentlyContinue
                }

                $obj = $true
            }

            # CcmSQLCE.log has not been updated for two days. We are good for now.
            Else {
                $obj = $false 
            }
        }

        # we are good
        Else {
            $obj = $false 
        }
        Write-Output $obj

    }

    function Test-CCMCertIficateError {
        Param([Parameter(Mandatory = $true)]$Log)
        # More checks to come
        $logdir = Get-CCMLogDirectory
        $logFile1 = "$logdir\ClientIDManagerStartup.log"
        $error1 = 'Failed to find the certIficate in the store'
        $error2 = '[RegTask] - Server rejected registration 3'
        $content = Get-Content -Path $logFile1

        $ok = $true

        If ($content -match $error1) {
            $ok = $false
            $Text = 'ConfigMgr Client CertIficate: Error failed to find the certIficate in store. Attempting fix.'
            Write-Warning $Text
            Stop-Service -Name ccmexec -Force
            # Name is persistant across systems.
            $cert = "$env:ProgramData\Microsoft\Crypto\RSA\MachineKeys\19c5cf9c7b5dc9de3e548adb70398402_50e417e0-e461-474b-96e2-077b80325612"
            # CCM creates new certIficate when missing.
            Remove-Item -Path $cert -Force -ErrorAction SilentlyContinue | Out-Null
            # Remove the error from the logfile to avoid double remediations based on false positives
            $newContent = $content | Select-String -Pattern $Error1 -NotMatch
            Out-File -FilePath $logfile -InputObject $newContent -Encoding utf8 -Force
            Start-Service -Name ccmexec

            # Update log object
            $log.ClientCertIficate = $error1
        }

        #$content = Get-Content -Path $logFile2
        If ($content -match $error2) {
            $ok = $false
            $Text = 'ConfigMgr Client CertIficate: Error! Server rejected client registration. Client CertIficate not valid. No auto-remediation.'
            Write-Error $Text
            $log.ClientCertIficate = $error2
        }

        If ($ok -eq $true) {
            $Text = 'ConfigMgr Client CertIficate: OK'
            Write-Output $Text
            $log.ClientCertIficate = 'OK'
        }
    }

    Function Test-InTaskSequence {
        Try {
            $tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment 
        } Catch {
            $tsenv = $null 
        }

        If ($tsenv) {
            Write-Host "Configuration Manager Task Sequence detected on computer. Exiting script"
            Exit 2
        }
    }


    Function Test-BITS {
        Param(
            [Parameter(Mandatory = $true)]$Log
        )
        If ($BitsCheckEnabled -eq $true) {
            $Errors = Get-BitsTransfer -AllUsers | Where-Object {($_.JobState -like "TransientError") -or ($_.JobState -like "Transient_Error") -or ($_.JobState -like "Error")}
            If ($Errors) {
                $Fix = Get-XMLConfigBITSCheckFix
                If ($Fix -eq "True") {
                    $Text = "BITS: Error. Remediating"
                    $Errors | Remove-BitsTransfer -ErrorAction SilentlyContinue
                    & $env:WINDIR\System32\sc.exe sdset bits "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)" | Out-Null
                    $Log.BITS = 'Remediated'
                    $Obj = $true
                } Else {
                    $Text = "BITS: Error. Monitor only"
                    $Log.BITS = 'Error'
                    $Obj = $false
                }
            } Else {
                $Text = "BITS: OK"
                $Log.BITS = 'OK'
                $Obj = $false
            }
        } Else {
            $Text = "BITS: PowerShell Module BitsTransfer missing. Skipping check"
            $Log.BITS = "BitsTransfer PowerShell module missing"
            $Obj = $false
        }
        Write-Host $Text
        Write-Output $Obj
    }

    Function Test-ClientSettingsConfiguration {
        Param([Parameter(Mandatory = $true)]$Log)

        $ClientSettingsConfig = @(Get-WmiObject -Namespace "root\ccm\Policy\DefaultMachine\RequestedConfig" -Class CCM_ClientAgentConfig -ErrorAction SilentlyContinue | Where-Object { $_.PolicySource -eq "CcmTaskSequence" })

        If ($ClientSettingsConfig.Count -gt 0) {

            $fix = (Get-XMLConfigClientSettingsCheckFix).ToLower()

            If ($fix -eq "true") {
                $Text = "ClientSettings: Error. Remediating"
                Do {
                    Get-WmiObject -Namespace "root\ccm\Policy\DefaultMachine\RequestedConfig" -Class CCM_ClientAgentConfig | Where-Object { $_.PolicySource -eq "CcmTaskSequence" } | Select-Object -First 1000 | ForEach-Object { Remove-WmiObject -InputObject $_ }
                } Until (!(Get-WmiObject -Namespace "root\ccm\Policy\DefaultMachine\RequestedConfig" -Class CCM_ClientAgentConfig | Where-Object { $_.PolicySource -eq "CcmTaskSequence" } | Select-Object -First 1))
                $log.ClientSettings = 'Remediated'
                $obj = $true
            } Else {
                $Text = "ClientSettings: Error. Monitor only"
                $log.ClientSettings = 'Error'
                $obj = $false
            }
        }

        Else {
            $Text = "ClientSettings: OK"
            $log.ClientSettings = 'OK'
            $Obj = $false
        }
        Write-Host $Text
        #Write-Output $Obj
    }

    Function New-ClientInstalledReason {
        Param(
            [Parameter(Mandatory = $true)]$Message,
            [Parameter(Mandatory = $true)]$Log
        )

        If ($null -eq $log.ClientInstalledReason) {
            $log.ClientInstalledReason = $Message 
        } Else {
            $log.ClientInstalledReason += " $Message" 
        }
    }


    function Get-PendingReboot {
        $result = @{
            CBSRebootPending            = $false
            WindowsUpdateRebootRequired = $false
            FileRenamePending           = $false
            SCCMRebootPending           = $false
        }

        #Check CBS Registry
        $key = Get-ChildItem "HKLM:Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
        If ($null -ne $key) {
            $result.CBSRebootPending = $true 
        }

        #Check Windows Update
        $key = Get-Item 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue
        If ($null -ne $key) {
            $result.WindowsUpdateRebootRequired = $true 
        }

        #Check PendingFileRenameOperations
        $prop = Get-ItemProperty 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        If ($null -ne $prop) {
            #PendingFileRenameOperations is not *must* to reboot?
            #$result.FileRenamePending = $true
        }

        Try {
            $util = [wmiclass]'\\.\root\ccm\clientsdk:CCM_ClientUtilities'
            $status = $util.DetermineIfRebootPending()
            If (($null -ne $status) -and $status.RebootPending) {
                $result.SCCMRebootPending = $true 
            }
        } Catch {
        }

        #Return Reboot required
        If ($result.ContainsValue($true)) {
            #$Text = 'Pending Reboot: YES'
            $obj = $true
            $log.PendingReboot = 'Pending Reboot'
        } Else {
            $obj = $false
            $log.PendingReboot = 'OK'
        }
        Write-Output $obj
    }

    Function Get-ProvisioningMode {
        $RegistryPath = 'HKLM:\SOFTWARE\Microsoft\CCM\CcmExec'
        $provisioningMode = (Get-ItemProperty -Path $RegistryPath).ProvisioningMode
        If ($provisioningMode -eq 'true') {
            $obj = $true 
        } Else {
            $obj = $false 
        }
        Write-Output $obj
    }

    Function Get-OSDiskFreeSpace {

        If ($PowerShellVersion -ge 6) {
            $driveC = Get-CimInstance -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object FreeSpace, Size 
        } Else {
            $driveC = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object FreeSpace, Size 
        }
        $freeSpace = (($driveC.FreeSpace / $driveC.Size) * 100)
        Write-Output ([math]::Round($freeSpace, 2))
    }

    Function Get-Computername {
        If ($PowerShellVersion -ge 6) {
            $obj = (Get-CimInstance Win32_ComputerSystem).Name 
        } Else {
            $obj = (Get-WmiObject Win32_ComputerSystem).Name 
        }
        Write-Output $obj
    }

    Function Get-LastBootTime {
        If ($PowerShellVersion -ge 6) {
            $wmi = Get-CimInstance Win32_OperatingSystem 
        } Else {
            $wmi = Get-WmiObject Win32_OperatingSystem 
        }
        $obj = $wmi.ConvertToDateTime($wmi.LastBootUpTime)
        Write-Output $obj
    }

    Function Get-LastInstalledPatches {
        Param([Parameter(Mandatory = $true)]$Log)
        # Reading date from Windows Update COM object.
        $Session = New-Object -ComObject Microsoft.Update.Session
        $Searcher = $Session.CreateUpdateSearcher()
        $HistoryCount = $Searcher.GetTotalHistoryCount()

        $OS = Get-OperatingSystem
        Switch -Wildcard ($OS) {
            "*Windows 7*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'AutomaticUpdates' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
            "*Windows 8*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'AutomaticUpdatesWuApp' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
            "*Windows 10*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'UpdateOrchestrator' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
            "*Server 2008*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'AutomaticUpdates' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
            "*Server 2012*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'AutomaticUpdatesWuApp' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
            "*Server 2016*" {
                $Date = $Searcher.QueryHistory(0, $HistoryCount) | Where-Object {
                    ($_.ClientApplicationID -eq 'UpdateOrchestrator' -or $_.ClientApplicationID -eq 'ccmexec') -and ($_.Title -notmatch "Security Intelligence Update|Definition Update")
                } | Select-Object -ExpandProperty Date | Measure-Latest
            }
        }

        # Reading date from PowerShell Get-Hotfix
        #$now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        #$Hotfix = Get-Hotfix | Where-Object {$_.InstalledOn -le $now} | Select-Object -ExpandProperty InstalledOn -ErrorAction SilentlyContinue

        #$Hotfix = Get-Hotfix | Select-Object -ExpandProperty InstalledOn -ErrorAction SilentlyContinue

        If ($PowerShellVersion -ge 6) {
            $Hotfix = Get-CimInstance -ClassName Win32_QuickFixEngineering | Select-Object @{Name = "InstalledOn"; Expression = { [DateTime]::Parse($_.InstalledOn, $([System.Globalization.CultureInfo]::GetCultureInfo("en-US"))) } } 
        } Else {
            $Hotfix = Get-HotFix | Select-Object @{l = "InstalledOn"; e = { [DateTime]::Parse($_.psbase.properties["installedon"].value, $([System.Globalization.CultureInfo]::GetCultureInfo("en-US"))) } } 
        }

        $Hotfix = $Hotfix | Select-Object -ExpandProperty InstalledOn

        $Date2 = $null

        If ($null -ne $hotfix) {
            $Date2 = Get-Date($hotfix | Measure-Latest) -ErrorAction SilentlyContinue 
        }

        If (($Date -ge $Date2) -and ($null -ne $Date)) {
            $Log.OSUpdates = Get-SmallDateTime -Date $Date 
        } Elseif (($Date2 -gt $Date) -and ($null -ne $Date2)) {
            $Log.OSUpdates = Get-SmallDateTime -Date $Date2 
        }
    }

    function Measure-Latest {
        BEGIN {
            $latest = $null 
        }
        PROCESS {
            If (($null -ne $_) -and (($null -eq $latest) -or ($_ -gt $latest))) {
                $latest = $_ 
            } 
        }
        END {
            $latest 
        }
    }

    Function Test-LogFileHistory {
        Param([Parameter(Mandatory = $true)]$Logfile)
        $startString = '<--- ConfigMgr Client Health Check starting --->'
        $content = ''

        # Handle the network share log file
        If (Test-Path $logfile -ErrorAction SilentlyContinue) {
            $content = Get-Content $logfile -ErrorAction SilentlyContinue 
        } Else {
            Return 
        }
        $maxHistory = Get-XMLConfigLoggingMaxHistory
        $startCount = [regex]::matches($content, $startString).count

        # Delete logfile If more start and stop entries than max history
        If ($startCount -ge $maxHistory) {
            Remove-Item $logfile -Force 
        }
    }

    Function Test-DNSConfiguration {
        Param([Parameter(Mandatory = $true)]$Log)
        #$dnsdomain = (Get-WmiObject Win32_NetworkAdapterConfiguration -filter "ipenabled = 'true'").DNSDomain
        $fqdn = [System.Net.Dns]::GetHostEnTry([string]"localhost").HostName
        If ($PowerShellVersion -ge 6) {
            $localIPs = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -Match "True" } |  Select-Object -ExpandProperty IPAddress 
        } Else {
            $localIPs = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -Match "True" } |  Select-Object -ExpandProperty IPAddress 
        }
        $dnscheck = [System.Net.DNS]::GetHostByName($fqdn)

        $OSName = Get-OperatingSystem
        If (($OSName -notlike "*Windows 7*") -and ($OSName -notlike "*Server 2008*")) {
            # This method is supported on Windows 8 / Server 2012 and higher. More acurate than using .NET object method
            Try {
                $ActiveAdapters = (Get-NetAdapter | Where-Object { $_.Status -like "Up" }).Name
                $dnsServers = Get-DnsClientServerAddress | Where-Object { $ActiveAdapters -contains $_.InterfaceAlias } | Where-Object { $_.AddressFamily -eq 2 } | Select-Object -ExpandProperty ServerAddresses
                $dnsAddressList = Resolve-DnsName -Name $fqdn -Server ($dnsServers | Select-Object -First 1) -Type A -DnsOnly | Select-Object -ExpandProperty IPAddress
            } Catch {
                # Fallback to depreciated method
                $dnsAddressList = $dnscheck.AddressList | Select-Object -ExpandProperty IPAddressToString
                $dnsAddressList = $dnsAddressList -replace ("%(.*)", "")
            }
        }

        Else {
            # This method cannot guarantee to only resolve against DNS sever. Local cache can be used in some circumstances.
            # For Windows 7 only

            $dnsAddressList = $dnscheck.AddressList | Select-Object -ExpandProperty IPAddressToString
            $dnsAddressList = $dnsAddressList -replace ("%(.*)", "")
        }

        $dnsFail = ''
        $logFail = ''

        Write-Verbose 'VerIfy that local machines FQDN matches DNS'
        If ($dnscheck.HostName -like $fqdn) {
            $obj = $true
            Write-Verbose 'Checking If one local IP matches on IP from DNS'
            Write-Verbose 'Loop through each IP address published in DNS'
            foreach ($dnsIP in $dnsAddressList) {
                #Write-Host "Testing If IP address: $dnsIP published in DNS exist in local IP configuration."
                ##If ($dnsIP -notin $localIPs) { ## Requires PowerShell 3. Works fine :(
                If ($localIPs -notcontains $dnsIP) {
                    $dnsFail += "IP '$dnsIP' in DNS record do not exist locally`n"
                    $logFail += "$dnsIP "
                    $obj = $false
                }
            }
        } Else {
            $hn = $dnscheck.HostName
            $dnsFail = 'DNS name: ' + $hn + ' local fqdn: ' + $fqdn + ' DNS IPs: ' + $dnsAddressList + ' Local IPs: ' + $localIPs
            $obj = $false
            Write-Host $dnsFail
        }

        $FileLogLevel = ((Get-XMLConfigLoggingLevel).ToString()).ToLower()

        switch ($obj) {
            $false {
                $fix = (Get-XMLConfigDNSFix).ToLower()
                If ($fix -eq "true") {
                    $Text = 'DNS Check: FAILED. IP address published in DNS do not match IP address on local machine. Trying to resolve by registerting with DNS server'
                    If ($PowerShellVersion -ge 4) {
                        Register-DnsClient | Out-Null 
                    } Else {
                        ipconfig /registerdns | Out-Null 
                    }
                    Write-Host $Text
                    $log.DNS = $logFail
                    If (-NOT($FileLogLevel -like "clientlocal")) {
                        Out-LogFile -Xml $xml -Text $Text -Severity 2
                        Out-LogFile -Xml $xml -Text $dnsFail -Severity 2
                    }

                } Else {
                    $Text = 'DNS Check: FAILED. IP address published in DNS do not match IP address on local machine. Monitor mode only, no remediation'
                    $log.DNS = $logFail
                    If (-NOT($FileLogLevel -like "clientlocal")) {
                        Out-LogFile -Xml $xml -Text $Text -Severity 2 
                    }
                    Write-Host $Text
                }

            }
            $true {
                $Text = 'DNS Check: OK'
                Write-Output $Text
                $log.DNS = 'OK'
            }
        }
        #Write-Output $obj
    }

    # Function to test that 'HKU:\S-1-5-18\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\' is set to '%USERPROFILE%\AppData\Roaming'. CCMSETUP will fail If not.
    # Reference: https://www.systemcenterdudes.com/could-not-access-network-location-appdata-ccmsetup-log/
    Function Test-CCMSetup1 {
        New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null
        $correctValue = '%USERPROFILE%\AppData\Roaming'
        $currentValue = (Get-Item 'HKU:\S-1-5-18\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\').GetValue('AppData', $null, 'DoNotExpandEnvironmentNames')

        # Only fix If the value is wrong
        If ($currentValue -ne $correctValue) {
            Set-ItemProperty -Path 'HKU:\S-1-5-18\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\' -Name 'AppData' -Value $correctValue 
        }
    }

    Function Test-Update {
        Param([Parameter(Mandatory = $true)]$Log)

        #If (($Xml.Configuration.Option | Where-Object {$_.Name -like 'Updates'} | Select-Object -ExpandProperty 'Enable') -like 'True') {

        $UpdateShare = Get-XMLConfigUpdatesShare
        #$UpdateShare = $Xml.Configuration.Option | Where-Object {$_.Name -like 'Updates'} | Select-Object -ExpandProperty 'Share'


        Write-Verbose "Validating required updates is installed on the client. Required updates will be installed If missing on client."
        #$OS = Get-WmiObject -class Win32_OperatingSystem
        $OSName = Get-OperatingSystem


        $build = $null
        If ($OSName -like "*Windows 10*") {
            $build = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber
            switch ($build) {
                10240 {
                    $OSName = $OSName + " 1507" 
                }
                10586 {
                    $OSName = $OSName + " 1511" 
                }
                14393 {
                    $OSName = $OSName + " 1607" 
                }
                15063 {
                    $OSName = $OSName + " 1703" 
                }
                16299 {
                    $OSName = $OSName + " 1709" 
                }
                17134 {
                    $OSName = $OSName + " 1803" 
                }
                17763 {
                    $OSName = $OSName + " 1809" 
                }
                default {
                    $OSName = $OSName + " Insider Preview" 
                }
            }
        }

        $Updates = (Join-Path $UpdateShare $OSName)
        If ((Test-Path $Updates) -eq $true) {
            $regex = '(?i)^.+-kb[0-9]{6,}-(?:v[0-9]+-)?x[0-9]+\.msu$'
            $hotfixes = @(Get-ChildItem $Updates | Where-Object { $_.Name -match $regex } | Select-Object -ExpandProperty Name)

            If ($PowerShellVersion -ge 6) {
                $installedUpdates = @((Get-CimInstance Win32_QuickFixEngineering).HotFixID) 
            } Else {
                $installedUpdates = @(Get-HotFix | Select-Object -ExpandProperty HotFixID) 
            }

            $count = $hotfixes.count

            If (($count -eq 0) -or ($count -eq $null)) {
                $Text = 'Updates: No mandatory updates to install.'
                Write-Output $Text
                $log.Updates = 'OK'
            } Else {
                $logEnTry = $null

                $regex = '\b(?!(KB)+(\d+)\b)\w+'
                foreach ($hotfix in $hotfixes) {
                    $kb = $hotfix -replace $regex -replace "\." -replace "-"
                    If ($installedUpdates -contains $kb) {
                        $Text = "Update $hotfix" + ": OK"
                        Write-Output $Text
                    } Else {
                        If ($null -eq $logEnTry) {
                            $logEnTry = $kb 
                        } Else {
                            $logEnTry += ", $kb" 
                        }

                        $fix = (Get-XMLConfigUpdatesFix).ToLower()
                        If ($fix -eq "true") {
                            $kbfullpath = Join-Path $updates $hotfix
                            $Text = "Update $hotfix" + ": Missing. Installing now..."
                            Write-Warning $Text

                            $temppath = Join-Path (Get-LocalFilesPath) "Temp"

                            If ((Test-Path $temppath) -eq $false) {
                                New-Item -Path $temppath -ItemType Directory | Out-Null 
                            }

                            Copy-Item -Path $kbfullpath -Destination $temppath
                            $install = Join-Path $temppath $hotfix

                            wusa.exe $install /quiet /norestart
                            While (Get-Process wusa -ErrorAction SilentlyContinue) {
                                Start-Sleep -Seconds 2 
                            }
                            Remove-Item $install -Force -Recurse

                        } Else {
                            $Text = "Update $hotfix" + ": Missing. Monitor mode only, no remediation."
                            Write-Warning $Text
                        }
                    }

                    If ($null -eq $logEnTry) {
                        $log.Updates = 'OK' 
                    } Else {
                        $log.Updates = $logEnTry 
                    }
                }
            }
        } Else {
            $log.Updates = 'Failed'
            Write-Warning "Updates Failed: Could not locate update folder '$($Updates)'."
        }
    }

    Function Test-ConfigMgrClient {
        Param(
            [Parameter(Mandatory = $true)]$Log
        )
        # Check if the ConfigMgr client is installed or not and if installed, perform tests to decide if reinstall is needed or not.
        If (Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue) {
            $Text = "Configuration Manager Client is installed"
            Write-Host $Text
            # Lets assume we don't need to reinstall the client unless tests tells us to.
            $Reinstall = $false
            # We test that the local database files exists. Less than 7 means the client is horrible broken and requires reinstall.
            $LocalDBFilesPresent = Test-CcmSDF
            If ($LocalDBFilesPresent -eq $false) {
                New-ClientInstalledReason -Log $Log -Message "ConfigMgr Client database files missing."
                Write-Host "ConfigMgr Client database files missing. Reinstalling..."
                $Reinstall = $true
                $Uninstall = $true
            }
            # Only test ConfigMgr local DB if this check is enabled
            $TestLocalDB = Get-XMLConfigCcmSQLCELog
            If ($TestLocalDB -eq "True") {
                Write-Host "Testing CcmSQLCELog"
                $LocalDB = Test-CcmSQLCELog
                If ($LocalDB -eq $true) {
                    # LocalDB is messed up
                    New-ClientInstalledReason -Log $Log -Message "ConfigMgr Client database corrupt."
                    Write-Host "ConfigMgr Client database corrupt. Reinstalling..."
                    $Reinstall = $true
                    $Uninstall = $true
                }
            }
            $CCMService = Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue
            # Reinstall if we are unable to start the ConfigMgr client
            If (($CCMService.Status -eq "Stopped") -and ($LocalDB -eq $false)) {
                Try {
                    Write-Host "ConfigMgr Agent not running. Attempting to start it."
                    If ($CCMService.StartType -ne "Automatic") {
                        $Text = "Configuring service CcmExec StartupType to: Automatic (Delayed Start)..."
                        Write-Output $Text
                        Set-Service -Name "CcmExec" -StartupType Automatic
                    }
                    Start-Service -Name "CcmExec"
                } Catch {
                    $Reinstall = $true
                    New-ClientInstalledReason -Log $Log -Message "Service not running, failed to start."
                }
            }
            # Test that we are able to connect to SMS_Client WMI class
            Try {
                Get-CimInstance -Namespace "Root/Ccm" -Class "SMS_Client" -ErrorAction Stop
            } Catch {
                Write-Verbose 'Failed to connect to WMI namespace "Root/Ccm" class "SMS_Client". Repairing WMI and tagging client for reinstall to fix.'
                Repair-WMI
                $Reinstall = $true
                New-ClientInstalledReason -Log $Log -Message "Failed to connect to SMS_Client WMI class."
            }
            # Check if client is installed, but failing to register by checking CCMMessaging.log for 'ccm_system_windowsauth/request cannot be fulfilled since use of metered network is not allowed'
            $LogDir = Get-CCMLogDirectory
            $LogFile = "$LogDir\CCMMessaging.log"
            $LogLine = Search-CMLogFile -LogFile $LogFile -SearchStrings @('ccm_system_windowsauth/request cannot be fulfilled since use of metered network is not allowed')
            If ($LogLine) {
                $AdaptersGuids = Get-NetAdapter | Select-Object -ExpandProperty InterfaceGuid
                Foreach ($Guid in $AdaptersGuids) {
                    New-Item -Path "HKLM:\SOFTWARE\Microsoft\DusmSvc\Profiles\$Guid\*" -Force | Out-Null
                    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\DusmSvc\Profiles\$Guid\*" -Name UserCost -Value 0 -Type DWord -Force | Out-Null
                }
                # Restart the Data Usage service to commit the changes to network cost
                Restart-Service -Name DusmSvc -Force
                # To avoid false positives, remove CCMMessaging.log
                Remove-Item -Path $LogFile -Force -Confirm:$false | Out-Null
            }
            # Check if any of the above steps tagged for a repair
            If ($Reinstall -eq $true) {
                If ($Uninstall -eq $true) {
                    $Text = "ConfigMgrClientHealth thinks the agent need to be uninstalled and reinstalled.."
                } Else {
                    $Text = "ConfigMgrClientHealth thinks the agent need to be reinstalled.."
                }
                Write-Host $Text
                # Lets check that Registry settings are OK before we Try a new installation.
                Test-CCMSetupRegValue
                # Reinstall the client
                Resolve-Client -Xml $Xml -ClientInstallProperties $ClientInstallProperties -FirstInstall $false
                $Log.ClientInstalled = Get-SmallDateTime
                Start-Sleep 600
            }
        } Else {
            $Text = "Configuration Manager client is not installed. Installing..."
            Write-Host $Text
            Resolve-Client -Xml $Xml -ClientInstallProperties $ClientInstallProperties -FirstInstall $true
            New-ClientInstalledReason -Log $Log -Message "No client found."
            $Log.ClientInstalled = Get-SmallDateTime
            Start-Sleep 600
            # Test again if client is installed
            If (Get-Service -Name "CcmExec" -ErrorAction SilentlyContinue) {
                # Service now found, so client is installed
            } Else {
                Out-LogFile "ConfigMgr Client installation failed. Agent not detected 10 minutes after triggering installation."  -Mode "ClientInstall" -Severity 3
            }
        }
    }

    Function Test-ClientCacheSize {
        Param([Parameter(Mandatory = $true)]$Log)
        $ClientCacheSize = Get-XMLConfigClientCache
        #If ($PowerShellVersion -ge 6) { $Cache = Get-CimInstance -Namespace "ROOT\CCM\SoftMgmtAgent" -Class CacheConfig }
        #Else { $Cache = Get-WmiObject -Namespace "ROOT\CCM\SoftMgmtAgent" -Class CacheConfig }

        $CurrentCache = Get-ClientCache

        If ($ClientCacheSize -match '%') {
            $type = 'percentage'
            # percentage based cache based on disk space
            $num = $ClientCacheSize -replace '%'
            $num = ($num / 100)
            # TotalDiskSpace in Byte
            If ($PowerShellVersion -ge 6) {
                $TotalDiskSpace = (Get-CimInstance -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object -ExpandProperty Size) 
            } Else {
                $TotalDiskSpace = (Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object -ExpandProperty Size) 
            }
            $ClientCacheSize = ([math]::Round(($TotalDiskSpace * $num) / 1048576))
        } Else {
            $type = 'fixed' 
        }

        If ($CurrentCache -eq $ClientCacheSize) {
            $Text = "ConfigMgr Client Cache Size: OK"
            Write-Host $Text
            $Log.CacheSize = $CurrentCache
            $obj = $false
        }

        Else {
            switch ($type) {
                'fixed' {
                    $Text = "ConfigMgr Client Cache Size: $CurrentCache. Expected: $ClientCacheSize. Redmediating." 
                }
                'percentage' {
                    $percent = Get-XMLConfigClientCache
                    If ($ClientCacheSize -gt "99999") {
                        $ClientCacheSize = "99999" 
                    }
                    $Text = "ConfigMgr Client Cache Size: $CurrentCache. Expected: $ClientCacheSize ($percent). (99999 maxium). Redmediating."
                }
            }

            Write-Warning $Text
            #$Cache.Size = $ClientCacheSize
            #$Cache.Put()
            $log.CacheSize = $ClientCacheSize
            (New-Object -ComObject UIResource.UIResourceMgr).GetCacheInfo().TotalSize = "$ClientCacheSize"
            $obj = $true
        }
        Write-Output $obj
    }

    Function Test-ClientVersion {
        Param([Parameter(Mandatory = $true)]$Log)
        $ClientVersion = Get-XMLConfigClientVersion
        [String]$ClientAutoUpgrade = Get-XMLConfigClientAutoUpgrade
        $ClientAutoUpgrade = $ClientAutoUpgrade.ToLower()
        $installedVersion = Get-ClientVersion
        $log.ClientVersion = $installedVersion

        If ($installedVersion -ge $ClientVersion) {
            $Text = 'ConfigMgr Client version is: ' + $installedVersion + ': OK'
            Write-Output $Text
            $obj = $false
        } Elseif ($ClientAutoUpgrade -like 'true') {
            $Text = 'ConfigMgr Client version is: ' + $installedVersion + ': Tagging client for upgrade to version: ' + $ClientVersion
            Write-Warning $Text
            $obj = $true
        } Else {
            $Text = 'ConfigMgr Client version is: ' + $installedVersion + ': Required version: ' + $ClientVersion + ' AutoUpgrade: false. Skipping upgrade'
            Write-Output $Text
            $obj = $false
        }
        Write-Output $obj
    }

    Function Test-ClientSiteCode {
        Param([Parameter(Mandatory = $true)]$Log)
        $sms = New-Object -ComObject "Microsoft.SMS.Client"
        $ClientSiteCode = Get-XMLConfigClientSitecode
        #[String]$currentSiteCode = Get-Sitecode
        $currentSiteCode = $sms.GetAssignedSite()
        $currentSiteCode = $currentSiteCode.Trim()
        $Log.Sitecode = $currentSiteCode

        # Do more investigation and testing on WMI Method "SetAssignedSite" to possible avoid reinstall of client for this check.
        If ($ClientSiteCode -like $currentSiteCode) {
            $Text = "ConfigMgr Client Site Code: OK"
            Write-Host $Text
            #$obj = $false
        } Else {
            $Text = 'ConfigMgr Client Site Code is "' + $currentSiteCode + '". Expected: "' + $ClientSiteCode + '". Changing sitecode.'
            Write-Warning $Text
            $sms.SetAssignedSite($ClientSiteCode)
            #$obj = $true
        }
        #Write-Output $obj
    }

    function Test-PendingReboot {
        Param([Parameter(Mandatory = $true)]$Log)
        # Only run pending reboot check If enabled in config
        If (($Xml.Configuration.Option | Where-Object { $_.Name -like 'PendingReboot' } | Select-Object -ExpandProperty 'Enable') -like 'True') {
            $result = @{
                CBSRebootPending            = $false
                WindowsUpdateRebootRequired = $false
                FileRenamePending           = $false
                SCCMRebootPending           = $false
            }

            #Check CBS Registry
            $key = Get-ChildItem "HKLM:Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
            If ($null -ne $key) {
                $result.CBSRebootPending = $true 
            }

            #Check Windows Update
            $key = Get-Item 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ErrorAction SilentlyContinue
            If ($null -ne $key) {
                $result.WindowsUpdateRebootRequired = $true 
            }

            #Check PendingFileRenameOperations
            $prop = Get-ItemProperty 'HKLM:SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
            If ($null -ne $prop) {
                #PendingFileRenameOperations is not *must* to reboot?
                #$result.FileRenamePending = $true
            }

            Try {
                $util = [wmiclass]'\\.\root\ccm\clientsdk:CCM_ClientUtilities'
                $status = $util.DetermineIfRebootPending()
                If (($null -ne $status) -and $status.RebootPending) {
                    $result.SCCMRebootPending = $true 
                }
            } Catch {
            }

            #Return Reboot required
            If ($result.ContainsValue($true)) {
                $Text = 'Pending Reboot: Computer is in pending reboot'
                Write-Warning $Text
                $log.PendingReboot = 'Pending Reboot'

                If ((Get-XMLConfigPendingRebootApp) -eq $true) {
                    Start-RebootApplication
                    $log.RebootApp = Get-SmallDateTime
                }
            } Else {
                $Text = 'Pending Reboot: OK'
                Write-Output $Text
                $log.PendingReboot = 'OK'
            }
            #Out-LogFile -Xml $xml -Text $Text
        }
    }

    # Functions to detect and fix errors
    Function Test-ProvisioningMode {
        Param([Parameter(Mandatory = $true)]$Log)
        $RegistryPath = 'HKLM:\SOFTWARE\Microsoft\CCM\CcmExec'
        $provisioningMode = (Get-ItemProperty -Path $RegistryPath).ProvisioningMode

        If ($provisioningMode -eq 'true') {
            $Text = 'ConfigMgr Client Provisioning Mode: YES. Remediating...'
            Write-Warning $Text
            Set-ItemProperty -Path $RegistryPath -Name ProvisioningMode -Value "false"
            $ArgumentList = @($false)
            If ($PowerShellVersion -ge 6) {
                Invoke-CimMethod -Namespace 'root\ccm' -Class 'SMS_Client' -MethodName 'SetClientProvisioningMode' -Arguments @{bEnable = $false } | Out-Null 
            } Else {
                Invoke-WmiMethod -Namespace 'root\ccm' -Class 'SMS_Client' -Name 'SetClientProvisioningMode' -ArgumentList $ArgumentList | Out-Null 
            }
            $log.ProvisioningMode = 'Repaired'
        } Else {
            $Text = 'ConfigMgr Client Provisioning Mode: OK'
            Write-Output $Text
            $log.ProvisioningMode = 'OK'
        }
    }

    Function Update-State {
        Write-Verbose "Start Update-State"
        $SCCMUpdatesStore = New-Object -ComObject Microsoft.CCM.UpdatesStore
        $SCCMUpdatesStore.RefreshServerComplianceState()
        $log.StateMessages = 'OK'
        Write-Verbose "End Update-State"
    }

    Function Test-UpdateStore {
        Param([Parameter(Mandatory = $true)]$Log)
        Write-Verbose "Check StateMessage.log If State Messages are successfully forwarded to Management Point"
        $logdir = Get-CCMLogDirectory
        $logfile = "$logdir\StateMessage.log"
        $StateMessage = Get-Content($logfile)
        If ($StateMessage -match 'Successfully forwarded State Messages to the MP') {
            $Text = 'StateMessage: OK'
            $log.StateMessages = 'OK'
            Write-Output $Text
        } Else {
            $Text = 'StateMessage: ERROR. Remediating...'
            Write-Warning $Text
            Update-State
            $log.StateMessages = 'Repaired'
        }
    }

    Function Test-RegistryPol {
        Param(
            [datetime]$StartTime = [datetime]::MinValue,
            $Days,
            [Parameter(Mandatory = $true)]$Log)
        $log.WUAHandler = "Checking"
        $RepairReason = ""
        $MachineRegistryFile = "$($env:WinDir)\System32\GroupPolicy\Machine\Registry.pol"

        # Check 1 - Error in WUAHandler.log
        Write-Verbose "Check WUAHandler.log for errors since $($StartTime)."
        $logdir = Get-CCMLogDirectory
        $logfile = "$logdir\WUAHandler.log"
        $logLine = Search-CMLogFile -LogFile $logfile -StartTime $StartTime -SearchStrings @('0x80004005', '0x87d00692')
        If ($logLine) {
            $RepairReason = "WUAHandler Log" 
        }

        # Check 2 - Registry.pol is too old.
        If ($Days) {
            Write-Verbose "Check machine Registry file to see If it's older than $($Days) days."
            Try {
                $file = Get-ChildItem -Path $MachineRegistryFile -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty LastWriteTime
                $regPolDate = Get-Date($file)
                $now = Get-Date
                If (($now - $regPolDate).Days -ge $Days) {
                    $RepairReason = "File Age" 
                }
            } Catch {
                Write-Warning "GPO Cache: Failed to check machine policy age." 
            }
        }

        # Check 3 - Look back through the last 7 days for group policy processing errors.
        #Event IDs documented here: https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-vista/cc749336(v=ws.10)#troubleshooting-group-policy-using-event-logs-1
        Try {
            Write-Verbose "Checking the Group Policy event log for errors since $($StartTime)."
            $numberOfGPOErrors = (Get-WinEvent -Verbose:$false -FilterHashtable @{LogName = 'Microsoft-Windows-GroupPolicy/Operational'; Level = 2; StartTime = $StartTime } -ErrorAction SilentlyContinue | Where-Object { ($_.ID -ge 7000 -and $_.ID -le 7007) -or ($_.ID -ge 7018 -and $_.ID -le 7299) -or ($_.ID -eq 1096) }).Count
            If ($numberOfGPOErrors -gt 0) {
                $RepairReason = "Event Log" 
            }

        } Catch {
            Write-Warning "GPO Cache: Failed to check the event log for policy errors." 
        }

        #If we need to repart the policy files then do so.
        If ($RepairReason -ne "") {
            $log.WUAHandler = "Broken ($RepairReason)"
            Write-Output "GPO Cache: Broken ($RepairReason)"
            Write-Verbose 'Deleting Registry.pol and running gpupdate...'

            Try {
                If (Test-Path -Path $MachineRegistryFile) {
                    Remove-Item $MachineRegistryFile -Force 
                } 
            } Catch {
                Write-Warning "GPO Cache: Failed to remove the Registry file ($($MachineRegistryFile))." 
            } finally {
                & Write-Output n | gpupdate.exe /force /target:computer | Out-Null 
            }

            #Write-Verbose 'Sleeping for 1 minute to allow for group policy to refresh'
            #Start-Sleep -Seconds 60

            Write-Verbose 'Refreshing update policy'
            Invoke-SCCMPolicyScanUpdateSource
            Invoke-SCCMPolicySourceUpdateMessage

            $log.WUAHandler = "Repaired ($RepairReason)"
            Write-Output "GPO Cache: $($log.WUAHandler)"
        } Else {
            $log.WUAHandler = 'OK'
            Write-Output "GPO Cache: OK"
        }
    }

    Function Test-ClientLogSize {
        Param([Parameter(Mandatory = $true)]$Log)
        Try {
            [int]$currentLogSize = Get-ClientMaxLogSize 
        } Catch {
            [int]$currentLogSize = 0 
        }
        Try {
            [int]$currentMaxHistory = Get-ClientMaxLogHistory 
        } Catch {
            [int]$currentMaxHistory = 0 
        }
        Try {
            $logLevel = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Logging\@Global').logLevel 
        } Catch {
            $logLevel = 1 
        }

        $clientLogSize = Get-XMLConfigClientMaxLogSize
        $clientLogMaxHistory = Get-XMLConfigClientMaxLogHistory

        $Text = ''

        If ( ($currentLogSize -eq $clientLogSize) -and ($currentMaxHistory -eq $clientLogMaxHistory) ) {
            $Log.MaxLogSize = $currentLogSize
            $Log.MaxLogHistory = $currentMaxHistory
            $Text = "ConfigMgr Client Max Log Size: OK ($currentLogSize)"
            Write-Host $Text
            $Text = "ConfigMgr Client Max Log History: OK ($currentMaxHistory)"
            Write-Host $Text
            $obj = $false
        } Else {
            If ($currentLogSize -ne $clientLogSize) {
                $Text = 'ConfigMgr Client Max Log Size: Configuring to ' + $clientLogSize + ' KB'
                $Log.MaxLogSize = $clientLogSize
                Write-Warning $Text
            } Else {
                $Text = "ConfigMgr Client Max Log Size: OK ($currentLogSize)"
                Write-Host $Text
            }
            If ($currentMaxHistory -ne $clientLogMaxHistory) {
                $Text = 'ConfigMgr Client Max Log History: Configuring to ' + $clientLogMaxHistory
                $Log.MaxLogHistory = $clientLogMaxHistory
                Write-Warning $Text
            } Else {
                $Text = "ConfigMgr Client Max Log History: OK ($currentMaxHistory)"
                Write-Host $Text
            }

            $newLogSize = [int]$clientLogSize
            $newLogSize = $newLogSize * 1000

            <#
            If ($PowerShellVersion -ge 6) {Invoke-CimMethod -Namespace "root/ccm" -ClassName "sms_client" -MethodName SetGlobalLoggingConfiguration -Arguments @{LogLevel=$loglevel; LogMaxHistory=$clientLogMaxHistory; LogMaxSize=$newLogSize} }
            Else {
                $smsClient = [wmiclass]"root/ccm:sms_client"
                $smsClient.SetGlobalLoggingConfiguration($logLevel, $newLogSize, $clientLogMaxHistory)
            }
            #Write-Verbose 'Returning true to trigger restart of ccmexec service'
            #>
            
            # Rewrote after the WMI Method stopped working in previous CM client version
            New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\CCM\Logging\@GLOBAL" -Name LogMaxHistory -PropertyType DWORD -Value $clientLogMaxHistory -Force | Out-Null
            New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\CCM\Logging\@GLOBAL" -Name LogMaxSize -PropertyType DWORD -Value $newLogSize -Force | Out-Null

            #Write-Verbose 'Sleeping for 5 seconds to allow WMI method complete before we collect new results...'
            #Start-Sleep -Seconds 5

            Try {
                $Log.MaxLogSize = Get-ClientMaxLogSize 
            } Catch {
                $Log.MaxLogSize = 0 
            }
            Try {
                $Log.MaxLogHistory = Get-ClientMaxLogHistory 
            } Catch {
                $Log.MaxLogHistory = 0 
            }
            $obj = $true
        }
        Write-Output $obj
    }

    Function Remove-CCMOrphanedCache {
        Write-Host "Clearing ConfigMgr orphaned Cache items."
        Try {
            $CCMCache = "$env:SystemDrive\Windows\ccmcache"
            $CCMCache = (New-Object -ComObject "UIResource.UIResourceMgr").GetCacheInfo().Location
            If ($null -eq $CCMCache) {
                $CCMCache = "$env:SystemDrive\Windows\ccmcache" 
            }
            $ValidCachedFolders = (New-Object -ComObject "UIResource.UIResourceMgr").GetCacheInfo().GetCacheElements() | ForEach-Object { $_.Location }
            $AllCachedFolders = (Get-ChildItem -Path $CCMCache) | Select-Object Fullname -ExpandProperty Fullname

            ForEach ($CachedFolder in $AllCachedFolders) {
                If ($ValidCachedFolders -notcontains $CachedFolder) {
                    #Don't delete new folders that might be syncing data with BITS
                    If ((Get-ItemProperty $CachedFolder).LastWriteTime -le (Get-Date).AddDays(-14)) {
                        Write-Verbose "Removing orphaned folder: $CachedFolder - LastWriteTime: $((Get-ItemProperty $CachedFolder).LastWriteTime)"
                        Remove-Item -Path $CachedFolder -Force -Recurse
                    }
                }
            }
        } Catch {
            Write-Host "Failed Clearing ConfigMgr orphaned Cache items." 
        }
    }

    Function Resolve-Client {
        Param(
            [Parameter(Mandatory = $false)]$Xml,
            [Parameter(Mandatory = $true)]$ClientInstallProperties,
            [Parameter(Mandatory = $false)]$FirstInstall = $false
        )
        $ClientShare = Get-XMLConfigClientShare
        If (Test-Path -Path $ClientShare) {
            If ($FirstInstall -eq $true) {
                $Text = 'Installing Configuration Manager Client.'
            } Else {
                $Text = 'Client tagged for reinstall. Reinstalling client...'
            }
            Write-Output $Text
            Write-Verbose "Perform a test on a specific Registry key required for ccmsetup to succeed."
            Test-CCMSetupRegValue
            Write-Verbose "Enforce registration of common DLL files to make sure CCM client works."
            $DllFiles = 'actxprxy.dll', 'atl.dll', 'Bitsprx2.dll', 'Bitsprx3.dll', 'browseui.dll', 'cryptdlg.dll', 'dssenh.dll', 'gpkcsp.dll', 'initpki.dll', 'jscript.dll', 'mshtml.dll', 'msi.dll', 'mssip32.dll', 'msxml.dll', 'msxml3.dll', 'msxml3a.dll', 'msxml3r.dll', 'msxml4.dll', 'msxml4a.dll', 'msxml4r.dll', 'msxml6.dll', 'msxml6r.dll', 'muweb.dll', 'ole32.dll', 'oleaut32.dll', 'Qmgr.dll', 'Qmgrprxy.dll', 'rsaenh.dll', 'sccbase.dll', 'scrrun.dll', 'shdocvw.dll', 'shell32.dll', 'slbcsp.dll', 'softpub.dll', 'rlmon.dll', 'userenv.dll', 'vbscript.dll', 'Winhttp.dll', 'wintrust.dll', 'wuapi.dll', 'wuaueng.dll', 'wuaueng1.dll', 'wucltui.dll', 'wucltux.dll', 'wups.dll', 'wups2.dll', 'wuweb.dll', 'wuwebv.dll', 'Xpob2res.dll', 'WBEM\wmisvc.dll'
            Foreach ($Dll in $DllFiles) {
                $File = "$env:WINDIR\System32\$Dll"
                Start-Process -FilePath "$env:WINDIR\System32\regsvr32.exe" -Args "/s $File" -Wait -NoNewWindow
            }
            # If tagged for uninstall, uninstall the client first
            If ($Uninstall -eq $true) {
                Write-Verbose "Trigger ConfigMgr Client uninstallation."
                If (Test-Path -Path "$env:WINDIR\ccmsetup\ccmsetup.exe") {
                    & $env:WINDIR\ccmsetup\ccmsetup.exe /uninstall
                } Else {
                    & $ClientShare\ccmsetup.exe /uninstall
                }
                Wait-Process "ccmsetup" -ErrorAction SilentlyContinue
            }
            # Install or reinstall the client
            Write-Verbose "Trigger ConfigMgr Client installation."
            Write-Verbose "Client install string: $ClientShare\ccmsetup.exe $ClientInstallProperties"
            & $ClientShare\ccmsetup.exe $ClientInstallProperties
            Wait-Process "ccmsetup" -ErrorAction SilentlyContinue
            # Wait for sync if this was the first time installation
            If ($FirstInstall -eq $true) {
                Write-Host "ConfigMgr Client was installed for the first time. Waiting 10 minutes for client to syncronize policy before proceeding."
                Start-Sleep -Seconds 300
            }
            # Trigger a Machine Policy Request & Evaluation Cycle after client installation
            Start-Sleep -Seconds 300
            Invoke-SCCMPolicyMachineRequest
            Invoke-SCCMPolicyMachineEvaluation
        } Else {
            $Text = 'ERROR: Client tagged for reinstall, but failed to access fileshare: ' + $ClientShare
            Write-Error $Text
            Exit 1
        }
    }

    function Register-DLLFile {
        [CmdletBinding()]
        Param ([string]$FilePath)

        Try {
            $Result = Start-Process -FilePath 'regsvr32.exe' -Args "/s `"$FilePath`"" -Wait -NoNewWindow -PassThru 
        } Catch {
        }
    }

    Function Test-WMI {
        Param(
            [Parameter(Mandatory = $true)]$Log
        )
        $Vote = 0
        $Obj = $false
        $Result = & $env:WINDIR\System32\wbem\winmgmt.exe /verifyrepository
        Switch -Wildcard ($Result) {
            # Always fix if this Returns inconsistent (multiple languages)
            "*inconsistent*"   {$Vote = 100}
            "*not consistent*" {$Vote = 100}
            "*inkonsekvent*"   {$Vote = 100}
            "*epäyhtenäinen*"  {$Vote = 100}
            "*inkonsistent*"   {$Vote = 100}
        }
        Try {
            Get-CimInstance Win32_ComputerSystem -ErrorAction Stop | Out-Null
        } Catch {
            Write-Verbose 'Failed to connect to WMI class "Win32_ComputerSystem". Voting for WMI fix...'
            $Vote++
        } Finally {
            If ($Vote -eq 0) {
                $Text = 'WMI Check: OK'
                $Log.WMI = 'OK'
                Write-Host $Text
            } Else {
                $Fix = Get-XMLConfigWMIRepairEnable
                If ($Fix -eq "True") {
                    $Text = 'WMI Check: Corrupt. Attempting to repair WMI and reinstall ConfigMgr client.'
                    Write-Warning $Text
                    Repair-WMI
                    $Log.WMI = 'Repaired'
                } Else {
                    $Text = 'WMI Check: Corrupt. Autofix is not-enabled'
                    Write-Warning $Text
                    $Log.WMI = 'Corrupt'
                }
                Write-Verbose "Returning true to tag client for reinstall"
                $Obj = $true
            }
            Write-Output $Obj
        }
    }

    Function Repair-WMI {
        # Retrieve list of MOF files (excluding any that contain "Uninstall", "Remove", or "AutoRecover"), MFL files (excluding any that contain "Uninstall", or "Remove"), & DLL files from the WBEM folder
        $WbemContents = Get-ChildItem -Path "$env:WINDIR\System32\Wbem" -Recurse -File -Force
        $MofFiles = $WbemContents | Where-Object {$_.Extension -eq ".mof"} | Where-Object {$_.FullName -notmatch "Uninstall|Remove|Autorecover"} | Select-Object -ExpandProperty FullName
        $MflFiles = $WbemContents | Where-Object {$_.Extension -eq ".mfl"} | Where-Object {$_.FullName -notmatch "Uninstall|Remove"} | Select-Object -ExpandProperty FullName
        $DllFiles = $WbemContents | Where-Object {$_.Extension -eq ".dll"} | Select-Object -ExpandProperty FullName
        # Set Services for Volume Shadow Copy (VSS) and Microsoft Storage Spaces (SMPHost) to manual and stopped state prior to repository reset
        Set-Service -Name "vss" -Status Stopped -StartupType Manual -ErrorAction SilentlyContinue | Out-Null
        Set-Service -Name "smphost" -Status Stopped -StartUpType Manual -ErrorAction SilentlyContinue | Out-Null
        # Disable and Stop winmgmt service (Windows Management Instrumentation)
        Set-Service -Name "winmgmt" -Status Stopped -StartUpType Disabled -ErrorAction SilentlyContinue | Out-Null
        # This line resets the WMI repository, which renames current repository folder %systemroot%\system32\wbem\Repository to Repository.001
        & $env:WINDIR\System32\wbem\winmgmt.exe /resetrepository | Out-Null
        # These DLL Registers will help fix broken GPUpdate
        Start-Process -FilePath "$env:WINDIR\System32\regsvr32.exe" -ArgumentList "/s $env:WINDIR\System32\scecli.dll" -Wait
        Start-Process -FilePath "$env:WINDIR\System32\regsvr32.exe" -ArgumentList "/s $env:WINDIR\System32\userenv.dll" -Wait
        # These dll registers help ensure all DLLs for WMI are registered
        Foreach ($DllFilePath in $DllFiles) {Start-Process -FilePath "$env:WINDIR\System32\regsvr32.exe" -ArgumentList "/s $DllFilePath" -Wait}
        # Enable winmgmt service (WMI) and start the service
        Set-Service -Name "winmgmt" -Status Running -StartUpType Automatic -ErrorAction SilentlyContinue | Out-Null
        # Wait to let WMI Service start
        Start-Sleep -Seconds 15
        # Parse MOF and MFL files to add classes and class instances to WMI repository
        Foreach ($MofFilePath in $MofFiles) {& $env:WINDIR\System32\Wbem\mofcomp.exe $MofFilePath | Out-Null}
        Foreach ($MflFilePath in $MflFiles) {& $env:WINDIR\System32\Wbem\mofcomp.exe $MflFilePath | Out-Null}
    }

    # Test If the compliance state messages should be resent.
    Function Test-RefreshComplianceState {
        Param(
            $Days = 0,
            [Parameter(Mandatory = $true)]$RegistryKey,
            [Parameter(Mandatory = $true)]$Log
        )
        $RegValueName = "RefreshServerComplianceState"

        #Get the last time this script was ran.  If the Registry isn't found just use the current date.
        Try {
            [datetime]$LastSent = Get-RegistryValue -Path $RegistryKey -Name $RegValueName 
        } Catch {
            [datetime]$LastSent = Get-Date 
        }

        Write-Verbose "The compliance states were last sent on $($LastSent)"
        #Determine the number of days until the next run.
        $NumberOfDays = (New-TimeSpan -Start (Get-Date) -End ($LastSent.AddDays($Days))).Days

        #Resend complianc states If the next interval has already arrived or randomly based on the number of days left until the next interval.
        If (($NumberOfDays -le 0) -or ((Get-Random -Maximum $NumberOfDays) -eq 0 )) {
            Try {
                Write-Verbose "Resending compliance states."
                (New-Object -ComObject Microsoft.CCM.UpdatesStore).RefreshServerComplianceState()
                $LastSent = Get-Date
                Write-Output "Compliance States: Refreshed."
            } Catch {
                Write-Error "Failed to resend the compliance states."
                $LastSent = [datetime]::MinValue
            }
        } Else {
            Write-Output "Compliance States: OK."
        }

        Set-RegistryValue -Path $RegistryKey -Name $RegValueName -Value $LastSent
        $Log.RefreshComplianceState = Get-SmallDateTime $LastSent


    }

    # Start ConfigMgr Agent If not already running
    Function Test-SCCMService {
        If ($service.Status -ne 'Running') {
            Try {
                Start-Service -Name CcmExec | Out-Null 
            } Catch {
            }
        }
    }

    Function Test-SMSTSMgr {
        $service = Get-Service smstsmgr
        If (($service.ServicesDependedOn).name -contains "ccmexec") {
            Write-Host "SMSTSMgr: Removing dependency on CCMExec service."
            Start-Process sc.exe -ArgumentList "config smstsmgr depend= winmgmt" -Wait
        }

        # WMI service depenency is present by default
        If (($service.ServicesDependedOn).name -notcontains "Winmgmt") {
            Write-Host "SMSTSMgr: Adding dependency on Windows Management Instrumentaion service."
            Start-Process sc.exe -ArgumentList "config smstsmgr depend= winmgmt" -Wait
        } Else {
            Write-Host "SMSTSMgr: OK" 
        }
    }


    # Windows Service Functions
    Function Test-Services {
        Param([Parameter(Mandatory = $false)]$Xml, $log, $Webservice, $ProfileID)

        $log.Services = 'OK'

        # Test services defined by config.xml
        Write-Verbose 'Test services from XML configuration file'
        foreach ($service in $Xml.Configuration.Service) {
            $startuptype = ($service.StartupType).ToLower()

            If ($startuptype -like "automatic (delayed start)") {
                $service.StartupType = "automaticd" 
            }

            If ($service.uptime) {
                $uptime = ($service.Uptime).ToLower()
                Test-Service -Name $service.Name -StartupType $service.StartupType -State $service.State -Log $log -Uptime $uptime
            } Else {
                Test-Service -Name $service.Name -StartupType $service.StartupType -State $service.State -Log $log
            }
        }
    }

    Function Test-Service {
        Param(
            [Parameter(Mandatory = $True,
                HelpMessage = 'Name')]
            [string]$Name,
            [Parameter(Mandatory = $True,
                HelpMessage = 'StartupType: Automatic, Automatic (Delayed Start), Manual, Disabled')]
            [string]$StartupType,
            [Parameter(Mandatory = $True,
                HelpMessage = 'State: Running, Stopped')]
            [string]$State,
            [Parameter(Mandatory = $False,
                HelpMessage = 'Updatime in days')]
            [int]$Uptime,
            [Parameter(Mandatory = $True)]$log
        )

        $OSName = Get-OperatingSystem

        # Handle all sorts of casing and mispelling of delayed and triggerd start in config.xml services
        $val = $StartupType.ToLower()
        switch -Wildcard ($val) {
            "automaticd*" {
                $StartupType = "Automatic (Delayed Start)" 
            }
            "automatic(d*" {
                $StartupType = "Automatic (Delayed Start)" 
            }
            "automatic(t*" {
                $StartupType = "Automatic (Trigger Start)" 
            }
            "automatict*" {
                $StartupType = "Automatic (Trigger Start)" 
            }
        }

        $path = "HKLM:\SYSTEM\CurrentControlSet\Services\$name"

        $DelayedAutostart = (Get-ItemProperty -Path $path).DelayedAutostart
        If ($DelayedAutostart -ne 1) {
            $DelayedAutostart = 0
        }

        $service = Get-Service -Name $Name
        If ($PowerShellVersion -ge 6) {
            $WMIService = Get-CimInstance -Class Win32_Service -Property StartMode, ProcessID, Status -Filter "Name='$Name'" 
        } Else {
            $WMIService = Get-WmiObject -Class Win32_Service -Property StartMode, ProcessID, Status -Filter "Name='$Name'" 
        }
        $StartMode = ($WMIService.StartMode).ToLower()

        switch -Wildcard ($StartMode) {
            "auto*" {
                If ($DelayedAutostart -eq 1) {
                    $serviceStartType = "Automatic (Delayed Start)" 
                } Else {
                    $serviceStartType = "Automatic" 
                }
            }

            <# This will be implemented at a later time.
            "automatic d*" {$serviceStartType = "Automatic (Delayed Start)"}
            "automatic (d*" {$serviceStartType = "Automatic (Delayed Start)"}
            "automatic (t*" {$serviceStartType = "Automatic (Trigger Start)"}
            "automatic t*" {$serviceStartType = "Automatic (Trigger Start)"}
            #>
            "manual" {
                $serviceStartType = "Manual" 
            }
            "disabled" {
                $serviceStartType = "Disabled" 
            }
        }

        Write-Verbose "VerIfy startup type"
        If ($serviceStartType -eq $StartupType) {
            $Text = "Service $Name startup: OK"
            Write-Output $Text
        } Elseif ($StartupType -like "Automatic (Delayed Start)") {
            # Handle Automatic Trigger Start the dirty way for these two services. Implement in a nice way in future version.
            If ( (($name -eq "wuauserv") -or ($name -eq "W32Time")) -and (($OSName -like "Windows 10*") -or ($OSName -like "*Server 2016*")) ) {
                If ($service.StartType -ne "Automatic") {
                    $Text = "Configuring service $Name StartupType to: Automatic (Trigger Start)..."
                    Set-Service -Name $service.Name -StartupType Automatic
                } Else {
                    $Text = "Service $Name startup: OK" 
                }
                Write-Output $Text
            } Else {
                # Automatic delayed requires the use of sc.exe
                & sc.exe config $service start= delayed-auto | Out-Null
                $Text = "Configuring service $Name StartupType to: $StartupType..."
                Write-Output $Text
                $log.Services = 'Started'
            }
        }

        Else {
            Try {
                $Text = "Configuring service $Name StartupType to: $StartupType..."
                Write-Output $Text
                Set-Service -Name $service.Name -StartupType $StartupType
                $log.Services = 'Started'
            } Catch {
                $Text = "Failed to set $StartupType StartupType on service $Name"
                Write-Error $Text
            }
        }

        Write-Verbose 'VerIfy service is running'
        If ($service.Status -eq "Running") {
            $Text = 'Service ' + $Name + ' running: OK'
            Write-Output $Text

            #If we are checking uptime.
            If ($Uptime) {
                Write-Verbose "VerIfy the $($Name) service hasn't exceeded uptime of $($Uptime) days."
                $ServiceUptime = Get-ServiceUpTime -Name $Name
                If ($ServiceUptime -ge $Uptime) {
                    Try {

                        #Before restarting the service wait for some known processes to end.  Restarting the service While an app or updates is installing might cause issues.
                        $Timer = [Diagnostics.Stopwatch]::StartNew()
                        $WaitMinutes = 30
                        $ProcessesStopped = $True
                        While ((Get-Process -Name WUSA, wuauclt, setup, TrustedInstaller, msiexec, TiWorker, ccmsetup -ErrorAction SilentlyContinue).Count -gt 0) {
                            $MinutesLeft = $WaitMinutes - $Timer.Elapsed.Minutes

                            If ($MinutesLeft -le 0) {
                                Write-Warning "Timed out waiting $($WaitMinutes) minutes for installation processes to complete.  Will not restart the $($Name) service."
                                $ProcessesStopped = $False
                                Break
                            }
                            Write-Warning "Waiting $($MinutesLeft) minutes for installation processes to complete."
                            Start-Sleep -Seconds 30
                        }
                        $Timer.Stop()

                        #If the processes are not running the restart the service.
                        If ($ProcessesStopped) {
                            Write-Output "Restarting service: $($Name)..."
                            Restart-Service -Name $service.Name -Force
                            Write-Output "Restarted service: $($Name)..."
                            $log.Services = 'Restarted'
                        }
                    } Catch {
                        $Text = "Failed to restart service $($Name)"
                        Write-Error $Text
                    }
                } Else {
                    Write-Output "Service $($Name) uptime: OK"
                }
            }
        } Else {
            If ($WMIService.Status -eq 'Degraded') {
                Try {
                    Write-Warning "IdentIfied $Name service in a 'Degraded' state. Will force $Name process to stop."
                    $ServicePID = $WMIService | Select-Object -ExpandProperty ProcessID
                    Stop-Process -Id $ServicePID -Force:$true -Confirm:$false -ErrorAction Stop
                    Write-Verbose "Succesfully stopped the $Name service process which was in a degraded state."
                } Catch {
                    Write-Error "Failed to force $Name process to stop."
                }
            }
            Try {
                $ReTryService = $False
                $Text = 'Starting service: ' + $Name + '...'
                Write-Output $Text
                Start-Service -Name $service.Name -ErrorAction Stop
                $log.Services = 'Started'
            } Catch {
                #Error 1290 (-2146233087) indicates that the service is sharing a thread with another service that is protected and cannot share its thread.
                #This is resolved by configuring the service to run on its own thread.
                If ($_.Exception.Hresult -eq '-2146233087') {
                    Write-Output "Failed to start service $Name because it's sharing a thread with another process.  Changing to use its own thread."
                    & cmd /c sc config $Name type= own
                    $ReTryService = $True
                } Else {
                    $Text = 'Failed to start service ' + $Name
                    Write-Error $Text
                }
            }

            #If a recoverable error was found, Try starting it again.
            If ($ReTryService) {
                Try {
                    Start-Service -Name $service.Name -ErrorAction Stop
                    $log.Services = 'Started'
                } Catch {
                    $Text = 'Failed to start service ' + $Name
                    Write-Error $Text
                }
            }
        }
    }

    function Test-AdminShare {
        Param([Parameter(Mandatory = $true)]$Log)
        Write-Verbose "Test the ADMIN$ and C$"
        If ($PowerShellVersion -ge 6) {
            $share = Get-CimInstance Win32_Share | Where-Object { $_.Name -like 'ADMIN$' } 
        } Else {
            $share = Get-WmiObject Win32_Share | Where-Object { $_.Name -like 'ADMIN$' } 
        }
        #$shareClass = [WMICLASS]"WIN32_Share"  # Depreciated

        If ($share.Name -contains 'ADMIN$') {
            $Text = 'Adminshare Admin$: OK'
            Write-Output $Text
        } Else {
            $fix = $true 
        }

        If ($PowerShellVersion -ge 6) {
            $share = Get-CimInstance Win32_Share | Where-Object { $_.Name -like 'C$' } 
        } Else {
            $share = Get-WmiObject Win32_Share | Where-Object { $_.Name -like 'C$' } 
        }
        #$shareClass = [WMICLASS]'WIN32_Share'  # Depreciated

        If ($share.Name -contains "C$") {
            $Text = 'Adminshare C$: OK'
            Write-Output $Text
        } Else {
            $fix = $true 
        }

        If ($fix -eq $true) {
            $Text = 'Error with Adminshares. Remediating...'
            $log.AdminShare = 'Repaired'
            Write-Warning $Text
            Stop-Service server -Force
            Start-Service server
        } Else {
            $log.AdminShare = 'OK' 
        }
    }

    Function Test-DiskSpace {
        $XMLDiskSpace = Get-XMLConfigOSDiskFreeSpace
        If ($PowerShellVersion -ge 6) {
            $driveC = Get-CimInstance -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object FreeSpace, Size 
        } Else {
            $driveC = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq "$env:SystemDrive" } | Select-Object FreeSpace, Size 
        }
        $freeSpace = (($driveC.FreeSpace / $driveC.Size) * 100)

        If ($freeSpace -le $XMLDiskSpace) {
            $Text = "Local disk $env:SystemDrive Less than $XMLDiskSpace % free space"
            Write-Error $Text
        } Else {
            $Text = "Free space $env:SystemDrive OK"
            Write-Output $Text
        }
    }


    Function Test-CCMSoftwareDistribution {
        # TODO Implement this function
        Get-WmiObject -Class CCM_SoftwareDistributionClientConfig
    }

    Function Get-UBR {
        $UBR = (Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion').UBR
        Write-Output $UBR
    }

    Function Get-LastReboot {
        Param([Parameter(Mandatory = $false)][xml]$Xml)

        # Only run If option in config is enabled
        If (($Xml.Configuration.Option | Where-Object { $_.Name -like 'RebootApplication' } | Select-Object -ExpandProperty 'Enable') -like 'True') {
            $execute = $true 
        }

        If ($execute -eq $true) {

            [float]$maxRebootDays = Get-XMLConfigMaxRebootDays
            If ($PowerShellVersion -ge 6) {
                $wmi = Get-CimInstance Win32_OperatingSystem 
            } Else {
                $wmi = Get-WmiObject Win32_OperatingSystem 
            }

            $lastBootTime = $wmi.ConvertToDateTime($wmi.LastBootUpTime)

            $uptime = (Get-Date) - ($wmi.ConvertToDateTime($wmi.lastbootuptime))
            If ($uptime.TotalDays -lt $maxRebootDays) {
                $Text = 'Last boot time: ' + $lastBootTime + ': OK'
                Write-Output $Text
            } Elseif (($uptime.TotalDays -ge $maxRebootDays) -and (Get-XMLConfigRebootApplicationEnable -eq $true)) {
                $Text = 'Last boot time: ' + $lastBootTime + ': More than ' + $maxRebootDays + ' days since last reboot. Starting reboot application.'
                Write-Warning $Text
                Start-RebootApplication
            } Else {
                $Text = 'Last boot time: ' + $lastBootTime + ': More than ' + $maxRebootDays + ' days since last reboot. Reboot application disabled.'
                Write-Warning $Text
            }
        }
    }

    Function Start-RebootApplication {
        $taskName = 'ConfigMgr Client Health - Reboot on demand'
        #$OS = Get-OperatingSystem
        #If ($OS -like "*Windows 7*") {
        $task = schtasks.exe /query | FIND /I "ConfigMgr Client Health - Reboot"
        #}
        #Else { $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue }
        If ($task -eq $null) {
            New-RebootTask -taskName $taskName 
        }
        #If ($OS -notlike "*Windows 7*") {Start-ScheduledTask -TaskName $taskName }
        #Else {
        schtasks.exe /Run /TN $taskName
        #}
    }

    Function New-RebootTask {
        Param([Parameter(Mandatory = $true)]$taskName)
        $rebootApp = Get-XMLConfigRebootApplication

        # $execute is the executable file, $arguement is all the arguments added to it.
        $execute, $arguments = $rebootApp.Split(' ')
        $argument = $null

        foreach ($i in $arguments) {
            $argument += $i + " " 
        }

        # Trim the " " from argument If present
        $i = $argument.Length - 1
        If ($argument.Substring($i) -eq ' ') {
            $argument = $argument.Substring(0, $argument.Length - 1) 
        }

        #$OS = Get-OperatingSystem
        #If ($OS -like "*Windows 7*") {
        schtasks.exe /Create /tn $taskName /tr "$execute $argument" /ru "BUILTIN\Users" /sc ONCE /st 00:00 /sd 01/01/1901
        #}
        <#
        Else {
            $action = New-ScheduledTaskAction -Execute $execute -Argument $argument
            $userPrincipal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545"
            Register-ScheduledTask -Action $action -TaskName $taskName -Principal $userPrincipal | Out-Null
        }
        #>
    }

    Function Start-Ccmeval {
        Write-Host "Starting Built-in Configuration Manager Client Health Evaluation"
        $task = "Microsoft\Configuration Manager\Configuration Manager Health Evaluation"
        schtasks.exe /Run /TN $task | Out-Null
    }

    Function Test-MissingDrivers {
        Param([Parameter(Mandatory = $true)]$Log)
        $FileLogLevel = ((Get-XMLConfigLoggingLevel).ToString()).ToLower()
        $i = 0
        If ($PowerShellVersion -ge 6) {
            $devices = Get-CimInstance Win32_PNPEntity | Where-Object { ($_.ConfigManagerErrorCode -ne 0) -and ($_.ConfigManagerErrorCode -ne 22) -and ($_.Name -notlike "*PS/2*") } | Select-Object Name, DeviceID 
        } Else {
            $devices = Get-WmiObject Win32_PNPEntity | Where-Object { ($_.ConfigManagerErrorCode -ne 0) -and ($_.ConfigManagerErrorCode -ne 22) -and ($_.Name -notlike "*PS/2*") } | Select-Object Name, DeviceID 
        }
        $devices | ForEach-Object { $i++ }

        If ($devices -ne $null) {
            $Text = "Drivers: $i unknown or faulty device(s)"
            Write-Warning $Text
            $log.Drivers = "$i unknown or faulty driver(s)"

            foreach ($device in $devices) {
                $Text = 'Missing or faulty driver: ' + $device.Name + '. Device ID: ' + $device.DeviceID
                Write-Warning $Text
                If (-NOT($FileLogLevel -like "clientlocal")) {
                    Out-LogFile -Xml $xml -Text $Text -Severity 2 
                }
            }
        } Else {
            $Text = "Drivers: OK"
            Write-Output $Text
            $log.Drivers = 'OK'
        }
    }

    # Function to store SCCM log file changes to be processed
    Function New-SCCMLogFileJob {
        Param(
            [Parameter(Mandatory = $true)]$Logfile,
            [Parameter(Mandatory = $true)]$Text,
            [Parameter(Mandatory = $true)]$SCCMLogJobs
        )

        $path = Get-CCMLogDirectory
        $file = "$path\$LogFile"
        $SCCMLogJobs.Rows.Add($file, $Text)
    }

    # Function to remove info in SCCM logfiles after remediation. This to prevent false positives triggering remediation next time script runs
    Function Update-SCCMLogFile {
        Param([Parameter(Mandatory = $true)]$SCCMLogJobs)
        Write-Verbose "Start Update-SCCMLogFile"
        foreach ($job in $SCCMLogJobs) {
            Get-Content -Path $job.File | Where-Object { $_ -notmatch $job.Text } | Out-File $job.File -Force 
        }
        Write-Verbose "End Update-SCCMLogFile"
    }

    Function Test-SCCMHardwareInventoryScan {
        Param([Parameter(Mandatory = $true)]$Log)

        Write-Verbose "Start Test-SCCMHardwareInventoryScan"
        $days = Get-XMLConfigHardwareInventoryDays
        If ($PowerShellVersion -ge 6) {
            $wmi = Get-CimInstance -Namespace root\ccm\invagt -Class InventoryActionStatus | Where-Object { $_.InventoryActionID -eq '{00000000-0000-0000-0000-000000000001}' } | Select-Object @{label = 'HWSCAN'; expression = { $_.ConvertToDateTime($_.LastCycleStartedDate) } } 
        } Else {
            $wmi = Get-WmiObject -Namespace root\ccm\invagt -Class InventoryActionStatus | Where-Object { $_.InventoryActionID -eq '{00000000-0000-0000-0000-000000000001}' } | Select-Object @{label = 'HWSCAN'; expression = { $_.ConvertToDateTime($_.LastCycleStartedDate) } } 
        }
        $HWScanDate = $wmi | Select-Object -ExpandProperty HWSCAN
        $HWScanDate = Get-SmallDateTime $HWScanDate
        $minDate = Get-SmallDateTime((Get-Date).AddDays(-$days))
        If ($HWScanDate -le $minDate) {
            $fix = (Get-XMLConfigHardwareInventoryFix).ToLower()
            If ($fix -eq "true") {
                $Text = "ConfigMgr Hardware Inventory scan: $HWScanDate. Starting hardware inventory scan of the client."
                Write-Host $Text
                Invoke-SCCMPolicyHardwareInventory

                # Get the new date after policy trigger
                If ($PowerShellVersion -ge 6) {
                    $wmi = Get-CimInstance -Namespace root\ccm\invagt -Class InventoryActionStatus | Where-Object { $_.InventoryActionID -eq '{00000000-0000-0000-0000-000000000001}' } | Select-Object @{label = 'HWSCAN'; expression = { $_.ConvertToDateTime($_.LastCycleStartedDate) } } 
                } Else {
                    $wmi = Get-WmiObject -Namespace root\ccm\invagt -Class InventoryActionStatus | Where-Object { $_.InventoryActionID -eq '{00000000-0000-0000-0000-000000000001}' } | Select-Object @{label = 'HWSCAN'; expression = { $_.ConvertToDateTime($_.LastCycleStartedDate) } } 
                }
                $HWScanDate = $wmi | Select-Object -ExpandProperty HWSCAN
                $HWScanDate = Get-SmallDateTime -Date $HWScanDate
            } Else {
                # No need to update anything If fix = false. Last date will still be set in log
            }


        } Else {
            $Text = "ConfigMgr Hardware Inventory scan: OK"
            Write-Output $Text
        }
        $log.HWInventory = $HWScanDate
        Write-Verbose "End Test-SCCMHardwareInventoryScan"
    }

    # TODO: Implement so result of this remediation is stored in WMI log object, next to result of previous WMI check. This do not require db or webservice update
    # ref: https://social.technet.microsoft.com/Forums/de-DE/1f48e8d8-4e13-47b5-ae1b-dcb831c0a93b/setup-was-unable-to-compile-the-file-discoverystatusmof-the-error-code-is-8004100e?forum=configmanagerdeployment
    Function Test-PolicyPlatform {
        Param(
            [Parameter(Mandatory = $true)]$Log
        )
        Try {
            If (Get-CimInstance -Namespace 'Root/Microsoft' -ClassName '__Namespace' -Filter 'Name = "PolicyPlatform"') {
                Write-Host "PolicyPlatform: OK"
            } Else {
                Write-Warning "PolicyPlatform: Not found, recompiling WMI 'Microsoft Policy Platform\ExtendedStatus.mof'"
                Repair-WMI
                # Update WMI log object
                $Text = 'PolicyPlatform Recompiled.'
                If ($Log.WMI -eq 'OK') {
                    $Log.WMI = $Text
                } Else {
                    $Log.WMI += ". $Text"
                }
            }
        } Catch {
            Write-Warning "PolicyPlatform: RecompilePolicyPlatform failed!"
        }
    }


    # Get the clients SiteName in Active Directory
    Function Get-ClientSiteName {
        Try {
            If ($PowerShellVersion -ge 6) {
                $obj = (Get-CimInstance Win32_NTDomain).ClientSiteName 
            } Else {
                $obj = (Get-WmiObject Win32_NTDomain).ClientSiteName 
            }
        } Catch {
            $obj = $false 
        } finally {
            If ($obj -ne $false) {
                Write-Output ($obj | Select-Object -First 1) 
            } 
        }
    }

    Function Test-SoftwareMeteringPrepDriver {
        Param([Parameter(Mandatory = $true)]$Log)
        # To execute function: If (Test-SoftwareMeteringPrepDriver -eq $false) {$restartCCMExec = $true}
        # Thanks to Paul Andrews for letting me know about this issue.
        # And Sherry Kissinger for a nice fix: https://mnscug.org/blogs/sherry-kissinger/481-configmgr-ccmrecentlyusedapps-blank-or-mtrmgr-log-with-startprepdriver-openservice-failed-with-error-issue

        Write-Verbose "Start Test-SoftwareMeteringPrepDriver"

        $logdir = Get-CCMLogDirectory
        $logfile = "$logdir\mtrmgr.log"
        $content = Get-Content -Path $logfile
        $error1 = "StartPrepDriver - OpenService Failed with Error"
        $error2 = "Software Metering failed to start PrepDriver"

        If (($content -match $error1) -or ($content -match $error2)) {
            $fix = (Get-XMLConfigSoftwareMeteringFix).ToLower()

            If ($fix -eq "true") {
                $Text = "Software Metering - PrepDriver: Error. Remediating..."
                Write-Host $Text
                $CMClientDIR = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\SMS\Client\Configuration\Client Properties" -Name 'Local SMS Path').'Local SMS Path'
                $ExePath = $env:windir + '\system32\RUNDLL32.EXE'
                $CLine = ' SETUPAPI.DLL,InstallHinfSection DefaultInstall 128 ' + $CMClientDIR + 'prepdrv.inf'
                $ExePath = $env:windir + '\system32\RUNDLL32.EXE'
                $Prms = $Cline.Split(" ")
                & "$Exepath" $Prms

                $newContent = $content | Select-String -Pattern $error1, $error2 -NotMatch
                Stop-Service -Name CcmExec
                Out-File -FilePath $logfile -InputObject $newContent -Encoding utf8 -Force
                Start-Service -Name CcmExec

                $Obj = $false
                $Log.SWMetering = "Remediated"
            } Else {
                # Set $obj to true as we don't want to do anything with the CM agent.
                $obj = $true
                $Log.SWMetering = "Error"
            }
        } Else {
            $Text = "Software Metering - PrepDriver: OK"
            Write-Host $Text
            $Obj = $true
            $Log.SWMetering = "OK"
        }
        $content = $null # Clean the variable containing the log file.

        Write-Output $Obj
        Write-Verbose "End Test-SoftwareMeteringPrepDriver"
    }

    Function Test-SCCMHWScanErrors {
        # Function to test and fix errors that prevent a computer to perform a HW scan. Not sure If this is really needed or not.
    }

    # Functions to run certain ConfigMgr client actions
    Function Invoke-SCCMPolicySourceUpdateMessage {
        $Trigger = "{00000000-0000-0000-0000-000000000032}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }
    Function Invoke-SCCMPolicyMachineRequest {
        $Trigger = "{00000000-0000-0000-0000-000000000021}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }
    Function Invoke-SCCMPolicySendUnsentStateMessages {
        $Trigger = "{00000000-0000-0000-0000-000000000111}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }
    Function Invoke-SCCMPolicyScanUpdateSource {
        $Trigger = "{00000000-0000-0000-0000-000000000113}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }
    Function Invoke-SCCMPolicyHardwareInventory {
        $Trigger = "{00000000-0000-0000-0000-000000000001}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }
    Function Invoke-SCCMPolicyMachineEvaluation {
        $Trigger = "{00000000-0000-0000-0000-000000000022}"
        Invoke-CimMethod -Namespace 'Root\Ccm' -ClassName 'SMS_Client' -MethodName TriggerSchedule -Arguments @{sScheduleID = $Trigger} -ErrorAction SilentlyContinue | Out-Null
    }

    Function Get-Version {
        $Text = 'ConfigMgr Client Health Version ' + $Version
        Write-Output $Text
        Out-LogFile -Xml $xml -Text $Text -Severity 1
    }

    <# Trigger codes
    {00000000-0000-0000-0000-000000000001} Hardware Inventory
    {00000000-0000-0000-0000-000000000002} Software Inventory
    {00000000-0000-0000-0000-000000000003} Discovery Inventory
    {00000000-0000-0000-0000-000000000010} File Collection
    {00000000-0000-0000-0000-000000000011} IDMIf Collection
    {00000000-0000-0000-0000-000000000012} Client Machine Authentication
    {00000000-0000-0000-0000-000000000021} Request Machine Assignments
    {00000000-0000-0000-0000-000000000022} Evaluate Machine Policies
    {00000000-0000-0000-0000-000000000023} Refresh Default MP Task
    {00000000-0000-0000-0000-000000000024} LS (Location Service) Refresh Locations Task
    {00000000-0000-0000-0000-000000000025} LS (Location Service) Timeout Refresh Task
    {00000000-0000-0000-0000-000000000026} Policy Agent Request Assignment (User)
    {00000000-0000-0000-0000-000000000027} Policy Agent Evaluate Assignment (User)
    {00000000-0000-0000-0000-000000000031} Software Metering Generating Usage Report
    {00000000-0000-0000-0000-000000000032} Source Update Message
    {00000000-0000-0000-0000-000000000037} Clearing proxy settings cache
    {00000000-0000-0000-0000-000000000040} Machine Policy Agent Cleanup
    {00000000-0000-0000-0000-000000000041} User Policy Agent Cleanup
    {00000000-0000-0000-0000-000000000042} Policy Agent Validate Machine Policy / Assignment
    {00000000-0000-0000-0000-000000000043} Policy Agent Validate User Policy / Assignment
    {00000000-0000-0000-0000-000000000051} ReTrying/Refreshing certIficates in AD on MP
    {00000000-0000-0000-0000-000000000061} Peer DP Status reporting
    {00000000-0000-0000-0000-000000000062} Peer DP Pending package check schedule
    {00000000-0000-0000-0000-000000000063} SUM Updates install schedule
    {00000000-0000-0000-0000-000000000071} NAP action
    {00000000-0000-0000-0000-000000000101} Hardware Inventory Collection Cycle
    {00000000-0000-0000-0000-000000000102} Software Inventory Collection Cycle
    {00000000-0000-0000-0000-000000000103} Discovery Data Collection Cycle
    {00000000-0000-0000-0000-000000000104} File Collection Cycle
    {00000000-0000-0000-0000-000000000105} IDMIf Collection Cycle
    {00000000-0000-0000-0000-000000000106} Software Metering Usage Report Cycle
    {00000000-0000-0000-0000-000000000107} Windows Installer Source List Update Cycle
    {00000000-0000-0000-0000-000000000108} Software Updates Assignments Evaluation Cycle
    {00000000-0000-0000-0000-000000000109} Branch Distribution Point Maintenance Task
    {00000000-0000-0000-0000-000000000110} DCM policy
    {00000000-0000-0000-0000-000000000111} Send Unsent State Message
    {00000000-0000-0000-0000-000000000112} State System policy cache cleanout
    {00000000-0000-0000-0000-000000000113} Scan by Update Source
    {00000000-0000-0000-0000-000000000114} Update Store Policy
    {00000000-0000-0000-0000-000000000115} State system policy bulk send high
    {00000000-0000-0000-0000-000000000116} State system policy bulk send low
    {00000000-0000-0000-0000-000000000120} AMT Status Check Policy
    {00000000-0000-0000-0000-000000000121} Application manager policy action
    {00000000-0000-0000-0000-000000000122} Application manager user policy action
    {00000000-0000-0000-0000-000000000123} Application manager global evaluation action
    {00000000-0000-0000-0000-000000000131} Power management start summarizer
    {00000000-0000-0000-0000-000000000221} Endpoint deployment reevaluate
    {00000000-0000-0000-0000-000000000222} Endpoint AM policy reevaluate
    {00000000-0000-0000-0000-000000000223} External event detection
    #>

    function Test-SQLConnection {
        $SQLServer = Get-XMLConfigSQLServer
        $Database = 'ClientHealth'
        $FileLogLevel = ((Get-XMLConfigLoggingLevel).ToString()).ToLower()

        $ConnectionString = "Server={0};Database={1};Integrated Security=True;" -f $SQLServer, $Database

        Try {
            $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
            $sqlConnection.Open()
            $sqlConnection.Close()

            $obj = $true
            Write-Verbose "SQL connection test successfull"
        } Catch {
            $Text = "Error connecting to SQLDatabase $Database on SQL Server $SQLServer"
            Write-Error -Message $Text
            If (-NOT($FileLogLevel -like "clientinstall")) {
                Out-LogFile -Xml $xml -Text $Text -Severity 3 
            }
            $obj = $false
            Write-Verbose "SQL connection test failed"
        } finally {
            Write-Output $obj 
        }
    }

    # Invoke-SqlCmd2 - Created by Chad Miller
    function Invoke-Sqlcmd2 {
        [CmdletBinding()]
        Param(
            [Parameter(Position = 0, Mandatory = $true)] [string]$ServerInstance,
            [Parameter(Position = 1, Mandatory = $false)] [string]$Database,
            [Parameter(Position = 2, Mandatory = $false)] [string]$Query,
            [Parameter(Position = 3, Mandatory = $false)] [string]$Username,
            [Parameter(Position = 4, Mandatory = $false)] [string]$Password,
            [Parameter(Position = 5, Mandatory = $false)] [Int32]$QueryTimeout = 600,
            [Parameter(Position = 6, Mandatory = $false)] [Int32]$ConnectionTimeout = 15,
            [Parameter(Position = 7, Mandatory = $false)] [ValidateScript({ Test-Path $_ })] [string]$InputFile,
            [Parameter(Position = 8, Mandatory = $false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As = "DataRow"
        )

        If ($InputFile) {
            $filePath = $(Resolve-Path $InputFile).path
            $Query = [System.IO.File]::ReadAllText("$filePath")
        }

        $conn = New-Object System.Data.SqlClient.SQLConnection

        If ($Username) {
            $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance, $Database, $Username, $Password, $ConnectionTimeout 
        } Else {
            $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance, $Database, $ConnectionTimeout 
        }

        $conn.ConnectionString = $ConnectionString

        #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose Parameter specIfied by caller
        If ($PSBoundParameters.Verbose) {
            $conn.FireInfoMessageEventOnUserErrors = $true
            $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] { Write-Verbose "$($_)" }
            $conn.add_InfoMessage($handler)
        }

        $conn.Open()
        $cmd = New-Object system.Data.SqlClient.SqlCommand($Query, $conn)
        $cmd.CommandTimeout = $QueryTimeout
        $ds = New-Object system.Data.DataSet
        $da = New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
        [void]$da.fill($ds)
        $conn.Close()
        switch ($As) {
            'DataSet' {
                Write-Output ($ds) 
            }
            'DataTable' {
                Write-Output ($ds.Tables) 
            }
            'DataRow' {
                Write-Output ($ds.Tables[0]) 
            }
        }
    }


    # Gather info about the computer
    Function Get-Info {
        If ($PowerShellVersion -ge 6) {
            $OS = Get-CimInstance Win32_OperatingSystem
            $ComputerSystem = Get-CimInstance Win32_ComputerSystem
            If ($ComputerSystem.Manufacturer -like 'Lenovo') {
                $Model = (Get-CimInstance Win32_ComputerSystemProduct).Version 
            } Else {
                $Model = $ComputerSystem.Model 
            }
        } Else {
            $OS = Get-WmiObject Win32_OperatingSystem
            $ComputerSystem = Get-WmiObject Win32_ComputerSystem
            If ($ComputerSystem.Manufacturer -like 'Lenovo') {
                $Model = (Get-WmiObject Win32_ComputerSystemProduct).Version 
            } Else {
                $Model = $ComputerSystem.Model 
            }
        }

        $obj = New-Object PSObject -Property @{
            Hostname         = $ComputerSystem.Name
            Manufacturer     = $ComputerSystem.Manufacturer
            Model            = $Model
            Operatingsystem  = $OS.Caption
            Architecture     = $OS.OSArchitecture
            Build            = $OS.BuildNumber
            InstallDate      = Get-SmallDateTime -Date ($OS.ConvertToDateTime($OS.InstallDate))
            LastLoggedOnUser = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\').LastLoggedOnUser
        }

        $obj = $obj
        Write-Output $obj
    }

    # Start Getters - XML config file
    Function Get-LocalFilesPath {
        If ($config) {
            $obj = $Xml.Configuration.LocalFiles
        }
        $obj = $ExecutionContext.InvokeCommand.ExpandString($obj)
        If ($obj -eq $null) {
            $obj = Join-Path $env:SystemDrive "ClientHealth" 
        }
        Return $obj
    }

    Function Get-XMLConfigClientVersion {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Version' } | Select-Object -ExpandProperty '#text'
        }

        Write-Output $obj
    }

    Function Get-XMLConfigClientSitecode {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'SiteCode' } | Select-Object -ExpandProperty '#text'
        }

        Write-Output $obj
    }

    Function Get-XMLConfigClientDomain {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Domain' } | Select-Object -ExpandProperty '#text'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientAutoUpgrade {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'AutoUpgrade' } | Select-Object -ExpandProperty '#text'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientMaxLogSize {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Log' } | Select-Object -ExpandProperty 'MaxLogSize'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientMaxLogHistory {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Log' } | Select-Object -ExpandProperty 'MaxLogHistory'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientMaxLogSizeEnabled {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Log' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientCache {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'CacheSize' } | Select-Object -ExpandProperty 'Value'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientCacheDeleteOrphanedData {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'CacheSize' } | Select-Object -ExpandProperty 'DeleteOrphanedData'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientCacheEnable {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'CacheSize' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientShare {
        If ($config) {
            $obj = $Xml.Configuration.Client | Where-Object { $_.Name -like 'Share' } | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
        }

        If (!($obj)) {
            $obj = $global:ScriptPath 
        } #If Client share is empty, default to the script folder.
        Write-Output $obj
    }

    Function Get-XMLConfigUpdatesShare {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'Updates' } | Select-Object -ExpandProperty 'Share'
        }

        If (!($obj)) {
            $obj = Join-Path $global:ScriptPath "Updates" 
        }
        Return $obj
    }

    Function Get-XMLConfigUpdatesEnable {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'Updates' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigUpdatesFix {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'Updates' } | Select-Object -ExpandProperty 'Fix' 
        }
        Write-Output $obj
    }

    Function Get-XMLConfigLoggingShare {
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'File' } | Select-Object -ExpandProperty 'Share'
        }

        $obj = $ExecutionContext.InvokeCommand.ExpandString($obj)
        Return $obj
    }

    Function Get-XMLConfigLoggingLocalFile {
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'File' } | Select-Object -ExpandProperty 'LocalLogFile'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigLoggingEnable {
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'File' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigLoggingMaxHistory {
        # Currently not configurable through console extension and webservice. TODO
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'File' } | Select-Object -ExpandProperty 'MaxLogHistory'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigLoggingLevel {
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'File' } | Select-Object -ExpandProperty 'Level'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigLoggingTimeFormat {
        If ($config) {
            $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'Time' } | Select-Object -ExpandProperty 'Format'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigPendingRebootApp {
        # TODO verIfy this function
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'PendingReboot' } | Select-Object -ExpandProperty 'StartRebootApplication'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigMaxRebootDays {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'MaxRebootDays' } | Select-Object -ExpandProperty 'Days'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRebootApplication {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'RebootApplication' } | Select-Object -ExpandProperty 'Application'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRebootApplicationEnable {
        ### TODO implement in webservice
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'RebootApplication' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigDNSCheck {
        # TODO verIfy switch, skip test and monitor for console extension
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'DNSCheck' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigCcmSQLCELog {
        # TODO implement monitor mode
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'CcmSQLCELog' } | Select-Object -ExpandProperty 'Enable'
        }

        Write-Output $obj
    }

    Function Get-XMLConfigDNSFix {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'DNSCheck' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigDrivers {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'Drivers' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigPatchLevel {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'PatchLevel' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigOSDiskFreeSpace {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'OSDiskFreeSpace' } | Select-Object -ExpandProperty '#text'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigHardwareInventoryEnable {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'HardwareInventory' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigHardwareInventoryFix {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'HardwareInventory' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigSoftwareMeteringEnable {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'SoftwareMetering' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigSoftwareMeteringFix {
        # TODO implement this check in console extension and webservice
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'SoftwareMetering' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigHardwareInventoryDays {
        # TODO implement this check in console extension and webservice
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'HardwareInventory' } | Select-Object -ExpandProperty 'Days'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationAdminShare {
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'AdminShare' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationClientProvisioningMode {
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'ClientProvisioningMode' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationClientStateMessages {
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'ClientStateMessages' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationClientWUAHandler {
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'ClientWUAHandler' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationClientWUAHandlerDays {
        # TODO implement days in console extension and webservice
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'ClientWUAHandler' } | Select-Object -ExpandProperty 'Days'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigBITSCheck {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'BITSCheck' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigBITSCheckFix {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'BITSCheck' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigClientSettingsCheck {
        # TODO implement in console extension and webservice
        $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'ClientSettingsCheck' } | Select-Object -ExpandProperty 'Enable'
        Write-Output $obj
    }

    Function Get-XMLConfigClientSettingsCheckFix {
        # TODO implement in console extension and webservice
        $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'ClientSettingsCheck' } | Select-Object -ExpandProperty 'Fix'
        Write-Output $obj
    }

    Function Get-XMLConfigWMI {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'WMI' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigWMIRepairEnable {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'WMI' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRefreshComplianceState {
        # Measured in days
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'RefreshComplianceState' } | Select-Object -ExpandProperty 'Enable'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRefreshComplianceStateDays {
        If ($config) {
            $obj = $Xml.Configuration.Option | Where-Object { $_.Name -like 'RefreshComplianceState' } | Select-Object -ExpandProperty 'Days'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigRemediationClientCertIficate {
        If ($config) {
            $obj = $Xml.Configuration.Remediation | Where-Object { $_.Name -like 'ClientCertIficate' } | Select-Object -ExpandProperty 'Fix'
        }
        Write-Output $obj
    }

    Function Get-XMLConfigSQLServer {
        $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'SQL' } | Select-Object -ExpandProperty 'Server'
        Write-Output $obj
    }

    Function Get-XMLConfigSQLLoggingEnable {
        $obj = $Xml.Configuration.Log | Where-Object { $_.Name -like 'SQL' } | Select-Object -ExpandProperty 'Enable'
        Write-Output $obj
    }



    # End Getters - XML config file

    Function GetComputerInfo {
        $info = Get-Info | Select-Object HostName, OperatingSystem, Architecture, Build, InstallDate, Manufacturer, Model, LastLoggedOnUser
        #$Text = 'Computer info'+ "`n"
        $Text = 'Hostname: ' + $info.HostName
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'Operatingsystem: ' + $info.OperatingSystem
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'Architecture: ' + $info.Architecture
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'Build: ' + $info.Build
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'Manufacturer: ' + $info.Manufacturer
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'Model: ' + $info.Model
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'InstallDate: ' + $info.InstallDate
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
        $Text = 'LastLoggedOnUser: ' + $info.LastLoggedOnUser
        Write-Output $Text
        #Out-LogFile -Xml $xml $Text
    }

    Function Test-ConfigMgrHealthLogging {
        # VerIfies that logfiles are not bigger than max history

        
        $localLogging = (Get-XMLConfigLoggingLocalFile).ToLower()
        $fileshareLogging = (Get-XMLConfigLoggingEnable).ToLower()

        If ($localLogging -eq "true") {
            $clientpath = Get-LocalFilesPath
            $logfile = "$clientpath\ClientHealth.log"
            Test-LogFileHistory -Logfile $logfile
        }


        If ($fileshareLogging -eq "true") {
            $logfile = Get-LogFileName
            Test-LogFileHistory -Logfile $logfile
        }
    }

    Function CleanUp {
        $clientpath = (Get-LocalFilesPath).ToLower()
        $forbidden = "$env:SystemDrive", "$env:SystemDrive\", "$env:SystemDrive\windows", "$env:SystemDrive\windows\"
        $NoDelete = $false
        foreach ($item in $forbidden) {
            If ($clientpath -like $item) {
                $NoDelete = $true 
            } 
        }

        If (((Test-Path "$clientpath\Temp" -ErrorAction SilentlyContinue) -eq $True) -and ($NoDelete -eq $false) ) {
            Write-Verbose "Cleaning up temporary files in $clientpath\ClientHealth"
            Remove-Item "$clientpath\Temp" -Recurse -Force | Out-Null
        }

        $LocalLogging = ((Get-XMLConfigLoggingLocalFile).ToString()).ToLower()
        If (($LocalLogging -ne "true") -and ($NoDelete -eq $false)) {
            Write-Verbose "Local logging disabled. Removing $clientpath\ClientHealth"
            Remove-Item "$clientpath\Temp" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Function New-LogObject {
        # Write-Verbose "Start New-LogObject"

        If ($PowerShellVersion -ge 6) {
            $OS = Get-CimInstance -class Win32_OperatingSystem
            $CS = Get-CimInstance -class Win32_ComputerSystem
            If ($CS.Manufacturer -like 'Lenovo') {
                $Model = (Get-CimInstance Win32_ComputerSystemProduct).Version 
            } Else {
                $Model = $CS.Model 
            }
        } Else {
            $OS = Get-WmiObject -Class Win32_OperatingSystem
            $CS = Get-WmiObject -Class Win32_ComputerSystem
            If ($CS.Manufacturer -like 'Lenovo') {
                $Model = (Get-WmiObject Win32_ComputerSystemProduct).Version 
            } Else {
                $Model = $CS.Model 
            }
        }

        # Handles dIfferent OS languages
        $Hostname = Get-Hostname
        $OperatingSystem = $OS.Caption
        $Architecture = ($OS.OSArchitecture -replace ('([^0-9])(\.*)', '')) + '-Bit'
        $Build = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').BuildLabEx
        $Manufacturer = $CS.Manufacturer
        $ClientVersion = 'Unknown'
        $Sitecode = Get-Sitecode
        $Domain = Get-Domain
        [int]$MaxLogSize = 0
        $MaxLogHistory = 0
        If ($PowerShellVersion -ge 6) {
            $InstallDate = Get-SmallDateTime -Date ($OS.InstallDate) 
        } Else {
            $InstallDate = Get-SmallDateTime -Date ($OS.ConvertToDateTime($OS.InstallDate)) 
        }
        $InstallDate = $InstallDate -replace '\.', ':'
        $LastLoggedOnUser = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\').LastLoggedOnUser
        $CacheSize = Get-ClientCache
        $Services = 'Unknown'
        $Updates = 'Unknown'
        $DNS = 'Unknown'
        $Drivers = 'Unknown'
        $ClientCertIficate = 'Unknown'
        $PendingReboot = 'Unknown'
        $RebootApp = 'Unknown'
        If ($PowerShellVersion -ge 6) {
            $LastBootTime = Get-SmallDateTime -Date ($OS.LastBootUpTime) 
        } Else {
            $LastBootTime = Get-SmallDateTime -Date ($OS.ConvertToDateTime($OS.LastBootUpTime)) 
        }
        $LastBootTime = $LastBootTime -replace '\.', ':'
        $OSDiskFreeSpace = Get-OSDiskFreeSpace
        $AdminShare = 'Unknown'
        $ProvisioningMode = 'Unknown'
        $StateMessages = 'Unknown'
        $WUAHandler = 'Unknown'
        $WMI = 'Unknown'
        $RefreshComplianceState = Get-SmallDateTime
        $smallDateTime = Get-SmallDateTime
        $smallDateTime = $smallDateTime -replace '\.', ':'
        [float]$PSVersion = [float]$psVersion = [float]$PSVersionTable.PSVersion.Major + ([float]$PSVersionTable.PSVersion.Minor / 10)
        [int]$PSBuild = [int]$PSVersionTable.PSVersion.Build
        If ($PSBuild -le 0) {
            $PSBuild = $null 
        }
        $UBR = Get-UBR
        $BITS = $null
        $ClientSettings = $null

        $obj = New-Object PSObject -Property @{
            Hostname               = $Hostname
            Operatingsystem        = $OperatingSystem
            Architecture           = $Architecture
            Build                  = $Build
            Manufacturer           = $Manufacturer
            Model                  = $Model
            InstallDate            = $InstallDate
            OSUpdates              = $null
            LastLoggedOnUser       = $LastLoggedOnUser
            ClientVersion          = $ClientVersion
            PSVersion              = $PSVersion
            PSBuild                = $PSBuild
            Sitecode               = $Sitecode
            Domain                 = $Domain
            MaxLogSize             = $MaxLogSize
            MaxLogHistory          = $MaxLogHistory
            CacheSize              = $CacheSize
            ClientCertIficate      = $ClientCertIficate
            ProvisioningMode       = $ProvisioningMode
            DNS                    = $DNS
            Drivers                = $Drivers
            Updates                = $Updates
            PendingReboot          = $PendingReboot
            LastBootTime           = $LastBootTime
            OSDiskFreeSpace        = $OSDiskFreeSpace
            Services               = $Services
            AdminShare             = $AdminShare
            StateMessages          = $StateMessages
            WUAHandler             = $WUAHandler
            WMI                    = $WMI
            RefreshComplianceState = $RefreshComplianceState
            ClientInstalled        = $null
            Version                = $Version
            Timestamp              = $smallDateTime
            HWInventory            = $null
            SWMetering             = $null
            ClientSettings         = $null
            BITS                   = $BITS
            PatchLevel             = $UBR
            ClientInstalledReason  = $null
            RebootApp              = $RebootApp
        }
        Write-Output $obj
        # Write-Verbose "End New-LogObject"
    }

    Function Get-SmallDateTime {
        Param([Parameter(Mandatory = $false)]$Date)
        #Write-Verbose "Start Get-SmallDateTime"

        $UTC = (Get-XMLConfigLoggingTimeFormat).ToLower()

        If ($null -ne $Date) {
            If ($UTC -eq "utc") {
                $obj = (Get-UTCTime -DateTime $Date).ToString("yyyy-MM-dd HH:mm:ss") 
            } Else {
                $obj = ($Date).ToString("yyyy-MM-dd HH:mm:ss") 
            }
        } Else {
            $obj = Get-DateTime 
        }
        $obj = $obj -replace '\.', ':'
        Write-Output $obj
        #Write-Verbose "End Get-SmallDateTime"
    }

    # Test some values are whole numbers before attempting to insert / update database
    Function Test-ValuesBeforeLogUpdate {
        Write-Verbose "Start Test-ValuesBeforeLogUpdate"
        [int]$Log.MaxLogSize = [Math]::Round($Log.MaxLogSize)
        [int]$Log.MaxLogHistory = [Math]::Round($Log.MaxLogHistory)
        [int]$Log.PSBuild = [Math]::Round($Log.PSBuild)
        [int]$Log.CacheSize = [Math]::Round($Log.CacheSize)
        Write-Verbose "End Test-ValuesBeforeLogUpdate"
    }

    Function Update-SQL {
        Param(
            [Parameter(Mandatory = $true)]$Log,
            [Parameter(Mandatory = $false)]$Table
        )

        Write-Verbose "Start Update-SQL"
        Test-ValuesBeforeLogUpdate

        $SQLServer = Get-XMLConfigSQLServer
        $Database = 'ClientHealth'
        $table = 'dbo.Clients'
        $smallDateTime = Get-SmallDateTime

        If ($null -ne $log.OSUpdates) {
            # UPDATE
            $q1 = "OSUpdates='" + $log.OSUpdates + "', "
            # INSERT INTO
            $q2 = "OSUpdates, "
            # VALUES
            $q3 = "'" + $log.OSUpdates + "', "
        } Else {
            $q1 = $null
            $q2 = $null
            $q3 = $null
        }

        If ($null -ne $log.ClientInstalled) {
            # UPDATE
            $q10 = "ClientInstalled='" + $log.ClientInstalled + "', "
            # INSERT INTO
            $q20 = "ClientInstalled, "
            # VALUES
            $q30 = "'" + $log.ClientInstalled + "', "
        } Else {
            $q10 = $null
            $q20 = $null
            $q30 = $null
        }

        #ADD ClientSettings.log...
        $query = "begin tran
        If exists (SELECT * FROM $table WITH (updlock,serializable) WHERE Hostname='"+ $log.Hostname + "')
        begin
            UPDATE $table SET Operatingsystem='"+ $log.Operatingsystem + "', Architecture='" + $log.Architecture + "', Build='" + $log.Build + "', Manufacturer='" + $log.Manufacturer + "', Model='" + $log.Model + "', InstallDate='" + $log.InstallDate + "', $q1 LastLoggedOnUser='" + $log.LastLoggedOnUser + "', ClientVersion='" + $log.ClientVersion + "', PSVersion='" + $log.PSVersion + "', PSBuild='" + $log.PSBuild + "', Sitecode='" + $log.Sitecode + "', Domain='" + $log.Domain + "', MaxLogSize='" + $log.MaxLogSize + "', MaxLogHistory='" + $log.MaxLogHistory + "', CacheSize='" + $log.CacheSize + "', ClientCertIficate='" + $log.ClientCertIficate + "', ProvisioningMode='" + $log.ProvisioningMode + "', DNS='" + $log.DNS + "', Drivers='" + $log.Drivers + "', Updates='" + $log.Updates + "', PendingReboot='" + $log.PendingReboot + "', LastBootTime='" + $log.LastBootTime + "', OSDiskFreeSpace='" + $log.OSDiskFreeSpace + "', Services='" + $log.Services + "', AdminShare='" + $log.AdminShare + "', StateMessages='" + $log.StateMessages + "', WUAHandler='" + $log.WUAHandler + "', WMI='" + $log.WMI + "', RefreshComplianceState='" + $log.RefreshComplianceState + "', HWInventory='" + $log.HWInventory + "', Version='" + $Version + "', $q10 Timestamp='" + $smallDateTime + "', SWMetering='" + $log.SWMetering + "', BITS='" + $log.BITS + "', PatchLevel='" + $Log.PatchLevel + "', ClientInstalledReason='" + $log.ClientInstalledReason + "'
            WHERE Hostname = '"+ $log.Hostname + "'
        end
        Else
        begin
            INSERT INTO $table (Hostname, Operatingsystem, Architecture, Build, Manufacturer, Model, InstallDate, $q2 LastLoggedOnUser, ClientVersion, PSVersion, PSBuild, Sitecode, Domain, MaxLogSize, MaxLogHistory, CacheSize, ClientCertIficate, ProvisioningMode, DNS, Drivers, Updates, PendingReboot, LastBootTime, OSDiskFreeSpace, Services, AdminShare, StateMessages, WUAHandler, WMI, RefreshComplianceState, HWInventory, Version, $q20 Timestamp, SWMetering, BITS, PatchLevel, ClientInstalledReason)
            VALUES ('"+ $log.Hostname + "', '" + $log.Operatingsystem + "', '" + $log.Architecture + "', '" + $log.Build + "', '" + $log.Manufacturer + "', '" + $log.Model + "', '" + $log.InstallDate + "', $q3 '" + $log.LastLoggedOnUser + "', '" + $log.ClientVersion + "', '" + $log.PSVersion + "', '" + $log.PSBuild + "', '" + $log.Sitecode + "', '" + $log.Domain + "', '" + $log.MaxLogSize + "', '" + $log.MaxLogHistory + "', '" + $log.CacheSize + "', '" + $log.ClientCertIficate + "', '" + $log.ProvisioningMode + "', '" + $log.DNS + "', '" + $log.Drivers + "', '" + $log.Updates + "', '" + $log.PendingReboot + "', '" + $log.LastBootTime + "', '" + $log.OSDiskFreeSpace + "', '" + $log.Services + "', '" + $log.AdminShare + "', '" + $log.StateMessages + "', '" + $log.WUAHandler + "', '" + $log.WMI + "', '" + $log.RefreshComplianceState + "', '" + $log.HWInventory + "', '" + $log.Version + "', $q30 '" + $smallDateTime + "', '" + $log.SWMetering + "', '" + $log.BITS + "', '" + $Log.PatchLevel + "', '" + $Log.ClientInstalledReason + "')
        end
        commit tran"

        Try {
            Invoke-SqlCmd2 -ServerInstance $SQLServer -Database $Database -Query $query 
        } Catch {
            $ErrorMessage = $_.Exception.Message
            $Text = "Error updating SQL with the following query: $query. Error: $ErrorMessage"
            Write-Error $Text
            Out-LogFile -Xml $Xml -Text "ERROR Insert/Update SQL. SQL Query: $query `nSQL Error: $ErrorMessage" -Severity 3
        }
        Write-Verbose "End Update-SQL"
    }

    Function Update-LogFile {
        Param(
            [Parameter(Mandatory = $true)]$Log,
            [Parameter(Mandatory = $false)]$Mode
        )
        # Start the logfile
        Write-Verbose "Start Update-LogFile"
        #$share = Get-XMLConfigLoggingShare

        Test-ValuesBeforeLogUpdate
        $logfile = $logfile = Get-LogFileName
        Test-LogFileHistory -Logfile $logfile
        $Text = "<--- ConfigMgr Client Health Check starting --->"
        $Text += $log | Select-Object Hostname, Operatingsystem, Architecture, Build, Model, InstallDate, OSUpdates, LastLoggedOnUser, ClientVersion, PSVersion, PSBuild, SiteCode, Domain, MaxLogSize, MaxLogHistory, CacheSize, CertIficate, ProvisioningMode, DNS, PendingReboot, LastBootTime, OSDiskFreeSpace, Services, AdminShare, StateMessages, WUAHandler, WMI, RefreshComplianceState, ClientInstalled, Version, Timestamp, HWInventory, SWMetering, BITS, ClientSettings, PatchLevel, ClientInstalledReason | Out-String
        $Text = $Text.replace("`t", "")
        $Text = $Text.replace("  ", "")
        $Text = $Text.replace(" :", ":")
        $Text = $Text -creplace '(?m)^\s*\r?\n', ''

        If ($Mode -eq 'Local') {
            Out-LogFile -Xml $xml -Text $Text -Mode $Mode -Severity 1 
        } Elseif ($Mode -eq 'ClientInstalledFailed') {
            Out-LogFile -Xml $xml -Text $Text -Mode $Mode -Severity 1 
        } Else {
            Out-LogFile -Xml $xml -Text $Text -Severity 1 
        }
        Write-Verbose "End Update-LogFile"
    }

    # Write-Log : CMTrace compatible log file


    #endregion

    # Set default restart values to false
    $newinstall = $false
    $restartCCMExec = $false
    $Reinstall = $false


    # If config.xml is used
    If ($Config) {

        # Build the ConfigMgr Client Install Property string
        $propertyString = ""
        foreach ($property in $Xml.Configuration.ClientInstallProperty) {
            $propertyString = $propertyString + $property
            $propertyString = $propertyString + ' '
        }
        $clientCacheSize = Get-XMLConfigClientCache
        #replace to account for multiple skipreqs and escapee the character
        $clientInstallProperties = $propertyString.Replace(';', '`;')
        $clientAutoUpgrade = (Get-XMLConfigClientAutoUpgrade).ToLower()
        $AdminShare = Get-XMLConfigRemediationAdminShare
        $ClientProvisioningMode = Get-XMLConfigRemediationClientProvisioningMode
        $ClientStateMessages = Get-XMLConfigRemediationClientStateMessages
        $ClientWUAHandler = Get-XMLConfigRemediationClientWUAHandler
        $LogShare = Get-XMLConfigLoggingShare
    }

    # Create a DataTable to store all changes to log files to be processed later. This to prevent false positives to remediate the next time script runs If error is already remediated.
    $SCCMLogJobs = New-Object System.Data.DataTable
    [Void]$SCCMLogJobs.Columns.Add("File")
    [Void]$SCCMLogJobs.Columns.Add("Text")

}

Process {
    Write-Verbose "Starting precheck. Determing If script will run or not."
    # Veriy script is running with administrative priveleges.
    If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        $Text = 'ERROR: Powershell not running as Administrator! Client Health aborting.'
        Out-LogFile -Xml $Xml -Text $Text -Severity 3
        Write-Error $Text
        Exit 1
    } Else {
        # Will exit with errorcode 2 If in task sequence
        Test-InTaskSequence

        $StartupText1 = "PowerShell version: " + $PSVersionTable.PSVersion + ". Script executing with Administrator rights."
        Write-Host $StartupText1

        Write-Verbose "Determing If a task sequence is running."
        Try {
            $tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment | Out-Null 
        } Catch {
            $tsenv = $null 
        }

        If ($tsenv -ne $null) {
            $TSName = $tsenv.Value("_SMSTSAdvertID")
            Write-Host "Task sequence $TSName is active executing on computer. ConfigMgr Client Health will not execute."
            Exit 1
        } Else {
            $StartupText2 = "ConfigMgr Client Health " + $Version + " starting."
            Write-Host $StartupText2
        }
    }


    # If config.xml is used
    $LocalLogging = ((Get-XMLConfigLoggingLocalFile).ToString()).ToLower()
    $FileLogging = ((Get-XMLConfigLoggingEnable).ToString()).ToLower()
    $FileLogLevel = ((Get-XMLConfigLoggingLevel).ToString()).ToLower()
    $SQLLogging = ((Get-XMLConfigSQLLoggingEnable).ToString()).ToLower()


    $RegistryKey = "HKLM:\Software\ConfigMgrClientHealth"
    $LastRunRegistryValueName = "LastRun"

    #Get the last run from the Registry, defaulting to the minimum date value If the script has never ran.
    Try {
        [datetime]$LastRun = Get-RegistryValue -Path $RegistryKey -Name $LastRunRegistryValueName 
    } Catch {
        $LastRun = [datetime]::MinValue 
    }
    Write-Output "Script last ran: $($LastRun)"

    Write-Verbose "Testing If log files are bigger than max history for logfiles."
    Test-ConfigMgrHealthLogging

    # Create the log object containing the result of health check
    $Log = New-LogObject

    # Only test this is not using webservice
    If ($config) {
        Write-Verbose 'Testing SQL Server connection'
        If (($SQLLogging -like 'true') -and ((Test-SQLConnection) -eq $false)) {
            # Failed to create SQL connection. Logging this error to fileshare and aborting script.
            #Exit 1
        }
    }


    Write-Verbose 'Validating WMI is not corrupt...'
    $WMI = Get-XMLConfigWMI
    If ($WMI -like 'True') {
        Write-Verbose 'Checking If WMI is corrupt. Will reinstall configmgr client If WMI is rebuilt.'
        If ((Test-WMI -log $Log) -eq $true) {
            $reinstall = $true
            New-ClientInstalledReason -Log $Log -Message "Corrupt WMI."
        }
    }

    Write-Verbose 'Determining If compliance state should be resent...'
    $RefreshComplianceState = Get-XMLConfigRefreshComplianceState
    If ( ($RefreshComplianceState -like 'True') -or ($RefreshComplianceState -ge 1)) {
        $RefreshComplianceStateDays = Get-XMLConfigRefreshComplianceStateDays

        Write-Verbose "Checking If compliance state should be resent after $($RefreshComplianceStateDays) days."
        Test-RefreshComplianceState -Days $RefreshComplianceStateDays -RegistryKey $RegistryKey -log $Log
    }

    Write-Verbose 'Testing If ConfigMgr client is installed. Installing If not.'
    Test-ConfigMgrClient -Log $Log

    Write-Verbose 'Validating If ConfigMgr client is running the minimum version...'
    If ((Test-ClientVersion -Log $log) -eq $true) {
        If ($clientAutoUpgrade -like 'true') {
            $reinstall = $true
            New-ClientInstalledReason -Log $Log -Message "Below minimum verison."
        }
    }

    <#
    Write-Verbose 'Validate that ConfigMgr client do not have CcmSQLCE.log and are not in debug mode'
    If (Test-CcmSQLCELog -eq $true) {
        # This is a very bad situation. ConfigMgr agent is fubar. Local SDF files are deleted by the test itself, now reinstalling client immediatly. Waiting 10 minutes before continuing with health check.
        Resolve-Client -Xml $xml -ClientInstallProperties $ClientInstallProperties
        Start-Sleep -Seconds 600
    }
    #>

    Write-Verbose 'Validating services...'
    Test-Services -Xml $Xml -log $log

    Write-Verbose 'Validating SMSTSMgr service is depenent on CCMExec service...'
    Test-SMSTSMgr

    Write-Verbose 'Validating ConfigMgr SiteCode...'
    Test-ClientSiteCode -Log $Log

    Write-Verbose 'Validating client cache size. Will restart configmgr client If cache size is changed'

    $CacheCheckEnabled = Get-XMLConfigClientCacheEnable
    If ($CacheCheckEnabled -like 'True') {
        $TestClientCacheSzie = Test-ClientCacheSize -Log $Log
        # This check is now able to set ClientCacheSize without restarting CCMExec service.
        If ($TestClientCacheSzie -eq $true) {
            $restartCCMExec = $false 
        }
    }


    If ((Get-XMLConfigClientMaxLogSizeEnabled -like 'True') -eq $true) {
        Write-Verbose 'Validating Max CCMClient Log Size...'
        $TestClientLogSize = Test-ClientLogSize -Log $Log
        If ($TestClientLogSize -eq $true) {
            $restartCCMExec = $true 
        }
    }

    Write-Verbose 'Validating CCMClient provisioning mode...'
    If (($ClientProvisioningMode -like 'True') -eq $true) {
        Test-ProvisioningMode -log $log 
    }
    Write-Verbose 'Validating CCMClient certIficate...'

    If ((Get-XMLConfigRemediationClientCertIficate -like 'True') -eq $true) {
        Test-CCMCertIficateError -Log $Log 
    }
    If (Get-XMLConfigHardwareInventoryEnable -like 'True') {
        Test-SCCMHardwareInventoryScan -Log $log 
    }


    If (Get-XMLConfigSoftwareMeteringEnable -like 'True') {
        Write-Verbose "Testing software metering prep driver check"
        If ((Test-SoftwareMeteringPrepDriver -Log $Log) -eq $false) {
            $restartCCMExec = $true 
        }
    }

    Write-Verbose 'Validating DNS...'
    If ((Get-XMLConfigDNSCheck -like 'True' ) -eq $true) {
        Test-DNSConfiguration -Log $log 
    }

    Write-Verbose 'Validating BITS'
    If (Get-XMLConfigBITSCheck -like 'True') {
        If ((Test-BITS -Log $Log) -eq $true) {
            #$Reinstall = $true
        }
    }

    Write-Verbose 'Validating ClientSettings'
    If (Get-XMLConfigClientSettingsCheck -like 'True') {
        Test-ClientSettingsConfiguration -Log $log
    }

    If (($ClientWUAHandler -like 'True') -eq $true) {
        Write-Verbose 'Validating Windows Update Scan not broken by bad group policy...'
        $days = Get-XMLConfigRemediationClientWUAHandlerDays
        Test-RegistryPol -Days $days -log $log -StartTime $LastRun

    }


    If (($ClientStateMessages -like 'True') -eq $true) {
        Write-Verbose 'Validating that CCMClient is sending state messages...'
        Test-UpdateStore -log $log
    }

    Write-Verbose 'Validating Admin$ and C$ are shared...'
    If (($AdminShare -like 'True') -eq $true) {
        Test-AdminShare -log $log 
    }

    Write-Verbose 'Testing that all devices have functional drivers.'
    If ((Get-XMLConfigDrivers -like 'True') -eq $true) {
        Test-MissingDrivers -Log $log 
    }

    $UpdatesEnabled = Get-XMLConfigUpdatesEnable
    If ($UpdatesEnabled -like 'True') {
        Write-Verbose 'Validating required updates are installed...'
        Test-Update -Log $log
    }

    Write-Verbose "Validating $env:SystemDrive free diskspace (Only warning, no remediation)..."
    Test-DiskSpace
    Write-Verbose 'Getting install date of last OS patch for SQL log'
    Get-LastInstalledPatches -Log $log
    Write-Verbose 'Sending unsent state messages If any'
    Invoke-SCCMPolicySendUnsentStateMessages
    Write-Verbose 'Getting Source Update Message policy and policy to trigger scan update source'

    If ($newinstall -eq $false) {
        Invoke-SCCMPolicySourceUpdateMessage
        Invoke-SCCMPolicyScanUpdateSource
        Invoke-SCCMPolicySendUnsentStateMessages
    }
    Invoke-SCCMPolicyMachineEvaluation

    # Restart ConfigMgr client If tagged for restart and no reinstall tag
    If (($restartCCMExec -eq $true) -and ($Reinstall -eq $false)) {
        Write-Output "Restarting service CcmExec..."

        If ($SCCMLogJobs.Rows.Count -ge 1) {
            Stop-Service -Name CcmExec
            Write-Verbose "Processing changes to SCCM logfiles after remediation to prevent remediation again next time script runs."
            Update-SCCMLogFile
            Start-Service -Name CcmExec
        } Else {
            Restart-Service -Name CcmExec 
        }

        $Log.MaxLogSize = Get-ClientMaxLogSize
        $Log.MaxLogHistory = Get-ClientMaxLogHistory
        $log.CacheSize = Get-ClientCache
    }

    # Updating SQL Log object with current version number
    $log.Version = $Version

    Write-Verbose 'Cleaning up after healthcheck'
    CleanUp
    Write-Verbose 'Validating pending reboot...'
    Test-PendingReboot -log $log
    Write-Verbose 'Getting last reboot time'
    Get-LastReboot -Xml $xml

    If (Get-XMLConfigClientCacheDeleteOrphanedData -like "true") {
        Write-Verbose "Removing orphaned ccm client cache items."
        Remove-CCMOrphanedCache
    }

    # Reinstall client If tagged for reinstall and configmgr client is not already installing
    $proc = Get-Process ccmsetup -ErrorAction SilentlyContinue

    If (($reinstall -eq $true) -and ($null -ne $proc) ) {
        Write-Warning "ConfigMgr Client set to reinstall, but ccmsetup.exe is already running." 
    } Elseif (($Reinstall -eq $true) -and ($null -eq $proc)) {
        Write-Verbose 'Reinstalling ConfigMgr Client'
        Resolve-Client -Xml $Xml -ClientInstallProperties $ClientInstallProperties
        # Add smalldate timestamp in SQL for when client was installed by Client Health.
        $log.ClientInstalled = Get-SmallDateTime
        $Log.MaxLogSize = Get-ClientMaxLogSize
        $Log.MaxLogHistory = Get-ClientMaxLogHistory
        $log.CacheSize = Get-ClientCache

        # VerIfy that installed client version is now equal or better that minimum required client version
        $NewClientVersion = Get-ClientVersion
        $MinimumClientVersion = Get-XMLConfigClientVersion

        If ( $NewClientVersion -lt $MinimumClientVersion) {
            # ConfigMgr client version is still not at expected level.
            # Log for now, remediation is comming
            $Log.ClientInstalledReason += " Upgrade failed."
        }

    }

    # Get the latest client version in case it was reinstalled by the script
    $log.ClientVersion = Get-ClientVersion

    # Trigger default Microsoft CM client health evaluation
    Start-Ccmeval
    Write-Verbose "End Process"
}

End {
    # Update database and logfile with results

    #Set the last run.
    $Date = Get-Date
    Set-RegistryValue -Path $RegistryKey -Name $LastRunRegistryValueName -Value $Date
    Write-Output "Setting last ran to $($Date)"

    If ($LocalLogging -like 'true') {
        Write-Output 'Updating local logfile with results'
        Update-LogFile -Log $log -Mode 'Local'
    }

    If (($FileLogging -like 'true') -and ($FileLogLevel -like 'full')) {
        Write-Output 'Updating fileshare logfile with results'
        Update-LogFile -Log $log
    }

    If (($SQLLogging -eq 'true') -and -not $PSBoundParameters.ContainsKey('Webservice')) {
        Write-Output 'Updating SQL database with results'
        Update-SQL -Log $log
    }

    If ($PSBoundParameters.ContainsKey('Webservice')) {
        Write-Output 'Updating SQL database with results using webservice'
        Update-Webservice -URI $Webservice -Log $Log
    }
    Write-Verbose "Client Health script finished"
}

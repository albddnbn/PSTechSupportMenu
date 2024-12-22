Function GetTargets {
    <#
    .SYNOPSIS
        Queries Active Directory for computer names based on hostname, hostname substring, or comma-separated list of hostnames.

    .DESCRIPTION
        Long description

    .PARAMETER TargetComputer
        Parameter description

    .EXAMPLE
        GetTargets -TargetComputer "t-client-01"

    .EXAMPLE
        GetTargets -TargetComputer "t-client-01,t-client-02"

    .EXAMPLE
        GetTargets -TargetComputer "t-client-"

    .NOTES
        General notes
    #>
    param(
        [String[]]$TargetComputer
    )

    if ($TargetComputer -in @('', '127.0.0.1', 'localhost')) {
        $TargetComputer = @('127.0.0.1')
    }
    elseif ($(Test-Path $Targetcomputer -erroraction SilentlyContinue) -and ($TargetComputer.count -eq 1)) {
        $TargetComputer = Get-Content $TargetComputer
    }

    else {
        ## Prepare TargetComputer for LDAP query in ForEach loop
        ## if TargetComputer contains commas - it's either multiple comma separated hostnames, or multiple comma separated hostname substrings - either way LDAP query will verify
        if ($Targetcomputer -like "*,*") {
            $TargetComputer = $TargetComputer -split ','
        }
        else {
            $Targetcomputer = @($Targetcomputer)
        }

        ## LDAP query each TargetComputer item, create new list / sets back to Targetcomputer when done.
        $NewTargetComputer = [System.Collections.Arraylist]::new()
        foreach ($computer in $TargetComputer) {
            ## CREDITS FOR The code this was adapted from: https://intunedrivemapping.azurewebsites.net/DriveMapping
            if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN) -and [string]::IsNullOrEmpty($searchRoot)) {
                Write-Error "LDAP query `$env:USERDNSDOMAIN is not available!"
                Write-Warning "You can override your AD Domain in the `$overrideUserDnsDomain variable"
            }
            else {

                # if no domain specified fallback to PowerShell environment variable
                if ([string]::IsNullOrEmpty($searchRoot)) {
                    $searchRoot = $env:USERDNSDOMAIN
                }

                ## Thank you https://github.com/Jreece321 for this snippet - it shortened 10 lines of code to the 3 that you see below.
                $matching_hostnames = (([adsisearcher]"(&(objectCategory=Computer)(name=$computer*))").findall()).properties
                $matching_hostnames = $matching_hostnames.name
                $NewTargetComputer += $matching_hostnames
            }
        }
        $TargetComputer = $NewTargetComputer
    }

    # }
    $TargetComputer = $TargetComputer | Where-object { $_ -ne $null } | Select -Unique
    # Safety catch
    if ($null -eq $TargetComputer) {
        return
    }
    # }
    return $TargetComputer
}

function TestConnectivity {
    <#
    .SYNOPSIS
        Tests connectivity to a single computer or list of computers by using Test-Connection -Quiet.

    .DESCRIPTION
        Works fairly quickly, but doesn't give you any information about the computer's name, IP, or latency - judges online/offline by the 1 ping.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PingCount
        Number of pings sent to each target machine. Default is 1.

    .EXAMPLE
        Check all hostnames starting with t-client- for online/offline status.
        TestConnectivityQuick -TargetComputer "t-client-"

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        $PingCount = 1
    )
    ## 1. Set PingCount - # of pings sent to each target machine.
    ## 2. Handle Targetcomputer if not supplied through the pipeline.
    ## 1. Set PingCount - # of pings sent to each target machine.
    $PING_COUNT = $PingCount

    # $list_of_online_computers = [system.collections.arraylist]::new()
    # $list_of_offline_computers = [system.collections.arraylist]::new()
    $online_results = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {
        if (Test-Connection $single_computer -Count $PING_COUNT -Quiet) {
            Write-Host "$single_computer is online." -ForegroundColor Green
            
            $online_results.Add($single_computer) | Out-Null
        }
        else {
            Write-Host "$single_computer is offline." -ForegroundColor Red
        }
    
    }

    return $online_results

}

function Clear-FirefoxProfile {
    <#
    .SYNOPSIS
        Removes the profiles.ini file from the user's roaming profile, which will cause Firefox to recreate it on next launch.

    .PARAMETER ClearLocalProfile
        'y' clears Local Firefox profile, which ' will delete all bookmarks, history, etc.
        Submitting any other value will only target the profiles.ini file.

    .PARAMETER TargetPC
        The single target computer from which to remove the profiles.ini file (Hasn't seemed useful to add ability to target multiple computers yet).
        Enter '' (Press [ENTER]) for localhost.

    .PARAMETER Username
        Username of the user to remove the profiles.ini file from - used to target that users folder on the target computer.

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param (
        [Parameter(Mandatory = $false)]
        [String]$Username,
        [Parameter(Mandatory = $false)]
        [String]$TargetPC,
        [string]$ClearLocalProfile
    )
    if ($ClearLocalProfile.ToLower() -eq 'y') { $ClearLocalProfile = $true }
    else { $ClearLocalProfile = $false }
    
    ######################################################################################
    ## Scriptblock - removes Firefox profiles.ini / profile data from local profile folder
    ######################################################################################
    $delete_firefox_profile_Scriptblock = {
        param(
            $Targeted_user,
            $ClearProfile
        )
        # stop the running firefox processes
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Stopping any running firefox processes..."
        get-process 'firefox' | stop-process -force

        if ($ClearProfile) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Clearing local profiles..." -foregroundcolor yellow
            Remove-Item -Path "C:\Users\$targetuser\AppData\Local\Mozilla\Firefox\Profiles\*" -Recurse -Force
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Skipping deletion of local Firefox profiles."
        }
        # get rid of profiles.ini in roaming
        $profilesini = Get-ChildItem -Path "C:\Users\$Targeted_user\AppData\Roaming\Mozilla\Firefox" -Filter "profiles.ini" -File -ErrorAction SilentlyContinue
        if ($profilesini) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($profilesini.fullname), removing $($profilesini.fullname) from $env:COMPUTERNAME..."
            Remove-Item -Path $($profilesini.fullname) -Force

            Write-Host "Please ask the user to start firefox back up to see if the issue is resolved." -ForegroundColor Green
            return 0
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No profiles.ini file found in C:\Users\$Targeted_user\AppData\Roaming\Mozilla\Firefox" -ForegroundColor Red
            return 1
        }
    }

    ## If TargetPC isn't supplied, script is run on local computer
    if ($TargetPC -eq '') {
        $result = Invoke-Command -Scriptblock $delete_firefox_profile_Scriptblock -ArgumentList $Username, $ClearLocalProfile
        $result | Add-Member -MemberType NoteProperty -Name PScomputerName -Value $env:COMPUTERNAME
    }
    else {
        $result = Invoke-Command -ComputerName $TargetPC -Scriptblock $delete_firefox_profile_Scriptblock -ArgumentList $Username, $ClearLocalProfile
    }

    ## Display success/failure message
    if ($result -eq 0) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Successfully removed profiles.ini from $TargetPC." -ForegroundColor Green
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to remove profiles.ini from $TargetPC, or it wasn't there for $Username." -ForegroundColor Red
    }

    # read-host "Press enter to continue."
}

function Copy-RemoteFiles {
    <#
    .SYNOPSIS
        Recursively grabs target files or folders from remote computer(s) and copies them to specified directory on local computer.

    .DESCRIPTION
        TargetPath specifies the target file(s) or folder(s) to target on remote machines.
        TargetPath can be supplied as a single absolute path, comma-separated list, or array.
        OutputPath specifies the directory to store the retrieved files.
        Creates a subfolder for each target computer to store it's retrieved files.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputPath
        Path to folder to store retrieved files. Ex: 'C:\users\abuddenb\Desktop\grabbed-files'

    .PARAMETER TargetPath
        Path to file(s)/folder(s) to be grabbed from remote machines. Ex: 'C:\users\abuddenb\Desktop\test.txt'

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Copy-RemoteFiles -TargetPath "Users\Public\Desktop" -OutputPath "C:\Users\Public\Desktop" -TargetComputer "t-client-"

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(        
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$TargetPath,
        [string]$OutputFolder,
        $SendPings
    )

    ## 1. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 2. Make sure output folder path exists for remote files to be copied to.
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    ## 2. Make sure the outputpath folder exists (remote files are copied here):

    if (-not(Test-Path "$OutputFolder" -erroraction SilentlyContinue)) {
        New-Item -ItemType Directory -Path "$OutputFolder" -ErrorAction SilentlyContinue | out-null
    }
        
    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. Copy file from pssession on target machine, to local computer.
    ##    Report on success/fail
    ## 4. Remove the pssession.
    ForEach ($single_computer in $ComputerName) {
        ## 1. no empty Targetcomputer values past this point
        if ($single_computer) {
            ## 2. Make sure machine is responsive on the network
            $target_network_path = $targetpath -replace 'C:', "\\$single_computer\c$"
            if ([system.IO.Directory]::Exists("\\$single_computer\c$")) {
                if (Test-Path "$target_network_path" -erroraction SilentlyContinue) {

                    
                    $target_session = New-PSSession $single_computer

                    $target_filename = $targetpath | split-path -leaf


                    Copy-Item -Path "$targetpath" -Destination "$OutputFolder\$single_computer-$target_filename" -FromSession $target_session -Recurse
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $targetpath ($single_computer) to $OutputFolder\$single_computer-$target_filename  complete." -foregroundcolor green
                    
                    Remove-PSSession $target_session

                }
                else {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $targetpath on $single_computer to $OutputFolder on local computer." -foregroundcolor red
                }
                ## 4. Bye pssession
            }
        }
    }
    ## Open output folder, pause.
    if (Test-Path "$OutputFolder" -erroraction SilentlyContinue) {
        Invoke-item "$OutputFolder"
    }
    # read-host "Press enter to continue."
}

function Send-Files {
    <#
    .SYNOPSIS
        Sends a target file/folder from local computer to target path on remote computers.

    .DESCRIPTION
        You can enter both paths as if they're on local filesystem, the script should cut out any drive letters and insert the \\hostname\c$ for UNC path. The script only works for C drive on target computers right now.

    .PARAMETER SourcePath
        The path of the file/folder you want to send to target computers. 
        ex: C:\users\public\desktop\test.txt, 
        ex: \\networkshare\folder\test.txt

    .PARAMETER DestinationPath
        The path on the target computer where you want to send the file/folder. 
        The script will cut off any preceding drive letters and insert \\hostname\c$ - so destination paths should be on C drive of target computers.
        ex: C:\users\public\desktop\
        The path should be a destination folder, script should create directories/subdirectories as required.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        copy the test.txt file to all computers in stanton open lab
        Send-Files -sourcepath "C:\Users\Public\Desktop\test.txt" -destinationpath "Users\Public\Desktop" -targetcomputer "t-client-"

    .EXAMPLE
        Get-User -ComputerName "t-client-28"

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [ValidateScript({
                if (Test-Path $_ -ErrorAction SilentlyContinue) {
                    return $true
                }
                else {
                    Write-Error "SourcePath does not exist. Please enter a valid path."
                    return $false                
                }
            })]
        [string]$sourcepath,
        [string]$destinationpath,
        $SendPings
    )
    ## 1. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }        

    $informational_string = ""

    ## make sure destination path is formatted correctly:
    # if destination path is a folder path, add the source file name to the end of the path
    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. Use session to copy file from local computer.
    ##    Report on success/fail
    ## 4. Remove the pssession.
    ForEach ($single_computer in $ComputerName) {
        ## 1. no empty Targetcomputer values past this point
        if ($single_computer) {
            ## 2. Make sure machine is responsive on network
            $file_copied = $false
            if ([System.IO.Directory]::Exists("\\$single_computer\c$")) {
                ## make sure target directory exists:
                Invoke-Command -ComputerName $single_computer -Scriptblock {

                    $destination_folder = $using:destinationpath
                    if (-not (Test-Path $destination_folder -ErrorAction SilentlyContinue)) {
                        Write-Host "Creating $destination_folder directory on $env:COMPUTERNAME."
                        New-Item -Path $destination_folder -ItemType Directory -Force | Out-Null
                    }
                }

                $target_session = New-PSSession $single_computer

                Copy-Item -Path "$sourcepath" -Destination "$destinationpath" -ToSession $target_session -Recurse
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $sourcepath to $destinationpath ($single_computer) complete." -foregroundcolor green
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $sourcepath to $destinationpath ($single_computer) complete.`n"

                $file_copied = $true


                Remove-PSSession $target_session
                            
            }
            
            if (-not $file_copied) {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $sourcepath to $destinationpath on $single_computer." -foregroundcolor red
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $sourcepath to $destinationpath on $single_computer.`n"
            }
            
        }
    }
    ## 1. Write an ending message to terminal.
    ## announcement file for when function is run as background job
    if (-not $env:PSMENU_DIR) {
        $env:PSMENU_DIR = $(pwd)
    }
    ## create simple output path to reports directory
    $thedate = Get-Date -Format 'yyyy-MM-dd'
    $DIRECTORY_NAME = 'SendFiles'
    $OUTPUT_FILENAME = 'SendFiles'
    if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
    }
        
    $counter = 0
    do {
        $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
        $counter++
    } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))
        
        
    ## Append text to file here:
    $informational_string | Out-File -FilePath $output_filepath -Append
    # $ComputerName | Out-File -FilePath $output_filepath -Append
    "`nThe Scan-ForApporFilepath function can be used to verify file/folders' existence on target computers." | Out-File -FilePath $output_filepath -Append -Force
        
    ## then open the file:
    Invoke-Item "$output_filepath"
        
    # read-host "`nPress [ENTER] to continue."
}


function Send-Reboots {
    <#
    .SYNOPSIS
        Reboots the target computer(s) either with/without a message displayed to logged in users.

    .DESCRIPTION
        If a reboot msg isn't provided, no reboot msg/warning will be shown to logged in users.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .EXAMPLE
        Send-Reboot -TargetComputer "t-client-" -RebootMessage "This computer will reboot in 5 minutes." -RebootTimeInSeconds 300

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [Parameter(Mandatory = $false)]
        [string]$RebootMessage,
        # the time before reboot in seconds, 3600 = 1hr, 300 = 5min
        [Parameter(Mandatory = $false)]
        [string]$RebootTimeInSeconds = 300,
        $SendPings
    )
    ## 1. Confirm time before reboot w/user
    ## 2. Handling of TargetComputer input
    ## 3. typecast reboot time to double to be sure
    ## 4. container for offline computers
    ## 1. Confirmation
    $reply = Read-Host "Sending reboot in $RebootTimeInSeconds seconds, or $([double]$RebootTimeInSeconds / 60) minutes, OK? (y/n)"
    if ($reply.ToLower() -eq 'y') {
    
        $ComputerName = GetTargets -TargetComputer $ComputerName

        ## Ping Test for Connectivity:
        if ($SendPings -eq 'y') {
            $ComputerName = TestConnectivity -ComputerName $ComputerName
        }
    }
    ## 3. typecast to double
    $RebootTimeInSeconds = [double]$RebootTimeInSeconds

    ## 4. container for offline computers
    $offline_computers = [system.collections.arraylist]::new()

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session and/or reboot.
    ## 3. Send reboot either with or without message
    ## 4. If machine was offline - add it to list to output at end.
    if ($RebootMessage) {
        Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            shutdown  /r /t $using:RebootTimeInSeconds /c "$using:RebootMessage"
        }
        $reboot_method = "Reboot w/popup msg"
    }
    else {
        Restart-Computer $ComputerName
        $reboot_method = "Reboot using Restart-Computer (no Force)" # 2-28-2024
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Reboot(s) sent."
}

Function Set-ChromeClearDataOnExit {
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName
    )
    ## 1. Handling TargetComputer input if not supplied through pipeline.
    ## 2. Define scriptblock that sets Chrome data deletion registry settings.
    BEGIN {
        ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
        if ($null -eq $ComputerName) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected pipeline input for targetcomputer." -Foregroundcolor Yellow
        }
        else {
            ## Assigns localhost value
            if ($ComputerName -in @('', '127.0.0.1', 'localhost')) {
                $ComputerName = @('127.0.0.1')
            }
            ## If input is a file, gets content
            elseif ($(Test-Path $ComputerName -erroraction SilentlyContinue) -and ($ComputerName.count -eq 1)) {
                $ComputerName = Get-Content $ComputerName
            }
            ## A. Separates any comma-separated strings into an array, otherwise just creates array
            ## B. Then, cycles through the array to process each hostname/hostname substring using LDAP query
            else {
                ## A.
                if ($ComputerName -like "*,*") {
                    $ComputerName = $ComputerName -split ','
                }
                else {
                    $ComputerName = @($ComputerName)
                }
        
                ## B. LDAP query each TargetComputer item, create new list / sets back to Targetcomputer when done.
                $NewTargetComputer = [System.Collections.Arraylist]::new()
                foreach ($computer in $ComputerName) {
                    ## CREDITS FOR The code this was adapted from: https://intunedrivemapping.azurewebsites.net/DriveMapping
                    if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN) -and [string]::IsNullOrEmpty($searchRoot)) {
                        Write-Error "LDAP query `$env:USERDNSDOMAIN is not available!"
                        Write-Warning "You can override your AD Domain in the `$overrideUserDnsDomain variable"
                    }
                    else {
        
                        # if no domain specified fallback to PowerShell environment variable
                        if ([string]::IsNullOrEmpty($searchRoot)) {
                            $searchRoot = $env:USERDNSDOMAIN
                        }
                        
                        $matching_hostnames = (([adsisearcher]"(&(objectCategory=Computer)(name=$computer*))").findall()).properties
                        $matching_hostnames = $matching_hostnames.name
                        $NewTargetComputer += $matching_hostnames
                    }
                }
                $ComputerName = $NewTargetComputer
            }
            $ComputerName = $ComputerName | Where-object { $_ -ne $null } | Select -Unique
            # Safety catch
            if ($null -eq $ComputerName) {
                return
            }
        }
        ## 2. Scriptblock that runs on each target computer, setting registry values to cause Chrome to auto-delete
        ##    specified categories of browsing data on exit of the application.
        ##    This is useful for 'guest accounts' or 'testing center' computers, that are not likely to have to be 
        ##    reused by the same person.
        $chrome_setting_scriptblock = {
            $testforchromekey = Test-Path -Path "HKLM:\SOFTWARE\Policies\Google\Chrome\ClearBrowsingDataOnExitList" -erroraction silentlycontinue
            if (-not $testforchromekey) {
                New-Item -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "ClearBrowsingDataOnExitList" -Force
            }
    
            New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Google\Chrome" -Name "SyncDisabled" -Value 1 -PropertyType DWord -Force
    
            $chromehash = @{
                "1" = "browsing_history"
                "2" = "download_history"
                "3" = "cookies_and_other_site_data"
                "4" = "cached_images_and_files"
                "5" = "password_signin"
                "6" = "autofill"
                "7" = "site_settings"
                "8" = "hosted_app_data"
            }
            ForEach ($key in $chromehash.keys) {
                New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Google\Chrome\ClearBrowsingDataOnExitList" -Name $key -Value $chromehash[$key] -PropertyType String -Force
            }
            New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "BackgroundModeEnabled" -Value 0 -PropertyType DWORD -Force
            
            New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "ClearBrowsingDataOnExit" -Value 1 -PropertyType DWORD -Force
        }
    }
    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, Collect local asset information from computer
    PROCESS {
        ForEach ($single_computer in $ComputerName) {

            if ($single_computer) {

                ## test with ping first:
                $pingreply = Test-Connection $single_computer -Count 1 -Quiet
                if ($pingreply) {
                    Invoke-Command -ComputerName $single_computer -ScriptBlock $chrome_setting_scriptblock
                }
                else {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer didn't respond to one ping, skipping." -ForegroundColor Yellow
                }
            }
        }
    }

    END {
        ## create announcement file for when function is run as background job:
        if (-not $env:PSMENU_DIR) {
            $env:PSMENU_DIR = pwd
        }
        ## create simple output path to reports directory
        $thedate = Get-Date -Format 'yyyy-MM-dd'
        $DIRECTORY_NAME = 'ChromeClearDataOnExit'
        $OUTPUT_FILENAME = 'results'
        if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
            New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
        }
        
        $counter = 0
        do {
            $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
        } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Finished setting Chrome to clear data on exit for computers listed below." | Out-File -FilePath $output_filepath -Append
        $ComputerName | Out-File -FilePath $output_filepath -Append

        Invoke-Item "$output_filepath"
        
        # Read-Host "Press enter to continue."
    }
}

function Stop-WifiAdapters {
    <#
    .SYNOPSIS
        Attempts to turn off (and disable if 'y' entered for 'DisableWifiAdapter') Wi-Fi adapter of target device(s).
        Having an active Wi-Fi adapter/connection when the Ethernet adapter is also active can cause issues.

    .DESCRIPTION
        Function needs work - 1/13/2024.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER DisableWifiAdapter
        Optional parameter to disable the Wi-Fi adapter. If not specified, the function will only turn off wifi adapter.
        'y' or 'Y' will disable target Wi-Fi adapters.

    .EXAMPLE
        Stop-WifiAdapters -TargetComputer s-tc136-02 -DisableWifiAdapter y
        Turns off and disables Wi-Fi adapter on single computer/hostname s-tc136-02.

    .EXAMPLE
        Stop-WifiAdapters -TargetComputer t-client- -DisableWifiAdapter n
        Turns off Wi-Fi adapters on computers w/hostnames starting with t-client-, without disabling them.

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$DisableWifiAdapter = 'n'
    )
    ## 1. Handling of TargetComputer input
    ## 2. Define Turn off / disable wifi adapter scriptblock that gets run on each target computer
    BEGIN {
        ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
        if ($null -eq $ComputerName) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected pipeline input for targetcomputer." -Foregroundcolor Yellow
        }
        else {
            ## Assigns localhost value
            if ($ComputerName -in @('', '127.0.0.1', 'localhost')) {
                $ComputerName = @('127.0.0.1')
            }
            ## If input is a file, gets content
            elseif ($(Test-Path $ComputerName -erroraction SilentlyContinue) -and ($ComputerName.count -eq 1)) {
                $ComputerName = Get-Content $ComputerName
            }
            ## A. Separates any comma-separated strings into an array, otherwise just creates array
            ## B. Then, cycles through the array to process each hostname/hostname substring using LDAP query
            else {
                ## A.
                if ($ComputerName -like "*,*") {
                    $ComputerName = $ComputerName -split ','
                }
                else {
                    $ComputerName = @($ComputerName)
                }
        
                ## B. LDAP query each TargetComputer item, create new list / sets back to Targetcomputer when done.
                $NewTargetComputer = [System.Collections.Arraylist]::new()
                foreach ($computer in $ComputerName) {
                    ## CREDITS FOR The code this was adapted from: https://intunedrivemapping.azurewebsites.net/DriveMapping
                    if ([string]::IsNullOrEmpty($env:USERDNSDOMAIN) -and [string]::IsNullOrEmpty($searchRoot)) {
                        Write-Error "LDAP query `$env:USERDNSDOMAIN is not available!"
                        Write-Warning "You can override your AD Domain in the `$overrideUserDnsDomain variable"
                    }
                    else {
        
                        # if no domain specified fallback to PowerShell environment variable
                        if ([string]::IsNullOrEmpty($searchRoot)) {
                            $searchRoot = $env:USERDNSDOMAIN
                        }
                        
                        $matching_hostnames = (([adsisearcher]"(&(objectCategory=Computer)(name=$computer*))").findall()).properties
                        $matching_hostnames = $matching_hostnames.name
                        $NewTargetComputer += $matching_hostnames
                    }
                }
                $ComputerName = $NewTargetComputer
            }
            $ComputerName = $ComputerName | Where-object { $_ -ne $null } | Select -Unique
            # Safety catch
            if ($null -eq $ComputerName) {
                return
            }
        }

        $DisableWifiAdapter = $DisableWifiAdapter.ToLower()
        ## 2. Turn off / disable wifi adapter scriptblock
        $turnoff_wifi_adapter_scriptblock = {
            param(
                $DisableWifi
            )
            $EthStatus = (Get-Netadapter | where-object { $_.Name -eq 'Ethernet' }).status
            if ($ethstatus -eq 'Up') {
                Write-Host "[$env:COMPUTERNAME] :: Eth is up, turning off Wi-Fi..." -foregroundcolor green
                Set-NetAdapterAdvancedProperty -Name "Wi-Fi" -AllProperties -RegistryKeyword "SoftwareRadioOff" -RegistryValue "1"
                # should these be uncommented?
                if ($DisableWifi -eq 'y') {
                    Disable-NetAdapterPowerManagement -Name "Wi-Fi"
                    Disable-NetAdapter -Name "Wi-Fi" -Force
                }            
            }
            else {
                Write-Host "[$env:COMPUTERNAME] :: Eth is down, leaving Wi-Fi alone..." -foregroundcolor red
            }
        }
    } 
    
    ## Test connection to target machine(s) and then run scriptblock to disable wifi adapter if ethernet adapter is active
    PROCESS {
        ForEach ($single_computer in $ComputerName) {

            ## empty Targetcomputer values will cause errors to display during test-connection / rest of code
            if ($single_computer) {
                ## Ping test
                $ping_result = Test-Connection $single_computer -count 1 -Quiet
                if ($ping_result) {
                    if ($single_computer -eq '127.0.0.1') {
                        $single_computer = $env:COMPUTERNAME
                    }
        
                    Invoke-Command -ComputerName $single_computer -Scriptblock $turnoff_wifi_adapter_scriptblock -ArgumentList $DisableWifiAdapter
                }
                else {
                    Write-Host "[$env:COMPUTERNAME] :: $single_computer is offline, skipping." -Foregroundcolor Yellow
                }
            }
        }
    }

    ## Pause before continuing back to terminal menu
    END {

        if (-not $env:PSMENU_DIR) {
            $env:PSMENU_DIR = pwd
        }
        ## create simple output path to reports directory
        $thedate = Get-Date -Format 'yyyy-MM-dd'
        $DIRECTORY_NAME = 'StopWifi'
        $OUTPUT_FILENAME = 'StopWifi'
        if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
            New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
        }
        
        $counter = 0
        do {
            $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
        } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))


        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Stop-WifiAdapters function executed on computers listed below." | Out-File -FilePath $output_filepath -Append -Force
        $ComputerName | Out-File -FilePath $output_filepath -Append -Force
        
        Invoke-Item "$output_filepath"

        # Read-Host "`nPress [ENTER] to continue."
    }
}

## export all functions except GetTargets and GetOutputFileString
Export-ModuleMember -Function *-*
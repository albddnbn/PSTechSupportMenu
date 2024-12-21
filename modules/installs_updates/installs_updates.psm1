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

function Add-PrinterLogicPrinter {
    <#
    .SYNOPSIS
        Connects local or remote computer to target printerlogic printer by executing C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin\PrinterInstallerConsole.exe.
        Connection fails if PrinterLogic Client software is not installed on target machine(s).
        The user does NOT have to be a PrinterLogic user to be able to access connected PrinterLogic printers.

    .DESCRIPTION
        PrinterLogic Client software has to be installed on target machine(s) and connecting to your organization's 'Printercloud' instance using the registration key.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PrinterName
        Name of the printer in printerlogic. Ex: 't-prt-lib-01', you can use the name or the full 'path' to the printer, ex: 'STANTON\B WING..\s-prt-b220-01'
        Name must match the exact 'name' of the printer, as listed in PrinterLogic.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, printer name, connection status, and client software status.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Connect single remote target computer to t-prt-lib-01 printer:
        Add-PrinterLogicPrinter -TargetComputer "t-client-28" -PrinterName "t-prt-lib-01"
        
    .EXAMPLE
        Connect a group of computers using hostname txt filepath to t-prt-lib-02 printer:
        Add-PrinterLogicPrinter -TargetComputer "D:\computers.txt" -PrinterName "t-prt-lib-02"

    .EXAMPLE
        Submit target computer(s) by hostname through pipeline and add specified printerlogic printer.
        @('t-client-01', 't-client-02', 't-client-03') | Add-PrinterLogicPrinter -PrinterName "t-prt-lib-03"
        
    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param (
        [Parameter(
            Mandatory = $true, ValueFromPipeline = $true
        )]
        [String[]]$ComputerName,
        [string]$PrinterName,
        $SendPings
    )
    ## 1. TargetComputer handling if not supplied through pipeline
    ## 2. Scriptblock to connect to printerlogic printer
    ## 3. Results containers for overall results and skipped computers.
    $ComputerName = GetTargets -TargetComputer $ComputerName

    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    $thedate = Get-Date -Format 'yyyy-MM-dd'

        
    ## 2. Define the scriptblock that connects machine to target printer in printerlogic cloud instance
    $connect_to_printer_block = {
        param(
            $printer_name
        )

        $obj = [pscustomobject]@{
            hostname       = $env:COMPUTERNAME
            printer        = $printer_name
            connectstatus  = 'NO'
            clientsoftware = 'NO'
        }
        # get installerconsole.exe
        $exepath = get-childitem -path "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin" -Filter "PrinterInstallerConsole.exe" -File -Erroraction SilentlyContinue
        if (-not $exepath) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: PrinterLogic PrinterInstallerConsole.exe was not found in C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin." -Foregroundcolor Red
            return $obj
        }
        
        $obj.clientsoftware = 'YES'
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Found $($exepath.fullname), mapping $printer_name now..."
        $map_result = (Start-Process "$($exepath.fullname)" -Argumentlist "InstallPrinter=$printer_name" -Wait -Passthru).ExitCode
        
        # 0 = good, 1 = bad
        if ($map_result -eq 0) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Connected to $printer_name successfully." -Foregroundcolor Green
            # Write-Host "*Remember that this script does not set default printer, user has to do that themselves."
            $obj.connectstatus = 'YES'
            return $obj
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: failed to connect to $printer_name." -Foregroundcolor Red
            return $obj
        }
    }


    $results = Invoke-Command -ComputerName $single_computer -scriptblock $connect_to_printer_block -ArgumentList $PrinterName -ErrorVariable RemoteError 

    ## errored out invoke-commands:
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, run scriptblock to attempt connection to printerlogic printer, save results to object

    ## 1. If there are any results - output to file
    ## 2. Output missed computers to terminal.
    ## 3. Return results arraylist
    ## 1.
    if ($results) {

        $results | out-gridview -Title "PrinterLogic Connect Results"

    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output." | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append
        "These computers did not respond to ping:" | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append
        $errored_machines | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append

        Invoke-Item "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt"
    }
    ## 2. Output unresponsive computers
    Write-Host "These computers did not respond to ping:"
    Write-Host ""
    $errored_machines
    Write-Host ""
    # read-host "Press enter to return results."

    ## 3. return results arraylist
    return $results
}

Function Apply-DellCommandUpdates {
    <#
    .SYNOPSIS
        Looks for & executes the dcu-cli.exe on target machine(s) using the '/applyUpdates -reboot=enable' switches.
        Function will automatically pass over currently occupied computers.
        BIOS updates will not be applied unless BIOS password is added to dcu-cli.exe command in some form.

    .DESCRIPTION
        End of function prints 3 lists to terminal:
            1. Computers with users logged in, and were skipped.
            2. Computers without Dell Command | Update installed, and were skipped.
            3. Computers that have begun to apply updates and should be rebooting momentarily.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER WeeklyUpdates
        'y'     will use this command: &"$($dellcommandexe.fullname)" /configure -scheduleWeekly=Sun,02:00
                to set weekly updates on Sundays at 2am.           
        'n'     will not set weekly update times, leaving update schedule as default (every 3 days).


    .NOTES
        Executing like this: Start-Process "$($dellcommandexe.fullname)" -argumentlist $params doesn't show output.
        Trying $result = (Start-process file -argumentlist $params -wait -passthru).exitcode -> this still rebooted s-c136-03
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        $WeeklyUpdates
    )
    ## 1. Handling TargetComputer input if not supplied through pipeline.
    ## 2. Create empty results container.
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, check for dell command update executable and apply updates, report on results.
    $results = Invoke-Command -ComputerName $single_computer -scriptblock {

        $weeklyupdates_setting = $using:WeeklyUpdates
        $LoggedInUser = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username

        $obj = [pscustomobject]@{
            DCUInstalled   = "NO"
            UpdatesStarted = "NO"
            # DCUExitCode    = ""
            Username       = $LoggedInUser
        }
        if ($obj.UserName) {
            Write-Host "[$env:COMPUTERNAME] :: $LoggedInUser is logged in, skipping computer." -Foregroundcolor Yellow
        }
        # $apply_updates_process_started = $false
        # find Dell Command Update executable
        $dellcommandexe = Get-ChildItem -Path "${env:ProgramFiles(x86)}" -Filter "dcu-cli.exe" -File -Recurse # -ErrorAction SilentlyContinue
        if (-not ($dellcommandexe)) {
            Write-Host "[$env:COMPUTERNAME] :: NO Dell Command | Update found." -Foregroundcolor Red
        }
        else {
            $obj.DCUInstalled = "YES"
        }
        if (($obj.Username) -or (-not $dellcommandexe)) {
            $obj
            break
        }

        Write-Host "[$env:COMPUTERNAME] :: Found $($dellcommandexe.fullname), executing with the /applyupdates -reboot=enable parameters." -ForegroundColor Yellow
        # NOTE: abuddenb - Haven't been able to get the -reboot=disable switch to work yet.
        Write-Host "[$env:COMPUTERNAME] :: Configuring weekly update schedule."
        if ($weeklyupdates_setting.ToLower() -eq 'y') {
            &"$($dellcommandexe.fullname)" /configure -scheduleWeekly='Sun,02:00' -silent
        }
        Write-Host "[$env:COMPUTERNAME] :: Applying updates now." -ForegroundColor Yellow
        &$($dellcommandexe.fullname) /applyUpdates -reboot=enable -silent

        $obj.UpdatesStarted = $true

        # return results of a command or any other type of object, so it will be addded to the $results list
        $obj
    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunSpaceId, PSShowComputerName # filters out some properties that don't seem necessary for these functions

    $errored_machines = $RemoteError.CategoryInfo.TargetName


    ## 1. Output results to gridview and terminal
    if ($results) {
        $results | Out-GridView - "Dell Command | Update Results"
        $occupied_computers = $($results | where-object { $_.Username } | select-object -expandproperty PSComputerName) -join ', '
        $need_dell_command_update_installed = $($results | where-object { $_.DCUInstalled -eq 'NO' } | select-object -expandproperty PSComputerName) -join ', '
    
        $currently_applying_updates = $($results | where-object { $_.UpdatesStarted -eq "YES" } | select-object -expandproperty PSComputerName) -join ', '
        Write-Host "These computers are occupied: " -nonewline
        Write-Host "$occupied_computers" -Foregroundcolor Yellow
        Write-Host "`nThese computers need Dell Command | Update installed: " -nonewline
        Write-Host "$need_dell_command_update_installed" -Foregroundcolor Red
        Write-Host "`nThese computers have begun to apply updates and should be rebooting momentarily: " -nonewline
        Write-Host "$currently_applying_updates" -Foregroundcolor Green
        Write-Host "`nThese computers are unresponsive:"
        $errored_machines
    }
    return $results
}

function Install-ActiveDirectoryModule {
    <#
    .SYNOPSIS
        Installs Powershell Active Directory module without installing RSAT or using DISM.

    .DESCRIPTION
        After installing the Active Directory module on a computer the traditional way, follow the guide in the notes to find the necessary module files.

    .EXAMPLE
        Get-ActiveDirectoryModule

    .NOTES
        Get the .dll file you have to import for AD module:
        https://petertheautomator.com/2020/10/05/use-the-active-directory-powershell-module-without-installing-rsat/

        abuddenb / 02-17-2024
    #>
    $activedirectory_folder = Get-Childitem -path "$env:SUPPORTFILES_DIR" -filter "ActiveDirectory" -Directory -ErrorAction SilentlyContinue
    if (-not $activedirectory_folder) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ActiveDirectory folder not found in $env:SUPPORTFILES_DIR" -foregroundcolor red
        return
    }

    Copy-Item -Path "$($activedirectory_folder.fullname)" -Destination "C:\Program Files\WindowsPowerShell\Modules\" -Recurse -Force -ErrorAction SilentlyContinue

    $Microsoft_AD_Management_DLL = Get-childitem -path "C:\Program Files\WindowsPowerShell\Modules\" -Filter 'Microsoft.ActiveDirectory.Management.dll' -File -Recurse -Erroraction SilentlyContinue
    if (-not $Microsoft_AD_Management_DLL) {
        Write-Host "Unable to find Microsoft.ActiveDirectory.Management.dll in ./ActiveDirectory. Exiting script." -foregroundcolor red
        exit
    }
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($Microsoft_AD_Management_DLL.fullname). Importing ActiveDirectory module."
    Import-Module "$($Microsoft_AD_Management_DLL.fullname)" -Force

}

function Install-Application {
    <#
	.SYNOPSIS
        Uses $env:PSMENU_DIR/deploy/applications folder to present menu to user. 
        Multiple app selections can be made, and then installed sequentially on target machine(s).
        Application folders should be PSADT folders, this function uses the traditional PSADT silent installation line to execute.
            Ex: For the Notepad++ application, the folder name is 'Notepad++', and installation script is 'Deploy-Notepad++.ps1'.

	.DESCRIPTION
        You can find some pre-made PSADT installation folders here:
        https://dtccedu-my.sharepoint.com/:f:/g/personal/abuddenb_dtcc_edu/Ervb5x-KkbdHvVcCBb9SK5kBCINk2Jtuvh240abVnpsS_A?e=kRsjKx
        Applications in the 'working' folder have been tested and are working for the most part.

	.PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).
        
    .PARAMETER AppName
        If supplied, the function will look for a folder in $env:PSMENU_DIR\deploy\applications with a name that = $AppName.
        If not supplied, the function will present menu of all folders in $env:PSMENU_DIR\deploy\applications to user.
    
    .PARAMETER SkipOccupied
        If anything other than 'n' is supplied, the function will skip over computers that have users logged in.
        If 'n' is supplied, the function will install the application(s) regardless of users logged in.

	.EXAMPLE
        Run installation(s) on all hostnames starting with 's-a231-':
		Install-Application -TargetComputer 's-a231-'

    .EXAMPLE
        Run installation(s) on local computer:
        Install-Application

    .EXAMPLE
        Install Chrome on all hostnames starting with 's-c137-'.
        Install-Application -Targetcomputer 's-c137-' -AppName 'Chrome'

	.NOTES
        PSADT Folders: https://dtccedu-my.sharepoint.com/:f:/g/personal/abuddenb_dtcc_edu/Ervb5x-KkbdHvVcCBb9SK5kBCINk2Jtuvh240abVnpsS_A?e=kRsjKx
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
                if (Test-Path "$env:PSMENU_DIR\deploy\applications\$_" -ErrorAction SilentlyContinue) {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($_) in $env:PSMENU_DIR\deploy\applications." -Foregroundcolor Green
                    return $true
                }
                else {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $($_) not found in $env:PSMENU_DIR\deploy\applications." -Foregroundcolor Red
                    return $false
                }
            })]
        [string]$AppName,
        [string]$SkipOccupied
    )
    ## 1. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 2. If AppName parameter was not supplied, apps chosen through menu will be installed on target machine(s).
    ##    - menu presented uses the 'PS-Menu' module: https://github.com/chrisseroka/ps-menu
    ## 3. Define scriptblock - installs specified app using PSADT folder/script on local machine.
    ## 4. Prompt - should this script skip over computers that have users logged in?
    ## 5. create empty containers for reports:
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $AllComputers = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    $ComputerName = $AllComputers | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }
    
    $unresponsive_computers = $AllComputers | Where-Object { $ComputerName -notcontains $_ }


    ## 2. If AppName parameter was not supplied, apps chosen through menu will be installed on target machine(s).
    if (-not $appName) {
        # present script/ADMIN with a list of apps to choose from
        $ApplicationList = Get-ChildItem -Path "$env:PSMENU_DIR\deploy\applications" -Directory | Where-object { $_.name -notin @('Veyon', 'SMARTNotebook') }
        Clear-Host
        $applist = ($ApplicationList).Name

        # divide applist into app_list_one and app_list_two
        ## THIS HAS TO BE DONE (AT LEAST WITH THIS PS-MENU MODULE - ANYTHING OVER ~30 ITEMS WILL CAUSE THE MENU TO FREEZE UP)
        $app_list_one = [system.collections.arraylist]::new()
        $app_list_two = [system.collections.arraylist]::new()
        $app_name_counter = 0
        ForEach ($single_app_name in $applist) {
            if ($app_name_counter -lt 29) {
                $app_list_one.Add($single_app_name) | Out-Null
            }
            else {
                $app_list_two.Add($single_app_name) | Out-Null
            }
            $app_name_counter++
        }
        # First 29 apps|
        Write-Host "Displaying the first 29 applications available:" -ForegroundColor Yellow
        $chosen_apps_one = Menu $app_list_one -MultiSelect
        Write-Host "More applications:" -ForegroundColor Yellow
        # Last 29 apps
        $chosen_apps_two = Menu $app_list_two -MultiSelect

        $chosen_apps = @()
        ForEach ($applist in @($chosen_apps_one, $chosen_apps_two)) {
            $appnames = $applist -split ' '
            $chosen_apps += $appnames
        }
    }
    elseif ($AppName) {
        $chosen_apps = $AppName -split ','

        if ($chosen_apps -isnot [array]) {
            $chosen_apps = @($chosen_apps)
        }
        # validate the applist:
        ForEach ($single_app in $chosen_apps) {
            if (-not (Test-Path "$env:PSMENU_DIR\deploy\applications\$single_app")) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_app not found in $env:PSMENU_DIR\deploy\applications." -Foregroundcolor Red
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Ending function." -Foregroundcolor Red
                return
            }
        }
    }

    ## 3. Define scriptblock - installs specified app using PSADT folder/script on local machine.
    $install_local_psadt_block = {
        param(
            $app_to_install,
            $do_not_disturb
        )
        ## Remove previous psadt folders:
        # Remove-Item -Path "C:\temp\$app_to_install" -Recurse -Force -ErrorAction SilentlyContinue
        # Safety net since psadt script silent installs close app-related processes w/o prompting user
        $check_for_user = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
        if ($check_for_user) {
            if ($($do_not_disturb) -eq 'y') {
                Write-Host "[$env:COMPUTERNAME] :: Skipping, $check_for_user logged in."
                Continue
            }
        }
        # get the installation script
        $Installationscript = Get-ChildItem -Path "C:\temp" -Filter "Deploy-$app_to_install.ps1" -File -Recurse -ErrorAction SilentlyContinue
        # unblock files:
        Get-ChildItem -Path "C:\temp" -Recurse | Unblock-File
        # $AppFolder = Get-ChildItem -Path 'C:\temp' -Filter "$app_to_install" -Directory -Erroraction silentlycontinue
        if ($Installationscript) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Found $($Installationscript.Fullname), installing."
            Set-Location "$($Installationscript.DirectoryName)"
            Powershell.exe -ExecutionPolicy Bypass ".\Deploy-$($app_to_install).ps1" -DeploymentType "Install" -DeployMode "Silent"
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: ERROR - Couldn't find the app deployment script!" -Foregroundcolor Red
        }
    }

    Clear-Host

    ## 4. Prompt - should this script skip over computers that have users logged in?
    ##    - script runs 'silent' installation of PSADT - this means installations will likely close the app / associated processes
    ##      before uninstalling / installing. This could disturb users.
    if ($SkipOccupied.ToLower() -eq 'n') {
        $skip_pcs = 'n'
    }
    else {
        $skip_pcs = 'y'
    }
    
    ## 5. create empty containers for reports:
    ## computers that were unresponsive
    ## apps that weren't able to be installed (weren't found in deployment folder for some reason.)
    ## - If they were presented in menu / chosen, apps should definitely be in deployment folder, though.
    $skipped_applications = [system.collections.arraylist]::new()

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, cycle through chosen apps and run the local psadt install scriptblock for each one,
    ##    on each target machine.
    ##    3.1 --> Check for app/deployment folder in ./deploy/applications, move on to next installation if not found
    ##    3.2 --> Copy PSADT folder to target machine/session
    ##    3.3 --> Execute PSADT installation script on target machine/session
    ##    3.4 --> Cleanup PSADT folder in C:\temp

    ## create sesion
    $TargetSessions = New-PSSession $ComputerName
    ## 3. Install chosen apps by creating remote session and cycling through list
    ForEach ($single_application in $chosen_apps) {

        $DeploymentFolder = Get-ChildItem -Path "$env:PSMENU_DIR\deploy\applications\" -Filter "$single_application" -Directory -ErrorAction SilentlyContinue
        if (-not $DeploymentFolder) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_application not found in $env:PSMENU_DIR\deploy\applications." -Foregroundcolor Red
            $skipped_applications.Add($single_application) | Out-Null
            Continue
        }

        ## Make sure there isn't an existing deployment folder on target machine:
        Invoke-Command -Session $TargetSessions -scriptblock {
            Remove-Item -Path "C:\temp\$($using:single_application)" -Recurse -Force -ErrorAction SilentlyContinue
        }

        ## 3.2 Copy PSADT folder to target machine/session
        $copy_application_error = $null
        $ComputerName | % { Copy-Item -Path "$($DeploymentFolder.fullname)" -Destination "\\$_\c$\temp\" -Recurse -ErrorAction Continue }

        # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $($DeploymentFolder.name) copied to $single_computer."
        ## 3.3 Execute PSADT installation script on target mach8ine/session
        Invoke-Command -Session $TargetSessions -scriptblock $install_local_psadt_block -ArgumentList $single_application, $skip_pcs
        # Start-Sleep -Seconds 1
        ## 3.4 Cleanup PSADT folder in temp
        $folder_to_delete = "C:\temp\$($DeploymentFolder.Name)"
        Invoke-Command -Session $TargetSessions -command {
            Remove-Item -Path "$($using:folder_to_delete)" -Recurse -Force -ErrorAction SilentlyContinue
        }

        # $installation_completed.add($single_computer) | out-null
    }
    ## 1. Open the folder that will contain reports if necessary.

    ## installation COMPLETED list - not necessarily completed successfully. just to help with tracking / reporting.
    ## outfile to temp to store unresponsive computers
    $unresponsive_computers | Out-File "$env:PSMENU_DIR\reports\$thedate\unresponsive-computers-$thedate.txt" -Force

    # if ($unresponsive_computers) {
    #     Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Unresponsive computers:" -Foregroundcolor Yellow
    #     $unresponsive_computers | Sort-Object
    # }
    # if ($skipped_applications) {
    #     Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Skipped applications:" -Foregroundcolor Yellow
    #     $skipped_applications | Sort-Object
    # }

    "Function completed execution on: [$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')]" | Out-File "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" -append -Force
    "Installation process completed (not necessarily successfully) on the following computers:" | Out-File "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" -append -Force

    $ComputerName | Sort-Object | Out-File "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" -append -Force

    "Unresponsive computers:" | Out-File "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" -append -Force

    $unresponsive_computers | Sort-Object | Out-File "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" -append -Force

    Invoke-Item "$env:PSMENU_DIR\reports\$thedate\install-application-$thedate.txt" 
    
}

Function Install-NewTeams {
    <#
    .SYNOPSIS
        Installs the 'new' Microsoft Teams (work or school account) client on target computers using the 'network-required' method (only teamsbootstrapper.exe, not .msix).

    .DESCRIPTION
        Doesn't create desktop icon like Teams Classic - users will have to go to Start Menu and search for Teams. They should see a Teams app listed that uses the new icon.
        Sometimes it's hidden toward the bottom of the results in Start Menu search.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .PARAMETER skip_occupied_computers
        This script KILLS ANY RUNNING TEAMS PROCESSES.
        'y' will skip computers that have users logged in, 'n' will not skip them.
        Default is 'y'.

    .EXAMPLE
        Install new teams client and provision for future users - all AD computers with hostnames starting with pc-a227-
        Install-NewTeams -TargetComputer "pc-a227-"

    .NOTES
        The script used for actual installation of the new Teams client was created by:
        Author:     Sassan Fanai
        Date:       2023-11-22
        *need to get a link to script on github to put here*
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$skip_occupied_computers = 'y',
        $SendPings
    )
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    # $skip_occupied_computers = Read-Host "Install of Teams will stop any running teams processes on target machines - skip computers that have users logged in? [y/n]"
    # get newteams folder from irregular applications
    $NewTeamsFolder = Get-ChildItem -Path "$env:PSMENU_DIR\deploy\irregular" -Filter 'NewTeams' -Directory -ErrorAction SilentlyContinue
    if (-not $NewTeamsFolder) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: NewTeams folder not found in $env:PSMENU_DIR\deploy\irregular" -foregroundcolor red
        return
    }
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($NewTeamsFolder.fullname) folder."
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Beginning installation/provisioning of 'New Teams' on $($ComputerName -join ', ')"
    
    ## Create empty collections list for skipped target machines
    $missed_computers = [system.collections.arraylist]::new()
    ## *Remove teams classic and install/provision new teams client on machines*
    ## 1. Check $ComputerName for $null / '' values
    ## 2. Send one test ping
    ## 3. Create PSSession on on target machine
    ## 4. Copy the New Teams folder to target, using the session
    ## 5. Execute script - removes Teams classic / installs & provisions new Teams client
    ForEach ($single_computer in $ComputerName) {

        ## 1. Check $single_computer for $null / '' values
        if ($single_computer) {
            ## 2. Send one test ping
            $ping_result = Test-Connection $single_computer -count 1 -Quiet
            if ($ping_result) {
                if ($single_computer -eq '127.0.0.1') {
                    $single_computer = $env:COMPUTERNAME
                }
                ## 3. Create PSSession on on target machine
                $single_target_session = New-PSSession $single_computer
                ## 4. Copy the New Teams folder to target, using the session
                try {
                    Copy-Item -Path "$($NewTeamsFolder.fullname)" -Destination 'C:\temp' -ToSession $single_target_session -Recurse -Force
                }
                catch {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Error copying NewTeams folder to $single_computer" -foregroundcolor red
                    continue
                }

                ## 5. Execute script - removes Teams classic / installs & provisions new Teams client
                ##    - Skips computer if skip_occupied_computers is 'y' and user is logged in
                $new_teams_install_result = Invoke-Command -ComputerName $single_computer -scriptblock {

                    $obj = [pscustomobject]@{
                        UserLoggedIn = ''
                        Skipped      = 'no'
                    }
                    $check_for_user = $using:skip_occupied_computers
                    if ($check_for_user.ToLower() -eq 'y') {
                        $user_logged_in = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
                        if ($user_logged_in) {
                            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -NoNewline
                            Write-Host "[$env:COMPUTERNAME] :: Found $user_logged_in logged in, skipping installation of Teams." -foregroundcolor yellow
                            $obj.UserLoggedIn = $user_logged_in
                            $obj.Skipped = 'yes'
                            return $obj                            
                        }
                    }
                    $installteamsscript = get-childitem -path 'C:\temp\newteams' -filter "install-msteams.ps1" -file -erroraction SilentlyContinue
                    if (-not $installteamsscript) {
                        Write-Host "No install-msteams.ps1 on $env:computername" -foregroundcolor red
                        $obj.Skipped = 'yes'
                        return $obj
                    }
                    # execute script (NewTeams installation folder can be found in ./deploy/irregular)
                    $install_params = @{
                        logfile      = "C:\WINDOWS\Logs\Software\NewTeamsClient-Install-$(Get-Date -Format 'yyyy-MM-dd').log"
                        forceinstall = $true
                        setrunonce   = $true
                    }
                    ## If the msix file is present - run offline install
                    if (Test-Path -Path 'C:\temp\newteams\Teams_windows_x64.msi') {
                        $install_params['offline'] = $true
                    }
                    &"$($installteamsscript.fullname)" @install_params
                }

                ## Check if the computer was skipped
                if ($new_teams_install_result.Skipped -eq 'yes') {
                    $missed_computers.add($single_computer) | out-null
                    continue
                }
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is not reachable, skipping." -foregroundcolor red
                $missed_computers.add($single_computer) | out-null
                continue
            }
        }
    }
    ## 1. Output the missed computers list to a text file in reports dir / new teams
    ## 2. Function complete - read-host to allow pause for debugging errors/etc
    if (-not $env:PSMENU_DIR) {
        $env:PSMENU_DIR = pwd
    }
    $DIRECTORY_NAME = 'NewTeams'
    $OUTPUT_FILENAME = 'newteamsresults'
    if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
    }
        
    $counter = 0
    do {
        $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
    } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: New Teams not installed on these computers." | Out-File -Append -FilePath $output_filepath
    $missed_computers | Out-File -Append -FilePath $output_filepath

    "`n" | Out-File -Append -FilePath $output_filepath

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Complete list of target computers below." | Out-File -Append -FilePath $output_filepath
    $ComputerName | Out-File -Append -FilePath $output_filepath

    Invoke-Item $output_filepath
        
    # read-host "Press enter to continue."
}

function Install-SMARTNotebookSoftware {
    <#
    .SYNOPSIS
        Installs SMART Learning Suite software on target computers.
        Info: https://www.smarttech.com/en/education/products/software/smart-notebook

    .DESCRIPTION
        May also be able to use a hostname file eventually.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Install-SMARTNotebookSoftware -TargetComputer "s-c136-02"

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
        $SendPings
    )
    ## 1. Handling TargetComputer input if not supplied through pipeline.
    ## 2. Make sure SMARTNotebook folder is in ./deploy/irregular
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    ## 2. Get the smartnotebook folder from irregular applications
    $SmartNotebookFolder = Get-ChildItem -path "$env:PSMENU_DIR\deploy\irregular" -Filter 'SMARTNotebook' -Directory -Erroraction SilentlyContinue
    if ($SmartNotebookFolder) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($SmartNotebookFolder.FullName), copying to target computers." -foregroundcolor green
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: SMARTNotebook folder not found in irregular applications, exiting." -foregroundcolor red
        exit
    }

    ## For each target computer - assign installation method - either office or classroom. Classroom installs the smartboard and ink drivers.
    Write-Host "Please choose installation method for target computers:"
    $InstallationTypeReply = Menu @('Office', 'Classroom')

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, find PSADT Folder and install SMARTNotebook software.
    Invoke-Command -ComputerName $ComputerName -Scriptblock {
        $installation_method = $using:InstallationTypeReply
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME]:: Installation method set to: $installation_method"
        # unblock files
        Get-ChildItem -Path "C:\TEMP\SMARTNotebook" -Recurse | Unblock-File
        # get Deploy-SMARTNotebook.ps1
        $DeployScript = Get-ChildItem -Path "C:\TEMP\SMARTNotebook" -Filter 'Deploy-SMARTNotebook.ps1' -File -ErrorAction SilentlyContinue
        if ($DeployScript) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($DeployScript.FullName), executing." -foregroundcolor green
            Powershell.exe -ExecutionPolicy Bypass "$($DeployScript.FullName)" -DeploymentType "Install" -DeployMode "Silent" -InstallationType "$installation_method"
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Deploy-SMARTNotebook.ps1 not found, exiting." -foregroundcolor red
            exit
        }
    } -ErrorVariable RemoteError

    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## create file to announce completion for when being run as background job
    if (-not $env:PSMENU_DIR) {
        $env:PSMENU_DIR = pwd
    }
    ## create simple output path to reports directory
    $thedate = Get-Date -Format 'yyyy-MM-dd'
    $DIRECTORY_NAME = 'SMARTNotebook'
    $OUTPUT_FILENAME = 'SMARTInstall'
    if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
    }
        
    $counter = 0
    do {
        $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter"
    } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))

    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Finished SMART Learning Suite installation(s)." | Out-File -FilePath $output_filepath -Append
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Installation method: $InstallationTypeReply" | Out-File -FilePath $output_filepath -Append
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Target computer(s):" | Out-File -FilePath $output_filepath -Append
    $ComputerName | Out-File -FilePath "$output_filepath.txt" -Append
    $errored_machines | Out-File -FilePath "$output_filepath-errors.txt" -Append
}

function Install-VeyonRoom {
    <#
    .SYNOPSIS
        Uses the Veyon PS App Deployment Toolkit folder in 'deploy\irregular' to install Veyon on target computers, then creates a script to run on the master computer to create the room list of PCs you can view.

    .DESCRIPTION
        Specify the master computer using the parameter, the /NoMaster installation switch is used to install on all other computers.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER RoomName
        The name of the room to be used in the Veyon room list - only used in the script that's output, not in actual installation.

    .PARAMETER Master_Computer
        The name of the master computer, Veyon installation is run without the /NoMaster switch on this computer.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

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
        [string]$RoomName,
        # comma-separated list of computers that get Veyon master installation.
        [string]$Master_Computer,
        $SendPings
    )
    ## 1. Handling TargetComputer input if not supplied through pipeline.
    ## 2. Create list of master computers from $Master_computer parameter
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    ## 2. Creating list of master computers, trim whitespace
    $master_computer_selection = $($Master_Computers -split ',')
    ## ^ Need to verify master computers are in computername (w/any suffixes)
    $master_computer_selection | % { $_ = $_.Trim() }

    if (-not $RoomName) {
        $RoomName = Read-Host "Please enter the name of the room/location: "
    }

    # copy the ps app deployment folders over to temp dir
    # get veyon directory from irregular applications:
    $VeyonDeploymentFolder = Get-ChildItem "$env:PSMENU_DIR\deploy\irregular" -Filter "Veyon" -Directory -ErrorAction SilentlyContinue
    if (-not $VeyonDeploymentFolder) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Couldn't find Veyon deployment folder in $env:PSMENU_DIR\deploy\irregular\Veyon\Veyon, exiting." -foregroundcolor red
        return
    }

    $install_veyon_scriptblock = {
        param(
            $MasterInstall,
            $VeyonDirectory,
            $DeleteFolder
        )
        Get-ChildItem $VeyonDirectory -Recurse | Unblock-File
        Set-Location $VeyonDirectory
        $DeployVeyonPs1 = Get-ChildItem -Path $VeyonDirectory -Filter "*Deploy-Veyon.ps1" -File -Recurse -ErrorAction SilentlyContinue
        if (-not $DeployVeyonPs1) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Couldn't find Deploy-Veyon.ps1 in $VeyonDirectory, exiting." -foregroundcolor red
            return
        }
        if ($MasterInstall -eq 'y') {
            Powershell.exe -ExecutionPolicy Bypass "$($DeployVeyonPs1.Fullname)" -DeploymentType "Install" -DeployMode "Silent" -MasterPC
        }
        else {
            Powershell.exe -ExecutionPolicy Bypass "$($DeployVeyonPs1.Fullname)" -DeploymentType "Install" -DeployMode "Silent"
        }
        if ($DeleteFolder -eq 'y') {
            Remove-Item -Path $VeyonDirectory -Recurse -Force
        }
    }

    ## MASTER INSTALLATIONS:
    Write-Host "`n[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Beginning MASTER INSTALLATIONS."
    Write-Host "[1] Veyon Master, [2] Veyon Student, [3] No Veyon"
    $reply = Read-Host "Would you like your local computer to be installed with Veyon?"
    if ($reply -eq '1') {
        Invoke-Command  -ScriptBlock $install_veyon_scriptblock -ArgumentList 'y', $VeyonDeploymentFolder.FullName, 'n'
        $master_computer_selection += @($env:COMPUTERNAME)
    }
    elseif ($reply -eq '2') {
        Invoke-Command  -ScriptBlock $install_veyon_scriptblock -ArgumentList 'n', $VeyonDeploymentFolder.FullName, 'n'
    }

    ## Missed computers container
    $missed_computers = [system.collections.arraylist]::new()
    ## Student computers list:
    $Student_Computers = [system.collections.arraylist]::new()
    ## 1. check targetcomputer for null / empty values
    ## 2. ping test target machine
    ## 3. If responsive - copy Veyon install folder over to session, and install Veyon student or master
    ## 4. Add any missed computers to a list, also add student installations to a list for end output.
    ForEach ($single_computer in $ComputerName) {

        # If there aren't any master computers yet - ask if this one should be master
        if (-not $master_computer_selection) {
            $reply = Read-Host "Would you like to make $single_computer a master computer? [y/n]"
            if ($reply.tolower() -eq 'y') {
                $master_computer_selection = @($single_computer)
            }
        }

        ## 1.
        if ($single_computer) {

            ## 2. Test with ping
            $pingreply = Test-Connection $single_computer -Count 1 -Quiet
            if ($pingreply) {
                ## Create Session
                $target_session = New-PSSession -ComputerName $single_computer
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is responsive to ping, proceeding." -foregroundcolor green
                ## Remove any existing veyon folder
                Invoke-Command -Session $target_session -Scriptblock {
                    Remove-Item -Path "C:\temp\Veyon" -Recurse -Force -ErrorAction SilentlyContinue
                }
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Copying Veyon source files to $single_computer." -foregroundcolor green
                ## 3. Copy source files
                Copy-Item -Path "$($VeyonDeploymentFolder.fullname)" -Destination C:\temp\ -ToSession $target_session -Recurse -Force

                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Installing Veyon on $single_computer." -foregroundcolor green
                ## If its a master computer:
                if ($single_computer -in $master_computer_selection) {
                    Invoke-Command -Session $target_session -ScriptBlock $install_veyon_scriptblock -ArgumentList 'y', 'C:\Temp\Veyon', 'y'
                }
                else {
                    Invoke-Command -Session $target_session -ScriptBlock $install_veyon_scriptblock -ArgumentList 'n', 'C:\Temp\Veyon', 'y'
                    $Student_Computers.Add($single_computer) | Out-Null
                }
            }
            else {
                ## 4. Missed list is below, student list = $student_computers
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is not responding to ping, skipping." -foregroundcolor red
                $missed_computers.Add($single_computer) | Out-null
            }
        }
    }
    ## 1. Create script that needs to be run on master computers, to add client computers & make them visible.
    ## 2. Check if user wants to run master script on local computer.
    ## 3. Open an RDP window, targeting the first computer in the master computers list - RDP or some type of graphical
    ##    login is necessary to add clients.
    ## 1. Create script that needs to be run on master computers, to add client computers & make them visible.
    $scriptstring = @"
`$Student_Computers = @('$($Student_Computers -join "', '")')
`$RoomName = `'$RoomName`'
`$veyon = "C:\Program Files\Veyon\veyon-cli"
&`$veyon networkobjects add location `$RoomName
ForEach (`$single_computer in `$Student_Computers) {		
    Write-Host "[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ::  Adding student computer: `$single_computer."

    If ( Test-Connection -BufferSize 32 -Count 1 -ComputerName `$single_computer -Quiet ) {
        Start-Sleep -m 300
        `$IPAddress = (Resolve-DNSName `$single_computer).IPAddress
        `$MACAddress = Invoke-Command -Computername `$single_computer -scriptblock {
            `$obj = (get-netadapter -physical | where-object {`$_.name -eq 'Ethernet'}).MAcaddress
            `$obj
        }
        Write-Host " `$veyon networkobjects add computer `$single_computer `$IPAddress `$MACAddress `$RoomName "
        &`$veyon networkobjects add computer `$single_computer `$IPADDRESS `$MACAddress `$RoomName
    }
    Else {
        Write-Host "Didn't add `$single_computer because it's offline." -foregroundcolor Red
    }
}
"@
    # create the script that needs to be run on the master computer while RDP'd in (invoke-command is generating errors)
    $scriptfilename = "$RoomName-RunWhileLoggedIntoMaster.ps1"
    New-Item -Path "$env:PSMENU_DIR\output\$scriptfilename" -ItemType "file" -Value $scriptstring -Force | out-null

    $reply = Read-Host "Veyon room build script created, execute script to add student computers on local computer? [y/n]"
    if ($reply.ToLower() -eq 'y') {
        Get-Item "$env:PSMENU_DIR\output\$scriptfilename" | Unblock-File
        try {
            & "$env:PSMENU_DIR\output\$scriptfilename"
        }
        catch {
            WRite-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -nonewline
            Write-Host "[ERROR]" -NoNewline -foregroundcolor red
            Write-Host " :: Something went wrong when trying to add student computers to local Veyon master installation. "
        }
    }

    Write-Host "Please run $scriptfilename on $($master_computer_selection -join ', ') to create the room list of PCs you can view." -ForegroundColor Green
    Write-Host ""
    Write-host "These student computers were skipped because they're unresponsive:"
    Write-host "$($missed_computers -join ', ')" -foregroundcolor red

    Invoke-Item "$env:PSMENU_DIR\output\$scriptfilename"

    # select first master computer, open rdp window to it
    $first_master = $master_computer_selection[0]
    Open-RDP -SingleTargetComputer $first_master
}




Function Remove-TeamsClassic {
    <#
    .SYNOPSIS
        Attempts to remove any user installations of Microsoft Teams Classic, and any system installation of 'Teams Machine-Wide Installer'

    .DESCRIPTION
        The Teams Machine-Wide Installer .msi uninstallation WILL return an exit code indicating the product is not currently installed - this is expected.
        The script goes on to remove the Teams Machine-Wide Installer registry key, and then checks for any user installations of Teams Classic.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER DoNotDisturbUsers
        'y' will skip any computers that are occupied by a user.
        'n' will attempt to remove Teams Classic from all computers, including those with users logged in.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Remove Microsoft Teams Classic from all computers that have hostnames starting with 't-computer-'
        Remove-TeamsClassic -TargetComputer 't-computer-' -DoNotDisturbUsers 'y'

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
        [Parameter(
            Mandatory = $true        
        )]
        [string]$DoNotDisturbUsers,
        $SendPings
    )

    Function Purge-TeamsClassic {
        <#
    .SYNOPSIS
        WARNING: Script will force close the 'Teams' process if not told to skip occupied computers.
        Performs a 'purge' of Microsoft Teams Classic, and the Teams Machine-Wide Installer from local computer in preparation for installation of the 'new' Teams (work or school) client.
        Attempts to uninstall the user installations of Teams Classic for any user logged in (if not skipping computer), then attempts to remove the Teams Machine-Wide Installer.
        Script will delete any remaining Teams folders for user or computer, and delete corresponding registry keys.
    
    .NOTES
        Info on new Teams upgrade: https://learn.microsoft.com/en-us/microsoftteams/new-teams-bulk-install-client
        PSADT Silent install of Microsoft Teams Classic (contains uninstall section used as source): https://silentinstallhq.com/microsoft-teams-install-and-uninstall-powershell/
        Info on Teams install/uninstall/cleanup: https://lazyadmin.nl/powershell/microsoft-teams-uninstall-reinstall-and-cleanup-guide-scripts/
    #>
        param(
            $skip_occupied_computers
        )
        $check_for_user = $skip_occupied_computers
        if ($check_for_user.ToLower() -eq 'y') {
            $user_logged_in = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            if ($user_logged_in) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " -NoNewline
                Write-Host "[$env:COMPUTERNAME] :: Found $user_logged_in logged in, skipping removal of Teams Classic." -foregroundcolor yellow
                return
            }
        }
    
        Write-Host ""
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Beginning removal of Teams classic." -foregroundcolor yellow
    
        # install nuget / foce
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | out-null
        # install psadt / force
        Install-Module PSADT -Force | out-null
    
        # kill teams processes:
        Get-PRocess -Name 'Teams' -erroraction silentlycontinue | Stop-Process -Force
    
        $user_folders = Get-ChildItem -Path "C:\Users" -Directory
    
        # uninstall user profile teams isntalations
        ForEach ($single_user_folder in $user_folders) {
    
            $username = $single_user_folder.name
            # get teams update.exe
            $update_exe = Get-ChildItem -Path "$($single_user_folder.fullname)\AppData\Local\Microsoft\Teams\" -Filter "Update.exe" -File -Recurse -ErrorAction SilentlyContinue
            if ($update_exe) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Found $($update_exe.fullname), attempting to uninstall Teams for $($single_user_folder.name)."
                Execute-ProcessAsUser -username $username -Path "$($update_exe.fullname)" -Parameters "--uninstall -s" -Wait
                Start-Sleep -Seconds 5
                Update-Desktop
            }
    
            # remove local app data teams folder
            if (Test-Path  "$($single_user_folder.fullname)\AppData\Local\Microsoft\Teams\" -erroraction silentlycontinue) {
                Remove-item -path  "$($single_user_folder.fullname)\AppData\Local\Microsoft\Teams\" -recurse -force
            }
            # remove any start menu links to old teams for the user:
            Remove-Item -path "$($single_user_folder.fullname)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\*Teams*.lnk" -Force -ErrorAction SilentlyContinue
    
            # remove any desktop shortcuts for teams:
            Remove-Item -path "$($single_user_folder.fullname)\Desktop\*Teams*.lnk" -Force -Erroraction SilentlyContinue
    
            Update-desktop
        }
    
        Write-host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Finished user uninstallations/deletions, attempting Teams Machine-Wide uninstall and ignoring error exit codes." -Foregroundcolor Yellow
    
        Remove-MSIApplications -Name 'Teams Machine-Wide Installer' -ContinueOnError $true
    
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Finished Teams Machine-Wide uninstall, deleting Teams registry keys in HKCU and HKLM."
    
        Remove-RegistryKey -Key 'HKCU:\SOFTWARE\Microsoft\Teams\' -Recurse
        Remove-RegistryKey -Key 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{731F6BAA-A986-45A4-8936-7C3AAAAA760B}' -Recurse
        Remove-RegistryKey -Key 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams' -Recurse
    
    
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Deleting any existing 'Teams Installer' Folder in C:\Program Files (x86)."
    
        Remove-Item -Path 'C:\Program Files (x86)\Teams Installer' -Recurse -Force -ErrorAction SilentlyContinue
    
        Start-Sleep -Seconds 3
    
        Update-Desktop
    }



    ## 1. Handling of TargetComputer input
    ## 2. ask to skip occupied computers
    ## 3. find the Purge-TeamsClassic.ps1 file.
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }
        

    ## 2. Ask to skip occupied computers
    # $DoNotDisturbUsers = Read-Host "Removal of Teams Classic will stop any running teams processes on target machines - skip computers that have users logged in? [y/n]"

    try {
        $DoNotDisturbUsers = $DoNotDisturbUsers.ToLower()
        if ($DoNotDisturbUsers -eq 'y') {
            Write-Host "Skipping occupied computers - acknowledged."
        }
        else {
            Write-Host "Shutting down Teams and removing Teams Classic, even on occupied computers - acknowledged."
        }

    }
    catch {
        Write-Host "Wasn't able to convert $DoNotDisturbUsers to lowercase, assuming 'y'."
        $DoNotDisturbUsers = 'y'
    }

    ## 3. Find the Purge-TeamsClassic.ps1 file.
    # $teamsclassic_scrubber_ps1 = Get-ChildItem -Path "$env:LOCAL_SCRIPTS" -Filter "Purge-TeamsClassic.ps1" -File -ErrorAction SilentlyContinue
    # if (-not $teamsclassic_scrubber_ps1) {
    #     Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Purge-TeamsClassic.ps1 not found in $env:PSMENU_DIR\files, ending function." -Foregroundcolor Red
    #     return
    # }

    # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($teamsclassic_scrubber_ps1.fullname)."
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Beginning removal of Teams Classic on $($ComputerName -join ', ')"

    ## Use PURGE-TEAMSCLASSIC.PS1 file from LOCALSCRIPTS, on each target computer to remove Teams Classic for all users / system.
    ForEach ($single_computer in $ComputerName) {

        if ($single_computer) {

            ## test with ping first:
            $pingreply = Test-Connection $single_computer -Count 1 -Quiet
            if ($pingreply) {
                # Invoke-Command -ComputerName $single_computer -FilePath "$($teamsclassic_scrubber_ps1.fullname)" -ArgumentList $DoNotDisturbUsers
                Invoke-Command -ComputerName $single_computer -ScriptBlock {
                    param(
                        $PurgeTeamsFunction,
                        $DoNotDisturb
                    )
                    $PurgeTeams = [scriptblock]::Create($PurgeTeamsFunction)
                    $PurgeTeams.Invoke($DoNotDisturb)

                } -ArgumentList (Get-Command Purge-TeamsClassic).Definition, $DoNotDisturbUsers

                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Teams Classic removal attempt on $single_computer completed."
            }
        }
    }

    ## Function completion msg

    ## create file to announce completion, for when function is run as background job
    if (-not $env:PSMENU_DIR) {
        $env:PSMENU_DIR = pwd
    }
    ## create simple output path to reports directory
    $thedate = Get-Date -Format 'yyyy-MM-dd'
    $DIRECTORY_NAME = 'TeamsRemoval'
    $OUTPUT_FILENAME = 'TeamsRemoval'
    if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
    }
    
    $counter = 0
    do {
        $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
    } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))



    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Finished removing Microsoft Teams 'Classic' from these computers." | Out-File -FilePath $output_filepath -Append
    $ComputerName | Out-File -FilePath $output_filepath -Append
    # Read-Host "Press enter to continue."
}

Export-ModuleMember -Function *-*
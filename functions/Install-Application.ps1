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
    $AllComputers = Get-Targets -TargetComputer $ComputerName

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

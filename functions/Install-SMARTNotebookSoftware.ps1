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
    $ComputerName = Get-Targets -TargetComputer $ComputerName
    if ($SendPings -eq 'y') {
        $ComputerName = Test-Connectivity -ComputerName $ComputerName
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
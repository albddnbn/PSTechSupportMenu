Function Get-IntuneHardwareIDs {
    <#
    .SYNOPSIS
        Generates a .csv containing hardware ID info for target device(s), which can then be imported into Intune / Autopilot.
        If $ComputerName = '', function is run on local computer.
        Specify GroupTag using DeviceGroupTag parameter.

    .DESCRIPTION
        Uses Get-WindowsAutopilotInfo from: https://github.com/MikePohatu/Get-WindowsAutoPilotInfo/blob/main/Get-WindowsAutoPilotInfo.ps1
        Get-WindowsAutopilotInfo.ps1 is in the supportfiles directory, so it doesn't have to be installed/downloaded from online.

    .PARAMETER TargetComputer
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER DeviceGroupTag
        Specifies the group tag that will be set in target devices' hardware ID info.
        DeviceGroupTag value is used with the -GroupTag parameter of Get-WindowsAutopilotInfo.

    .PARAMETER OutputFile
        Used to create the name of the output .csv file, output to local computer.
        If not supplied, an output filepath will be created using formatted string.

    .INPUTS
        [String[]] - an array of hostnames can be submitted through pipeline for Targetcomputer parameter.

    .OUTPUTS
        Outputs .csv file containing HWID information for target devices, to upload them into Intune.

    .EXAMPLE
        Get Intune Hardware IDs from all computers in room A227 on Stanton campus:
        Get-IntuneHardwareIDs -TargetComputer "t-client-" -OutputFile "TClientIDs" -DeviceGroupTag 'Student Laptops'

    .EXAMPLE
        Get Intune Hardware ID of single target computer
        Get-IntuneHardwareIDs -TargetComputer "t-client-01" -OutputFile "TClient01-ID"

    .NOTES
        Needs utility functions and menu environment variables to run at this point in time.
        Basically just a wrapper for the Get-WindowsAutopilotInfo function, not created by abuddenb.
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param (
        [Parameter(
            Mandatory = $true,
        )]
        $ComputerName,
        [string]$Outputfile,
        [string]$DeviceGroupTag
    )
    ## 1. Set date and report directory variables.
    ## 2. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 3. Create output filepath
    ## 4. Find Get-WindowsAutopilotInfo script and dot source - hopefully from Supportfiles.
    ##    *Making change soon to get rid of the Run-GetWindowsAutopilotinfo file / function setup [02-27-2024].
    ## 1. Date / Report Directory (for output file creation / etc.)
    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    $ComputerName = $ComputerName | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }
    
    ## 3. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "IntuneHardwareIDs"
    if ($Outputfile.tolower() -eq 'n') {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected 'N' input for outputfile, skipping creation of outputfile."
    }
    else {
        if ($Outputfile.toLower() -eq '') {
            $REPORT_DIRECTORY = "$str_title_var"
        }
        else {
            $REPORT_DIRECTORY = $outputfile            
        }
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ## make sure there's a .csv on the end of output file?
    if ($outputfile -notlike "*.csv") {
        $outputfile += ".csv"
    }
    ## 4. Find Get-WindowsAutopilotInfo script and dot source - hopefully from Supportfiles, will check internet if necessary.
    $getwindowsautopilotinfo = Get-ChildItem -Path "$env:SUPPORTFILES_DIR" -Filter "Get-WindowsAutoPilotInfo.ps1" -File -ErrorAction SilentlyContinue
    if (-not $getwindowsautopilotinfo) {
        # Attempt to download script if there's Internet
        # $check_internet_connection = Test-NetConnection "google.com" -ErrorAction SilentlyContinue
        $check_internet_connection = Test-Connection "google.com" -Count 2 -ErrorAction SilentlyContinue
        if ($check_internet_connection) {
            # check for nuget / install
            $check_for_nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
            if ($null -eq $check_for_nuget) {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: NuGet not found, installing now."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
            }
            Install-Script -Name 'Get-WindowsAutopilotInfo' -Force 
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: " -NoNewline
            Write-Host "No internet connection detected, unable to generate hardware ID .csv." -ForegroundColor Red
            return
        }
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Get-WindowsAutopilotInfo.ps1 not found in supportfiles directory, unable to generate hardware ID .csv." -ForegroundColor Red
        return
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Found $($getwindowsautopilotinfo.fullname), importing.`r" -NoNewline
        Get-ChildItem "$env:SUPPORTFILES_DIR" -recurse | unblock-file
    }

    ## 1/2. Filter Targetcomputer for null/empty values and ping test machine.
    ## 3. If machine was responsive:
    ##    - Attempt to use cmdlet to get hwid
    ##    - if Fails (unrecognized because wasn't installed using install-script)
    ##    - Execute from support files.
    ##    * I read that using a @splat like this for parameters gives you the advantage of having only one set to modify,
    ##      as opposed to having to modify two sets of parameters (one for each command in the try/catch)

    ## 1. empty Targetcomputer values will cause errors to display during test-connection / rest of code
    ## 2. Make sure machine is responsive on the network
    ## chop the domain off end of computer name
    ## Define parameters to be used when executing Get-WindowsAutoPilotInfo
    $params = @{
        ComputerName = $ComputerName
        OutputFile   = "$outputfile"
        GroupTag     = $DeviceGroupTag
        Append       = $true
    }
    ## Attempt to use cmdlet from installing script from internet, if fails - revert to script in support 
    ## files (it should have to exist at this point).
    try {
        . "$($getwindowsautopilotinfo.fullname)" @params
    }
    catch {
        Get-WindowsAutoPilotInfo @params
    }

    ## 1. Open the folder that will contain reports if necessary.
    ## 1. Open reports folder
    ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
    try {
        Invoke-item "$($outputfile | split-path -Parent)"
    }
    catch {
        # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
        Invoke-item "$outputfile"
    }

    # read-host "Press enter to return to menu."
}

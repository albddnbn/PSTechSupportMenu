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


function GetOutputFileString {
    <#
        .SYNOPSIS
            Takes input values for part of the filename, the root directory, subfolder title, and whether it should 
            be in the reports or executables directory, and returns an acceptable filename.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$TitleString,
        [Parameter(Mandatory = $true)]
        [string]$Rootdirectory,
        [string]$FolderTitle,
        [switch]$ReportOutput,
        [switch]$ExecutableOutput
    )
    ForEach ($file_ext in @('.csv', '.xlsx', '.ps1', '.exe')) {
        Write-Verbose "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Checking for file extension: $file_ext."
        $TitleString = $TitleString -replace $file_ext, ''
        Write-Verbose "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Removed file extension, TitleString is now: $TitleString."
    }

    $thedate = Get-Date -Format 'yyyy-MM-dd'

    # create outputfolder
    if ($Reportoutput) {
        $subfolder = 'reports'
    }
    elseif ($ExecutableOutput) {
        $subfolder = 'executables'
    }

    Write-Verbose "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Subfolder determined to be: $subfolder."

    $outputfolder = "$Rootdirectory\$subfolder\$thedate\$FolderTitle"

    if (-not (Test-Path $outputfolder)) {
        Write-Verbose "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Couldn't find folder at: $outputfolder, creating it now."
        New-Item $outputfolder -itemtype 'directory' -force | Out-null
    }
    $filename = "$TitleString-$thedate"
    $full_output_path = "$outputfolder\$filename"
    # make sure outputfiles dont exist
    if ($ReportOutput) {
        $x = 0
        while ((Test-Path "$full_output_path.csv") -or (Test-Path "$full_output_path.xlsx")) {
            $x++
            $full_output_path = "$outputfolder\$filename-$x"
        }
    }
    elseif ($ExecutableOutput) {
        $x = 0
        while ((Test-Path "$full_output_path.ps1") -or (Test-Path "$full_output_path.exe")) {
            $x++
            $full_output_path = "$outputfolder\$filename-$x"
        }
    }

    Write-Verbose "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Full output path determined to be: $full_output_path."

    return $full_output_path
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
    $PingCount = $PingCount

    # $list_of_online_computers = [system.collections.arraylist]::new()
    # $list_of_offline_computers = [system.collections.arraylist]::new()
    $online_results = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {
        if (Test-Connection $single_computer -Count $PingCount -Quiet) {
            Write-Host "$single_computer is online." -ForegroundColor Green
            
            $online_results.Add($single_computer) | Out-Null
        }
        else {
            Write-Host "$single_computer is offline." -ForegroundColor Red
        }
    
    }

    return $online_results

}

function Count-TempProfiles {
    <#
    .SYNOPSIS
        Assumes temporary user folder suffix is: "*.$env:USERDOMAIN*" unless otherwise specified.
        Counts the number of temporary user folders/profiles on each target computer.
        Outputs a gridview, excel spreadsheet, and/or csv file.
    
    .DESCRIPTION
        This function can be useful to find:
        1. Specific user accounts or profiles that are having issues on the network and need assistance.
        2. Specific computers on the network that are having issues syncing with domain / network shares.
        3. Specific files that cause issues with redirected folders and roaming user profiles.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER TempFolderSuffix
        Suffix for temporary user folders. Default is "*.$env:USERDOMAIN*"
        Ex: C:\Users\Tsmith28.LLDC.000

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects.
        Multiple objects can be output for one computer - each object will pertain to each user that has a temporary folder on that computer.
        Ex: One object (row in output table/grid), could be for user: Tsmith28 with folder count 2.
        This would signify that they have 2 temp folders on that computer in addition to any regular user folder.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Count-TempProfiles -TargetComputer 't-client-'
        Count-TempProfiles -TargetComputer 't-client-01,t-client-02,t-client-03' -OutputFile 'n'
        Count-TempProfiles -TargetComputer 'C:\users\public\computers.txt' -OutputFile 'A220'

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
        [string]$Outputfile = '',
        [string]$TempFolderSuffix = "$env:USERDOMAIN"
    )
    ## allow user to press enter for default suffix.
    if ($TempFolderSuffix -eq '') {
        $TempFolderSuffix = "$env:USERDOMAIN"
    }
    Write-Host "Temporary folder suffix set to: " -NoNewline
    Write-Host "$TempFolderSuffix" -ForegroundColor Yellow
    ## 4. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "TempProfiles"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ForEach ($single_computer in $ComputerName) {
        if ($single_computer) {
            ## 2. Check for connectivity
            if ([System.IO.Directory]::Exists("\\$single_computer\c$\Users")) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online." -Foregroundcolor Green

                # get all temp profiles on computer / count / create object
                $temp_profile_folders = Get-Childitem -Path "\\$single_computer\c$\Users" -Filter "*.$env:USERDOMAIN*" -Directory -ErrorAction SilentlyContinue
                ## create object with  user, computer name, folder count to object, add to arraylist
                ForEach ($single_folder in $temp_profile_folders) {

                    $foldername = $single_folder.name

                    $username = $foldername.split('.')[0]
                    ## if the user and computer combo are not in results - add with count of 1
                    if ($results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) }) {
                        $results | Where-Object { ($_.User -eq $username) -and ($_.Computer -eq $single_computer) } | ForEach-Object { $_.FolderCount++ }
                        Write-Host "Found existing entry for $username and $single_computer increased FolderCount by 1."
                    }
                    else {
                        $temp_profile = [pscustomobject]@{
                            User        = $username
                            Computer    = $single_computer
                            FolderCount = 1
                        }
                        $results.Add($temp_profile) | Out-Null
                        Write-Host "Added new entry for $username and $single_computer."
                    }
                }
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is offline." -Foregroundcolor Red
                continue
            }               
        }
    }

    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $str_title_var
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-CurrentUser." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    # read-host "Press enter to return results."
    return $results
}

function Get-AssetInformation {
    <#
    .SYNOPSIS
        Attempts to use Dell Command Configure to get asset tag, if not available uses built-in powershell cmdlets.

    .DESCRIPTION
        Function will work as a part of the Terminal menu or outside of it.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, computer model, bios version/release date, asset tag/serial number, and connected monitor information.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Get-AssetInformation

    .EXAMPLE
        Get-AssetInformation -TargetComputer s-c127-01 -Outputfile C127-01-AssetInfo

    .NOTES
        Monitor details show up in .csv but not .xlsx right now - 12.1.2023
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param (
        [Parameter(
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }
   
    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "AssetInfo"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }


    ## Asset info scriptblock used to get local asset info from each target computer.
    $asset_info_scriptblock = {
        # computer model (ex: 'precision 3630 tower'), BIOS version, and BIOS release date
        $computer_model = get-ciminstance -class win32_computersystem | select -exp model
        $biosversion = get-ciminstance -class win32_bios | select -exp smbiosbiosversion
        $bioreleasedate = get-ciminstance -class win32_bios | select -exp releasedate
        # Asset tag from BIOS (tested with dell computer)
        try {
            $command_configure_exe = Get-ChildItem -Path "${env:ProgramFiles(x86)}\Dell\Command Configure\x86_64" -Filter "cctk.exe" -File -ErrorAction Silentlycontinue
            # returns a string like: 'Asset=2001234'
            $asset_tag = &"$($command_configure_exe.fullname)" --asset
            $asset_tag = $asset_tag -replace 'Asset=', ''
        }
        catch {
            $asset_tag = Get-Ciminstance -class win32_systemenclosure | select -exp smbiosassettag
            # asus motherboard returned 'default string'
            if ($asset_tag.ToLower() -eq 'default string') {
                $asset_tag = 'No asset tag set in BIOS'
            }    
        }
        $computer_serial_num = get-ciminstance -class win32_bios | select -exp serialnumber
        # get monitor info and create a string from it (might be unnecessary, or a lengthy approach):
        $monitors = Get-CimInstance WmiMonitorId -Namespace root\wmi -ComputerName $ComputerName -ErrorAction SilentlyContinue
        if ($monitors) {
            $monitors = $monitors | Select Active, ManufacturerName, UserFriendlyName, SerialNumberID, YearOfManufacture
            $monitor_string = ""
            $monitor_count = 0
            $monitors | ForEach-Object {
                $_.UserFriendlyName = [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)
                $_.SerialNumberID = [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)
                $_.ManufacturerName = [System.Text.Encoding]::ASCII.GetString($_.ManufacturerName)
                $manufacturername = $($_.ManufacturerName).trim()
                $monitor_string += "Maker: $manufacturername,Mod: $($_.UserFriendlyName),Ser: $($_.SerialNumberID),Yr: $($_.YearOfManufacture)"
                $monitor_count++
            }
        }
        else {
            $monitor_string = "No monitor information available."
            $monitor_count = 0
        }
        $obj = [PSCustomObject]@{
            model               = $computer_model
            biosversion         = $biosversion
            bioreleasedate      = $bioreleasedate
            asset_tag           = $asset_tag
            computer_serial_num = $computer_serial_num
            monitors            = $monitor_string
            NumMonitors         = $monitor_count
        }
        return $obj
    }

    $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock $asset_info_scriptblock -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

    ## Collects hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## If there were any results - output them to terminal and/or report files as necessary.
    if ($results) {
        ## Sort the results
        $results = $results | sort -property pscomputername
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -Title $str_title_var
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            
            
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                ## xlsx attempt:
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params # -ErrorAction SilentlyContinue
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            Invoke-item "$($outputfile | split-path -Parent)"
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

    }
    # read-host "Press enter to return results."
    return $results
}

function Get-ComputerDetails {
    <#
    .SYNOPSIS
        Collects: Manufacturer, Model, Current User, Windows Build, BIOS Version, BIOS Release Date, and Total RAM from target machine(s).
        Outputs: A .csv and .xlsx report file if anything other than 'n' is supplied for the $OutputFile parameter.

    .DESCRIPTION

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        'n' or 'N' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220-Info', output file(s) will be in the $env:PSMENU_DIR\reports\2023-11-1\A220-Info\ directory.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, computer model, bios version/release date, last boot time, and other info.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Output details for a single hostname to "sa227-28-details.csv" and "sa227-28-details.xlsx" in the 'reports' directory.
        Get-ComputerDetails -TargetComputer "t-client-28" -Outputfile "tclient-28-details"

    .EXAMPLE
        Output details for all hostnames starting with g-pc-0 to terminal.
        Get-ComputerDetails -TargetComputer 'g-pc-0' -outputfile 'n'

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param (
        [Parameter(
            Mandatory = $true, ValueFromPipeline = $true
        )]
        $ComputerName,
        [string]$Outputfile,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }


    $str_title_var = "PCdetails"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ## Save results to variable
    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        # Gets active user, computer manufacturer, model, BIOS version & release date, Win Build number, total RAM, last boot time, and total system up time.
        # object returned to $results list
        $computersystem = Get-CimInstance -Class Win32_Computersystem
        $bios = Get-CimInstance -Class Win32_BIOS
        $operatingsystem = Get-CimInstance -Class Win32_OperatingSystem

        $lastboot = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
        $uptime = ((Get-Date) - $lastboot).ToString("dd\.hh\:mm\:ss")
        $obj = [PSCustomObject]@{
            Manufacturer    = $($computersystem.manufacturer)
            Model           = $($computersystem.model)
            CurrentUser     = $((get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username)
            WindowsBuild    = $($operatingsystem.buildnumber)
            BiosVersion     = $($bios.smbiosbiosversion)
            BiosReleaseDate = $($bios.releasedate)
            TotalRAM        = $((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb)
            LastBoot        = $lastboot
            SystemUptime    = $uptime
        }
        $obj
    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername -ErrorAction SilentlyContinue

    ## Collects hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {
        ## Sort the results
        $results = $results | sort -property pscomputername
        if ($outputfile.tolower() -eq 'n') {
            # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected 'N' input for outputfile, skipping creation of outputfile."
            if ($results.count -le 2) {
                $results | Format-List
                # $results | Out-GridView
            }
            else {
                $results | out-gridview -Title $str_title_var
            }
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                ## xlsx attempt:
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-ComputerDetails." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    # read-host "Press enter to return results."
    return $results
    
}

function Get-ConnectedPrinters {
    <#
    .SYNOPSIS
        Checks the target computer, and returns the user that's logged in, and the printers that user has access to.

    .DESCRIPTION
        This function, unlike some others, only takes a single string DNS hostname of a target computer.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        'n' or 'no' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER FolderTitleSubstring
        If specified, the function will create a folder in the 'reports' directory with the specified substring in the title, appended to the $REPORT_DIRECTORY String (relates to the function title).

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and list of connected printers.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Get-ConnectedPrinters -TargetComputer 't-client-07'

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
        [string]$Outputfile,
        $SendPings
    )
    $ComputerName = GetTargets -TargetComputer $ComputerName
    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }
     
        
    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "Printers"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ## Scriptblock - lists connected/default printers
    $list_local_printers_block = {
        # Everything will stay null, if there is no user logged in
        $obj = [PScustomObject]@{
            Username          = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            DefaultPrinter    = $null
            ConnectedPrinters = $null
        }

        # Only need to check for connected printers if a user is logged in.
        if ($obj.Username) {
            # get connected printers:
            get-ciminstance -class win32_printer | select name, Default | ForEach-Object {
                if (($_.name -notin ('Microsoft Print to PDF', 'Fax')) -and ($_.name -notlike "*OneNote*")) {
                    if ($_.name -notlike "Send to*") {
                        $obj.ConnectedPrinters = "$($obj.ConnectedPrinters), $($_.name)"
                    }
                }   
            }
        }
        $obj
    }

    ## Create empty results container to use during process block
    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock $list_local_printers_block  -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

    ## Collects hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -Title $str_title_var
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
            
            ## Try ImportExcel
            try {
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-ConnectedPrinters." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}

function Get-CurrentUser {
    <#
    .SYNOPSIS
        Gets user logged into target system(s).
        Checks if teams or zoom processes are running and returns True/False for each in report/terminal output.

    .DESCRIPTION
        Creates report with current user, computer model, and if Teams or Zoom are running.
        If no output file is specified, terminal output only ($Outputfile = 'n').

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and whether the Teams/Zoom processes are running.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get users on all S-A231 computers:
        Get-CurrentUser -Targetcomputer "s-a231-"

    .EXAMPLE
        2. Get user on a single target computer:
        Get-CurrentUser -TargetComputer "t-client-28"

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
        [string]$Outputfile,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }

    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "CurrentUsers"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        $obj = [PSCustomObject]@{
            Model        = (get-ciminstance -class win32_computersystem).model
            CurrentUser  = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            TeamsRunning = $(if (Get-PRocess -Name 'Teams' -ErrorAction SilentlyContinue) { $true } else { $false })
            ZoomRunning  = $(if (Get-PRocess -Name 'Zoom' -ErrorAction SilentlyContinue) { $true } else { $false })

        }
        $obj
    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername
    ## Collects hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $str_title_var
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
         
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-CurrentUser." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }
    return $results
}

function Get-InstalledDotNetversions {
    <#
    .SYNOPSIS
        Gets a list of installed dotnet versions on target computers. Returns results.

    .DESCRIPTION
        Creates report if anything except 'n' is supplied for Outputfile.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing the hostname and info on installed .net versions.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get dotnet versions on single computer, output results to terminal/gridview
        Get-InstalledDotNetVersions -TargetComputer "t-client-01" -outputfile 'n'

    .EXAMPLE
        2. Get user on group of computers with hostnames starting with t-client-, output default filename reports
        Get-InstalledDotNetVersions -TargetComputer "t-client-" -outputfile ''

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
        [string]$Outputfile,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }    

    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "InstalledDotNet"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse | `
            Get-ItemProperty -Name version -EA 0 | Where { $_.PSChildName -Match '^(?!S)\p{L}' } |`
            Select PSChildName, version

    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

    $errored_machines = $RemoteError.CategoryInfo.TargetName

    if ($results) {

        # ForEach ($single_result in $results) {
        #     $single_result
        # }

        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -Title "Installed .NET Versions"
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
            $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append
           
            ## Try ImportExcel
            try {
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Get-InstalledDotNetVersions." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}

Function Get-IntuneHardwareIDs {
    <#
    .SYNOPSIS
        Generates a .csv containing hardware ID info for target device(s), which can then be imported into Intune / Autopilot.
        If $ComputerName = '', function is run on local computer.
        Specify GroupTag using DeviceGroupTag parameter.

    .DESCRIPTION
        Uses Get-WindowsAutopilotInfo from: https://github.com/MikePohatu/Get-WindowsAutoPilotInfo/blob/main/Get-WindowsAutoPilotInfo.ps1
        Get-WindowsAutopilotInfo.ps1 is in the supportfiles directory, so it doesn't have to be installed/downloaded from online.

    .PARAMETER ComputerName
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

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

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
            Mandatory = $true
        )]
        $ComputerName,
        [string]$Outputfile,
        [string]$DeviceGroupTag,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }    
    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ## make sure there's a .csv on the end of output file?
    if ($outputfile -notlike "*.csv") {
        $outputfile += ".csv"
    }
    ## Find Get-WindowsAutopilotInfo script and dot source - hopefully from Supportfiles, will check internet if necessary.
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

    ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
    try {
        Invoke-item "$($outputfile | split-path -Parent)"
    }
    catch {
        Invoke-item "$outputfile"
    }

}

function Get-InventoryDetails {
    <#
    .SYNOPSIS
        Targets supplied computer names, and takes inventory of computer asset tag/serial number, and any other
        details that can be gathered from the connected monitors.
        Outputs a csv with results.

    .DESCRIPTION
        This has mainly been tested with Dell equipment - computers and monitors.
        Still in testing/development phase but should work.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, logged in user, and whether the Teams/Zoom processes are running.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        1. Get users on all S-A231 computers:
        Sample-Function -Targetcomputer "s-a231-"

    .EXAMPLE
        2. Get user on a single target computer:
        Sample-Function -TargetComputer "t-client-28"

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
        [string]$Outputfile,
        $SendPings
    )

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    } 
    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    ###
    ### *** INSERT THE TITLE OF YOUR FUNCTION / REPORT FOR $str_title_var ***
    ###
    $str_title_var = "Inventory"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    $results = Invoke-Command -ComputerName $ComputerName -scriptblock {
        $pc_asset_tag = Get-Ciminstance -class win32_systemenclosure | select -exp smbiosassettag
        $pc_model = Get-Ciminstance -class win32_computersystem | select -exp model
        $pc_serial = Get-Ciminstance -class Win32_SystemEnclosure | select -exp serialnumber
        $pc_manufacturer = Get-Ciminstance -class Win32_ComputerSystem | select -exp manufacturer
        $monitors = Get-CimInstance WmiMonitorId -Namespace root\wmi | Select SerialNumberID, ManufacturerName, UserFriendlyName
        $monitors | % { 
            # $_.serialnumberid = [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)
            # 
            $_.UserFriendlyName = [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)
            if ($_.UserFriendlyName -like "*P19*") {
                $_.serialnumberid = $(([System.Text.Encoding]::ASCII.GetString($_.SerialNumberID -notmatch 0)).Trim())
            }
            else {
                ## from copilot: his will replace any character that is not in the range from hex 20 (space) to hex 7E (tilde), which includes all printable ASCII characters, with nothing.
                $_.serialnumberid = ($([System.Text.Encoding]::ASCII.GetString($_.SerialNumberID ).Trim()) -replace '[^\x20-\x7E]', '')
            }
                
            $_.ManufacturerName = [System.Text.Encoding]::ASCII.GetString($_.ManufacturerName)
        }
            
        $obj = [pscustomobject]@{
                
            computer_asset        = $pc_asset_tag
            computer_location     = $(($env:COMPUTERNAME -split '-')[1]) ## at least make an attempt to get location.
            computer_model        = $pc_model
            computer_serial       = $pc_serial
            computer_manufacturer = $pc_manufacturer
            monitor_serials       = $(($monitors.serialnumberid) -join ',')
            monitor_manufacturers = $(($monitors.ManufacturerName) -join ',')
            monitor_models        = $(($monitors.UserFriendlyName) -join ',')
            inventoried           = $true
        }
        # Write-Host "Gathered details from $env:COMPUTERNAME"
        # Write-Host "$obj"
        $obj
    } -ErrorVariable RemoteError | Select * -ExcludeProperty PSShowComputerName, RunspaceId

    $not_inventoried = $ComputerName | ? { $_ -notin $results.pscomputername }
    $not_inventoried += $RemoteError.CategoryInfo.TargetName | ? { $_ -notin $not_inventoried }

    ## This section will attempt to output a CSV and XLSX report if anything other than 'n' was used for $Outputfile.
    ## If $Outputfile = 'n', results will be displayed in a gridview, with title set to $str_title_var.
    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -title $str_title_var
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {

                Import-Module ImportExcel

                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output from Sample-Function." | Out-File -FilePath "$outputfile.csv"

        Invoke-Item "$outputfile.csv"
    }

    return $results
}

function Ping-TestReport {
    <#
    .SYNOPSIS
        Pings a group of computers a specified amount of times, and outputs the successes / total pings to a .csv and .xlsx report.

    .DESCRIPTION
        Script will output to ./reports/<date>/ folder. It calculates average response time, and packet loss percentage.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PingCount
        Number of times to ping each computer.

    .PARAMETER OutputFile
        'n' or 'no' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'Room1', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - Room1\

    .EXAMPLE
        Ping-TestReport -Targetcomputer "g-client-" -PingCount 10 -Outputfile "GClientPings"
    
    .EXAMPLE
        Ping-TestReport -Targetcomputer "g-client-" -PingCount 2

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
        $PingCount,
        [string]$Outputfile = ''
    )
    ## 1. Set date and AM / PM variables
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    ## 3. If provided, use outputfile input to create report output filepath.
    ## 4. Create arraylist to store results

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## 1. Set date and AM / PM variables
    $am_pm = (Get-Date).ToString('tt')

    ## 2. If provided, use outputfile input to create report output filepath.
    ## Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "Pings-$Outputfile-$(Get-Date -Format 'hh-MM')$($am_pm)"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }


    ## 3. Create arraylist to store results
    $results = [system.collections.arraylist]::new()

    $PingCount = [int]$PingCount

    ## 1. Ping EACH Target computer / record results into ps object, add to arraylist (results_container)
    ## 2. Set object property values:
    ## 3. Send pings - object property values are derived from resulting object
    ## 4. Number of responses
    ## 5. Calculate average response time for successful responses
    ## 6. Calculate packet loss percentage
    ForEach ($single_computer in $ComputerName) {
        ## 1. empty Targetcomputer values will cause errors to display during test-connection / rest of code
        if ($single_computer) {

            ## check if network path exists first - that way we don't waste time pinging machine thats offline?
            if (-not ([System.IO.Directory]::Exists("\\$single_computer\c$"))) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is not online." -foregroundcolor red
                continue
            }
            


            ## 2. Create object to store results of ping test on single machine
            $obj = [pscustomobject]@{
                Sourcecomputer       = $env:COMPUTERNAME
                ComputerHostName     = $single_computer
                TotalPings           = $pingcount
                Responses            = 0
                AvgResponseTime      = 0
                PacketLossPercentage = 0
            }
            Write-Host "Sending $pingcount pings to $single_computer..."
            ## 3. Send $PINGCOUNT number of pings to target device, store results
            $send_pings = Test-Connection -ComputerName $single_computer -count $PingCount -ErrorAction SilentlyContinue
            ## 4. Set number of responses from target machine
            $obj.responses = $send_pings.count
            ## 5. Calculate average response time for successful responses
            $sum_of_response_times = $($send_pings | measure-object responsetime -sum)
            if ($obj.Responses -eq 0) {
                $obj.AvgResponseTime = 0
            }
            else {
                $obj.avgresponsetime = $sum_of_response_times.sum / $obj.responses
            }
            ## 6. Calculate packet loss percentage - divide total pings by responses
            $total_drops = $obj.TotalPings - $obj.Responses
            $obj.PacketLossPercentage = ($total_drops / $($obj.TotalPings)) * 100

            ## 7. Add object to container created in BEGIN block
            $results.add($obj) | Out-Null
        }
    }

    ## Report file creation or terminal output
    if ($results) {
        ## 1. Sort any existing results by computername
        $results = $results | sort -property pscomputername
        ## 2. Output to gridview if user didn't choose report output.
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview
        }
        else {
            ## 3. Create .csv/.xlsx reports if possible
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
            ## Try ImportExcel
            try {
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = $str_title_var
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = $str_title_var
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$Outputfile.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
            ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
            try {
                Invoke-item "$($outputfile | split-path -Parent)"
            }
            catch {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
                Invoke-item "$outputfile.csv"
            }
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."
    }
    # read-host "`nPress [ENTER] to return results."
    return $results
}

function Scan-ForAppOrFilePath {
    <#
    .SYNOPSIS
        Scan a group of computers for a specified file/folder or application, and output the results to a .csv and .xlsx report.

    .DESCRIPTION
        The script searches application DisplayNames when the -type 'app' argument is used, and searches for files/folders when the -type 'path' argument is used.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER Item
        The item to search for. 
        If the -SearchType 'app' argument is used, this should be the application's DisplayName. 
        If the -SearchType 'path' argument is used, this should be the path to search for, Ex: C:\users\public\test.txt.

    .PARAMETER OutputFile
        Used to create the output filename/path if supplied.

    .PARAMETER SearchType
        The type of search to perform. 
        This can be either 'app' or 'path'. 
        If 'app' is specified, the script will search for the specified application in the registry. 
        If 'path' is specified, the script will search for the specified file/folder path on the target's filesystem.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Scan-ForAppOrFilePath -ComputerList 't-client-01' -SearchType 'app' -Item 'Microsoft Teams' -outputfile 'teams'

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
        [Parameter(Mandatory = $true)]
        [ValidateSet('Path', 'App', 'File', 'Folder')]
        [String]$SearchType,
        [Parameter(Mandatory = $true)]
        [String]$Item,
        [String]$Outputfile,
        $SendPings
    )
    ## 1. Set date
    ## 2. Handle targetcomputer if not submitted through pipeline
    ## 3. Create output filepath, clean any input file search paths that are local,  
    ## and handle TargetComputer input / filter offline hosts.
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }
        

    ## 3. Outputfile handling - either create default, create filenames using input - report files are mandatory 
    ##    in this function.
    $str_title_var = "$SearchType-scan"
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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }
        
    if (@('path', 'file', 'folder') -contains $SearchType.ToLower()) {

        $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            $obj = [PSCustomObject]@{
                Name           = $env:COMPUTERNAME
                Path           = $using:item
                PathPresent    = $false
                PathType       = $null
                LastWriteTime  = $null
                CreationTime   = $null
                LastAccessTime = $null
                Attributes     = $null
            }
            $GetSpecifiedItem = Get-Item -Path "$using:item" -ErrorAction SilentlyContinue
            if ($GetSpecifiedItem.Exists) {
                $details = $GetSpecifiedItem | Select FullName, *Time, Attributes, Length
                $obj.PathPresent = $true
                if ($GetSpecifiedItem.PSIsContainer) {
                    $obj.PathType = 'Folder'
                }
                else {
                    $obj.PathType = 'File'
                }
                $obj.LastWriteTime = $details.LastWriteTime
                $obj.CreationTime = $details.CreationTime
                $obj.LastAccessTime = $details.LastAccessTime
                $obj.Attributes = $details.Attributes
            }
            else {
                $obj.PathPresent = "Filepath not found"
            }
            $obj
        } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername
    
    }
    ## Application search
    elseif ($SearchType -eq 'App') {

        $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
            # $app_matches = [System.Collections.ArrayList]::new()
            # Define the registry paths for uninstall information
            $registryPaths = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            $obj = $null
            # Loop through each registry path and retrieve the list of subkeys
            foreach ($path in $registryPaths) {
                $uninstallKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
                # Skip if the registry path doesn't exist
                if (-not $uninstallKeys) {
                    continue
                }
                # Loop through each uninstall key and display the properties
                foreach ($key in $uninstallKeys) {
                    $keyPath = Join-Path -Path $path -ChildPath $key.PSChildName
                    $displayName = (Get-ItemProperty -Path $keyPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                    if ($displayName -like "*$using:Item*") {
                        $uninstallString = (Get-ItemProperty -Path $keyPath -Name "UninstallString" -ErrorAction SilentlyContinue).UninstallString
                        $version = (Get-ItemProperty -Path $keyPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
                        $publisher = (Get-ItemProperty -Path $keyPath -Name "Publisher" -ErrorAction SilentlyContinue).Publisher
                        $installLocation = (Get-ItemProperty -Path $keyPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
                        # $productcode = (Get-ItemProperty -Path $keyPath -Name "productcode" -ErrorAction SilentlyContinue).productcode
                        $installdate = (Get-ItemProperty -Path $keyPath -Name "installdate" -ErrorAction SilentlyContinue).installdate

                        $obj = [PSCustomObject]@{
                            ComputerName    = $env:COMPUTERNAME
                            AppName         = $displayName
                            AppVersion      = $version
                            InstallDate     = $installdate
                            InstallLocation = $installLocation
                            Publisher       = $publisher
                            UninstallString = $uninstallString
                        }
                        $obj
                    }
                }
            }
            # if ($null -eq $obj) {
            #     $obj = [PSCustomObject]@{
            #         ComputerName    = $single_computer
            #         AppName         = "No matching apps found for $using:Item"
            #         AppVersion      = $null
            #         InstallDate     = $null
            #         InstallLocation = $null
            #         Publisher       = $null
            #         UninstallString = "No matching apps found"
            #     }
            #     $obj
            # }
        } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

        # $search_result
        # read-host "enter"
    }

    ## Collects hostnames from any Invoke-Command error messages
    $errored_machines = $RemoteError.CategoryInfo.TargetName
 
    ## 1. Output findings (if any) to report files or terminal
    if ($results) {
        $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
        "These machines errored out:`r" | Out-File -FilePath "$outputfile-Errors.csv"
        $errored_machines | Out-File -FilePath "$outputfile-Errors.csv" -Append      
        ## Try ImportExcel
        try {
            ## xlsx attempt:
            $params = @{
                AutoSize             = $true
                TitleBackgroundColor = 'Blue'
                TableName            = "$REPORT_DIRECTORY"
                TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                BoldTopRow           = $true
                WorksheetName        = "$SearchType-Search"
                PassThru             = $true
                Path                 = "$Outputfile.xlsx" # => Define where to save it here!
            }
            $Content = Import-Csv "$Outputfile.csv"
            $xlsx = $Content | Export-Excel @params
            $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
            $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
            Close-ExcelPackage $xlsx
        }
        catch {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
        }
        ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
        try {
            Invoke-item "$($outputfile | split-path -Parent)"
        }
        catch {
            # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder." -Foregroundcolor Yellow
            Invoke-item "$outputfile.csv"
        }
            
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."
    }
    # Read-Host "`nPress [ENTER] to return results."
    return $results
}

function Scan-SoftwareInventory {
    <#
    .SYNOPSIS
        Scans a group of computers for installed applications and exports results to .csv/.xlsx - one per computer.

    .DESCRIPTION
        Scan-SoftwareInventory can handle a single string hostname as a target, a single string filepath to hostname list, or an array/arraylist of hostnames.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .PARAMETER Outputfile
        A string used to create the output .csv and .xlsx files. If not specified, a default filename is created.

    .PARAMETER AppsToLookFor
        Comma-separated list.
        Optional parameter to specify a list of applications/strings to look for. If not specified, all applications are scanned.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .EXAMPLE
        Scan-SoftwareInventory -TargetComputer "t-client-28" -Title "tclient-28-details"

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
            Mandatory = $true)]
        [string]$OutputFile,
        $AppsToLookFor,
        $SendPings
    )
    ## 1. Define title, date variables
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    ## 3. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    ## 4. Create empty results container
    $AppsToLookFor = $AppsToLookFor.split(",")
    if ($AppsToLookFor -isnot [array]) {
        $AppsToLookFor = @($AppsToLookFor)
    }
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)

    $ComputerName = GetTargets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = TestConnectivity -ComputerName $ComputerName
    }        
    ## 3. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    $str_title_var = "SoftwareScan"

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
        $OutputFile = GetOutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }
        
    $results = invoke-command -computername $ComputerName -scriptblock {

        $targetapps = ($using:AppsToLookFor)
        $registryPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        foreach ($path in $registryPaths) {
            $uninstallKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            # Skip if the registry path doesn't exist
            if (-not $uninstallKeys) {
                continue
            }
            # Loop through each uninstall key and display the properties
            foreach ($key in $uninstallKeys) {
                $keyPath = Join-Path -Path $path -ChildPath $key.PSChildName
                $displayName = (Get-ItemProperty -Path $keyPath -Name "DisplayName" -ErrorAction SilentlyContinue).DisplayName
                $uninstallString = (Get-ItemProperty -Path $keyPath -Name "UninstallString" -ErrorAction SilentlyContinue).UninstallString
                $version = (Get-ItemProperty -Path $keyPath -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
                $publisher = (Get-ItemProperty -Path $keyPath -Name "Publisher" -ErrorAction SilentlyContinue).Publisher
                $installLocation = (Get-ItemProperty -Path $keyPath -Name "InstallLocation" -ErrorAction SilentlyContinue).InstallLocation
                $productcode = (Get-ItemProperty -Path $keyPath -Name "productcode" -ErrorAction SilentlyContinue).productcode
                $installdate = (Get-ItemProperty -Path $keyPath -Name "installdate" -ErrorAction SilentlyContinue).installdate
                $application_size = $null ## define as null for each loopthru


                if (($displayname -ne '') -and ($null -ne $displayname)) {
                    # if a target app list was provided, cycle through it and see if we're dealing with an app installation that is being searched for.
                    if ($targetapps) {
                        $matched_app = $false

                        $targetapps | % {
                            if ($displayname -like "*$_*") {
                                $matched_app = $true
                            }
                        }
                        ## If a search list was provided and there was no match, skip this app listing and move on to next
                        if (-not $matched_app) {
                            continue
                        }
                    }

                    ## Attempt to get approx 'size' of install location folder:
                    if ($installlocation) {
                        $application_size = (Get-ChildItem -Path "$installLocation" -Recurse -ErrorAction SilentlyContinue | MEasure-Object -Property Length -Sum -ErrorAction SilentlyContinue).sum / 1GB
                        $application_size = [Math]::Round($application_size, 2)
                        $application_size = "$application_size GB"
                    }

                    $obj = [pscustomobject]@{
                        DisplayName     = $displayName
                        UninstallString = $uninstallString
                        Version         = $version
                        Publisher       = $publisher
                        InstallLocation = $installLocation
                        ProductCode     = $productcode
                        InstallDate     = $installdate
                        ApplicationSize = $application_size
                    }
                    $obj    
                }        
            }
        } 
    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## 1. Get list of unique computer names from results - use it to sort through all results to create a list of apps for 
    ##    a specific computer, output apps to report, then move on to next iteration of loop.
    if ($results) {
        ## 1. get list of UNIQUE pscomputername s from the results - a file needs to be created for EACH computer.
        $unique_hostnames = $($results.pscomputername) | select -Unique

        if ($errored_machines) {
            Write-Host "These machines errored out during Invoke-Command." -ForegroundColor Red
            $errored_machines
        }

        ForEach ($single_computer_name in $unique_hostnames) {
            # get that computers apps
            $apps = $results | where-object { $_.pscomputername -eq $single_computer_name }
            # create the full filepaths
            $output_filepath = "$outputfile-$single_computer_name"
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Exporting files for $single_computername to $output_filepath."

            $apps | Export-Csv -Path "$outputfile-$single_computer_name.csv" -NoTypeInformation
            ## Try ImportExcel
            try {
                ## xlsx attempt:
                $params = @{
                    AutoSize             = $true
                    TitleBackgroundColor = 'Blue'
                    TableName            = "$REPORT_DIRECTORY"
                    TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
                    BoldTopRow           = $true
                    WorksheetName        = "$single_computer_name Apps"
                    PassThru             = $true
                    Path                 = "$Outputfile.xlsx" # => Define where to save it here!
                }
                $Content = Import-Csv "$outputfile-$single_computer_name.csv"
                $xlsx = $Content | Export-Excel @params
                $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
                $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
                Close-ExcelPackage $xlsx
            }
            catch {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: ImportExcel module not found, skipping xlsx creation." -Foregroundcolor Yellow
            }
                
        }
        ## Try opening directory (that might contain xlsx and csv reports), default to opening csv which should always exist
        try {
            Invoke-item "$($outputfile | split-path -Parent)"
        }
        catch {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Could not open output folder, attempting to open first .csv in list." -Foregroundcolor Yellow
            Invoke-item "$outputfile-$($unique_hostnames | select -first 1).csv"
        }
    }
    # read-host "`nPress [ENTER] to return results."
    return $results
}

function Test-ConnectivityQuick {
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
    $PingCount = $PingCount
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    if ($null -eq $ComputerName) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Detected pipeline for targetcomputer." -Foregroundcolor Yellow
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

    ## COLLECTIONS LISTS - successful/failed pings.
    $results = [system.collections.arraylist]::new()
    # $list_of_online_computers = [system.collections.arraylist]::new()
    # $list_of_offline_computers = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {

        if ($single_computer) {
            $connection_result = Test-Connection $single_computer -count $PingCount -ErrorAction SilentlyContinue
            # $connection_result
            # $ping_responses = $([string[]]($connection_result | where-object { $_.status -eq 'Success' })).count
            $ping_responses = $connection_result.count
            ## Create object
            $ping_response_obj = [pscustomobject]@{
                ComputerName  = $single_computer
                Status        = ""
                PingResponses = $ping_responses
                NumberPings   = $PingCount
            }

            if ($connection_result) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: $single_computer is online [$ping_responses responses]" -foregroundcolor green
                # $list_of_online_computers.add($single_computer) | Out-Null
                $ping_response_obj.Status = 'online'
            }
            else {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: " -NoNewline
                Write-Host "$single_computer is not online." -foregroundcolor red
                # $list_of_offline_computers.add($single_computer) | Out-Null
                $ping_response_obj.Status = 'offline'
            }

            $results.add($ping_response_obj) | Out-Null
        }
    
    }
    ## Open results in gridview since this is just supposed to be quick test for connectivity
    $results | out-gridview -Title "Results: $PingCount Pings"
}

Export-ModuleMember -Function *-*
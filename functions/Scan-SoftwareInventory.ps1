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
        $AppsToLookFor
    )
    ## 1. Define title, date variables
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)
    ## 3. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    ## 4. Create empty results container
    $AppsToLookFor = $AppsToLookFor.split(",")
    if ($AppsToLookFor -isnot [array]) {
        $AppsToLookFor = @($AppsToLookFor)
    }

    ## 1. Define title, date variables
    $thedate = Get-Date -Format 'yyyy-MM-dd'
    ## 2. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)

    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    $ComputerName = $ComputerName | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }
        
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
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
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


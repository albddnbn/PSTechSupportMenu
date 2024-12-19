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

    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = Test-Connectivity -ComputerName $ComputerName
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
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }
        
    ## Collecting the results
    $results = [System.Collections.ArrayList]::new()
    ## 1/2. Check Targetcomputer for null/empty values and test ping.
    ## 3. If machine was responsive, check for file/folder or application, add to $results.
    ##    --> If searching for filepaths - creates object with some details / file attributes
    ##    --> If searching for apps - creates object with some details / app attributes

    ## 1.

    ## 2. Test with ping first:
    ## File/Folder search
    if (@('path', 'file', 'folder') -contains $SearchType.ToLower()) {

        $results = Invoke-Command -ComputerName $single_computer -ScriptBlock {
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

        $results = Invoke-Command -ComputerName $single_computer -Scriptblock {
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
            if ($null -eq $obj) {
                $obj = [PSCustomObject]@{
                    ComputerName    = $single_computer
                    AppName         = "No matching apps found for $using:Item"
                    AppVersion      = $null
                    InstallDate     = $null
                    InstallLocation = $null
                    Publisher       = $null
                    UninstallString = "No matching apps found"
                }
                $obj
            }
        } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername

        # $search_result
        # read-host "enter"
    }

    ## errored out invoke-commands:
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

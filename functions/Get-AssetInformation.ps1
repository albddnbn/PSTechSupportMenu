function Get-AssetInformation {
    <#
    .SYNOPSIS
        Attempts to use Dell Command Configure to get asset tag, if not available uses built-in powershell cmdlets.

    .DESCRIPTION
        Function will work as a part of the Terminal menu or outside of it.

    .PARAMETER TargetComputer
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER OutputFile
        'n' = terminal output only
        Entering anything else will create an output file in the 'reports' directory, in a folder with name based on function name, and OutputFile input.
        Ex: Outputfile = 'A220', output file(s) will be in $env:PSMENU_DIR\reports\AssetInfo - A220\

    .INPUTS
        [String[]] - an array of hostnames can be submitted through pipeline for Targetcomputer parameter.

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
            Mandatory = $true,
            ValueFromPipeline = $true,
            Position = 0
        )]
        $ComputerName,
        [string]$Outputfile
    )
    ## 1. Handling TargetComputer input if not supplied through pipeline.
    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
    ## 3. Define scriptblock that retrieves asset info from local computer.
    ## 4. Create empty results container.
    $thedate = Get-Date -Format 'yyyy-MM-dd'

    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    $ComputerName = $ComputerName | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }
   
    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
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
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }


    ## 3. Asset info scriptblock used to get local asset info from each target computer.
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
        $monitors = Get-CimInstance WmiMonitorId -Namespace root\wmi -ComputerName $ComputerName | Select Active, ManufacturerName, UserFriendlyName, SerialNumberID, YearOfManufacture
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

    $results = Invoke-Command -ComputerName $ComputerName -ScriptBlock $asset_info_scriptblock `
    | Select PSComputerName, * -ExcludeProperty RunspaceId, PSshowcomputername -ErrorAction SilentlyContinue

    ## 1. If there were any results - output them to terminal and/or report files as necessary.
    if ($results) {
        ## Sort the results
        $results = $results | sort -property pscomputername
        if ($outputfile.tolower() -eq 'n') {
            $results | out-gridview -Title $str_title_var
        }
        else {
            $results | Export-Csv -Path "$outputfile.csv" -NoTypeInformation
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

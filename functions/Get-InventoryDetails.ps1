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

    .INPUTS
        [String[]] - an array of hostnames can be submitted through pipeline for Targetcomputer parameter.

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
        [string]$Outputfile
    )

    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    #$ComputerName = $ComputerName | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }

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
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    $results = Invoke-Command -ComputerName $single_computer -scriptblock {
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
    # read-host "Press enter to return results."
    return $results
}

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

    .INPUTS
        [String[]] - an array of hostnames can be submitted through pipeline for Targetcomputer parameter.

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
        [string]$Outputfile
    )
    ## 1. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 2. Create output filepath if necessary.
    ## 3. Create empty results arraylist to hold results from each target machine (collected during the PROCESS block).
    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    if ($SendPings -eq 'y') {
        $ComputerName = Test-Connectivity -ComputerName $ComputerName
    }

    ## 2. Outputfile handling - either create default, create filenames using input, or skip creation if $outputfile = 'n'.
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
        $OutputFile = Get-OutputFileString -TitleString $REPORT_DIRECTORY -Rootdirectory $env:USERPROFILE\PSTechSupportMenu -FolderTitle $REPORT_DIRECTORY -ReportOutput
    }

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responseive, run scriptblock to logged in user, info on teams/zoom processes, etc.
    $results = Invoke-Command -ComputerName $ComputerName -Scriptblock {
        $obj = [PSCustomObject]@{
            Model        = (get-ciminstance -class win32_computersystem).model
            CurrentUser  = (get-process -name 'explorer' -includeusername -erroraction silentlycontinue).username
            TeamsRunning = $(if (Get-PRocess -Name 'Teams' -ErrorAction SilentlyContinue) { $true } else { $false })
            ZoomRunning  = $(if (Get-PRocess -Name 'Zoom' -ErrorAction SilentlyContinue) { $true } else { $false })

        }
        $obj
    } -ErrorVariable RemoteError | Select * -ExcludeProperty RunspaceId, PSshowcomputername
    
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## 1. If there are results - sort them by the hostname (pscomputername) property.
    ## 2. If the user specified 'n' for outputfile - just output to terminal or gridview.
    ## 3. Create .csv/.xlsx reports as necessary.
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
    # read-host "Press enter to return results."
    return $results
}

function Send-Files {
    <#
    .SYNOPSIS
        Sends a target file/folder from local computer to target path on remote computers.

    .DESCRIPTION
        You can enter both paths as if they're on local filesystem, the script should cut out any drive letters and insert the \\hostname\c$ for UNC path. The script only works for C drive on target computers right now.

    .PARAMETER SourcePath
        The path of the file/folder you want to send to target computers. 
        ex: C:\users\public\desktop\test.txt, 
        ex: \\networkshare\folder\test.txt

    .PARAMETER DestinationPath
        The path on the target computer where you want to send the file/folder. 
        The script will cut off any preceding drive letters and insert \\hostname\c$ - so destination paths should be on C drive of target computers.
        ex: C:\users\public\desktop\
        The path should be a destination folder, script should create directories/subdirectories as required.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: g-labpc- will create a list of all hostnames that start with 
        g-labpc- (g-labpc-01. g-labpc-02, g-labpc-03..).

    .EXAMPLE
        copy the test.txt file to all computers in stanton open lab
        Send-Files -sourcepath "C:\Users\Public\Desktop\test.txt" -destinationpath "Users\Public\Desktop" -targetcomputer "t-client-"

    .EXAMPLE
        Get-User -ComputerName "t-client-28"

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
        [ValidateScript({
                if (Test-Path $_ -ErrorAction SilentlyContinue) {
                    return $true
                }
                else {
                    Write-Error "SourcePath does not exist. Please enter a valid path."
                    return $false                
                }
            })]
        [string]$sourcepath,
        [string]$destinationpath
    )
    ## 1. Handle Targetcomputer input if it's not supplied through pipeline.
    ## 1. Handle TargetComputer input if not supplied through pipeline (will be $null in BEGIN if so)

    $ComputerName = Get-Targets -TargetComputer $ComputerName

    ## Ping Test for Connectivity:
    $ComputerName = $ComputerName | Where-Object { Test-Connection -ComputerName $_ -Count 1 -Quiet }
        

    $informational_string = ""

    ## make sure destination path is formatted correctly:
    # if destination path is a folder path, add the source file name to the end of the path
    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. Use session to copy file from local computer.
    ##    Report on success/fail
    ## 4. Remove the pssession.
    ForEach ($single_computer in $ComputerName) {
        ## 1. no empty Targetcomputer values past this point
        if ($single_computer) {
            ## 2. Make sure machine is responsive on network
            $file_copied = $false
            if ([System.IO.Directory]::Exists("\\$single_computer\c$")) {
                ## make sure target directory exists:
                Invoke-Command -ComputerName $single_computer -Scriptblock {

                    $destination_folder = $using:destinationpath
                    if (-not (Test-Path $destination_folder -ErrorAction SilentlyContinue)) {
                        Write-Host "Creating $destination_folder directory on $env:COMPUTERNAME."
                        New-Item -Path $destination_folder -ItemType Directory -Force | Out-Null
                    }
                }

                $target_session = New-PSSession $single_computer

                Copy-Item -Path "$sourcepath" -Destination "$destinationpath" -ToSession $target_session -Recurse
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $sourcepath to $destinationpath ($single_computer) complete." -foregroundcolor green
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Transfer of $sourcepath to $destinationpath ($single_computer) complete.`n"

                $file_copied = $true


                Remove-PSSession $target_session
                            
            }
            
            if (-not $file_copied) {
                # Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $sourcepath to $destinationpath on $single_computer." -foregroundcolor red
                $informational_string += "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to copy $sourcepath to $destinationpath on $single_computer.`n"
            }
            
        }
    }
    ## 1. Write an ending message to terminal.
    ## announcement file for when function is run as background job
    if (-not $env:PSMENU_DIR) {
        $env:PSMENU_DIR = $(pwd)
    }
    ## create simple output path to reports directory
    $thedate = Get-Date -Format 'yyyy-MM-dd'
    $DIRECTORY_NAME = 'SendFiles'
    $OUTPUT_FILENAME = 'SendFiles'
    if (-not (Test-Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ErrorAction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME" -ItemType Directory -Force | Out-Null
    }
        
    $counter = 0
    do {
        $output_filepath = "$env:PSMENU_DIR\reports\$thedate\$DIRECTORY_NAME\$OUTPUT_FILENAME-$counter.txt"
        $counter++
    } until (-not (Test-Path $output_filepath -ErrorAction SilentlyContinue))
        
        
    ## Append text to file here:
    $informational_string | Out-File -FilePath $output_filepath -Append
    # $ComputerName | Out-File -FilePath $output_filepath -Append
    "`nThe Scan-ForApporFilepath function can be used to verify file/folders' existence on target computers." | Out-File -FilePath $output_filepath -Append -Force
        
    ## then open the file:
    Invoke-Item "$output_filepath"
        
    # read-host "`nPress [ENTER] to continue."
}

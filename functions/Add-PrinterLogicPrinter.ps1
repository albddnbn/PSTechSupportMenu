function Add-PrinterLogicPrinter {
    <#
    .SYNOPSIS
        Connects local or remote computer to target printerlogic printer by executing C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin\PrinterInstallerConsole.exe.
        Connection fails if PrinterLogic Client software is not installed on target machine(s).
        The user does NOT have to be a PrinterLogic user to be able to access connected PrinterLogic printers.

    .DESCRIPTION
        PrinterLogic Client software has to be installed on target machine(s) and connecting to your organization's 'Printercloud' instance using the registration key.

    .PARAMETER ComputerName
        Target computer or computers of the function.
        Single hostname, ex: 't-client-01' or 't-client-01.domain.edu'
        Path to text file containing one hostname per line, ex: 'D:\computers.txt'
        First section of a hostname to generate a list, ex: t-pc-0 will create a list of all hostnames that start with t-pc-0. (Possibly t-pc-01, t-pc-02, t-pc-03, etc.)

    .PARAMETER PrinterName
        Name of the printer in printerlogic. Ex: 't-prt-lib-01', you can use the name or the full 'path' to the printer, ex: 'STANTON\B WING..\s-prt-b220-01'
        Name must match the exact 'name' of the printer, as listed in PrinterLogic.

    .PARAMETER SendPings
        'y' = Ping test for connectivity before attempting main purpose of function.
        Anything else - will not conduct the ping test.

    .INPUTS
        [String[]] - an array of hostnames can be submitted through pipeline for Targetcomputer parameter.

    .OUTPUTS
        [System.Collections.ArrayList] - Returns an arraylist of objects containing hostname, printer name, connection status, and client software status.
        The results arraylist is also displayed in a GridView.

    .EXAMPLE
        Connect single remote target computer to t-prt-lib-01 printer:
        Add-PrinterLogicPrinter -TargetComputer "t-client-28" -PrinterName "t-prt-lib-01"
        
    .EXAMPLE
        Connect a group of computers using hostname txt filepath to t-prt-lib-02 printer:
        Add-PrinterLogicPrinter -TargetComputer "D:\computers.txt" -PrinterName "t-prt-lib-02"

    .EXAMPLE
        Submit target computer(s) by hostname through pipeline and add specified printerlogic printer.
        @('t-client-01', 't-client-02', 't-client-03') | Add-PrinterLogicPrinter -PrinterName "t-prt-lib-03"
        
    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param (
        [Parameter(
            Mandatory = $true, ValueFromPipeline = $true
        )]
        [String[]]$ComputerName,
        [string]$PrinterName,
        $SendPings
    )
    ## 1. TargetComputer handling if not supplied through pipeline
    ## 2. Scriptblock to connect to printerlogic printer
    ## 3. Results containers for overall results and skipped computers.
    $ComputerName = Get-Targets -TargetComputer $ComputerName

    if ($SendPings -eq 'y') {
        $ComputerName = Test-Connectivity -ComputerName $ComputerName
    }

    $thedate = Get-Date -Format 'yyyy-MM-dd'

        
    ## 2. Define the scriptblock that connects machine to target printer in printerlogic cloud instance
    $connect_to_printer_block = {
        param(
            $printer_name
        )

        $obj = [pscustomobject]@{
            hostname       = $env:COMPUTERNAME
            printer        = $printer_name
            connectstatus  = 'NO'
            clientsoftware = 'NO'
        }
        # get installerconsole.exe
        $exepath = get-childitem -path "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin" -Filter "PrinterInstallerConsole.exe" -File -Erroraction SilentlyContinue
        if (-not $exepath) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: PrinterLogic PrinterInstallerConsole.exe was not found in C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin." -Foregroundcolor Red
            return $obj
        }
        
        $obj.clientsoftware = 'YES'
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Found $($exepath.fullname), mapping $printer_name now..."
        $map_result = (Start-Process "$($exepath.fullname)" -Argumentlist "InstallPrinter=$printer_name" -Wait -Passthru).ExitCode
        
        # 0 = good, 1 = bad
        if ($map_result -eq 0) {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: Connected to $printer_name successfully." -Foregroundcolor Green
            # Write-Host "*Remember that this script does not set default printer, user has to do that themselves."
            $obj.connectstatus = 'YES'
            return $obj
        }
        else {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$env:COMPUTERNAME] :: failed to connect to $printer_name." -Foregroundcolor Red
            return $obj
        }
    }


    $results = Invoke-Command -ComputerName $single_computer -scriptblock $connect_to_printer_block -ArgumentList $PrinterName -ErrorVariable RemoteError 

    ## errored out invoke-commands:
    $errored_machines = $RemoteError.CategoryInfo.TargetName

    ## 1. Make sure no $null or empty values are submitted to the ping test or scriptblock execution.
    ## 2. Ping the single target computer one time as test before attempting remote session.
    ## 3. If machine was responsive, run scriptblock to attempt connection to printerlogic printer, save results to object

    ## 1. If there are any results - output to file
    ## 2. Output missed computers to terminal.
    ## 3. Return results arraylist
    ## 1.
    if ($results) {

        $results | out-gridview -Title "PrinterLogic Connect Results"

    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output."

        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: No results to output." | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append
        "These computers did not respond to ping:" | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append
        $errored_machines | Out-File -FilePath "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt" -Append

        Invoke-Item "$env:PSMENU_DIR\reports\PrinterLogicConnectResults_$thedate.txt"
    }
    ## 2. Output unresponsive computers
    Write-Host "These computers did not respond to ping:"
    Write-Host ""
    $errored_machines
    Write-Host ""
    # read-host "Press enter to return results."

    ## 3. return results arraylist
    return $results
}
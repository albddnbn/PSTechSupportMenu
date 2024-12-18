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
        Test-ConnectivityQuick -TargetComputer "t-client-"

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
    $PING_COUNT = $PingCount

    # $list_of_online_computers = [system.collections.arraylist]::new()
    # $list_of_offline_computers = [system.collections.arraylist]::new()
    $online_results = [system.collections.arraylist]::new()

    ## Ping target machines $PingCount times and log result to terminal.
    ForEach ($single_computer in $ComputerName) {
        if (Test-Connection $single_computer -Count $PING_COUNT -Quiet) {
            Write-Host "$single_computer is online." -ForegroundColor Green
            
            $online_results.Add($single_computer) | Out-Null
        }
        else {
            Write-Host "$single_computer is offline." -ForegroundColor Red
        }
    
    }

    return $online_results

}

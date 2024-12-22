function Copy-CannedResponse {
    <#
    .SYNOPSIS
        Pulls files from the ./canned_responses directory, and presents menu to user for selection of canned response.
        User is then prompted for values for any variables in the selected response, and the response is copied to the clipboard.

    .DESCRIPTION
        Canned reponses can contain multiple variables indicated by (($variable$)).
        This function cycles through a list of unique variables in the selected response, and prompts the user to enter values for each one.
        The function inserts these values into the original response in place of the variable names, and copies the canned response to the user's clipboard.

    .EXAMPLE
        Copy-CannedResponse

    .NOTES
        One possible use for this function and it's companion function New-CannedResponse is to create canned responses for use in a ticket system, like OSTicket.
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>

    
    ## 1. Creates a list of all .txt and .html files in the canned_responses directory.
    $CannedResponses = Get-ChildItem -Path "$env:CANNED_RESPONSES_DIRECTORY" -Include *.txt, *.html -File -Recurse -ErrorAction SilentlyContinue
    if (-not $CannedResponses) {
        Write-Host "No canned responses found in $env:CANNED_RESPONSES_DIRECTORY" -ForegroundColor Red
        return
    }

    ## 2. If created by New-CannedResponse, the files will have names with words separated by dashes (-).
    ##    - To offer a menu of options - the filenames are split at all occurences of - by replacing them with spaces.
    $CannedResponseOptions = $CannedResponses | ForEach-Object { $_.BaseName -replace '-', ' ' } | Sort-Object

    ## 3. Checks for PS-Menu module necessary to display interactive terminal menu.
    ##    - if not found - checks for nuget / tries to install ps-menu
    if (-not (Get-Module -Name PS-Menu -ListAvailable)) {
        Write-Host "Installing PS-Menu module..." -ForegroundColor Yellow
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable)) {
            Write-Host "Installing NuGet package provider..." -ForegroundColor Yellow
            Install-PackageProvider -Name NuGet -MinimumVersion
        }
        Install-Module -Name PS-Menu -Force
    }
    Import-Module -Name PS-Menu -Force

    # presents menu
    $chosen_response = Menu $CannedResponseOptions

    # reconstructs filenames, re-inserting dashes for spaces
    $chosen_response = $chosen_response -replace ' ', '-'

    ## 4. Get the content of the correct file - using original list of .html/.txt file objects
    $chosen_response_file = $CannedResponses | Where-Object { $_.BaseName -eq $chosen_response }
    $chosen_response_content = Get-Content "$($chosen_response_file.fullname)"

    ## 5. Get the variables from the content - variables are enclosed in (($variable$))
    ##    - for loop cycles through each unique variable, prompting the user for values.

    $variables = $chosen_response_content | Select-String -Pattern '\(\(\$.*\$\)\)' -AllMatches | ForEach-Object { $_.Matches.Value } | Sort-Object -Unique
    ForEach ($single_variable in $variables) {
        $formatted_variable_name = $single_variable -split '=' | select -first 1
        # $formatted_variable_name = $formatted_variable_name.replace('(($', '')

        ForEach ($str_item in @('(($', '$))')) {
            $formatted_variable_name = $formatted_variable_name.replace($str_item, '')
        }


        if ($single_variable -like "*=*") {
            $variable_description = $single_variable -split '=' | select -last 1
            $variable_description = $variable_description.replace('$))', '')
        }
        Write-Host "Description: " -nonewline -foregroundcolor yellow
        Write-host "$variable_description"
        $variable_value = Read-Host "Enter value for $formatted_variable_name"
        Write-Host "Replacing $single_variable with $variable_value"
        $chosen_response_content = $chosen_response_content.replace($single_Variable, $variable_value)
    }
    ## 6. Copy chosen response to clipboard, with new variable values inserted
    $chosen_response_content | Set-Clipboard
    Write-Host "Canned response copied to clipboard." -ForegroundColor Green
}

# function Copy-SnippetToClipboard {
#     <#
#     .SYNOPSIS
#         Uses .txt files found in the ./snippets directory to present terminal menu to user.
#         User selects item to be copied to clipboard.

#     .DESCRIPTION
#         IT Helpdesk Contact Info.txt                  --> Campus help desk phone numbers.
#         OSTicket - Big Number Ordered List.txt        --> HTML/CSS for an ordered list that uses big numbers, horizontal to text.
#         PS1 1-Liner - Get Size of Folder.txt          --> Powershell one-liner to get size of all items in folder.
#         PS1 1-Liner - Last Boot Time.txt              --> Powershell one-liner to get last boot time of computer.
#         PS1 1-Liner - List All Apps.txt               --> Powershell one-liner to list all installed applications.
#         PS1 Code - List App Info From Reg.txt         --> Powershell code to list application info from registry.
#         PS1 Code - PS App Deployment Install Line.txt --> Powershell PSADT module silent installation line.

#     .EXAMPLE
#         Copy-SnippetToClipboard
        
#     .NOTES
#         ---
#         Author: albddnbn (Alex B.)
#         Project Site: https://github.com/albddnbn/PSTerminalMenu
#     #>

#     ## 1. Checks for PS-Menu module necessary to display interactive terminal menu.
#     ##    - if not found - checks for nuget / tries to install ps-menu
#     if (-not (Get-Module -Name PS-Menu -ListAvailable)) {
#         Write-Host "Installing PS-Menu module..." -ForegroundColor Yellow
#         if (-not (Get-PackageProvider -Name NuGet -ListAvailable)) {
#             Write-Host "Installing NuGet package provider..." -ForegroundColor Yellow
#             Install-PackageProvider -Name NuGet -MinimumVersion -Force
#         }
#         Install-Module -Name PS-Menu -Force
#     }
#     Import-Module -Name PS-Menu -Force | Out-Null

#     ## 2. Creates a list of filenames from the ./snippets directory that end in .txt (resulting filenames in list will
#     ##    not include the .txt extension because of the BaseName property).
#     $snippet_options = Get-ChildItem -Path "$env:PSMENU_DIR\snippets" -Include *.txt -Recurse -ErrorAction SilentlyContinue | ForEach-Object { $_.BaseName }
#     $snippet_choice = menu $snippet_options

#     ## 3. Get the content of the chosen file and copy to clipboard
#     Get-Content -Path "$env:PSMENU_DIR\snippets\$snippet_choice.txt" | Set-Clipboard
#     Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Content from $env:PSMENU_DIR\snippets\$snippet_choice.txt copied to clipboard."
# }

function New-CannedResponse {
    <#
    .SYNOPSIS
        Allows user to create a txt/html file, that has variables defined. 
        When the Copy-CannedResponse function is used on this newly created file - it will prompt user for values for any variabels before copying the text to clipboard.

    .DESCRIPTION
        The user will have to enter variables of their own, defined in the text like this: (($variable_name$)).
    
    .PARAMETER FileType
        Allows user to create either a TXT or HTML canned response file. With HTML files - the user will have to click the <> (code) button in OSTicket before pasting in the canned response HTML.
        TXT files can be directly pasted without clicking the <> button.
        Options include: txt, html

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param(
        [string]$FileType
    )
    Write-Host "Canned Response Instructions:"
    Write-Host "-----------------------------"
    Write-Host "Variables are defined like this: " -nonewline -foregroundcolor Yellow
    Write-Host "((`$variable_name_here`$)), Get-CannedResponse will pick them up and prompt you for their values before copying the response content to your clipboard."
    Write-Host "You should be able to use any html tags that work correctly in OSTicket, for ex: <b>text</b> for bold text."
    Write-Host "Enter response content, a blank line will end the input session." -Foregroundcolor Green
    # source: https://stackoverflow.com/questions/36650961/use-read-host-to-enter-multiple-lines / majkinetor
    $response_content = while (1) { read-host | set r; if (!$r) { break }; $r }

    if ($FileType.ToLower() -eq 'html') {

        # create html:
        $html_template = @"
<html>`n
<body>`n

"@
        ForEach ($single_line in $response_content) {
            $html_string = "<p>$single_line</p>`n"
            $html_template += $html_string
        }
        $html_template += @"

</body>`n
</html>
"@
    }
    elseif ($FileType.ToLower() -eq 'txt') {
        # $html_template variable is eventually output to file, so set it to equal just the text content if user chose txt
        $html_template = $response_content
    }
    # write to file - extension will be .html even if its just text content to keep things simple
    do {
        $filename = Read-Host "Enter filename for canned response, no extension needed."
        if ($filename -notlike "*.html") {
            $filename = "$filename.html"
        }
    } until (-not (Test-Path "$env:CANNED_RESPONSES_DIRECTORY\$filename" -ErrorAction SilentlyContinue))
    # Set-Content will keep formatting / newlines - Out-File will not as-is
    Set-Content -Path "$env:CANNED_RESPONSES_DIRECTORY\$filename" -Value $html_template
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ::Canned response file created at canned_responses\$filename" -ForegroundColor Green

    Invoke-Item "$env:PSMENU_DIR\canned_responses\$filename"

}

function New-PSADTFolder {
    <#
    .SYNOPSIS
        Downloads specified (hopefully latest) version of Powershell App Deployment Toolkit and creates directory structure for application deployment.

    .DESCRIPTION
        Creates basic directory structure for PSADT deployment after downloading specified PSADT.zip folder from Github.

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>

    $PSADT_DOWNLOAD_URL = "https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/releases/download/3.10.2/PSAppDeployToolkit_3.10.2.zip"
    ## Downloads the PSADT Toolkit and extracts it
    ## Renames Toolkit folder to the given Application Name.

    $ApplicationName = Read-Host "Enter the name of the application"
    $DeploymentFolder_Destination = Read-Host "Enter destination for PSADT folder (press enter for default: $env:USERPROFILE\$APPLICATIONNAME)"
    if ($DeploymentFolder_Destination -eq '') {
        $DeploymentFolder_Destination = "$env:USERPROFILE\$APPLICATIONNAME"
    }

    ## create destination directory if not exist
    if (-not (Test-Path $DeploymentFolder_Destination -PathType Container -Erroraction SilentlyContinue)) {
        New-Item -Path $DeploymentFolder_Destination -ItemType Directory | out-null
    }

    ## Download the PSADT .zip folder (current latest):
    Invoke-WebRequest "$PSADT_DOWNLOAD_URL" -Outfile "$env:USERPROFILE\PSADT.zip"

    Expand-Archive -Path "$env:USERPROFILE\PSADT.zip" -DestinationPath "$env:USERPROFILE\PSADT"

    ## Get the toolkit folder:
    $toolkit_folder = Get-ChildItem -Path "$env:USERPROFILE\PSADT" -Filter "Toolkit" -Directory | Select -Exp FullName

    Copy-Item -Path "$toolkit_folder\*" -Destination "$DeploymentFolder_Destination\"

    ## Delete everything not needed:
    REmove-Item -Path "$env:USERPROFILE\PSADT*" -Recurse -Force -ErrorAction SilentlyContinue

    ## Rename the Deploy-Application.ps1 file
    Rename-Item -Path "$DeploymentFolder_Destination\Deploy-Application.ps1" -NewName "Deploy-$ApplicationName.ps1"
}

function Open-Guide {
    
    <#
    .SYNOPSIS
        Opens PSTerminalMenu Wiki in browser.

    .DESCRIPTION
        The guide explains the menu's directory structure, functionality, and configuration.

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>

    $HELP_URL = "https://github.com/albddnbn/PSTerminalMenu/wiki"
    Write-Host "Attempting to open " -nonewline
    Write-Host "$HELP_URL" -Foregroundcolor Yellow -NoNewline
    Write-Host " in default browser."


    try {
        $chrome_exe = Get-ChildItem -Path "C:\Program Files\Google\Chrome\Application" -Filter "chrome.exe" -File -ErrorAction SilentlyContinue
        Start-Process "$($chrome_exe.fullname)" -argumentlist "$HELP_URL"
    }
    catch {
        Start-Process "msedge.exe" -argumentlist "$HELP_URL"
    }
    Read-Host "Press enter to continue."
}

function Open-RDP {
    <#
    .SYNOPSIS
        Basic function to open RDP window with target computer and _admin username inserted.

    .PARAMETER SingleTargetComputer
        Single hostname of target computer - script will open RDP window with _admin username and target computer hostname already inserted.

    .NOTES
        Function is just meant to save a few seconds.
    #>
    param(
        $SingleTargetComputer
    )
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Opening RDP session to " -Nonewline
    Write-Host "$SingleTargetComputer" -ForegroundColor Green
    mstsc /v:$($singletargetcomputer):3389
}



function Scan-Inventory {
    <#
    .SYNOPSIS
        Helps to generate an inventory spreadsheet of items, their current stock values, along with other item details.
        The script tries to link scanned input to an item in 3 ways, in this order: 
        1. currently scanned item list
        2. master item list
        3. API lookup

    .DESCRIPTION
        If the script fails to automatically link scanned input to an item, and also fails to find the item details 
        through the API - it will prompt the user to manually enter item details. The item object will then be added to 
        the currently scanned items list (and also master item list).

    .PARAMETER ScanTitle
        Title of the scan, will be used to name the output file(s).
    
    .PARAMETER ScanList
        Absolute path to a file containing codes scanned using barcode scanner.
        You can scan all codes to a text file and then submit them as a list to speed things up.
        Note to add catch in the list scanning - functionality so user can make sure items don't look off.
        For Ex: I've scanned some printer-related items, and RARELY had a wierd item returned, like Denim Jeans.
        This isn't common, especially for major brand toner-related items - Dell, HP, Lexmark, Brother.

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param(
        [String]$ScanTitle = 'Inventory',
        [string]$ScanList
    )
    # if (-not (Get-Module -Name 'PSADT' -ListAvailable)) {
    #     Install-Module -Name 'PSADT' -Force
    # }
    # ipmo psadt


    ## Creates item with specified properties (in parameters) and then prompts user to keep or change each value, 
    ## returns changed object
    Function Create-NewItemObject {
        param(
            $Description = '',
            $Compatible_printers = '',
            $code = '',
            $UPC = '',
            $Stock = '',
            $Part_num = '',
            $color = '',
            $brand = '',
            $yield = '',
            $type = '',
            $comments = '',
            $ean = '',
            $Seek_input = $true
        )
        ## "description","compatible_printers","code","upc","stock","part_num","color","brand","yield","type","comments","ean"
        $obj = [pscustomobject]@{
            Description         = $Description
            Compatible_printers = $Compatible_printers
            code                = $code
            UPC                 = $UPC
            Stock               = $Stock
            # need to see how i can parse this
            Part_num            = $Part_num
            color               = $color
            brand               = $brand
            yield               = $yield
            type                = $type
            comments            = $comments
            ean                 = $ean
        }
        if (-not $Seek_input) {
            return $obj
        }
        # These property values are SKIPPED when item is scanned / prompting for new values:
        $SKIPPED_PROPERTY_VALUES = @('upc', 'type', 'Compatible_printers', 'ean', 'yield', 'color', 'stock')
        $obj.PSObject.Properties | ForEach-Object {
            if ($_.Name -notin $SKIPPED_PROPERTY_VALUES) {
                
                $current_value = $_.Value
                Write-Host "Current value of $($_.Name) is: $($_.Value), enter 'k' to retain value."
                $_.Value = Read-Host "Enter value for $($_.Name)"
                if ($_.Value -eq 'k') {
                    $_.Value = $current_value
                }
            }
        }
        return $obj
    }

    ## --- SCRIPT VARIABLES ---
    $MASTER_CSV_FILE = "$env:PSMENU_DIR\inventory\master-inventory.csv" # holds all items that have been scanned in, for all time
    $API_URL = "https://api.upcitemdb.com/prod/trial/lookup?upc="

    # SOUNDS for success/failure at linking scanned items to existing or api items.
    $POSITIVE_BEEP_PATH = "$env:PSMENU_DIR\inventory\sounds\positivebeep.wav"
    $NEGATIVE_BEEP_PATH = "$env:PSMENU_DIR\inventory\sounds\negativebeep.wav"

    # making sure master .csv file exists
    if (-not (Test-path "$env:PSMENU_DIR\inventory\master-inventory.csv" -erroraction SilentlyContinue)) {
        New-Item -Path "$env:PSMENU_DIR\inventory\master-inventory.csv" -ItemType 'File' -Force | Out-Null
    }
    # New-File -Path "$env:PSMENU_DIR\inventory\master-inventory.csv"
    # get filesize of master csv file
    $reply = Read-Host "Import items from $MASTER_CSV_FILE ? (y/n)"
    if ($reply.tolower() -eq 'y') {
        try {
            $MASTER_ITEM_LIST = import-csv $MASTER_CSV_FILE
        }
        catch {
            Write-Host "Failed to import $MASTER_CSV_FILE, continuing in 5 seconds.." -foregroundcolor yellow
            Start-Sleep -Seconds 5        
        }
    }
    if (Get-Command -Name 'Get-Outputfilestring' -ErrorAction SilentlyContinue) {
        # create output filename for inventory report (this is different from the master list file, this is the report being generated by this current/run of the function)
        $outputfile = Get-OutputFileString -Titlestring $ScanTitle -rootdirectory $env:PSMENU_DIR -foldertitle $REPORT_DIRECTORY -reportoutput
    }
    else {
        $outputfile = "Inventory-$ScanTitle"
    }
    Write-Host "Output file(s) will be at: $outputfile .csv and/or .xlsx" -foregroundcolor green
    
    ## Holds items scanned during this run of the function, will be output to report .csv/.xlsx
    $CURRENT_INVENTORY = [System.collections.arraylist]::new()

    ## add items from master item list to current inventory, set stock to 0
    $MASTER_ITEM_LIST | ForEach-Object {
        $_.stock = '0'
        $CURRENT_INVENTORY.Add($_) | Out-Null
    }

    Function Get-ScannedItems {
        param(
            $scanned_codes,
            $inventory_list
        )

        $missed_items = [System.Collections.ArrayList]::new()

        # if api lookup can't recognize a code, the code won't be looked up again.
        $invalid_upcs = [System.Collections.ArrayList]::new()
        $API_LOOKUP_LIMIT_REACHED = $false
        ForEach ($scanned_code in $scanned_codes) {
            write-host "Checking: $scanned_code"
            $ITEM_FOUND = $false
            # CHECK CURRENT INVENTORY ARRAYLIST
            $inventory_list | ForEach-Object {
                ## Search upc codes
                # write-host "checking $($_.upc), $($_.description)"
                # $upctest = $_.upc
                # $upctest.gettype().name
                if ($scanned_code -eq ([string]$($_.upc))) {
                    Write-Host "Found match for $scanned_code in upc/ean columns of current inventory, increasing stock by 1."
                    $player = New-Object System.Media.SoundPlayer
                    $player.SoundLocation = $POSITIVE_BEEP_PATH
                    $player.Load()
                    $player.Play()

                    $matched_item_index = $inventory_list.IndexOf($_)
                    $matched_item = $inventory_list[$matched_item_index]
                    Write-Host "Set matched item to: $($matched_item.description)"

                    $inventory_list | ForEach-Object {
                        if ($matched_item.description -eq $_.description) {
                            $_.stock = [string]$([int]$_.stock + 1)
                            Write-Host "`rIncreased stock of $($_.description) by 1, to: $($_.stock).`n"
                        }
                    }
                    $ITEM_FOUND = $true
                }
            }

            ## Continue to next iteration / scanned item if item was found
            if ($ITEM_FOUND) {
                continue
            }
            else {
                ## If the API has already said the code is invalid, or returned an 'exceeded limit' response
                ## - no point in continuing
                if ($scanned_code -in $invalid_upcs) {
                    Write-Host "Skipping $scanned_code, it was already determined to be an invalid upc." -foregroundcolor yellow
                    $missed_items.add($scanned_code) | Out-Null
                    continue
                }
                elseif ($API_LOOKUP_LIMIT_REACHED) {
                    Write-Host "API LOOKUP LIMIT REACHED FOR THE DAY! (~100 for free last checked)" -Foregroundcolor red
                    $missed_items.add($scanned_code) | Out-Null
                    continue
                }
                ## IF - the user input has a space or dash - it's not gonna work, no point continuing. 
                elseif (($scanned_code -like "* *") -or ($scanned_code -like "*-*")) {
                    Write-Host "Skipping $scanned_code, it contains a space or dash." -foregroundcolor yellow
                    $missed_items.add($scanned_code) | Out-Null
                    continue
                }
                ## IF - the user input has any letters in it - not a upc code
                elseif ($scanned_code -cmatch "^*[A-Z][a-z]*$") { 
                    Write-Host "Skipping $scanned_code, it has letters." -foregroundcolor yellow
                    $missed_items.add($scanned_code) | Out-Null
                    continue
                }

                ## If CODE not 'invalid' / limit not reached - continue with API lookup

                Write-Host "input: $scanned_code not found in current inventory list, checking api.."
                ## API LOOKUP
                ## Play 'BAD BEEP'       
                $player = New-Object System.Media.SoundPlayer
                $player.SoundLocation = $NEGATIVE_BEEP_PATH
                $player.Load()
                $player.Play()
                # if reaches this point, then item is not already in the lists so API has to be checked.
                try {
                    $api_response = Invoke-WebRequest -URI "$API_URL$($scanned_code)"
                }
                catch [System.Net.WebException] {
                    # Handle WebException here
                    Write-Host "API response was for invalud upc or to slow down."
                    ## if the exception error message contains TOO_FAST - start - sleep 15 seconds
                    Write-Host "This is the exception:`n$($_.errordetails)"
                    if ($($_.errordetails) -like "*INVALID_UPC*") {
                        Write-Host "API returned INVALID UPC for: $scanned_code." -Foregroundcolor Yellow
                        $missed_items.add($scanned_code) | Out-Null
                        $invalid_upcs.add($scanned_code) | Out-Null

                        Write-Host "Adding: $scanned_code to list of items that need to be accounted for manually." -foregroundcolor yellow
                    }
                    ## Sending API lookups too fast - slow down for 90s
                    elseif ($($_.errordetails) -like "*TOO_FAST*") {
                        Write-Host "API returned TOO_FAST for: $scanned_code." -Foregroundcolor Yellow
                        $missed_items.add($scanned_code) | Out-Null
                        Write-Host "Slowing down for 90 seconds." -foregroundcolor yellow
                        Start-Sleep -Seconds 90
                    }
                    ## API Lookup limit reached for the day - api lookups won't continue in next loop iterations
                    elseif ($($_.errordetails) -like "*EXCEED_LIMIT*") {
                        Write-Host "API LOOKUP LIMIT REACHED FOR THE DAY! (~100 for free last checked)" -Foregroundcolor red
                        $missed_items.add($scanned_code) | Out-Null
                        Write-Host "Slowing down for 90 seconds." -foregroundcolor yellow
                        $API_LOOKUP_LIMIT_REACHED = $true
                    }
                    continue
                }
                # convert api_response content from json
                $json_response_content = $api_response | ConvertFrom-Json
                # make sure response was 'ok'
                ## if the items.title isn't a system/object (psobject)?
                if (($([string]($json_response_content.code)) -eq 'OK') -and ($json_response_content.items.title)) {
                    Write-Host "API response for $scanned_code was 'OK'" -Foregroundcolor Green
                    $menu_options = $json_response_content.items.title | sort
                    ## "description","compatible_printers","code","upc","stock","part_num","color","brand","yield","type","comments","ean"
                    ## this gets info from first result
                    # present menu - user can choose between multiple item results (theoretically):
                    $chosen_item_name = $Menu_options | select -first 1

                    # get item object - translate it to new object
                    $item_obj = $json_response_content.items | where-object { $_.title -eq $chosen_item_name }

                    ## create item object using default / values from the itenm object
                    $obj = Create-NewItemObject -Description $chosen_item_name -code $item_obj.model -stock '1' -UPC $item_obj.upc -Part_num $item_obj.model -color $item_obj.color -brand $item_obj.brand -ean $item_obj.ean -seek_input $false
                    # read-host "press enter to continue"
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Adding the object to inventory_csv variable.." -ForegroundColor green
                    $inventory_list.add($obj) | Out-Null

                }
                else {
                    $missed_items.add($scanned_code) | Out-Null
                    Write-Host "Adding: $scanned_code to list of items that need to be accounted for manually." -foregroundcolor yellow
                }
            }
        }

        ## Safety export of inventory_list to csv (in case changes messed up how list is passed back out to report creation)
        $inventory_list | export-csv "inventoryscan-$(get-date -format 'yyyy-MM-dd-hh-mm-ss').csv" -NoTypeInformation -Force
        $missed_items | out-file "misseditems-$(get-date -format 'yyyy-MM-dd-hh-mm-ss').csv"
        $return_obj = [pscustomobject]@{
            misseditems   = $missed_items
            inventorylist = $inventory_list
        }
        
        return $return_obj
    }

    if (test-path $Scanlist -erroraction SilentlyContinue) {
        read-host "scan list detected."
        $scanned_codes = Get-Content $Scanlist

        $results = Get-ScannedItems -scanned_codes $scanned_codes -inventory_list $CURRENT_INVENTORY
        $missed_items = $results.misseditems
        # $written_inventory_list = $results.inventorylist
        # # Write-Host "This is the missed items array: $missed_items"
        # # Write-Host "This is inventory list $written_inventory_list"
        # # Read-Host "The results array and inventory lists should be above this."

        if ($missed_items) {

            $datetimestring = get-date -format 'yyyy-MM-dd-hh-mm-ss'
            "These items were missed during scan at: $datetimestring`r" | Out-File "$env:PSMENU_DIR\inventory\missed_items-$datetimestring.txt"
            $missed_items | Out-File "$env:PSMENU_DIR\inventory\missed_items-$datetimestring.txt" -Append
            Write-Host "Items that were not found in the master list or API lookup have been saved to $env:PSMENU_DIR\inventory\missed_items-$datetimestring.txt"
        
            Invoke-item "$env:PSMENU_DIR\inventory\"
        }
        else {
            Write-Host "All items were found in the master list or API lookup." -Foregroundcolor Green
        
        }
        $CURRENT_INVENTORY = $results.inventorylist
    }
    else {
        # Begin 'scan loop' where user scans items.
        ## First - the scanned code is checked against the master item list, 
        ## If it's not found, the script sends a request to the API URL with scanned code to try to get details.
        while ($true) {

            $user_input = Read-Host "Scan code, or type 'exit' to quit"
            if ($user_input.ToLower() -eq 'exit') {
                break
            }
            $ITEM_FOUND = $false

            # CHECK CURRENT INVENTORY ARRAYLIST
            $CURRENT_INVENTORY | ForEach-Object {
                ## Search upc codes
                write-host "checking $($_.upc), $($_.description)"
                $upctest = $_.upc
                $upctest.gettype().name
                if (($user_input -eq ([string]$($_.upc))) -or ($user_input -eq ([string]$($_.ean)))) {
                    Write-Host "Found match for $user_input in upc/ean columns of current inventory, increasing stock by 1."
                    $player = New-Object System.Media.SoundPlayer
                    $player.SoundLocation = $POSITIVE_BEEP_PATH
                    $player.Load()
                    $player.Play()

                    $matched_item_index = $CURRENT_INVENTORY.IndexOf($_)
                    $matched_item = $CURRENT_INVENTORY[$matched_item_index]
                    Write-Host "Set matched item to: $($matched_item.description)"

                    $CURRENT_INVENTORY | ForEach-Object {
                        if ($matched_item.description -eq $_.description) {
                            $_.stock = [string]$([int]$_.stock + 1)
                            Write-Host "`rIncreased stock of $($_.description) by 1, to: $($_.stock).`n"
                        }
                    }
                    $ITEM_FOUND = $true
                }
            }

            ## Continue to next iteration / scanned item if item was found
            if ($ITEM_FOUND) {
                continue
            }
            ## API LOOKUP
            ## Play 'BAD BEEP'       
            $player = New-Object System.Media.SoundPlayer
            $player.SoundLocation = $NEGATIVE_BEEP_PATH
            $player.Load()
            $player.Play()
            # if reaches this point, then item is not already in the lists so API has to be checked.
            $api_response = Invoke-WebRequest -URI "$API_URL$($user_input)"

            # convert api_response content from json
            $json_response_content = $api_response.content | ConvertFrom-Json
            # make sure response was 'ok'
            if ($json_response_content.code -eq 'OK') {
                Write-Host "API response for $user_input was 'OK', here are the results:" -Foregroundcolor Green
                Write-Host ""
                $menu_options = $json_response_content.items.title | sort
                ## "description","compatible_printers","code","upc","stock","part_num","color","brand","yield","type","comments","ean"
                ## this gets info from first result
                $item_details = $json_response_content.items | where-object { $_.title -eq $($menu_options | select -first 1) }
                $useful_details = $item_details | Select title, description, model, upc, color, brand, ean, dimension, category
                # $useful_details = $item_details | Select ean, title, description, upc, brand, model, color, dimension, category

                $useful_details | Format-List

                Write-Host ""

                # present menu - user can choose between multiple item results (theoretically):

                $chosen_item_name = Menu $menu_options

                # get item object - translate it to new object
                $item_obj = $json_response_content.items | where-object { $_.title -eq $chosen_item_name }

                ## create item object using default / values from the itenm object
                $obj = Create-NewItemObject -Description $chosen_item_name -code $item_obj.model -UPC $item_obj.upc -Part_num $item_obj.model -color $item_obj.color -brand $item_obj.brand -ean $item_obj.ean

                read-host "press enter to continue"
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Adding the object to inventory_csv variable.."
                $CURRENT_INVENTORY.add($obj) | Out-Null

                $CURRENT_INVENTORY
                Read-Host "This is current inventory"

            }
            else {

                Write-Host "API was unable to find matching item for: $user_input."
                $response = Menu @('Scan again?', 'Create item manually')
                if ($response -eq 'Scan again?') {
                    continue
                }
                else {
                    $obj = Create-NewItemObject -UPC $user_input
                    $CURRENT_INVENTORY.add($obj) | Out-Null
                }
            }    
        }
    }
    #########################################################
    ## INVENTORY SPREADSHEET CREATION (end of scanning loop):
    #########################################################
    # script creates a .csv / .xlsx report from CURRENT_INVENTORY arraylist
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Exporting scanned inventory to " -NoNewline
    Write-Host "$outputfile.csv and $outputfile.xlsx" -Foregroundcolor Green
    $CURRENT_INVENTORY | Export-CSV "$outputfile.csv" -NoTypeInformation -Force

    # look for the importexcel powershell module
    $CheckForimportExcel = Get-InstalledModule -Name 'importexcel' -ErrorAction SilentlyContinue
    if (-not $CheckForimportExcel) {
        Install-Module -Name ImportExcel -Force
    }

    $params = @{
        AutoSize             = $true
        TitleBackgroundColor = 'Blue'
        TableName            = "$ScanTitle"
        TableStyle           = 'Medium9' # => Here you can chosse the Style you like the most
        BoldTopRow           = $true
        WorksheetName        = $ScanTitle
        PassThru             = $true
        Path                 = "$OutPutFile.xlsx" # => Define where to save it here!
    }

    $xlsx = $CURRENT_INVENTORY | Export-Excel @params
    $ws = $xlsx.Workbook.Worksheets[$params.Worksheetname]
    $ws.View.ShowGridLines = $false # => This will hide the GridLines on your file
    Close-ExcelPackage $xlsx

    # Explorer.exe $env:PSMENU_DIR\reports\$thedate\$REPORT_DIRECTORY
    try {
        Invoke-Item "$outputfile.xlsx"
    }
    catch {
        Write-Host "Failed to open $outputfile.xlsx, opening .csv in notepad." -Foregroundcolor Yellow
        notepad.exe "$outputfile.csv"
    }
    Read-Host "`nPress [ENTER] to continue."
}

Export-ModuleMember -Function *-*


<#
.SYNOPSIS
Reads 'config.json' and creates a menu based on the categories and functions contained within the .json file.

.DESCRIPTION
The user can select any category, and then any function. They'll be prompted for any parameters contained in the 
function, and then the function will attempt to execute based on the input values.

.EXAMPLE
.\Menu.ps1

.NOTES
Alex B. - 08/25/2024
References for menu: https://michael-casey.com/2019/07/03/powershell-terminal-menu-template/ 
PS-Menu module: https://github.com/chrisseroka/ps-menu
PSADT: https://psappdeploytoolkit.com/

#>
function Read-HostNoColon {
    # Read-HostNoColon is just Read-Host except it doesn't automatically add the colon at the end
    param (
        [Parameter(Mandatory = $true)]
        [string]$Prompt
    )
    Write-Host $Prompt -NoNewLine -ForegroundColor Yellow
    return $Host.UI.ReadLine()
}

Try { Set-ExecutionPolicy -ExecutionPolicy 'Unrestricted' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

# Set window title:
$host.ui.RawUI.WindowTitle = "Menu - $(Get-Date -Format 'MM-dd-yyyy')"

########################################################################################################
## CONFIG.JSON file --> found in ./config. It's where the script gets categories/function listings 
## that appear in terminal menu, along with some other configuration variables.
########################################################################################################
$CONFIG_FILENAME = "config.json"

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: " -nonewline
Write-Host "Loading $CONFIG_FILENAME.." -ForegroundColor Yellow

$config_file = Get-Content -Path "$PSSCRIPTROOT\config\config.json" -ErrorAction SilentlyContinue
if (-not $config_file) {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: " -nonewline
    Write-Host "Couldn't find $CONFIG_FILENAME in $PSSCRIPTROOT\config, exiting." -foregroundcolor red
    read-host "couldn't fgind config"
    exit
}
$config_file = $config_file | ConvertFrom-Json

## $env:PSMENU_DIR --> Base directory of terminal menu.
Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Setting " -nonewline
Write-Host "`$env:PSMENU_DIR" -ForegroundColor Green -NoNewline
Write-Host " environment variable to $((Get-Item .).FullName)."
$env:PSMENU_DIR = (Get-Item $PSSCRIPTROOT).FullName

## $env:LOCAL_SCRIPTS
Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Setting " -NoNewline
Write-Host "`$env:LOCAL_SCRIPTS" -foregroundcolor green -NoNewline
Write-Host " environment variable to $((Get-Item $PSSCRIPTROOT\scripts).FullName)."
$env:LOCAL_SCRIPTS = (Get-Item $PSSCRIPTROOT\scripts).FullName

## Set reports directory environment variable
$env:REPORTS_DIRECTORY = "$env:USERPROFILE\PSTechSupportMenu\reports"

## Create user-specific inventory/canned_responses directories
$env:INVENTORY_DIRECTORY = "$env:USERPROFILE\PSTechSupportMenu\inventory"

$env:CANNED_RESPONSES_DIRECTORY = "$env:USERPROFILE\PSTechSupportMenu\canned_responses"

@("$env:REPORTS_DIRECTORY", "$env:INVENTORY_DIRECTORY", "$env:CANNED_RESPONSES_DIRECTORY") | % {
    if (-not (Test-Path $_)) {
        New-Item -Path $_ -Force -ItemType Directory | Out-Null
    }
}

$functions = @{}
# Keys = Category Names
# Values = List of functions for that category
foreach ($category in $config_file.categories.PSObject.Properties) {
    $functions[$category.Name] = $category.Value
}

## Import modules / functions
get-childitem modules | % { ipmo "$($_.fullname)" -WarningAction SilentlyContinue }

# this line is here just so it will stop if there are errors when trying to install/import modules
Write-Host "`nDebugging point in case errors are encountered - please screenshot and share if you're able." -Foregroundcolor Yellow
Read-Host "Thank you! Press enter to continue."
# create variable, when true - exits program loop
$exit_program = $false
while ($exit_program -eq $false) {
    $target_computers = $null

    Clear-Host
    # splitting function options into a list - it was string divided by spaces before
    $split_options = [string[]]$config_file.categories.PSObject.Properties.Name
    $options = [system.collections.arraylist]::new()
    Foreach ($option in $split_options) {
        $newoption = $option -replace '_', ' '
        $options.add($newoption) | out-null
    }

    ## Add Help to options list
    $options.add('Help') | Out-Null

    ## Add Search to options list
    $options.add('Search') | Out-Null

    ## Allow the user to choose category - CORRESPONDS to categories listed in config.json
    $chosen_category = Menu $($options)
    
    ## SEARCH functionality - returns any file/functions in the functions directory with name containing the input search term.
    if ($chosen_category -eq 'search') {
        Write-Host "Search will return any functions that contain the search term." -Foregroundcolor Yellow
        $search_term = Read-Host "Enter search term"
        # $function_list = Get-ChildItem -Path "$env:PSMENU_DIR\functions" -Filter "*$search_term*.ps1" -File -ErrorAction SilentlyContinue | Select -Exp BaseName

        $function_list = Get-ChildItem ./Modules | % { gcm -module $_ } | Select -Exp Name | ? { $_ -like "*$search_term*" }

        ## if there are no results - allow user to continue back to original menu
        if (-not $function_list) {
            Write-Host "No results found for search term: " -NoNewLine
            Write-Host "$search_term" -ForegroundColor Red
            Read-Host "Press enter to return to original menu."
            continue
        }
    }
    ## HELP Functionality - open README on Github
    elseif ($chosen_category -eq 'Help') {
        Open-Guide
        continue
    } ## If the user chooses to exit, the program will exit
    else {
        #reassemble the chosen category with _, if that method was chosen, if it wasn't - nothing will happen
        # $chosen_category = $chosen_category -replace ' ', '_'
        # get the functions from the chosen category so they can be presented in the second menu
        $function_list = $functions["$chosen_category"]
    }
    Clear-Host
    $function_list = $function_list -replace '-', ' '
    $function_list = $function_list | sort
    if ($function_list.getType().name -eq 'String') {
        $function_list = @($function_list)
    }
    # put 'Return-ToPreviousMenu' at bottom of list - allows user to return to category choices
    $function_list += 'Return to previous menu'

    ## FUNCTION SELECTION - menu is presented using psmenu module
    $function_selection = Menu $function_list

    Clear-Host
    # reconstruct the actual filename
    # If return was chosen, continue to next iteration of infinite loop - continues until user chooses to exit
    if ($function_selection -eq 'Return to previous menu') {
        continue
    }
    $function_selection = $function_selection -replace ' ', '-'

    # 'Get' info on the command, if there is 1+ parameter(s) - cycle through them prompting the user for values.
    $command = Get-Command $function_selection
    if ($command.Parameters.Count -gt 0) {
        # Getting detailed info on the command is what allows printing of parameter descriptions to terminal, above where user is being prompted for their values.
        $functionhelper = get-help $command -Detailed

        $functions_synopsis = $functionhelper.synopsis.text
        $function_description = $functionhelper.description.text

        Write-Host "$function_selection -> " -foregroundcolor Green
        Write-Host "Function description: " -nonewline -foregroundcolor yellow
        Write-Host "$function_description"
        Write-Host "`n$functions_synopsis"
        Write-Host ""

        # Parameter names
        $parameters = $command.Parameters.Keys
        # Hashtable: Keys = Parameter names, Values = values input by user
        $splat = @{}

        foreach ($parameter in $parameters) {
            ## SKIPS COMMON PARAMETERS: Verbose, Debug, ErrorAction, WarningAction, InformationAction, ErrorVariable, WarningVariable, InformationVariable, OutVariable, OutBuffer, PipelineVariable
            # if ($parameter -in ('Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'ProgressAction')) {
            #     continue
            # }

            $current_parameter_info = $functionhelper.parameters.parameter | Where-Object { $_.name -eq $parameter }
            # For each line in that text block (underneath .PARAMETER parameterName)
            Write-Host "`nParameter $parameter DESCRIPTION: `n" -NoNewLine -Foregroundcolor Yellow
            ForEach ($textitem in $current_parameter_info.description) {
                # Write each line to terminal.
                $textitem.text
            }
            Write-Host "`n"


            ## This is the one thing specifically added in for a function - Install-Application
            ## Will display a list of applications available when user is prompted to enter name/names.
            if (($function_selection -eq 'Install-Application') -and ($parameter -eq 'AppName')) {
                Write-Host "Available applications: " -Foregroundcolor Green
                    (Get-ChildItem -Path "$env:PSMENU_DIR\deploy\applications" -Directory -ErrorAction SilentlyContinue).Name
                Write-Host "`n"
            }

            # Read-HostNoColon is just Read-Host without the colon at the end, so that an = can be used.
            $value = Read-HostNoColon -Prompt "$parameter = "
            if ($value) {
                $splat.Add($parameter, $value)
            }
        }

        $splat_String = $splat.getenumerator() | % { "-$($_.name) $($_.value) " }

        Invoke-Expression -command "$command $splat_string" | Out-Null
    }
    else {
        # execute the command without parameters if it doesn't have any.
        Invoke-Expression -Command $command | Out-Null
    }

    ## USER can press x to exit, or enter to return to main menu (category selection)
    Write-Host "`nPress " -NoNewLine
    Write-Host "'x'" -ForegroundColor Red -NoNewline
    Write-Host " to exit, or " -NoNewline 
    Write-Host "[ENTER] to return to menu." -ForegroundColor Yellow -NoNewline

    $key = $Host.UI.RawUI.ReadKey()
    [String]$character = $key.Character
    if ($($character.ToLower()) -eq 'x') {
        return
    }
    elseif ($($character.ToLower()) -eq '') {
        continue
    }

}

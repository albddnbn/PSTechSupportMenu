1. Teamsbootstrapper.exe needs to be downloaded in the script - Install-NewTeams
2. Create inventory and canned response user-specific folder paths in menu.ps1 - environment variables?
3. Make sure modules in ./config/modules.csv that have menu=y are added to config.json in 'other' category if not already present in one of the other categories.
4. Add the following to the config.json file:
    - ensure they're installed / imported correctly, ideally with no continuing errors
4. Test job functionality
5. edit inventory and canned response functions to include new environment variables.

1. Teamsbootstrapper.exe needs to be downloaded in the script - Install-NewTeams
2. Create inventory and canned response user-specific folder paths in menu.ps1 - environment variables?
3. Remove any traces of modules.csv - it shouldn't be too bad to just put modules into modules directory.
4. Add the following to the config.json file:
    - ensure they're installed / imported correctly, ideally with no continuing errors
4. Test job functionality
5. edit inventory and canned response functions to include new environment variables.

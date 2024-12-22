1. Teamsbootstrapper.exe needs to be downloaded in the script - Install-NewTeams
2. edit inventory and canned response functions to include new environment variables.
3. Any errors that occur during invoke-command on report-generating functions will interfere with parsing error machine hostnames from winrm error messages as things are now.
4. Is copy-snippettoclipboard pointless? Can just create canned response with no variables.
5. All of installs/updates still untested.
6. you will have to use the menu to be able to create/copy canned responses.
7. Installs and Updates are going to be an afterthought - most Tech Support professionals will be able to deploy/update applications more efficiently.

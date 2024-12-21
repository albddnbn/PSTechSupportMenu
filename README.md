# PSTechSupportMenu
 Powershell module that presents menu with functions, prompts user for parameter values of chosen function.

## Currently 'under construction' with goal of simplifying functions and creating a module (if there's a point) from:

<a href="https://github.com/albddnbn/PSTerminalMenu">https://github.com/albddnbn/PSTerminalMenu</a>

Notes:
1. 12.20.24 - Report-generating/scan functions will be a module or submodule of their own. Unsure about whether the actual menu script will be a module. Other major groups of functions will also be modules.
2. The script will check for Modules folder in current directory and import them all. It will also check a 'modules list' to import all listed modules to give access to commands. Note that for advanced functions you will still have to deal with extra parameters since they're no longer expllcitly skipped over. This should make the menu more versatile than ever!
   - Hopefully there will be a way to quickly update the config.json as well.

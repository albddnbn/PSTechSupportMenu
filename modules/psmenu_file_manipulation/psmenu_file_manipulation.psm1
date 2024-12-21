Function Convert-ImgToBitmap {
    <#
    .SYNOPSIS
        Uses the Get-Image and Convert-ToBitmap functions from: https://github.com/dwj7738/PowershellModules/tree/master/PSImageTools to convert specified image to bitmap.
    
    .DESCRIPTION
        Get-Image: https://github.com/dwj7738/PowershellModules/blob/master/PSImageTools/Get-Image.ps1
        Convert-ToBitmap: https://github.com/dwj7738/PowershellModules/blob/master/PSImageTools/ConvertTo-Bitmap.ps1
    
    .EXAMPLE
        Convert-ImgToBitmap -ImagePath 'C:\Users\user\Downloads\image.jpg'
    
    .NOTES
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>

    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )
    function Get-Image {
        <#
        .Synopsis
            Gets information about images.
    
        .Description
            Get-Image gets an object that represents each image file. 
            The object has many properties and methods that you can use to edit images in Windows PowerShell. 
    
        .Notes
            Get-Image uses Wia.ImageFile, a Windows Image Acquisition COM object to get image data.
            Then it uses the Add-Member cmdlet to add note properties and script methods to the object. 
    
            The Resize script method uses the Add-ScaleFilter function. It has the following syntax:
            Resize ($width, $height, [switch]$DoNotPreserveAspectRation). 
            Width and Height can be specified in pixels or percentages. 
            For a description of these parameters, type "get-help Add-ScaleFilter –par *".
    
            The Crop script method uses the uses the Add-CropFilter function. It has the following syntax:
            Crop ([Double]$left, [Double]$top, [Double]$right, [Double]$bottom).
            All dimensions are measured in pixels.
            For a description of these parameters, type "get-help Add-CropFilter –par *".
    
            The FlipVertical, FlipHorizontal, RotateClockwise and RotateCounterClockwise script methods use the Add-RotateFlipFilter function.
            For a description of these parameters, type "get-help Add-RotateFlipFilter –par *".
    
        .Parameter File
            [Required] Specifies the image files. Enter the path and file name of an image file, such as $home\pictures\MyPhoto.jpg.
            You can also pipe one or more image files to Get-Image, such as those from Get-Item or Get-Childitem (dir). 
    
        .Example
            Get-Image –file C:\myPics\MyPhoto.jpg
    
        .Example
            Get-ChildItem $home\Pictures -Recurse | Get-Image        
    
        .Example
            (Get-Image –file C:\myPics\MyPhoto.jpg).resize(80, 120)
    
        .Example
            # Crops 8 pixels from the top of the image.
            $CatPhoto = Get-Image –file $home\Pictures\Cat.jpg
            $CatPhoto.crop(0,8,0,0)
    
        .Example
            $CatPhoto = Get-Image –file $home\Pictures\Cat.jpg
            $CatPhoto.flipvertical()
    
        .Example
            dir $home\pictures\Vacation*.jpg | get-image | format-table fullname, horizontalResolution, PixelDepth –autosize
    
        .Link
            "Image Manipulation in PowerShell": http://blogs.msdn.com/powershell/archive/2009/03/31/image-manipulation-in-powershell.aspx
        .Link
            Add-CropFilter
        .Link
            Add-ScaleFilter
        .Link
            Add-RotateFlipFilter
        .Link
            Get-ImageProperties
        #>
        param(    
            [Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
            [Alias('FullName')]
            [string]$file)
        
        process {
            $realItem = Get-Item $file -ErrorAction SilentlyContinue     
            if (-not $realItem) { return }
            $image = New-Object -ComObject Wia.ImageFile        
            try {        
                $image.LoadFile($realItem.FullName)
                $image | 
                Add-Member NoteProperty FullName $realItem.FullName -PassThru | 
                Add-Member ScriptMethod Resize {
                    param($width, $height, [switch]$DoNotPreserveAspectRatio)                    
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-ScaleFilter @psBoundParameters -passThru -image $image
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru | 
                Add-Member ScriptMethod Crop {
                    param([Double]$left, [Double]$top, [Double]$right, [Double]$bottom)
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-CropFilter @psBoundParameters -passThru -image $image
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru | 
                Add-Member ScriptMethod FlipVertical {
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-RotateFlipFilter -flipVertical -passThru 
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru | 
                Add-Member ScriptMethod FlipHorizontal {
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-RotateFlipFilter -flipHorizontal -passThru 
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru |
                Add-Member ScriptMethod RotateClockwise {
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-RotateFlipFilter -angle 90 -passThru 
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru |
                Add-Member ScriptMethod RotateCounterClockwise {
                    $image = New-Object -ComObject Wia.ImageFile
                    $image.LoadFile($this.FullName)
                    $filter = Add-RotateFlipFilter -angle 270 -passThru 
                    $image = $image | Set-ImageFilter -filter $filter -passThru
                    Remove-Item $this.Fullname
                    $image.SaveFile($this.FullName)                    
                } -PassThru 
                    
            }
            catch {
                Write-Verbose $_
            }
        }    
    }
    function ConvertTo-Bitmap {
        <#
        .Synopsis
        Converts an image to a bitmap (.bmp) file.
    
        .Description
        The ConvertTo-Bitmap function converts image files to .bmp file format.
        You can specify the desired image quality on a scale of 1 to 100.
    
        ConvertTo-Bitmap takes only COM-based image objects of the type that Get-Image returns.
        To use this function, use the Get-Image function to create image objects for the image files, 
        then submit the image objects to ConvertTo-Bitmap.
    
        The converted files have the same name and location as the original files but with a .bmp file name extension. 
        If a file with the same name already exists in the location, ConvertTo-Bitmap declares an error. 
    
        .Parameter Image
        Specifies the image objects to convert.
        The objects must be of the type that the Get-Image function returns.
        Enter a variable that contains the image objects or a command that gets the image objects, such as a Get-Image command.
        This parameter is optional, but if you do not include it, ConvertTo-Bitmap has no effect.
    
        .Parameter Quality
        A number from 1 to 100 that indicates the desired quality of the .bmp file.
        The default is 100, which represents the best possible quality.
    
        .Parameter HideProgress
        Hides the progress bar.
    
        .Parameter Remove
        Deletes the original file. By default, both the original file and new .bmp file are saved. 
    
        .Notes
        ConvertTo-Bitmap uses the Windows Image Acquisition (WIA) Layer to convert files.
    
        .Link
        "Image Manipulation in PowerShell": http://blogs.msdn.com/powershell/archive/2009/03/31/image-manipulation-in-powershell.aspx
    
        .Link
        "ImageProcess object": http://msdn.microsoft.com/en-us/library/ms630507(VS.85).aspx
    
        .Link 
        Get-Image
    
        .Link
        ConvertTo-JPEG
    
        .Example
        Get-Image .\MyPhoto.png | ConvertTo-Bitmap
    
        .Example
        # Deletes the original BMP files.
        dir .\*.jpg | get-image | ConvertTo-Bitmap –quality 100 –remove -hideProgress
    
        .Example
        $photos = dir $home\Pictures\Vacation\* -recurse –include *.jpg, *.png, *.gif
        $photos | get-image | ConvertTo-Bitmap
        #>
        param(
            [Parameter(ValueFromPipeline = $true)]    
            $Image,
        
            [ValidateRange(1, 100)]
            [int]$Quality = 100,
        
            [switch]$HideProgress,
        
            [switch]$Remove
        )
        process {
            if (-not $image.Loadfile -and 
                -not $image.Fullname) { return }
            $realItem = Get-Item -ErrorAction SilentlyContinue $image.FullName
            if (-not $realItem) { return }
            if (".bmp" -ne $realItem.Extension) {
                $noExtension = $realItem.FullName.Substring(0, 
                    $realItem.FullName.Length - $realItem.Extension.Length)
                $process = New-Object -ComObject Wia.ImageProcess
                $convertFilter = $process.FilterInfos.Item("Convert").FilterId
                $process.Filters.Add($convertFilter)
                $process.Filters.Item(1).Properties.Item("Quality") = $quality
                $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
                $newImg = $process.Apply($image.PSObject.BaseObject)
                $newImg.SaveFile("$noExtension.bmp")
                if (-not $hideProgress) {
                    Write-Progress "Converting Image" $realItem.Fullname
                }
                if ($remove) {
                    $realItem | Remove-Item
                }
            }
        
        }
    }

    if (Test-Path $ImagePath -ErrorAction SilentlyContinue) {

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Successfully tested path at $ImagePath, attempting to convert to bitmap."

        Get-Image $ImagePath | ConvertTo-Bitmap

    }
    else {
            
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Path at $ImagePath does not exist, please check the path and try again." -ForegroundColor Red
    }

}

function Convert-PNGtoICO {
    <#
    .SYNOPSIS
        Converts image to icons, verified with .png files.

    .DESCRIPTION
        Easier than trying to find/navigate reputable online/executable converters.

    .Example
        Convert-PNGtoICO -File .\Logo.png -OutputFile .\Favicon.ico

    .NOTES
        SOURCE: https://www.powershellgallery.com/packages/RoughDraft/0.1/Content/ConvertTo-Icon.ps1
        # wrapped for Terminal menu: Alex B. (albddnbn)
    #>
    
    param(
        [string]$File,
        # If set, will output bytes instead of creating a file
        # [switch]$InMemory,
        # If provided, will output the icon to a location
        # [Parameter(Position = 1, ValueFromPipelineByPropertyName = $true)]
        [string]$OutputFile
    )
    
    begin {
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        
    }
    
    process {
        #region Load Icon
        $resolvedFile = $ExecutionContext.SessionState.Path.GetResolvedPSPathFromPSPath($file)
        if (-not $resolvedFile) { return }
        $loadedImage = [Drawing.Image]::FromFile($resolvedFile)
        $intPtr = New-Object IntPtr
        $thumbnail = $loadedImage.GetThumbnailImage(72, 72, $null, $intPtr)
        $bitmap = New-Object Drawing.Bitmap $thumbnail 
        $bitmap.SetResolution(72, 72); 
        $icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon());         
        #endregion Load Icon

        #region Save Icon
        if ($InMemory) {                        
            $memStream = New-Object IO.MemoryStream
            $icon.Save($memStream) 
            $memStream.Seek(0, 0)
            $bytes = New-Object Byte[] $memStream.Length
            $memStream.Read($bytes, 0, $memStream.Length)                        
            $bytes
        }
        elseif ($OutputFile) {
            $resolvedOutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputFile)
            $fileStream = [IO.File]::Create("$resolvedOutputFile")                               
            $icon.Save($fileStream) 
            $fileStream.Close()               
        }
        #endregion Save Icon

        #region Cleanup
        $icon.Dispose()
        $bitmap.Dispose()
        #endregion Cleanup

    }
}

function New-DriveMappingExe {
    <#
    .SYNOPSIS
        Uses the PS2exe module to generate an executable, that when double-clicked by a user, will map the specified drive / drive letter.
        The user must have correct permissions assigned to access specified drive path.

    .DESCRIPTION
        Generates an executable that, when double-clicked, will map a specified drive for user.
        Add a Drivemap.ico file to ./SupportFiles to add custom icon to executable (right now, it's just a hard drive icon).

    .PARAMETER TargetPath
        UNC path of folder to be mapped, ex: '\\server-01\users\lvoldemort'

    .PARAMETER DriveLetter
        Drive letter to map the path to. Ex: 'Z'

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    
    param(
        [Parameter(Mandatory = $true)]        
        [string]$TargetPath,
        [Parameter(Mandatory = $true)]
        [string]$DriveLetter
    )
    ## Script that's compiled into an executable
    $drive_map_scriptblock = @"
Write-Host "Attempting to map  " -NoNewLine
Write-Host "$TargetPath" -Foregroundcolor Yellow -NoNewLine
Write-Host " to " -NoNewLine
Write-Host "$DriveLetter" -NoNewLine -Foregroundcolor Green
Write-Host "..."
# creates the persistent drive mapping
try {
    New-PSDrive -Name $driveletter -PSProvider FileSystem -Root `'$TargetPath`' -Persist
    Write-Host "Successfully mapped " -NoNewLine
    Write-Host "$TargetPath" -Foregroundcolor Yellow -NoNewLine
    Write-Host " to " -NoNewLine
    Write-Host "$DriveLetter" -NoNewLine -Foregroundcolor Green
    Write-Host "."
    Start-Sleep -Seconds 5
} catch {
    Write-Host "Failed to map $TargetPath to $DriveLetter." -Foregroundcolor Red
    Start-Sleep -Seconds 5
}
"@

    # create output filename: get last string from targetpath
    $splitup_targetpath = $TargetPath -split '\\'
    $last_string = $splitup_targetpath[-1]
    $outputfile = "drivemap-$last_string-$driveletter.ps1"
    # create executable output path
    $exe_outputfile = $outputfile -replace '.ps1', '.exe'

    # define absolute paths for .ps1 script output, and generated .exe output
    $ps1_output_path = "$env:PSMENU_DIR\output\$thedate\$EXECUTABLES_DIRECTORY\$outputfile"
    $exe_output_path = "$env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY\$exe_outputfile"
    foreach ($singledir in @("$env:PSMENU_DIR\output\$thedate\$EXECUTABLES_DIRECTORY", "$env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY")) {
        if (-not (Test-Path $singledir -ErrorAction SilentlyContinue)) {
            New-Item -Path $singledir -ItemType 'Directory' -Force | Out-Null
        }
    }
    
    # $drive_map_scriptblock | Out-File "$ps1_output_path" -Force
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Creating $ps1_output_path..."
    Set-Content -Path "$ps1_output_path" -Value $drive_map_scriptblock -Force


    # supportfiles\drivemap.ico just adds a hard drive icon to the executable that's created
    $drivemap_icon_file = Get-ChildItem -Path "$env:SUPPORTFILES_DIR" -Filter "drivemap.ico" -File -ErrorAction SilentlyContinue

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Compiling $ps1_output_path into $exe_output_path..."
    Start-Sleep -Seconds 1
    if ($drivemap_icon_file) {
        Invoke-PS2exe -inputFile "$ps1_output_path" -outputFile "$exe_output_path" -iconFile "$($drivemap_icon_file.fullname)" -title "Map $($DriveLetter): Drive" -company "Delaware Technical Community College"
    }
    else {
        Invoke-PS2exe -inputFile "$ps1_output_path" -outputFile "$exe_output_path" -title "Map $($DriveLetter): Drive" -company "Delaware Technical Community College"
    }
    try {
        Invoke-Item $($exe_output_path | split-path -parent)
    }
    catch {
        Write-Host "Failed to open directory for $exe_output_path." -Foregroundcolor Yellow
    }

    Read-Host "Press Enter to continue."
}

function New-IntuneWin32App {
    <#
    .SYNOPSIS
        Creates an .intunewin package file using a PSADT Deployment Folder, given certain conditions.

    .DESCRIPTION
        Uses IntuneWinAppUtil.exe from: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool

    .PARAMETER ScriptFolder
        Path to folder holding source files and/or script.

    .PARAMETER InstallationFile
        Path to the installation executable or script. If your package is installed using a script - target the script. If it's only using an .msi, or .exe - target this.
        This is the file that will be targeted in your 'Installation Line' when creating the Win32 App in Intune.

    .PARAMETER OutputFolder
        Path to the folder where the .intunewin file will be created.

    .EXAMPLE
        Generate-IntuneWin32AppFromPSADT -ScriptFolder "C:\Users\williamwonka\Desktop\PSADT-Deployments\AdobeAcrobatReaderDC-2021.001.20155" -InstallationFile "C:\Users\abuddenb\Desktop\PSADT-Deployments\AdobeAcrobatReaderDC-2021.001.20155\Deploy-Application.ps1" -OutputFolder "C:\Users\abuddenb\Desktop\PSADT-Deployments\AdobeAcrobatReaderDC-2021.001.20155"

    .NOTES
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scriptfolder,
        [Parameter(Mandatory = $true)]
        [string]$InstallationFile,
        [Parameter(Mandatory = $true)]
        [string]$OutputFolder
    )

    # get intunewinapputil.exe
    $Intunewinapputil_exe = Get-Childitem -Path "$env:SUPPORTFILES_DIR" -Filter 'IntuneWinAppUtil.exe' -File -Recurse -Erroraction SilentlyContinue
    if (-not $Intunewinapputil_exe) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Unable to find IntuneWinAppUtil.exe in $env:PSMENU_DIR, attempting to download from github..." -ForegroundColor Red
        # download the png from : https://iit.dtcc.edu/wp-content/uploads/2023/07/it-logo.png
        try {
            Invoke-WebRequest -Uri "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/archive/refs/tags/v1.8.5.zip" -OutFile "$env:PSMENU_DIR\IntuneWinAppUtil.zip"
            Expand-Archive -Path "$env:PSMENU_DIR\IntuneWinAppUtil.zip" -DestinationPath "$env:PSMENU_DIR"
            $Intunewinapputil_exe = Get-Childitem -Path "$env:PSMENU_DIR" -Filter 'IntuneWinAppUtil.exe' -File -Recurse -Erroraction SilentlyContinue
            if (-not $Intunewinapputil_exe) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Unable to find IntuneWinAppUtil.exe in $env:PSMENU_DIR, exiting." -ForegroundColor Red
                exit
            }
        }
        catch {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Unable to find IntuneWinAppUtil.exe from github." -ForegroundColor Red
            Exit
        }
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Found $($Intunewinapputil_exe.fullname) in $env:PSMENU_DIR."
    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Creating .intunewin package from " -NoNewLine
    Write-Host "$ScriptFolder" -Foregroundcolor Green -NoNewline
    Write-Host ", using " -NoNewLine
    Write-Host "$Installationfile" -NoNewLine -Foregroundcolor Yellow
    Write-Host ", and saving to " -NoNewLine
    Write-Host "$OutputFolder" -Foregroundcolor Cyan -NoNewline
    Write-Host "."

    Start-Process -FilePath "$($Intunewinapputil_exe.fullname)" -ArgumentList "-c ""$ScriptFolder"" -s ""$Installationfile"" -o $OutputFolder -q" -Wait

    Start-Sleep -Seconds 2
    Invoke-Item $OutputFolder

    Write-Host "If the package follows traditional PSADT format, the execution line should be:"
    Write-Host "powershell.exe -ExecutionPolicy Bypass .\$($InstallationFile | Split-Path -Leaf) -DeploymentType 'Install' -DeployMode 'Silent'" -Foregroundcolor Yellow
    # Read-Host "This is an experimental function, press ENTER to continue after screenshotting any error messages, please!"
}

function New-PrinterConnectionExe {
    <#
    .SYNOPSIS
        Generates executable using the PS2exe powershell module that will map a printer.
        Either PrinterLogic or on print server, specified by $PrinterType parameter.

    .DESCRIPTION
        PrinterLogic  - Printer Installer Client software must be installed on machines where executable is being used.
        Print servers - Print server must be accessible over the network, on machines where executable will be used.

    .PARAMETER Printername
        Name of the printer to map. Ex: 'printer-c136-01'
        PrinterLogic - needs to match printer's name as listed in PrinterLogic Printercloud instance.
        Print server - needs to match printer's hostname as listed in print server and on DNS server.

    .PARAMETER PrinterType
        Set to 'printerlogic' or the name of a printer server, ex: 's-ps-02', 'org-printserv-001'

    .EXAMPLE
        Generate-PrinterLogicExe -PrinterName "printer-c136-01"

    .NOTES
        Executable will be created in the 'executables' directory.
        ---
        Author: albddnbn (Alex B.)
        Project Site: https://github.com/albddnbn/PSTerminalMenu
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$PrinterName,
        # [ValidateSet('printerlogic', 'ps')]
        $PrinterType
    )
    $EXECUTABLES_DIRECTORY = 'PrinterMapping'
    $thedate = Get-Date -Format 'yyyy-MM-dd'

    ### Make sure the executables and output directories for today exist / create if not.
    foreach ($singledir in @("$env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY", "$env:PSMENU_DIR\output\$thedate")) {
        if (-not (Test-Path $singledir -ErrorAction SilentlyContinue)) {
            New-Item -Path $singledir -ItemType 'Directory' -Force | Out-Null
        }
    }

    ## creates $exe_script variable depending on what kind of printer needs to be mapped
    if ($PrinterType -notlike "*print*logic*") {
        $exe_script = @"
`$printername = '\\$PrinterType\$PrinterName'
try {
    (New-Object -comobject wscript.network).addwindowsprinterconnection(`$printername)
    (New-Object -comobject wscript.network).setdefaultprinter(`$printername)
    Write-Host "Mapped `$printername successfully." -Foregroundcolor Green
} catch {
    Write-Host "Failed to map printer: `$printername, please let Tech Support know." -Foregroundcolor Red
}
Start-Sleep -Seconds 5
"@

        # create output filename for print server mapping:
        $output_filename = "$PrinterType-connect-$PrinterName"
    }
    else {
        # generate text for .ps1 file
        $exe_script = @"
# 'get' the .exe
`$printerinstallerconsoleexe = Get-Childitem -Path "C:\Program Files (x86)\Printer Properties Pro\Printer Installer Client\bin\" -Filter "PrinterInstallerconsole.exe" -File -ErrorAction silentlycontinue
# run install command:
`$execution_result = Start-Process "`$(`$printerinstallerconsoleexe.FullName)" -ArgumentList "InstallPrinter=$printername"
# if (`$execution_result -eq 0) {
#     Write-Host "[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Successfully installed printer: `$printername" -Foregroundcolor Green
# } else {
#     Write-Host "[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Failed to install printer: `$printername, please let tech support know." -Foregroundcolor Red
# }
Start-Sleep -Seconds 5
"@

        $output_filename = "plogic-connect-$PrinterName"

    }

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Creating executable that will map $PrinterType printer: $printername when double-clicked by user..."

    $exe_script | Out-File -FilePath "$env:PSMENU_DIR\output\$thedate\$output_filename.ps1" -Force

    Invoke-PS2EXE -inputfile "$env:PSMENU_DIR\output\$thedate\$output_filename.ps1" `
        -outputfile "$env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY\$output_filename.exe"

    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] :: Executable created successfully: $env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY\$output_filename.exe" -ForegroundColor Green

    Invoke-Item "$env:PSMENU_DIR\executables\$thedate\$EXECUTABLES_DIRECTORY"
}

Export-ModuleMember -Function *-*

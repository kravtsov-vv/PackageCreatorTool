<#

.SYNOPSIS
    Create_FSv2.ps1 - Package Creator Tool

.DESCRIPTION
    A comprehensive GUI-based tool for preparing and creating standardized software package structures and deployment scripts for enterprise software distribution via SCCM/MECM and Microsoft Intune.

    The script provides a user-friendly Windows Forms interface to:
    - Input package metadata (manufacturer, product name, version, architecture, language, revision, zone code)
    - Browse and select installer files (MSI or EXE)
    - Automatically extract MSI metadata (product code, version, manufacturer, etc.) when an MSI is selected
    - Configure detection methods for MSI (via product code) or EXE (via file path, name, and version)
    - Specify display name from Add/Remove Programs

    Upon package creation, the tool:
    - Generates a standardized folder structure: RootPath\Vendor\Product\Version_Arch_Lang_Rev_Zone\
    - Creates subfolders: Documentation\, Resources\, Package\
    - Copies PSADT (PowerShell App Deployment Toolkit) templates to Package\ folder
    - Updates Deploy-Application.ps1 with package-specific variables and installer commands
    - Generates test installation/uninstallation scripts in Resources\ using PsExec
    - Creates SCCM deployment script (Deploy_SCCM.ps1) with full application configuration
    - Generates Intune conversion script (Convert_Intunewin.ps1) for packaging as .intunewin
    - Copies installer files (and optionally entire source folder) to Package\Files\
    - Creates documentation template in Documentation\

    Supports zone-specific packages for multi-environment deployments and includes validation for required fields.

.PARAMETER RootPath
    Specifies the target root folder where package folders will be created. Defaults to "T:\Software" if not provided.

.NOTES
    Version:        4.3
    Author:         Viktor Kravtsov
    Date:           10/05/2026
    Changes:        Updated header documentation and minor improvements to the GUI tool for package creation.

.EXAMPLE
    .\Create_FSv2.ps1 'T:\Software'
    Launches the GUI tool with T:\Software as the default root path for package creation.

#>

Param (
    [Parameter(Mandatory = $false)] [string]$global:RootPath = "T:\Software"
    )
	
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

#*************** FUNCTIONS REGION  *************************************************************************
Function Get-MSIFileInformation {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo[]]$FilePath
    ) 
  
    # https://learn.microsoft.com/en-us/windows/win32/msi/installer-opendatabase
    $msiOpenDatabaseModeReadOnly = 0
    
    $productLanguageHashTable = @{
        '1025' = 'Arabic'
        '1026' = 'Bulgarian'
        '1027' = 'Catalan'
        '1028' = 'Chinese - Traditional'
        '1029' = 'Czech'
        '1030' = 'Danish'
        '1031' = 'German'
        '1032' = 'Greek'
        '1033' = 'English'
        '1034' = 'Spanish'
        '1035' = 'Finnish'
        '1036' = 'French'
        '1037' = 'Hebrew'
        '1038' = 'Hungarian'
        '1040' = 'Italian'
        '1041' = 'Japanese'
        '1042' = 'Korean'
        '1043' = 'Dutch'
        '1044' = 'Norwegian'
        '1045' = 'Polish'
        '1046' = 'Brazilian'
        '1048' = 'Romanian'
        '1049' = 'Russian'
        '1050' = 'Croatian'
        '1051' = 'Slovak'
        '1053' = 'Swedish'
        '1054' = 'Thai'
        '1055' = 'Turkish'
        '1058' = 'Ukrainian'
        '1060' = 'Slovenian'
        '1061' = 'Estonian'
        '1062' = 'Latvian'
        '1063' = 'Lithuanian'
        '1081' = 'Hindi'
        '1087' = 'Kazakh'
        '2052' = 'Chinese - Simplified'
        '2070' = 'Portuguese'
        '2074' = 'Serbian'
    }

    $summaryInfoHashTable = @{
        1  = 'Codepage'
        2  = 'Title'
        3  = 'Subject'
        4  = 'Author'
        5  = 'Keywords'
        6  = 'Comment'
        7  = 'Template'
        8  = 'LastAuthor'
        9  = 'RevisionNumber'
        10 = 'EditTime'
        11 = 'LastPrinted'
        12 = 'CreationDate'
        13 = 'LastSaved'
        14 = 'PageCount'
        15 = 'WordCount'
        16 = 'CharacterCount'
        18 = 'ApplicationName'
        19 = 'Security'
    }

    $properties = @('ProductVersion', 'ProductCode', 'ProductName', 'Manufacturer', 'ProductLanguage', 'UpgradeCode')
   
    try {
        $file = Get-ChildItem $FilePath -ErrorAction Stop
    }
    catch {
        Write-Warning "Unable to get file $FilePath $($_.Exception.Message)"
        return
    }

    $object = [PSCustomObject][ordered]@{
        FileName     = $file.Name
        FilePath     = $file.FullName
        'Length(MB)' = $file.Length / 1MB
    }

    # Read property from MSI database
    $windowsInstallerObject = New-Object -ComObject WindowsInstaller.Installer

    # open read only    
    $msiDatabase = $windowsInstallerObject.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstallerObject, @($file.FullName, $msiOpenDatabaseModeReadOnly))

    foreach ($property in $properties) {
        $view = $null
        $query = "SELECT Value FROM Property WHERE Property = '$($property)'"
        $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
        $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
        $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)

        try {
            $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
        }
        catch {
            Write-Verbose "Unable to get '$property' $($_.Exception.Message)"
            $value = ''
        }
        
        if ($property -eq 'ProductLanguage') {
            $value = "$value ($($productLanguageHashTable[$value]))"
        }

        $object | Add-Member -MemberType NoteProperty -Name $property -Value $value
    }

    $summaryInfo = $msiDatabase.GetType().InvokeMember('SummaryInformation', 'GetProperty', $null, $msiDatabase, $null)
    $summaryInfoPropertiesCount = $summaryInfo.GetType().InvokeMember('PropertyCount', 'GetProperty', $null, $summaryInfo, $null)

    (1..$summaryInfoPropertiesCount) | ForEach-Object {
        $value = $SummaryInfo.GetType().InvokeMember("Property", "GetProperty", $Null, $SummaryInfo, $_)

        if ($null -eq $value) {
            $object | Add-Member -MemberType NoteProperty -Name $summaryInfoHashTable[$_] -Value ''
        }
        else {
            $object | Add-Member -MemberType NoteProperty -Name $summaryInfoHashTable[$_] -Value $value
        }
    }

    #$msiDatabase.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $msiDatabase, $null)
    $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
 
    # Run garbage collection and release ComObject
    $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstallerObject) 
    [System.GC]::Collect()

    return $object  
}

Function CreateFolders {
    param
    (
        [string]$RootPath
    )

    $Application = $_AppObject
    $AppVendor = ($Application.AppVendor -replace '\s','').Trim()
    $AppName = ($Application.AppName -replace '\s','').Trim()
    $AppVersion = ($Application.AppVersion + "_" + $_AppObject.AppArch + "_" + $_AppObject.AppLang + "_" + $_AppObject.AppRev + $_AppObject.AppZone).Trim()   

    $PackagePath= $RootPath+'\'+$AppVendor+'\'+$AppName+'\'+$AppVersion
    If (Test-Path -Path $PackagePath){
		[System.Windows.MessageBox]::Show("The folder already exist:`n`n $($PackagePath) `n`nPress 'OK' to back","",0,16)
		MAIN
		}  

    New-Item ($RootPath,$AppVendor,$AppName,$AppVersion, 'Documentation' -join "\") -Type Directory -Force | Out-Null
    New-Item ($RootPath,$AppVendor,$AppName,$AppVersion, 'Resources' -join "\") -Type Directory -Force | Out-Null
    $RootPath,$AppVendor,$AppName,$AppVersion -join "\"
}

Function Get-FolderName {
<#
.SYNOPSIS
   Show a Folder Browser Dialog and return the directory selected by the user

.DESCRIPTION
  Show a Folder Browser Dialog and return the directory selected by the user

.PARAMETER SelectedPath
   Initial Directory for browsing
   Mandatory - [string]

.PARAMETER Description
   Message Box Title
   Optional - [string] - Default : "Select a Folder"

.PARAMETER  ShowNewFolderButton
   Show New Folder Button when unused (default) or doesn't show New Folder when used with $false value
   Optional - [Switch]

 .EXAMPLE
   Get-FolderName
    cmdlet Get-FileFolder at position 1 of the command pipeline
    Provide values for the following parameters:
    SelectedPath: C:\temp
    C:\Temp\

   Choose only one Directory. It's possible to create a new folder (default)

.EXAMPLE
   Get-FolderName -SelectedPath c:\temp -Description "Select a folder" -ShowNewFolderButton
   C:\Temp\Test

   Choose only one Directory. It's possible to create a new folder

.EXAMPLE
   Get-FolderName -SelectedPath c:\temp -Description "Select a folder"
   C:\Temp\Test
   Choose only one Directory. It's not possible to create a new folder

.EXAMPLE
   Get-FolderName  -SelectedPath c:\temp
   C:\Temp\Test

   Choose only one Directory. It's possible to create a new folder (default)


.EXAMPLE
 Get-Help Get-FolderName -Full
#>

[CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Initial Directory for browsing",
            Position = 0)]
        [String]$SelectedPath,

        # Description help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Message Box Title")]
        [String]$Description="Select a Folder",

        # ShowNewFolderButton help description
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Show New Folder Button when used")]
        [Switch]$ShowNewFolderButton
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $FolderBrowser= New-Object System.Windows.Forms.FolderBrowserDialog

   # Define Title
    $FolderBrowser.Description = $Description

    # Define Initial Directory
    if (-Not [String]::IsNullOrWhiteSpace($SelectedPath))
    {
        $FolderBrowser.SelectedPath=$SelectedPath
    }

    if($folderBrowser.ShowDialog() -eq "OK")
    {
        $global:RootPath = $FolderBrowser.SelectedPath
    }
    return $global:RootPath
}

Function UI_Form {
	$form = $null
	$form = New-Object System.Windows.Forms.Form
	$form.Text = 'Package Creator Tool v.'+$MSIVer
	$form.Size = New-Object System.Drawing.Size(820,580)
	$form.MaximumSize = New-Object System.Drawing.Size(820,580)
	$form.MinimumSize = New-Object System.Drawing.Size(820,580)
	#$form.Topmost = $False
	#$form.TopLevel = $False
	$Form.SizeGripStyle = "Hide"
	$form.StartPosition = 'CenterScreen'

	$Font = New-Object System.Drawing.Font("Arial",8,[System.Drawing.FontStyle]::Bold)
	# Font styles are: Regular, Bold, Italic, Underline, Strikeout

	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = New-Object System.Drawing.Point(670,480)
	$okButton.Size = New-Object System.Drawing.Size(95,23)
	$okButton.Text = 'Create package'
	$okButton.Font = $Font
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)

	$mypen = new-object Drawing.Pen blue
	$mypen.width = 3

	$formGraphics = $form.createGraphics()
	$form.add_paint({$formGraphics.DrawLine($mypen, 415, 10, 415, 500)})

	$mypen1 = new-object Drawing.Pen orange
	$mypen1.width = 2

	$formGraphics = $form.createGraphics()
	$form.add_paint({$formGraphics.DrawLine($mypen1, 10, 180, 390, 180)})

	$labelTargetFolder = New-Object System.Windows.Forms.Label
	$labelTargetFolder.Location = New-Object System.Drawing.Point(10,20)
	$labelTargetFolder.Size = New-Object System.Drawing.Size(185,20)
	$labelTargetFolder.Text = 'Select a target folder for package:'
	$labelTargetFolder.Font = $Font
	$form.Controls.Add($labelTargetFolder)

		$textBoxTargetFolder = New-Object System.Windows.Forms.TextBox
		$textBoxTargetFolder.Location = New-Object System.Drawing.Point(10,40)
		$textBoxTargetFolder.Size = New-Object System.Drawing.Size(300,20)
		$textBoxTargetFolder.Text = $global:RootPath
		$form.Controls.Add($textBoxTargetFolder)

			$BrowseButton = New-Object System.Windows.Forms.Button
			$BrowseButton.Location = New-Object System.Drawing.Point(310,40)
			$BrowseButton.Size = New-Object System.Drawing.Size(75,20)
			$BrowseButton.Text = 'Browse'
			$BrowseButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
			$form.Controls.Add($BrowseButton)

	$labelFileBrowser = New-Object System.Windows.Forms.Label
	$labelFileBrowser.Location = New-Object System.Drawing.Point(10,80)
	$labelFileBrowser.Size = New-Object System.Drawing.Size(175,20)
	$labelFileBrowser.Text = 'Select installer file:'
	$labelFileBrowser.Font = $Font
	$form.Controls.Add($labelFileBrowser)

		$textBoxFileBrowser = New-Object System.Windows.Forms.TextBox
		$textBoxFileBrowser.Location = New-Object System.Drawing.Point(10,100)
		$textBoxFileBrowser.Size = New-Object System.Drawing.Size(300,20)
		$textBoxFileBrowser.Text = $global:Installer
		$form.Controls.Add($textBoxFileBrowser)

			$FileBrowseButton = New-Object System.Windows.Forms.Button
			$FileBrowseButton.Location = New-Object System.Drawing.Point(310,100)
			$FileBrowseButton.Size = New-Object System.Drawing.Size(75,20)
			$FileBrowseButton.Text = 'Browse'
			$FileBrowseButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
			$form.Controls.Add($FileBrowseButton)

				$checkboxIncludeFolder = new-object System.Windows.Forms.checkbox
				$checkboxIncludeFolder.Location = new-object System.Drawing.Size(10,125)
				$checkboxIncludeFolder.Size = new-object System.Drawing.Size(300,20)
				$checkboxIncludeFolder.Text = "Copy all content from installer's folder"
				If ($global:CBoxIncludeFolder -eq $true){$checkboxIncludeFolder.Checked = $true}
				else {$checkboxIncludeFolder.Checked = $false}
				$Form.Controls.Add($checkboxIncludeFolder)  

	$labelDetections = New-Object System.Windows.Forms.Label
	$labelDetections.Location = New-Object System.Drawing.Point(110,195)
	$labelDetections.Size = New-Object System.Drawing.Size(250,20)
	$labelDetections.Text = 'Set detection method for MSI/EXE:'
	$labelDetections.Font = $Font
	$form.Controls.Add($labelDetections)

	$labelMSIGUID = New-Object System.Windows.Forms.Label
	$labelMSIGUID.Location = New-Object System.Drawing.Point(10,225)
	$labelMSIGUID.Size = New-Object System.Drawing.Size(100,15)
	$labelMSIGUID.Text = 'MSI product code:'
	$labelMSIGUID.Font = $Font
	$form.Controls.Add($labelMSIGUID)

		$textBoxMSIGUID = New-Object System.Windows.Forms.TextBox
		$textBoxMSIGUID.Location = New-Object System.Drawing.Point(110,225)
		$textBoxMSIGUID.Size = New-Object System.Drawing.Size(250,15)
		$textBoxMSIGUID.Text = $global:MSIProductCode
		$form.Controls.Add($textBoxMSIGUID)

	$labelMSIVer = New-Object System.Windows.Forms.Label
	$labelMSIVer.Location = New-Object System.Drawing.Point(10,245)
	$labelMSIVer.Size = New-Object System.Drawing.Size(70,15)
	$labelMSIVer.Text = 'MSI vesion:'
	$labelMSIVer.Font = $Font
	$form.Controls.Add($labelMSIVer)

		$textBoxMSIVer = New-Object System.Windows.Forms.TextBox
		$textBoxMSIVer.Location = New-Object System.Drawing.Point(110,245)
		$textBoxMSIVer.Size = New-Object System.Drawing.Size(250,15)
		$textBoxMSIVer.Text = $global:MSIVersion
		$form.Controls.Add($textBoxMSIVer)


	$labelOR = New-Object System.Windows.Forms.Label
	$labelOR.Location = New-Object System.Drawing.Point(210,275)
	$labelOR.Size = New-Object System.Drawing.Size(50,20)
	$labelOR.Text = 'OR'
	$labelOR.Font = $Font
	$form.Controls.Add($labelOR)

	$labelDetectDir = New-Object System.Windows.Forms.Label
	$labelDetectDir.Location = New-Object System.Drawing.Point(10,295)
	$labelDetectDir.Size = New-Object System.Drawing.Size(100,15)
	$labelDetectDir.Text = 'EXE file folder:'
	$labelDetectDir.Font = $Font
	$form.Controls.Add($labelDetectDir)

		$textBoxDetectDir = New-Object System.Windows.Forms.TextBox
		$textBoxDetectDir.Location = New-Object System.Drawing.Point(110,295)
		$textBoxDetectDir.Size = New-Object System.Drawing.Size(250,15)
		$textBoxDetectDir.Text = $global:DetectDir
		$form.Controls.Add($textBoxDetectDir)

	$labelDetectFile = New-Object System.Windows.Forms.Label
	$labelDetectFile.Location = New-Object System.Drawing.Point(10,315)
	$labelDetectFile.Size = New-Object System.Drawing.Size(100,15)
	$labelDetectFile.Text = 'EXE file name:'
	$labelDetectFile.Font = $Font
	$form.Controls.Add($labelDetectFile)

		$textBoxDetectFile = New-Object System.Windows.Forms.TextBox
		$textBoxDetectFile.Location = New-Object System.Drawing.Point(110,315)
		$textBoxDetectFile.Size = New-Object System.Drawing.Size(250,15)
		$textBoxDetectFile.Text = $global:DetectFile
		$form.Controls.Add($textBoxDetectFile)

	$labelDetectFileVer = New-Object System.Windows.Forms.Label
	$labelDetectFileVer.Location = New-Object System.Drawing.Point(10,335)
	$labelDetectFileVer.Size = New-Object System.Drawing.Size(100,15)
	$labelDetectFileVer.Text = 'EXE file version:'
	$labelDetectFileVer.Font = $Font
	$form.Controls.Add($labelDetectFileVer)

		$textBoxDetectFileVer = New-Object System.Windows.Forms.TextBox
		$textBoxDetectFileVer.Location = New-Object System.Drawing.Point(110,335)
		$textBoxDetectFileVer.Size = New-Object System.Drawing.Size(250,15)
		$textBoxDetectFileVer.Text = $global:DetectFileVer
		$form.Controls.Add($textBoxDetectFileVer)

			$CheckBoxDetectFile_32x64 = new-object System.Windows.Forms.checkbox
			$CheckBoxDetectFile_32x64.Location = new-object System.Drawing.Size(10,355)
			$CheckBoxDetectFile_32x64.Size = new-object System.Drawing.Size(350,20)
			$CheckBoxDetectFile_32x64.Text = "This file is associated with 32-bit application on 64-bit systems"
			If ($global:CBoxDetectFile_32x64 -eq $true){$CheckBoxDetectFile_32x64.Checked = $true}
			else {$checkboxIncludeFolder.Checked = $false}
			$Form.Controls.Add($CheckBoxDetectFile_32x64)  

	$formGraphics = $form.createGraphics()
	$form.add_paint({$formGraphics.DrawLine($mypen1, 10, 385, 390, 385)})

	$labelDisplayName = New-Object System.Windows.Forms.Label
	$labelDisplayName.Location = New-Object System.Drawing.Point(10,405)
	$labelDisplayName.Size = New-Object System.Drawing.Size(250,20)
	$labelDisplayName.Text = 'Display name (from Add/Remove program):'
	$labelDisplayName.Font = $Font
	$form.Controls.Add($labelDisplayName)

		$textBoxDisplayName = New-Object System.Windows.Forms.TextBox
		$textBoxDisplayName.Location = New-Object System.Drawing.Point(10,425)
		$textBoxDisplayName.Size = New-Object System.Drawing.Size(350,20)
		$textBoxDisplayName.Text = $global:DisplayName
		$form.Controls.Add($textBoxDisplayName)

#RIGHT SIDE

	$labelManufacturer = New-Object System.Windows.Forms.Label
	$labelManufacturer.Location = New-Object System.Drawing.Point(440,20)
	$labelManufacturer.Size = New-Object System.Drawing.Size(280,20)
	$labelManufacturer.Text = 'Manufacturer:'
	$labelManufacturer.Font = $Font
	$form.Controls.Add($labelManufacturer)

		$textBoxManufacturer = New-Object System.Windows.Forms.TextBox
		$textBoxManufacturer.Location = New-Object System.Drawing.Point(440,40)
		$textBoxManufacturer.Size = New-Object System.Drawing.Size(325,20)
		$textBoxManufacturer.Text = $global:Manufacturer
		$form.Controls.Add($textBoxManufacturer)

	$labelProductName = New-Object System.Windows.Forms.Label
	$labelProductName.Location = New-Object System.Drawing.Point(440,80)
	$labelProductName.Size = New-Object System.Drawing.Size(280,20)
	$labelProductName.Text = 'ProductName (like Reader DC):'
	$labelProductName.Font = $Font
	$form.Controls.Add($labelProductName)

		$textBoxProductName = New-Object System.Windows.Forms.TextBox
		$textBoxProductName.Location = New-Object System.Drawing.Point(440,100)
		$textBoxProductName.Size = New-Object System.Drawing.Size(325,20)
		$textBoxProductName.Text = $global:ProductName
		$form.Controls.Add($textBoxProductName)

	$labelVersion = New-Object System.Windows.Forms.Label
	$labelVersion.Location = New-Object System.Drawing.Point(440,140)
	$labelVersion.Size = New-Object System.Drawing.Size(280,20)
	$labelVersion.Text = 'Version:'
	$labelVersion.Font = $Font
	$form.Controls.Add($labelVersion)

		$textBoxVersion = New-Object System.Windows.Forms.TextBox
		$textBoxVersion.Location = New-Object System.Drawing.Point(440,160)
		$textBoxVersion.Size = New-Object System.Drawing.Size(325,20)
		$textBoxVersion.Text = $global:Version
		$form.Controls.Add($textBoxVersion)

	$labelArchitecture = New-Object System.Windows.Forms.Label
	$labelArchitecture.Location = New-Object System.Drawing.Point(440,200)
	$labelArchitecture.Size = New-Object System.Drawing.Size(80,20)
	$labelArchitecture.Text = 'Architecture:'
	$labelArchitecture.Font = $Font
	$form.Controls.Add($labelArchitecture)

		$checkbox64 = new-object System.Windows.Forms.checkbox
		$checkbox64.Location = new-object System.Drawing.Size(530,198)
		$checkbox64.Size = new-object System.Drawing.Size(60,20)
		$checkbox64.Text = "x64"
		If ($global:Architecture -eq 'All'){$checkbox64.Checked = $true}
		If ($global:Architecture -eq 'x64'){$checkbox64.Checked = $true}
		If ($global:Architecture -eq 'x86'){$checkbox64.Checked = $false}
		$Form.Controls.Add($checkbox64)  

		$checkbox86 = new-object System.Windows.Forms.checkbox
		$checkbox86.Location = new-object System.Drawing.Size(590,198)
		$checkbox86.Size = new-object System.Drawing.Size(60,20)
		$checkbox86.Text = "x86"
		If ($global:Architecture -eq 'All'){$checkbox86.Checked = $true}
		If ($global:Architecture -eq 'x64'){$checkbox86.Checked = $false}
		If ($global:Architecture -eq 'x86'){$checkbox86.Checked = $true}
		$Form.Controls.Add($checkbox86) 

	$labelLanguage = New-Object System.Windows.Forms.Label
	$labelLanguage.Location = New-Object System.Drawing.Point(440,230)
	$labelLanguage.Size = New-Object System.Drawing.Size(155,20)
	$labelLanguage.Text = 'Language (En,Fr,En-Fr,ML):'
	$labelLanguage.Font = $Font
	$form.Controls.Add($labelLanguage)

		$textBoxLanguage = New-Object System.Windows.Forms.TextBox
		$textBoxLanguage.Location = New-Object System.Drawing.Point(595,230)
		$textBoxLanguage.Size = New-Object System.Drawing.Size(170,20)
		$textBoxLanguage.Text = $global:Language
		$form.Controls.Add($textBoxLanguage)

	$labelRevision = New-Object System.Windows.Forms.Label
	$labelRevision.Location = New-Object System.Drawing.Point(440,260)
	$labelRevision.Size = New-Object System.Drawing.Size(55,20)
	$labelRevision.Text = 'Revision'
	$labelRevision.Font = $Font
	$form.Controls.Add($labelRevision)

		$textBoxRevision = New-Object System.Windows.Forms.TextBox
		$textBoxRevision.Location = New-Object System.Drawing.Point(495,258)
		$textBoxRevision.Size = New-Object System.Drawing.Size(25,20)
		$textBoxRevision.Text = $global:Revision
		$form.Controls.Add($textBoxRevision)

	$labelZoneCode = New-Object System.Windows.Forms.Label
	$labelZoneCode.Location = New-Object System.Drawing.Point(440,290)
	$labelZoneCode.Size = New-Object System.Drawing.Size(210,20)
	$labelZoneCode.Text = 'ZoneCode, for zone-specific package:'
	$labelZoneCode.Font = $Font
	$form.Controls.Add($labelZoneCode)


		$RadioButton1 = New-Object System.Windows.Forms.RadioButton
		$RadioButton1.Location = '650,288'
		$RadioButton1.size = '50,20'
		if ($global:ZoneCode -eq ''){$RadioButton1.Checked=$true}
		$RadioButton1.Text = "None"
		$form.Controls.Add($RadioButton1)

		$RadioButton2 = New-Object System.Windows.Forms.RadioButton
		$RadioButton2.Location = '720,288'
		$RadioButton2.size = '65,20'
		if ($global:ZoneCode -eq '_MAIN'){$RadioButton2.Checked=$true}
		$RadioButton2.Text = "MAIN"
		$form.Controls.Add($RadioButton2) 

		$RadioButton3 = New-Object System.Windows.Forms.RadioButton
		$RadioButton3.Location = '440,318'
		$RadioButton3.size = '60,20'
		if ($global:ZoneCode -eq '_ZN01'){$RadioButton3.Checked=$true} 
		$RadioButton3.Text = "ZN01"
		$form.Controls.Add($RadioButton3)

		$RadioButton4 = New-Object System.Windows.Forms.RadioButton
		$RadioButton4.Location = '440,338'
		$RadioButton4.size = '60,20'
		if ($global:ZoneCode -eq '_ZN02'){$RadioButton4.Checked=$true} 
		$RadioButton4.Text = "ZN02"
		$form.Controls.Add($RadioButton4)

		$RadioButton5 = New-Object System.Windows.Forms.RadioButton
		$RadioButton5.Location = '510,318'
		$RadioButton5.size = '60,20'
		if ($global:ZoneCode -eq '_ZN03'){$RadioButton5.Checked=$true} 
		$RadioButton5.Text = "ZN03"
		$form.Controls.Add($RadioButton5)

		$RadioButton6 = New-Object System.Windows.Forms.RadioButton
		$RadioButton6.Location = '510,338'
		$RadioButton6.size = '60,20'
		if ($global:ZoneCode -eq '_ZN04'){$RadioButton6.Checked=$true} 
		$RadioButton6.Text = "ZN04"
		$form.Controls.Add($RadioButton6)

		$RadioButton7 = New-Object System.Windows.Forms.RadioButton
		$RadioButton7.Location = '580,318'
		$RadioButton7.size = '60,20'
		if ($global:ZoneCode -eq '_ZN51'){$RadioButton7.Checked=$true}
		$RadioButton7.Text = "ZN51"
		$form.Controls.Add($RadioButton7)

		$RadioButton8 = New-Object System.Windows.Forms.RadioButton
		$RadioButton8.Location = '580,338'
		$RadioButton8.size = '60,20'
		if ($global:ZoneCode -eq '_ZN05'){$RadioButton8.Checked=$true} 
		$RadioButton8.Text = "ZN05"
		$form.Controls.Add($RadioButton8)

		$RadioButton9 = New-Object System.Windows.Forms.RadioButton
		$RadioButton9.Location = '650,318'
		$RadioButton9.size = '60,20'
		if ($global:ZoneCode -eq '_ZN06'){$RadioButton9.Checked=$true} 
		$RadioButton9.Text = "ZN06"
		$form.Controls.Add($RadioButton9)

		$RadioButton10 = New-Object System.Windows.Forms.RadioButton
		$RadioButton10.Location = '650,338'
		$RadioButton10.size = '60,20'
		if ($global:ZoneCode -eq '_ZN07'){$RadioButton10.Checked=$true}
		$RadioButton10.Text = "ZN07"
		$form.Controls.Add($RadioButton10)

		$RadioButton11 = New-Object System.Windows.Forms.RadioButton
		$RadioButton11.Location = '720,318'
		$RadioButton11.size = '60,20'
		if ($global:ZoneCode -eq '_ZN08'){$RadioButton11.Checked=$true}
		$RadioButton11.Text = "ZN08"
		$form.Controls.Add($RadioButton11)

		$RadioButton12 = New-Object System.Windows.Forms.RadioButton
		$RadioButton12.Location = '720,338'
		$RadioButton12.size = '60,20'
		if ($global:ZoneCode -eq '_ZN09'){$RadioButton12.Checked=$true} 
		$RadioButton12.Text = "ZN09"
		$form.Controls.Add($RadioButton12)

		$RadioButton13 = New-Object System.Windows.Forms.RadioButton
		$RadioButton13.Location = '440,368'
		$RadioButton13.size = '60,20'
		if ($global:ZoneCode -eq '_ZN10'){$RadioButton13.Checked=$true}
		$RadioButton13.Text = "ZN10"
		$form.Controls.Add($RadioButton13)

		$RadioButton14 = New-Object System.Windows.Forms.RadioButton
		$RadioButton14.Location = '440,388'
		$RadioButton14.size = '60,20'
		if ($global:ZoneCode -eq '_ZN11'){$RadioButton14.Checked=$true}
		$RadioButton14.Text = "ZN11"
		$form.Controls.Add($RadioButton14)



		$RadioButton15 = New-Object System.Windows.Forms.RadioButton
		$RadioButton15.Location = '510,368'
		$RadioButton15.size = '60,20'
		if ($global:ZoneCode -eq '_ZN12'){$RadioButton15.Checked=$true} 
		$RadioButton15.Text = "ZN12"
		$form.Controls.Add($RadioButton15)

		$RadioButton16 = New-Object System.Windows.Forms.RadioButton
		$RadioButton16.Location = '510,388'
		$RadioButton16.size = '60,20'
		if ($global:ZoneCode -eq '_ZN13'){$RadioButton16.Checked=$true} 
		$RadioButton16.Text = "ZN13"
		$form.Controls.Add($RadioButton16)



		$RadioButton17 = New-Object System.Windows.Forms.RadioButton
		$RadioButton17.Location = '580,368'
		$RadioButton17.size = '60,20'
		if ($global:ZoneCode -eq '_ZN14'){$RadioButton17.Checked=$true}
		$RadioButton17.Text = "ZN14"
		#$form.Controls.Add($RadioButton17)

		$RadioButton18 = New-Object System.Windows.Forms.RadioButton
		$RadioButton18.Location = '580,388'
		$RadioButton18.size = '60,20'
		if ($global:ZoneCode -eq '_ZN15'){$RadioButton18.Checked=$true} 
		$RadioButton18.Text = "ZN15"
		$form.Controls.Add($RadioButton18)


#END RIGHT SIDE

###################### RUN UI ######################################
	$result = $form.ShowDialog()
	# 'Cancel' pressed
	if ($result -eq [System.Windows.Forms.DialogResult]::Cancel){
		exit
		}
		
	# 'Browse' pressed
	if ($result -eq [System.Windows.Forms.DialogResult]::Retry){
		Get-FolderName -SelectedPath c:
		$global:Manufacturer = $textBoxManufacturer.Text
		$global:ProductName = $textBoxProductName.Text
		$global:Version = $textBoxVersion.Text
		$global:Language = $textBoxLanguage.Text
		$global:Revision = $textBoxRevision.Text
		$global:SCCM_SiteCode = $textBoxSiteCode.Text
		$global:DisplayName = $textBoxDisplayName.Text

		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='All'}
		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $false)) {$global:Architecture='x64'}
		If (($checkbox64.Checked -eq $false) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='x86'}
		If ($checkboxIncludeFolder.Checked -eq $true) {$global:CBoxIncludeFolder=$true}
		else {$global:CBoxIncludeFolder=$false}
		If ($CBoxDetectFile_32x64.Checked -eq $true) {$global:CBoxDetectFile_32x64=$true}
		else {$global:CBoxDetectFile_32x64=$false}

		if ($RadioButton1.Checked -eq $true) {$global:ZoneCode=''}
		if ($RadioButton2.Checked -eq $true) {$global:ZoneCode='_MAIN'}
		if ($RadioButton3.Checked -eq $true) {$global:ZoneCode='_ZN01'}
		if ($RadioButton4.Checked -eq $true) {$global:ZoneCode='_ZN02'}
		if ($RadioButton5.Checked -eq $true) {$global:ZoneCode='_ZN03'}
		if ($RadioButton6.Checked -eq $true) {$global:ZoneCode='_ZN04'}
		if ($RadioButton7.Checked -eq $true) {$global:ZoneCode='_ZN51'}
		if ($RadioButton8.Checked -eq $true) {$global:ZoneCode='_ZN05'}
		if ($RadioButton9.Checked -eq $true) {$global:ZoneCode='_ZN06'}
		if ($RadioButton10.Checked -eq $true) {$global:ZoneCode='_ZN07'}
		if ($RadioButton11.Checked -eq $true) {$global:ZoneCode='_ZN08'}
		if ($RadioButton12.Checked -eq $true) {$global:ZoneCode='_ZN09'}
		if ($RadioButton13.Checked -eq $true) {$global:ZoneCode='_ZN10'}
		if ($RadioButton14.Checked -eq $true) {$global:ZoneCode='_ZN11'}
		if ($RadioButton15.Checked -eq $true) {$global:ZoneCode='_ZN12'}
		if ($RadioButton16.Checked -eq $true) {$global:ZoneCode='_ZN13'}
		if ($RadioButton17.Checked -eq $true) {$global:ZoneCode='_ZN14'}
		if ($RadioButton18.Checked -eq $true) {$global:ZoneCode='_ZN15'}



		UI_Form
		}

	# 'BrowseFile' pressed
	if ($result -eq [System.Windows.Forms.DialogResult]::Yes){
		$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
			InitialDirectory = [Environment]::GetFolderPath('Desktop') 
			#Filter = 'MSI (*.msi)|*.msi|EXE (*.exe)|*.exe'
			Filter = 'MSI (*.msi); EXE (*.exe)|*.msi;*.exe'}
		$null = $FileBrowser.ShowDialog()
		$global:Installer = $FileBrowser.FileName

		If ((Get-Item $FileBrowser.FileName).Extension -eq '.msi'){
			$MSIObj = Get-MSIFileInformation -FilePath $global:Installer
			<#
			FileName        : InstEdPlus-1.1.9.10.msi
			FilePath        : T:\#PackagingTools\InstEdPlus-1.1.9.10.msi
			Length(MB)      : 2.8828125
			ProductVersion  : 1.1.9.10
			ProductCode     : {2ADB3488-300B-471A-80C0-C0FFAB8222E5}
			ProductName     : InstEd Plus 1.1.9.10
			Manufacturer    : instedit.com
			ProductLanguage : 1033 (English)
			UpgradeCode     : {678AC25E-6AB9-4E93-8807-C2DF99ABFCA8}
			Codepage        : 1252
			Title           : Installation Database
			Subject         : InstEd Plus 1.1.9.10
			Author          : instedit.com
			Keywords        : Installer
			Comment         : This installer database contains the logic and data required to install InstEd Plus 1.1.9.10.
				  
			Template        : Intel;1033
			LastAuthor      : instedit.com
			RevisionNumber  : {F8F2F05D-5EDB-4A06-ACD6-3F2159DCF4E4}
			EditTime        : 
			LastPrinted     : 
			CreationDate    : 4/30/2012 8:49:04 AM
			LastSaved       : 4/30/2012 4:56:34 PM
			PageCount       : 200
			WordCount       : 2
			#>

			$Platform=($MSIObj.template -Split �;�)[0]
			If($Platform -eq 'Intel'){
				$global:Architecture='x86'
				$checkbox64.Checked = $false
				$checkbox86.Checked = $true
				}
			If($Platform -eq ''){
				$global:Architecture=''
				$checkbox64.Checked = $false
				$checkbox86.Checked = $false
				}
			$global:Manufacturer = $MSIObj.Manufacturer
			$global:ProductName = $MSIObj.ProductName
			$global:Version = $MSIObj.ProductVersion
			$global:MSIProductCode = $MSIObj.ProductCode
			$global:MSIVersion = $MSIObj.ProductVersion
			$global:DisplayName = $MSIObj.ProductName

			$global:DetectDir = ''
			$global:DetectFile = ''
			$global:DetectFileVer = ''
			}
		else{
			$global:MSIProductCode = ''
			$global:MSIVersion = ''
			$global:DetectDir = 'C:\Program Files\Application'
			$global:DetectFile = 'App.exe'
			$global:DetectFileVer = '1.0.0'
			$global:Manufacturer = 'Manufacturer name'
			$global:ProductName = 'Product name'
			$global:Version = 'xx.xxx.xxx'
			$global:DisplayName = '<Copy_from_ARP>'
			}

		$global:Language = $textBoxLanguage.Text
		$global:Revision = $textBoxRevision.Text
		$global:SCCM_SiteCode = $textBoxSiteCode.Text

		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='All'}
		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $false)) {$global:Architecture='x64'}
		If (($checkbox64.Checked -eq $false) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='x86'}
		If ($checkboxIncludeFolder.Checked -eq $true) {$global:CBoxIncludeFolder=$true}
		else {$global:CBoxIncludeFolder=$false}
		If ($CBoxDetectFile_32x64.Checked -eq $true) {$global:CBoxDetectFile_32x64=$true}
		else {$global:CBoxDetectFile_32x64=$false}

		if ($RadioButton1.Checked -eq $true) {$global:ZoneCode=''}
		if ($RadioButton2.Checked -eq $true) {$global:ZoneCode='_MAIN'}
		if ($RadioButton3.Checked -eq $true) {$global:ZoneCode='_ZN01'}
		if ($RadioButton4.Checked -eq $true) {$global:ZoneCode='_ZN02'}
		if ($RadioButton5.Checked -eq $true) {$global:ZoneCode='_ZN03'}
		if ($RadioButton6.Checked -eq $true) {$global:ZoneCode='_ZN04'}
		if ($RadioButton7.Checked -eq $true) {$global:ZoneCode='_ZN51'}
		if ($RadioButton8.Checked -eq $true) {$global:ZoneCode='_ZN05'}
		if ($RadioButton9.Checked -eq $true) {$global:ZoneCode='_ZN06'}
		if ($RadioButton10.Checked -eq $true) {$global:ZoneCode='_ZN07'}
		if ($RadioButton11.Checked -eq $true) {$global:ZoneCode='_ZN08'}
		if ($RadioButton12.Checked -eq $true) {$global:ZoneCode='_ZN09'}
		if ($RadioButton13.Checked -eq $true) {$global:ZoneCode='_ZN10'}
		if ($RadioButton14.Checked -eq $true) {$global:ZoneCode='_ZN11'}
		if ($RadioButton15.Checked -eq $true) {$global:ZoneCode='_ZN12'}
		if ($RadioButton16.Checked -eq $true) {$global:ZoneCode='_ZN13'}
		if ($RadioButton17.Checked -eq $true) {$global:ZoneCode='_ZN14'}
		if ($RadioButton18.Checked -eq $true) {$global:ZoneCode='_ZN15'}

		UI_Form
		}

	# 'Create package' pressed
	if ($result -eq [System.Windows.Forms.DialogResult]::OK){
		$global:Manufacturer = $textBoxManufacturer.Text
		$global:ProductName = $textBoxProductName.Text
		$global:Version = $textBoxVersion.Text
		$global:Language = $textBoxLanguage.Text
		$global:Revision = $textBoxRevision.Text
		$global:DisplayName = $textBoxDisplayName.Text

		$global:MSIProductCode = $textBoxMSIGUID.Text
		$global:MSIVersion = $textBoxMSIVer.Text
		$global:DetectDir = $textBoxDetectDir.Text
		$global:DetectFile = $textBoxDetectFile.Text
		$global:DetectFileVer = $textBoxDetectFileVer.Text

		$global:Installer = $textBoxFileBrowser.Text
		$global:CBoxIncludeFolder = $checkboxIncludeFolder.Checked
		$global:CBoxDetectFile_32x64 = $checkboxDetectFile_32x64.Checked

		if ($RadioButton1.Checked -eq $true) {$global:ZoneCode=''}
		if ($RadioButton2.Checked -eq $true) {$global:ZoneCode='_MAIN'}
		if ($RadioButton3.Checked -eq $true) {$global:ZoneCode='_ZN01'}
		if ($RadioButton4.Checked -eq $true) {$global:ZoneCode='_ZN02'}
		if ($RadioButton5.Checked -eq $true) {$global:ZoneCode='_ZN03'}
		if ($RadioButton6.Checked -eq $true) {$global:ZoneCode='_ZN04'}
		if ($RadioButton7.Checked -eq $true) {$global:ZoneCode='_ZN04'}
		if ($RadioButton8.Checked -eq $true) {$global:ZoneCode='_ZN05'}
		if ($RadioButton9.Checked -eq $true) {$global:ZoneCode='_ZN06'}
		if ($RadioButton10.Checked -eq $true) {$global:ZoneCode='_ZN07'}
		if ($RadioButton11.Checked -eq $true) {$global:ZoneCode='_ZN08'}
		if ($RadioButton12.Checked -eq $true) {$global:ZoneCode='_ZN09'}
		if ($RadioButton13.Checked -eq $true) {$global:ZoneCode='_ZN10'}

		if ($RadioButton14.Checked -eq $true) {$global:ZoneCode='_ZN11'}
		if ($RadioButton15.Checked -eq $true) {$global:ZoneCode='_ZN12'}
		if ($RadioButton16.Checked -eq $true) {$global:ZoneCode='_ZN13'}
		if ($RadioButton17.Checked -eq $true) {$global:ZoneCode='_ZN14'}
		if ($RadioButton18.Checked -eq $true) {$global:ZoneCode='_ZN15'}



		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='All'}
		If (($checkbox64.Checked -eq $true) -and ($checkbox86.Checked -eq $false)) {$global:Architecture='x64'}
		If (($checkbox64.Checked -eq $false) -and ($checkbox86.Checked -eq $true)) {$global:Architecture='x86'}
		If ($checkboxIncludeFolder.Checked -eq $true) {$global:CBoxIncludeFolder=$true}
		If ($checkboxDetectFile_32x64.Checked -eq $true) {$global:CBoxDetectFile_32x64=$true}

		If (($checkbox64.Checked -eq $false) -and ($checkbox86.Checked -eq $false)) {
			[System.Windows.MessageBox]::Show('Architecture not set',"",0,48)
			UI_Form
			}            
		If ($global:Manufacturer -eq '') {
			[System.Windows.MessageBox]::Show('Vendor not set',"",0,48)
			UI_Form
			}
		If ($global:ProductName -eq '') {
			[System.Windows.MessageBox]::Show('ProductName not set',"",0,48)
			UI_Form
			}
		If ($global:Version -eq '') {
			[System.Windows.MessageBox]::Show('Version not set',"",0,48)
			UI_Form
			}
		If ($global:Language -eq '') {
			[System.Windows.MessageBox]::Show('Language not set',"",0,48)
			UI_Form
			}
		If ($global:Revision -eq '') {
			[System.Windows.MessageBox]::Show('Revision not set',"",0,48)
			UI_Form
			}
		}
}

Function CreateApplicationObj {
    $_AppObject = "" | Select-Object -Property AppVendor, AppName, AppVersion, AppArch, AppLang, AppRev, AppZone, LocalizedName, DeploymentTypeName, MSIProductCode, MSIVersion, DetectDir, DetectFile, DetectFileVer, AppPackageName, DisplayName
    $_AppObject.AppVendor = $global:Manufacturer
    $_AppObject.AppName = $global:ProductName
    $_AppObject.AppVersion = $global:Version
    $_AppObject.AppArch = $global:Architecture
    $_AppObject.AppLang = $global:Language
    $_AppObject.AppRev = $global:Revision
    $_AppObject.AppZone = $global:ZoneCode
    $_AppObject.DisplayName = $global:DisplayName
    $_AppObject.LocalizedName = $LocalizedName
    $_AppObject.DeploymentTypeName = $DeploymentTypeName
    $_AppObject.MSIProductCode = $MSIProductCode
    $_AppObject.MSIVersion = $MSIVersion
    $_AppObject.DetectDir = $DetectDir
    $_AppObject.DetectFile = $DetectFile
    $_AppObject.DetectFileVer = $DetectFileVer

    If ($global:ZoneCode -ne ""){$_AppObject.AppPackageName = ($global:Manufacturer -replace '\s','') +"_" + ($global:ProductName -replace '\s','') +"_" + $global:Version +"_" + $global:Architecture +"_" + $global:Language +"_" + $global:Revision + $global:ZoneCode}
    else {$_AppObject.AppPackageName = ($global:Manufacturer -replace '\s','') +"_" + ($global:ProductName -replace '\s','') +"_" + $global:Version +"_" + $global:Architecture +"_" + $global:Language +"_" + $global:Revision}

    If ($global:ZoneCode -ne ""){$_AppObject.LocalizedName = ($global:Manufacturer) +" " + ($global:ProductName) +" " + $global:Version +" " + $global:Architecture +" " + $global:Language +" " + $global:Revision + ($global:ZoneCode -replace '_',' ')}
    else {$_AppObject.LocalizedName = ($global:Manufacturer) +" " + ($global:ProductName) +" " + $global:Version +" " + $global:Architecture +" " + $global:Language +" " + $global:Revision}

    $_AppObject.DeploymentTypeName = ($global:Manufacturer -replace '\s','') +"_" + ($global:ProductName -replace '\s','') +"_" + $global:Version +"_" + $global:Architecture +"_" + $global:Language +"-I"

    return $_AppObject
}

Function CreateFolderStructure {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Object[]]$InputObject
    )

	$InstallationSource = CreateFolders -RootPath $RootPath
	If (-not (Test-Path -Path $InstallationSource)){
		[System.Windows.MessageBox]::Show("The folder $($InstallationSource) was not created`n","",0,16)
		exit
		}

	$PackageDocName=$InputObject.AppPackageName + " - Documentation.docx"

	Copy-Item -Path ($scriptPath, "Templates", "Package" -join "\") -Destination ($InstallationSource -join "\") -Recurse 

	If ($global:Installer){
		If ($global:CBoxIncludeFolder){
			Copy-item -Force -Recurse -Path ((get-item $Installer).DirectoryName+"\*") -Destination ($InstallationSource, "Package", "Files" -join "\")
			}
		else{
			Copy-Item -Path $global:Installer -Destination ($InstallationSource, "Package", "Files" -join "\")     
			}
		}

	Copy-Item -Path ($scriptPath, "Templates", "PackageName - Documentation.docx" -join "\") -Destination ($InstallationSource, "Documentation", $PackageDocName -join "\") 

	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<date>',(Get-Date -UFormat "%m/%d/%Y") } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<UserName>',$UserPUID } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppVendor>',$InputObject.AppVendor } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppName>',$InputObject.AppName } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppVersion>',$InputObject.AppVersion } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppArch>',$InputObject.AppArch } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppLang>',$InputObject.AppLang } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppRev>',($InputObject.AppRev+$InputObject.AppZone) } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<AppPackageName>',$InputObject.AppPackageName } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
	(Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace '<Copy_from_ARP>',$InputObject.DisplayName } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
    
    If ($Installer -ne ''){
	    If ((Get-Item $Installer -ErrorAction SilentlyContinue).Extension -eq '.msi'){
		    (Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace "#Execute-MSI -Action 'Install' -Path '<MSI_File>.msi' -private",("Execute-MSI -Action 'Install' -Path '"+(Get-Item $Installer).Name+"' -private") } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
		    (Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace "#Execute-MSI -Action 'Uninstall' -Path '{AAAAAAAA-BBBB-CCCC-DDDDD-000001000000}'",("Execute-MSI -Action 'Uninstall' -Path '"+$InputObject.MSIProductCode+"'")} ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
		    }
	    If ((Get-Item $Installer -ErrorAction SilentlyContinue).Extension -eq '.exe'){
		    (Get-Content $InstallationSource'\Package\Deploy-Application.ps1' | ForEach-Object { $_ -replace "#Execute-Process -Path '<EXE_File>.exe'",("Execute-Process -Path '"+(Get-Item $Installer).Name+"'") } ) | Set-Content $InstallationSource'\Package\Deploy-Application.ps1'
		    }
    }

	#*****************************************************************************************************


	$test_Install = ($InstallationSource, 'Resources', 'test_Install.ps1' -join "\")
	$test_Uninstall = ($InstallationSource, 'Resources', 'test_Uninstall.ps1' -join "\")
	$deploy_SCCM =  ($InstallationSource, 'Resources', 'Deploy_SCCM.ps1' -join "\")
	$convert_Intune =  ($InstallationSource, 'Resources', 'Convert_Intunewin.ps1' -join "\")

	#template for deployment script
	$tmpl_deploy = ($scriptPath, "Templates", "Deploy_template.ps1" -join "\")
	#template for Convert to Intune script
	$tmpl_intune = ($scriptPath, "Templates", "Convert_Intunewin_template.ps1" -join "\")
	$IntuneWinAppUtil =  ($scriptPath, "Templates", 'IntuneWinAppUtil.exe' -join "\")

	# Create test-instal/uninstal scripts
	If (Test-Path $test_Install){Remove-Item $test_Install}; If (Test-Path $test_Uninstall){Remove-Item $test_Uninstall}
	Add-Content $test_Install, $test_Uninstall ('Clear-Host')
	Add-Content $test_Install, $test_Uninstall ('Write-Host "Copying ..." -foregroundcolor "DarkGreen"')    
    
	Add-Content $test_Install, $test_Uninstall ('$SourceFolder = (get-item (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)).parent.FullName + ''\Package*''')
	#Add-Content $test_Install, $test_Uninstall ('$SourceFolder = ' + "`"$InstallationSource\Package*`"")
	Add-Content $test_Install, $test_Uninstall ('$TempInstall = '+"`"C:\TempInstall`"")
	Add-Content $test_Install, $test_Uninstall ('$DestFolder = '+"`""+'$TempInstall'+"\" +$InputObject.AppPackageName + "\Package`"")
	Add-Content $test_Install ('$Logname = '+"`""+$InputObject.AppPackageName+"_Install.log`"")
	Add-Content $test_Uninstall ('$Logname = '+"`""+$InputObject.AppPackageName+"_Uninstall.log`"")
	Add-Content $test_Install, $test_Uninstall ('$SourcePsexec = '+"`"c:\ProgramData\MyCompany\PackageCreator\Templates\psexec.exe`"")
	Add-Content $test_Install, $test_Uninstall ('$Username = $env:USERDOMAIN +''\'' +$env:username; If (-not($env:username -Like ''*t2adm_*'')){$Username = $env:USERDOMAIN +''\'' +''t2adm_''+ $env:username}')
	Add-Content $test_Install, $test_Uninstall ("If (Test-Path "+'$DestFolder'+"){Remove-Item "+'$DestFolder -Recurse -Force}')
	Add-Content $test_Install, $test_Uninstall ('copy-item $SourceFolder $DestFolder -force -recurse')
	Add-Content $test_Install, $test_Uninstall ('copy-item $SourcePsexec $TempInstall -force')
	Add-Content $test_Install, $test_Uninstall 'Set-Location $TempInstall'
	Add-Content $test_Install ('$arg=''-accepteula -i -s cmd.exe /c '' + """"+ $DestFolder + ''\Deploy-Application.exe'' + """" + '' >> '' + $TempInstall+''\''+ $Logname')
	Add-Content $test_Uninstall ('$arg=''-accepteula -i -s cmd.exe /c '' + """"+ $DestFolder + ''\Deploy-Application.exe'' + """" + '' uninstall'' + '' >> '' + $TempInstall+''\''+ $Logname')
	Add-Content $test_Install, $test_Uninstall ('Start-Process -FilePath ($TempInstall+''\PsExec.exe'') -ArgumentList $arg -Verb runas')

	# Create Intune convert Script
	If (Test-Path $convert_Intune){Remove-Item $convert_Intune}
	(get-content $tmpl_intune) | foreach-object {$_ -replace "<tmpl_intune>", "`"$IntuneWinAppUtil`""} | set-content $convert_Intune
	(get-content $convert_Intune) | foreach-object {$_ -replace '<tmpl_intune_SCCMPackageName>', $InputObject.AppPackageName} | set-content $convert_Intune

	# Create Deployment Script
	If (Test-Path $deploy_SCCM){Remove-Item $deploy_SCCM}
	(get-content $tmpl_deploy | ForEach-Object { $_ -replace '<date>',(Get-Date -UFormat "%m/%d/%Y") } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<UserName>',$UserPUID } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppVendor>',$InputObject.AppVendor } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppName>',$InputObject.AppName } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppVersion>',$InputObject.AppVersion } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppArch>',$InputObject.AppArch } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppLang>',$InputObject.AppLang } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppRev>',($InputObject.AppRev+$InputObject.AppZone) } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<AppPackageName>',$InputObject.AppPackageName } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DefaultSiteCode>',$SCCM_SiteCode } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DefaultPkgPath>',$SCCM_Repository } ) | Set-Content $deploy_SCCM

	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<LocalizedName>',$InputObject.LocalizedName } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DeploymentTypeName>',$InputObject.DeploymentTypeName } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<MSIProductCode>',$InputObject.MSIProductCode } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<MSIVersion>',$InputObject.MSIVersion } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DetectDir>',$InputObject.DetectDir } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DetectFile>',$InputObject.DetectFile } ) | Set-Content $deploy_SCCM
	(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<DetectFileVer>',$InputObject.DetectFileVer } ) | Set-Content $deploy_SCCM


	If($global:CBoxDetectFile_32x64){
		(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<CheckBoxDetectFile_32x64_value>','$True' } ) | Set-Content $deploy_SCCM    
		}
	else{
		(get-content $deploy_SCCM| ForEach-Object { $_ -replace '<CheckBoxDetectFile_32x64_value>','$False' } ) | Set-Content $deploy_SCCM    
		}

	#*****************************************************************************************************

	[System.Windows.MessageBox]::Show('Package created: ' + """$InstallationSource""","",0,64)
}

Function MAIN {
    #Create UI form
    UI_Form
    #Create application object
    $_AppObject = CreateApplicationObj
    #Create folders
    CreateFolderStructure -inputObject $_AppObject
<#
    #Re-initialise variables 
    $global:Manufacturer = 'Manufacturer name'
    $global:ProductName = 'Product name'
    $global:Version = 'xx.xxx.xxx'
    $global:Language = 'En'
    $global:Architecture = 'x64'
    $global:Revision = '001'
    $global:ZoneCode = ''
    $global:Installer = ''	
    $global:CBoxIncludeFolder = $false
    $global:CBoxDetectFile_32x64 = $false
    $global:DisplayName = '<Copy_from_ARP>'
    $global:MSIProductCode = '00000000-0000-0000-0000-000000000000'
    $global:MSIVersion = '1.0.0'
    $global:DetectDir = 'C:\Program Files\Application'
    $global:DetectFile = 'App.exe'
    $global:DetectFileVer = '1.0.0'
#>
    #recursive call
    MAIN
} 

############ INITIALIZATION ###############
#Get current user PUID
If ($env:USERNAME.length -le 7) {$UserPUID = $env:USERNAME}
else {$UserPUID = $env:USERNAME.substring(6)} 

#Get current script folder
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

#Get target folder
$RegRootPath=Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\MyCompany\PackageCreator\' -Name TargetFolder -ErrorAction SilentlyContinue
If (Test-Path -Path $RegRootPath){$global:RootPath = $RegRootPath}
#Get current version of 'PackageCreator'
$global:MSIVer=Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\MyCompany\PackageCreator\' -Name MSIVer -ErrorAction SilentlyContinue
#Get SCCM SiteCode
$global:SCCM_SiteCode=Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\MyCompany\PackageCreator\' -Name SCCM_SiteCode -ErrorAction SilentlyContinue
#Get SCCM DP Repository path ('\\MECMServer\MECMSources$\Software')
$global:SCCM_Repository=Get-ItemPropertyValue -Path 'HKCU:\SOFTWARE\MyCompany\PackageCreator\' -Name SCCM_Repository -ErrorAction SilentlyContinue

[string]$global:Manufacturer = 'Manufacturer name'
[string]$global:ProductName = 'Product name'
[string]$global:Version = 'xx.xxx.xxx'
[string]$global:Language = 'En'
[string]$global:Architecture = 'x64'
[string]$global:Revision = '001'
[string]$global:ZoneCode = ''
[string]$global:Installer = ''
[bool]$global:CBoxIncludeFolder = $false
[bool]$global:CBoxDetectFile_32x64 = $false
[string]$global:DisplayName = '<Copy_from_ARP>'
[string]$global:MSIProductCode = '00000000-0000-0000-0000-000000000000'
[string]$global:MSIVersion = '1.0.0'
[string]$global:DetectDir = 'C:\Program Files\Application'
[string]$global:DetectFile = 'App.exe'
[string]$global:DetectFileVer = '1.0.0'
[string]$global:LocalizedName = ''
[string]$global:DeploymentTypeName = ''

############ INITIALIZATION END ###############

MAIN

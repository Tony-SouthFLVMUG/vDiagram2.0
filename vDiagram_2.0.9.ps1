<# 
.SYNOPSIS 
   vDiagram Visio Drawing Tool

.DESCRIPTION
   vDiagram Visio Drawing Tool

.NOTES 
   File Name	: vDiagram_2.0.9.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 2.0.9

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data
		MS Visio

.CHANGE LOG
	- 09/25/2019 - v2.0.9
		Moved from Get-<Item> to Get-View.
		Added Pop-up bubbles to all items in GUI to provide direction.
		
	- 07/17/2019 - v2.0.8
		Typo found out capture output. Added CpuHotRemoveEnabled, CpuHotAddEnabled & MemoryHotAddEnabled to VM & Template outputs.
		Added additional properties to VMHost object.
		
	- 04/15/2019 - v2.0.7
		New drawing added for Linked vCenters.
		
	- 04/06/2019 - v2.0.6
		New drawing added for VMs with snapshots.

	- 10/22/2018 - v2.0.5
		Dupliacte Resource Pools for same cluster were being drawn in Visio.
		
	- 10/22/2018 - v2.0.4
		Slight changes post presenting at Orlando VMUG UserCon
		Removed target vCenter box
		Cleaned up global variables for CSVs & vCenter
		File saves as .vsd then converts to .vsdx and deletes .vsd
		File save now in .vsdx vs .vsd as it saves as a smaller file
		Changed date format of Visio file from yyyy_MM_dd-HH_mm to yyyy-MM-dd_HH-mm
				
	- 10/17/2018 - v2.0.3
		Fixed IP and MAC address capture on VMHost and VMs, not listing all IPs and MACs
	
	- 10/02/2018 - v2.0.2
		Added Open CSV Folder Button to Capture Tab
		Once Open CSV Folder or OPen Visio Button is clicked form now resets
		Separated sections into regions for ease of modification later
	
	- 04/12/2018 - v2.0.1
		Added MAC Addresses to VMs & Templates
		Added a check to see if prior CSVs are still present
		Added option to copy prior CSVs to new folder
		Consolidate the object placement into functions for ease of management

	- 04/11/2018 - v2.0.0
		Presented as a Community Theater Session at South Florida VMUG
		Feature enhancement requests collected
#>

#region ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
#endregion ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Post-Constructor Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< About >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
$MyVer = "2.0.9"
$LastUpdated = "August  25, 2019"
$About = 
@"

	vDiagram $MyVer
	
	Contributors:	Tony Gonzalez
			Jason Hopkins
	
	Description:	vDiagram $MyVer - Based off of Alan Renouf's vDiagram
	
	Created:		February 13, 2018
	
	Last Updated:	$LastUpdated                   

"@
#endregion ~~< About >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TestShapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TestShapes = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes\vDiagram_" + $MyVer + ".vssx"
if (!(Test-Path $TestShapes))
{
	$CurrentLocation = Get-Location
	$UpdatedShapes = "$CurrentLocation" + "\vDiagram_" + "$MyVer" + ".vssx"
	copy $UpdatedShapes $TestShapes
	Write-Host "Copying Shapes File to My Shapes"
}
$shpFile = "\vDiagram_" + $MyVer + ".vssx"
#endregion ~~< TestShapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Set_WindowStyle >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Set_WindowStyle {
param(
    [Parameter()]
    [ValidateSet('FORCEMINIMIZE', 'HIDE', 'MAXIMIZE', 'MINIMIZE', 'RESTORE', 
                 'SHOW', 'SHOWDEFAULT', 'SHOWMAXIMIZED', 'SHOWMINIMIZED', 
                 'SHOWMINNOACTIVE', 'SHOWNA', 'SHOWNOACTIVATE', 'SHOWNORMAL')]
    $Style = 'SHOW',
    [Parameter()]
    $MainWindowHandle = (Get-Process -Id $pid).MainWindowHandle
)
    $WindowStates = @{
        FORCEMINIMIZE   = 11; HIDE            = 0
        MAXIMIZE        = 3;  MINIMIZE        = 6
        RESTORE         = 9;  SHOW            = 5
        SHOWDEFAULT     = 10; SHOWMAXIMIZED   = 3
        SHOWMINIMIZED   = 2;  SHOWMINNOACTIVE = 7
        SHOWNA          = 8;  SHOWNOACTIVATE  = 4
        SHOWNORMAL      = 1
    }
    Write-Verbose ("Set Window Style {1} on handle {0}" -f $MainWindowHandle, $($WindowStates[$style]))

    $Win32ShowWindowAsync = Add-Type –memberDefinition @” 
    [DllImport("user32.dll")] 
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru

    $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$Style]) | Out-Null
}
Set_WindowStyle MINIMIZE
#endregion ~~< Set_WindowStyle >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< About_Config >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function About_Config 
{

	$About

    # Add objects for About
    $AboutForm = New-Object System.Windows.Forms.Form
    $AboutTextBox = New-Object System.Windows.Forms.RichTextBox
    
    # About Form
    $AboutForm.Icon = $Icon
    $AboutForm.AutoScroll = $True
    $AboutForm.ClientSize = New-Object System.Drawing.Size(464,500)
    $AboutForm.DataBindings.DefaultDataSourceUpdateMode = 0
    $AboutForm.Name = "About"
    $AboutForm.StartPosition = 1
    $AboutForm.Text = "About vDiagram $MyVer"
    
    $AboutTextBox.Anchor = 15
    $AboutTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,240,240,240)
    $AboutTextBox.BorderStyle = 0
    $AboutTextBox.Font = "Tahoma"
    $AboutTextBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $AboutTextBox.Location = New-Object System.Drawing.Point(13,13)
    $AboutTextBox.Name = "AboutTextBox"
    $AboutTextBox.ReadOnly = $True
    $AboutTextBox.Size = New-Object System.Drawing.Size(440,500)
    $AboutTextBox.Text = $About
        
    $AboutForm.Controls.Add($AboutTextBox)

    $AboutForm.Show() | Out-Null
}
#endregion ~~< About_Config >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Post-Constructor Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region Form Creation
#~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vDiagram = New-Object System.Windows.Forms.Form
$vDiagram.ClientSize = New-Object System.Drawing.Size(1008, 661)
$CurrentLocation = Get-Location
$Icon = "$CurrentLocation" + "\vDiagram.ico"
$vDiagram.Icon = $Icon
$vDiagram.Text = "vDiagram " + $MyVer 
$vDiagram.BackColor = [System.Drawing.Color]::DarkCyan
#~~< SubTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab = New-Object System.Windows.Forms.TabControl
$SubTab.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$SubTab.Location = New-Object System.Drawing.Point(10, 136)
$SubTab.Size = New-Object System.Drawing.Size(990, 512)
$SubTab.TabIndex = 2
$SubTab.Text = "Draw Visio"
#~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDirections = New-Object System.Windows.Forms.TabPage
$TabDirections.Location = New-Object System.Drawing.Point(4, 22)
$TabDirections.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDirections.Size = New-Object System.Drawing.Size(982, 486)
$TabDirections.TabIndex = 0
$TabDirections.Text = "Directions"
$TabDirections.UseVisualStyleBackColor = $true
#~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawDirections = New-Object System.Windows.Forms.Label
$DrawDirections.Location = New-Object System.Drawing.Point(8, 288)
$DrawDirections.Size = New-Object System.Drawing.Size(900, 130)
$DrawDirections.TabIndex = 7
$DrawDirections.Text = "1. Click on "+[char]34+"Select Input Folder"+[char]34+" button and select location where CSVs can be found."+[char]13+[char]10+"2. Click on "+[char]34+"Check for CSVs"+[char]34+" button to validate presence of required files."+[char]13+[char]10+"3. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select where location where you would like to save the Visio drawing."+[char]13+[char]10+"4. Select drawing that you would like to produce."+[char]13+[char]10+"5. Click on "+[char]34+"Draw Visio"+[char]34+" button."+[char]13+[char]10+"6. Click on "+[char]34+"Open Visio Drawing"+[char]34+" button once "+[char]34+"Draw Visio"+[char]34+" button says it has completed."
#~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawHeading = New-Object System.Windows.Forms.Label
$DrawHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DrawHeading.Location = New-Object System.Drawing.Point(8, 264)
$DrawHeading.Size = New-Object System.Drawing.Size(149, 23)
$DrawHeading.TabIndex = 6
$DrawHeading.Text = "Draw Visio Tab"
#~~< CaptureDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureDirections = New-Object System.Windows.Forms.Label
$CaptureDirections.Location = New-Object System.Drawing.Point(8, 200)
$CaptureDirections.Size = New-Object System.Drawing.Size(900, 65)
$CaptureDirections.TabIndex = 5
$CaptureDirections.Text = "1. Click on "+[char]34+"Capture CSVs for Visio"+[char]34+" tab."+[char]13+[char]10+"2. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select folder where you would like to output the CSVs to."+[char]13+[char]10+"3. Select items you wish to grab data on."+[char]13+[char]10+"4. Click on "+[char]34+"Collect CSV Data"+[char]34+" button."
#~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvHeading = New-Object System.Windows.Forms.Label
$CaptureCsvHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CaptureCsvHeading.Location = New-Object System.Drawing.Point(8, 176)
$CaptureCsvHeading.Size = New-Object System.Drawing.Size(216, 23)
$CaptureCsvHeading.TabIndex = 4
$CaptureCsvHeading.Text = "Capture CSVs for Visio Tab"
#~~< vCenterInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfoDirections = New-Object System.Windows.Forms.Label
$vCenterInfoDirections.Location = New-Object System.Drawing.Point(8, 96)
$vCenterInfoDirections.Size = New-Object System.Drawing.Size(900, 70)
$vCenterInfoDirections.TabIndex = 3
$vCenterInfoDirections.Text = "1. Click on"+[char]34+"vCenter Info"+[char]34+" tab."+[char]13+[char]10+"2. Enter name of vCenter"+[char]13+[char]10+"3. Enter User Name and Password (password will be hashed and not plain text)."+[char]13+[char]10+"4. Click on "+[char]34+"Connect to vCenter"+[char]34+" button."
#~~< vCenterInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfoHeading = New-Object System.Windows.Forms.Label
$vCenterInfoHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$vCenterInfoHeading.Location = New-Object System.Drawing.Point(8, 72)
$vCenterInfoHeading.Size = New-Object System.Drawing.Size(149, 23)
$vCenterInfoHeading.TabIndex = 2
$vCenterInfoHeading.Text = "vCenter Info Tab"
#~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesDirections = New-Object System.Windows.Forms.Label
$PrerequisitesDirections.Location = New-Object System.Drawing.Point(8, 32)
$PrerequisitesDirections.Size = New-Object System.Drawing.Size(900, 30)
$PrerequisitesDirections.TabIndex = 1
$PrerequisitesDirections.Text = "1. Verify that prerequisites are met on the "+[char]34+"Prerequisites"+[char]34+" tab."+[char]34+[char]13+[char]10+"2. If not please install needed requirements."+[char]13+[char]10
#~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesHeading = New-Object System.Windows.Forms.Label
$PrerequisitesHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$PrerequisitesHeading.Location = New-Object System.Drawing.Point(8, 8)
$PrerequisitesHeading.Size = New-Object System.Drawing.Size(149, 23)
$PrerequisitesHeading.TabIndex = 0
$PrerequisitesHeading.Text = "Prerequisites Tab"
$TabDirections.Controls.Add($DrawDirections)
$TabDirections.Controls.Add($DrawHeading)
$TabDirections.Controls.Add($CaptureDirections)
$TabDirections.Controls.Add($CaptureCsvHeading)
$TabDirections.Controls.Add($vCenterInfoDirections)
$TabDirections.Controls.Add($vCenterInfoHeading)
$TabDirections.Controls.Add($PrerequisitesDirections)
$TabDirections.Controls.Add($PrerequisitesHeading)
#~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCapture = New-Object System.Windows.Forms.TabPage
$TabCapture.Location = New-Object System.Drawing.Point(4, 22)
$TabCapture.Padding = New-Object System.Windows.Forms.Padding(3)
$TabCapture.Size = New-Object System.Drawing.Size(982, 486)
$TabCapture.TabIndex = 1
$TabCapture.Text = "Capture CSVs for Visio"
$TabCapture.UseVisualStyleBackColor = $true
#~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton = New-Object System.Windows.Forms.Button
$OpenCaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenCaptureButton.Location = New-Object System.Drawing.Point(668, 215)
$OpenCaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenCaptureButton.TabIndex = 53
$OpenCaptureButton.Text = "Open CSV Output Folder"
$OpenCaptureButton.UseVisualStyleBackColor = $false
$OpenCaptureButton.BackColor = [System.Drawing.Color]::LightGray
#~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$components = New-Object System.ComponentModel.Container
#~~< OpenCaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$OpenCaptureButtonToolTip.AutoPopDelay = 5000
$OpenCaptureButtonToolTip.InitialDelay = 50
$OpenCaptureButtonToolTip.IsBalloon = $true
$OpenCaptureButtonToolTip.ReshowDelay = 100
$OpenCaptureButtonToolTip.SetToolTip($OpenCaptureButton, "Click once collection is complete to open output folder"+[char]13+[char]10+"seleted above.")
#~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton = New-Object System.Windows.Forms.Button
$CaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureButton.Location = New-Object System.Drawing.Point(448, 215)
$CaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureButton.TabIndex = 52
$CaptureButton.Text = "Collect CSV Data"
$CaptureButton.UseVisualStyleBackColor = $false
$CaptureButton.BackColor = [System.Drawing.Color]::LightGray
#~~< CaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureButtonToolTip.AutoPopDelay = 5000
$CaptureButtonToolTip.InitialDelay = 50
$CaptureButtonToolTip.IsBalloon = $true
$CaptureButtonToolTip.ReshowDelay = 100
$CaptureButtonToolTip.SetToolTip($CaptureButton, "Click to begin collecting environment information"+[char]13+[char]10+"on options selected above.")
#~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton = New-Object System.Windows.Forms.Button
$CaptureCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCheckButton.Location = New-Object System.Drawing.Point(228, 215)
$CaptureCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureCheckButton.TabIndex = 51
$CaptureCheckButton.Text = "Check All"
$CaptureCheckButton.UseVisualStyleBackColor = $false
$CaptureCheckButton.BackColor = [System.Drawing.Color]::LightGray
#~~< CaptureCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureCheckButtonToolTip.AutoPopDelay = 5000
$CaptureCheckButtonToolTip.InitialDelay = 50
$CaptureCheckButtonToolTip.IsBalloon = $true
$CaptureCheckButtonToolTip.ReshowDelay = 100
$CaptureCheckButtonToolTip.SetToolTip($CaptureCheckButton, "Click to check all check boxes above.")
#~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton = New-Object System.Windows.Forms.Button
$CaptureUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureUncheckButton.Location = New-Object System.Drawing.Point(8, 215)
$CaptureUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureUncheckButton.TabIndex = 50
$CaptureUncheckButton.Text = "Uncheck All"
$CaptureUncheckButton.UseVisualStyleBackColor = $false
$CaptureUncheckButton.BackColor = [System.Drawing.Color]::LightGray
#~~< CaptureUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureUncheckButtonToolTip.AutoPopDelay = 5000
$CaptureUncheckButtonToolTip.InitialDelay = 50
$CaptureUncheckButtonToolTip.IsBalloon = $true
$CaptureUncheckButtonToolTip.ReshowDelay = 100
$CaptureUncheckButtonToolTip.SetToolTip($CaptureUncheckButton, "Click to clear all check boxes above.")
#~~< LinkedvCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 180)
$LinkedvCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$LinkedvCenterCsvValidationComplete.TabIndex = 49
$LinkedvCenterCsvValidationComplete.Text = ""
#~~< LinkedvCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$LinkedvCenterCsvCheckBox.Checked = $true
$LinkedvCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$LinkedvCenterCsvCheckBox.Location = New-Object System.Drawing.Point(620, 180)
$LinkedvCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$LinkedvCenterCsvCheckBox.TabIndex = 48
$LinkedvCenterCsvCheckBox.Text = "Export Linked vCenter Info"
$LinkedvCenterCsvCheckBox.UseVisualStyleBackColor = $true
#~~< LinkedvCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$LinkedvCenterCsvToolTip.AutoPopDelay = 5000
$LinkedvCenterCsvToolTip.InitialDelay = 50
$LinkedvCenterCsvToolTip.IsBalloon = $true
$LinkedvCenterCsvToolTip.ReshowDelay = 100
$LinkedvCenterCsvToolTip.SetToolTip($LinkedvCenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Linked vCenters in this vCenter.")
#~~< SnapshotCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidationComplete = New-Object System.Windows.Forms.Label
$SnapshotCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 160)
$SnapshotCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$SnapshotCsvValidationComplete.TabIndex = 47
$SnapshotCsvValidationComplete.Text = ""
#~~< SnapshotCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$SnapshotCsvCheckBox.Checked = $true
$SnapshotCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SnapshotCsvCheckBox.Location = New-Object System.Drawing.Point(620, 160)
$SnapshotCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$SnapshotCsvCheckBox.TabIndex = 46
$SnapshotCsvCheckBox.Text = "Export Snapshot Info"
$SnapshotCsvCheckBox.UseVisualStyleBackColor = $true
#~~< SnapshotCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$SnapshotCsvToolTip.AutoPopDelay = 5000
$SnapshotCsvToolTip.InitialDelay = 50
$SnapshotCsvToolTip.IsBalloon = $true
$SnapshotCsvToolTip.ReshowDelay = 100
$SnapshotCsvToolTip.SetToolTip($SnapshotCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Snapshots in this vCenter.")
#~~< ResourcePoolCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidationComplete = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 140)
$ResourcePoolCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ResourcePoolCsvValidationComplete.TabIndex = 45
$ResourcePoolCsvValidationComplete.Text = ""
#~~< ResourcePoolCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ResourcePoolCsvCheckBox.Checked = $true
$ResourcePoolCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ResourcePoolCsvCheckBox.Location = New-Object System.Drawing.Point(620, 140)
$ResourcePoolCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ResourcePoolCsvCheckBox.TabIndex = 44
$ResourcePoolCsvCheckBox.Text = "Export Resource Pool Info"
$ResourcePoolCsvCheckBox.UseVisualStyleBackColor = $true
#~~< ResourcePoolCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ResourcePoolCsvToolTip.AutoPopDelay = 5000
$ResourcePoolCsvToolTip.InitialDelay = 50
$ResourcePoolCsvToolTip.IsBalloon = $true
$ResourcePoolCsvToolTip.ReshowDelay = 100
$ResourcePoolCsvToolTip.SetToolTip($ResourcePoolCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Resource Pools in this vCenter.")
#~~< DrsVmHostRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 120)
$DrsVmHostRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsVmHostRuleCsvValidationComplete.TabIndex = 43
$DrsVmHostRuleCsvValidationComplete.Text = ""
#~~< DrsVmHostRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsVmHostRuleCsvCheckBox.Checked = $true
$DrsVmHostRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsVmHostRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 120)
$DrsVmHostRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsVmHostRuleCsvCheckBox.TabIndex = 42
$DrsVmHostRuleCsvCheckBox.Text = "Export DRS VMHost Rule Info"
$DrsVmHostRuleCsvCheckBox.UseVisualStyleBackColor = $true
#~~< DrsVmHostRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsVmHostRuleCsvToolTip.AutoPopDelay = 5000
$DrsVmHostRuleCsvToolTip.InitialDelay = 50
$DrsVmHostRuleCsvToolTip.IsBalloon = $true
$DrsVmHostRuleCsvToolTip.ReshowDelay = 100
$DrsVmHostRuleCsvToolTip.SetToolTip($DrsVmHostRuleCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Host Rules"+[char]13+[char]10+"(DRS Host Rules) in this vCenter.")
#~~< DrsClusterGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 100)
$DrsClusterGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsClusterGroupCsvValidationComplete.TabIndex = 41
$DrsClusterGroupCsvValidationComplete.Text = ""
#~~< DrsClusterGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsClusterGroupCsvCheckBox.Checked = $true
$DrsClusterGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsClusterGroupCsvCheckBox.Location = New-Object System.Drawing.Point(620, 100)
$DrsClusterGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsClusterGroupCsvCheckBox.TabIndex = 40
$DrsClusterGroupCsvCheckBox.Text = "Export DRS Cluster Group Info"
$DrsClusterGroupCsvCheckBox.UseVisualStyleBackColor = $true
#~~< DrsClusterGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsClusterGroupCsvToolTip.AutoPopDelay = 5000
$DrsClusterGroupCsvToolTip.InitialDelay = 50
$DrsClusterGroupCsvToolTip.IsBalloon = $true
$DrsClusterGroupCsvToolTip.ReshowDelay = 100
$DrsClusterGroupCsvToolTip.SetToolTip($DrsClusterGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Cluster Rules"+[char]13+[char]10+"(DRS Cluster Rules) in this vCenter.")
#~~< DrsRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 80)
$DrsRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsRuleCsvValidationComplete.TabIndex = 39
$DrsRuleCsvValidationComplete.Text = ""
#~~< DrsRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsRuleCsvCheckBox.Checked = $true
$DrsRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 80)
$DrsRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsRuleCsvCheckBox.TabIndex = 38
$DrsRuleCsvCheckBox.Text = "Export DRS Rule Info"
$DrsRuleCsvCheckBox.UseVisualStyleBackColor = $true
#~~< DrsRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsRuleCsvToolTip.AutoPopDelay = 5000
$DrsRuleCsvToolTip.InitialDelay = 50
$DrsRuleCsvToolTip.IsBalloon = $true
$DrsRuleCsvToolTip.ReshowDelay = 100
$DrsRuleCsvToolTip.SetToolTip($DrsRuleCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Rules"+[char]13+[char]10+"(DRS Rules) in this vCenter.")
#~~< RdmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidationComplete = New-Object System.Windows.Forms.Label
$RdmCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 60)
$RdmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$RdmCsvValidationComplete.TabIndex = 37
$RdmCsvValidationComplete.Text = ""
#~~< RdmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$RdmCsvCheckBox.Checked = $true
$RdmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$RdmCsvCheckBox.Location = New-Object System.Drawing.Point(620, 60)
$RdmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$RdmCsvCheckBox.TabIndex = 36
$RdmCsvCheckBox.Text = "Export RDM Info"
$RdmCsvCheckBox.UseVisualStyleBackColor = $true
#~~< RdmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$RdmCsvToolTip.AutoPopDelay = 5000
$RdmCsvToolTip.InitialDelay = 50
$RdmCsvToolTip.IsBalloon = $true
$RdmCsvToolTip.ReshowDelay = 100
$RdmCsvToolTip.SetToolTip($RdmCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Raw Device Mappings (RDMs) in this vCenter.")
#~~< FolderCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidationComplete = New-Object System.Windows.Forms.Label
$FolderCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 40)
$FolderCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$FolderCsvValidationComplete.TabIndex = 35
$FolderCsvValidationComplete.Text = ""
#~~< FolderCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$FolderCsvCheckBox.Checked = $true
$FolderCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FolderCsvCheckBox.Location = New-Object System.Drawing.Point(620, 40)
$FolderCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$FolderCsvCheckBox.TabIndex = 34
$FolderCsvCheckBox.Text = "Export Folder Info"
$FolderCsvCheckBox.UseVisualStyleBackColor = $true
#~~< FolderCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$FolderCsvToolTip.AutoPopDelay = 5000
$FolderCsvToolTip.InitialDelay = 50
$FolderCsvToolTip.IsBalloon = $true
$FolderCsvToolTip.ReshowDelay = 100
$FolderCsvToolTip.SetToolTip($FolderCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Folders in this vCenter.")
#~~< VdsPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 180)
$VdsPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPnicCsvValidationComplete.TabIndex = 33
$VdsPnicCsvValidationComplete.Text = ""
#~~< VdsPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPnicCsvCheckBox.Checked = $true
$VdsPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 180)
$VdsPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPnicCsvCheckBox.TabIndex = 32
$VdsPnicCsvCheckBox.Text = "Export VDS pNIC Info"
$VdsPnicCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VdsPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsPnicCsvToolTip.AutoPopDelay = 5000
$VdsPnicCsvToolTip.InitialDelay = 50
$VdsPnicCsvToolTip.IsBalloon = $true
$VdsPnicCsvToolTip.ReshowDelay = 100
$VdsPnicCsvToolTip.SetToolTip($VdsPnicCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch Physical NICs in"+[char]13+[char]10+"this vCenter.")
#~~< VdsVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 160)
$VdsVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsVmkernelCsvValidationComplete.TabIndex = 31
$VdsVmkernelCsvValidationComplete.Text = ""
#~~< VdsVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsVmkernelCsvCheckBox.Checked = $true
$VdsVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 160)
$VdsVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsVmkernelCsvCheckBox.TabIndex = 30
$VdsVmkernelCsvCheckBox.Text = "Export VDS VMkernel Info"
$VdsVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VdsVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsVmkernelCsvToolTip.AutoPopDelay = 5000
$VdsVmkernelCsvToolTip.InitialDelay = 50
$VdsVmkernelCsvToolTip.IsBalloon = $true
$VdsVmkernelCsvToolTip.ReshowDelay = 100
$VdsVmkernelCsvToolTip.SetToolTip($VdsVmkernelCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch VMkernels in"+[char]13+[char]10+"this vCenter.")
#~~< VdsPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 140)
$VdsPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPortGroupCsvValidationComplete.TabIndex = 29
$VdsPortGroupCsvValidationComplete.Text = ""
#~~< VdsPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPortGroupCsvCheckBox.Checked = $true
$VdsPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 140)
$VdsPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPortGroupCsvCheckBox.TabIndex = 28
$VdsPortGroupCsvCheckBox.Text = "Export VDS Port Group Info"
$VdsPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VdsPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsPortGroupCsvToolTip.AutoPopDelay = 5000
$VdsPortGroupCsvToolTip.InitialDelay = 50
$VdsPortGroupCsvToolTip.IsBalloon = $true
$VdsPortGroupCsvToolTip.ReshowDelay = 100
$VdsPortGroupCsvToolTip.SetToolTip($VdsPortGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch Port Groups in"+[char]13+[char]10+"this vCenter.")
#~~< VdSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 120)
$VdSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdSwitchCsvValidationComplete.TabIndex = 27
$VdSwitchCsvValidationComplete.Text = ""
#~~< VdSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdSwitchCsvCheckBox.Checked = $true
$VdSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 120)
$VdSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdSwitchCsvCheckBox.TabIndex = 26
$VdSwitchCsvCheckBox.Text = "Export Distributed Switch Info"
$VdSwitchCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VdSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdSwitchCsvToolTip.AutoPopDelay = 5000
$VdSwitchCsvToolTip.InitialDelay = 50
$VdSwitchCsvToolTip.IsBalloon = $true
$VdSwitchCsvToolTip.ReshowDelay = 100
$VdSwitchCsvToolTip.SetToolTip($VdSwitchCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switches in this vCenter.")
#~~< VssPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 100)
$VssPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPnicCsvValidationComplete.TabIndex = 25
$VssPnicCsvValidationComplete.Text = ""
#~~< VssPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPnicCsvCheckBox.Checked = $true
$VssPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 100)
$VssPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPnicCsvCheckBox.TabIndex = 24
$VssPnicCsvCheckBox.Text = "Export VSS pNIC Info"
$VssPnicCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VssPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssPnicCsvToolTip.AutoPopDelay = 5000
$VssPnicCsvToolTip.InitialDelay = 50
$VssPnicCsvToolTip.IsBalloon = $true
$VssPnicCsvToolTip.ReshowDelay = 100
$VssPnicCsvToolTip.SetToolTip($VssPnicCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch Physical NICs in"+[char]13+[char]10+"this vCenter.")
#~~< VssVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 80)
$VssVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssVmkernelCsvValidationComplete.TabIndex = 23
$VssVmkernelCsvValidationComplete.Text = ""
#~~< VssVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssVmkernelCsvCheckBox.Checked = $true
$VssVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 80)
$VssVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssVmkernelCsvCheckBox.TabIndex = 22
$VssVmkernelCsvCheckBox.Text = "Export VSS VMkernel Info"
$VssVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VssVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssVmkernelCsvToolTip.AutoPopDelay = 5000
$VssVmkernelCsvToolTip.InitialDelay = 50
$VssVmkernelCsvToolTip.IsBalloon = $true
$VssVmkernelCsvToolTip.ReshowDelay = 100
$VssVmkernelCsvToolTip.SetToolTip($VssVmkernelCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch VMkernels in"+[char]13+[char]10+"this vCenter.")
#~~< VssPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 60)
$VssPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPortGroupCsvValidationComplete.TabIndex = 21
$VssPortGroupCsvValidationComplete.Text = ""
#~~< VssPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPortGroupCsvCheckBox.Checked = $true
$VssPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 60)
$VssPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPortGroupCsvCheckBox.TabIndex = 20
$VssPortGroupCsvCheckBox.Text = "Export VSS Port Group Info"
$VssPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VssPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssPortGroupCsvToolTip.AutoPopDelay = 5000
$VssPortGroupCsvToolTip.InitialDelay = 50
$VssPortGroupCsvToolTip.IsBalloon = $true
$VssPortGroupCsvToolTip.ReshowDelay = 100
$VssPortGroupCsvToolTip.SetToolTip($VssPortGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch Port Groups in"+[char]13+[char]10+"this vCenter.")
#~~< VsSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 40)
$VsSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VsSwitchCsvValidationComplete.TabIndex = 19
$VsSwitchCsvValidationComplete.Text = ""
#~~< VsSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VsSwitchCsvCheckBox.Checked = $true
$VsSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VsSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 40)
$VsSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VsSwitchCsvCheckBox.TabIndex = 18
$VsSwitchCsvCheckBox.Text = "Export Standard Switch Info"
$VsSwitchCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VsSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VsSwitchCsvToolTip.AutoPopDelay = 5000
$VsSwitchCsvToolTip.InitialDelay = 50
$VsSwitchCsvToolTip.IsBalloon = $true
$VsSwitchCsvToolTip.ReshowDelay = 100
$VsSwitchCsvToolTip.SetToolTip($VsSwitchCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switches in this vCenter.")
#~~< DatastoreCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 180)
$DatastoreCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreCsvValidationComplete.TabIndex = 17
$DatastoreCsvValidationComplete.Text = ""
#~~< DatastoreCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreCsvCheckBox.Checked = $true
$DatastoreCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreCsvCheckBox.Location = New-Object System.Drawing.Point(10, 180)
$DatastoreCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreCsvCheckBox.TabIndex = 16
$DatastoreCsvCheckBox.Text = "Export Datastore Info"
$DatastoreCsvCheckBox.UseVisualStyleBackColor = $true
#~~< DatacenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatacenterCsvToolTip.AutoPopDelay = 5000
$DatacenterCsvToolTip.InitialDelay = 50
$DatacenterCsvToolTip.IsBalloon = $true
$DatacenterCsvToolTip.ReshowDelay = 100
$DatacenterCsvToolTip.SetToolTip($DatastoreCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datastores in this vCenter.")
#~~< DatastoreClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 160)
$DatastoreClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreClusterCsvValidationComplete.TabIndex = 15
$DatastoreClusterCsvValidationComplete.Text = ""
#~~< DatastoreClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreClusterCsvCheckBox.Checked = $true
$DatastoreClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 160)
$DatastoreClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreClusterCsvCheckBox.TabIndex = 14
$DatastoreClusterCsvCheckBox.Text = "Export Datastore Cluster Info"
$DatastoreClusterCsvCheckBox.UseVisualStyleBackColor = $true
#~~< DatastoreClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatastoreClusterCsvToolTip.AutoPopDelay = 5000
$DatastoreClusterCsvToolTip.InitialDelay = 50
$DatastoreClusterCsvToolTip.IsBalloon = $true
$DatastoreClusterCsvToolTip.ReshowDelay = 100
$DatastoreClusterCsvToolTip.SetToolTip($DatastoreClusterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datastore Clusters in this vCenter.")
#~~< TemplateCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidationComplete = New-Object System.Windows.Forms.Label
$TemplateCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 140)
$TemplateCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$TemplateCsvValidationComplete.TabIndex = 13
$TemplateCsvValidationComplete.Text = ""
#~~< TemplateCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$TemplateCsvCheckBox.Checked = $true
$TemplateCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$TemplateCsvCheckBox.Location = New-Object System.Drawing.Point(10, 140)
$TemplateCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$TemplateCsvCheckBox.TabIndex = 12
$TemplateCsvCheckBox.Text = "Export Template Info"
$TemplateCsvCheckBox.UseVisualStyleBackColor = $true
#~~< TemplateCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$TemplateCsvToolTip.AutoPopDelay = 5000
$TemplateCsvToolTip.InitialDelay = 50
$TemplateCsvToolTip.IsBalloon = $true
$TemplateCsvToolTip.ReshowDelay = 100
$TemplateCsvToolTip.SetToolTip($TemplateCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Templates in this vCenter.")
#~~< VmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 120)
$VmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmCsvValidationComplete.TabIndex = 11
$VmCsvValidationComplete.Text = ""
#~~< VmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmCsvCheckBox.Checked = $true
$VmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmCsvCheckBox.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmCsvCheckBox.TabIndex = 10
$VmCsvCheckBox.Text = "Export Virtual Machine Info"
$VmCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VmCsvToolTip.AutoPopDelay = 5000
$VmCsvToolTip.InitialDelay = 50
$VmCsvToolTip.IsBalloon = $true
$VmCsvToolTip.ReshowDelay = 100
$VmCsvToolTip.SetToolTip($VmCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Machines in this vCenter.")
#~~< VmHostCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmHostCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 100)
$VmHostCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationComplete.TabIndex = 9
$VmHostCsvValidationComplete.Text = ""
#~~< VmHostCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmHostCsvCheckBox.Checked = $true
$VmHostCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmHostCsvCheckBox.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmHostCsvCheckBox.TabIndex = 8
$VmHostCsvCheckBox.Text = "Export Host Info"
$VmHostCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VmHostCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VmHostCsvToolTip.AutoPopDelay = 5000
$VmHostCsvToolTip.InitialDelay = 50
$VmHostCsvToolTip.IsBalloon = $true
$VmHostCsvToolTip.ReshowDelay = 100
$VmHostCsvToolTip.SetToolTip($VmHostCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all ESXi Hosts in this vCenter.")
#~~< ClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$ClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 80)
$ClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ClusterCsvValidationComplete.TabIndex = 7
$ClusterCsvValidationComplete.Text = ""
#~~< ClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ClusterCsvCheckBox.Checked = $true
$ClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 80)
$ClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ClusterCsvCheckBox.TabIndex = 6
$ClusterCsvCheckBox.Text = "Export Cluster Info"
$ClusterCsvCheckBox.UseVisualStyleBackColor = $true
#~~< ClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ClusterCsvToolTip.AutoPopDelay = 5000
$ClusterCsvToolTip.InitialDelay = 50
$ClusterCsvToolTip.IsBalloon = $true
$ClusterCsvToolTip.ReshowDelay = 100
$ClusterCsvToolTip.SetToolTip($ClusterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Clusters in this vCenter.")
#~~< DatacenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatacenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 60)
$DatacenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatacenterCsvValidationComplete.TabIndex = 5
$DatacenterCsvValidationComplete.Text = ""
#~~< DatacenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatacenterCsvCheckBox.Checked = $true
$DatacenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatacenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 60)
$DatacenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatacenterCsvCheckBox.TabIndex = 4
$DatacenterCsvCheckBox.Text = "Export Datacenter Info"
$DatacenterCsvCheckBox.UseVisualStyleBackColor = $true
$DatacenterCsvToolTip.SetToolTip($DatacenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datacenters in this vCenter.")
#~~< vCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$vCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 40)
$vCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$vCenterCsvValidationComplete.TabIndex = 3
$vCenterCsvValidationComplete.Text = ""
#~~< vCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenterCsvCheckBox.Checked = $true
$vCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$vCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$vCenterCsvCheckBox.TabIndex = 2
$vCenterCsvCheckBox.Text = "Export vCenter Info"
$vCenterCsvCheckBox.UseVisualStyleBackColor = $true
#~~< VcenterToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VcenterToolTip.AutoPopDelay = 5000
$VcenterToolTip.InitialDelay = 50
$VcenterToolTip.IsBalloon = $true
$VcenterToolTip.ReshowDelay = 100
$VcenterToolTip.SetToolTip($vCenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"vCenter.")
#~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton = New-Object System.Windows.Forms.Button
$CaptureCsvOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCsvOutputButton.Location = New-Object System.Drawing.Point(220, 10)
$CaptureCsvOutputButton.Size = New-Object System.Drawing.Size(750, 25)
$CaptureCsvOutputButton.TabIndex = 1
$CaptureCsvOutputButton.Text = "Select Output Folder"
$CaptureCsvOutputButton.UseVisualStyleBackColor = $false
$CaptureCsvOutputButton.BackColor = [System.Drawing.Color]::LightGray
#~~< CaptureCsvOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureCsvOutputButtonToolTip.AutoPopDelay = 5000
$CaptureCsvOutputButtonToolTip.InitialDelay = 50
$CaptureCsvOutputButtonToolTip.IsBalloon = $true
$CaptureCsvOutputButtonToolTip.ReshowDelay = 100
$CaptureCsvOutputButtonToolTip.SetToolTip($CaptureCsvOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"CSV"+[char]39+"s."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green."+[char]13+[char]10+[char]13+[char]10+"If the folder has files in it you will be presented with an "+[char]13+[char]10+"option to move or delete the files that are currently there.")
#~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputLabel = New-Object System.Windows.Forms.Label
$CaptureCsvOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CaptureCsvOutputLabel.Location = New-Object System.Drawing.Point(10, 10)
$CaptureCsvOutputLabel.Size = New-Object System.Drawing.Size(210, 25)
$CaptureCsvOutputLabel.TabIndex = 0
$CaptureCsvOutputLabel.Text = "CSV Output Folder:"
$TabCapture.Controls.Add($OpenCaptureButton)
$TabCapture.Controls.Add($CaptureButton)
$TabCapture.Controls.Add($CaptureCheckButton)
$TabCapture.Controls.Add($CaptureUncheckButton)
$TabCapture.Controls.Add($LinkedvCenterCsvValidationComplete)
$TabCapture.Controls.Add($LinkedvCenterCsvCheckBox)
$TabCapture.Controls.Add($SnapshotCsvValidationComplete)
$TabCapture.Controls.Add($SnapshotCsvCheckBox)
$TabCapture.Controls.Add($ResourcePoolCsvValidationComplete)
$TabCapture.Controls.Add($ResourcePoolCsvCheckBox)
$TabCapture.Controls.Add($DrsVmHostRuleCsvValidationComplete)
$TabCapture.Controls.Add($DrsVmHostRuleCsvCheckBox)
$TabCapture.Controls.Add($DrsClusterGroupCsvValidationComplete)
$TabCapture.Controls.Add($DrsClusterGroupCsvCheckBox)
$TabCapture.Controls.Add($DrsRuleCsvValidationComplete)
$TabCapture.Controls.Add($DrsRuleCsvCheckBox)
$TabCapture.Controls.Add($RdmCsvValidationComplete)
$TabCapture.Controls.Add($RdmCsvCheckBox)
$TabCapture.Controls.Add($FolderCsvValidationComplete)
$TabCapture.Controls.Add($FolderCsvCheckBox)
$TabCapture.Controls.Add($VdsPnicCsvValidationComplete)
$TabCapture.Controls.Add($VdsPnicCsvCheckBox)
$TabCapture.Controls.Add($VdsVmkernelCsvValidationComplete)
$TabCapture.Controls.Add($VdsVmkernelCsvCheckBox)
$TabCapture.Controls.Add($VdsPortGroupCsvValidationComplete)
$TabCapture.Controls.Add($VdsPortGroupCsvCheckBox)
$TabCapture.Controls.Add($VdSwitchCsvValidationComplete)
$TabCapture.Controls.Add($VdSwitchCsvCheckBox)
$TabCapture.Controls.Add($VssPnicCsvValidationComplete)
$TabCapture.Controls.Add($VssPnicCsvCheckBox)
$TabCapture.Controls.Add($VssVmkernelCsvValidationComplete)
$TabCapture.Controls.Add($VssVmkernelCsvCheckBox)
$TabCapture.Controls.Add($VssPortGroupCsvValidationComplete)
$TabCapture.Controls.Add($VssPortGroupCsvCheckBox)
$TabCapture.Controls.Add($VsSwitchCsvValidationComplete)
$TabCapture.Controls.Add($VsSwitchCsvCheckBox)
$TabCapture.Controls.Add($DatastoreCsvValidationComplete)
$TabCapture.Controls.Add($DatastoreCsvCheckBox)
$TabCapture.Controls.Add($DatastoreClusterCsvValidationComplete)
$TabCapture.Controls.Add($DatastoreClusterCsvCheckBox)
$TabCapture.Controls.Add($TemplateCsvValidationComplete)
$TabCapture.Controls.Add($TemplateCsvCheckBox)
$TabCapture.Controls.Add($VmCsvValidationComplete)
$TabCapture.Controls.Add($VmCsvCheckBox)
$TabCapture.Controls.Add($VmHostCsvValidationComplete)
$TabCapture.Controls.Add($VmHostCsvCheckBox)
$TabCapture.Controls.Add($ClusterCsvValidationComplete)
$TabCapture.Controls.Add($ClusterCsvCheckBox)
$TabCapture.Controls.Add($DatacenterCsvValidationComplete)
$TabCapture.Controls.Add($DatacenterCsvCheckBox)
$TabCapture.Controls.Add($vCenterCsvValidationComplete)
$TabCapture.Controls.Add($vCenterCsvCheckBox)
$TabCapture.Controls.Add($CaptureCsvOutputButton)
$TabCapture.Controls.Add($CaptureCsvOutputLabel)
#~~< TabCaptureToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCaptureToolTip = New-Object System.Windows.Forms.ToolTip($components)
$TabCaptureToolTip.AutoPopDelay = 5000
$TabCaptureToolTip.InitialDelay = 50
$TabCaptureToolTip.IsBalloon = $true
$TabCaptureToolTip.ReshowDelay = 100
$TabCaptureToolTip.SetToolTip($TabCapture, "This must be ran first in order to collect the information"+[char]13+[char]10+"about your environment.")
#~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDraw = New-Object System.Windows.Forms.TabPage
$TabDraw.Location = New-Object System.Drawing.Point(4, 22)
$TabDraw.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDraw.Size = New-Object System.Drawing.Size(982, 486)
$TabDraw.TabIndex = 2
$TabDraw.Text = "Draw Visio"
$TabDraw.UseVisualStyleBackColor = $true
#~~< vCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$vCenterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 40)
$vCenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$vCenterCsvValidationCheck.TabIndex = 3
$vCenterCsvValidationCheck.Text = ""
$vCenterCsvValidationCheck.add_Click({VCenterCsvValidationCheckClick($vCenterCsvValidationCheck)})
#~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton = New-Object System.Windows.Forms.Button
$OpenVisioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenVisioButton.Location = New-Object System.Drawing.Point(668, 450)
$OpenVisioButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenVisioButton.TabIndex = 90
$OpenVisioButton.Text = "Open Visio Drawing"
$OpenVisioButton.UseVisualStyleBackColor = $false
$OpenVisioButton.BackColor = [System.Drawing.Color]::LightGray
#~~< OpenVisioButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$OpenVisioButtonToolTip.AutoPopDelay = 5000
$OpenVisioButtonToolTip.InitialDelay = 50
$OpenVisioButtonToolTip.IsBalloon = $true
$OpenVisioButtonToolTip.ReshowDelay = 100
$OpenVisioButtonToolTip.SetToolTip($OpenVisioButton, "Click to open Visio drawing once all above check boxes"+[char]13+[char]10+"are marked as completed.")
#~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton = New-Object System.Windows.Forms.Button
$DrawButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawButton.Location = New-Object System.Drawing.Point(448, 450)
$DrawButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawButton.TabIndex = 89
$DrawButton.Text = "Draw Visio"
$DrawButton.UseVisualStyleBackColor = $false
$DrawButton.BackColor = [System.Drawing.Color]::LightGray
#~~< DrawButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawButtonToolTip.AutoPopDelay = 5000
$DrawButtonToolTip.InitialDelay = 50
$DrawButtonToolTip.IsBalloon = $true
$DrawButtonToolTip.ReshowDelay = 100
$DrawButtonToolTip.SetToolTip($DrawButton, "Click to begin drawing environment based on"+[char]13+[char]10+"options selected above.")
#~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton = New-Object System.Windows.Forms.Button
$DrawCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCheckButton.Location = New-Object System.Drawing.Point(228, 450)
$DrawCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawCheckButton.TabIndex = 88
$DrawCheckButton.Text = "Check All"
$DrawCheckButton.UseVisualStyleBackColor = $false
$DrawCheckButton.BackColor = [System.Drawing.Color]::LightGray
#~~< DrawCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawCheckButtonToolTip.AutoPopDelay = 5000
$DrawCheckButtonToolTip.InitialDelay = 50
$DrawCheckButtonToolTip.IsBalloon = $true
$DrawCheckButtonToolTip.ReshowDelay = 100
$DrawCheckButtonToolTip.SetToolTip($DrawCheckButton, "Click to check all check boxes above.")
#~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton = New-Object System.Windows.Forms.Button
$DrawUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawUncheckButton.Location = New-Object System.Drawing.Point(8, 450)
$DrawUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawUncheckButton.TabIndex = 87
$DrawUncheckButton.Text = "Uncheck All"
$DrawUncheckButton.UseVisualStyleBackColor = $false
$DrawUncheckButton.BackColor = [System.Drawing.Color]::LightGray
#~~< DrawUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawUncheckButtonToolTip.AutoPopDelay = 5000
$DrawUncheckButtonToolTip.InitialDelay = 50
$DrawUncheckButtonToolTip.IsBalloon = $true
$DrawUncheckButtonToolTip.ReshowDelay = 100
$DrawUncheckButtonToolTip.SetToolTip($DrawUncheckButton, "Click to clear all check boxes above.")
#~~< Cluster_to_DRS_Rule_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_Complete = New-Object System.Windows.Forms.Label
$Cluster_to_DRS_Rule_Complete.Location = New-Object System.Drawing.Point(760, 400)
$Cluster_to_DRS_Rule_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Cluster_to_DRS_Rule_Complete.TabIndex = 86
$Cluster_to_DRS_Rule_Complete.Text = ""
#~~< Cluster_to_DRS_Rule_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Cluster_to_DRS_Rule_DrawCheckBox.Checked = $true
$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Cluster_to_DRS_Rule_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 400)
$Cluster_to_DRS_Rule_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$Cluster_to_DRS_Rule_DrawCheckBox.TabIndex = 85
$Cluster_to_DRS_Rule_DrawCheckBox.Text = "Cluster to DRS Rule Visio Drawing"
$Cluster_to_DRS_Rule_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< Cluster_to_DRS_Rule_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.InitialDelay = 50
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.IsBalloon = $true
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.ReshowDelay = 100
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.SetToolTip($Cluster_to_DRS_Rule_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to DRS Rules."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#~~< VDSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VDSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(760, 380)
$VDSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VDSPortGroup_to_VM_Complete.TabIndex = 84
$VDSPortGroup_to_VM_Complete.Text = ""
#~~< VDSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VDSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 380)
$VDSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDSPortGroup_to_VM_DrawCheckBox.TabIndex = 83
$VDSPortGroup_to_VM_DrawCheckBox.Text = "Distributed Switch Port Group to VM Visio Drawing"
$VDSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VDSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.SetToolTip($VDSPortGroup_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to Port Groups to VMs."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#~~< VMK_to_VDS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VDS_Complete.Location = New-Object System.Drawing.Point(760, 360)
$VMK_to_VDS_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMK_to_VDS_Complete.TabIndex = 82
$VMK_to_VDS_Complete.Text = ""
#~~< VMK_to_VDS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VDS_DrawCheckBox.Checked = $true
$VMK_to_VDS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VDS_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 360)
$VMK_to_VDS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VDS_DrawCheckBox.TabIndex = 81
$VMK_to_VDS_DrawCheckBox.Text = "VMkernel to Distributed Switch Visio Drawing"
$VMK_to_VDS_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VMK_to_VDS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMK_to_VDS_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMK_to_VDS_DrawCheckBoxToolTip.InitialDelay = 50
$VMK_to_VDS_DrawCheckBoxToolTip.IsBalloon = $true
$VMK_to_VDS_DrawCheckBoxToolTip.ReshowDelay = 100
$VMK_to_VDS_DrawCheckBoxToolTip.SetToolTip($VMK_to_VDS_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to VMkernels. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#~~< VDS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VDS_to_Host_Complete.Location = New-Object System.Drawing.Point(760, 340)
$VDS_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VDS_to_Host_Complete.TabIndex = 80
$VDS_to_Host_Complete.Text = ""
#~~< VDS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDS_to_Host_DrawCheckBox.Checked = $true
$VDS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 340)
$VDS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDS_to_Host_DrawCheckBox.TabIndex = 79
$VDS_to_Host_DrawCheckBox.Text = "Distributed Switch to Host Visio Drawing"
$VDS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VDS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VDS_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VDS_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VDS_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VDS_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VDS_to_Host_DrawCheckBoxToolTip.SetToolTip($VDS_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to Port Groups. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#~~< VSSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VSSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(760, 320)
$VSSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VSSPortGroup_to_VM_Complete.TabIndex = 78
$VSSPortGroup_to_VM_Complete.Text = ""
#~~< VSSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VSSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 320)
$VSSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSSPortGroup_to_VM_DrawCheckBox.TabIndex = 77
$VSSPortGroup_to_VM_DrawCheckBox.Text = "Standard Switch Port Group to VM Visio Drawing"
$VSSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VSSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.SetToolTip($VSSPortGroup_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Port Groups to VMs."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#~~< VMK_to_VSS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VSS_Complete.Location = New-Object System.Drawing.Point(760, 300)
$VMK_to_VSS_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMK_to_VSS_Complete.TabIndex = 76
$VMK_to_VSS_Complete.Text = ""
#~~< VMK_to_VSS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VSS_DrawCheckBox.Checked = $true
$VMK_to_VSS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VSS_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 300)
$VMK_to_VSS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VSS_DrawCheckBox.TabIndex = 75
$VMK_to_VSS_DrawCheckBox.Text = "VMkernel to Standard Switch Visio Drawing"
$VMK_to_VSS_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VMK_to_VSS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMK_to_VSS_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMK_to_VSS_DrawCheckBoxToolTip.InitialDelay = 50
$VMK_to_VSS_DrawCheckBoxToolTip.IsBalloon = $true
$VMK_to_VSS_DrawCheckBoxToolTip.ReshowDelay = 100
$VMK_to_VSS_DrawCheckBoxToolTip.SetToolTip($VMK_to_VSS_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to VMkernels. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#~~< VSS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VSS_to_Host_Complete.Location = New-Object System.Drawing.Point(760, 280)
$VSS_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VSS_to_Host_Complete.TabIndex = 74
$VSS_to_Host_Complete.Text = ""
#~~< VSS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSS_to_Host_DrawCheckBox.Checked = $true
$VSS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 280)
$VSS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSS_to_Host_DrawCheckBox.TabIndex = 73
$VSS_to_Host_DrawCheckBox.Text = "Standard Switch to Host Visio Drawing"
$VSS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VSS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VSS_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VSS_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VSS_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VSS_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VSS_to_Host_DrawCheckBoxToolTip.SetToolTip($VSS_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Port Groups. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#~~< PhysicalNIC_to_vSwitch_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_Complete = New-Object System.Windows.Forms.Label
$PhysicalNIC_to_vSwitch_Complete.Location = New-Object System.Drawing.Point(760, 260)
$PhysicalNIC_to_vSwitch_Complete.Size = New-Object System.Drawing.Size(90, 20)
$PhysicalNIC_to_vSwitch_Complete.TabIndex = 72
$PhysicalNIC_to_vSwitch_Complete.Text = ""
#~~< PhysicalNIC_to_vSwitch_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$PhysicalNIC_to_vSwitch_DrawCheckBox.Checked = $true
$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$PhysicalNIC_to_vSwitch_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 260)
$PhysicalNIC_to_vSwitch_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$PhysicalNIC_to_vSwitch_DrawCheckBox.TabIndex = 71
$PhysicalNIC_to_vSwitch_DrawCheckBox.Text = "PhysicalNIC to vSwitch Visio Drawing"
$PhysicalNIC_to_vSwitch_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.AutoPopDelay = 5000
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.InitialDelay = 50
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.IsBalloon = $true
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.ReshowDelay = 100
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.SetToolTip($PhysicalNIC_to_vSwitch_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Physical NIC. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#~~< Snapshot_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_Complete = New-Object System.Windows.Forms.Label
$Snapshot_to_VM_Complete.Location = New-Object System.Drawing.Point(315, 420)
$Snapshot_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Snapshot_to_VM_Complete.TabIndex = 70
$Snapshot_to_VM_Complete.Text = ""
#~~< Snapshot_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Snapshot_to_VM_DrawCheckBox.Checked = $true
$Snapshot_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Snapshot_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 420)
$Snapshot_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Snapshot_to_VM_DrawCheckBox.TabIndex = 69
$Snapshot_to_VM_DrawCheckBox.Text = "Snapshot to VM Visio Drawing"
$Snapshot_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< Snapshot_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Snapshot_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Snapshot_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$Snapshot_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$Snapshot_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$Snapshot_to_VM_DrawCheckBoxToolTip.SetToolTip($Snapshot_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Virtual Machines"+[char]13+[char]10+"to Snapshot Tree. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#~~< Datastore_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_Complete = New-Object System.Windows.Forms.Label
$Datastore_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 400)
$Datastore_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Datastore_to_Host_Complete.TabIndex = 68
$Datastore_to_Host_Complete.Text = ""
#~~< Datastore_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Datastore_to_Host_DrawCheckBox.Checked = $true
$Datastore_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Datastore_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 400)
$Datastore_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Datastore_to_Host_DrawCheckBox.TabIndex = 67
$Datastore_to_Host_DrawCheckBox.Text = "Datastore to Host Visio Drawing"
$Datastore_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< Datastore_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Datastore_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Datastore_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$Datastore_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$Datastore_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$Datastore_to_Host_DrawCheckBoxToolTip.SetToolTip($Datastore_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Datastores. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#~~< VM_to_ResourcePool_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_Complete = New-Object System.Windows.Forms.Label
$VM_to_ResourcePool_Complete.Location = New-Object System.Drawing.Point(315, 380)
$VM_to_ResourcePool_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_ResourcePool_Complete.TabIndex = 66
$VM_to_ResourcePool_Complete.Text = ""
#~~< VM_to_ResourcePool_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_ResourcePool_DrawCheckBox.Checked = $true
$VM_to_ResourcePool_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_ResourcePool_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 380)
$VM_to_ResourcePool_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_ResourcePool_DrawCheckBox.TabIndex = 65
$VM_to_ResourcePool_DrawCheckBox.Text = "VM to ResourcePool Visio Drawing"
$VM_to_ResourcePool_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VM_to_ResourcePool_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_ResourcePool_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_ResourcePool_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_ResourcePool_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_ResourcePool_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_ResourcePool_DrawCheckBoxToolTip.SetToolTip($VM_to_ResourcePool_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Resource Pools  to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#~~< VM_to_Datastore_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_Complete = New-Object System.Windows.Forms.Label
$VM_to_Datastore_Complete.Location = New-Object System.Drawing.Point(315, 360)
$VM_to_Datastore_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Datastore_Complete.TabIndex = 64
$VM_to_Datastore_Complete.Text = ""
#~~< VM_to_Datastore_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Datastore_DrawCheckBox.Checked = $true
$VM_to_Datastore_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Datastore_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 360)
$VM_to_Datastore_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Datastore_DrawCheckBox.TabIndex = 63
$VM_to_Datastore_DrawCheckBox.Text = "VM to Datastore Visio Drawing"
$VM_to_Datastore_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VM_to_Datastore_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Datastore_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Datastore_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Datastore_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Datastore_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Datastore_DrawCheckBoxToolTip.SetToolTip($VM_to_Datastore_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Datastore Clusters"+[char]13+[char]10+"to Datastores to Virtual Machines. This will also add all"+[char]13+[char]10+"metadata to the Visio shapes.")
#~~< SRM_Protected_VMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_Complete = New-Object System.Windows.Forms.Label
$SRM_Protected_VMs_Complete.Location = New-Object System.Drawing.Point(315, 340)
$SRM_Protected_VMs_Complete.Size = New-Object System.Drawing.Size(90, 20)
$SRM_Protected_VMs_Complete.TabIndex = 62
$SRM_Protected_VMs_Complete.Text = ""
#~~< SRM_Protected_VMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$SRM_Protected_VMs_DrawCheckBox.Checked = $true
$SRM_Protected_VMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SRM_Protected_VMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 340)
$SRM_Protected_VMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$SRM_Protected_VMs_DrawCheckBox.TabIndex = 61
$SRM_Protected_VMs_DrawCheckBox.Text = "SRM Protected VMs Visio Drawing"
$SRM_Protected_VMs_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< SRM_Protected_VMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$SRM_Protected_VMs_DrawCheckBoxToolTip.AutoPopDelay = 5000
$SRM_Protected_VMs_DrawCheckBoxToolTip.InitialDelay = 50
$SRM_Protected_VMs_DrawCheckBoxToolTip.IsBalloon = $true
$SRM_Protected_VMs_DrawCheckBoxToolTip.ReshowDelay = 100
$SRM_Protected_VMs_DrawCheckBoxToolTip.SetToolTip($SRM_Protected_VMs_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#~~< VMs_with_RDMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_Complete = New-Object System.Windows.Forms.Label
$VMs_with_RDMs_Complete.Location = New-Object System.Drawing.Point(315, 320)
$VMs_with_RDMs_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMs_with_RDMs_Complete.TabIndex = 60
$VMs_with_RDMs_Complete.Text = ""
#~~< VMs_with_RDMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMs_with_RDMs_DrawCheckBox.Checked = $true
$VMs_with_RDMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMs_with_RDMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 320)
$VMs_with_RDMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VMs_with_RDMs_DrawCheckBox.TabIndex = 59
$VMs_with_RDMs_DrawCheckBox.Text = "VMs with RDMs Visio Drawing"
$VMs_with_RDMs_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VMs_with_RDMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMs_with_RDMs_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMs_with_RDMs_DrawCheckBoxToolTip.InitialDelay = 50
$VMs_with_RDMs_DrawCheckBoxToolTip.IsBalloon = $true
$VMs_with_RDMs_DrawCheckBoxToolTip.ReshowDelay = 100
$VMs_with_RDMs_DrawCheckBoxToolTip.SetToolTip($VMs_with_RDMs_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Virtual Machines"+[char]13+[char]10+"to Raw Device Mappings (RDMs). This will also add all"+[char]13+[char]10+"metadata to the Visio shapes.")
#~~< VM_to_Folder_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_Complete = New-Object System.Windows.Forms.Label
$VM_to_Folder_Complete.Location = New-Object System.Drawing.Point(315, 300)
$VM_to_Folder_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Folder_Complete.TabIndex = 58
$VM_to_Folder_Complete.Text = ""
#~~< VM_to_Folder_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Folder_DrawCheckBox.Checked = $true
$VM_to_Folder_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Folder_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 300)
$VM_to_Folder_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Folder_DrawCheckBox.TabIndex = 57
$VM_to_Folder_DrawCheckBox.Text = "VM to Folder Visio Drawing"
$VM_to_Folder_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VM_to_Folder_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Folder_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Folder_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Folder_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Folder_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Folder_DrawCheckBoxToolTip.SetToolTip($VM_to_Folder_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Folders to Virtual Machines."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#~~< VM_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_Complete = New-Object System.Windows.Forms.Label
$VM_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 280)
$VM_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Host_Complete.TabIndex = 56
$VM_to_Host_Complete.Text = ""
#~~< VM_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Host_DrawCheckBox.Checked = $true
$VM_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 280)
$VM_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Host_DrawCheckBox.TabIndex = 55
$VM_to_Host_DrawCheckBox.Text = "VM to Host Visio Drawing"
$VM_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< VM_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Host_DrawCheckBoxToolTip.SetToolTip($VM_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes."+[char]13+[char]10)
#~~< vCenter_to_LinkedvCenter_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_Complete = New-Object System.Windows.Forms.Label
$vCenter_to_LinkedvCenter_Complete.Location = New-Object System.Drawing.Point(315, 260)
$vCenter_to_LinkedvCenter_Complete.Size = New-Object System.Drawing.Size(90, 20)
$vCenter_to_LinkedvCenter_Complete.TabIndex = 54
$vCenter_to_LinkedvCenter_Complete.Text = ""
#~~< vCenter_to_LinkedvCenter_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenter_to_LinkedvCenter_DrawCheckBox.Checked = $true
$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenter_to_LinkedvCenter_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 260)
$vCenter_to_LinkedvCenter_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$vCenter_to_LinkedvCenter_DrawCheckBox.TabIndex = 53
$vCenter_to_LinkedvCenter_DrawCheckBox.Text = "vCenter to Linked vCenter Visio Drawing"
$vCenter_to_LinkedvCenter_DrawCheckBox.UseVisualStyleBackColor = $true
#~~< vCenter_to_LinkedvCenter_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.AutoPopDelay = 5000
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.InitialDelay = 50
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.IsBalloon = $true
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.ReshowDelay = 100
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.SetToolTip($vCenter_to_LinkedvCenter_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to all Linked vCenters. This will also add all"+[char]13+[char]10+"metadata to the Visio shapes."+[char]13+[char]10)
#~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton = New-Object System.Windows.Forms.Button
$VisioOpenOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$VisioOpenOutputButton.Location = New-Object System.Drawing.Point(230, 230)
$VisioOpenOutputButton.Size = New-Object System.Drawing.Size(740, 25)
$VisioOpenOutputButton.TabIndex = 52
$VisioOpenOutputButton.Text = "Select Visio Output Folder"
$VisioOpenOutputButton.UseVisualStyleBackColor = $false
$VisioOpenOutputButton.BackColor = [System.Drawing.Color]::LightGray
#~~< VisioOpenOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VisioOpenOutputButtonToolTip.AutoPopDelay = 5000
$VisioOpenOutputButtonToolTip.InitialDelay = 50
$VisioOpenOutputButtonToolTip.IsBalloon = $true
$VisioOpenOutputButtonToolTip.ReshowDelay = 100
$VisioOpenOutputButtonToolTip.SetToolTip($VisioOpenOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"Visio Drawings."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green.")
#~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOutputLabel = New-Object System.Windows.Forms.Label
$VisioOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$VisioOutputLabel.Location = New-Object System.Drawing.Point(10, 230)
$VisioOutputLabel.Size = New-Object System.Drawing.Size(215, 25)
$VisioOutputLabel.TabIndex = 51
$VisioOutputLabel.Text = "Visio Output Folder:"
#~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton = New-Object System.Windows.Forms.Button
$CsvValidationButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CsvValidationButton.Location = New-Object System.Drawing.Point(8, 200)
$CsvValidationButton.Size = New-Object System.Drawing.Size(200, 25)
$CsvValidationButton.TabIndex = 50
$CsvValidationButton.Text = "Check for CSVs"
$CsvValidationButton.UseVisualStyleBackColor = $false
$CsvValidationButton.BackColor = [System.Drawing.Color]::LightGray
#~~< CsvValidationButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CsvValidationButtonToolTip.IsBalloon = $true
$CsvValidationButtonToolTip.SetToolTip($CsvValidationButton, "Click to validate that the required CSV files are present."+[char]13+[char]10+"You must validate files prior to drawing Visio.")
#~~< LinkedvCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 180)
$LinkedvCenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$LinkedvCenterCsvValidationCheck.TabIndex = 49
$LinkedvCenterCsvValidationCheck.Text = ""
#~~< LinkedvCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidation = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidation.Location = New-Object System.Drawing.Point(530, 180)
$LinkedvCenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$LinkedvCenterCsvValidation.TabIndex = 48
$LinkedvCenterCsvValidation.Text = "Linked vCenter CSV File:"
#~~< SnapshotCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidationCheck = New-Object System.Windows.Forms.Label
$SnapshotCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 160)
$SnapshotCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$SnapshotCsvValidationCheck.TabIndex = 47
$SnapshotCsvValidationCheck.Text = ""
#~~< SnapshotCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidation = New-Object System.Windows.Forms.Label
$SnapshotCsvValidation.Location = New-Object System.Drawing.Point(530, 160)
$SnapshotCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$SnapshotCsvValidation.TabIndex = 46
$SnapshotCsvValidation.Text = "Snapshot CSV File:"
#~~< ResourcePoolCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidationCheck = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 140)
$ResourcePoolCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ResourcePoolCsvValidationCheck.TabIndex = 45
$ResourcePoolCsvValidationCheck.Text = ""
#~~< ResourcePoolCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidation = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidation.Location = New-Object System.Drawing.Point(530, 140)
$ResourcePoolCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$ResourcePoolCsvValidation.TabIndex = 44
$ResourcePoolCsvValidation.Text = "Resource Pool CSV File:"
#~~< DrsVmHostRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 120)
$DrsVmHostRuleCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsVmHostRuleCsvValidationCheck.TabIndex = 43
$DrsVmHostRuleCsvValidationCheck.Text = ""
#~~< DrsVmHostRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidation = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidation.Location = New-Object System.Drawing.Point(530, 120)
$DrsVmHostRuleCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsVmHostRuleCsvValidation.TabIndex = 42
$DrsVmHostRuleCsvValidation.Text = "DRS VmHost Rule CSV File:"
#~~< DrsClusterGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 100)
$DrsClusterGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsClusterGroupCsvValidationCheck.TabIndex = 41
$DrsClusterGroupCsvValidationCheck.Text = ""
#~~< DrsClusterGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidation = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidation.Location = New-Object System.Drawing.Point(530, 100)
$DrsClusterGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsClusterGroupCsvValidation.TabIndex = 40
$DrsClusterGroupCsvValidation.Text = "DRS Cluster Group CSV File:"
#~~< DrsRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 80)
$DrsRuleCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsRuleCsvValidationCheck.TabIndex = 39
$DrsRuleCsvValidationCheck.Text = ""
#~~< DrsRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidation = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidation.Location = New-Object System.Drawing.Point(530, 80)
$DrsRuleCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsRuleCsvValidation.TabIndex = 38
$DrsRuleCsvValidation.Text = "DRS Rule CSV File:"
#~~< RdmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidationCheck = New-Object System.Windows.Forms.Label
$RdmCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 60)
$RdmCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$RdmCsvValidationCheck.TabIndex = 37
$RdmCsvValidationCheck.Text = ""
#~~< RdmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidation = New-Object System.Windows.Forms.Label
$RdmCsvValidation.Location = New-Object System.Drawing.Point(530, 60)
$RdmCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$RdmCsvValidation.TabIndex = 36
$RdmCsvValidation.Text = "RDM CSV File:"
#~~< FolderCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidationCheck = New-Object System.Windows.Forms.Label
$FolderCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 40)
$FolderCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$FolderCsvValidationCheck.TabIndex = 35
$FolderCsvValidationCheck.Text = ""
$FolderCsvValidationCheck.add_Click({Label1Click($FolderCsvValidationCheck)})
#~~< FolderCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidation = New-Object System.Windows.Forms.Label
$FolderCsvValidation.Location = New-Object System.Drawing.Point(530, 40)
$FolderCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$FolderCsvValidation.TabIndex = 34
$FolderCsvValidation.Text = "Folder CSV File:"
#~~< VdsPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 180)
$VdsPnicCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsPnicCsvValidationCheck.TabIndex = 33
$VdsPnicCsvValidationCheck.Text = ""
#~~< VdsPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidation = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 180)
$VdsPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsPnicCsvValidation.TabIndex = 32
$VdsPnicCsvValidation.Text = "Vds pNIC CSV File:"
#~~< VdsVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 160)
$VdsVmkernelCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsVmkernelCsvValidationCheck.TabIndex = 31
$VdsVmkernelCsvValidationCheck.Text = ""
#~~< VdsVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 160)
$VdsVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsVmkernelCsvValidation.TabIndex = 30
$VdsVmkernelCsvValidation.Text = "Vds VMkernel CSV File:"
#~~< VdsPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 140)
$VdsPortGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsPortGroupCsvValidationCheck.TabIndex = 29
$VdsPortGroupCsvValidationCheck.Text = ""
#~~< VdsPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidation = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidation.Location = New-Object System.Drawing.Point(270, 140)
$VdsPortGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsPortGroupCsvValidation.TabIndex = 28
$VdsPortGroupCsvValidation.Text = "Vds Port Group CSV File:"
#~~< VdSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 120)
$VdSwitchCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdSwitchCsvValidationCheck.TabIndex = 27
$VdSwitchCsvValidationCheck.Text = ""
#~~< VdSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 120)
$VdSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdSwitchCsvValidation.TabIndex = 26
$VdSwitchCsvValidation.Text = "Distributed Switch CSV File:"
#~~< VssPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssPnicCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 100)
$VssPnicCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssPnicCsvValidationCheck.TabIndex = 25
$VssPnicCsvValidationCheck.Text = ""
#~~< VssPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidation = New-Object System.Windows.Forms.Label
$VssPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 100)
$VssPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssPnicCsvValidation.TabIndex = 24
$VssPnicCsvValidation.Text = "Vss pNIC CSV File:"
#~~< VssVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 80)
$VssVmkernelCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssVmkernelCsvValidationCheck.TabIndex = 23
$VssVmkernelCsvValidationCheck.Text = ""
#~~< VssVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 80)
$VssVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssVmkernelCsvValidation.TabIndex = 22
$VssVmkernelCsvValidation.Text = "Vss VMkernel CSV File:"
#~~< VssPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 60)
$VssPortGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssPortGroupCsvValidationCheck.TabIndex = 21
$VssPortGroupCsvValidationCheck.Text = ""
$VssPortGroupCsvValidationCheck.add_Click({VssPortGroupCsvValidationCheckClick($VssPortGroupCsvValidationCheck)})
#~~< VssPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidation = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidation.Location = New-Object System.Drawing.Point(270, 60)
$VssPortGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssPortGroupCsvValidation.TabIndex = 20
$VssPortGroupCsvValidation.Text = "Vss Port Group CSV File:"
#~~< VsSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidationCheck = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 40)
$VsSwitchCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VsSwitchCsvValidationCheck.TabIndex = 19
$VsSwitchCsvValidationCheck.Text = ""
#~~< VsSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 40)
$VsSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VsSwitchCsvValidation.TabIndex = 18
$VsSwitchCsvValidation.Text = "Standard Switch CSV File:"
#~~< DatastoreCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatastoreCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 180)
$DatastoreCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreCsvValidationCheck.TabIndex = 17
$DatastoreCsvValidationCheck.Text = ""
#~~< DatastoreCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidation = New-Object System.Windows.Forms.Label
$DatastoreCsvValidation.Location = New-Object System.Drawing.Point(10, 180)
$DatastoreCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatastoreCsvValidation.TabIndex = 16
$DatastoreCsvValidation.Text = "Datastore CSV File:"
#~~< DatastoreClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 160)
$DatastoreClusterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreClusterCsvValidationCheck.TabIndex = 15
$DatastoreClusterCsvValidationCheck.Text = ""
#~~< DatastoreClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidation = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidation.Location = New-Object System.Drawing.Point(10, 160)
$DatastoreClusterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatastoreClusterCsvValidation.TabIndex = 14
$DatastoreClusterCsvValidation.Text = "Datastore Cluster CSV File:"
#~~< TemplateCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidationCheck = New-Object System.Windows.Forms.Label
$TemplateCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 140)
$TemplateCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$TemplateCsvValidationCheck.TabIndex = 13
$TemplateCsvValidationCheck.Text = ""
#~~< TemplateCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidation = New-Object System.Windows.Forms.Label
$TemplateCsvValidation.Location = New-Object System.Drawing.Point(10, 140)
$TemplateCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$TemplateCsvValidation.TabIndex = 12
$TemplateCsvValidation.Text = "Template CSV File:"
#~~< VmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidationCheck = New-Object System.Windows.Forms.Label
$VmCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 120)
$VmCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VmCsvValidationCheck.TabIndex = 11
$VmCsvValidationCheck.Text = ""
#~~< VmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidation = New-Object System.Windows.Forms.Label
$VmCsvValidation.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmCsvValidation.TabIndex = 10
$VmCsvValidation.Text = "Virtual Machine CSV File:"
#~~< VmHostCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationCheck = New-Object System.Windows.Forms.Label
$VmHostCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 100)
$VmHostCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationCheck.TabIndex = 9
$VmHostCsvValidationCheck.Text = ""
$VmHostCsvValidationCheck.add_Click({VmHostCsvValidationCheckClick($VmHostCsvValidationCheck)})
#~~< VmHostCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidation = New-Object System.Windows.Forms.Label
$VmHostCsvValidation.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmHostCsvValidation.TabIndex = 8
$VmHostCsvValidation.Text = "Host CSV File:"
#~~< ClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidationCheck = New-Object System.Windows.Forms.Label
$ClusterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 80)
$ClusterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ClusterCsvValidationCheck.TabIndex = 7
$ClusterCsvValidationCheck.Text = ""
#~~< ClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidation = New-Object System.Windows.Forms.Label
$ClusterCsvValidation.Location = New-Object System.Drawing.Point(10, 80)
$ClusterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$ClusterCsvValidation.TabIndex = 6
$ClusterCsvValidation.Text = "Cluster CSV File:"
#~~< DatacenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatacenterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 60)
$DatacenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatacenterCsvValidationCheck.TabIndex = 5
$DatacenterCsvValidationCheck.Text = ""
#~~< DatacenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidation = New-Object System.Windows.Forms.Label
$DatacenterCsvValidation.Location = New-Object System.Drawing.Point(10, 60)
$DatacenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatacenterCsvValidation.TabIndex = 4
$DatacenterCsvValidation.Text = "Datacenter CSV File:"
#~~< vCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidation = New-Object System.Windows.Forms.Label
$vCenterCsvValidation.Location = New-Object System.Drawing.Point(10, 40)
$vCenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$vCenterCsvValidation.TabIndex = 2
$vCenterCsvValidation.Text = "vCenter CSV File:"
#~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton = New-Object System.Windows.Forms.Button
$DrawCsvInputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCsvInputButton.Location = New-Object System.Drawing.Point(220, 10)
$DrawCsvInputButton.Size = New-Object System.Drawing.Size(750, 25)
$DrawCsvInputButton.TabIndex = 1
$DrawCsvInputButton.Text = "Select CSV Input Folder"
$DrawCsvInputButton.UseVisualStyleBackColor = $false
$DrawCsvInputButton.BackColor = [System.Drawing.Color]::LightGray
#~~< DrawCsvInputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawCsvInputButtonToolTip.AutoPopDelay = 5000
$DrawCsvInputButtonToolTip.InitialDelay = 50
$DrawCsvInputButtonToolTip.IsBalloon = $true
$DrawCsvInputButtonToolTip.ReshowDelay = 100
$DrawCsvInputButtonToolTip.SetToolTip($DrawCsvInputButton, "Click to select the folder where the CSV"+[char]39+"s are located."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green.")
#~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputLabel = New-Object System.Windows.Forms.Label
$DrawCsvInputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DrawCsvInputLabel.Location = New-Object System.Drawing.Point(10, 10)
$DrawCsvInputLabel.Size = New-Object System.Drawing.Size(190, 25)
$DrawCsvInputLabel.TabIndex = 0
$DrawCsvInputLabel.Text = "CSV Input Folder:"
$TabDraw.Controls.Add($vCenterCsvValidationCheck)
$TabDraw.Controls.Add($OpenVisioButton)
$TabDraw.Controls.Add($DrawButton)
$TabDraw.Controls.Add($DrawCheckButton)
$TabDraw.Controls.Add($DrawUncheckButton)
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_DrawCheckBox)
$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)
$TabDraw.Controls.Add($VDSPortGroup_to_VM_DrawCheckBox)
$TabDraw.Controls.Add($VMK_to_VDS_Complete)
$TabDraw.Controls.Add($VMK_to_VDS_DrawCheckBox)
$TabDraw.Controls.Add($VDS_to_Host_Complete)
$TabDraw.Controls.Add($VDS_to_Host_DrawCheckBox)
$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)
$TabDraw.Controls.Add($VSSPortGroup_to_VM_DrawCheckBox)
$TabDraw.Controls.Add($VMK_to_VSS_Complete)
$TabDraw.Controls.Add($VMK_to_VSS_DrawCheckBox)
$TabDraw.Controls.Add($VSS_to_Host_Complete)
$TabDraw.Controls.Add($VSS_to_Host_DrawCheckBox)
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_DrawCheckBox)
$TabDraw.Controls.Add($Snapshot_to_VM_Complete)
$TabDraw.Controls.Add($Snapshot_to_VM_DrawCheckBox)
$TabDraw.Controls.Add($Datastore_to_Host_Complete)
$TabDraw.Controls.Add($Datastore_to_Host_DrawCheckBox)
$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)
$TabDraw.Controls.Add($VM_to_ResourcePool_DrawCheckBox)
$TabDraw.Controls.Add($VM_to_Datastore_Complete)
$TabDraw.Controls.Add($VM_to_Datastore_DrawCheckBox)
$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)
$TabDraw.Controls.Add($SRM_Protected_VMs_DrawCheckBox)
$TabDraw.Controls.Add($VMs_with_RDMs_Complete)
$TabDraw.Controls.Add($VMs_with_RDMs_DrawCheckBox)
$TabDraw.Controls.Add($VM_to_Folder_Complete)
$TabDraw.Controls.Add($VM_to_Folder_DrawCheckBox)
$TabDraw.Controls.Add($VM_to_Host_Complete)
$TabDraw.Controls.Add($VM_to_Host_DrawCheckBox)
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_DrawCheckBox)
$TabDraw.Controls.Add($VisioOpenOutputButton)
$TabDraw.Controls.Add($VisioOutputLabel)
$TabDraw.Controls.Add($CsvValidationButton)
$TabDraw.Controls.Add($LinkedvCenterCsvValidationCheck)
$TabDraw.Controls.Add($LinkedvCenterCsvValidation)
$TabDraw.Controls.Add($SnapshotCsvValidationCheck)
$TabDraw.Controls.Add($SnapshotCsvValidation)
$TabDraw.Controls.Add($ResourcePoolCsvValidationCheck)
$TabDraw.Controls.Add($ResourcePoolCsvValidation)
$TabDraw.Controls.Add($DrsVmHostRuleCsvValidationCheck)
$TabDraw.Controls.Add($DrsVmHostRuleCsvValidation)
$TabDraw.Controls.Add($DrsClusterGroupCsvValidationCheck)
$TabDraw.Controls.Add($DrsClusterGroupCsvValidation)
$TabDraw.Controls.Add($DrsRuleCsvValidationCheck)
$TabDraw.Controls.Add($DrsRuleCsvValidation)
$TabDraw.Controls.Add($RdmCsvValidationCheck)
$TabDraw.Controls.Add($RdmCsvValidation)
$TabDraw.Controls.Add($FolderCsvValidationCheck)
$TabDraw.Controls.Add($FolderCsvValidation)
$TabDraw.Controls.Add($VdsPnicCsvValidationCheck)
$TabDraw.Controls.Add($VdsPnicCsvValidation)
$TabDraw.Controls.Add($VdsVmkernelCsvValidationCheck)
$TabDraw.Controls.Add($VdsVmkernelCsvValidation)
$TabDraw.Controls.Add($VdsPortGroupCsvValidationCheck)
$TabDraw.Controls.Add($VdsPortGroupCsvValidation)
$TabDraw.Controls.Add($VdSwitchCsvValidationCheck)
$TabDraw.Controls.Add($VdSwitchCsvValidation)
$TabDraw.Controls.Add($VssPnicCsvValidationCheck)
$TabDraw.Controls.Add($VssPnicCsvValidation)
$TabDraw.Controls.Add($VssVmkernelCsvValidationCheck)
$TabDraw.Controls.Add($VssVmkernelCsvValidation)
$TabDraw.Controls.Add($VssPortGroupCsvValidationCheck)
$TabDraw.Controls.Add($VssPortGroupCsvValidation)
$TabDraw.Controls.Add($VsSwitchCsvValidationCheck)
$TabDraw.Controls.Add($VsSwitchCsvValidation)
$TabDraw.Controls.Add($DatastoreCsvValidationCheck)
$TabDraw.Controls.Add($DatastoreCsvValidation)
$TabDraw.Controls.Add($DatastoreClusterCsvValidationCheck)
$TabDraw.Controls.Add($DatastoreClusterCsvValidation)
$TabDraw.Controls.Add($TemplateCsvValidationCheck)
$TabDraw.Controls.Add($TemplateCsvValidation)
$TabDraw.Controls.Add($VmCsvValidationCheck)
$TabDraw.Controls.Add($VmCsvValidation)
$TabDraw.Controls.Add($VmHostCsvValidationCheck)
$TabDraw.Controls.Add($VmHostCsvValidation)
$TabDraw.Controls.Add($ClusterCsvValidationCheck)
$TabDraw.Controls.Add($ClusterCsvValidation)
$TabDraw.Controls.Add($DatacenterCsvValidationCheck)
$TabDraw.Controls.Add($DatacenterCsvValidation)
$TabDraw.Controls.Add($vCenterCsvValidation)
$TabDraw.Controls.Add($DrawCsvInputButton)
$TabDraw.Controls.Add($DrawCsvInputLabel)
$SubTab.Controls.Add($TabDirections)
$SubTab.Controls.Add($TabCapture)
$SubTab.Controls.Add($TabDraw)
$SubTab.SelectedIndex = 0
#~~< MainTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainTab = New-Object System.Windows.Forms.TabControl
$MainTab.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$MainTab.ItemSize = New-Object System.Drawing.Size(85, 20)
$MainTab.Location = New-Object System.Drawing.Point(10, 30)
$MainTab.Size = New-Object System.Drawing.Size(990, 98)
$MainTab.TabIndex = 1
$MainTab.Text = "Prerequisites"
#~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Prerequisites = New-Object System.Windows.Forms.TabPage
$Prerequisites.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Prerequisites.Location = New-Object System.Drawing.Point(4, 24)
$Prerequisites.Padding = New-Object System.Windows.Forms.Padding(3)
$Prerequisites.Size = New-Object System.Drawing.Size(982, 70)
$Prerequisites.TabIndex = 0
$Prerequisites.Text = "Prerequisites"
$Prerequisites.ToolTipText = "Prerequisites: These items are needed in order to run this script."
$Prerequisites.BackColor = [System.Drawing.Color]::LightGray
#~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioInstalled = New-Object System.Windows.Forms.Label
$VisioInstalled.Location = New-Object System.Drawing.Point(490, 40)
$VisioInstalled.Size = New-Object System.Drawing.Size(320, 20)
$VisioInstalled.TabIndex = 7
$VisioInstalled.Text = ""
#~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioLabel = New-Object System.Windows.Forms.Label
$VisioLabel.Location = New-Object System.Drawing.Point(450, 40)
$VisioLabel.Size = New-Object System.Drawing.Size(40, 20)
$VisioLabel.TabIndex = 6
$VisioLabel.Text = "Visio:"
#~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliInstalled = New-Object System.Windows.Forms.Label
$PowerCliInstalled.Location = New-Object System.Drawing.Point(520, 15)
$PowerCliInstalled.Size = New-Object System.Drawing.Size(400, 20)
$PowerCliInstalled.TabIndex = 5
$PowerCliInstalled.Text = ""
#~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliLabel = New-Object System.Windows.Forms.Label
$PowerCliLabel.Location = New-Object System.Drawing.Point(450, 15)
$PowerCliLabel.Size = New-Object System.Drawing.Size(64, 20)
$PowerCliLabel.TabIndex = 4
$PowerCliLabel.Text = "PowerCLI:"
#~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleInstalled = New-Object System.Windows.Forms.Label
$PowerCliModuleInstalled.Location = New-Object System.Drawing.Point(128, 40)
$PowerCliModuleInstalled.Size = New-Object System.Drawing.Size(320, 20)
$PowerCliModuleInstalled.TabIndex = 3
$PowerCliModuleInstalled.Text = ""
#~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleLabel = New-Object System.Windows.Forms.Label
$PowerCliModuleLabel.Location = New-Object System.Drawing.Point(10, 40)
$PowerCliModuleLabel.Size = New-Object System.Drawing.Size(110, 20)
$PowerCliModuleLabel.TabIndex = 2
$PowerCliModuleLabel.Text = "PowerCLI Module:"
#~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellInstalled = New-Object System.Windows.Forms.Label
$PowershellInstalled.Location = New-Object System.Drawing.Point(96, 15)
$PowershellInstalled.Size = New-Object System.Drawing.Size(350, 20)
$PowershellInstalled.TabIndex = 1
$PowershellInstalled.Text = ""
#~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellLabel = New-Object System.Windows.Forms.Label
$PowershellLabel.Location = New-Object System.Drawing.Point(10, 15)
$PowershellLabel.Size = New-Object System.Drawing.Size(75, 20)
$PowershellLabel.TabIndex = 0
$PowershellLabel.Text = "Powershell:"
$Prerequisites.Controls.Add($VisioInstalled)
$Prerequisites.Controls.Add($VisioLabel)
$Prerequisites.Controls.Add($PowerCliInstalled)
$Prerequisites.Controls.Add($PowerCliLabel)
$Prerequisites.Controls.Add($PowerCliModuleInstalled)
$Prerequisites.Controls.Add($PowerCliModuleLabel)
$Prerequisites.Controls.Add($PowershellInstalled)
$Prerequisites.Controls.Add($PowershellLabel)
#~~< vCenterInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfo = New-Object System.Windows.Forms.TabPage
$vCenterInfo.Location = New-Object System.Drawing.Point(4, 24)
$vCenterInfo.Padding = New-Object System.Windows.Forms.Padding(3)
$vCenterInfo.Size = New-Object System.Drawing.Size(982, 70)
$vCenterInfo.TabIndex = 1
$vCenterInfo.Text = "vCenter Info"
$vCenterInfo.BackColor = [System.Drawing.Color]::LightGray
#~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$ConnectButton.Location = New-Object System.Drawing.Point(8, 37)
$ConnectButton.Size = New-Object System.Drawing.Size(345, 25)
$ConnectButton.TabIndex = 6
$ConnectButton.Text = "Connect to vCenter"
$ConnectButton.UseVisualStyleBackColor = $true
#~~< ConnectButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ConnectButtonToolTip.AutoPopDelay = 5000
$ConnectButtonToolTip.InitialDelay = 50
$ConnectButtonToolTip.IsBalloon = $true
$ConnectButtonToolTip.ReshowDelay = 100
$ConnectButtonToolTip.SetToolTip($ConnectButton, "Click to connect to vCenter."+[char]13+[char]10+[char]13+[char]10+"If connected this button will turn green and show connected to the name entered in the vCenter box."+[char]13+[char]10+[char]13+[char]10+"If disconnected or unable to connect this button will display red text, indicating that you were unable to"+[char]13+[char]10+"connect to vCenter either due to bad creditials, not being on the same network or insufficient access to vCenter.")
#~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Location = New-Object System.Drawing.Point(734, 8)
$PasswordTextBox.Size = New-Object System.Drawing.Size(238, 21)
$PasswordTextBox.TabIndex = 5
$PasswordTextBox.Text = ""
$PasswordTextBox.UseSystemPasswordChar = $true
#~~< PasswordToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordToolTip = New-Object System.Windows.Forms.ToolTip($components)
$PasswordToolTip.AutoPopDelay = 5000
$PasswordToolTip.InitialDelay = 50
$PasswordToolTip.IsBalloon = $true
$PasswordToolTip.ReshowDelay = 100
$PasswordToolTip.SetToolTip($PasswordTextBox, "Enter Passwrd."+[char]13+[char]10+[char]13+[char]10+"Characters will not be seen.")
#~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Location = New-Object System.Drawing.Point(656, 11)
$PasswordLabel.Size = New-Object System.Drawing.Size(70, 20)
$PasswordLabel.TabIndex = 4
$PasswordLabel.Text = "Password:"
#~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameTextBox = New-Object System.Windows.Forms.TextBox
$UserNameTextBox.Location = New-Object System.Drawing.Point(402, 8)
$UserNameTextBox.Size = New-Object System.Drawing.Size(238, 21)
$UserNameTextBox.TabIndex = 3
$UserNameTextBox.Text = ""
#~~< UserNameToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameToolTip = New-Object System.Windows.Forms.ToolTip($components)
$UserNameToolTip.AutoPopDelay = 5000
$UserNameToolTip.InitialDelay = 50
$UserNameToolTip.IsBalloon = $true
$UserNameToolTip.ReshowDelay = 100
$UserNameToolTip.SetToolTip($UserNameTextBox, "Enter User Name."+[char]13+[char]10+[char]13+[char]10+"Example:"+[char]13+[char]10+"administrator@vsphere.local"+[char]13+[char]10+"Domain\User")
#~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameLabel = New-Object System.Windows.Forms.Label
$UserNameLabel.Location = New-Object System.Drawing.Point(324, 11)
$UserNameLabel.Size = New-Object System.Drawing.Size(70, 20)
$UserNameLabel.TabIndex = 2
$UserNameLabel.Text = "User Name:"
#~~< VcenterTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterTextBox = New-Object System.Windows.Forms.TextBox
$VcenterTextBox.Location = New-Object System.Drawing.Point(78, 8)
$VcenterTextBox.Size = New-Object System.Drawing.Size(238, 21)
$VcenterTextBox.TabIndex = 1
$VcenterTextBox.Text = ""
$VcenterToolTip.SetToolTip($VcenterTextBox, "Enter vCenter name")
#~~< VcenterLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterLabel = New-Object System.Windows.Forms.Label
$VcenterLabel.Location = New-Object System.Drawing.Point(8, 11)
$VcenterLabel.Size = New-Object System.Drawing.Size(70, 20)
$VcenterLabel.TabIndex = 0
$VcenterLabel.Text = "vCenter:"
$vCenterInfo.Controls.Add($ConnectButton)
$vCenterInfo.Controls.Add($PasswordTextBox)
$vCenterInfo.Controls.Add($PasswordLabel)
$vCenterInfo.Controls.Add($UserNameTextBox)
$vCenterInfo.Controls.Add($UserNameLabel)
$vCenterInfo.Controls.Add($VcenterTextBox)
$vCenterInfo.Controls.Add($VcenterLabel)
$MainTab.Controls.Add($Prerequisites)
$MainTab.Controls.Add($vCenterInfo)
$MainTab.SelectedIndex = 0
#~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu = New-Object System.Windows.Forms.MenuStrip
$MainMenu.Location = New-Object System.Drawing.Point(0, 0)
$MainMenu.Size = New-Object System.Drawing.Size(1010, 24)
$MainMenu.TabIndex = 0
$MainMenu.Text = "MainMenu"
#~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$FileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
$FileToolStripMenuItem.Text = "File"
#~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ExitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExitToolStripMenuItem.Size = New-Object System.Drawing.Size(92, 22)
$ExitToolStripMenuItem.Text = "Exit"
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($ExitToolStripMenuItem)))
#~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HelpToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$HelpToolStripMenuItem.Size = New-Object System.Drawing.Size(44, 20)
$HelpToolStripMenuItem.Text = "Help"
#~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutToolStripMenuItem.Size = New-Object System.Drawing.Size(107, 22)
$AboutToolStripMenuItem.Text = "About"
$HelpToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($AboutToolStripMenuItem)))
$MainMenu.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@($FileToolStripMenuItem, $HelpToolStripMenuItem)))
$vDiagram.Controls.Add($SubTab)
$vDiagram.Controls.Add($MainTab)
$vDiagram.Controls.Add($MainMenu)
#~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$VisioBrowse.Description = "Select a directory"
$VisioBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$CaptureCsvBrowse.Description = "Select a directory"
$CaptureCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#~~< DatastoreCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatastoreCsvToolTip.AutoPopDelay = 5000
$DatastoreCsvToolTip.InitialDelay = 50
$DatastoreCsvToolTip.IsBalloon = $true
$DatastoreCsvToolTip.ReshowDelay = 100
#~~< vCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$vCenterCsvToolTip.AutoPopDelay = 5000
$vCenterCsvToolTip.InitialDelay = 50
$vCenterCsvToolTip.IsBalloon = $true
$vCenterCsvToolTip.ReshowDelay = 100
#~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$DrawCsvBrowse.Description = "Select a directory"
$DrawCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer

#endregion

#region ~~< Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellCheck = $PSVersionTable.PSVersion
if ($PowershellCheck.Major -ge 4)
{
	$PowershellInstalled.Forecolor = "Green"
	$PowershellInstalled.Text = "Installed Version $PowershellCheck"
}
else
{
	$PowershellInstalled.Forecolor = "Red"
	$PowershellInstalled.Text = "Not installed or Powershell version lower than 4"
}
#endregion ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleCheck = (Get-Module VMware.PowerCLI -ListAvailable | Where-Object { $_.Name -eq "VMware.PowerCLI" })
$PowerCliModuleVersion = ($PowerCliModuleCheck.Version)
if ($PowerCliModuleCheck -ne $null)
{
	$PowerCliModuleInstalled.Forecolor = "Green"
	$PowerCliModuleInstalled.Text = "Installed Version $PowerCliModuleVersion"
}
else
{
	$PowerCliModuleInstalled.Forecolor = "Red"
	$PowerCliModuleInstalled.Text = "Not Installed"
}
#endregion ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ((Get-PSSnapin -registered | Where-Object { $_.Name -eq "VMware.VimAutomation.Core" }) -ne $null)
{
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerClI Installed"
}
elseif ($PowerCliModuleCheck -ne $null)
{
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerCLI Module Installed"
}
else
{
	$PowerCliInstalled.Forecolor = "Red"
	$PowerCliInstalled.Text = "PowerCLI or PowerCli Module not installed"
}
#endregion ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ((Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*"} | Select-Object DisplayName) -or (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*"} | Select-Object DisplayName) -ne $null)
{
	$VisioInstalled.Forecolor = "Green"
	$VisioInstalled.Text = "Installed"
}
else
{
	$VisioInstalled.Forecolor = "Red"
	$VisioInstalled.Text = "Visio is Not Installed"
}
#endregion ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton.Add_MouseClick({ $Connected = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager ; 
	if ($Connected -eq $null)
	{
		$ConnectButton.Forecolor = [System.Drawing.Color]::Red ; 
		$ConnectButton.Text = "Unable to Connect"
	}
	else
	{
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green ;
		$ConnectButton.Text = "Connected to $DefaultViserver"
	}
} )
$ConnectButton.Add_Click({ Connect_vCenter })
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton.Add_Click( { Find_CaptureCsvFolder ; 
	if ($CaptureCsvFolder -eq $null) 
	{
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$CaptureCsvOutputButton.Text = "Folder Not Selected"
	}
	else
	{
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$CaptureCsvOutputButton.Text = $CaptureCsvFolder
	}
	Check_CaptureCsvFolder
} )
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton.Add_Click({ $vCenterCsvCheckBox.CheckState = "UnChecked" ;
	$DatacenterCsvCheckBox.CheckState = "UnChecked" ;
	$ClusterCsvCheckBox.CheckState = "UnChecked" ;
	$VmHostCsvCheckBox.CheckState = "UnChecked" ;
	$VmCsvCheckBox.CheckState = "UnChecked" ;
	$TemplateCsvCheckBox.CheckState = "UnChecked" ;
	$DatastoreClusterCsvCheckBox.CheckState = "UnChecked" ;
	$DatastoreCsvCheckBox.CheckState = "UnChecked" ;
	$VsSwitchCsvCheckBox.CheckState = "UnChecked" ;
	$VssPortGroupCsvCheckBox.CheckState = "UnChecked" ;
	$VssVmkernelCsvCheckBox.CheckState = "UnChecked" ;
	$VssPnicCsvCheckBox.CheckState = "UnChecked" ;
	$VdSwitchCsvCheckBox.CheckState = "UnChecked" ;
	$VdsPortGroupCsvCheckBox.CheckState = "UnChecked" ;
	$VdsVmkernelCsvCheckBox.CheckState = "UnChecked" ;
	$VdsPnicCsvCheckBox.CheckState = "UnChecked" ;
	$FolderCsvCheckBox.CheckState = "UnChecked" ;
	$RdmCsvCheckBox.CheckState = "UnChecked" ;
	$DrsRuleCsvCheckBox.CheckState = "UnChecked" ;
	$DrsClusterGroupCsvCheckBox.CheckState = "UnChecked" ;
	$DrsVmHostRuleCsvCheckBox.CheckState = "UnChecked" ;
	$ResourcePoolCsvCheckBox.CheckState = "UnChecked";
	$SnapshotCsvCheckBox.CheckState = "UnChecked";
	$LinkedvCenterCsvCheckBox.CheckState = "UnChecked"	
} )
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton.Add_Click({ $vCenterCsvCheckBox.CheckState = "Checked" ;
	$DatacenterCsvCheckBox.CheckState = "Checked" ;
	$ClusterCsvCheckBox.CheckState = "Checked" ;
	$VmHostCsvCheckBox.CheckState = "Checked" ;
	$VmCsvCheckBox.CheckState = "Checked" ;
	$TemplateCsvCheckBox.CheckState = "Checked" ;
	$DatastoreClusterCsvCheckBox.CheckState = "Checked" ;
	$DatastoreCsvCheckBox.CheckState = "Checked" ;
	$VsSwitchCsvCheckBox.CheckState = "Checked" ;
	$VssPortGroupCsvCheckBox.CheckState = "Checked" ;
	$VssVmkernelCsvCheckBox.CheckState = "Checked" ;
	$VssPnicCsvCheckBox.CheckState = "Checked" ;
	$VdSwitchCsvCheckBox.CheckState = "Checked" ;
	$VdsPortGroupCsvCheckBox.CheckState = "Checked" ;
	$VdsVmkernelCsvCheckBox.CheckState = "Checked" ;
	$VdsPnicCsvCheckBox.CheckState = "Checked" ;
	$FolderCsvCheckBox.CheckState = "Checked" ;
	$RdmCsvCheckBox.CheckState = "Checked" ;
	$DrsRuleCsvCheckBox.CheckState = "Checked" ;
	$DrsClusterGroupCsvCheckBox.CheckState = "Checked" ;
	$DrsVmHostRuleCsvCheckBox.CheckState = "Checked" ;
	$ResourcePoolCsvCheckBox.CheckState = "Checked";
	$SnapshotCsvCheckBox.CheckState = "Checked";
	$LinkedvCenterCsvCheckBox.CheckState = "Checked"
})
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton.Add_Click({
	if($CaptureCsvFolder -eq $null)
	{
		$CaptureButton.Forecolor = [System.Drawing.Color]::Red; $CaptureButton.Text = "Folder Not Selected"
	}
	else
	{ 
		if ($vCenterCsvCheckBox.Checked -eq "True")
		{
			$vCenterCsvValidationComplete.Forecolor = "Blue"
			$vCenterCsvValidationComplete.Text = "Processing ....."
			vCenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$vCenterExportFileComplete = $CsvCompleteDir + "-vCenterExport.csv"
			$vCenterCsvComplete = Test-Path $vCenterExportFileComplete
			if ($vCenterCsvComplete -eq $True)
			{
				$vCenterCsvValidationComplete.Forecolor = "Green"
				$vCenterCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$vCenterCsvValidationComplete.Forecolor = "Red"
				$vCenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		Connect_vCenter
		$Connected = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
		if ($Connected -eq $null) { Connect_vCenter }
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green
		$ConnectButton.Text = "Connected to $DefaultViserver"
		if ($DatacenterCsvCheckBox.Checked -eq "True")
		{
			$DatacenterCsvValidationComplete.Forecolor = "Blue"
			$DatacenterCsvValidationComplete.Text = "Processing ....."
			Datacenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatacenterExportFileComplete = $CsvCompleteDir + "-DatacenterExport.csv"
			$DatacenterCsvComplete = Test-Path $DatacenterExportFileComplete
			if ($DatacenterCsvComplete -eq $True)
			{
				$DatacenterCsvValidationComplete.Forecolor = "Green"
				$DatacenterCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DatacenterCsvValidationComplete.Forecolor = "Red"
				$DatacenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($ClusterCsvCheckBox.Checked -eq "True")
		{
			$ClusterCsvValidationComplete.Forecolor = "Blue"
			$ClusterCsvValidationComplete.Text = "Processing ....."
			Cluster_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$ClusterExportFileComplete = $CsvCompleteDir + "-ClusterExport.csv"
			$ClusterCsvComplete = Test-Path $ClusterExportFileComplete
			if ($ClusterCsvComplete -eq $True)
			{
				$ClusterCsvValidationComplete.Forecolor = "Green"
				$ClusterCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$ClusterCsvValidationComplete.Forecolor = "Red"
				$ClusterCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VmHostCsvCheckBox.Checked -eq "True")
		{
			$VmHostCsvValidationComplete.Forecolor = "Blue"
			$VmHostCsvValidationComplete.Text = "Processing ....."
			VmHost_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VmHostExportFileComplete = $CsvCompleteDir + "-VmHostExport.csv"
			$VmHostCsvComplete = Test-Path $VmHostExportFileComplete
			if ($VmHostCsvComplete -eq $True)
			{
				$VmHostCsvValidationComplete.Forecolor = "Green"
				$VmHostCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VmHostCsvValidationComplete.Forecolor = "Red"
				$VmHostCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VmCsvCheckBox.Checked -eq "True")
		{
			$VmCsvValidationComplete.Forecolor = "Blue"
			$VmCsvValidationComplete.Text = "Processing ....."
			Vm_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VmExportFileComplete = $CsvCompleteDir + "-VmExport.csv"
			$VmCsvComplete = Test-Path $VmExportFileComplete
			if ($VmCsvComplete -eq $True)
			{
				$VmCsvValidationComplete.Forecolor = "Green"
				$VmCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VmCsvValidationComplete.Forecolor = "Red"
				$VmCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($TemplateCsvCheckBox.Checked -eq "True")
		{
			$TemplateCsvValidationComplete.Forecolor = "Blue"
			$TemplateCsvValidationComplete.Text = "Processing ....."
			Template_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$TemplateExportFileComplete = $CsvCompleteDir + "-TemplateExport.csv"
			$TemplateCsvComplete = Test-Path $TemplateExportFileComplete
			if ($TemplateCsvComplete -eq $True)
			{
				$TemplateCsvValidationComplete.Forecolor = "Green"
				$TemplateCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$TemplateCsvValidationComplete.Forecolor = "Red"
				$TemplateCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($DatastoreClusterCsvCheckBox.Checked -eq "True")
		{
			$DatastoreClusterCsvValidationComplete.Forecolor = "Blue"
			$DatastoreClusterCsvValidationComplete.Text = "Processing ....."
			DatastoreCluster_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatastoreClusterExportFileComplete = $CsvCompleteDir + "-DatastoreClusterExport.csv"
			$DatastoreClusterCsvComplete = Test-Path $DatastoreClusterExportFileComplete
			if ($DatastoreClusterCsvComplete -eq $True)
			{
				$DatastoreClusterCsvValidationComplete.Forecolor = "Green"
				$DatastoreClusterCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DatastoreClusterCsvValidationComplete.Forecolor = "Red"
				$DatastoreClusterCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($DatastoreCsvCheckBox.Checked -eq "True")
		{
			$DatastoreCsvValidationComplete.Forecolor = "Blue"
			$DatastoreCsvValidationComplete.Text = "Processing ....."
			Datastore_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatastoreExportFileComplete = $CsvCompleteDir + "-DatastoreExport.csv"
			$DatastoreCsvComplete = Test-Path $DatastoreExportFileComplete
			if ($DatastoreCsvComplete -eq $True)
			{
				$DatastoreCsvValidationComplete.Forecolor = "Green"
				$DatastoreCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DatastoreCsvValidationComplete.Forecolor = "Red"
				$DatastoreCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VsSwitchCsvCheckBox.Checked -eq "True")
		{
			$VsSwitchCsvValidationComplete.Forecolor = "Blue"
			$VsSwitchCsvValidationComplete.Text = "Processing ....."
			VsSwitch_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$vSSwitchExportFileComplete = $CsvCompleteDir + "-vSSwitchExport.csv"
			$vSSwitchCsvComplete = Test-Path $vSSwitchExportFileComplete
			if ($vSSwitchCsvComplete -eq $True)
			{
				$vSSwitchCsvValidationComplete.Forecolor = "Green"
				$vSSwitchCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$vSSwitchCsvValidationComplete.Forecolor = "Red"
				$vSSwitchCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VssPortGroupCsvCheckBox.Checked -eq "True")
		{
			$VssPortGroupCsvValidationComplete.Forecolor = "Blue"
			$VssPortGroupCsvValidationComplete.Text = "Processing ....."
			VssPort_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssPortGroupExportFileComplete = $CsvCompleteDir + "-VssPortGroupExport.csv"
			$VssPortGroupCsvComplete = Test-Path $VssPortGroupExportFileComplete
			if ($VssPortGroupCsvComplete -eq $True)
			{
				$VssPortGroupCsvValidationComplete.Forecolor = "Green"
				$VssPortGroupCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VssPortGroupCsvValidationComplete.Forecolor = "Red"
				$VssPortGroupCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VssVmkernelCsvCheckBox.Checked -eq "True")
		{
			$VssVmkernelCsvValidationComplete.Forecolor = "Blue"
			$VssVmkernelCsvValidationComplete.Text = "Processing ....."
			VssVmk_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssVmkernelExportFileComplete = $CsvCompleteDir + "-VssVmkernelExport.csv"
			$VssVmkernelCsvComplete = Test-Path $VssVmkernelExportFileComplete
			if ($VssVmkernelCsvComplete -eq $True)
			{
				$VssVmkernelCsvValidationComplete.Forecolor = "Green"
				$VssVmkernelCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VssVmkernelCsvValidationComplete.Forecolor = "Red"
				$VssVmkernelCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VssPnicCsvCheckBox.Checked -eq "True")
		{
			$VssPnicCsvValidationComplete.Forecolor = "Blue"
			$VssPnicCsvValidationComplete.Text = "Processing ....."
			VssPnic_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssPnicExportFileComplete = $CsvCompleteDir + "-VssPnicExport.csv"
			$VssPnicCsvComplete = Test-Path $VssPnicExportFileComplete
			if ($VssPnicCsvComplete -eq $True)
			{
				$VssPnicCsvValidationComplete.Forecolor = "Green"
				$VssPnicCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VssPnicCsvValidationComplete.Forecolor = "Red"
				$VssPnicCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VdSwitchCsvCheckBox.Checked -eq "True")
		{
			$VdSwitchCsvValidationComplete.Forecolor = "Blue"
			$VdSwitchCsvValidationComplete.Text = "Processing ....."
			VdSwitch_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdSwitchExportFileComplete = $CsvCompleteDir + "-VdSwitchExport.csv"
			$VdSwitchCsvComplete = Test-Path $VdSwitchExportFileComplete
			if ($VdSwitchCsvComplete -eq $True)
			{
				$VdSwitchCsvValidationComplete.Forecolor = "Green"
				$VdSwitchCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VdSwitchCsvValidationComplete.Forecolor = "Red"
				$VdSwitchCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VdsPortGroupCsvCheckBox.Checked -eq "True")
		{
			$VdsPortGroupCsvValidationComplete.Forecolor = "Blue"
			$VdsPortGroupCsvValidationComplete.Text = "Processing ....."
			VdsPort_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsPortGroupExportFileComplete = $CsvCompleteDir + "-VdsPortGroupExport.csv"
			$VdsPortGroupCsvComplete = Test-Path $VdsPortGroupExportFileComplete
			if ($VdsPortGroupCsvComplete -eq $True)
			{
				$VdsPortGroupCsvValidationComplete.Forecolor = "Green"
				$VdsPortGroupCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VdsPortGroupCsvValidationComplete.Forecolor = "Red"
				$VdsPortGroupCsvValidationComplete.Text = "Not Complete"
				
			}
		}
		if ($VdsVmkernelCsvCheckBox.Checked -eq "True")
		{
			$VdsVmkernelCsvValidationComplete.Forecolor = "Blue"
			$VdsVmkernelCsvValidationComplete.Text = "Processing ....."
			VdsVmk_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsVmkernelExportFileComplete = $CsvCompleteDir + "-VdsVmkernelExport.csv"
			$VdsVmkernelCsvComplete = Test-Path $VdsVmkernelExportFileComplete
			if ($VdsVmkernelCsvComplete -eq $True)
			{
				$VdsVmkernelCsvValidationComplete.Forecolor = "Green"
				$VdsVmkernelCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VdsVmkernelCsvValidationComplete.Forecolor = "Red"
				$VdsVmkernelCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($VdsPnicCsvCheckBox.Checked -eq "True")
		{
			$VdsPnicCsvValidationComplete.Forecolor = "Blue"
			$VdsPnicCsvValidationComplete.Text = "Processing ....."
			VdsPnic_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsPnicExportFileComplete = $CsvCompleteDir + "-VdsPnicExport.csv"
			$VdsPnicCsvComplete = Test-Path $VdsPnicExportFileComplete
			if ($VdsPnicCsvComplete -eq $True)
			{
				$VdsPnicCsvValidationComplete.Forecolor = "Green"
				$VdsPnicCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$VdsPnicCsvValidationComplete.Forecolor = "Red"
				$VdsPnicCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($FolderCsvCheckBox.Checked -eq "True")
		{
			$FolderCsvValidationComplete.Forecolor = "Blue"
			$FolderCsvValidationComplete.Text = "Processing ....."
			Folder_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$FolderExportFileComplete = $CsvCompleteDir + "-FolderExport.csv"
			$FolderCsvComplete = Test-Path $FolderExportFileComplete
			if ($FolderCsvComplete -eq $True)
			{
				$FolderCsvValidationComplete.Forecolor = "Green"
				$FolderCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$FolderCsvValidationComplete.Forecolor = "Red"
				$FolderCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($RdmCsvCheckBox.Checked -eq "True")
		{
			$RdmCsvValidationComplete.Forecolor = "Blue"
			$RdmCsvValidationComplete.Text = "Processing ....."
			Rdm_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$RdmExportFileComplete = $CsvCompleteDir + "-RdmExport.csv"
			$RdmCsvComplete = Test-Path $RdmExportFileComplete
			if ($RdmCsvComplete -eq $True)
			{
				$RdmCsvValidationComplete.Forecolor = "Green"
				$RdmCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$RdmCsvValidationComplete.Forecolor = "Red"
				$RdmCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($DrsRuleCsvCheckBox.Checked -eq "True")
		{
			$DrsRuleCsvValidationComplete.Forecolor = "Blue"
			$DrsRuleCsvValidationComplete.Text = "Processing ....."
			Drs_Rule_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsRuleExportFileComplete = $CsvCompleteDir + "-DrsRuleExport.csv"
			$DrsRuleCsvComplete = Test-Path $DrsRuleExportFileComplete
			if ($DrsRuleCsvComplete -eq $True)
			{
				$DrsRuleCsvValidationComplete.Forecolor = "Green"
				$DrsRuleCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DrsRuleCsvValidationComplete.Forecolor = "Red"
				$DrsRuleCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($DrsClusterGroupCsvCheckBox.Checked -eq "True")
		{
			$DrsClusterGroupCsvValidationComplete.Forecolor = "Blue"
			$DrsClusterGroupCsvValidationComplete.Text = "Processing ....."
			Drs_Cluster_Group_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsClusterGroupExportFileComplete = $CsvCompleteDir + "-DrsClusterGroupExport.csv"
			$DrsClusterGroupCsvComplete = Test-Path $DrsClusterGroupExportFileComplete
			if ($DrsClusterGroupCsvComplete -eq $True)
			{
				$DrsClusterGroupCsvValidationComplete.Forecolor = "Green"
				$DrsClusterGroupCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DrsClusterGroupCsvValidationComplete.Forecolor = "Red"
				$DrsClusterGroupCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($DrsVmHostRuleCsvCheckBox.Checked -eq "True")
		{
			$DrsVmHostRuleCsvValidationComplete.Forecolor = "Blue"
			$DrsVmHostRuleCsvValidationComplete.Text = "Processing ....."
			Drs_VmHost_Rule_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsVmHostRuleExportFileComplete = $CsvCompleteDir + "-DrsVmHostRuleExport.csv"
			$DrsVmHostRuleCsvComplete = Test-Path $DrsVmHostRuleExportFileComplete
			if ($DrsVmHostRuleCsvComplete -eq $True)
			{
				$DrsVmHostRuleCsvValidationComplete.Forecolor = "Green"
				$DrsVmHostRuleCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$DrsVmHostRuleCsvValidationComplete.Forecolor = "Red"
				$DrsVmHostRuleCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($ResourcePoolCsvCheckBox.Checked -eq "True")
		{
			$ResourcePoolCsvValidationComplete.Forecolor = "Blue"
			$ResourcePoolCsvValidationComplete.Text = "Processing ....."
			Resource_Pool_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$ResourcePoolExportFileComplete = $CsvCompleteDir + "-ResourcePoolExport.csv"
			$ResourcePoolCsvComplete = Test-Path $ResourcePoolExportFileComplete
			if ($ResourcePoolCsvComplete -eq $True)
			{
				$ResourcePoolCsvValidationComplete.Forecolor = "Green"
				$ResourcePoolCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$ResourcePoolCsvValidationComplete.Forecolor = "Red"
				$ResourcePoolCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($SnapshotCsvCheckBox.Checked -eq "True")
		{
			$SnapshotCsvValidationComplete.Forecolor = "Blue"
			$SnapshotCsvValidationComplete.Text = "Processing ....."
			Snapshot_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$SnapshotExportFileComplete = $CsvCompleteDir + "-SnapshotExport.csv"
			$SnapshotCsvComplete = Test-Path $SnapshotExportFileComplete
			if ($SnapshotCsvComplete -eq $True)
			{
				$SnapshotCsvValidationComplete.Forecolor = "Green"
				$SnapshotCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$SnapshotCsvValidationComplete.Forecolor = "Red"
				$SnapshotCsvValidationComplete.Text = "Not Complete"
			}
		}
		if ($LinkedvCenterCsvCheckBox.Checked -eq "True")
		{
			$LinkedvCenterCsvValidationComplete.Forecolor = "Blue"
			$LinkedvCenterCsvValidationComplete.Text = "Processing ....."
			Linked_vCenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$LinkedvCenterExportFileComplete = $CsvCompleteDir + "-LinkedvCenterExport.csv"
			$LinkedvCenterCsvComplete = Test-Path $LinkedvCenterExportFileComplete
			if ($LinkedvCenterCsvComplete -eq $True)
			{
				$LinkedvCenterCsvValidationComplete.Forecolor = "Green"
				$LinkedvCenterCsvValidationComplete.Text = "Complete"
			}
			else
			{
				$LinkedvCenterCsvValidationComplete.Forecolor = "Red"
				$LinkedvCenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		Disconnect_vCenter
		$ConnectButton.Forecolor = [System.Drawing.Color]::Red
		$ConnectButton.Text = "Disconnected"
		$CaptureButton.Forecolor = [System.Drawing.Color]::Green ; $CaptureButton.Text = "CSV Collection Complete"
	}
})
#endregion ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton.Add_Click({Open_Capture_Folder;
	$VcenterTextBox.Text = "" ;
	$UserNameTextBox.Text = "" ;
	$PasswordTextBox.Text = "" ;
	$PasswordTextBox.UseSystemPasswordChar = $true ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to vCenter" ;
	$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Black ;
	$CaptureCsvOutputButton.Text = "Select Output Folder" ;
	$CaptureButton.Forecolor = [System.Drawing.Color]::Black ;
	$CaptureButton.Text = "Collect CSV Data" ;
	$vCenterCsvCheckBox.CheckState = "Checked" ;
	$vCenterCsvValidationComplete.Text = "" ;
	$DatacenterCsvCheckBox.CheckState = "Checked" ;
	$DatacenterCsvValidationComplete.Text = "" ;
	$ClusterCsvCheckBox.CheckState = "Checked" ;
	$ClusterCsvValidationComplete.Text = "" ;
	$VmHostCsvCheckBox.CheckState = "Checked" ;
	$VmHostCsvValidationComplete.Text = "" ;
	$VmCsvCheckBox.CheckState = "Checked" ;
	$VmCsvValidationComplete.Text = "" ;
	$TemplateCsvCheckBox.CheckState = "Checked" ;
	$TemplateCsvValidationComplete.Text = "" ;
	$DatastoreClusterCsvCheckBox.CheckState = "Checked" ;
	$DatastoreClusterCsvValidationComplete.Text = "" ;
	$DatastoreCsvCheckBox.CheckState = "Checked" ;
	$DatastoreCsvValidationComplete.Text = "" ;
	$VsSwitchCsvCheckBox.CheckState = "Checked" ;
	$VsSwitchCsvValidationComplete.Text = "" ;
	$VssPortGroupCsvCheckBox.CheckState = "Checked" ;
	$VssPortGroupCsvValidationComplete.Text = "" ;
	$VssVmkernelCsvCheckBox.CheckState = "Checked" ;
	$VssVmkernelCsvValidationComplete.Text = "" ;
	$VssPnicCsvCheckBox.CheckState = "Checked" ;
	$VssPnicCsvValidationComplete.Text = "" ;
	$VdSwitchCsvCheckBox.CheckState = "Checked" ;
	$VdSwitchCsvValidationComplete.Text = "" ;
	$VdsPortGroupCsvCheckBox.CheckState = "Checked" ;
	$VdsPortGroupCsvValidationComplete.Text = "" ;
	$VdsVmkernelCsvCheckBox.CheckState = "Checked" ;
	$VdsVmkernelCsvValidationComplete.Text = "" ;
	$VdsPnicCsvCheckBox.CheckState = "Checked" ;
	$VdsPnicCsvValidationComplete.Text = "" ;
	$FolderCsvCheckBox.CheckState = "Checked" ;
	$FolderCsvValidationComplete.Text = "" ;
	$RdmCsvCheckBox.CheckState = "Checked" ;
	$RdmCsvValidationComplete.Text = "" ;
	$DrsRuleCsvCheckBox.CheckState = "Checked" ;
	$DrsRuleCsvValidationComplete.Text = "" ;
	$DrsClusterGroupCsvCheckBox.CheckState = "Checked" ;
	$DrsClusterGroupCsvValidationComplete.Text = "" ;
	$DrsVmHostRuleCsvCheckBox.CheckState = "Checked" ;
	$DrsVmHostRuleCsvValidationComplete.Text = "" ;
	$ResourcePoolCsvValidationComplete.Text = "" ;
	$ResourcePoolCsvCheckBox.CheckState = "Checked" ;
	$SnapshotCsvValidationComplete.Text = "" ;
	$SnapshotCsvCheckBox.CheckState = "Checked" ;
	$LinkedvCenterCsvValidationComplete.Text = "" ;
	$LinkedvCenterCsvCheckBox.CheckState = "Checked" ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to vCenter"
})
#endregion ~~< CaptureOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton.Add_MouseClick({ Find_DrawCsvFolder ;
	if ($DrawCsvFolder -eq $null)
	{
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Red ;
		$DrawCsvInputButton.Text = "Folder Not Selected"
	}
	else
	{
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Green ;
		$DrawCsvInputButton.Text = $DrawCsvFolder
	}
} )
$TabDraw.Controls.Add($DrawCsvInputButton)
#endregion ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton.Add_Click(
{
	$CsvInputDir = $DrawCsvFolder+"\"+$VcenterTextBox.Text
	$vCenterExportFile = $CsvInputDir + "-vCenterExport.csv"
	$vCenterCsvExists = Test-Path $vCenterExportFile
	$TabDraw.Controls.Add($vCenterCsvValidationCheck)
	if ($vCenterCsvExists -eq $True)
	{
							
		$vCenterCsvValidationCheck.Forecolor = "Green"
		$vCenterCsvValidationCheck.Text = "Present"
	}
	else
	{
		$vCenterCsvValidationCheck.Forecolor = "Red"
		$vCenterCsvValidationCheck.Text = "Not Present"
	}
	
	$DatacenterExportFile = $CsvInputDir + "-DatacenterExport.csv"
	$DatacenterCsvExists = Test-Path $DatacenterExportFile
	$TabDraw.Controls.Add($DatacenterCsvValidationCheck)
			
	if ($DatacenterCsvExists -eq $True)
	{
		$DatacenterCsvValidationCheck.Forecolor = "Green"
		$DatacenterCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DatacenterCsvValidationCheck.Forecolor = "Red"
		$DatacenterCsvValidationCheck.Text = "Not Present"
	}
	
	$ClusterExportFile = $CsvInputDir + "-ClusterExport.csv"
	$ClusterCsvExists = Test-Path $ClusterExportFile
	$TabDraw.Controls.Add($ClusterCsvValidationCheck)
			
	if ($ClusterCsvExists -eq $True)
	{
		$ClusterCsvValidationCheck.Forecolor = "Green"
		$ClusterCsvValidationCheck.Text = "Present"
	}
	else
	{
		$ClusterCsvValidationCheck.Forecolor = "Red"
		$ClusterCsvValidationCheck.Text = "Not Present"
	}
			
	$VmHostExportFile = $CsvInputDir + "-VmHostExport.csv"
	$VmHostCsvExists = Test-Path $VmHostExportFile
	$TabDraw.Controls.Add($VmHostCsvValidationCheck)
			
	if ($VmHostCsvExists -eq $True)
	{
		$VmHostCsvValidationCheck.Forecolor = "Green"
		$VmHostCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VmHostCsvValidationCheck.Forecolor = "Red"
		$VmHostCsvValidationCheck.Text = "Not Present"
	}
			
	$VmExportFile = $CsvInputDir + "-VmExport.csv"
	$VmCsvExists = Test-Path $VmExportFile
	$TabDraw.Controls.Add($VmCsvValidationCheck)
			
	if ($VmCsvExists -eq $True)
	{
		$VmCsvValidationCheck.Forecolor = "Green"
		$VmCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VmCsvValidationCheck.Forecolor = "Red"
		$VmCsvValidationCheck.Text = "Not Present"
	}
			
	$TemplateExportFile = $CsvInputDir + "-ClusterExport.csv"
	$TemplateCsvExists = Test-Path $TemplateExportFile
	$TabDraw.Controls.Add($TemplateCsvValidationCheck)
			
	if ($TemplateCsvExists -eq $True)
	{
		$TemplateCsvValidationCheck.Forecolor = "Green"
		$TemplateCsvValidationCheck.Text = "Present"
	}
	else
	{
		$TemplateCsvValidationCheck.Forecolor = "Red"
		$TemplateCsvValidationCheck.Text = "Not Present"
	}
			
	$DatastoreClusterExportFile = $CsvInputDir + "-DatastoreClusterExport.csv"
	$DatastoreClusterCsvExists = Test-Path $DatastoreClusterExportFile
	$TabDraw.Controls.Add($DatastoreClusterCsvValidationCheck)
			
	if ($DatastoreClusterCsvExists -eq $True)
	{
		$DatastoreClusterCsvValidationCheck.Forecolor = "Green"
		$DatastoreClusterCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DatastoreClusterCsvValidationCheck.Forecolor = "Red"
		$DatastoreClusterCsvValidationCheck.Text = "Not Present"
	}
			
	$DatastoreExportFile = $CsvInputDir + "-DatastoreExport.csv"
	$DatastoreCsvExists = Test-Path $DatastoreExportFile
	$TabDraw.Controls.Add($DatastoreCsvValidationCheck)
			
	if ($DatastoreCsvExists -eq $True)
	{
		$DatastoreCsvValidationCheck.Forecolor = "Green"
		$DatastoreCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DatastoreCsvValidationCheck.Forecolor = "Red"
		$DatastoreCsvValidationCheck.Text = "Not Present"
	}
			
	$VsSwitchExportFile = $CsvInputDir + "-VsSwitchExport.csv"
	$VsSwitchCsvExists = Test-Path $VsSwitchExportFile
	$TabDraw.Controls.Add($VsSwitchCsvValidationCheck)
			
	if ($VsSwitchCsvExists -eq $True)
	{
		$VsSwitchCsvValidationCheck.Forecolor = "Green"
		$VsSwitchCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VsSwitchCsvValidationCheck.Forecolor = "Red"
		$VsSwitchCsvValidationCheck.Text = "Not Present"
		$VSS_to_Host_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssPortGroupExportFile = $CsvInputDir + "-VssPortGroupExport.csv"
	$VssPortGroupCsvExists = Test-Path $VssPortGroupExportFile
	$TabDraw.Controls.Add($VssPortGroupCsvValidationCheck)
			
	if ($VssPortGroupCsvExists -eq $True)
	{
		$VssPortGroupCsvValidationCheck.Forecolor = "Green"
		$VssPortGroupCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VssPortGroupCsvValidationCheck.Forecolor = "Red"
		$VssPortGroupCsvValidationCheck.Text = "Not Present"
		$VSSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssVmkernelExportFile = $CsvInputDir + "-VssVmkernelExport.csv"
	$VssVmkernelCsvExists = Test-Path $VssVmkernelExportFile
	$TabDraw.Controls.Add($VssVmkernelCsvValidationCheck)
			
	if ($VssVmkernelCsvExists -eq $True)
	{
		$VssVmkernelCsvValidationCheck.Forecolor = "Green"
		$VssVmkernelCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VssVmkernelCsvValidationCheck.Forecolor = "Red"
		$VssVmkernelCsvValidationCheck.Text = "Not Present"
		$VMK_to_VSS_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssPnicExportFile = $CsvInputDir + "-VssPnicExport.csv"
	$VssPnicCsvExists = Test-Path $VssPnicExportFile
	$TabDraw.Controls.Add($VssPnicCsvValidationCheck)
			
	if ($VssPnicCsvExists -eq $True)
	{
		$VssPnicCsvValidationCheck.Forecolor = "Green"
		$VssPnicCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VssPnicCsvValidationCheck.Forecolor = "Red"
		$VssPnicCsvValidationCheck.Text = "Not Present"
	}
			
	$VdSwitchExportFile = $CsvInputDir + "-VdSwitchExport.csv"
	$VdSwitchCsvExists = Test-Path $VdSwitchExportFile
	$TabDraw.Controls.Add($VdSwitchCsvValidationCheck)
			
	if ($VdSwitchCsvExists -eq $True)
	{
		$VdSwitchCsvValidationCheck.Forecolor = "Green"
		$VdSwitchCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VdSwitchCsvValidationCheck.Forecolor = "Red"
		$VdSwitchCsvValidationCheck.Text = "Not Present"
		$VDS_to_Host_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsPortGroupExportFile = $CsvInputDir + "-VdsPortGroupExport.csv"
	$VdsPortGroupCsvExists = Test-Path $VdsPortGroupExportFile
	$TabDraw.Controls.Add($VdsPortGroupCsvValidationCheck)
			
	if ($VdsPortGroupCsvExists -eq $True)
	{
		$VdsPortGroupCsvValidationCheck.Forecolor = "Green"
		$VdsPortGroupCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VdsPortGroupCsvValidationCheck.Forecolor = "Red"
		$VdsPortGroupCsvValidationCheck.Text = "Not Present"
		$VDSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsVmkernelExportFile = $CsvInputDir + "-VdsVmkernelExport.csv"
	$VdsVmkernelCsvExists = Test-Path $VdsVmkernelExportFile
	$TabDraw.Controls.Add($VdsVmkernelCsvValidationCheck)
			
	if ($VdsVmkernelCsvExists -eq $True)
	{
		$VdsVmkernelCsvValidationCheck.Forecolor = "Green"
		$VdsVmkernelCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VdsVmkernelCsvValidationCheck.Forecolor = "Red"
		$VdsVmkernelCsvValidationCheck.Text = "Not Present"
		$VMK_to_VDS_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsPnicExportFile = $CsvInputDir + "-VdsPnicExport.csv"
	$VdsPnicCsvExists = Test-Path $VdsPnicExportFile
	$TabDraw.Controls.Add($VdsPnicCsvValidationCheck)
			
	if ($VdsPnicCsvExists -eq $True)
	{
		$VdsPnicCsvValidationCheck.Forecolor = "Green"
		$VdsPnicCsvValidationCheck.Text = "Present"
	}
	else
	{
		$VdsPnicCsvValidationCheck.Forecolor = "Red"
		$VdsPnicCsvValidationCheck.Text = "Not Present"
	}
			
	$FolderExportFile = $CsvInputDir + "-FolderExport.csv"
	$FolderCsvExists = Test-Path $FolderExportFile
	$TabDraw.Controls.Add($FolderCsvValidationCheck)
			
	if ($FolderCsvExists -eq $True)
	{
		$FolderCsvValidationCheck.Forecolor = "Green"
		$FolderCsvValidationCheck.Text = "Present"
	}
	else
	{
		$FolderCsvValidationCheck.Forecolor = "Red"
		$FolderCsvValidationCheck.Text = "Not Present"
	}
			
	$RdmExportFile = $CsvInputDir + "-RdmExport.csv"
	$RdmCsvExists = Test-Path $RdmExportFile
	$TabDraw.Controls.Add($RdmCsvValidationCheck)
			
	if ($RdmCsvExists -eq $True)
	{
		$RdmCsvValidationCheck.Forecolor = "Green"
		$RdmCsvValidationCheck.Text = "Present"
	}
	else
	{
		$RdmCsvValidationCheck.Forecolor = "Red"
		$RdmCsvValidationCheck.Text = "Not Present"
		$VMs_with_RDMs_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$DrsRuleExportFile = $CsvInputDir + "-DrsRuleExport.csv"
	$DrsRuleCsvExists = Test-Path $DrsRuleExportFile
	$TabDraw.Controls.Add($DrsRuleCsvValidationCheck)
			
	if ($DrsRuleCsvExists -eq $True)
	{
		$DrsRuleCsvValidationCheck.Forecolor = "Green"
		$DrsRuleCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DrsRuleCsvValidationCheck.Forecolor = "Red"
		$DrsRuleCsvValidationCheck.Text = "Not Present"
		$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$DrsClusterGroupExportFile = $CsvInputDir + "-DrsClusterGroupExport.csv"
	$DrsClusterGroupCsvExists = Test-Path $DrsClusterGroupExportFile
	$TabDraw.Controls.Add($DrsClusterGroupCsvValidationCheck)
			
	if ($DrsClusterGroupCsvExists -eq $True)
	{
		$DrsClusterGroupCsvValidationCheck.Forecolor = "Green"
		$DrsClusterGroupCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DrsClusterGroupCsvValidationCheck.Forecolor = "Red"
		$DrsClusterGroupCsvValidationCheck.Text = "Not Present"
	}
			
	$DrsVmHostRuleExportFile = $CsvInputDir + "-DrsVmHostRuleExport.csv"
	$DrsVmHostRuleCsvExists = Test-Path $DrsVmHostRuleExportFile
	$TabDraw.Controls.Add($DrsVmHostRuleCsvValidationCheck)
			
	if ($DrsVmHostRuleCsvExists -eq $True)
	{
		$DrsVmHostRuleCsvValidationCheck.Forecolor = "Green"
		$DrsVmHostRuleCsvValidationCheck.Text = "Present"
	}
	else
	{
		$DrsVmHostRuleCsvValidationCheck.Forecolor = "Red"
		$DrsVmHostRuleCsvValidationCheck.Text = "Not Present"
	}
			
	$ResourcePoolExportFile = $CsvInputDir + "-ResourcePoolExport.csv"
	$ResourcePoolCsvExists = Test-Path $ResourcePoolExportFile
	$TabDraw.Controls.Add($ResourcePoolCsvValidationCheck)
			
	if ($ResourcePoolCsvExists -eq $True)
	{
		$ResourcePoolCsvValidationCheck.Forecolor = "Green"
		$ResourcePoolCsvValidationCheck.Text = "Present"
	}
	else
	{
		$ResourcePoolCsvValidationCheck.Forecolor = "Red"
		$ResourcePoolCsvValidationCheck.Text = "Not Present"
		$VM_to_ResourcePool_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$SnapshotExportFile = $CsvInputDir + "-SnapshotExport.csv"
	$SnapshotCsvExists = Test-Path $SnapshotExportFile
	$TabDraw.Controls.Add($SnapshotCsvValidationCheck)
			
	if ($SnapshotCsvExists -eq $True)
	{
		$SnapshotCsvValidationCheck.Forecolor = "Green"
		$SnapshotCsvValidationCheck.Text = "Present"
	}
	else
	{
		$SnapshotCsvValidationCheck.Forecolor = "Red"
		$SnapshotCsvValidationCheck.Text = "Not Present"
		$Snapshot_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
	
	$LinkedvCenterExportFile = $CsvInputDir + "-LinkedvCenterExport.csv"
	$LinkedvCenterCsvExists = Test-Path $LinkedvCenterExportFile
	$TabDraw.Controls.Add($LinkedvCenterCsvValidationCheck)
			
	if ($LinkedvCenterCsvExists -eq $True)
	{
		$LinkedvCenterCsvValidationCheck.Forecolor = "Green"
		$LinkedvCenterCsvValidationCheck.Text = "Present"
	}
	else
	{
		$LinkedvCenterCsvValidationCheck.Forecolor = "Red"
		$LinkedvCenterCsvValidationCheck.Text = "Not Present"
		$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = "UnChecked"
	}
} )
$CsvValidationButton.Add_MouseClick({ $CsvValidationButton.Forecolor = [System.Drawing.Color]::Green ; $CsvValidationButton.Text = "CSV Validation Complete" })
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton.Add_MouseClick({Find_DrawVisioFolder; 
	if($VisioFolder -eq $null)
	{
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$VisioOpenOutputButton.Text = "Folder Not Selected"
	}
	else
	{
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$VisioOpenOutputButton.Text = $VisioFolder
	}
} )
#endregion ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton.Add_Click( { $vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = "UnChecked" ;
	$VM_to_Host_DrawCheckBox.CheckState = "UnChecked" ;
	$VM_to_Folder_DrawCheckBox.CheckState = "UnChecked" ;
	$VMs_with_RDMs_DrawCheckBox.CheckState = "UnChecked" ;
	$SRM_Protected_VMs_DrawCheckBox.CheckState = "UnChecked" ;
	$VM_to_Datastore_DrawCheckBox.CheckState = "UnChecked" ;
	$VM_to_ResourcePool_DrawCheckBox.CheckState = "UnChecked" ;
	$Datastore_to_Host_DrawCheckBox.CheckState = "UnChecked" ;
	$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = "UnChecked" ;
	$VSS_to_Host_DrawCheckBox.CheckState = "UnChecked" ;
	$VMK_to_VSS_DrawCheckBox.CheckState = "UnChecked" ;
	$VSSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked" ;
	$VDS_to_Host_DrawCheckBox.CheckState = "UnChecked" ;
	$VMK_to_VDS_DrawCheckBox.CheckState = "UnChecked" ;
	$VDSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked" ;
	$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = "UnChecked";
	$Snapshot_to_VM_DrawCheckBox.CheckState = "UnChecked"
} )
#endregion ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton.Add_Click( { $vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Folder_DrawCheckBox.CheckState = "Checked" ;
	$VMs_with_RDMs_DrawCheckBox.CheckState = "Checked" ;
	$SRM_Protected_VMs_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Datastore_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_ResourcePool_DrawCheckBox.CheckState = "Checked" ;
	$Datastore_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = "Checked" ;
	$VSS_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VMK_to_VSS_DrawCheckBox.CheckState = "Checked" ;
	$VSSPortGroup_to_VM_DrawCheckBox.CheckState = "Checked" ;
	$VDS_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VMK_to_VDS_DrawCheckBox.CheckState = "Checked" ;
	$VDSPortGroup_to_VM_DrawCheckBox.CheckState = "Checked" ;
	$VdsVmkernelCsvCheckBox.CheckState = "Checked" ;
	$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = "Checked";
	$Snapshot_to_VM_DrawCheckBox.CheckState = "Checked"
} )
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton.Add_Click({
if($VisioFolder -eq $null)
{
	$DrawButton.Forecolor = [System.Drawing.Color]::Red ;
	$DrawButton.Text = "Folder Not Selected"
}
else
{
$DrawButton.Forecolor = [System.Drawing.Color]::Blue ;
$DrawButton.Text = "Drawing Please Wait" ;
Create_Visio_Base;
if ($vCenter_to_LinkedvCenter_DrawCheckBox.Checked -eq "True")
{
	$vCenter_to_LinkedvCenter_Complete.Forecolor = "Blue"
	$vCenter_to_LinkedvCenter_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)
	vCenter_to_LinkedvCenter
	$vCenter_to_LinkedvCenter_Complete.Forecolor = "Green"
	$vCenter_to_LinkedvCenter_Complete.Text = "Complete"
};
if ($VM_to_Host_DrawCheckBox.Checked -eq "True")
{
	$VM_to_Host_Complete.Forecolor = "Blue"
	$VM_to_Host_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VM_to_Host_Complete)
	VM_to_Host
	$VM_to_Host_Complete.Forecolor = "Green"
	$VM_to_Host_Complete.Text = "Complete"
	$TabDraw.Controls.Add($VM_to_Host_Complete)
}
if ($VM_to_Folder_DrawCheckBox.Checked -eq "True")
{
	$VM_to_Folder_Complete.Forecolor = "Blue"
	$VM_to_Folder_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VM_to_Folder_Complete)
	VM_to_Folder
	$VM_to_Folder_Complete.Forecolor = "Green"
	$VM_to_Folder_Complete.Text = "Complete"
}
if ($VMs_with_RDMs_DrawCheckBox.Checked -eq "True")
{
	$VMs_with_RDMs_Complete.Forecolor = "Blue"
	$VMs_with_RDMs_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VMs_with_RDMs_Complete)
	VMs_with_RDMs
	$VMs_with_RDMs_Complete.Forecolor = "Green"
	$VMs_with_RDMs_Complete.Text = "Complete"
}
if ($SRM_Protected_VMs_DrawCheckBox.Checked -eq "True")
{
	$SRM_Protected_VMs_Complete.Forecolor = "Blue"
	$SRM_Protected_VMs_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)
	SRM_Protected_VMs
	$SRM_Protected_VMs_Complete.Forecolor = "Green"
	$SRM_Protected_VMs_Complete.Text = "Complete"
}
if ($VM_to_Datastore_DrawCheckBox.Checked -eq "True")
{
	$VM_to_Datastore_Complete.Forecolor = "Blue"
	$VM_to_Datastore_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VM_to_Datastore_Complete)
	VM_to_Datastore
	$VM_to_Datastore_Complete.Forecolor = "Green"
	$VM_to_Datastore_Complete.Text = "Complete"
}
if ($VM_to_ResourcePool_DrawCheckBox.Checked -eq "True")
{
	$VM_to_ResourcePool_Complete.Forecolor = "Blue"
	$VM_to_ResourcePool_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)
	VM_to_ResourcePool
	$VM_to_ResourcePool_Complete.Forecolor = "Green"
	$VM_to_ResourcePool_Complete.Text = "Complete"
}
if ($Datastore_to_Host_DrawCheckBox.Checked -eq "True")
{
	$Datastore_to_Host_Complete.Forecolor = "Blue"
	$Datastore_to_Host_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($Datastore_to_Host_Complete)
	Datastore_to_Host
	$Datastore_to_Host_Complete.Forecolor = "Green"
	$Datastore_to_Host_Complete.Text = "Complete"
}
if ($Snapshot_to_VM_DrawCheckBox.Checked -eq "True")
{
	$Snapshot_to_VM_Complete.Forecolor = "Blue"
	$Snapshot_to_VM_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($Snapshot_to_VM_Complete)
	Snapshot_to_VM
	$Snapshot_to_VM_Complete.Forecolor = "Green"
	$Snapshot_to_VM_Complete.Text = "Complete"
};
if ($PhysicalNIC_to_vSwitch_DrawCheckBox.Checked -eq "True")
{
	$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
	$PhysicalNIC_to_vSwitch_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)
	PhysicalNIC_to_vSwitch
	$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Green"
	$PhysicalNIC_to_vSwitch_Complete.Text = "Complete"
}
if ($VSS_to_Host_DrawCheckBox.Checked -eq "True")
{
	$VSS_to_Host_Complete.Forecolor = "Blue"
	$VSS_to_Host_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VSS_to_Host_Complete)
	VSS_to_Host
	$VSS_to_Host_Complete.Forecolor = "Green"
	$VSS_to_Host_Complete.Text = "Complete"
}
if ($VMK_to_VSS_DrawCheckBox.Checked -eq "True")
{
	$VMK_to_VSS_Complete.Forecolor = "Blue"
	$VMK_to_VSS_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VMK_to_VSS_Complete)
	VMK_to_VSS
	$VMK_to_VSS_Complete.Forecolor = "Green"
	$VMK_to_VSS_Complete.Text = "Complete"
}
if ($VSSPortGroup_to_VM_DrawCheckBox.Checked -eq "True")
{
	$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
	$VSSPortGroup_to_VM_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)
	VSSPortGroup_to_VM
	$VSSPortGroup_to_VM_Complete.Forecolor = "Green"
	$VSSPortGroup_to_VM_Complete.Text = "Complete"
}
if ($VDS_to_Host_DrawCheckBox.Checked -eq "True")
{
	$VDS_to_Host_Complete.Forecolor = "Blue"
	$VDS_to_Host_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VDS_to_Host_Complete)
	VDS_to_Host
	$VDS_to_Host_Complete.Forecolor = "Green"
	$VDS_to_Host_Complete.Text = "Complete"
}
if ($VMK_to_VDS_DrawCheckBox.Checked -eq "True")
{
	$VMK_to_VDS_Complete.Forecolor = "Blue"
	$VMK_to_VDS_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VMK_to_VDS_Complete)
	VMK_to_VDS
	$VMK_to_VDS_Complete.Forecolor = "Green"
	$VMK_to_VDS_Complete.Text = "Complete"
}
if ($VDSPortGroup_to_VM_DrawCheckBox.Checked -eq "True")
{
	$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
	$VDSPortGroup_to_VM_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)
	VDSPortGroup_to_VM
	$VDSPortGroup_to_VM_Complete.Forecolor = "Green"
	$VDSPortGroup_to_VM_Complete.Text = "Complete"
}
if ($Cluster_to_DRS_Rule_DrawCheckBox.Checked -eq "True")
{
	$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
	$Cluster_to_DRS_Rule_Complete.Text = "Processing ..."
	$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
	Cluster_to_DRS_Rule
	$Cluster_to_DRS_Rule_Complete.Forecolor = "Green"
	$Cluster_to_DRS_Rule_Complete.Text = "Complete"
};
$DrawButton.Forecolor = [System.Drawing.Color]::Green; $DrawButton.Text = "Visio Drawings Complete"}})
#endregion ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton.Add_Click({Open_Final_Visio ;
	$VcenterTextBox.Text = "" ;
	$UserNameTextBox.Text = "" ;
	$PasswordTextBox.Text = "" ;
	$PasswordTextBox.UseSystemPasswordChar = $true ;
	$ConnectButton.Forecolor = [System.Drawing.Color]::Black ;
	$ConnectButton.Text = "Connect to vCenter" ;
	$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Black ;
	$DrawCsvInputButton.Text = "Select CSV Input Folder" ;
	$vCenterCsvValidationCheck.Text = "" ;
	$DatacenterCsvValidationCheck.Text = "" ;
	$ClusterCsvValidationCheck.Text = "" ;
	$VmHostCsvValidationCheck.Text = "" ;
	$VmCsvValidationCheck.Text = "" ;
	$TemplateCsvValidationCheck.Text = "" ;
	$DatastoreClusterCsvValidationCheck.Text = "" ;
	$DatastoreCsvValidationCheck.Text = "" ;
	$VsSwitchCsvValidationCheck.Text = "" ;
	$VssPortGroupCsvValidationCheck.Text = "" ;
	$VssVmkernelCsvValidationCheck.Text = "" ;
	$VssPnicCsvValidationCheck.Text = "" ;
	$VdSwitchCsvValidationCheck.Text = "" ;
	$VdsPortGroupCsvValidationCheck.Text = "" ;
	$VdsVmkernelCsvValidationCheck.Text = "" ;
	$VdsPnicCsvValidationCheck.Text = "" ;
	$FolderCsvValidationCheck.Text = "" ;
	$RdmCsvValidationCheck.Text = "" ;
	$DrsRuleCsvValidationCheck.Text = "" ;
	$DrsClusterGroupCsvValidationCheck.Text = "" ;
	$DrsVmHostRuleCsvValidationCheck.Text = "" ;
	$ResourcePoolCsvValidationCheck.Text = "" ;
	$LinkedvCenterCsvValidationCheck.Text = "" ;
	$SnapshotCsvValidationCheck.Text = "" ;
	$CsvValidationButton.Forecolor = [System.Drawing.Color]::Black ;
	$CsvValidationButton.Text = "Check for CSVs" ;
	$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Black ;
	$VisioOpenOutputButton.Text = "Select Visio Output Folder" ;
	$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = "Checked" ;
	$vCenter_to_LinkedvCenter_Complete.Text = "" ;
	$VM_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Host_Complete.Text = "" ;
	$VM_to_Folder_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Folder_Complete.Text = "" ;
	$VMs_with_RDMs_DrawCheckBox.CheckState = "Checked" ;
	$VMs_with_RDMs_Complete.Text = "" ;
	$SRM_Protected_VMs_DrawCheckBox.CheckState = "Checked" ;
	$SRM_Protected_VMs_Complete.Text = "" ;
	$VM_to_Datastore_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_Datastore_Complete.Text = "" ;
	$VM_to_ResourcePool_DrawCheckBox.CheckState = "Checked" ;
	$VM_to_ResourcePool_Complete.Text = "" ;
	$Datastore_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$Datastore_to_Host_Complete.Text = "" ;
	$Snapshot_to_VM_DrawCheckBox.CheckState = "Checked" ;
	$Snapshot_to_VM_Complete.Text = "" ;
	$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = "Checked" ;
	$PhysicalNIC_to_vSwitch_Complete.Text = "" ;
	$VSS_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VSS_to_Host_Complete.Text = "" ;
	$VMK_to_VSS_DrawCheckBox.CheckState = "Checked" ;
	$VMK_to_VSS_Complete.Text = "" ;
	$VSSPortGroup_to_VM_DrawCheckBox.CheckState = "Checked" ;
	$VSSPortGroup_to_VM_Complete.Text = "" ;
	$VDS_to_Host_DrawCheckBox.CheckState = "Checked" ;
	$VDS_to_Host_Complete.Text = "" ;
	$VMK_to_VDS_DrawCheckBox.CheckState = "Checked" ;
	$VMK_to_VDS_Complete.Text = "" ;
	$VDSPortGroup_to_VM_DrawCheckBox.CheckState = "Checked" ;
	$VDSPortGroup_to_VM_Complete.Text = "" ;
	$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = "Checked" ;
	$Cluster_to_DRS_Rule_Complete.Text = "" ;
	$DrawButton.Forecolor = [System.Drawing.Color]::Black ;
	$DrawButton.Text = "Draw Visio"
} )
#endregion ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Event Loop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Main
{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($vDiagram)
}
#endregion ~~< Event Loop >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Event Handlers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenter Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Connect_vCenter
{
	$global:vCenter = $VcenterTextBox.Text
	$User = $UserNameTextBox.Text
	Connect-VIServer $Vcenter -user $User -password $PasswordTextBox.Text
}
#endregion ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Disconnect_vCenter
{
	$Disconnect = Disconnect-ViServer * -Confirm:$false
}
#endregion ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< vCenter Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Folder Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Find_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_CaptureCsvFolder
{
	$CaptureCsvBrowseLoop = $True
	while ($CaptureCsvBrowseLoop)
	{
		if ($CaptureCsvBrowse.ShowDialog() -eq "OK")
		{
			$CaptureCsvBrowseLoop = $False
		}
		else
		{
			$CaptureCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
			if ($CaptureCsvBrowseRes -eq "Cancel")
			{
				return
			}
		}
	}
	$global:CaptureCsvFolder = $CaptureCsvBrowse.SelectedPath
}
#endregion ~~< Find_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Check_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Check_CaptureCsvFolder
{
	$CheckContentPath = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
	$CheckContentDir = $CheckContentPath + "*.csv"
	$CheckContent = Test-Path $CheckContentDir
	if ($CheckContent -eq "True")
	{
		$CheckContents_CaptureCsvFolder =  [System.Windows.MessageBox]::Show("Files where found in the folder. Would you like to delete these files? Click 'Yes' to delete and 'No' move files to a new folder.","Warning!","YesNo","Error")
		switch  ($CheckContents_CaptureCsvFolder) { 
		'Yes' 
		{
		del $CheckContentDir
		}
		
		'No'
		{
		$CheckContentCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
		$CheckContentCsvBrowse.Description = "Select a directory to copy files to"
		$CheckContentCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
		$CheckContentCsvBrowse.ShowDialog()
		$global:NewContentCsvFolder = $CheckContentCsvBrowse.SelectedPath
		copy-item -Path $CheckContentDir -Destination $NewContentCsvFolder
		del $CheckContentDir
		}
	}
  }
}
#endregion ~~< Check_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Find_DrawCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_DrawCsvFolder
{
	$DrawCsvBrowseLoop = $True
	while ($DrawCsvBrowseLoop)
	{
		if ($DrawCsvBrowse.ShowDialog() -eq "OK")
		{
			$DrawCsvBrowseLoop = $False
		}
		else
		{
			$DrawCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
			if ($DrawCsvBrowseRes -eq "Cancel")
			{
				return
			}
		}
	}
	$global:DrawCsvFolder = $DrawCsvBrowse.SelectedPath
}
#endregion ~~< Find_DrawCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Find_DrawVisioFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_DrawVisioFolder
{
	$VisioBrowseLoop = $True
	while($VisioBrowseLoop)
	{
		if ($VisioBrowse.ShowDialog() -eq "OK")
		{
			$VisioBrowseLoop = $False
		}
		else
		{
			$VisioBrowseRes = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
			if($VisioBrowseRes -eq "Cancel")
			{
				return
			}
		}
	}
	$global:VisioFolder = $VisioBrowse.SelectedPath
}
#endregion ~~< Find_DrawVisioFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Folder Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Export Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function vCenter_Export
{
	$vCenterExportFile = "$CaptureCsvFolder\$vCenter-vCenterExport.csv"
	$global:DefaultVIServer | 
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } }, 
			@{ Name = "Version" ; Expression = { $_.Version } }, 
			@{ Name = "Build" ; Expression = { $_.Build } },
			@{ Name = "OsType" ; Expression = { $_.ExtensionData.Content.About.OsType } },
			@{ Name = "IsConnected" ; Expression = { $_.IsConnected } },
			@{ Name = "ServiceUri" ; Expression = { $_.ServiceUri } },
			@{ Name = "Port" ; Expression = { $_.Port } },
			@{ Name = "ProductLine" ; Expression = { $_.ProductLine } },
			@{ Name = "InstanceUuid" ; Expression = { $_.InstanceUuid } },
			@{ Name = "RefCount" ; Expression = { $_.RefCount } },
			@{ Name = "ExtensionData_ServerClock" ; Expression = { $_.ExtensionData.ServerClock } },
			@{ Name = "ExtensionData_Capability_ProvisioningSupported" ; Expression = { $_.ExtensionData.Capability.ProvisioningSupported } },
			@{ Name = "ExtensionData_Capability_MultiHostSupported" ; Expression = { $_.ExtensionData.Capability.MultiHostSupported } },
			@{ Name = "ExtensionData_Capability_UserShellAccessSupported" ; Expression = { $_.ExtensionData.Capability.UserShellAccessSupported } },
			@{ Name = "ExtensionData_Capability_NetworkBackupAndRestoreSupported" ; Expression = { $_.ExtensionData.Capability.NetworkBackupAndRestoreSupported } },
			@{ Name = "ExtensionData_Capability_FtDrsWithoutEvcSupported" ; Expression = { $_.ExtensionData.Capability.FtDrsWithoutEvcSupported } },
			@{ Name = "ExtensionData_Capability_HciWorkflowSupported" ; Expression = { $_.ExtensionData.Capability.HciWorkflowSupported } },
			@{ Name = "ExtensionData_Content_RootFolder" ; Expression = { Get-Folder -Id ( $_.ExtensionData.Content.RootFolder ) } },
			@{ Name = "ExtensionData_Content_About_Name" ; Expression = { $_.ExtensionData.Content.About.Name } },
			@{ Name = "ExtensionData_Content_About_FullName" ; Expression = { $_.ExtensionData.Content.About.FullName } },
			@{ Name = "ExtensionData_Content_About_Vendor" ; Expression = { $_.ExtensionData.Content.About.Vendor } },
			@{ Name = "ExtensionData_Content_About_Version" ; Expression = { $_.ExtensionData.Content.About.Version } },
			@{ Name = "ExtensionData_Content_About_Build" ; Expression = { $_.ExtensionData.Content.About.Build } },
			@{ Name = "ExtensionData_Content_About_LocaleVersion" ; Expression = { $_.ExtensionData.Content.About.LocaleVersion } },
			@{ Name = "ExtensionData_Content_About_LocaleBuild" ; Expression = { $_.ExtensionData.Content.About.LocaleBuild } },
			@{ Name = "ExtensionData_Content_About_OsType" ; Expression = { $_.ExtensionData.Content.About.OsType } },
			@{ Name = "ExtensionData_Content_About_ProductLineId" ; Expression = { $_.ExtensionData.Content.About.ProductLineId } },
			@{ Name = "ExtensionData_Content_About_ApiType" ; Expression = { $_.ExtensionData.Content.About.ApiType } },
			@{ Name = "ExtensionData_Content_About_ApiVersion" ; Expression = { $_.ExtensionData.Content.About.ApiVersion } },
			@{ Name = "ExtensionData_Content_About_InstanceUuid" ; Expression = { $_.ExtensionData.Content.About.InstanceUuid } },
			@{ Name = "ExtensionData_Content_About_LicenseProductName" ; Expression = { $_.ExtensionData.Content.About.LicenseProductName } },
			@{ Name = "ExtensionData_Content_About_LicenseProductVersion" ; Expression = { $_.ExtensionData.Content.About.LicenseProductVersion } } |
		Export-Csv $vCenterExportFile -Append -NoTypeInformation
}
#endregion ~~< vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datacenter_Export
{
	$DatacenterExportFile = "$CaptureCsvFolder\$vCenter-DatacenterExport.csv"
	Get-View -ViewType Datacenter | 
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } },
			@{ Name = "VmFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Type VM | Where-Object { $_.MoRef -eq $_.VmFolder } | Sort-Object Name  ) ) } },
			@{ Name = "HostFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Type HostAndCluster | Where-Object { $_.MoRef -eq $_.HostFolder } | Sort-Object Name ) ) } },
			@{ Name = "DatastoreFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Datastore | Where-Object { $_.MoRef -eq $_.DatastoreFolder } | Sort-Object Name ) ) } },
			@{ Name = "NetworkFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Network | Where-Object { $_.MoRef -eq $_.NetworkFolder } | Sort-Object Name ) ) } },
			@{ Name = "Datastore" ; Expression = { [string]::Join(", ", ( Get-Datastore -Id $_.Datastore | Sort-Object Name ) ) } },
			@{ Name = "Network" ; Expression = { [string]::Join(", ", ( Get-VirtualPortGroup | Where-Object { $_.MoRef -eq $_.Network } | Sort-Object Name ) ) } },
			@{ Name = "Parent" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Datacenter | Where-Object { $_.MoRef -eq $_.Parent } | Sort-Object Name ) ) } },
			@{ Name = "OverallStatus" ; Expression = { [string]::Join(", ", ( $_.OverallStatus ) ) } },
			@{ Name = "ConfigStatus" ; Expression = { [string]::Join(", ", ( $_.ConfigStatus ) ) } },
			@{ Name = "ConfigIssue" ; Expression = { [string]::Join( ", ", ( $_.ConfigIssue ) ) } },
			@{ Name = "EffectiveRole" ; Expression = { [string]::Join( ", ", ( $_.EffectiveRole ) ) } },
			@{ Name = "AlarmActionsEnabled" ; Expression = { [string]::Join(", ", ( $_.AlarmActionsEnabled ) ) } },
			@{ Name = "MoRef" ; Expression = { [string]::Join(", ", ( $_.MoRef ) ) } } | 
		Export-Csv $DatacenterExportFile -Append -NoTypeInformation
}
#endregion ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Cluster_Export
{
	$ClusterExportFile = "$CaptureCsvFolder\$vCenter-ClusterExport.csv"
	Get-View -ViewType ClusterComputeResource | 
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } },
			@{ Name = "Datacenter" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -Cluster $_.Name ) ) } },
			@{ Name = "HAEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster -Id $_.MoRef ).HAEnabled ) ) } },
			@{ Name = "HAAdmissionControlEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).HAAdmissionControlEnabled ) ) } },
			@{ Name = "AdmissionControlPolicyCpuFailoverResourcesPercent" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.AdmissionControlPolicy.CpuFailoverResourcesPercent ) ) } },
			@{ Name = "AdmissionControlPolicyMemoryFailoverResourcesPercent" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.MemoryFailoverResourcesPercent ) ) } },
			@{ Name = "AdmissionControlPolicyFailoverLevel" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.FailoverLevel ) ) } },
			@{ Name = "AdmissionControlPolicyAutoComputePercentages" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.AutoComputePercentages ) ) } },
			@{ Name = "AdmissionControlPolicyResourceReductionToToleratePercent" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.ResourceReductionToToleratePercent ) ) } },
			@{ Name = "DrsEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).DrsEnabled ) ) } },
			@{ Name = "DrsAutomationLevel" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).DrsAutomationLevel ) ) } },
			@{ Name = "VmMonitoring" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.VmMonitoring ) ) } },
			@{ Name = "HostMonitoring" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.HostMonitoring ) ) } },
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } | 
		Export-Csv $ClusterExportFile -Append -NoTypeInformation
}
#endregion ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VmHost_Export
{
	$VmHostExportFile = "$CaptureCsvFolder\$vCenter-VmHostExport.csv"
	$ServiceInstance = Get-View ServiceInstance
	$LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
	$LicenseAssignmentManager = Get-View $LicenseManager.LicenseAssignmentManager
	Get-View -ViewType HostSystem | 
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", (  $_.Name ) ) } },
            @{	N = "Datacenter" ; E = { $Datacenter = Get-View -Id $_.Parent -Property Name, Parent
				while ($Datacenter -isnot [VMware.Vim.Datacenter] -and $Datacenter.Parent)
				{
					$Datacenter = Get-View -Id $Datacenter.Parent -Property Name, Parent
				}
				if ($Datacenter -is [VMware.Vim.Datacenter])
				{
				$Datacenter.Name } } }, 
			@{ N = "Cluster" ; E = { $Cluster = Get-View -Id $_.Parent -Property Name, Parent
				while ($Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent)
				{
					$Cluster = Get-View -Id $Cluster.Parent -Property Name, Parent
				}
				if ($Cluster -is [VMware.Vim.ClusterComputeResource]) { $Cluster.Name } } },
			@{ Name = "Vm" ; Expression = { [string]::Join( ", ", ( Get-VM -Id $_.Vm | Sort-Object Name ) ) } },
            @{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( ( Get-Datastore -Id $_.Datastore | Sort-Object Name ) ) ) } },
            @{ Name = "Version" ; Expression = { $_.Config.Product.Version } },
		    @{ Name = "Build" ; Expression = { $_.Config.Product.Build } },
		    @{ Name = "Manufacturer" ; Expression = { $_.Summary.Hardware.Vendor } },
		    @{ Name = "Model" ; Expression = { $_.Summary.Hardware.Model } },
		    @{ Name = "LicenseType" ; Expression = { $LicenseAssignmentManager.QueryAssignedLicenses($_.Config.Host.Value).AssignedLicense.Name  } },
		    @{ Name = "BiosVersion" ; Expression = { ( Get-VMHost $_.Name | Get-VMHostHardware -WaitForAllData -SkipAllSslCertificateChecks -ErrorAction SilentlyContinue ).BiosVersion } },
		    @{ Name = "BIOSReleaseDate" ; Expression = { ( ( ( Get-VMHost $_.Name ).ExtensionData.Hardware.BiosInfo.ReleaseDate -split " " )[0] ) } },
		    @{ Name = "ProcessorType" ; Expression = { $_.Summary.Hardware.CpuModel } },
		    @{ Name = "CpuMhz" ; Expression = { $_.Summary.Hardware.CpuMhz } },
		    @{ Name = "NumCpuPkgs" ; Expression = { $_.Summary.Hardware.NumCpuPkgs } },
		    @{ Name = "NumCpuCores" ; Expression = { $_.Summary.Hardware.NumCpuCores } },
		    @{ Name = "NumCpuThreads" ; Expression = { $_.Summary.Hardware.NumCpuThreads } },
		    @{ Name = "Memory" ; Expression = { [math]::Round([decimal]$_.Summary.Hardware.MemorySize / 1073741824) } },
		    @{ Name = "MaxEVCMode" ; Expression = { $_.Summary.MaxEVCModeKey } },
		    @{ Name = "NumNics" ; Expression = { $_.Summary.Hardware.NumNics } },
		    @{ Name = "ManagemetIP" ; Expression = { Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty IP } },
		    @{ Name = "ManagemetMacAddress" ; Expression = { Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Mac } },
		    @{ Name = "ManagemetVMKernel" ; Expression = { Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Name } },
		    @{ Name = "ManagemetSubnetMask" ; Expression = { Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask } },
		    @{ Name = "vMotionIP" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } },
		    @{ Name = "vMotionMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } },
		    @{ Name = "vMotionVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } },
		    @{ Name = "vMotionSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } },
		    @{ Name = "FtIP" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } },
		    @{ Name = "FtMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } },
		    @{ Name = "FtVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } },
		    @{ Name = "FtSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } },
		    @{ Name = "VSANIP" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } },
		    @{ Name = "VSANMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } },
		    @{ Name = "VSANVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } },
		    @{ Name = "VSANSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostNetworkAdapter | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } },
		    @{ Name = "NumHBAs" ; Expression = { $_.Summary.Hardware.NumHBAs } },
		    @{ Name = "iSCSIIP" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).IPv4 ) } },
		    @{ Name = "iSCSIMac" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).MACAddress ) } },
		    @{ Name = "iSCSIVMKernel" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).Vmknic ) } },
		    @{ Name = "iSCSISubnetMask" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).IPv4SubnetMask ) } },
		    @{ Name = "iSCSIAdapter" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).Adapter ) } },
		    @{ Name = "iSCSILinkUp" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).LinkUp ) } },
		    @{ Name = "iSCSIMTU" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).MTU ) } },
		    @{ Name = "iSCSINICDriver" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICDriver ) } },
		    @{ Name = "iSCSINICDriverVersion" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICDriverVersion ) } },
		    @{ Name = "iSCSINICFirmwareVersion" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICFirmwareVersion ) } },
		    @{ Name = "iSCSIPathStatus" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).PathStatus ) } },
		    @{ Name = "iSCSIVlanID" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).VlanID ) } },
		    @{ Name = "iSCSIVswitch" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).Vswitch ) } },
		    @{ Name = "iSCSICompliantStatus" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).CompliantStatus ) } },
		    @{ Name = "IScsiName" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostHBA -Type IScsi ).IScsiName ) } },
            @{ Name = "PortGroup" ; Expression = { if ( $_.Network -like "DistributedVirtualPortgroup*" ) { [string]::Join( ", ", ( Get-VDPortGroup -Id $_.Network ) ) }  
                elseif ( $_.Network -like "VmwareDistributedVirtualSwitch*" ) { [string]::Join( ", ", ( Get-VDSwitch -Id $_.Network ) ) }  
                elseif ( $_.Network -like "Network*" ) { [string]::Join( ", ", ( Get-VirtualNetwork -Id $_.Network ) ) } } } |
		Export-Csv $VmHostExportFile -Append -NoTypeInformation
}
#endregion ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Vm_Export
{
	$VmExportFile = "$CaptureCsvFolder\$vCenter-VmExport.csv"
	Get-View -ViewType VirtualMachine | 
		Where-Object { $_.Config.Template -eq $False } |
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", ( ( $_.Name ) ) ) } },
			@{ Name = "Datacenter" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( $_.Config.DatastoreUrl.Name ) ) } },
			@{ Name = "ResourcePool" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -VM ( Get-VM -Id $_.MoRef ) | Where-Object { $_ -notlike "Resources" } ) ) } },
			@{ Name = "vSwitch" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualSwitch -VM ( Get-VM -Id $_.MoRef ) ) ) ) } },
			@{ Name = "PortGroup" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualPortGroup -VM ( Get-VM -Id $_.MoRef ) ) ) ) } },
			@{ Name = "OS" ; Expression = { [string]::Join( ", ", ( $_.Config.GuestFullName ) ) } },
			@{ Name = "Version" ; Expression = { [string]::Join( ", ", ( $_.Config.Version ) ) } },
			@{ Name = "VMToolsVersion" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersion ) ) } },
			@{ Name = "ToolsVersionStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersionStatus ) ) } },
			@{ Name = "ToolsStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsStatus ) ) } },
			@{ Name = "ToolsRunningStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsRunningStatus ) ) } },
			@{ Name = "Folder" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).Name ) ) } },
			@{ Name = "NumCPU" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCPU ) ) } },
			@{ Name = "CoresPerSocket" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCoresPerSocket ) ) } },
			@{ Name = "MemoryGB" ; Expression = { [string]::Join( ", ", ( [math]::Round([decimal] ( $_.Config.Hardware.MemoryMB / 1024 ), 0 ) ) ) } },
			@{ Name = "IP" ; Expression = { [string]::Join(", ", ( $_.Guest.IpAddress ) ) } },
			@{ Name = "MacAddress" ; Expression = { [string]::Join(", ", ( $_.Guest.Net.MacAddress ) ) } },
			@{ Name = "ProvisionedSpaceGB" ; Expression = { [string]::Join( ", ", ( [math]::Round([decimal] ( $_.ProvisionedSpaceGB - $_.MemoryGB ), 0 ) ) ) } },
			@{ Name = "NumEthernetCards" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumEthernetCards ) ) } },
			@{ Name = "NumVirtualDisks" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumVirtualDisks ) ) } },
			@{ Name = "CpuReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.CpuReservation ) ) } },
			@{ Name = "MemoryReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.MemoryReservation ) ) } },
			@{ Name = "CpuHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotAddEnabled ) ) } },
			@{ Name = "CpuHotRemoveEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotRemoveEnabled ) ) } },
			@{ Name = "MemoryHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.MemoryHotAddEnabled ) ) } },
			@{ Name = "SRM" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.ManagedBy.Type ) ) } },
			@{ Name = "Snapshot" ; Expression = { [string]::Join( ", ", ( Get-Snapshot -VM $_.Name -Id ( $_.Snapshot.CurrentSnapshot ) ) ) } },
			@{ Name = "RootSnapshot" ; Expression = { [string]::Join( ", ", ( ( Get-Snapshot -VM $_.Name -Id $_.RootSnapshot ).Name ) ) } },
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( ( $_.MoRef ) ) ) } } |
		Export-Csv $VmExportFile -Append -NoTypeInformation
}
#endregion ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Template_Export
{
	$TemplateExportFile = "$CaptureCsvFolder\$vCenter-TemplateExport.csv"
	Get-View -ViewType VirtualMachine | 
		Where-Object { $_.Config.Template -eq $True } |
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", ( ( $_.Name ) ) ) } },
			@{ Name = "Datacenter" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -VM ( Get-VM -Id $_.MoRef ) ) ) } },
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( $_.Config.DatastoreUrl.Name ) ) } },
			@{ Name = "ResourcePool" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -VM ( Get-VM -Id $_.MoRef ) | Where-Object { $_ -notlike "Resources" } ) ) } },
			@{ Name = "vSwitch" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualSwitch -VM ( Get-VM -Id $_.MoRef ) ) ) ) } },
			@{ Name = "PortGroup" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualPortGroup -VM ( Get-VM -Id $_.MoRef ) ) ) ) } },
			@{ Name = "OS" ; Expression = { [string]::Join( ", ", ( $_.Config.GuestFullName ) ) } },
			@{ Name = "Version" ; Expression = { [string]::Join( ", ", ( $_.Config.Version ) ) } },
			@{ Name = "VMToolsVersion" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersion ) ) } },
			@{ Name = "ToolsVersionStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersionStatus ) ) } },
			@{ Name = "ToolsStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsStatus ) ) } },
			@{ Name = "ToolsRunningStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsRunningStatus ) ) } },
			@{ Name = "Folder" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).Name ) ) } },
			@{ Name = "NumCPU" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCPU ) ) } },
			@{ Name = "CoresPerSocket" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCoresPerSocket ) ) } },
			@{ Name = "MemoryGB" ; Expression = { [string]::Join( ", ", ( [math]::Round([decimal] ( $_.Config.Hardware.MemoryMB / 1024 ), 0 ) ) ) } },
			@{ Name = "MacAddress" ; Expression = { [string]::Join(", ", ( $_.Guest.Net.MacAddress ) ) } },
			@{ Name = "NumEthernetCards" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumEthernetCards ) ) } },
			@{ Name = "NumVirtualDisks" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumVirtualDisks ) ) } },
			@{ Name = "CpuReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.CpuReservation ) ) } },
			@{ Name = "MemoryReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.MemoryReservation ) ) } },
			@{ Name = "CpuHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotAddEnabled ) ) } },
			@{ Name = "CpuHotRemoveEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotRemoveEnabled ) ) } },
			@{ Name = "MemoryHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.MemoryHotAddEnabled ) ) } },
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( ( $_.MoRef ) ) ) } } |
		Export-Csv $TemplateExportFile -Append -NoTypeInformation
}
#endregion ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function DatastoreCluster_Export
{
	$DatastoreClusterExportFile = "$CaptureCsvFolder\$vCenter-DatastoreClusterExport.csv"
	Get-View -ViewType StoragePod |
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } },
			@{ Name = "Datacenter" ; Expression = { ( Get-DatastoreCluster -Id $_.MoRef | Get-Datastore ).Datacenter | Select-Object -Unique } },
			@{ Name = "Cluster" ; Expression = { ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost ).Parent | Select-Object -Unique } },
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost | Sort-Object Name ).Name ) ) ) } },
			@{ Name = "SdrsAutomationLevel" ; Expression = { $_.PodStorageDrsEntry.StorageDrsConfig.PodConfig.DefaultVmBehavior } },
			@{ Name = "IOLoadBalanceEnabled" ; Expression = { $_.PodStorageDrsEntry.StorageDrsConfig.PodConfig.IoLoadBalanceEnabled } },
			@{ Name = "CapacityGB" ; Expression = { [math]::Round( [decimal]$_.Summary.Capacity/1073741824, 0 ) } },
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | 
	Export-Csv $DatastoreClusterExportFile -Append -NoTypeInformation
}
#endregion ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datastore_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datastore_Export
{
	$DatastoreExportFile = "$CaptureCsvFolder\$vCenter-DatastoreExport.csv"
	Get-View -ViewType Datastore |
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } },
			@{ Name = "Datacenter" ; Expression = { ( Get-Datastore -Id $_.MoRef ).Datacenter } },
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster (Get-VmHost -Id $_.Host.Key).Parent.Name ) ) } },
			@{ Name = "DatastoreCluster" ; Expression = { Get-DatastoreCluster -Datastore ( Get-Datastore -Id $_.MoRef ) } },
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( Get-VmHost -Id $_.Host.Key | Sort-Object Name ) ) } },
			@{ Name = "Type" ; Expression = { $_.Info.Vmfs.Type } },
			@{ Name = "FileSystemVersion" ; Expression = { $_.Info.Vmfs.Version } },
			@{ Name = "StorageIOControlEnabled" ; Expression = { $_.IormConfiguration.Enabled } },
			@{ Name = "CapacityGB" ; Expression = { [math]::Round( [decimal] $_.Summary.Capacity / 1073741824, 0 ) } },
			@{ Name = "FreeSpaceGB" ; Expression = { [math]::Round( [decimal] $_.Summary.FreeSpace / 1073741824, 0 ) } },
			@{ Name = "CongestionThresholdMillisecond" ; Expression = { $_.IormConfiguration.CongestionThreshold } },
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | 
		Export-Csv $DatastoreExportFile -Append -NoTypeInformation
}
#endregion ~~< Datastore_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VsSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VsSwitch_Export
{
	$VsSwitchExportFile = "$CaptureCsvFolder\$vCenter-VsSwitchExport.csv"
	Get-VirtualSwitch -Standard | Sort-Object Name | 
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Datacenter" ; E = { Get-Datacenter -VmHost $_.VmHost } }, 
	@{ N = "Cluster" ; E = { Get-Cluster -VmHost $_.VmHost } }, 
	@{ N = "VmHost" ; E = { $_.VmHost } }, 
	@{ N = "Nic" ; E = { $_.Nic } }, 
	@{ N = "NumPorts" ; E = { $_.ExtensionData.Spec.NumPorts } }, 
	@{ N = "AllowPromiscuous" ; E = { $_.ExtensionData.Spec.Policy.Security.AllowPromiscuous } }, 
	@{ N = "MacChanges" ; E = { $_.ExtensionData.Spec.Policy.Security.MacChanges } }, 
	@{ N = "ForgedTransmits" ; E = { $_.ExtensionData.Spec.Policy.Security.ForgedTransmits } }, 
	@{ N = "Policy" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.Policy } }, 
	@{ N = "ReversePolicy" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.ReversePolicy } }, 
	@{ N = "NotifySwitches" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.NotifySwitches } }, 
	@{ N = "RollingOrder" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.RollingOrder } }, 
	@{ N = "ActiveNic" ; E = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic)) } }, 
	@{ N = "StandbyNic" ; E = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.StandbyNic)) } } | Export-Csv $VsSwitchExportFile -Append -NoTypeInformation
}
#endregion ~~< VsSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssPort_Export
{
	$VssPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VssPortGroupExport.csv"
	foreach ($VMHost in Get-VMHost)
	{
		foreach ($VsSwitch in(Get-VirtualSwitch -Standard -VMHost $VmHost))
		{
			Get-VirtualPortGroup -Standard -VirtualSwitch $VsSwitch | Sort-Object Name | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -VMHost $VMHost.Name } }, 
			@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.Name } }, 
			@{ N = "VmHost" ; E = { $VMHost.Name } }, 
			@{ N = "VsSwitch" ; E = { $VsSwitch.Name } }, 
			@{ N = "VLanId" ; E = { $_.VLanId } }, 
			@{ N = "ActiveNic" ; E = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.ActiveNic)) } }, 
			@{ N = "StandbyNic" ; E = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.StandbyNic)) } } | Export-Csv $VssPortGroupExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< VssPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssVmk_Export
{
	$VssVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VssVmkernelExport.csv"
	foreach ($VMHost in Get-VMHost)
	{
		foreach ($VsSwitch in(Get-VirtualSwitch -VMHost $VmHost -Standard))
		{
			foreach ($VssPort in(Get-VirtualPortGroup -Standard -VMHost $VmHost | Sort-Object Name))
			{
				Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VsSwitch -PortGroup $VssPort | Sort-Object Name | 
				Select-Object @{ N = "Name" ; E = { $_.Name } }, 
				@{ N = "Datacenter" ; E = { Get-Datacenter -VMHost $VMHost.Name } }, 
				@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.Name } }, 
				@{ N = "VmHost" ; E = { $VMHost.Name } }, 
				@{ N = "VsSwitch" ; E = { $VsSwitch.Name } }, 
				@{ N = "PortGroupName" ; E = { $_.PortGroupName } }, 
				@{ N = "DhcpEnabled" ; E = { $_.DhcpEnabled } }, 
				@{ N = "IP" ; E = { $_.IP } }, 
				@{ N = "Mac" ; E = { $_.Mac } }, 
				@{ N = "ManagementTrafficEnabled" ; E = { $_.ManagementTrafficEnabled } }, 
				@{ N = "VMotionEnabled" ; E = { $_.VMotionEnabled } }, 
				@{ N = "FaultToleranceLoggingEnabled" ; E = { $_.FaultToleranceLoggingEnabled } }, 
				@{ N = "VsanTrafficEnabled" ; E = { $_.VsanTrafficEnabled } }, 
				@{ N = "Mtu" ; E = { $_.Mtu } } | Export-Csv $VssVmkernelExportFile -Append -NoTypeInformation
			}
		}
	}
}
#endregion ~~< VssVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssPnic_Export
{
	$VssPnicExportFile = "$CaptureCsvFolder\$vCenter-VssPnicExport.csv"
	foreach ($VMHost in Get-VMHost)
	{
		foreach ($VsSwitch in(Get-VirtualSwitch -Standard -VMHost $VmHost))
		{
			Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VsSwitch -VMHost $VmHost | Sort-Object Name | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -VmHost $VmHost } }, 
			@{ N = "Cluster" ; E = { Get-Cluster -VmHost $_.VmHost } }, 
			@{ N = "VmHost" ; E = { $_.VmHost } }, 
			@{ N = "VsSwitch" ; E = { $VsSwitch.Name } }, 
			@{ N = "Mac" ; E = { $_.Mac } } | Export-Csv $VssPnicExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< VssPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdSwitch_Export
{
	$VdSwitchExportFile = "$CaptureCsvFolder\$vCenter-VdSwitchExport.csv"
	Get-View -ViewType DistributedVirtualSwitch | 
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } }, 
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) ) } },
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ) ) ) } },
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ) ) } },
			@{ Name = "Vendor" ; Expression = { $_.Summary.ProductInfo.Vendor } }, 
			@{ Name = "Version" ; Expression = { $_.Summary.ProductInfo.Version } }, 
			@{ Name = "NumUplinkPorts" ; Expression = { ($_.Config.UplinkPortPolicy.UplinkPortName).Count } }, 
			@{ Name = "UplinkPortName" ; Expression = { [string]::Join( ", ", ( $_.Config.UplinkPortPolicy.UplinkPortName | Sort-Object Name  ) ) } },
			@{ Name = "Mtu" ; Expression = { $_.Config.MaxMtu } },
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | 
		Export-Csv $VdSwitchExportFile -Append -NoTypeInformation
}
#endregion ~~< VdSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsPort_Export
{
	$VdsPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VdsPortGroupExport.csv"
	Get-View -ViewType DistributedVirtualPortgroup | 
		Sort-Object Name |
		Where-Object { $_.Name -notlike "*DVUplinks*" } | 
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } }, 
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Host ) ) } }, 
			@{ Name = "Cluster" ; Expression = { Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Host ) ) } }, 
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Host ) | Sort-Object Name  ) ) } }, 
			@{ Name = "VlanConfiguration" ; Expression = { "VLAN "+ $_.Config.DefaultPortConfig.Vlan.VlanId } }, 
			@{ Name = "VdSwitch" ; Expression = { ( Get-VdSwitch  -Id $_.Config.DistributedVirtualSwitch ) } }, 
			@{ Name = "NumPorts" ; Expression = { $_.Config.NumPorts } }, 
			@{ Name = "ActiveUplinkPort" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.ActiveUplinkPort)) } }, 
			@{ Name = "StandbyUplinkPort" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.StandbyUplinkPort)) } }, 
			@{ Name = "Policy" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value } }, 
			@{ Name = "ReversePolicy" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.ReversePolicy.Value } }, 
			@{ Name = "NotifySwitches" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.NotifySwitches.Value } }, 
			@{ Name = "PortBinding" ; Expression = { ( Get-VDPortgroup  -Id $_.MoRef ).PortBinding } },
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | 
		Export-Csv $VdsPortGroupExportFile -Append -NoTypeInformation
	}
#endregion ~~< VdsPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsVmk_Export
{
	$VdsVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VdsVmkernelExport.csv"
	foreach ($VmHost in Get-VmHost)
	{
		foreach ($VdSwitch in(Get-VdSwitch -VMHost $VmHost))
		{
			Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VdSwitch -VMHost $VmHost | Sort-Object -Property Name -Unique | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -VMHost $VMHost.name } }, 
			@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.name } }, 
			@{ N = "VmHost" ; E = { $VMHost.Name } }, 
			@{ N = "VdSwitch" ; E = { $VdSwitch.Name } }, 
			@{ N = "PortGroupName" ; E = { $_.PortGroupName } }, 
			@{ N = "DhcpEnabled" ; E = { $_.DhcpEnabled } }, 
			@{ N = "IP" ; E = { $_.IP } }, 
			@{ N = "Mac" ; E = { $_.Mac } }, 
			@{ N = "ManagementTrafficEnabled" ; E = { $_.ManagementTrafficEnabled } }, 
			@{ N = "VMotionEnabled" ; E = { $_.VMotionEnabled } }, 
			@{ N = "FaultToleranceLoggingEnabled" ; E = { $_.FaultToleranceLoggingEnabled } }, 
			@{ N = "VsanTrafficEnabled" ; E = { $_.VsanTrafficEnabled } }, 
			@{ N = "Mtu" ; E = { $_.Mtu } } | Export-Csv $VdsVmkernelExportFile -Append -NoTypeInformation
					
		}
	}
}
#endregion ~~< VdsVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsPnic_Export
{
	$VdsPnicExportFile = "$CaptureCsvFolder\$vCenter-VdsPnicExport.csv"
	foreach ($VmHost in Get-VmHost)
	{
		foreach ($VdSwitch in(Get-VdSwitch -VMHost $VmHost))
		{
			Get-VDPort -VdSwitch $VdSwitch -Uplink | Sort-Object -Property ConnectedEntity -Unique | 
			Select-Object @{ N = "Name" ; E = { $_.ConnectedEntity } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -VMHost $VMHost.name } }, 
			@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.name } }, 
			@{ N = "VmHost" ; E = { $VMHost.Name } }, 
			@{ N = "VdSwitch" ; E = { $VdSwitch } }, 
			@{ N = "Portgroup" ; E = { $_.Portgroup } }, 
			@{ N = "ConnectedEntity" ; E = { $_.Name } }, 
			@{ N = "VlanConfiguration" ; E = { $_.VlanConfiguration } } | Export-Csv $VdsPnicExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< VdsPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Folder_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Folder_Export
{
	$FolderExportFile = "$CaptureCsvFolder\$vCenter-FolderExport.csv"
	foreach ($Datacenter in Get-Datacenter)
	{
		Get-View -ViewType Folder |
			Sort-Object Name | 
			Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } },
				@{ Name = "Datacenter" ; Expression = { $Datacenter.Name } },
				@{ Name = "ChildType" ; Expression = { [string]::Join( ", ", ( $_.ChildType ) ) } },
				@{ Name = "ChildEntity" ; Expression = { if ( $_.ChildEntity -like "Datacenter*" ) { [string]::Join( ", ", ( Get-Datacenter -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "ClusterComputeResource*" ) { [string]::Join( ", ", ( Get-Cluster -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "DistributedVirtualPortgroup*" ) { [string]::Join( ", ", ( Get-VDPortGroup -Id $_.ChildEntity ) ) }  
					elseif ( $_.ChildEntity -like "VmwareDistributedVirtualSwitch*" ) { [string]::Join( ", ", ( Get-VDSwitch -Id $_.ChildEntity ) ) }  
					elseif ( $_.ChildEntity -like "Network*" ) { [string]::Join( ", ", ( Get-VirtualSwitch -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "Datastore*" ) { [string]::Join( ", ", ( Get-Datastore -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "StoragePod*" ) { [string]::Join( ", ", ( Get-DatastoreCluster -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "Folder*" ) { [string]::Join( ", ", ( Get-Folder -Id $_.ChildEntity ) ) } 
					elseif ( $_.ChildEntity -like "VirtualMachine*" ) { [string]::Join( ", ", ( Get-VM -Id $_.ChildEntity ) ) } } },
				@{ Name = "LinkedView" ; Expression = { [string]::Join( ", ", ( $_.LinkedView ) ) } },
				@{ Name = "Parent" ; Expression = { if ( $_.Parent -like "Datacenter*" ) { [string]::Join( ", ", ( Get-Datacenter -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "ClusterComputeResource*" ) { [string]::Join( ", ", ( Get-Cluster -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "DistributedVirtualPortgroup*" ) { [string]::Join( ", ", ( Get-VDPortGroup -Id $_.Parent ) ) }  
					elseif ( $_.Parent -like "VmwareDistributedVirtualSwitch*" ) { [string]::Join( ", ", ( Get-VDSwitch -Id $_.Parent ) ) }  
					elseif ( $_.Parent -like "Network*" ) { [string]::Join( ", ", ( Get-VirtualSwitch -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "Datastore*" ) { [string]::Join( ", ", ( Get-Datastore -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "StoragePod*" ) { [string]::Join( ", ", ( Get-DatastoreCluster -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "Folder*" ) { [string]::Join( ", ", ( Get-Folder -Id $_.Parent ) ) } 
					elseif ( $_.Parent -like "VirtualMachine*" ) { [string]::Join( ", ", ( Get-VM -Id $_.Parent ) ) } } },
				@{ Name = "CustomValue" ; Expression = { [string]::Join( ", ", ( $_.CustomValue ) ) } },
				@{ Name = "OverallStatus" ; Expression = { [string]::Join( ", ", ( $_.OverallStatus ) ) } },
				@{ Name = "ConfigStatus" ; Expression = { [string]::Join( ", ", ( $_.ConfigStatus ) ) } },
				@{ Name = "ConfigIssue" ; Expression = { [string]::Join( ", ", ( $_.ConfigIssue ) ) } },
				@{ Name = "EffectiveRole" ; Expression = { [string]::Join( ", ", ( $_.EffectiveRole ) ) } },
				@{ Name = "Permission" ; Expression = { [string]::Join( ", ", ( $_.Permission ) ) } },
				@{ Name = "DisabledMethod" ; Expression = { [string]::Join( ", ", ( $_.DisabledMethod ) ) } },
				@{ Name = "AlarmActionsEnabled" ; Expression = { [string]::Join( ", ", ( $_.AlarmActionsEnabled ) ) } },
				@{ Name = "Tag" ; Expression = { [string]::Join( ", ", ( $_.Tag ) ) } },
				@{ Name = "Value" ; Expression = { [string]::Join( ", ", ( $_.Value ) ) } },
				@{ Name = "AvailableField" ; Expression = { [string]::Join( ", ", ( $_.AvailableField ) ) } },
				@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } | 
			Export-Csv $FolderExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Folder_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Rdm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Rdm_Export
{
	$RdmExportFile = "$CaptureCsvFolder\$vCenter-RdmExport.csv"
	Get-VM | Get-HardDisk | 
		Where-Object { $_.DiskType -like "Raw*" } | 
		Sort-Object Parent | 
		Select-Object @{ N = "ScsiCanonicalName" ; E = { $_.ScsiCanonicalName } },
			@{ N = "Cluster" ; E = { Get-Cluster -VM $_.Parent } },
			@{ N = "Vm" ; E = { $_.Parent } },
			@{ N = "Label" ; E = { $_.Name } },
			@{ N = "CapacityGB" ; E = { [math]::Round([decimal]$_.CapacityGB, 2) } },
			@{ N = "DiskType" ; E = { $_.DiskType } },
			@{ N = "Persistence" ; E = { $_.Persistence } },
			@{ N = "CompatibilityMode" ; E = { $_.ExtensionData.Backing.CompatibilityMode } },
			@{ N = "DeviceName" ; E = { $_.ExtensionData.Backing.DeviceName } },
			@{ N = "Sharing" ; E = { $_.ExtensionData.Backing.Sharing } } |
		Export-Csv $RdmExportFile -Append -NoTypeInformation
}
#endregion ~~< Rdm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Drs_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_Rule_Export
{
	$DrsRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsRuleExport.csv"
	foreach ($Cluster in Get-Cluster)
	{
		Get-Cluster $Cluster | Get-DrsRule | Sort-Object Name | 
		Select-Object @{ N = "Name" ; E = { $_.Name } }, 
		@{ N = "Datacenter" ; E = { Get-Datacenter -Cluster $Cluster.Name } }, 
		@{ N = "Cluster" ; E = { $_.Cluster } }, 
		@{ N = "Type" ; E = { $_.Type } }, 
		@{ N = "Enabled" ; E = { $_.Enabled } }, 
		@{ N = "Mandatory" ; E = { $_.Mandatory } } | Export-Csv $DrsRuleExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Drs_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Drs_Cluster_Group_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_Cluster_Group_Export
{
	$DrsClusterGroupExportFile = "$CaptureCsvFolder\$vCenter-DrsClusterGroupExport.csv"
	foreach ($Cluster in Get-Cluster)
	{
		Get-Cluster $Cluster | Get-DrsClusterGroup | Sort-Object Name | 
		Select-Object @{ N = "Name" ; E = { $_.Name } }, 
		@{ N = "Datacenter" ; E = { Get-Datacenter -Cluster $Cluster.Name } }, 
		@{ N = "Cluster" ; E = { $_.Cluster } }, 
		@{ N = "GroupType" ; E = { $_.GroupType } }, 
		@{ N = "Member" ; E = { $_.Member } } | Export-Csv $DrsClusterGroupExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Drs_Cluster_Group_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Drs_VmHost_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_VmHost_Rule_Export
{
	$DrsVmHostRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsVmHostRuleExport.csv"
	foreach ($Cluster in Get-Cluster)
	{
		foreach ($DrsClusterGroup in (Get-Cluster $Cluster | Get-DrsClusterGroup | Sort-Object Name))
		{
			Get-DrsVmHostRule -VMHostGroup $DRSClusterGroup | Sort-Object Name | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -Cluster $Cluster.Name } }, 
			@{ N = "Cluster" ; E = { $_.Cluster } }, 
			@{ N = "Enabled" ; E = { $_.Enabled } }, 
			@{ N = "Type" ; E = { $_.Type } }, 
			@{ N = "VMGroup" ; E = { $_.VMGroup } }, 
			@{ N = "VMHostGroup" ; E = { $_.VMHostGroup } }, 
			@{ N = "AffineHostGroupName" ; E = { $_.ExtensionData.AffineHostGroupName } }, 
			@{ N = "AntiAffineHostGroupName" ; E = { $_.ExtensionData.AntiAffineHostGroupName } } | Export-Csv $DrsVmHostRuleExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Drs_VmHost_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Resource_Pool_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Resource_Pool_Export
{
	$ResourcePoolExportFile = "$CaptureCsvFolder\$vCenter-ResourcePoolExport.csv"
	Get-View -ViewType ResourcePool | 
		Sort-Object Name | 
		Select-Object @{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } },
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster -ID (Get-ResourcePool -Id (Get-View -ViewType ResourcePool).MoRef ).Parent.ExtensionData.MoRef ) ) ) } }, 
			@{ Name = "CpuSharesLevel" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuSharesLevel ) ) } },
			@{ Name = "NumCpuShares" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).NumCpuShares ) ) } },
			@{ Name = "CpuReservationMHz" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuReservationMHz ) ) } }, 
			@{ Name = "CpuExpandableReservation" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuExpandableReservation ) ) } },
			@{ Name = "CpuLimitMHz" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuLimitMHz ) ) } },
			@{ Name = "MemSharesLevel" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemSharesLevel ) ) } },
			@{ Name = "NumMemShares" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).NumMemShares ) ) } },
			@{ Name = "MemReservationGB" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemReservationGB ) ) } }, 
			@{ Name = "MemExpandableReservation" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemExpandableReservation ) ) } },
			@{ Name = "MemLimitGB" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemLimitGB ) ) } },
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } |
		Export-Csv $ResourcePoolExportFile -Append -NoTypeInformation
}
#endregion ~~< Resource_Pool_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Snapshot_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Snapshot_Export
{
	$SnapshotExportFile = "$CaptureCsvFolder\$vCenter-SnapshotExport.csv"
	Get-VM | Get-Snapshot |
		Sort-Object  VM, Created | 
		Select-Object @{ Name = "VM" ; Expression = { $_.VM } }, 
			@{ Name = "Name" ; Expression = { $_.Name } }, 
			@{ Name = "Created" ; Expression = { $_.Created } }, 
			@{ Name = "Children" ; Expression = { $_.Children } }, 
			@{ Name = "ParentSnapshot" ; Expression = { $_.ParentSnapshot } },
			@{ Name = "IsCurrent" ; Expression = { $_.IsCurrent } } |
		Export-Csv $SnapshotExportFile -Append -NoTypeInformation
}
#endregion ~~< Snapshot_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Linked_vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Linked_vCenter_Export
{
	$LinkedvCenterExportFile = "$CaptureCsvFolder\$vCenter-LinkedvCenterExport.csv"
	Disconnect-ViServer * -Confirm:$false
	$global:vCenter = $VcenterTextBox.Text
	$User = $UserNameTextBox.Text
	Connect-VIServer $Vcenter -user $User -password $PasswordTextBox.Text -AllLinked
	$global:DefaultVIServers |
		Where-Object { $_.Name -ne "$vCenter" } |
		Select-Object @{ Name = "Name" ; Expression = { $_.Name } }, 
			@{ Name = "Version" ; Expression = { $_.Version } }, 
			@{ Name = "Build" ; Expression = { $_.Build } },
			@{ Name = "OsType" ; Expression = { $_.ExtensionData.Content.About.OsType } },
			@{ Name = "vCenter" ; Expression = { $vCenter } } |
		Export-Csv $LinkedvCenterExportFile -Append -NoTypeInformation
}
#endregion ~~< Linked_vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Export Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Visio Object Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Connect-VisioObject >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Connect-VisioObject($firstObj, $secondObj)
{
	$shpConn = $pagObj.Drop($pagObj.Application.ConnectorToolDataObject, 0, 0)
	$ConnectBegin = $shpConn.CellsU("BeginX").GlueTo($firstObj.CellsU("PinX"))
	$ConnectEnd = $shpConn.CellsU("EndX").GlueTo($secondObj.CellsU("PinX"))
}
#endregion ~~< Connect-VisioObject >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVC($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectDC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDC($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectCluster($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectHost($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVM($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectTemplate >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectTemplate($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectTemplate >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectSRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectSRM($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectSRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectDatastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDatastore($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDatastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectHardDisk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectHardDisk($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.ScsiCanonicalName
	return $shpObj
}
#endregion ~~< Add-VisioObjectHardDisk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectFolder($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVsSwitch($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectPG($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVssPNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVssPNIC($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVssPNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVMK >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVMK($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVMK >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVdSwitch($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVdsPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVdsPG($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVdsPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectVdsPNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVdsPNIC($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVdsPNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectDrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDrsRule($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectDrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDrsClusterGroup($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectDRSVMHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectDRSVMHostRule($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectDRSVMHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectResourcePool($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectRecoveryPlan Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectRecoveryPlan($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectRecoveryPlan Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectProtectionGroup Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectProtectionGroup($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectProtectionGroup Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Add-VisioObjectSnapshot Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectSnapshot($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectSnapshot Function >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Visio Object Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Visio Draw Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_vCenter
{
	# Name
	$VCObject.Cells("Prop.Name").Formula = '"' + $vCenterImport.Name + '"'
	# Version
	$VCObject.Cells("Prop.Version").Formula = '"' + $vCenterImport.Version + '"'
	# Build
	$VCObject.Cells("Prop.Build").Formula = '"' + $vCenterImport.Build + '"'
	# OsType
	$VCObject.Cells("Prop.OsType").Formula = '"' + $vCenterImport.OsType + '"'
}
#endregion ~~< Draw_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Datacenter
{
	# Name
	$DatacenterObject.Cells("Prop.Name").Formula = '"' + $Datacenter.Name + '"'
}
#endregion ~~< Draw_Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Cluster
{
	# Name
	$ClusterObject.Cells("Prop.Name").Formula = '"' + $Cluster.Name + '"'
	# HAEnabled
	$ClusterObject.Cells("Prop.HAEnabled").Formula = '"' + $Cluster.HAEnabled + '"'
	# HAAdmissionControlEnabled
	$ClusterObject.Cells("Prop.HAAdmissionControlEnabled").Formula = '"' + $Cluster.HAAdmissionControlEnabled + '"'
	# AdmissionControlPolicyCpuFailoverResourcesPercent
	$ClusterObject.Cells("Prop.AdmissionControlPolicyCpuFailoverResourcesPercent").Formula = '"' + $Cluster.AdmissionControlPolicyCpuFailoverResourcesPercent + '"'
	# AdmissionControlPolicyMemoryFailoverResourcesPercent
	$ClusterObject.Cells("Prop.AdmissionControlPolicyMemoryFailoverResourcesPercent").Formula = '"' + $Cluster.AdmissionControlPolicyMemoryFailoverResourcesPercent + '"'
	# AdmissionControlPolicyFailoverLevel
	$ClusterObject.Cells("Prop.AdmissionControlPolicyFailoverLevel").Formula = '"' + $Cluster.AdmissionControlPolicyFailoverLevel + '"'
	# AdmissionControlPolicyAutoComputePercentages
	$ClusterObject.Cells("Prop.AdmissionControlPolicyAutoComputePercentages").Formula = '"' + $Cluster.AdmissionControlPolicyAutoComputePercentages + '"'
	# AdmissionControlPolicyResourceReductionToToleratePercent
	$ClusterObject.Cells("Prop.AdmissionControlPolicyResourceReductionToToleratePercent").Formula = '"' + $Cluster.AdmissionControlPolicyResourceReductionToToleratePercent + '"'
	# DrsEnabled
	$ClusterObject.Cells("Prop.DrsEnabled").Formula = '"' + $Cluster.DrsEnabled + '"'
	# DrsAutomationLevel
	$ClusterObject.Cells("Prop.DrsAutomationLevel").Formula = '"' + $Cluster.DrsAutomationLevel + '"'
	# VmMonitoring
	$ClusterObject.Cells("Prop.VmMonitoring").Formula = '"' + $Cluster.VmMonitoring + '"'
	# HostMonitoring
	$ClusterObject.Cells("Prop.HostMonitoring").Formula = '"' + $Cluster.HostMonitoring + '"'
}
#endregion ~~< Draw_Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VmHost
{
	# Name
	$HostObject.Cells("Prop.Name").Formula = '"' + $VMHost.Name + '"'
	# Version
	$HostObject.Cells("Prop.Version").Formula = '"' + $VMHost.Version + '"'
	# Build
	$HostObject.Cells("Prop.Build").Formula = '"' + $VMHost.Build + '"'
	# Manufacturer
	$HostObject.Cells("Prop.Manufacturer").Formula = '"' + $VMHost.Manufacturer + '"'
	# Model
	$HostObject.Cells("Prop.Model").Formula = '"' + $VMHost.Model + '"'
	# LicenseType
	$HostObject.Cells("Prop.LicenseType").Formula = '"' + $VMHost.LicenseType + '"'
	# BiosVersion
	$HostObject.Cells("Prop.BiosVersion").Formula = '"' + $VMHost.BiosVersion + '"'
	# BIOSReleaseDate
	$HostObject.Cells("Prop.BIOSReleaseDate").Formula = '"' + $VMHost.BIOSReleaseDate + '"'
	# ProcessorType
	$HostObject.Cells("Prop.ProcessorType").Formula = '"' + $VMHost.ProcessorType + '"'
	# CpuMhz
	$HostObject.Cells("Prop.CpuMhz").Formula = '"' + $VMHost.CpuMhz + '"'
	# NumCpuPkgs
	$HostObject.Cells("Prop.NumCpuPkgs").Formula = '"' + $VMHost.NumCpuPkgs + '"'
	# NumCpuCores
	$HostObject.Cells("Prop.NumCpuCores").Formula = '"' + $VMHost.NumCpuCores + '"'
	# NumCpuThreads
	$HostObject.Cells("Prop.NumCpuThreads").Formula = '"' + $VMHost.NumCpuThreads + '"'
	# Memory
	$HostObject.Cells("Prop.Memory").Formula = '"' + $VMHost.Memory + '"'
	# MaxEVCMode
	$HostObject.Cells("Prop.MaxEVCMode").Formula = '"' + $VMHost.MaxEVCMode + '"'
	# NumNics
	$HostObject.Cells("Prop.NumNics").Formula = '"' + $VMHost.NumNics + '"'
	# ManagemetIP
	$HostObject.Cells("Prop.ManagemetIP").Formula = '"' + $VMHost.ManagemetIP + '"'
	# ManagemetMacAddress
	$HostObject.Cells("Prop.ManagemetMacAddress").Formula = '"' + $VMHost.ManagemetMacAddress + '"'
	# ManagemetVMKernel
	$HostObject.Cells("Prop.ManagemetVMKernel").Formula = '"' + $VMHost.ManagemetVMKernel + '"'
	# ManagemetSubnetMask
	$HostObject.Cells("Prop.ManagemetSubnetMask").Formula = '"' + $VMHost.ManagemetSubnetMask + '"'
	# vMotionIP
	$HostObject.Cells("Prop.vMotionIP").Formula = '"' + $VMHost.vMotionIP + '"'
	# vMotionMacAddress
	$HostObject.Cells("Prop.vMotionMacAddress").Formula = '"' + $VMHost.vMotionMacAddress + '"'
	# vMotionVMKernel
	$HostObject.Cells("Prop.vMotionVMKernel").Formula = '"' + $VMHost.vMotionVMKernel + '"'
	# vMotionSubnetMask
	$HostObject.Cells("Prop.vMotionSubnetMask").Formula = '"' + $VMHost.vMotionSubnetMask + '"'
	# FtIP
	$HostObject.Cells("Prop.FtIP").Formula = '"' + $VMHost.FtIP + '"'
	# FtMacAddress
	$HostObject.Cells("Prop.FtMacAddress").Formula = '"' + $VMHost.FtMacAddress + '"'
	# FtVMKernel
	$HostObject.Cells("Prop.FtVMKernel").Formula = '"' + $VMHost.FtVMKernel + '"'
	# FtSubnetMask
	$HostObject.Cells("Prop.FtSubnetMask").Formula = '"' + $VMHost.FtSubnetMask + '"'
	# VSANIP
	$HostObject.Cells("Prop.VSANIP").Formula = '"' + $VMHost.VSANIP + '"'
	# VSANMacAddress
	$HostObject.Cells("Prop.VSANMacAddress").Formula = '"' + $VMHost.VSANMacAddress + '"'
	# VSANVMKernel
	$HostObject.Cells("Prop.VSANVMKernel").Formula = '"' + $VMHost.VSANVMKernel + '"'
	# VSANSubnetMask
	$HostObject.Cells("Prop.VSANSubnetMask").Formula = '"' + $VMHost.VSANSubnetMask + '"'
	# NumHBAs
	$HostObject.Cells("Prop.NumHBAs").Formula = '"' + $VMHost.NumHBAs + '"'
	# iSCSIIP
	$HostObject.Cells("Prop.iSCSIIP").Formula = '"' + $VMHost.iSCSIIP + '"'
	# iSCSIMac
	$HostObject.Cells("Prop.iSCSIMac").Formula = '"' + $VMHost.iSCSIMac + '"'
	# iSCSIVMKernel
	$HostObject.Cells("Prop.iSCSIVMKernel").Formula = '"' + $VMHost.iSCSIVMKernel + '"'
	# iSCSISubnetMask
	$HostObject.Cells("Prop.iSCSISubnetMask").Formula = '"' + $VMHost.iSCSISubnetMask + '"'
	# iSCSIAdapter
	$HostObject.Cells("Prop.iSCSIAdapter").Formula = '"' + $VMHost.iSCSIAdapter + '"'
	# iSCSILinkUp
	$HostObject.Cells("Prop.iSCSILinkUp").Formula = '"' + $VMHost.iSCSILinkUp + '"'
	# iSCSIMTU
	$HostObject.Cells("Prop.iSCSIMTU").Formula = '"' + $VMHost.iSCSIMTU + '"'
	# iSCSINICDriver
	$HostObject.Cells("Prop.iSCSINICDriver").Formula = '"' + $VMHost.iSCSINICDriver + '"'
	# iSCSINICDriverVersion
	$HostObject.Cells("Prop.iSCSINICDriverVersion").Formula = '"' + $VMHost.iSCSINICDriverVersion + '"'
	# iSCSINICFirmwareVersion
	$HostObject.Cells("Prop.iSCSINICFirmwareVersion").Formula = '"' + $VMHost.iSCSINICFirmwareVersion + '"'
	# iSCSIPathStatus
	$HostObject.Cells("Prop.iSCSIPathStatus").Formula = '"' + $VMHost.iSCSIPathStatus + '"'
	# iSCSIVlanID
	$HostObject.Cells("Prop.iSCSIVlanID").Formula = '"' + $VMHost.iSCSIVlanID + '"'
	# iSCSIVswitch
	$HostObject.Cells("Prop.iSCSIVswitch").Formula = '"' + $VMHost.iSCSIVswitch + '"'
	# iSCSICompliantStatus
	$HostObject.Cells("Prop.iSCSICompliantStatus").Formula = '"' + $VMHost.iSCSICompliantStatus + '"'
	# IScsiName
	$HostObject.Cells("Prop.IScsiName").Formula = '"' + $VMHost.IScsiName + '"'
}
#endregion ~~< Draw_VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VM
{
	# Name
	$VMObject.Cells("Prop.Name").Formula = '"' + $VM.Name + '"'
	# OS
	$VMObject.Cells("Prop.OS").Formula = '"' + $VM.OS + '"'
	# Version
	$VMObject.Cells("Prop.Version").Formula = '"' + $VM.Version + '"'
	# VMToolsVersion
	$VMObject.Cells("Prop.VMToolsVersion").Formula = '"' + $VM.VMToolsVersion + '"'
	# ToolsVersionStatus
	$VMObject.Cells("Prop.ToolsVersionStatus").Formula = '"' + $VM.ToolsVersionStatus + '"'
	# ToolsStatus
	$VMObject.Cells("Prop.ToolsStatus").Formula = '"' + $VM.ToolsStatus + '"'
	# ToolsRunningStatus
	$VMObject.Cells("Prop.ToolsRunningStatus").Formula = '"' + $VM.ToolsRunningStatus + '"'
	# Folder
	$VMObject.Cells("Prop.Folder").Formula = '"' + $VM.Folder + '"'
	# NumCPU
	$VMObject.Cells("Prop.NumCPU").Formula = '"' + $VM.NumCPU + '"'
	# CoresPerSocket
	$VMObject.Cells("Prop.CoresPerSocket").Formula = '"' + $VM.CoresPerSocket + '"'
	# MemoryGB
	$VMObject.Cells("Prop.MemoryGB").Formula = '"' + $VM.MemoryGB + '"'
	# IP
	$VMObject.Cells("Prop.IP").Formula = '"' + $VM.Ip + '"'
	# MacAddress
	$VMObject.Cells("Prop.MacAddress").Formula = '"' + $VM.MacAddress + '"'
	# ProvisionedSpaceGB
	$VMObject.Cells("Prop.ProvisionedSpaceGB").Formula = '"' + $VM.ProvisionedSpaceGB + '"'
	# NumEthernetCards
	$VMObject.Cells("Prop.NumEthernetCards").Formula = '"' + $VM.NumEthernetCards + '"'
	# NumVirtualDisks
	$VMObject.Cells("Prop.NumVirtualDisks").Formula = '"' + $VM.NumVirtualDisks + '"'
	# CpuReservation
	$VMObject.Cells("Prop.CpuReservation").Formula = '"' + $VM.CpuReservation + '"'
	# MemoryReservation
	$VMObject.Cells("Prop.MemoryReservation").Formula = '"' + $VM.MemoryReservation + '"'
	# CpuHotAddEnabled
	$VMObject.Cells("Prop.CpuHotAddEnabled").Formula = '"' + $VM.CpuHotAddEnabled + '"'
	# CpuHotRemoveEnabled
	$VMObject.Cells("Prop.CpuHotRemoveEnabled").Formula = '"' + $VM.CpuHotRemoveEnabled + '"'
	# MemoryHotAddEnabled
	$VMObject.Cells("Prop.MemoryHotAddEnabled").Formula = '"' + $VM.MemoryHotAddEnabled + '"'
	# ProtectionGroup
	$VMObject.Cells("Prop.ProtectionGroup").Formula = '"' + $VM.ProtectionGroup + '"'
	# ProtectedVm
	$VMObject.Cells("Prop.ProtectedVm").Formula = '"' + $VM.ProtectedVm + '"'
	# PeerProtectedVm
	$VMObject.Cells("Prop.PeerProtectedVm").Formula = '"' + $VM.PeerProtectedVm + '"'
	# State
	$VMObject.Cells("Prop.State").Formula = '"' + $VM.State + '"'
	# PeerState
	$VMObject.Cells("Prop.PeerState").Formula = '"' + $VM.PeerState + '"'
	# NeedsConfiguration
	$VMObject.Cells("Prop.NeedsConfiguration").Formula = '"' + $VM.NeedsConfiguration + '"'
}
#endregion ~~< Draw_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Template
{
	# Name
	$TemplateObject.Cells("Prop.Name").Formula = '"' + $Template.Name + '"'
	# OS
	$TemplateObject.Cells("Prop.OS").Formula = '"' + $Template.OS + '"'
	# Version
	$TemplateObject.Cells("Prop.Version").Formula = '"' + $Template.Version + '"'
	# ToolsVersion
	$TemplateObject.Cells("Prop.ToolsVersion").Formula = '"' + $Template.ToolsVersion + '"'
	# ToolsVersionStatus
	$TemplateObject.Cells("Prop.ToolsVersionStatus").Formula = '"' + $Template.ToolsVersionStatus + '"'
	# ToolsStatus
	$TemplateObject.Cells("Prop.ToolsStatus").Formula = '"' + $Template.ToolsStatus + '"'
	# ToolsRunningStatus
	$TemplateObject.Cells("Prop.ToolsRunningStatus").Formula = '"' + $Template.ToolsRunningStatus + '"'
	# NumCPU
	$TemplateObject.Cells("Prop.NumCPU").Formula = '"' + $Template.NumCPU + '"'
	# NumCoresPerSocket
	$TemplateObject.Cells("Prop.NumCoresPerSocket").Formula = '"' + $Template.NumCoresPerSocket + '"'
	# MemoryGB
	$TemplateObject.Cells("Prop.MemoryGB").Formula = '"' + $Template.MemoryGB + '"'
	# MacAddress
	$TemplateObject.Cells("Prop.MacAddress").Formula = '"' + $Template.MacAddress + '"'
	# NumEthernetCards
	$TemplateObject.Cells("Prop.NumEthernetCards").Formula = '"' + $Template.NumEthernetCards + '"'
	# NumVirtualDisks
	$TemplateObject.Cells("Prop.NumVirtualDisks").Formula = '"' + $Template.NumVirtualDisks + '"'
	# CpuReservation
	$TemplateObject.Cells("Prop.CpuReservation").Formula = '"' + $Template.CpuReservation + '"'
	# MemoryReservation
	$TemplateObject.Cells("Prop.MemoryReservation").Formula = '"' + $Template.MemoryReservation + '"'
	# CpuHotAddEnabled
	$TemplateObject.Cells("Prop.CpuHotAddEnabled").Formula = '"' + $Template.CpuHotAddEnabled + '"'
	# CpuHotRemoveEnabled
	$TemplateObject.Cells("Prop.CpuHotRemoveEnabled").Formula = '"' + $Template.CpuHotRemoveEnabled + '"'
	# MemoryHotAddEnabled
	$TemplateObject.Cells("Prop.MemoryHotAddEnabled").Formula = '"' + $Template.MemoryHotAddEnabled + '"'
}
#endregion ~~< Draw_Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Folder
{
	#Name
	$FolderObject.Cells("Prop.Name").Formula = '"' + $Folder.Name + '"'
}
#endregion ~~< Draw_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_RDM
{
	# ScsiCanonicalName
	$RDMObject.Cells("Prop.ScsiCanonicalName").Formula = '"' + $HardDisk.ScsiCanonicalName + '"'
	# CapacityGB
	$RDMObject.Cells("Prop.CapacityGB").Formula = '"' + [math]::Round([decimal]$HardDisk.CapacityGB, 2) + '"'
	# DiskType
	$RDMObject.Cells("Prop.DiskType").Formula = '"' + $HardDisk.DiskType + '"'
	# CompatibilityMode
	$RDMObject.Cells("Prop.CompatibilityMode").Formula = '"' + $HardDisk.CompatibilityMode + '"'
	# DeviceName
	$RDMObject.Cells("Prop.DeviceName").Formula = '"' + $HardDisk.DeviceName + '"'
	# Sharing
	$RDMObject.Cells("Prop.Sharing").Formula = '"' + $HardDisk.Sharing + '"'
	# HardDisk
	$RDMObject.Cells("Prop.Label").Formula = '"' + $HardDisk.Label + '"'
	# Persistence
	$RDMObject.Cells("Prop.Persistence").Formula = '"' + $HardDisk.Persistence + '"'
}
#endregion ~~< Draw_RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_SRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SRM
{
	# Name
	$SrmObject.Cells("Prop.Name").Formula = '"' + $SrmVM.Name + '"'
	# OS
	$SrmObject.Cells("Prop.OS").Formula = '"' + $SrmVM.ConfigGuestFullName + '"'
	# Version
	$SrmObject.Cells("Prop.Version").Formula = '"' + $SrmVM.Version + '"'
	# Folder
	$SrmObject.Cells("Prop.Folder").Formula = '"' + $SrmVM.Folder + '"'
	# NumCPU
	$SrmObject.Cells("Prop.NumCPU").Formula = '"' + $SrmVM.NumCPU + '"'
	# CoresPerSocket
	$SrmObject.Cells("Prop.CoresPerSocket").Formula = '"' + $SrmVM.CoresPerSocket + '"'
	# MemoryGB
	$SrmObject.Cells("Prop.MemoryGB").Formula = '"' + $SrmVM.MemoryGB + '"'
}
#endregion ~~< Draw_SRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_DatastoreCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DatastoreCluster
{
	# Name
	$DatastoreClusObject.Cells("Prop.Name").Formula = '"' + $DatastoreCluster.Name + '"'
	# SdrsAutomationLevel
	$DatastoreClusObject.Cells("Prop.SdrsAutomationLevel").Formula = '"' + $DatastoreCluster.SdrsAutomationLevel + '"' 
	# IOLoadBalanceEnabled
	$DatastoreClusObject.Cells("Prop.IOLoadBalanceEnabled").Formula = '"' + $DatastoreCluster.IOLoadBalanceEnabled + '"'
	# CapacityGB
	$DatastoreClusObject.Cells("Prop.CapacityGB").Formula = '"' + $DatastoreCluster.CapacityGB + '"'
}
#endregion ~~< Draw_DatastoreCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Datastore
{
	# Name
	$DatastoreObject.Cells("Prop.Name").Formula = '"' + $Datastore.Name + '"'
	# Type
	$DatastoreObject.Cells("Prop.Type").Formula = '"' + $Datastore.Type + '"'
	# FileSystemVersion
	$DatastoreObject.Cells("Prop.FileSystemVersion").Formula = '"' + $Datastore.FileSystemVersion + '"'
	# DiskName
	$DatastoreObject.Cells("Prop.DiskName").Formula = '"' + $Datastore.DiskName + '"'
	# StorageIOControlEnabled
	$DatastoreObject.Cells("Prop.StorageIOControlEnabled").Formula = '"' + $Datastore.StorageIOControlEnabled + '"'
	# CapacityGB
	$DatastoreObject.Cells("Prop.CapacityGB").Formula = '"' + $Datastore.CapacityGB + '"'
	# FreeSpaceGB
	$DatastoreObject.Cells("Prop.FreeSpaceGB").Formula = '"' + $Datastore.FreeSpaceGB + '"'
	# Cluster
	$DatastoreObject.Cells("Prop.Cluster").Formula = '"' + $Datastore.Cluster + '"'
	# VmHost
	$DatastoreObject.Cells("Prop.VmHost").Formula = '"' + $Datastore.VmHost + '"'
	# Vm
	$DatastoreObject.Cells("Prop.Vm").Formula = '"' + $Datastore.Vm + '"'
	# State
	$DatastoreObject.Cells("Prop.State").Formula = '"' + $Datastore.State + '"'
}
#endregion ~~< Draw_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ResourcePool
{
	# Name
	$ResourcePoolObject.Cells("Prop.Name").Formula = '"' + $ResourcePool.Name + '"'
	# Cluster
	$ResourcePoolObject.Cells("Prop.Cluster").Formula = '"' + $ResourcePool.Cluster + '"'
	# CpuSharesLevel
	$ResourcePoolObject.Cells("Prop.CpuSharesLevel").Formula = '"' + $ResourcePool.CpuSharesLevel + '"'
	# NumCpuShares
	$ResourcePoolObject.Cells("Prop.NumCpuShares").Formula = '"' + $ResourcePool.NumCpuShares + '"'
	# CpuReservationMHz
	$ResourcePoolObject.Cells("Prop.CpuReservationMHz").Formula = '"' + $ResourcePool.CpuReservationMHz + '"'
	# CpuExpandableReservation
	$ResourcePoolObject.Cells("Prop.CpuExpandableReservation").Formula = '"' + $ResourcePool.CpuExpandableReservation + '"'
	# CpuLimitMHz
	$ResourcePoolObject.Cells("Prop.CpuLimitMHz").Formula = '"' + $ResourcePool.CpuLimitMHz + '"'
	# MemSharesLevel
	$ResourcePoolObject.Cells("Prop.MemSharesLevel").Formula = '"' + $ResourcePool.MemSharesLevel + '"'
	# NumMemShares
	$ResourcePoolObject.Cells("Prop.NumMemShares").Formula = '"' + $ResourcePool.NumMemShares + '"'
	# MemReservationGB
	$ResourcePoolObject.Cells("Prop.MemReservationGB").Formula = '"' + $ResourcePool.MemReservationGB + '"'
	# MemExpandableReservation
	$ResourcePoolObject.Cells("Prop.MemExpandableReservation").Formula = '"' + $ResourcePool.MemExpandableReservation + '"'
	# MemLimitGB
	$ResourcePoolObject.Cells("Prop.MemLimitGB").Formula = '"' + $ResourcePool.MemLimitGB + '"'
}
#endregion ~~< Draw_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VsSwitch
{
	# Name
	$VSSObject.Cells("Prop.Name").Formula = '"' + $VsSwitch.Name + '"'
	# NIC
	$VSSObject.Cells("Prop.NIC").Formula = '"' + $VsSwitch.Nic + '"'
	# NumPorts
	$VSSObject.Cells("Prop.NumPorts").Formula = '"' + $VsSwitch.NumPorts + '"'
	# SecurityAllowPromiscuous
	$VSSObject.Cells("Prop.AllowPromiscuous").Formula = '"' + $VsSwitch.AllowPromiscuous + '"'
	# SecurityMacChanges
	$VSSObject.Cells("Prop.MacChanges").Formula = '"' + $VsSwitch.MacChanges + '"'
	# SecurityForgedTransmits
	$VSSObject.Cells("Prop.ForgedTransmits").Formula = '"' + $VsSwitch.ForgedTransmits + '"'
	# NicTeamingPolicy
	$VSSObject.Cells("Prop.Policy").Formula = '"' + $VsSwitch.Policy + '"'
	# NicTeamingReversePolicy
	$VSSObject.Cells("Prop.ReversePolicy").Formula = '"' + $VsSwitch.ReversePolicy + '"'
	# NicTeamingNotifySwitches
	$VSSObject.Cells("Prop.NotifySwitches").Formula = '"' + $VsSwitch.NotifySwitches + '"'
	# NicTeamingRollingOrder
	$VSSObject.Cells("Prop.RollingOrder").Formula = '"' + $VsSwitch.RollingOrder + '"'
	# NicTeamingNicOrderActiveNic
	$VSSObject.Cells("Prop.ActiveNic").Formula = '"' + $VsSwitch.ActiveNic + '"'
	# NicTeamingNicOrderStandbyNic
	$VSSObject.Cells("Prop.StandbyNic").Formula = '"' + $VsSwitch.StandbyNic + '"'
}
#endregion ~~< Draw_VsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VssPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssPnic
{
	# Name
	$VssPNICObject.Cells("Prop.Name").Formula = '"' + $VssPnic.Name + '"'
	# Datacenter
	$VssPNICObject.Cells("Prop.Datacenter").Formula = '"' + $VssPnic.Datacenter + '"'
	# Cluster
	$VssPNICObject.Cells("Prop.Cluster").Formula = '"' + $VssPnic.Cluster + '"'
	# VmHost
	$VssPNICObject.Cells("Prop.VmHost").Formula = '"' + $VssPnic.VmHost + '"'
	# VsSwitch
	$VssPNICObject.Cells("Prop.VsSwitch").Formula = '"' + $VssPnic.VsSwitch + '"'
	# Mac
	$VssPNICObject.Cells("Prop.Mac").Formula = '"' + $VssPnic.Mac + '"'
}
#endregion ~~< Draw_VssPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VssPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssPort
{
	# Name
	$VssNicObject.Cells("Prop.Name").Formula = '"' + $VssPort.Name + '"'
	# VLanId
	$VssNicObject.Cells("Prop.VLanId").Formula = '"' + $VssPort.VLanId + '"'
	# ActiveNic
	$VssNicObject.Cells("Prop.ActiveNic").Formula = '"' + $VssPort.ActiveNic + '"'
	# StandbyNic
	$VssNicObject.Cells("Prop.StandbyNic").Formula = '"' + $VssPort.StandbyNic + '"'
}
#endregion ~~< Draw_VssPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VssVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssVmk
{
	# Name
	$VmkNicObject.Cells("Prop.Name").Formula = '"' + $VssVmk.Name + '"'
	# PortGroupName
	$VmkNicObject.Cells("Prop.PortGroupName").Formula = '"' + $VssVmk.PortGroupName + '"'
	# DhcpEnabled
	$VmkNicObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VssVmk.DhcpEnabled + '"'
	# IP
	$VmkNicObject.Cells("Prop.IP").Formula = '"' + $VssVmk.IP + '"'
	# Mac
	$VmkNicObject.Cells("Prop.Mac").Formula = '"' + $VssVmk.Mac + '"'
	# ManagementTrafficEnabled
	$VmkNicObject.Cells("Prop.ManagementTrafficEnabled").Formula = '"' + $VssVmk.ManagementTrafficEnabled + '"'
	# VMotionEnabled
	$VmkNicObject.Cells("Prop.VMotionEnabled").Formula = '"' + $VssVmk.VMotionEnabled + '"'
	# FaultToleranceLoggingEnabled
	$VmkNicObject.Cells("Prop.FaultToleranceLoggingEnabled").Formula = '"' + $VssVmk.FaultToleranceLoggingEnabled + '"'
	# VsanTrafficEnabled
	$VmkNicObject.Cells("Prop.VsanTrafficEnabled").Formula = '"' + $VssVmk.VsanTrafficEnabled + '"'
	# Mtu
	$VmkNicObject.Cells("Prop.Mtu").Formula = '"' + $VssVmk.Mtu + '"'
}
#endregion ~~< Draw_VssVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdSwitch
{
	# Name
	$VdSwitchObject.Cells("Prop.Name").Formula = '"' + $VdSwitch.Name + '"'
	# Vendor
	$VdSwitchObject.Cells("Prop.Vendor").Formula = '"' + $VdSwitch.Vendor + '"'
	# Version
	$VdSwitchObject.Cells("Prop.Version").Formula = '"' + $VdSwitch.Version + '"'
	# NumUplinkPorts
	$VdSwitchObject.Cells("Prop.NumUplinkPorts").Formula = '"' + $VdSwitch.NumUplinkPorts + '"'
	# UplinkPortName
	$VdSwitchObject.Cells("Prop.UplinkPortName").Formula = '"' + $VdSwitch.UplinkPortName + '"'
	# Mtu
	$VdSwitchObject.Cells("Prop.Mtu").Formula = '"' + $VdSwitch.Mtu + '"'
}
#endregion ~~< Draw_VdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VdsPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsPnic
{
	# Name
	$VdsPNICObject.Cells("Prop.Name").Formula = '"' + $VdsPnic.Name + '"'
	# Datacenter
	$VdsPNICObject.Cells("Prop.Datacenter").Formula = '"' + $VdsPnic.Datacenter + '"'
	# Cluster
	$VdsPNICObject.Cells("Prop.Cluster").Formula = '"' + $VdsPnic.Cluster + '"'
	# VmHost
	$VdsPNICObject.Cells("Prop.VmHost").Formula = '"' + $VdsPnic.VmHost + '"'
	# VdSwitch
	$VdsPNICObject.Cells("Prop.VdSwitch").Formula = '"' + $VdsPnic.VdSwitch + '"'
	# Portgroup
	$VdsPNICObject.Cells("Prop.Portgroup").Formula = '"' + $VdsPnic.Portgroup + '"'
	# ConnectedEntity
	$VdsPNICObject.Cells("Prop.ConnectedEntity").Formula = '"' + $VdsPnic.ConnectedEntity + '"'
	# VlanConfiguration
	$VdsPNICObject.Cells("Prop.VlanConfiguration").Formula = '"' + $VdsPnic.VlanConfiguration + '"'
}
#endregion ~~< Draw_VdsPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VdsPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsPort
{
	# Name
	$VPGObject.Cells("Prop.Name").Formula = '"' + $VdsPort.Name + '"'
	# VlanConfiguration
	$VPGObject.Cells("Prop.VlanConfiguration").Formula = '"' + $VdsPort.VlanConfiguration + '"'
	# NumPorts
	$VPGObject.Cells("Prop.NumPorts").Formula = '"' + $VdsPort.NumPorts + '"'
	# ActiveUplinkPort
	$VPGObject.Cells("Prop.ActiveUplinkPort").Formula = '"' + $VdsPort.ActiveUplinkPort + '"'
	# StandbyUplinkPort
	$VPGObject.Cells("Prop.StandbyUplinkPort").Formula = '"' + $VdsPort.StandbyUplinkPort + '"'
	# UplinkTeamingPolicy.Policy
	$VPGObject.Cells("Prop.Policy").Formula = '"' + $VdsPort.Policy + '"'
	# UplinkTeamingPolicy.ReversePolicy
	$VPGObject.Cells("Prop.ReversePolicy").Formula = '"' + $VdsPort.ReversePolicy + '"'
	#UplinkTeamingPolicy.NotifySwitches
	$VPGObject.Cells("Prop.NotifySwitches").Formula = '"' + $VdsPort.NotifySwitches + '"'
	# PortBinding
	$VPGObject.Cells("Prop.PortBinding").Formula = '"' + $VdsPort.PortBinding + '"'
}
#endregion ~~< Draw_VdsPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_VdsVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsVmk
{
	# Name
	$VmkNicObject.Cells("Prop.Name").Formula = '"' + $VdsVmk.Name + '"'
	# PortGroupName
	$VmkNicObject.Cells("Prop.PortGroupName").Formula = '"' + $VdsVmk.PortGroupName + '"'
	# DhcpEnabled
	$VmkNicObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VdsVmk.DhcpEnabled + '"'
	# IP
	$VmkNicObject.Cells("Prop.IP").Formula = '"' + $VdsVmk.IP + '"'
	# Mac
	$VmkNicObject.Cells("Prop.Mac").Formula = '"' + $VdsVmk.Mac + '"'
	# ManagementTrafficEnabled
	$VmkNicObject.Cells("Prop.ManagementTrafficEnabled").Formula = '"' + $VdsVmk.ManagementTrafficEnabled + '"'
	# VMotionEnabled
	$VmkNicObject.Cells("Prop.VMotionEnabled").Formula = '"' + $VdsVmk.VMotionEnabled + '"'
	# FaultToleranceLoggingEnabled
	$VmkNicObject.Cells("Prop.FaultToleranceLoggingEnabled").Formula = '"' + $VdsVmk.FaultToleranceLoggingEnabled + '"'
	# VsanTrafficEnabled
	$VmkNicObject.Cells("Prop.VsanTrafficEnabled").Formula = '"' + $VdsVmk.VsanTrafficEnabled + '"'
	# Mtu
	$VmkNicObject.Cells("Prop.Mtu").Formula = '"' + $VdsVmk.Mtu + '"'
}
#endregion ~~< Draw_VdsVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_DrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsRule
{
	# Name
	$DRSObject.Cells("Prop.Name").Formula = '"' + $DRSRule.Name + '"'
	# VM Affinity
	$DRSObject.Cells("Prop.Type").Formula = '"' + $DRSRule.Type + '"'
	# DRS Rule Enabled
	$DRSObject.Cells("Prop.Enabled").Formula = '"' + $DRSRule.Enabled + '"'
	# DRS Rule Mandatory
	$DRSObject.Cells("Prop.Mandatory").Formula = '"' + $DRSRule.Mandatory + '"'
}
#endregion ~~< Draw_DrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_DrsVmHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsVmHostRule
{
	# Name
	$DRSVMHostRuleObject.Cells("Prop.Name").Formula = '"' + $DrsVmHostRule.Name + '"'
	# Enabled
	$DRSVMHostRuleObject.Cells("Prop.Enabled").Formula = '"' + $DrsVmHostRule.Enabled + '"'
	# Type
	$DRSVMHostRuleObject.Cells("Prop.Type").Formula = '"' + $DrsVmHostRule.Type + '"'
	# VMGroup
	$DRSVMHostRuleObject.Cells("Prop.VMGroup").Formula = '"' + $DrsVmHostRule.VMGroup + '"'
	# VMHostGroup
	$DRSVMHostRuleObject.Cells("Prop.VMHostGroup").Formula = '"' + $DrsVmHostRule.VMHostGroup + '"'
	# AffineHostGroupName
	$DRSVMHostRuleObject.Cells("Prop.AffineHostGroupName").Formula = '"' + $DrsVmHostRule.ExtensionData.AffineHostGroupName + '"'
	# AntiAffineHostGroupName
	$DRSVMHostRuleObject.Cells("Prop.AntiAffineHostGroupName").Formula = '"' + $DrsVmHostRule.ExtensionData.AntiAffineHostGroupName + '"'
}
#endregion ~~< Draw_DrsVmHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_DrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsClusterGroup
{
	# Name
	$DrsClusterGroupObject.Cells("Prop.Name").Formula = '"' + $DrsClusterGroup.Name + '"'
	# GroupType
	$DrsClusterGroupObject.Cells("Prop.GroupType").Formula = '"' + $DrsClusterGroup.GroupType + '"'
	# Members
	$DrsClusterGroupObject.Cells("Prop.Member").Formula = '"' + $DrsClusterGroup.Member + '"'
}
#endregion ~~< Draw_DrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_ParentSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ParentSnapshot
{
	# Name
	$ParentSnapshotObject.Cells("Prop.Name").Formula = '"' + $ParentSnapshot.Name + '"'
	# Created
	$ParentSnapshotObject.Cells("Prop.Created").Formula = '"' + $ParentSnapshot.Created + '"'
	# Children
	$ParentSnapshotObject.Cells("Prop.Children").Formula = '"' + $ParentSnapshot.Children + '"'
	# ParentSnapshot
	$ParentSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ParentSnapshot.ParentSnapshot + '"'
	# IsCurrent
	$ParentSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ParentSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ParentSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_ChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildSnapshot
{
	# Name
	$ChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildSnapshot.Name + '"'
	# Created
	$ChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildSnapshot.Created + '"'
	# Children
	$ChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildSnapshot.Children + '"'
	# ParentSnapshot
	$ChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildSnapshot.ParentSnapshot + '"'
	# IsCurrent
	$ChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_ChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildChildSnapshot
{
	# Name
	$ChildChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildChildSnapshot.Name + '"'
	# Created
	$ChildChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildChildSnapshot.Created + '"'
	# Children
	$ChildChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildChildSnapshot.Children + '"'
	# ParentSnapshot
	$ChildChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildChildSnapshot.ParentSnapshot + '"'
	# IsCurrent
	$ChildChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_Draw_ChildChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildChildChildSnapshot
{
	# Name
	$ChildChildChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildChildChildSnapshot.Name + '"'
	# Created
	$ChildChildChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildChildChildSnapshot.Created + '"'
	# Children
	$ChildChildChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildChildChildSnapshot.Children + '"'
	# ParentSnapshot
	$ChildChildChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildChildChildSnapshot.ParentSnapshot + '"'
	# IsCurrent
	$ChildChildChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildChildChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_Draw_ChildChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Draw_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_LinkedvCenter
{
	# Name
	$LinkedvCenterObject.Cells("Prop.Name").Formula = '"' + $LinkedvCenter.Name + '"'
	# Version
	$LinkedvCenterObject.Cells("Prop.Version").Formula = '"' + $LinkedvCenter.Version + '"'
	# Build
	$LinkedvCenterObject.Cells("Prop.Build").Formula = '"' + $LinkedvCenter.Build + '"'
	# OsType
	$LinkedvCenterObject.Cells("Prop.OsType").Formula = '"' + $LinkedvCenter.OsType + '"'
}
#endregion ~~< Draw_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Visio Draw Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CSV >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function CSV_In_Out
{
	$global:DrawCsvFolder = $DrawCsvBrowse.SelectedPath
	# vCenter
	$global:vCenterExportFile = "$DrawCsvFolder\$vCenter-vCenterExport.csv"
	$global:vCenterImport = Import-Csv $vCenterExportFile
	# Datacenter
	$global:DatacenterExportFile = "$DrawCsvFolder\$vCenter-DatacenterExport.csv"
	$global:DatacenterImport = Import-Csv $DatacenterExportFile
	# Cluster
	$global:ClusterExportFile = "$DrawCsvFolder\$vCenter-ClusterExport.csv"
	$global:ClusterImport = Import-Csv $ClusterExportFile
	# VmHost
	$global:VmHostExportFile = "$DrawCsvFolder\$vCenter-VmHostExport.csv"
	$global:VmHostImport = Import-Csv $VmHostExportFile
	# Vm
	$global:VmExportFile = "$DrawCsvFolder\$vCenter-VmExport.csv"
	$global:VmImport = Import-Csv $VmExportFile
	#Template
	$global:TemplateExportFile = "$DrawCsvFolder\$vCenter-TemplateExport.csv"
	$global:TemplateImport = Import-Csv $TemplateExportFile
	# Folder
	$global:FolderExportFile = "$DrawCsvFolder\$vCenter-FolderExport.csv"
	$global:FolderImport = Import-Csv $FolderExportFile
	# Datastore Cluster
	$global:DatastoreClusterExportFile = "$DrawCsvFolder\$vCenter-DatastoreClusterExport.csv"
	$global:DatastoreClusterImport = Import-Csv $DatastoreClusterExportFile
	# Datastore
	$global:DatastoreExportFile = "$DrawCsvFolder\$vCenter-DatastoreExport.csv"
	$global:DatastoreImport = Import-Csv $DatastoreExportFile
	# RDM's
	$global:RdmExportFile = "$DrawCsvFolder\$vCenter-RdmExport.csv"
	$global:RdmImport = Import-Csv $RdmExportFile
	# ResourcePool
	$global:ResourcePoolExportFile = "$DrawCsvFolder\$vCenter-ResourcePoolExport.csv"
	$global:ResourcePoolImport = Import-Csv $ResourcePoolExportFile
	# Vss Switch
	$global:VsSwitchExportFile = "$DrawCsvFolder\$vCenter-VsSwitchExport.csv"
	$global:VsSwitchImport = Import-Csv $VsSwitchExportFile
	# Vss Port Group
	$global:VssPortExportFile = "$DrawCsvFolder\$vCenter-VssPortGroupExport.csv"
	$global:VssPortImport = Import-Csv $VssPortExportFile
	# Vss VMKernel
	$global:VssVmkExportFile = "$DrawCsvFolder\$vCenter-VssVmkernelExport.csv"
	$global:VssVmkImport = Import-Csv $VssVmkExportFile
	# Vss Pnic
	$global:VssPnicExportFile = "$DrawCsvFolder\$vCenter-VssPnicExport.csv"
	$global:VssPnicImport = Import-Csv $VssPnicExportFile
	# Vds Switch
	$global:VdSwitchExportFile = "$DrawCsvFolder\$vCenter-VdSwitchExport.csv"
	$global:VdSwitchImport = Import-Csv $VdSwitchExportFile
	# Vds Port Group
	$global:VdsPortExportFile = "$DrawCsvFolder\$vCenter-VdsPortGroupExport.csv"
	$global:VdsPortImport = Import-Csv $VdsPortExportFile
	# Vds VMKernel
	$global:VdsVmkExportFile = "$DrawCsvFolder\$vCenter-VdsVmkernelExport.csv"
	$global:VdsVmkImport = Import-Csv $VdsVmkExportFile
	# Vds Pnic
	$global:VdsPnicExportFile = "$DrawCsvFolder\$vCenter-VdsPnicExport.csv"
	$global:VdsPnicImport = Import-Csv $VdsPnicExportFile
	# DRS Rule
	$global:DrsRuleExportFile = "$DrawCsvFolder\$vCenter-DrsRuleExport.csv"
	$global:DrsRuleImport = Import-Csv $DrsRuleExportFile
	# DRS Cluster Group
	$global:DrsClusterGroupExportFile = "$DrawCsvFolder\$vCenter-DrsClusterGroupExport.csv"
	$global:DrsClusterGroupImport = Import-Csv $DrsClusterGroupExportFile
	# DRS VmHost Rule
	$global:DrsVmHostRuleExportFile = "$DrawCsvFolder\$vCenter-DrsVmHostRuleExport.csv"
	$global:DrsVmHostImport = Import-Csv $DrsVmHostRuleExportFile
	# Snapshot
	$global:SnapshotExportFile = "$DrawCsvFolder\$vCenter-SnapshotExport.csv"
	$global:SnapshotImport = Import-Csv $SnapshotExportFile
	# Linked vCenter
	$global:LinkedvCenterExportFile = "$DrawCsvFolder\$vCenter-LinkedvCenterExport.csv"
	$global:LinkedvCenterImport = Import-Csv $LinkedvCenterExportFile
}
#endregion ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< CSV >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Visio_Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Visio_Shapes
{
	$stnPath = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes"
	$stnObj = $AppVisio.Documents.Add($stnPath + $shpFile)
	# vCenter Object
	$global:VCObj = $stnObj.Masters.Item("Virtual Center Management Console")
	# Datacenter Object
	$global:DatacenterObj = $stnObj.Masters.Item("Datacenter")
	# Cluster Object
	$global:ClusterObj = $stnObj.Masters.Item("Cluster")
	# Host Object
	$global:HostObj = $stnObj.Masters.Item("ESX Host")
	# Microsoft VM Object
	$global:MicrosoftObj = $stnObj.Masters.Item("Microsoft Server")
	# Linux VM Object
	$global:LinuxObj = $stnObj.Masters.Item("Linux Server")
	# Other VM Object
	$global:OtherObj = $stnObj.Masters.Item("Other Server")
	# Template VM Object
	$global:TemplateObj = $stnObj.Masters.Item("Template")
	# Folder Object
	$global:FolderObj = $stnObj.Masters.Item("Folder")
	# RDM Object
	$global:RDMObj = $stnObj.Masters.Item("RDM")
	# SRM Protected VM Object
	$global:SRMObj = $stnObj.Masters.Item("SRM Protected Server")
	# Datastore Cluster Object
	$global:DatastoreClusObj = $stnObj.Masters.Item("Datastore Cluster")
	# Datastore Object
	$global:DatastoreObj = $stnObj.Masters.Item("Datastore")
	# Resource Pool Object
	$global:ResourcePoolObj = $stnObj.Masters.Item("Resource Pool")
	# VSS Object
	$global:VSSObj = $stnObj.Masters.Item("VSS")
	# VSS PNIC Object
	$global:VssPNICObj = $stnObj.Masters.Item("VSS Physical NIC")
	# VSSNIC Object
	$global:VssNicObj = $stnObj.Masters.Item("VSS NIC")
	# VDS Object
	$global:VDSObj = $stnObj.Masters.Item("VDS")
	# VDS PNIC Object
	$global:VdsPNICObj = $stnObj.Masters.Item("VDS Physical NIC")
	# VDSNIC Object
	$global:VdsNicObj = $stnObj.Masters.Item("VDS NIC")
	# VMK NIC Object
	$global:VmkNicObj = $stnObj.Masters.Item("VMKernel")
	# DRS Rule
	$global:DRSRuleObj = $stnObj.Masters.Item("DRS Rule")
	# DRS Cluster Group
	$global:DRSClusterGroupObj = $stnObj.Masters.Item("DRS Cluster group")
	# DRS Host Rule
	$global:DRSVMHostRuleObj = $stnObj.Masters.Item("DRS Host Rule")
	# Snapshot Object
	$global:SnapshotObj = $stnObj.Masters.Item("Snapshot")
	# Current Snapshot Object
	$global:CurrentSnapshotObj = $stnObj.Masters.Item("Current Snapshot")
}
#endregion ~~< Visio_Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Shapes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Create_Visio_Base >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Create_Visio_Base
{
	$global:vCenter = $VcenterTextBox.Text
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$DocObj = $docsObj.Add("")
	$DocObj.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< Create_Visio_Base >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenter_to_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function vCenter_to_LinkedvCenter
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "vCenter to Linked vCenters"
	$Page = $DocsObj.Pages('vCenter to Linked vCenters')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('vCenter to Linked vCenters')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	
	foreach ( $LinkedvCenter in ( $LinkedvCenterImport | Sort-Object Name ) )
	{
		$x += 2.50
		$LinkedvCenterObject = Add-VisioObjectVC $VCObj $LinkedvCenter
		#Draw_vCenter
		Draw_LinkedvCenter
		Connect-VisioObject $VCObject $LinkedvCenterObject
		$VCObject = $LinkedvCenterObject
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< vCenter_to_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Host
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VM to Host"
	$Page = $DocsObj.Pages('VM to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
						
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
				{
					$x += 2.50
					if ($VM.OS -eq "")
					{
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
					}
					else
					{
						if ($VM.OS.contains("Microsoft") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
						}
						else
						{
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
						}
					}	
					Connect-VisioObject $HostObject $VMObject
					$HostObject = $VMObject
				}
				foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template	
					Connect-VisioObject $HostObject $TemplateObject
					$HostObject = $TemplateObject
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50
						
			foreach ($VM in($VmImport | Sort-Object | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
			{
				$x += 2.50
				if ($VM.OS -eq "")
				{
					$VMObject = Add-VisioObjectVM $OtherObj $VM
					Draw_VM
				}
				else
				{
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
					}
				}	
				Connect-VisioObject $HostObject $VMObject
				$HostObject = $VMObject
			}
			foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) }))
			{
				$x += 2.50
				$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
				Draw_Template
				Connect-VisioObject $HostObject $TemplateObject
				$HostObject = $TemplateObject
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Folder
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VM to Folder"
	$Page = $DocsObj.Pages('VM to Folder')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Folder')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
				
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Folder in($FolderImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$FolderObject = Add-VisioObjectFolder $FolderObj $Folder
			Draw_Folder
			Connect-VisioObject $DatacenterObject $FolderObject
			$y += 1.50
				
			foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Folder.contains($Folder.Name) }))
			{
				$x += 2.50
				$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
				Draw_Template
				Connect-VisioObject $FolderObject $TemplateObject
				$FolderObject = $TemplateObject
			}
						
			foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Folder.contains($Folder.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
			{
				$x += 2.50
				if ($VM.OS -eq "")
				{
					$VMObject = Add-VisioObjectVM $OtherObj $VM
					Draw_VM
				}
				else
				{
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
					}
				}	
				Connect-VisioObject $FolderObject $VMObject
				$FolderObject = $VMObject
			}
		}
	}
				
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMs_with_RDMs
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VM w/ RDMs"
	$Page = $DocsObj.Pages('VM w/ RDMs')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM w/ RDMs')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
	
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.name) -and $RdmImport.Vm.contains($_.name) }))
			{
				$x = 6.00
				$y += 1.50
				if ($VM.OS -eq "")
				{
					$VMObject = Add-VisioObjectVM $OtherObj $VM
					Draw_VM
				}
				else
				{
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
					}
				}
				Connect-VisioObject $ClusterObject $VMObject
				$y += 1.50		
								
				foreach ($HardDisk in($RdmImport | Sort-Object Label | Where-Object { $_.Vm.contains($Vm.Name) }))
				{
					$x += 3.50
					$RDMObject = Add-VisioObjectHardDisk $RDMObj $HardDisk
					Draw_RDM
					Connect-VisioObject $VMObject $RDMObject
					$VMObject = $RDMObject
				}
			}		
		}	
		foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $RdmImport.Vm.contains($_.name) }))
		{
			$x = 6.00
			$y += 1.50
			if ($VM.OS -eq "")
			{
				$VMObject = Add-VisioObjectVM $OtherObj $VM
				Draw_VM
			}
			else
			{
				if ($VM.OS.contains("Microsoft") -eq $True)
				{
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
				}
				else
				{
					$VMObject = Add-VisioObjectVM $LinuxObj $VM
					Draw_VM
				}
			}
			Connect-VisioObject $DatacenterObject $VMObject
			$y += 1.50	
						
			foreach ($HardDisk in($RdmImport | Sort-Object Label | Where-Object { $_.Vm.contains($Vm.Name) }))
			{
				$x += 2.50
				$RDMObject = Add-VisioObjectHardDisk $RDMObj $HardDisk
				Draw_RDM
				Connect-VisioObject $VMObject $RDMObject
				$VMObject = $RDMObject
			}
		}		
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function SRM_Protected_VMs
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "SRM VM"
	$Page = $DocsObj.Pages('SRM VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('SRM VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50
								
				foreach ($SrmVM in($VmImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.SRM.contains("placeholderVm") }))
				{
					$x += 2.50
					$SrmObject = Add-VisioObjectSRM $SRMObj $SrmVM
					Draw_SRM
					Connect-VisioObject $HostObject $SrmObject
					$HostObject = $SrmObject
				}	
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50
						
			foreach ($SrmVM in($VmImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.SRM.contains("placeholderVm") }))
			{
				$x += 2.50
				$SrmObject = Add-VisioObjectSRM $SRMObj $SrmVM
				Draw_SRM
				Connect-VisioObject $HostObject $SrmObject
				$HostObject = $SrmObject
			}	
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Datastore
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VM to Datastore"
	$Page = $DocsObj.Pages('VM to Datastore')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Datastore')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($DatastoreCluster in($DatastoreClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$DatastoreClusObject = Add-VisioObjectDatastore $DatastoreClusObj $DatastoreCluster
				Draw_DatastoreCluster
				Connect-VisioObject $ClusterObject $DatastoreClusObject
									
				foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster.contains($DatastoreCluster.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
					Draw_Datastore
					Connect-VisioObject $DatastoreClusObject $DatastoreObject
					$y += 1.50
										
					foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster.contains($DatastoreCluster.Name) -and $_.Datastore.contains($Datastore.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
					{
						$x += 2.50
						if ($VM.OS.contains("Microsoft") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
						}
						else
						{
							if ($VM.OS.contains("Linux") -eq $True)
							{
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
							}
							else
							{
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
							}
						}	
						Connect-VisioObject $DatastoreObject $VMObject
						$DatastoreObject = $VMObject
					}
					foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster.contains($DatastoreCluster.Name) -and $_.Datastore.contains($Datastore.Name) }))
					{
						$x += 2.50
						$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
						Draw_Template
						Connect-VisioObject $DatastoreObject $TemplateObject
						$DatastoreObject = $TemplateObject
					}
				}
			}
			foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster -eq "" }))
			{
				$x = 8.00
				$y += 1.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				Connect-VisioObject $ClusterObject $DatastoreObject
				$y += 1.50
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster -eq "" -and $_.Datastore.contains($Datastore.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
				{
					$x += 2.50
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						if ($VM.OS.contains("Linux") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
						}
						else
						{
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
						}
					}	
					Connect-VisioObject $DatastoreObject $VMObject
					$DatastoreObject = $VMObject
				}
				foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster.contains($Cluster.Name) -and $_.DatastoreCluster -eq "" -and $_.Datastore.contains($Datastore.Name) }))
				{
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template	
					Connect-VisioObject $DatastoreObject $TemplateObject
					$DatastoreObject = $TemplateObject
				}
			}
		}
		foreach ($DatastoreCluster in($DatastoreClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$DatastoreClusObject = Add-VisioObjectDatastore $DatastoreClusObj $DatastoreCluster
			Draw_DatastoreCluster
			Connect-VisioObject $DatacenterObject $DatastoreClusObject
							
			foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.DatastoreCluster.contains($DatastoreCluster) }))
			{
				$x = 8.00
				$y += 1.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				Connect-VisioObject $DatastoreClusObject $DatastoreObject
				$y += 1.50
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.Datastore.contains($Datastore.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
				{
					$x += 2.50
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						if ($VM.OS.contains("Linux") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
						}
						else
						{
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
						}
					}	
					Connect-VisioObject $HostObject $VMObject
					$HostObject = $VMObject
				}
				foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.DatastoreCluster -eq "" -and $_.Datastore.contains($Datastore.Name) }))
				{
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template	
					Connect-VisioObject $HostObject $TemplateObject
					$HostObject = $TemplateObject
				}
			}
		}
		foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.DatastoreCluster -eq "" }))
		{
			$x = 8.00
			$y += 1.50
			$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
			Draw_Datastore
			Connect-VisioObject $DatacenterObject $DatastoreObject
			$y += 1.50
						
			foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.DatastoreCluster -eq "" -and $_.Datastore.contains($Datastore.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
			{
				$x += 2.50
				if ($VM.OS.contains("Microsoft") -eq $True)
				{
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
				}
				else
				{
					if ($VM.OS.contains("Linux") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
					}
					else
					{
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
					}
				}	
				Connect-VisioObject $DatastoreObject $VMObject
				$DatastoreObject = $VMObject
			}
			foreach ($Template in($TemplateImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" -and $_.DatastoreCluster -eq "" -and $_.Datastore.contains($Datastore.Name) }))
			{
				$x += 2.50
				$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
				Draw_Template
				Connect-VisioObject $DatastoreObject $TemplateObject
				$DatastoreObject = $TemplateObject
			}
		}	
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_ResourcePool
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VM to Resource Pool"
	$Page = $DocsObj.Pages('VM to Resource Pool')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Resource Pool')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($ResourcePool in($ResourcePoolImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) } ))
			{
				$x = 6.00
				$y += 1.50
				$ResourcePoolObject = Add-VisioObjectResourcePool $ResourcePoolObj $ResourcePool
				Draw_ResourcePool
				Connect-VisioObject $ClusterObject $ResourcePoolObject
				$y += 1.50
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.ResourcePool.contains($ResourcePool.Name) -and $_.Cluster.contains($Cluster.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
				{
					$x += 3.50
					if ($VM.OS.contains("Microsoft") -eq $True)
					{
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
					}
					else
					{
						if ($VM.OS.contains("Linux") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
						}
						else
						{
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
						}
					}	
					Connect-VisioObject $ResourcePoolObject $VMObject
					$ResourcePoolObject = $VMObject
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datastore_to_Host
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "Datastore to Host"
	$Page = $DocsObj.Pages('Datastore to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Datastore to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50
								
				foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x += 2.50
					$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
					Draw_Datastore
					Connect-VisioObject $HostObject $DatastoreObject
					$HostObject = $DatastoreObject
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50
						
			foreach ($Datastore in($DatastoreImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) }))
			{
				$x += 2.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				Connect-VisioObject $HostObject $DatastoreObject
				$HostObject = $DatastoreObject
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Snapshot_to_VM
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "Snapshot to VM"
	$Page = $DocsObj.Pages('Snapshot to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Snapshot to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
		
	foreach ($Datacenter in $DatacenterImport)
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($VM in($VmImport | Sort-Object Name | Where-Object { ($_.Snapshot -notlike "") }))
		{
			$x = 3.50
			$y += 1.50
			if ($VM.OS -eq "")
			{
				$VMObject = Add-VisioObjectVM $OtherObj $VM
				Draw_VM
			}
			else
			{
				if ($VM.OS.contains("Microsoft") -eq $True)
				{
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
				}
				else
				{
					$VMObject = Add-VisioObjectVM $LinuxObj $VM
					Draw_VM
				}
			}
			Connect-VisioObject $DatacenterObject $VMObject
			
			foreach ($ParentSnapshot in($SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains($VM.Name) -and ( $_.ParentSnapshot -like $null ) }))
			{
				$x = 6.00
				$y += 1.50
				if ($ParentSnapshot.IsCurrent -eq "FALSE")
				{
					$ParentSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ParentSnapshot
					Draw_ParentSnapshot
				}
				else
				{
					$ParentSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ParentSnapshot
					Draw_ParentSnapshot
				}
				Connect-VisioObject $VMObject $ParentSnapshotObject 
				
				foreach ($ChildSnapshot in($SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains($VM.Name) -and ($_.ParentSnapshot -like $ParentSnapshot.Name) }))
				{
					$x = 8.50
					$y += 1.50
					if ($ChildSnapshot.IsCurrent -eq "FALSE")
					{
						$ChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildSnapshot
						Draw_ChildSnapshot
					}
					else
					{
						$ChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildSnapshot
						Draw_ChildSnapshot
					}
					Connect-VisioObject $ParentSnapshotObject $ChildSnapshotObject
					
					foreach ($ChildChildSnapshot in($SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains($VM.Name) -and ($_.ParentSnapshot -like $ChildSnapshot.Name) }))
					{
						$x = 11.00
						$y += 1.50
						if ($ChildChildSnapshot.IsCurrent -eq "FALSE")
						{
							$ChildChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildChildSnapshot
							Draw_ChildChildSnapshot
						}
						else
						{
							$ChildChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildChildSnapshot
							Draw_ChildChildSnapshot
						} 
						Connect-VisioObject $ChildSnapshotObject $ChildChildSnapshotObject
						
						foreach ($ChildChildChildSnapshot in($SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains($VM.Name) -and ($_.ParentSnapshot -like $ChildChildSnapshot.Name) }))
						{
							$x += 2.50
							$y += 1.50
							if ($ChildChildChildSnapshot.IsCurrent -eq "FALSE")
							{
								$ChildChildChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildChildChildSnapshot
								Draw_ChildChildChildSnapshot
							}
							else
							{
								$ChildChildChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildChildChildSnapshot
								Draw_ChildChildChildSnapshot
							}
							Connect-VisioObject $ChildChildSnapshotObject $ChildChildChildSnapshotObject
							$ChildChildSnapshotObject = $ChildChildChildSnapshotObject	
						}
					}
				}
			}
		}	
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function PhysicalNIC_to_vSwitch
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
		
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "PNIC to switch"
	$Page = $DocsObj.Pages('PNIC to Switch')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('PNIC to Switch')
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ClusterObject.Cells("Prop.HostMonitoring").Formula = '"' + $Cluster.HostMonitoring + '"'
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VSSObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					Connect-VisioObject $HostObject $VSSObject
					$y += 1.50
										
					foreach ($VssPnic in($VssPnicImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch -eq $VsSwitch.Name }))
					{
						$x += 2.50
						$VssPNICObject = Add-VisioObjectVssPNIC $VssPNICObj $VssPnic
						Draw_VssPnic
						Connect-VisioObject $VSSObject $VssPNICObject
						$VSSObject = $VssPNICObject
					}
				}
				foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
										
					foreach ($VdsPnic in($VdsPnicImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VdSwitch.contains($VdSwitch.Name) }))
					{
						$x += 2.50
						$VdsPNICObject = Add-VisioObjectVdsPNIC $VdsPNICObj $VdsPnic
						Draw_VdsPnic
						Connect-VisioObject $VdSwitchObject $VdsPNICObject
						$VdSwitchObject = $VdsPNICObject
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VSSObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				Connect-VisioObject $HostObject $VSSObject
									
				foreach ($VssPnic in($VssPnicImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch -eq $VsSwitch.Name }))
				{
					$x += 2.50
					$VssPNICObject = Add-VisioObjectVssPNIC $VssPNICObj $VssPnic
					Draw_VssPnic
					Connect-VisioObject $VSSObject $VssPNICObject
					$VSSObject = $VssPNICObject
				}
			}
			foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50
								
				foreach ($VdsPnic in($VdsPnicImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VdSwitch.contains($VdSwitch.Name) }))
				{
					$x += 2.50
					$VdsPNICObject = Add-VisioObjectVdsPNIC $VdsPNICObj $VdsPnic
					Draw_VdsPnic
					Connect-VisioObject $VdSwitchObject $VdsPNICObject
					$VdSwitchObject = $VdsPNICObject
				}
			}
		}	
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VSS_to_Host
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
		
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VSS to Host"
	$Page = $DocsObj.Pages('VSS to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VSS to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					Connect-VisioObject $HostObject $VssObject
					$y += 1.50
										
					foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) }))
					{
						$x += 2.50
						$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
						Draw_VssPort
						Connect-VisioObject $VssObject $VssNicObject
						$VssObject = $VssNicObject
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				Connect-VisioObject $HostObject $VssObject
				$y += 1.50
								
				foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) }))
				{
					$x += 2.50
					$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
					Draw_VssPort
					Connect-VisioObject $VssObject $VssNicObject
					$VssObject = $VssNicObject
				}
			}
		}	
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMK_to_VSS
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VMK to VSS"
	$Page = $DocsObj.Pages('VMK to VSS')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VMK to VSS')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					Connect-VisioObject $HostObject $VssObject
					$y += 1.50
										
					foreach ($VssVmk in($VssVmkImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
					{
						$x += 1.50
						$VmkNicObject = Add-VisioObjectVMK $VmkNicObj $VssVmk
						Draw_VssVmk
						Connect-VisioObject $VssObject $VmkNicObject
						$VssObject = $VmkNicObject
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				Connect-VisioObject $HostObject $VssObject
				$y += 1.50
								
				foreach ($VssVmk in($VssVmkImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
				{
					$x += 1.50
					$VmkNicObject = Add-VisioObjectVMK $VmkNicObj $VssVmk
					Draw_VssVmk
					Connect-VisioObject $VssObject $VmkNicObject
					$VssObject = $VmkNicObject
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VSSPortGroup_to_VM
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VSSPortGroup to VM"
	$Page = $DocsObj.Pages('VSSPortGroup to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VSSPortGroup to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
						
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					Connect-VisioObject $HostObject $VssObject
										
					foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) }))
					{
						$x = 10.00
						$y += 1.50
						$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
						Draw_VssPort
						Connect-VisioObject $VssObject $VssNicObject
						$y += 1.50
												
						foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) -and $_.PortGroup.contains($VssPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
						{
							$x += 2.50
							if ($VM.OS -eq "")
							{
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
							}
							else
							{
								if ($VM.OS.contains("Microsoft") -eq $True)
								{
									$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
									Draw_VM
								}
								else
								{
									$VMObject = Add-VisioObjectVM $LinuxObj $VM
									Draw_VM
								}
							}	
							Connect-VisioObject $VssNicObject $VMObject
							$VssNicObject = $VMObject
						}
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VsSwitch in($VsSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VssObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				Connect-VisioObject $HostObject $VssObject
								
				foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) }))
				{
					$x = 10.00
					$y += 1.50
					$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
					Draw_VssPort
					Connect-VisioObject $VssObject $VssNicObject
					$y += 1.50
										
					foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and $_.vSwitch.contains($VsSwitch.Name) -and $_.PortGroup.contains($VssPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
					{
						$x += 2.50
						if ($VM.OS -eq "")
						{
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
						}
						else
						{
							if ($VM.OS.contains("Microsoft") -eq $True)
							{
								$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
								Draw_VM
							}
							else
							{
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
							}
						}	
						Connect-VisioObject $VssNicObject $VMObject
						$VssNicObject = $VMObject
					}
				}
			}
		}	
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VDS_to_Host
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VDS to Host"
	$Page = $DocsObj.Pages('VDS to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VDS to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
										
					foreach ($VdsPort in($VdsPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VdSwitch.VmHost) -and $_.VdSwitch.contains($VdSwitch.Name) }))
					{
						$x += 2.50
						$VPGObject = Add-VisioObjectVdsPG $VdsNicObj $VdsPort
						Draw_VdsPort
						Connect-VisioObject $VdSwitchObject $VPGObject
						$VdSwitchObject = $VPGObject
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50
								
				foreach ($VdsPort in($VdsPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VdSwitch.VmHost) -and $_.VdSwitch.contains($VdSwitch.Name) }))
				{
					$x += 2.50
					$VPGObject = Add-VisioObjectVdsPG $VdsNicObj $VdsPort
					Draw_VdsPort
					Connect-VisioObject $VdSwitchObject $VPGObject
					$VdSwitchObject = $VPGObject
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMK_to_VDS
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VMK to VDS"
	$Page = $DocsObj.Pages('VMK to VDS')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VMK to VDS')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
					
			foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				Connect-VisioObject $ClusterObject $HostObject
								
				foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) }))
				{
					$x = 8.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
										
					foreach ($VdsVmk in($VdsVmkImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VdSwitch.contains($VdSwitch.Name) }))
					{
						$x += 1.50
						$VmkNicObject = Add-VisioObjectVMK $VmkNicObj $VdsVmk
						Draw_VdsVmk
						Connect-VisioObject $VdSwitchObject $VmkNicObject
						$VdSwitchObject = $VmkNicObject
					}
				}
			}
		}
		foreach ($VmHost in($VmHostImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) -and $_.Cluster -eq "" }))
		{
			$x = 6.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			Connect-VisioObject $DatacenterObject $HostObject
						
			foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) }))
			{
				$x = 8.00
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50
								
				foreach ($VdsVmk in($VdsVmkImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VdSwitch.contains($VdSwitch.Name) }))
				{
					$x += 1.50
					$VmkNicObject = Add-VisioObjectVMK $VmkNicObj $VdsVmk
					Draw_VdsVmk
					Connect-VisioObject $VdSwitchObject $VmkNicObject
					$VdSwitchObject = $VmkNicObject
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VDSPortGroup_to_VM
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "VDSPortGroup to VM"
	$Page = $DocsObj.Pages('VDSPortGroup to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VDSPortGroup to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($VdSwitch in($VdSwitchImport | Sort-Object Name -Unique | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.00
			$y += 1.50
			$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
			Draw_VdSwitch
			Connect-VisioObject $DatacenterObject $VdSwitchObject
				
			foreach ($VdsPort in($VdsPortImport | Sort-Object Name -Unique | Where-Object { $_.VdSwitch.contains($VdSwitch.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$VPGObject = Add-VisioObjectVdsPG $VdsNicObj $VdsPort
				Draw_VdsPort
				Connect-VisioObject $VdSwitchObject $VPGObject
				$y += 1.50
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.vSwitch.contains($VdSwitch.Name) -and $_.PortGroup.contains($VdsPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
				{
					$x += 2.50
					if ($VM.OS -eq "")
					{
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
					}
					else
					{
						if ($VM.OS.contains("Microsoft") -eq $True)
						{
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
						}
						else
						{
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
						}
					}
					Connect-VisioObject $VPGObject $VMObject
					$VPGObject = $VMObject
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Cluster_to_DRS_Rule
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$Page = $AppVisio.ActivePage.Name = "Cluster to DRS Rule"
	$Page = $DocsObj.Pages('Cluster to DRS Rule')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Cluster to DRS Rule')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
		
	foreach ($Datacenter in ($DatacenterImport | Sort-Object Name) )
	{
		$x = 1.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		Connect-VisioObject $VCObject $DatacenterObject
				
		foreach ($Cluster in($ClusterImport | Sort-Object Name | Where-Object { $_.Datacenter.contains($Datacenter.Name) }))
		{
			$x = 3.50
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			Connect-VisioObject $DatacenterObject $ClusterObject
			$y += 1.50
						
			foreach ($DRSRule in($DrsRuleImport | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$DRSObject = Add-VisioObjectDrsRule $DRSRuleObj $DRSRule
				Draw_DrsRule
				Connect-VisioObject $ClusterObject $DRSObject
				$y += 1.50
			}		
			foreach ($DrsVmHostRule in($DrsVmHostImport | Where-Object { $_.Cluster.contains($Cluster.Name) }))
			{
				$x = 6.00
				$y += 1.50
				$DRSVMHostRuleObject = Add-VisioObjectDRSVMHostRule $DRSVMHostRuleObj $DrsVmHostRule
				Draw_DrsVmHostRule
				Connect-VisioObject $ClusterObject $DRSVMHostRuleObject
				$y += 1.50
				
				foreach ($DrsClusterGroup in($DrsClusterGroupImport | Where-Object { $_.Name.contains($DrsVmHostRule.VMHostGroup) }))
				{
					$x += 2.50
					$DrsClusterGroupObject = Add-VisioObjectDrsClusterGroup $DRSClusterGroupObj $DrsClusterGroup
					Draw_DrsClusterGroup
					Connect-VisioObject $DRSVMHostRuleObject $DrsClusterGroupObject
					$DRSVMHostRuleObject = $DrsClusterGroupObject
					
				}
			}
		}
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Quit()
}
#endregion ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Capture_Folder
{
	explorer.exe $CaptureCsvFolder
}
#endregion ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Final_Visio
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsd"
	$ConvertSaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$DateTime" + ".vsdx"
	$AppVisio = New-Object -ComObject Visio.Application
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Item(1).Delete(1) | Out-Null
	$AppVisio.Documents.SaveAs($SaveFile) | Out-Null
	$AppVisio.Documents.SaveAs($ConvertSaveFile) | Out-Null
	del $SaveFile
}
#endregion ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Event Handlers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False | Out-Null
Main
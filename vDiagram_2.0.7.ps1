<# 
.SYNOPSIS 
   vDiagram Visio Drawing Tool

.DESCRIPTION
   vDiagram Visio Drawing Tool

.NOTES 
   File Name	: vDiagram_2.0.7.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 2.0.7

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data
		MS Visio

.CHANGE LOG
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
$MyVer = "2.0.7"
$LastUpdated = "April 15, 2019"
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

#region ~~< Form Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vDiagram = New-Object System.Windows.Forms.Form
$vDiagram.ClientSize = New-Object System.Drawing.Size(1008, 661)
$CurrentLocation = Get-Location
$Icon = "$CurrentLocation" + "\vDiagram.ico"
$vDiagram.Icon = $Icon
$vDiagram.Text = "vDiagram " + $MyVer 
$vDiagram.BackColor = [System.Drawing.Color]::DarkCyan
#region ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu = New-Object System.Windows.Forms.MenuStrip
$MainMenu.Location = New-Object System.Drawing.Point(0, 0)
$MainMenu.Size = New-Object System.Drawing.Size(1008, 24)
$MainMenu.TabIndex = 1
$MainMenu.Text = "MainMenu"
#region ~~< ToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< File >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$FileToolStripMenuItem.Size = New-Object System.Drawing.Size(37, 20)
$FileToolStripMenuItem.Text = "File"
#endregion ~~< FileToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ExitToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$ExitToolStripMenuItem.Size = New-Object System.Drawing.Size(92, 22)
$ExitToolStripMenuItem.Text = "Exit"
$ExitToolStripMenuItem.Add_Click({$vDiagram.Close()})
#endregion ~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($ExitToolStripMenuItem)))
#endregion ~~< File >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Help >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HelpToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$HelpToolStripMenuItem.Size = New-Object System.Drawing.Size(44, 20)
$HelpToolStripMenuItem.Text = "Help"
#endregion ~~< HelpToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AboutToolStripMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$AboutToolStripMenuItem.Size = New-Object System.Drawing.Size(107, 22)
$AboutToolStripMenuItem.Text = "About"
$AboutToolStripMenuItem.Add_Click({About_Config})
#endregion ~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$HelpToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($AboutToolStripMenuItem)))
#endregion ~~< Help >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@($FileToolStripMenuItem, $HelpToolStripMenuItem)))
#endregion ~~< ToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< MainTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainTab = New-Object System.Windows.Forms.TabControl
$MainTab.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$MainTab.ItemSize = New-Object System.Drawing.Size(85, 20)
$MainTab.Location = New-Object System.Drawing.Point(10, 30)
$MainTab.Size = New-Object System.Drawing.Size(990, 98)
$MainTab.TabIndex = 0
$MainTab.Text = "MainTabs"
#region ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Prerequisites = New-Object System.Windows.Forms.TabPage
$Prerequisites.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Prerequisites.Location = New-Object System.Drawing.Point(4, 24)
$Prerequisites.Padding = New-Object System.Windows.Forms.Padding(3)
$Prerequisites.Size = New-Object System.Drawing.Size(982, 70)
$Prerequisites.TabIndex = 0
$Prerequisites.Text = "Prerequisites"
$Prerequisites.BackColor = [System.Drawing.Color]::LightGray
#region ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellLabel = New-Object System.Windows.Forms.Label
$PowershellLabel.Location = New-Object System.Drawing.Point(10, 15)
$PowershellLabel.Size = New-Object System.Drawing.Size(75, 20)
$PowershellLabel.TabIndex = 1
$PowershellLabel.Text = "Powershell:"
$Prerequisites.Controls.Add($PowershellLabel)
#endregion ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellInstalled = New-Object System.Windows.Forms.Label
$PowershellInstalled.Location = New-Object System.Drawing.Point(96, 15)
$PowershellInstalled.Size = New-Object System.Drawing.Size(350, 20)
$PowershellInstalled.TabIndex = 2
$PowershellInstalled.Text = ""
$PowershellInstalled.BackColor = [System.Drawing.Color]::LightGray
$Prerequisites.Controls.Add($PowershellInstalled)
#endregion ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCli Module >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleLabel = New-Object System.Windows.Forms.Label
$PowerCliModuleLabel.Location = New-Object System.Drawing.Point(10, 40)
$PowerCliModuleLabel.Size = New-Object System.Drawing.Size(110, 20)
$PowerCliModuleLabel.TabIndex = 3
$PowerCliModuleLabel.Text = "PowerCLI Module:"
$Prerequisites.Controls.Add($PowerCliModuleLabel)
#endregion ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleInstalled = New-Object System.Windows.Forms.Label
$PowerCliModuleInstalled.Location = New-Object System.Drawing.Point(128, 40)
$PowerCliModuleInstalled.Size = New-Object System.Drawing.Size(320, 20)
$PowerCliModuleInstalled.TabIndex = 4
$PowerCliModuleInstalled.Text = ""
$PowerCliModuleInstalled.BackColor = [System.Drawing.Color]::LightGray
$Prerequisites.Controls.Add($PowerCliModuleInstalled)
#endregion ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< PowerCli Module >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliLabel = New-Object System.Windows.Forms.Label
$PowerCliLabel.Location = New-Object System.Drawing.Point(450, 15)
$PowerCliLabel.Size = New-Object System.Drawing.Size(64, 20)
$PowerCliLabel.TabIndex = 5
$PowerCliLabel.Text = "PowerCLI:"
$Prerequisites.Controls.Add($PowerCliLabel)
#endregion ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliInstalled = New-Object System.Windows.Forms.Label
$PowerCliInstalled.Location = New-Object System.Drawing.Point(520, 15)
$PowerCliInstalled.Size = New-Object System.Drawing.Size(400, 20)
$PowerCliInstalled.TabIndex = 6
$PowerCliInstalled.Text = ""
$PowerCliInstalled.BackColor = [System.Drawing.Color]::LightGray
$Prerequisites.Controls.Add($PowerCliInstalled)
#endregion ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioLabel = New-Object System.Windows.Forms.Label
$VisioLabel.Location = New-Object System.Drawing.Point(450, 40)
$VisioLabel.Size = New-Object System.Drawing.Size(40, 20)
$VisioLabel.TabIndex = 7
$VisioLabel.Text = "Visio:"
$Prerequisites.Controls.Add($VisioLabel)
#endregion ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioInstalled = New-Object System.Windows.Forms.Label
$VisioInstalled.Location = New-Object System.Drawing.Point(490, 40)
$VisioInstalled.Size = New-Object System.Drawing.Size(320, 20)
$VisioInstalled.TabIndex = 8
$VisioInstalled.Text = ""
$VisioInstalled.BackColor = [System.Drawing.Color]::LightGray
$Prerequisites.Controls.Add($VisioInstalled)
#endregion ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainTab.Controls.Add($Prerequisites)
#endregion ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfo = New-Object System.Windows.Forms.TabPage
$vCenterInfo.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$vCenterInfo.Location = New-Object System.Drawing.Point(4, 24)
$vCenterInfo.Padding = New-Object System.Windows.Forms.Padding(3)
$vCenterInfo.Size = New-Object System.Drawing.Size(982, 70)
$vCenterInfo.TabIndex = 0
$vCenterInfo.Text = "vCenter Info"
$vCenterInfo.BackColor = [System.Drawing.Color]::LightGray
#region ~~< VcenterLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterLabel = New-Object System.Windows.Forms.Label
$VcenterLabel.Location = New-Object System.Drawing.Point(8, 11)
$VcenterLabel.Size = New-Object System.Drawing.Size(70, 20)
$VcenterLabel.TabIndex = 1
$VcenterLabel.Text = "vCenter:"
$vCenterInfo.Controls.Add($VcenterLabel)
#endregion ~~< VcenterLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VcenterTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterTextBox = New-Object System.Windows.Forms.TextBox
$VcenterTextBox.Location = New-Object System.Drawing.Point(78, 8)
$VcenterTextBox.Size = New-Object System.Drawing.Size(238, 21)
$VcenterTextBox.TabIndex = 2
$VcenterTextBox.Text = ""
$vCenterInfo.Controls.Add($VcenterTextBox)
#endregion ~~< VcenterTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameLabel = New-Object System.Windows.Forms.Label
$UserNameLabel.Location = New-Object System.Drawing.Point(324, 11)
$UserNameLabel.Size = New-Object System.Drawing.Size(70, 20)
$UserNameLabel.TabIndex = 3
$UserNameLabel.Text = "User Name:"
$vCenterInfo.Controls.Add($UserNameLabel)
#endregion ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameTextBox = New-Object System.Windows.Forms.TextBox
$UserNameTextBox.Location = New-Object System.Drawing.Point(402, 8)
$UserNameTextBox.Size = New-Object System.Drawing.Size(238, 21)
$UserNameTextBox.TabIndex = 4
$UserNameTextBox.Text = ""
$vCenterInfo.Controls.Add($UserNameTextBox)
#endregion ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Location = New-Object System.Drawing.Point(656, 11)
$PasswordLabel.Size = New-Object System.Drawing.Size(70, 20)
$PasswordLabel.TabIndex = 5
$PasswordLabel.Text = "Password:"
$vCenterInfo.Controls.Add($PasswordLabel)
#endregion ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Location = New-Object System.Drawing.Point(734, 8)
$PasswordTextBox.Size = New-Object System.Drawing.Size(238, 21)
$PasswordTextBox.TabIndex = 6
$PasswordTextBox.Text = ""
$PasswordTextBox.UseSystemPasswordChar = $true
$vCenterInfo.Controls.Add($PasswordTextBox)
#endregion ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$ConnectButton.Location = New-Object System.Drawing.Point(8, 37)
$ConnectButton.Size = New-Object System.Drawing.Size(345, 25)
$ConnectButton.TabIndex = 7
$ConnectButton.Text = "Connect to vCenter"
$ConnectButton.UseVisualStyleBackColor = $true
$vCenterInfo.Controls.Add($ConnectButton)
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainTab.Controls.Add($vCenterInfo)
#endregion ~~< vCenterInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainTab.SelectedIndex = 0
$vDiagram.Controls.Add($MainTab)
#endregion ~~< MainTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SubTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab = New-Object System.Windows.Forms.TabControl
$SubTab.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$SubTab.Location = New-Object System.Drawing.Point(10, 136)
$SubTab.Size = New-Object System.Drawing.Size(990, 512)
$SubTab.TabIndex = 0
$SubTab.Text = "SubTabs"
#region ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDirections = New-Object System.Windows.Forms.TabPage
$TabDirections.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$TabDirections.Location = New-Object System.Drawing.Point(4, 22)
$TabDirections.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDirections.Size = New-Object System.Drawing.Size(982, 486)
$TabDirections.TabIndex = 0
$TabDirections.Text = "Directions"
$TabDirections.UseVisualStyleBackColor = $true
#region ~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesHeading = New-Object System.Windows.Forms.Label
$PrerequisitesHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$PrerequisitesHeading.Location = New-Object System.Drawing.Point(8, 8)
$PrerequisitesHeading.Size = New-Object System.Drawing.Size(149, 23)
$PrerequisitesHeading.TabIndex = 0
$PrerequisitesHeading.Text = "Prerequisites Tab"
$TabDirections.Controls.Add($PrerequisitesHeading)
#endregion ~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesDirections = New-Object System.Windows.Forms.Label
$PrerequisitesDirections.Location = New-Object System.Drawing.Point(8, 32)
$PrerequisitesDirections.Size = New-Object System.Drawing.Size(900, 30)
$PrerequisitesDirections.TabIndex = 1
$PrerequisitesDirections.Text = "1. Verify that prerequisites are met on the "+[char]34+"Prerequisites"+[char]34+" tab."+[char]13+[char]10+"2. If not please install needed requirements."
$TabDirections.Controls.Add($PrerequisitesDirections)
#endregion ~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfoHeading = New-Object System.Windows.Forms.Label
$vCenterInfoHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$vCenterInfoHeading.Location = New-Object System.Drawing.Point(8, 72)
$vCenterInfoHeading.Size = New-Object System.Drawing.Size(149, 23)
$vCenterInfoHeading.TabIndex = 2
$vCenterInfoHeading.Text = "vCenter Info Tab"
$TabDirections.Controls.Add($vCenterInfoHeading)
#endregion ~~< vCenterInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfoDirections = New-Object System.Windows.Forms.Label
$vCenterInfoDirections.Location = New-Object System.Drawing.Point(8, 96)
$vCenterInfoDirections.Size = New-Object System.Drawing.Size(900, 70)
$vCenterInfoDirections.TabIndex = 3
$vCenterInfoDirections.Text = "1. Click on "+[char]34+"vCenter Info"+[char]34+" tab."+[char]13+[char]10+"2. Enter name of vCenter."+[char]13+[char]10+"3. Enter User Name and Password (password will be hashed and not plain text)."+[char]13+[char]10+"4. Click on "+[char]34+"Connect to vCenter"+[char]34+" button."
$TabDirections.Controls.Add($vCenterInfoDirections)
#endregion ~~< vCenterInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvHeading = New-Object System.Windows.Forms.Label
$CaptureCsvHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CaptureCsvHeading.Location = New-Object System.Drawing.Point(8, 176)
$CaptureCsvHeading.Size = New-Object System.Drawing.Size(216, 23)
$CaptureCsvHeading.TabIndex = 4
$CaptureCsvHeading.Text = "Capture CSVs for Visio Tab"
$TabDirections.Controls.Add($CaptureCsvHeading)
#endregion ~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureDirections = New-Object System.Windows.Forms.Label
$CaptureDirections.Location = New-Object System.Drawing.Point(8, 200)
$CaptureDirections.Size = New-Object System.Drawing.Size(900, 65)
$CaptureDirections.TabIndex = 5
$CaptureDirections.Text = "1. Click on "+[char]34+"Capture CSVs for Visio"+[char]34+" tab."+[char]13+[char]10+"2. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select folder where you would like to output the CSVs to."+[char]13+[char]10+"3. Select items you wish to grab data on."+[char]13+[char]10+"4. Click on "+[char]34+"Collect CSV Data"+[char]34+" button."
$TabDirections.Controls.Add($CaptureDirections)
#endregion ~~< CaptureDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawHeading = New-Object System.Windows.Forms.Label
$DrawHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DrawHeading.Location = New-Object System.Drawing.Point(8, 264)
$DrawHeading.Size = New-Object System.Drawing.Size(149, 23)
$DrawHeading.TabIndex = 6
$DrawHeading.Text = "Draw Visio Tab"
$TabDirections.Controls.Add($DrawHeading)
#endregion ~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawDirections = New-Object System.Windows.Forms.Label
$DrawDirections.Location = New-Object System.Drawing.Point(8, 288)
$DrawDirections.Size = New-Object System.Drawing.Size(900, 130)
$DrawDirections.TabIndex = 7
$DrawDirections.Text = "1. Click on "+[char]34+"Select Input Folder"+[char]34+" button and select location where CSVs can be found."+[char]13+[char]10+"2. Click on "+[char]34+"Check for CSVs"+[char]39+" button to validate presence of required files."+[char]13+[char]10+"3. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select where location where you would like to save the Visio drawing."+[char]13+[char]10+"4. Select drawing that you would like to produce."+[char]13+[char]10+"5. Click on "+[char]34+"Draw Visio"+[char]34+" button."+[char]13+[char]10+"6. Click on "+[char]34+"Open Visio Drawing"+[char]34+" button once "+[char]34+"Draw Visio"+[char]34+" button says it has completed."
$TabDirections.Controls.Add($DrawDirections)
#endregion ~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab.Controls.Add($TabDirections)
#endregion ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCapture = New-Object System.Windows.Forms.TabPage
$TabCapture.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$TabCapture.Location = New-Object System.Drawing.Point(4, 22)
$TabCapture.Padding = New-Object System.Windows.Forms.Padding(3)
$TabCapture.Size = New-Object System.Drawing.Size(982, 486)
$TabCapture.TabIndex = 3
$TabCapture.Text = "Capture CSVs for Visio"
$TabCapture.UseVisualStyleBackColor = $true
#region ~~< CSV >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputLabel = New-Object System.Windows.Forms.Label
$CaptureCsvOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CaptureCsvOutputLabel.Location = New-Object System.Drawing.Point(10, 10)
$CaptureCsvOutputLabel.Size = New-Object System.Drawing.Size(210, 25)
$CaptureCsvOutputLabel.TabIndex = 0
$CaptureCsvOutputLabel.Text = "CSV Output Folder:"
$TabCapture.Controls.Add($CaptureCsvOutputLabel)
#endregion ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton = New-Object System.Windows.Forms.Button
$CaptureCsvOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCsvOutputButton.Location = New-Object System.Drawing.Point(220, 10)
$CaptureCsvOutputButton.Size = New-Object System.Drawing.Size(750, 25)
$CaptureCsvOutputButton.TabIndex = 0
$CaptureCsvOutputButton.Text = "Select Output Folder"
$CaptureCsvOutputButton.UseVisualStyleBackColor = $false
$CaptureCsvOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureCsvOutputButton)
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$CaptureCsvBrowse.Description = "Select a directory"
$CaptureCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< CSV >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CheckBoxes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenterCsvCheckBox.Checked = $true
$vCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$vCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$vCenterCsvCheckBox.TabIndex = 1
$vCenterCsvCheckBox.Text = "Export vCenter Info"
$vCenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($vCenterCsvCheckBox)
#endregion ~~< vCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$vCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 40)
$vCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$vCenterCsvValidationComplete.TabIndex = 26
$vCenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($vCenterCsvValidationComplete)
#endregion ~~< vCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatacenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatacenterCsvCheckBox.Checked = $true
$DatacenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatacenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 60)
$DatacenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatacenterCsvCheckBox.TabIndex = 2
$DatacenterCsvCheckBox.Text = "Export Datacenter Info"
$DatacenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatacenterCsvCheckBox)
#endregion ~~< DatacenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatacenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatacenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 60)
$DatacenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatacenterCsvValidationComplete.TabIndex = 27
$DatacenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatacenterCsvValidationComplete)
#endregion ~~< DatacenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ClusterCsvCheckBox.Checked = $true
$ClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 80)
$ClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ClusterCsvCheckBox.TabIndex = 3
$ClusterCsvCheckBox.Text = "Export Cluster Info"
$ClusterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($ClusterCsvCheckBox)
#endregion ~~< ClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$ClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 80)
$ClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ClusterCsvValidationComplete.TabIndex = 28
$ClusterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($ClusterCsvValidationComplete)
#endregion ~~< ClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHostCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmHostCsvCheckBox.Checked = $true
$VmHostCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmHostCsvCheckBox.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmHostCsvCheckBox.TabIndex = 4
$VmHostCsvCheckBox.Text = "Export VmHost Info"
$VmHostCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VmHostCsvCheckBox)
#endregion ~~< VmHostCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHostCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmHostCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 100)
$VmHostCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationComplete.TabIndex = 29
$VmHostCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VmHostCsvValidationComplete)
#endregion ~~< VmHostCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmCsvCheckBox.Checked = $true
$VmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmCsvCheckBox.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmCsvCheckBox.TabIndex = 5
$VmCsvCheckBox.Text = "Export Vm Info"
$VmCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VmCsvCheckBox)
#endregion ~~< VmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 120)
$VmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmCsvValidationComplete.TabIndex = 30
$VmCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VmCsvValidationComplete)
#endregion ~~< VmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TemplateCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$TemplateCsvCheckBox.Checked = $true
$TemplateCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$TemplateCsvCheckBox.Location = New-Object System.Drawing.Point(10, 140)
$TemplateCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$TemplateCsvCheckBox.TabIndex = 6
$TemplateCsvCheckBox.Text = "Export Template Info"
$TemplateCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($TemplateCsvCheckBox)
#endregion ~~< TemplateCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TemplateCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidationComplete = New-Object System.Windows.Forms.Label
$TemplateCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 140)
$TemplateCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$TemplateCsvValidationComplete.TabIndex = 31
$TemplateCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($TemplateCsvValidationComplete)
#endregion ~~< TemplateCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreClusterCsvCheckBox.Checked = $true
$DatastoreClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 160)
$DatastoreClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreClusterCsvCheckBox.TabIndex = 7
$DatastoreClusterCsvCheckBox.Text = "Export Datastore Cluster Info"
$DatastoreClusterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatastoreClusterCsvCheckBox)
#endregion ~~< DatastoreClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 160)
$DatastoreClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreClusterCsvValidationComplete.TabIndex = 32
$DatastoreClusterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatastoreClusterCsvValidationComplete)
#endregion ~~< DatastoreClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreCsvCheckBox.Checked = $true
$DatastoreCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreCsvCheckBox.Location = New-Object System.Drawing.Point(10, 180)
$DatastoreCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreCsvCheckBox.TabIndex = 8
$DatastoreCsvCheckBox.Text = "Export Datastore Info"
$DatastoreCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatastoreCsvCheckBox)
#endregion ~~< DatastoreCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 180)
$DatastoreCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreCsvValidationComplete.TabIndex = 33
$DatastoreCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatastoreCsvValidationComplete)
#endregion ~~< DatastoreCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VsSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VsSwitchCsvCheckBox.Checked = $true
$VsSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VsSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 40)
$VsSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VsSwitchCsvCheckBox.TabIndex = 9
$VsSwitchCsvCheckBox.Text = "Export VsSwitch Info"
$VsSwitchCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VsSwitchCsvCheckBox)
#endregion ~~< VsSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VsSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 40)
$VsSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VsSwitchCsvValidationComplete.TabIndex = 34
$VsSwitchCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VsSwitchCsvValidationComplete)
#endregion ~~< VsSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPortGroupCsvCheckBox.Checked = $true
$VssPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 60)
$VssPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPortGroupCsvCheckBox.TabIndex = 10
$VssPortGroupCsvCheckBox.Text = "Export VSS Port Group Info"
$VssPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssPortGroupCsvCheckBox)
#endregion ~~< VssPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 60)
$VssPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPortGroupCsvValidationComplete.TabIndex = 35
$VssPortGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssPortGroupCsvValidationComplete)
#endregion ~~< VssPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssVmkernelCsvCheckBox.Checked = $true
$VssVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 80)
$VssVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssVmkernelCsvCheckBox.TabIndex = 11
$VssVmkernelCsvCheckBox.Text = "Export VSS Vmkernel Info"
$VssVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssVmkernelCsvCheckBox)
#endregion ~~< VssVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 80)
$VssVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssVmkernelCsvValidationComplete.TabIndex = 36
$VssVmkernelCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssVmkernelCsvValidationComplete)
#endregion ~~< VssVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPnicCsvCheckBox.Checked = $true
$VssPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 100)
$VssPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPnicCsvCheckBox.TabIndex = 12
$VssPnicCsvCheckBox.Text = "Export VSS Pnic Info"
$VssPnicCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssPnicCsvCheckBox)
#endregion ~~< VssPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 100)
$VssPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPnicCsvValidationComplete.TabIndex = 37
$VssPnicCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssPnicCsvValidationComplete)
#endregion ~~< VssPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdSwitchCsvCheckBox.Checked = $true
$VdSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 120)
$VdSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdSwitchCsvCheckBox.TabIndex = 13
$VdSwitchCsvCheckBox.Text = "Export VdSwitch Info"
$VdSwitchCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdSwitchCsvCheckBox)
#endregion ~~< VdSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 120)
$VdSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdSwitchCsvValidationComplete.TabIndex = 38
$VdSwitchCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdSwitchCsvValidationComplete)
#endregion ~~< VdSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPortGroupCsvCheckBox.Checked = $true
$VdsPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 140)
$VdsPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPortGroupCsvCheckBox.TabIndex = 14
$VdsPortGroupCsvCheckBox.Text = "Export VDS Port Group Info"
$VdsPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsPortGroupCsvCheckBox)
#endregion ~~< VdsPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 140)
$VdsPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPortGroupCsvValidationComplete.TabIndex = 39
$VdsPortGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsPortGroupCsvValidationComplete)
#endregion ~~< VdsPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsVmkernelCsvCheckBox.Checked = $true
$VdsVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 160)
$VdsVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsVmkernelCsvCheckBox.TabIndex = 15
$VdsVmkernelCsvCheckBox.Text = "Export VDS Vmkernel Info"
$VdsVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsVmkernelCsvCheckBox)
#endregion ~~< VdsVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 160)
$VdsVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsVmkernelCsvValidationComplete.TabIndex = 40
$VdsVmkernelCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsVmkernelCsvValidationComplete)
#endregion ~~< VdsVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPnicCsvCheckBox.Checked = $true
$VdsPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 180)
$VdsPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPnicCsvCheckBox.TabIndex = 16
$VdsPnicCsvCheckBox.Text = "Export VDS Pnic Info"
$VdsPnicCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsPnicCsvCheckBox)
#endregion ~~< VdsPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 180)
$VdsPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPnicCsvValidationComplete.TabIndex = 41
$VdsPnicCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsPnicCsvValidationComplete)
#endregion ~~< VdsPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< FolderCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$FolderCsvCheckBox.Checked = $true
$FolderCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FolderCsvCheckBox.Location = New-Object System.Drawing.Point(620, 40)
$FolderCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$FolderCsvCheckBox.TabIndex = 17
$FolderCsvCheckBox.Text = "Export Folder Info"
$FolderCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($FolderCsvCheckBox)
#endregion ~~< FolderCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< FolderCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidationComplete = New-Object System.Windows.Forms.Label
$FolderCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 40)
$FolderCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$FolderCsvValidationComplete.TabIndex = 42
$FolderCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($FolderCsvValidationComplete)
#endregion ~~< FolderCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< RdmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$RdmCsvCheckBox.Checked = $true
$RdmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$RdmCsvCheckBox.Location = New-Object System.Drawing.Point(620, 60)
$RdmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$RdmCsvCheckBox.TabIndex = 18
$RdmCsvCheckBox.Text = "Export RDM Info"
$RdmCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($RdmCsvCheckBox)
#endregion ~~< RdmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< RdmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidationComplete = New-Object System.Windows.Forms.Label
$RdmCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 60)
$RdmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$RdmCsvValidationComplete.TabIndex = 43
$RdmCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($RdmCsvValidationComplete)
#endregion ~~< RdmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsRuleCsvCheckBox.Checked = $true
$DrsRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 80)
$DrsRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsRuleCsvCheckBox.TabIndex = 19
$DrsRuleCsvCheckBox.Text = "Export DRS Rule Info"
$DrsRuleCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsRuleCsvCheckBox)
#endregion ~~< DrsRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 80)
$DrsRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsRuleCsvValidationComplete.TabIndex = 44
$DrsRuleCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsRuleCsvValidationComplete)
#endregion ~~< DrsRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsClusterGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsClusterGroupCsvCheckBox.Checked = $true
$DrsClusterGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsClusterGroupCsvCheckBox.Location = New-Object System.Drawing.Point(620, 100)
$DrsClusterGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsClusterGroupCsvCheckBox.TabIndex = 20
$DrsClusterGroupCsvCheckBox.Text = "Export DRS Cluster Group Info"
$DrsClusterGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsClusterGroupCsvCheckBox)
#endregion ~~< DrsClusterGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsClusterGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 100)
$DrsClusterGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsClusterGroupCsvValidationComplete.TabIndex = 45
$DrsClusterGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsClusterGroupCsvValidationComplete)
#endregion ~~< DrsClusterGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsVmHostRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsVmHostRuleCsvCheckBox.Checked = $true
$DrsVmHostRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsVmHostRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 120)
$DrsVmHostRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsVmHostRuleCsvCheckBox.TabIndex = 21
$DrsVmHostRuleCsvCheckBox.Text = "Export DRS VmHost Rule Info"
$DrsVmHostRuleCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsVmHostRuleCsvCheckBox)
#endregion ~~< DrsVmHostRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsVmHostRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 120)
$DrsVmHostRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsVmHostRuleCsvValidationComplete.TabIndex = 46
$DrsVmHostRuleCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsVmHostRuleCsvValidationComplete)
#endregion ~~< DrsVmHostRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ResourcePoolCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ResourcePoolCsvCheckBox.Checked = $true
$ResourcePoolCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ResourcePoolCsvCheckBox.Location = New-Object System.Drawing.Point(620, 140)
$ResourcePoolCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ResourcePoolCsvCheckBox.TabIndex = 22
$ResourcePoolCsvCheckBox.Text = "Export Resource Pool Info"
$ResourcePoolCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($ResourcePoolCsvCheckBox)
#endregion ~~< ResourcePoolCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ResourcePoolCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidationComplete = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 140)
$ResourcePoolCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ResourcePoolCsvValidationComplete.TabIndex = 47
$ResourcePoolCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($ResourcePoolCsvValidationComplete)
#endregion ~~< ResourcePoolCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SnapshotCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$SnapshotCsvCheckBox.Checked = $true
$SnapshotCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SnapshotCsvCheckBox.Location = New-Object System.Drawing.Point(620, 160)
$SnapshotCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$SnapshotCsvCheckBox.TabIndex = 23
$SnapshotCsvCheckBox.Text = "Export Snapshot Info"
$SnapshotCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($SnapshotCsvCheckBox)
#endregion ~~< SnapshotCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SnapshotCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidationComplete = New-Object System.Windows.Forms.Label
$SnapshotCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 160)
$SnapshotCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$SnapshotCsvValidationComplete.TabIndex = 48
$SnapshotCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($SnapshotCsvValidationComplete)
#endregion ~~< SnapshotCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< LinkedvCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$LinkedvCenterCsvCheckBox.Checked = $true
$LinkedvCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$LinkedvCenterCsvCheckBox.Location = New-Object System.Drawing.Point(620, 180)
$LinkedvCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$LinkedvCenterCsvCheckBox.TabIndex = 24
$LinkedvCenterCsvCheckBox.Text = "Export Linked vCenter Info"
$LinkedvCenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($LinkedvCenterCsvCheckBox)
#endregion ~~< LinkedvCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< LinkedvCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 180)
$LinkedvCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$LinkedvCenterCsvValidationComplete.TabIndex = 49
$LinkedvCenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($LinkedvCenterCsvValidationComplete)
#endregion ~~< LinkedvCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< CheckBoxes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton = New-Object System.Windows.Forms.Button
$CaptureUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureUncheckButton.Location = New-Object System.Drawing.Point(8, 215)
$CaptureUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureUncheckButton.TabIndex = 23
$CaptureUncheckButton.Text = "Uncheck All"
$CaptureUncheckButton.UseVisualStyleBackColor = $false
$CaptureUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureUncheckButton)
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton = New-Object System.Windows.Forms.Button
$CaptureCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCheckButton.Location = New-Object System.Drawing.Point(228, 215)
$CaptureCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureCheckButton.TabIndex = 24
$CaptureCheckButton.Text = "Check All"
$CaptureCheckButton.UseVisualStyleBackColor = $false
$CaptureCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureCheckButton)
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton = New-Object System.Windows.Forms.Button
$CaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureButton.Location = New-Object System.Drawing.Point(448, 215)
$CaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureButton.TabIndex = 25
$CaptureButton.Text = "Collect CSV Data"
$CaptureButton.UseVisualStyleBackColor = $false
$CaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureButton)
#endregion ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton = New-Object System.Windows.Forms.Button
$OpenCaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenCaptureButton.Location = New-Object System.Drawing.Point(668, 215)
$OpenCaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenCaptureButton.TabIndex = 83
$OpenCaptureButton.Text = "Open CSV Output Folder"
$OpenCaptureButton.UseVisualStyleBackColor = $false
$OpenCaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($OpenCaptureButton)
#endregion ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab.Controls.Add($TabCapture)
#endregion ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDraw = New-Object System.Windows.Forms.TabPage
$TabDraw.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$TabDraw.Location = New-Object System.Drawing.Point(4, 22)
$TabDraw.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDraw.Size = New-Object System.Drawing.Size(982, 486)
$TabDraw.TabIndex = 2
$TabDraw.Text = "Draw Visio"
$TabDraw.UseVisualStyleBackColor = $true
#region ~~< Input >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputLabel = New-Object System.Windows.Forms.Label
$DrawCsvInputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DrawCsvInputLabel.Location = New-Object System.Drawing.Point(10, 10)
$DrawCsvInputLabel.Size = New-Object System.Drawing.Size(190, 25)
$DrawCsvInputLabel.TabIndex = 0
$DrawCsvInputLabel.Text = "CSV Input Folder:"
$TabDraw.Controls.Add($DrawCsvInputLabel)
#endregion ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton = New-Object System.Windows.Forms.Button
$DrawCsvInputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCsvInputButton.Location = New-Object System.Drawing.Point(220, 10)
$DrawCsvInputButton.Size = New-Object System.Drawing.Size(750, 25)
$DrawCsvInputButton.TabIndex = 1
$DrawCsvInputButton.Text = "Select CSV Input Folder"
$DrawCsvInputButton.UseVisualStyleBackColor = $false
$DrawCsvInputButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawCsvInputButton)
#endregion ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$DrawCsvBrowse.Description = "Select a directory"
$DrawCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Input >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CSV Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidation = New-Object System.Windows.Forms.Label
$vCenterCsvValidation.Location = New-Object System.Drawing.Point(10, 40)
$vCenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$vCenterCsvValidation.TabIndex = 2
$vCenterCsvValidation.Text = "vCenter CSV File:"
$TabDraw.Controls.Add($vCenterCsvValidation)
#endregion ~~< vCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$vCenterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 40)
$vCenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$vCenterCsvValidationCheck.TabIndex = 3
$vCenterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($vCenterCsvValidationCheck)
#endregion ~~< vCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatacenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidation = New-Object System.Windows.Forms.Label
$DatacenterCsvValidation.Location = New-Object System.Drawing.Point(10, 60)
$DatacenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatacenterCsvValidation.TabIndex = 4
$DatacenterCsvValidation.Text = "Datacenter CSV File:"
$TabDraw.Controls.Add($DatacenterCsvValidation)
#endregion ~~< DatacenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatacenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatacenterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 60)
$DatacenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatacenterCsvValidationCheck.TabIndex = 5
$DatacenterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DatacenterCsvValidationCheck)
#endregion ~~< DatacenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidation = New-Object System.Windows.Forms.Label
$ClusterCsvValidation.Location = New-Object System.Drawing.Point(10, 80)
$ClusterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$ClusterCsvValidation.TabIndex = 6
$ClusterCsvValidation.Text = "Cluster CSV File:"
$TabDraw.Controls.Add($ClusterCsvValidation)
#endregion ~~< ClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidationCheck = New-Object System.Windows.Forms.Label
$ClusterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 80)
$ClusterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ClusterCsvValidationCheck.TabIndex = 7
$ClusterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($ClusterCsvValidationCheck)
#endregion ~~< ClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHostCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidation = New-Object System.Windows.Forms.Label
$VmHostCsvValidation.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmHostCsvValidation.TabIndex = 8
$VmHostCsvValidation.Text = "VmHost CSV File:"
$TabDraw.Controls.Add($VmHostCsvValidation)
#endregion ~~< VmHostCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHostCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationCheck = New-Object System.Windows.Forms.Label
$VmHostCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 100)
$VmHostCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationCheck.TabIndex = 9
$VmHostCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VmHostCsvValidationCheck)
#endregion ~~< VmHostCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidation = New-Object System.Windows.Forms.Label
$VmCsvValidation.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmCsvValidation.TabIndex = 10
$VmCsvValidation.Text = "VM CSV File:"
$TabDraw.Controls.Add($VmCsvValidation)
#endregion ~~< VmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidationCheck = New-Object System.Windows.Forms.Label
$VmCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 120)
$VmCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VmCsvValidationCheck.TabIndex = 11
$VmCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VmCsvValidationCheck)
#endregion ~~< VmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TemplateCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidation = New-Object System.Windows.Forms.Label
$TemplateCsvValidation.Location = New-Object System.Drawing.Point(10, 140)
$TemplateCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$TemplateCsvValidation.TabIndex = 12
$TemplateCsvValidation.Text = "Template CSV File:"
$TabDraw.Controls.Add($TemplateCsvValidation)
#endregion ~~< TemplateCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< TemplateCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidationCheck = New-Object System.Windows.Forms.Label
$TemplateCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 140)
$TemplateCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$TemplateCsvValidationCheck.TabIndex = 13
$TemplateCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($TemplateCsvValidationCheck)
#endregion ~~< TemplateCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidation = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidation.Location = New-Object System.Drawing.Point(10, 160)
$DatastoreClusterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatastoreClusterCsvValidation.TabIndex = 14
$DatastoreClusterCsvValidation.Text = "Datastore Cluster CSV File:"
$TabDraw.Controls.Add($DatastoreClusterCsvValidation)
#endregion ~~< DatastoreClusterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 160)
$DatastoreClusterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreClusterCsvValidationCheck.TabIndex = 15
$DatastoreClusterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DatastoreClusterCsvValidationCheck)
#endregion ~~< DatastoreClusterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidation = New-Object System.Windows.Forms.Label
$DatastoreCsvValidation.Location = New-Object System.Drawing.Point(10, 180)
$DatastoreCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DatastoreCsvValidation.TabIndex = 16
$DatastoreCsvValidation.Text = "Datastore CSV File:"
$TabDraw.Controls.Add($DatastoreCsvValidation)
#endregion ~~< DatastoreCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidationCheck = New-Object System.Windows.Forms.Label
$DatastoreCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 180)
$DatastoreCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreCsvValidationCheck.TabIndex = 17
$DatastoreCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DatastoreCsvValidationCheck)
#endregion ~~< DatastoreCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VsSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 40)
$VsSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VsSwitchCsvValidation.TabIndex = 18
$VsSwitchCsvValidation.Text = "VsSwitch CSV File:"
$TabDraw.Controls.Add($VsSwitchCsvValidation)
#endregion ~~< VsSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VsSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidationCheck = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 40)
$VsSwitchCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VsSwitchCsvValidationCheck.TabIndex = 19
$VsSwitchCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VsSwitchCsvValidationCheck)
#endregion ~~< VsSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidation = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidation.Location = New-Object System.Drawing.Point(270, 60)
$VssPortGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssPortGroupCsvValidation.TabIndex = 20
$VssPortGroupCsvValidation.Text = "Vss Port Group CSV File:"
$TabDraw.Controls.Add($VssPortGroupCsvValidation)
#endregion ~~< VssPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 60)
$VssPortGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssPortGroupCsvValidationCheck.TabIndex = 21
$VssPortGroupCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VssPortGroupCsvValidationCheck)
#endregion ~~< VssPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 80)
$VssVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssVmkernelCsvValidation.TabIndex = 22
$VssVmkernelCsvValidation.Text = "Vss Vmkernel CSV File:"
$TabDraw.Controls.Add($VssVmkernelCsvValidation)
#endregion ~~< VssVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 80)
$VssVmkernelCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssVmkernelCsvValidationCheck.TabIndex = 23
$VssVmkernelCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VssVmkernelCsvValidationCheck)
#endregion ~~< VssVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidation = New-Object System.Windows.Forms.Label
$VssPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 100)
$VssPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssPnicCsvValidation.TabIndex = 24
$VssPnicCsvValidation.Text = "Vss Pnic CSV File:"
$TabDraw.Controls.Add($VssPnicCsvValidation)
#endregion ~~< VssPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VssPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidationCheck = New-Object System.Windows.Forms.Label
$VssPnicCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 100)
$VssPnicCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VssPnicCsvValidationCheck.TabIndex = 25
$VssPnicCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VssPnicCsvValidationCheck)
#endregion ~~< VssPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 120)
$VdSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdSwitchCsvValidation.TabIndex = 26
$VdSwitchCsvValidation.Text = "VdSwitch CSV File:"
$TabDraw.Controls.Add($VdSwitchCsvValidation)
#endregion ~~< VdSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 120)
$VdSwitchCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdSwitchCsvValidationCheck.TabIndex = 27
$VdSwitchCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VdSwitchCsvValidationCheck)
#endregion ~~< VdSwitchCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidation = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidation.Location = New-Object System.Drawing.Point(270, 140)
$VdsPortGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsPortGroupCsvValidation.TabIndex = 28
$VdsPortGroupCsvValidation.Text = "Vds Port Group CSV File:"
$TabDraw.Controls.Add($VdsPortGroupCsvValidation)
#endregion ~~< VdsPortGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 140)
$VdsPortGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsPortGroupCsvValidationCheck.TabIndex = 29
$VdsPortGroupCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VdsPortGroupCsvValidationCheck)
#endregion ~~< VdsPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 160)
$VdsVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsVmkernelCsvValidation.TabIndex = 30
$VdsVmkernelCsvValidation.Text = "Vds Vmkernel CSV File:"
$TabDraw.Controls.Add($VdsVmkernelCsvValidation)
#endregion ~~< VdsVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 160)
$VdsVmkernelCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsVmkernelCsvValidationCheck.TabIndex = 31
$VdsVmkernelCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VdsVmkernelCsvValidationCheck)
#endregion ~~< VdsVmkernelCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidation = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 180)
$VdsPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsPnicCsvValidation.TabIndex = 32
$VdsPnicCsvValidation.Text = "Vds Pnic CSV File:"
$TabDraw.Controls.Add($VdsPnicCsvValidation)
#endregion ~~< VdsPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidationCheck = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidationCheck.Location = New-Object System.Drawing.Point(440, 180)
$VdsPnicCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VdsPnicCsvValidationCheck.TabIndex = 33
$VdsPnicCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($VdsPnicCsvValidationCheck)
#endregion ~~< VdsPnicCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< FolderCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidation = New-Object System.Windows.Forms.Label
$FolderCsvValidation.Location = New-Object System.Drawing.Point(530, 40)
$FolderCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$FolderCsvValidation.TabIndex = 34
$FolderCsvValidation.Text = "Folder CSV File:"
$TabDraw.Controls.Add($FolderCsvValidation)
#endregion ~~< FolderCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< FolderCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidationCheck = New-Object System.Windows.Forms.Label
$FolderCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 40)
$FolderCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$FolderCsvValidationCheck.TabIndex = 35
$FolderCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($FolderCsvValidationCheck)
#endregion ~~< FolderCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< RdmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidation = New-Object System.Windows.Forms.Label
$RdmCsvValidation.Location = New-Object System.Drawing.Point(530, 60)
$RdmCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$RdmCsvValidation.TabIndex = 36
$RdmCsvValidation.Text = "RDM CSV File:"
$TabDraw.Controls.Add($RdmCsvValidation)
#endregion ~~< RdmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< RdmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidationCheck = New-Object System.Windows.Forms.Label
$RdmCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 60)
$RdmCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$RdmCsvValidationCheck.TabIndex = 37
$RdmCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($RdmCsvValidationCheck)
#endregion ~~< RdmCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidation = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidation.Location = New-Object System.Drawing.Point(530, 80)
$DrsRuleCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsRuleCsvValidation.TabIndex = 38
$DrsRuleCsvValidation.Text = "DRS Rule CSV File:"
$TabDraw.Controls.Add($DrsRuleCsvValidation)
#endregion ~~< DrsRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 80)
$DrsRuleCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsRuleCsvValidationCheck.TabIndex = 39
$DrsRuleCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DrsRuleCsvValidationCheck)
#endregion ~~< DrsRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsClusterGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidation = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidation.Location = New-Object System.Drawing.Point(530, 100)
$DrsClusterGroupCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsClusterGroupCsvValidation.TabIndex = 40
$DrsClusterGroupCsvValidation.Text = "DRS Cluster Group CSV File:"
$TabDraw.Controls.Add($DrsClusterGroupCsvValidation)
#endregion ~~< DrsClusterGroupCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsClusterGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 100)
$DrsClusterGroupCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsClusterGroupCsvValidationCheck.TabIndex = 41
$DrsClusterGroupCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DrsClusterGroupCsvValidationCheck)
#endregion ~~< DrsClusterGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsVmHostRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidation = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidation.Location = New-Object System.Drawing.Point(530, 120)
$DrsVmHostRuleCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$DrsVmHostRuleCsvValidation.TabIndex = 42
$DrsVmHostRuleCsvValidation.Text = "DRS VmHost Rule CSV File:"
$TabDraw.Controls.Add($DrsVmHostRuleCsvValidation)
#endregion ~~< DrsVmHostRuleCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrsVmHostRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidationCheck = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 120)
$DrsVmHostRuleCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$DrsVmHostRuleCsvValidationCheck.TabIndex = 43
$DrsVmHostRuleCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($DrsVmHostRuleCsvValidationCheck)
#endregion ~~< DrsVmHostRuleCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ResourcePoolCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidation = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidation.Location = New-Object System.Drawing.Point(530, 140)
$ResourcePoolCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$ResourcePoolCsvValidation.TabIndex = 44
$ResourcePoolCsvValidation.Text = "Resource Pool CSV File:"
$TabDraw.Controls.Add($ResourcePoolCsvValidation)
#endregion ~~< ResourcePoolCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< ResourcePoolCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidationCheck = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 140)
$ResourcePoolCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$ResourcePoolCsvValidationCheck.TabIndex = 45
$ResourcePoolCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($ResourcePoolCsvValidationCheck)
#endregion ~~< ResourcePoolCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SnapshotCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidation = New-Object System.Windows.Forms.Label
$SnapshotCsvValidation.Location = New-Object System.Drawing.Point(530, 160)
$SnapshotCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$SnapshotCsvValidation.TabIndex = 46
$SnapshotCsvValidation.Text = "Snapshot CSV File:"
$TabDraw.Controls.Add($SnapshotCsvValidation)
#endregion ~~< SnapshotCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SnapshotCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidationCheck = New-Object System.Windows.Forms.Label
$SnapshotCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 160)
$SnapshotCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$SnapshotCsvValidationCheck.TabIndex = 47
$SnapshotCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($SnapshotCsvValidationCheck)
#endregion ~~< SnapshotCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< LinkedvCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidation = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidation.Location = New-Object System.Drawing.Point(530, 180)
$LinkedvCenterCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$LinkedvCenterCsvValidation.TabIndex = 48
$LinkedvCenterCsvValidation.Text = "Linked vCenter CSV File:"
$TabDraw.Controls.Add($LinkedvCenterCsvValidation)
#endregion ~~< LinkedvCenterCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< LinkedvCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvValidationCheck = New-Object System.Windows.Forms.Label
$LinkedvCenterCsvValidationCheck.Location = New-Object System.Drawing.Point(700, 180)
$LinkedvCenterCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$LinkedvCenterCsvValidationCheck.TabIndex = 3
$LinkedvCenterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($LinkedvCenterCsvValidationCheck)
#endregion ~~< LinkedvCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton = New-Object System.Windows.Forms.Button
$CsvValidationButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CsvValidationButton.Location = New-Object System.Drawing.Point(8, 200)
$CsvValidationButton.Size = New-Object System.Drawing.Size(200, 25)
$CsvValidationButton.TabIndex = 2
$CsvValidationButton.Text = "Check for CSVs"
$CsvValidationButton.UseVisualStyleBackColor = $false
$CsvValidationButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($CsvValidationButton)
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< CSV Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Output >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOutputLabel = New-Object System.Windows.Forms.Label
$VisioOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$VisioOutputLabel.Location = New-Object System.Drawing.Point(10, 230)
$VisioOutputLabel.Size = New-Object System.Drawing.Size(215, 25)
$VisioOutputLabel.TabIndex = 49
$VisioOutputLabel.Text = "Visio Output Folder:"
$TabDraw.Controls.Add($VisioOutputLabel)
#endregion ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton = New-Object System.Windows.Forms.Button
$VisioOpenOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$VisioOpenOutputButton.Location = New-Object System.Drawing.Point(230, 230)
$VisioOpenOutputButton.Size = New-Object System.Drawing.Size(740, 25)
$VisioOpenOutputButton.TabIndex = 50
$VisioOpenOutputButton.Text = "Select Visio Output Folder"
$VisioOpenOutputButton.UseVisualStyleBackColor = $false
$VisioOpenOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($VisioOpenOutputButton)
#endregion ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$VisioBrowse.Description = "Select a directory"
$VisioBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Output >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< CheckBoxes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenter_to_LinkedvCenter_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenter_to_LinkedvCenter_DrawCheckBox.Checked = $true
$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenter_to_LinkedvCenter_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 260)
$vCenter_to_LinkedvCenter_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$vCenter_to_LinkedvCenter_DrawCheckBox.TabIndex = 51
$vCenter_to_LinkedvCenter_DrawCheckBox.Text = "Create vCenter to Linked vCenter Visio Drawing"
$vCenter_to_LinkedvCenter_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_DrawCheckBox)
#endregion ~~< vCenter_to_LinkedvCenter_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< vCenter_to_LinkedvCenter_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_Complete = New-Object System.Windows.Forms.Label
$vCenter_to_LinkedvCenter_Complete.Location = New-Object System.Drawing.Point(315, 260)
$vCenter_to_LinkedvCenter_Complete.Size = New-Object System.Drawing.Size(90, 20)
$vCenter_to_LinkedvCenter_Complete.TabIndex = 52
$vCenter_to_LinkedvCenter_Complete.Text = ""
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)
#endregion ~~< vCenter_to_LinkedvCenter_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Host_DrawCheckBox.Checked = $true
$VM_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 280)
$VM_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Host_DrawCheckBox.TabIndex = 53
$VM_to_Host_DrawCheckBox.Text = "Create VM to Host Visio Drawing"
$VM_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Host_DrawCheckBox)
#endregion ~~< VM_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_Complete = New-Object System.Windows.Forms.Label
$VM_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 280)
$VM_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Host_Complete.TabIndex = 54
$VM_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Host_Complete)
#endregion ~~< VM_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Folder_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Folder_DrawCheckBox.Checked = $true
$VM_to_Folder_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Folder_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 300)
$VM_to_Folder_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Folder_DrawCheckBox.TabIndex = 55
$VM_to_Folder_DrawCheckBox.Text = "Create VM to Folder Visio Drawing"
$VM_to_Folder_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Folder_DrawCheckBox)
#endregion ~~< VM_to_Folder_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Folder_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_Complete = New-Object System.Windows.Forms.Label
$VM_to_Folder_Complete.Location = New-Object System.Drawing.Point(315, 300)
$VM_to_Folder_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Folder_Complete.TabIndex = 56
$VM_to_Folder_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Folder_Complete)
#endregion ~~< VM_to_Folder_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMs_with_RDMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMs_with_RDMs_DrawCheckBox.Checked = $true
$VMs_with_RDMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMs_with_RDMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 320)
$VMs_with_RDMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VMs_with_RDMs_DrawCheckBox.TabIndex = 57
$VMs_with_RDMs_DrawCheckBox.Text = "Create VMs with RDMs Visio Drawing"
$VMs_with_RDMs_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMs_with_RDMs_DrawCheckBox)
#endregion ~~< VMs_with_RDMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMs_with_RDMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_Complete = New-Object System.Windows.Forms.Label
$VMs_with_RDMs_Complete.Location = New-Object System.Drawing.Point(315, 320)
$VMs_with_RDMs_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMs_with_RDMs_Complete.TabIndex = 58
$VMs_with_RDMs_Complete.Text = ""
$TabDraw.Controls.Add($VMs_with_RDMs_Complete)
#endregion ~~< VMs_with_RDMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SRM_Protected_VMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$SRM_Protected_VMs_DrawCheckBox.Checked = $true
$SRM_Protected_VMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SRM_Protected_VMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 340)
$SRM_Protected_VMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$SRM_Protected_VMs_DrawCheckBox.TabIndex = 59
$SRM_Protected_VMs_DrawCheckBox.Text = "Create SRM Protected VMs Visio Drawing"
$SRM_Protected_VMs_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($SRM_Protected_VMs_DrawCheckBox)
#endregion ~~< SRM_Protected_VMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< SRM_Protected_VMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_Complete = New-Object System.Windows.Forms.Label
$SRM_Protected_VMs_Complete.Location = New-Object System.Drawing.Point(315, 340)
$SRM_Protected_VMs_Complete.Size = New-Object System.Drawing.Size(90, 20)
$SRM_Protected_VMs_Complete.TabIndex = 60
$SRM_Protected_VMs_Complete.Text = ""
$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)
#endregion ~~< SRM_Protected_VMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Datastore_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Datastore_DrawCheckBox.Checked = $true
$VM_to_Datastore_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Datastore_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 360)
$VM_to_Datastore_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Datastore_DrawCheckBox.TabIndex = 61
$VM_to_Datastore_DrawCheckBox.Text = "Create VM to Datastore Visio Drawing"
$VM_to_Datastore_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Datastore_DrawCheckBox)
#endregion ~~< VM_to_Datastore_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_Datastore_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_Complete = New-Object System.Windows.Forms.Label
$VM_to_Datastore_Complete.Location = New-Object System.Drawing.Point(315, 360)
$VM_to_Datastore_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_Datastore_Complete.TabIndex = 62
$VM_to_Datastore_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Datastore_Complete)
#endregion ~~< VM_to_Datastore_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_ResourcePool_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_ResourcePool_DrawCheckBox.Checked = $true
$VM_to_ResourcePool_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_ResourcePool_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 380)
$VM_to_ResourcePool_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_ResourcePool_DrawCheckBox.TabIndex = 63
$VM_to_ResourcePool_DrawCheckBox.Text = "Create VM to ResourcePool Visio Drawing"
$VM_to_ResourcePool_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_ResourcePool_DrawCheckBox)
#endregion ~~< VM_to_ResourcePool_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VM_to_ResourcePool_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_Complete = New-Object System.Windows.Forms.Label
$VM_to_ResourcePool_Complete.Location = New-Object System.Drawing.Point(315, 380)
$VM_to_ResourcePool_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VM_to_ResourcePool_Complete.TabIndex = 64
$VM_to_ResourcePool_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)
#endregion ~~< VM_to_ResourcePool_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datastore_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Datastore_to_Host_DrawCheckBox.Checked = $true
$Datastore_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Datastore_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 400)
$Datastore_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Datastore_to_Host_DrawCheckBox.TabIndex = 65
$Datastore_to_Host_DrawCheckBox.Text = "Create Datastore to Host Visio Drawing"
$Datastore_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Datastore_to_Host_DrawCheckBox)
#endregion ~~< Datastore_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datastore_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_Complete = New-Object System.Windows.Forms.Label
$Datastore_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 400)
$Datastore_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Datastore_to_Host_Complete.TabIndex = 66
$Datastore_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($Datastore_to_Host_Complete)
#endregion ~~< Datastore_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Snapshot_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Snapshot_to_VM_DrawCheckBox.Checked = $true
$Snapshot_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Snapshot_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 420)
$Snapshot_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Snapshot_to_VM_DrawCheckBox.TabIndex = 67
$Snapshot_to_VM_DrawCheckBox.Text = "Create Snapshot to VM Visio Drawing"
$Snapshot_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Snapshot_to_VM_DrawCheckBox)
#endregion ~~< Snapshot_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Snapshot_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_Complete = New-Object System.Windows.Forms.Label
$Snapshot_to_VM_Complete.Location = New-Object System.Drawing.Point(315, 420)
$Snapshot_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Snapshot_to_VM_Complete.TabIndex = 68
$Snapshot_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($Snapshot_to_VM_Complete)
#endregion ~~< Snapshot_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PhysicalNIC_to_vSwitch_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$PhysicalNIC_to_vSwitch_DrawCheckBox.Checked = $true
$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$PhysicalNIC_to_vSwitch_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 260)
$PhysicalNIC_to_vSwitch_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$PhysicalNIC_to_vSwitch_DrawCheckBox.TabIndex = 69
$PhysicalNIC_to_vSwitch_DrawCheckBox.Text = "Create PhysicalNIC to vSwitch Visio Drawing"
$PhysicalNIC_to_vSwitch_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_DrawCheckBox)
#endregion ~~< PhysicalNIC_to_vSwitch_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< PhysicalNIC_to_vSwitch_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_Complete = New-Object System.Windows.Forms.Label
$PhysicalNIC_to_vSwitch_Complete.Location = New-Object System.Drawing.Point(760, 260)
$PhysicalNIC_to_vSwitch_Complete.Size = New-Object System.Drawing.Size(90, 20)
$PhysicalNIC_to_vSwitch_Complete.TabIndex = 70
$PhysicalNIC_to_vSwitch_Complete.Text = ""
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)
#endregion ~~< PhysicalNIC_to_vSwitch_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSS_to_Host_DrawCheckBox.Checked = $true
$VSS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 280)
$VSS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSS_to_Host_DrawCheckBox.TabIndex = 71
$VSS_to_Host_DrawCheckBox.Text = "Create VSS to Host Visio Drawing"
$VSS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VSS_to_Host_DrawCheckBox)
#endregion ~~< VSS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VSS_to_Host_Complete.Location = New-Object System.Drawing.Point(760, 280)
$VSS_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VSS_to_Host_Complete.TabIndex = 72
$VSS_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VSS_to_Host_Complete)
#endregion ~~< VSS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VSS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VSS_DrawCheckBox.Checked = $true
$VMK_to_VSS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VSS_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 300)
$VMK_to_VSS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VSS_DrawCheckBox.TabIndex = 73
$VMK_to_VSS_DrawCheckBox.Text = "Create Vmkernel to VSS Visio Drawing"
$VMK_to_VSS_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMK_to_VSS_DrawCheckBox)
#endregion ~~< VMK_to_VSS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VSS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VSS_Complete.Location = New-Object System.Drawing.Point(760, 300)
$VMK_to_VSS_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMK_to_VSS_Complete.TabIndex = 74
$VMK_to_VSS_Complete.Text = ""
$TabDraw.Controls.Add($VMK_to_VSS_Complete)
#endregion ~~< VMK_to_VSS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VSSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 320)
$VSSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSSPortGroup_to_VM_DrawCheckBox.TabIndex = 75
$VSSPortGroup_to_VM_DrawCheckBox.Text = "Create Vss Port Group to VM Visio Drawing"
$VSSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VSSPortGroup_to_VM_DrawCheckBox)
#endregion ~~< VSSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VSSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VSSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(760, 320)
$VSSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VSSPortGroup_to_VM_Complete.TabIndex = 76
$VSSPortGroup_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)
#endregion ~~< VSSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDS_to_Host_DrawCheckBox.Checked = $true
$VDS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 340)
$VDS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDS_to_Host_DrawCheckBox.TabIndex = 77
$VDS_to_Host_DrawCheckBox.Text = "Create VDS to Host Visio Drawing"
$VDS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VDS_to_Host_DrawCheckBox)
#endregion ~~< VDS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VDS_to_Host_Complete.Location = New-Object System.Drawing.Point(760, 340)
$VDS_to_Host_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VDS_to_Host_Complete.TabIndex = 78
$VDS_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VDS_to_Host_Complete)
#endregion ~~< VDS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VDS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VDS_DrawCheckBox.Checked = $true
$VMK_to_VDS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VDS_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 360)
$VMK_to_VDS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VDS_DrawCheckBox.TabIndex = 79
$VMK_to_VDS_DrawCheckBox.Text = "Create Vmkernel to VDS Visio Drawing"
$VMK_to_VDS_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMK_to_VDS_DrawCheckBox)
#endregion ~~< VMK_to_VDS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VMK_to_VDS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VDS_Complete.Location = New-Object System.Drawing.Point(760, 360)
$VMK_to_VDS_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VMK_to_VDS_Complete.TabIndex = 80
$VMK_to_VDS_Complete.Text = ""
$TabDraw.Controls.Add($VMK_to_VDS_Complete)
#endregion ~~< VMK_to_VDS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VDSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 380)
$VDSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDSPortGroup_to_VM_DrawCheckBox.TabIndex = 81
$VDSPortGroup_to_VM_DrawCheckBox.Text = "Create Vds Port Group to VM Visio Drawing"
$VDSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VDSPortGroup_to_VM_DrawCheckBox)
#endregion ~~< VDSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VDSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VDSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(760, 380)
$VDSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(90, 20)
$VDSPortGroup_to_VM_Complete.TabIndex = 82
$VDSPortGroup_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)
#endregion ~~< VDSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Cluster_to_DRS_Rule_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Cluster_to_DRS_Rule_DrawCheckBox.Checked = $true
$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Cluster_to_DRS_Rule_DrawCheckBox.Location = New-Object System.Drawing.Point(425, 400)
$Cluster_to_DRS_Rule_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$Cluster_to_DRS_Rule_DrawCheckBox.TabIndex = 83
$Cluster_to_DRS_Rule_DrawCheckBox.Text = "Create Cluster to DRS Rule Visio Drawing"
$Cluster_to_DRS_Rule_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_DrawCheckBox)
#endregion ~~< Cluster_to_DRS_Rule_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Cluster_to_DRS_Rule_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_Complete = New-Object System.Windows.Forms.Label
$Cluster_to_DRS_Rule_Complete.Location = New-Object System.Drawing.Point(760, 400)
$Cluster_to_DRS_Rule_Complete.Size = New-Object System.Drawing.Size(90, 20)
$Cluster_to_DRS_Rule_Complete.TabIndex = 84
$Cluster_to_DRS_Rule_Complete.Text = ""
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
#endregion ~~< Cluster_to_DRS_Rule_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< CheckBoxes >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton = New-Object System.Windows.Forms.Button
$DrawUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawUncheckButton.Location = New-Object System.Drawing.Point(8, 450)
$DrawUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawUncheckButton.TabIndex = 80
$DrawUncheckButton.Text = "Uncheck All"
$DrawUncheckButton.UseVisualStyleBackColor = $false
$DrawUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawUncheckButton)
#endregion ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton = New-Object System.Windows.Forms.Button
$DrawCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCheckButton.Location = New-Object System.Drawing.Point(228, 450)
$DrawCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawCheckButton.TabIndex = 82
$DrawCheckButton.Text = "Check All"
$DrawCheckButton.UseVisualStyleBackColor = $false
$DrawCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawCheckButton)
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton = New-Object System.Windows.Forms.Button
$DrawButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawButton.Location = New-Object System.Drawing.Point(448, 450)
$DrawButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawButton.TabIndex = 81
$DrawButton.Text = "Draw Visio"
$DrawButton.UseVisualStyleBackColor = $false
$DrawButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawButton)
#endregion ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton = New-Object System.Windows.Forms.Button
$OpenVisioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenVisioButton.Location = New-Object System.Drawing.Point(668, 450)
$OpenVisioButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenVisioButton.TabIndex = 83
$OpenVisioButton.Text = "Open Visio Drawing"
$OpenVisioButton.UseVisualStyleBackColor = $false
$OpenVisioButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($OpenVisioButton)
#endregion ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab.Controls.Add($TabDraw)
#endregion ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SubTab.ForeColor = [System.Drawing.SystemColors]::ControlText
$SubTab.SelectedIndex = 0
$vDiagram.Controls.Add($SubTab)
#endregion ~~< SubTab >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vDiagram.Controls.Add($MainMenu)
#endregion ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#endregion ~~< Form Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Version" ; E = { $_.Version } }, 
	@{ N = "Build" ; E = { $_.Build } },
	@{ N = "OsType" ; E = { $_.ExtensionData.Content.About.OsType } } | Export-Csv $vCenterExportFile -Append -NoTypeInformation
}
#endregion ~~< vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datacenter_Export
{
	$DatacenterExportFile = "$CaptureCsvFolder\$vCenter-DatacenterExport.csv"
	Get-Datacenter | Sort-Object Name | 
	Select-Object Name | Export-Csv $DatacenterExportFile -Append -NoTypeInformation
}
#endregion ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Cluster_Export
{
	$ClusterExportFile = "$CaptureCsvFolder\$vCenter-ClusterExport.csv"
	Get-Cluster | Sort-Object Name | 
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Datacenter" ; E = { Get-Cluster $_.Name | Get-Datacenter } }, 
	@{ N = "HAEnabled" ; E = { $_.HAEnabled } }, 
	@{ N = "HAAdmissionControlEnabled" ; E = { $_.HAAdmissionControlEnabled } }, 
	@{ N = "AdmissionControlPolicyCpuFailoverResourcesPercent" ; E = { $_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.CpuFailoverResourcesPercent } }, 
	@{ N = "AdmissionControlPolicyMemoryFailoverResourcesPercent" ; E = { $_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.MemoryFailoverResourcesPercent } }, 
	@{ N = "AdmissionControlPolicyFailoverLevel" ; E = { $_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.FailoverLevel } }, 
	@{ N = "AdmissionControlPolicyAutoComputePercentages" ; E = { $_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.AutoComputePercentages } }, 
	@{ N = "AdmissionControlPolicyResourceDarkCyanuctionToToleratePercent" ; E = { $_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.ResourceDarkCyanuctionToToleratePercent } }, 
	@{ N = "DrsEnabled" ; E = { $_.DrsEnabled } }, 
	@{ N = "DrsAutomationLevel" ; E = { $_.DrsAutomationLevel } }, 
	@{ N = "VmMonitoring" ; E = { $_.ExtensionData.configuration.dasconfig.VmMonitoring } }, 
	@{ N = "HostMonitoring" ; E = { $_.ExtensionData.configuration.dasconfig.HostMonitoring } } | Export-Csv $ClusterExportFile -Append -NoTypeInformation
}
#endregion ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VmHost_Export
{
	$VmHostExportFile = "$CaptureCsvFolder\$vCenter-VmHostExport.csv"
	Get-View -ViewType HostSystem -Property Name, Config.Product, Summary.Hardware, Summary, Parent, Config.Network |
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
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
	@{ N = "Version" ; E = { $_.Config.Product.Version } },
	@{ N = "Build" ; E = { $_.Config.Product.Build } },
	@{ N = "Manufacturer" ; E = { $_.Summary.Hardware.Vendor } },
	@{ N = "Model" ; E = { $_.Summary.Hardware.Model } },
	@{ N = "ProcessorType" ; E = { $_.Summary.Hardware.CpuModel } },
	@{ N = "CpuMhz" ; E = { $_.Summary.Hardware.CpuMhz } },
	@{ N = "NumCpuPkgs" ; E = { $_.Summary.Hardware.NumCpuPkgs } },
	@{ N = "NumCpuCores" ; E = { $_.Summary.Hardware.NumCpuCores } },
	@{ N = "NumCpuThreads" ; E = { $_.Summary.Hardware.NumCpuThreads } },
	@{ N = "Memory" ; E = { [math]::Round([decimal]$_.Summary.Hardware.MemorySize / 1073741824) } },
	@{ N = "MaxEVCMode" ; E = { $_.Summary.MaxEVCModeKey } },
	@{ N = "NumNics" ; E = { $_.Summary.Hardware.NumNics } },
	@{ N = "IP" ; E = { [string]::Join(", ", ($_.Config.Network.Vnic.Spec.Ip.IpAddress)) } },
	@{ N = "MacAddress" ; E = { [string]::Join(", ", ($_.Config.Network.Vnic.Spec.Mac)) } },
	@{ N = "NumHBAs" ; E = { $_.Summary.Hardware.NumHBAs } } | Export-Csv $VmHostExportFile -Append -NoTypeInformation
}
#endregion ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Vm_Export
{
	$VmExportFile = "$CaptureCsvFolder\$vCenter-VmExport.csv"
	foreach ($Vm in(Get-View -ViewType VirtualMachine -Property Name, Config, Config.Tools, Guest, Guest.Net, Config.Hardware, Summary.Config, Config.DatastoreUrl, Parent, Runtime.Host, Snapshot, RootSnapshot -Server $vCenter | Sort-Object Name))
	{
		$Folder = Get-View -Id $Vm.Parent -Property Name
		$Vm |
		Select-Object Name ,
		@{ N = "Datacenter" ; E = { Get-Datacenter -VM $_.Name -Server $vCenter } },
		@{ N = "Cluster" ; E = { Get-Cluster -VM $_.Name -Server $vCenter } },
		@{ N = "VmHost" ; E = { Get-VmHost -VM $_.Name -Server $vCenter } },
		@{ N = "DatastoreCluster" ; E = { Get-DatastoreCluster -VM $_.Name } },
		@{ N = "Datastore" ; E = { $_.Config.DatastoreUrl.Name } },
		@{ N = "ResourcePool" ; E = { Get-Vm $_.Name | Get-ResourcePool | Where-Object { $_ -notlike "Resources" } } },
		@{ N = "VsSwitch" ; E = { Get-VirtualSwitch -VM $_.Name -Server $vCenter } },
		@{ N = "PortGroup" ; E = { Get-VirtualPortGroup -VM $_.Name -Server $vCenter } },
		@{ N = "OS" ; E = { $_.Config.GuestFullName } },
		@{ N = "Version" ; E = { $_.Config.Version } },
		@{ N = "VMToolsVersion" ; E = { $_.Guest.ToolsVersion } },
		@{ N = "ToolsVersionStatus" ; E = { $_.Guest.ToolsVersionStatus } },
		@{ N = "ToolsStatus" ; E = { $_.Guest.ToolsStatus } },
		@{ N = "ToolsRunningStatus" ; E = { $_.Guest.ToolsRunningStatus } },
		@{ N = 'Folder' ; E = { $Folder.Name } },
		@{ N = "NumCPU" ; E = { $_.Config.Hardware.NumCPU } },
		@{ N = "CoresPerSocket" ; E = { $_.Config.Hardware.NumCoresPerSocket } },
		@{ N = "MemoryGB" ; E = { [math]::Round([decimal] ( $_.Config.Hardware.MemoryMB / 1024 ), 0) } },
		@{ N = "IP" ; E = { [string]::Join(", ", ($_.Guest.Net.IpAddress)) } },
		@{ N = "MacAddress" ; E = { [string]::Join(", ", ($_.Guest.Net.MacAddress)) } },
		@{ N = "ProvisionedSpaceGB" ; E = { [math]::Round([decimal] ( $_.ProvisionedSpaceGB - $_.MemoryGB ), 0) } },
		@{ N = "NumEthernetCards" ; E = { $_.Summary.Config.NumEthernetCards } },
		@{ N = "NumVirtualDisks" ; E = { $_.Summary.Config.NumVirtualDisks } },
		@{ N = "CpuReservation" ; E = { $_.Summary.Config.CpuReservation } },
		@{ N = "MemoryReservation" ; E = { $_.Summary.Config.MemoryReservation } },
		@{ N = "SRM" ; E = { $_.Summary.Config.ManagedBy.Type } },
		@{ N = "Snapshot" ; E = { $_.Snapshot } },
		@{ N = "RootSnapshot" ; E = { $_.RootSnapshot } } | Export-Csv $VmExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Template_Export
{
	$TemplateExportFile = "$CaptureCsvFolder\$vCenter-TemplateExport.csv"
	foreach ($VmHost in Get-Cluster | Get-VmHost)
	{
		Get-Template -Location $VmHost | 
		Select-Object @{ N = "Name" ; E = { $_.Name } },
		@{ N = "Datacenter" ; E = { $VmHost | Get-Datacenter } },
		@{ N = "Cluster" ; E = { $VmHost | Get-Cluster } },
		@{ N = "VmHost" ; E = { $VmHost.name } },
		@{ N = "Datastore" ; E = { Get-Datastore -Id $_.DatastoreIdList } },
		@{ N = "Folder" ; E = { Get-Folder -Id $_.FolderId } },
		@{ N = "OS" ; E = { $_.ExtensionData.Config.GuestFullName } },
		@{ N = "Version" ; E = { $_.ExtensionData.Config.Version } },
		@{ N = "ToolsVersion" ; E = { $_.ExtensionData.Guest.ToolsVersion } },
		@{ N = "ToolsVersionStatus" ; E = { $_.ExtensionData.Guest.ToolsVersionStatus } },
		@{ N = "ToolsStatus" ; E = { $_.ExtensionData.Guest.ToolsStatus } },
		@{ N = "ToolsRunningStatus" ; E = { $_.ExtensionData.Guest.ToolsRunningStatus } },
		@{ N = "NumCPU" ; E = { $_.ExtensionData.Config.Hardware.NumCPU } },
		@{ N = "NumCoresPerSocket" ; E = { $_.ExtensionData.Config.Hardware.NumCoresPerSocket } },
		@{ N = "MemoryGB" ; E = { [math]::Round([decimal]$_.ExtensionData.Config.Hardware.MemoryMB / 1024, 0) } },
		@{ N = "MacAddress" ; E = { $_.ExtensionData.Config.Hardware.Device.MacAddress } },
		@{ N = "NumEthernetCards" ; E = { $_.ExtensionData.Summary.Config.NumEthernetCards } },
		@{ N = "NumVirtualDisks" ; E = { $_.ExtensionData.Summary.Config.NumVirtualDisks } },
		@{ N = "CpuReservation" ; E = { $_.ExtensionData.Summary.Config.CpuReservation } },
		@{ N = "MemoryReservation" ; E = { $_.ExtensionData.Summary.Config.MemoryReservation } } | Export-Csv $TemplateExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function DatastoreCluster_Export
{
	$DatastoreClusterExportFile = "$CaptureCsvFolder\$vCenter-DatastoreClusterExport.csv"
	Get-DatastoreCluster | Sort-Object Name | 
	Select-Object @{ N = "Name" ; E = { $_.Name } },
	@{ N = "Datacenter" ; E = { Get-DatastoreCluster $_.Name | Get-VmHost | Get-Datacenter } },
	@{ N = "Cluster" ; E = { Get-DatastoreCluster $_.Name | Get-VmHost | Get-Cluster } },
	@{ N = "VmHost" ; E = { Get-DatastoreCluster $_.Name | Get-VmHost } },
	@{ N = "SdrsAutomationLevel" ; E = { $_.SdrsAutomationLevel } },
	@{ N = "IOLoadBalanceEnabled" ; E = { $_.IoLoadBalanceEnabled } },
	@{ N = "CapacityGB" ; E = { [math]::Round([decimal]$_.CapacityGB, 0) } } | Export-Csv $DatastoreClusterExportFile -Append -NoTypeInformation
}
#endregion ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Datastore_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datastore_Export
{
	$DatastoreExportFile = "$CaptureCsvFolder\$vCenter-DatastoreExport.csv"
	Get-Datastore | 
	Select-Object @{ N = "Name" ; E = { $_.Name } },
	@{ N = "Datacenter" ; E = { $_.Datacenter } },
	@{ N = "Cluster" ; E = { Get-Datastore $_.Name | Get-VmHost | Get-Cluster } },
	@{ N = "DatastoreCluster" ; E = { Get-DatastoreCluster -Datastore $_.Name } },
	@{ N = "VmHost" ; E = { Get-VmHost -Datastore $_.Name } },
	@{ N = "Vm" ; E = { Get-Datastore $_.Name | Get-Vm } },
	@{ N = "Type" ; E = { $_.Type } },
	@{ N = "FileSystemVersion" ; E = { $_.FileSystemVersion } },
	@{ N = "DiskName" ; E = { $_.ExtensionData.Info.VMFS.Extent.DiskName } },
	@{ N = "StorageIOControlEnabled" ; E = { $_.StorageIOControlEnabled } },
	@{ N = "CapacityGB" ; E = { [math]::Round([decimal]$_.CapacityGB, 0) } },
	@{ N = "FreeSpaceGB" ; E = { [math]::Round([decimal]$_.FreeSpaceGB, 0) } },
	@{ N = "Accessible" ; E = { $_.State } },
	@{ N = "CongestionThresholdMillisecond" ; E = { $_.CongestionThresholdMillisecond } } | Export-Csv $DatastoreExportFile -Append -NoTypeInformation
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
	@{ N = "ActiveNic" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic } }, 
	@{ N = "StandbyNic" ; E = { $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.StandbyNic } } | Export-Csv $VsSwitchExportFile -Append -NoTypeInformation
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
			@{ N = "ActiveNic" ; E = { $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.ActiveNic } }, 
			@{ N = "StandbyNic" ; E = { $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.StandbyNic } } | Export-Csv $VssPortGroupExportFile -Append -NoTypeInformation
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
	foreach ($VmHost in Get-VmHost)
	{
		Get-VdSwitch -VMHost $VmHost | 
		Select-Object @{ N = "Name" ; E = { $_.Name } }, 
		@{ N = "Datacenter" ; E = { $_.Datacenter } }, 
		@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.name } }, 
		@{ N = "VmHost" ; E = { $VMHost.Name } }, 
		@{ N = "Vendor" ; E = { $_.Vendor } }, 
		@{ N = "Version" ; E = { $_.Version } }, 
		@{ N = "NumUplinkPorts" ; E = { $_.NumUplinkPorts } }, 
		@{ N = "UplinkPortName" ; E = { $_.ExtensionData.Config.UplinkPortPolicy.UplinkPortName } }, 
		@{ N = "Mtu" ; E = { $_.Mtu } } | Export-Csv $VdSwitchExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< VdSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< VdsPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsPort_Export
{
	$VdsPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VdsPortGroupExport.csv"
	foreach ($VmHost in Get-VmHost)
	{
		foreach ($VdSwitch in(Get-VdSwitch -VMHost $VmHost | Sort-Object -Property ConnectedEntity -Unique))
		{
			Get-VDPortGroup | Sort-Object Name | Where-Object { $_.Name -notlike "*DVUplinks*" } | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Datacenter" ; E = { Get-Datacenter -VMHost $VMHost.name } }, 
			@{ N = "Cluster" ; E = { Get-Cluster -VMHost $VMHost.name } }, 
			@{ N = "VmHost" ; E = { $VMHost.Name } }, 
			@{ N = "VlanConfiguration" ; E = { $_.VlanConfiguration } }, 
			@{ N = "VdSwitch" ; E = { $_.VdSwitch } }, 
			@{ N = "NumPorts" ; E = { $_.NumPorts } }, 
			@{ N = "ActiveUplinkPort" ; E = { $_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.ActiveUplinkPort } }, 
			@{ N = "StandbyUplinkPort" ; E = { $_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.StandbyUplinkPort } }, 
			@{ N = "Policy" ; E = { $_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value } }, 
			@{ N = "ReversePolicy" ; E = { $_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.ReversePolicy.Value } }, 
			@{ N = "NotifySwitches" ; E = { $_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.NotifySwitches.Value } }, 
			@{ N = "PortBinding" ; E = { $_.PortBinding } } | Export-Csv $VdsPortGroupExportFile -Append -NoTypeInformation
		}
	}
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
		Get-Folder -Location $Datacenter -type VM | Sort-Object Name | 
		Select-Object @{ N = "Name" ; E = { $_.Name } }, 
		@{ N = "Datacenter" ; E = { $Datacenter.Name } } | Export-Csv $FolderExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Folder_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Rdm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Rdm_Export
{
	$RdmExportFile = "$CaptureCsvFolder\$vCenter-RdmExport.csv"
	Get-VM | Get-HardDisk | Where-Object { $_.DiskType -like "Raw*" } | Sort-Object Parent | 
	Select-Object @{ N = "ScsiCanonicalName" ; E = { $_.ScsiCanonicalName } },
	@{ N = "Cluster" ; E = { Get-Cluster -VM $_.Parent } },
	@{ N = "Vm" ; E = { $_.Parent } },
	@{ N = "Label" ; E = { $_.Name } },
	@{ N = "CapacityGB" ; E = { [math]::Round([decimal]$_.CapacityGB, 2) } },
	@{ N = "DiskType" ; E = { $_.DiskType } },
	@{ N = "Persistence" ; E = { $_.Persistence } },
	@{ N = "CompatibilityMode" ; E = { $_.ExtensionData.Backing.CompatibilityMode } },
	@{ N = "DeviceName" ; E = { $_.ExtensionData.Backing.DeviceName } },
	@{ N = "Sharing" ; E = { $_.ExtensionData.Backing.Sharing } } | Export-Csv $RdmExportFile -Append -NoTypeInformation
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
	foreach ($Cluster in Get-Cluster)
	{
		foreach ($ResourcePool in(Get-Cluster $Cluster | Get-ResourcePool | Where-Object { $_.Name -ne "Resources" } | Sort-Object Name))
		{
			Get-ResourcePool $ResourcePool | Sort-Object Name | 
			Select-Object @{ N = "Name" ; E = { $_.Name } }, 
			@{ N = "Cluster" ; E = { $Cluster.Name } }, 
			@{ N = "CpuSharesLevel" ; E = { $_.CpuSharesLevel } }, 
			@{ N = "NumCpuShares" ; E = { $_.NumCpuShares } }, 
			@{ N = "CpuReservationMHz" ; E = { $_.CpuReservationMHz } }, 
			@{ N = "CpuExpandableReservation" ; E = { $_.CpuExpandableReservation } }, 
			@{ N = "CpuLimitMHz" ; E = { $_.CpuLimitMHz } }, 
			@{ N = "MemSharesLevel" ; E = { $_.MemSharesLevel } }, 
			@{ N = "NumMemShares" ; E = { $_.NumMemShares } }, 
			@{ N = "MemReservationGB" ; E = { $_.MemReservationGB } }, 
			@{ N = "MemExpandableReservation" ; E = { $_.MemExpandableReservation } }, 
			@{ N = "MemLimitGB" ; E = { $_.MemLimitGB } } | Export-Csv $ResourcePoolExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Resource_Pool_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#region ~~< Snapshot_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Snapshot_Export
{
	$SnapshotExportFile = "$CaptureCsvFolder\$vCenter-SnapshotExport.csv"
	(Get-VM | Get-Snapshot) | sort VM, Created | 
	Select-Object @{ N = "VM" ; E = { $_.VM } }, 
	@{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Created" ; E = { $_.Created } }, 
	@{ N = "Children" ; E = { $_.Children } }, 
	@{ N = "ParentSnapshot" ; E = { $_.ParentSnapshot } },
	@{ N = "IsCurrent" ; E = { $_.IsCurrent } } | Export-Csv $SnapshotExportFile -Append -NoTypeInformation
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
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Version" ; E = { $_.Version } }, 
	@{ N = "Build" ; E = { $_.Build } },
	@{ N = "OsType" ; E = { $_.ExtensionData.Content.About.OsType } },
	@{ N = "vCenter" ; E = { $vCenter } } | Export-Csv $LinkedvCenterExportFile -Append -NoTypeInformation
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
	# ProcessorType
	$HostObject.Cells("Prop.ProcessorType").Formula = '"' + $VMHost.ProcessorType + '"'
	# MaxEVCMode
	$HostObject.Cells("Prop.MaxEVCMode").Formula = '"' + $VMHost.MaxEVCMode + '"'
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
	# NumNics
	$HostObject.Cells("Prop.NumNics").Formula = '"' + $VMHost.NumNics + '"'
	# IP
	$HostObject.Cells("Prop.IP").Formula = '"' + $VMHost.IP + '"'
	# MacAddress
	$HostObject.Cells("Prop.Mac").Formula = '"' + $VMHost.MacAddress + '"'
	# NumHBAs
	$HostObject.Cells("Prop.NumHBAs").Formula = '"' + $VMHost.NumHBAs + '"'
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
	$SrmObject.Cells("Prop.OS").Formula = '"' + $SrmVM.OS + '"'
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
	# Vms
	$DatastoreObject.Cells("Prop.Vms").Formula = '"' + $Datastore.Vms + '"'
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
	# ConnectedEntity
	$VssPNICObject.Cells("Prop.ConnectedEntity").Formula = '"' + $VssPnic.ConnectedEntity + '"'
	# VlanConfiguration
	$VssPNICObject.Cells("Prop.VlanConfiguration").Formula = '"' + $VssPnic.VlanConfiguration + '"'
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
#region ~~< Draw_Snapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Snapshot
{
	# Name
	$SnapshotObject.Cells("Prop.Name").Formula = '"' + $Snapshot.Name + '"'
	# Created
	$SnapshotObject.Cells("Prop.Created").Formula = '"' + $Snapshot.Created + '"'
	# Children
	$SnapshotObject.Cells("Prop.Children").Formula = '"' + $Snapshot.Children + '"'
	# ParentSnapshot
	$SnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $Snapshot.ParentSnapshot + '"'
	# IsCurrent
	$SnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $Snapshot.IsCurrent + '"'
}
#endregion ~~< Draw_Snapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	
	foreach ($LinkedvCenter in $LinkedvCenterImport | Sort-Object Name | Where-Object { $_.Name -ne "$vCenter" })
	{
		$x += 2.50
		$LinkedvCenterObject = Add-VisioObjectDC $VCObj $LinkedvCenter
		Draw_vCenter
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
					
			foreach ($ResourcePool in($ResourcePoolImport | Sort-Object Name | Where-Object { $_.Cluster.contains($Cluster.Name) }  | Select -Last 1))
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
										
					foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
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
								
				foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
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
										
					foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
					{
						$x = 10.00
						$y += 1.50
						$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
						Draw_VssPort
						Connect-VisioObject $VssObject $VssNicObject
						$y += 1.50
												
						foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) -and $_.PortGroup.contains($VssPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
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
								
				foreach ($VssPort in($VssPortImport | Sort-Object Name | Where-Object { $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) }))
				{
					$x = 10.00
					$y += 1.50
					$VssNicObject = Add-VisioObjectPG $VssNicObj $VssPort
					Draw_VssPort
					Connect-VisioObject $VssObject $VssNicObject
					$y += 1.50
										
					foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.Cluster -eq "" -and $_.VmHost.contains($VmHost.Name) -and $_.VsSwitch.contains($VsSwitch.Name) -and $_.PortGroup.contains($VssPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
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
								
				foreach ($VM in($VmImport | Sort-Object Name | Where-Object { $_.VsSwitch.contains($VdSwitch.Name) -and $_.PortGroup.contains($VdsPort.Name) -and ($_.SRM.contains("placeholderVm") -eq $False) }))
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
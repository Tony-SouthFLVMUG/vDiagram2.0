<# 
.SYNOPSIS 
   vDiagram Visio Drawing Tool

.DESCRIPTION
   vDiagram Visio Drawing Tool

.NOTES 
   File Name	: vDiagram_2.0.12.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 2.0.12

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data
		MS Visio

.CHANGE LOG
	- 09/12/2021 - v2.0.12
		Added option to choose between vDiagram Visio Shapes and VMware Validated Design Shapes

	- 10/07/2020 - v2.0.11
		Resolved reported issue with standalone ESXi Host.
		Sorted Datastores by name accending.

	- 04/09/2020 - v2.0.10
		Added PowerCLI module version check.
		Added PowerCLI module install if missing.
		Added PowerCLI module upgrade to latest if desired.
		Added device count to capture progression.
		Added device count to draw progression.
		Added additional attributes to shapes.
		Added folder hierarchy to Visio drawing.
		Added DRS Rules hierarchy to Visio drawing.
		Added Resource Pool hierachy to Visio drawing.
		Script now auto hides errors.
		-debug was added to parameters to allow for troubleshooting. To use open Powershell browse to script directory and enter script name -debug ( Example: c:\scripts\vDiagram_2.0.10.ps1 -debug )
		-logcapture was added to parameters to allow for troubleshooting. To use open Powershell browse to script directory and enter script name -debug ( Example: c:\scripts\vDiagram_2.0.10.ps1 -logcapture ). Log capture will be placed in the same directory where script was ran from.
		-logdraw was added to parameters to allow for troubleshooting. To use open Powershell browse to script directory and enter script name -debug ( Example: c:\scripts\vDiagram_2.0.10.ps1 -logdraw ). Log draw will be placed in the same directory where script was ran from.
		All 3 parameters can be used at the same time. To use open Powershell browse to script directory and enter script name -debug ( Example: c:\scripts\vDiagram_2.0.10.ps1 -debug -logcapture -logdraw )
		
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

#region ~~< Parameters >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
param(
	[switch] $debug,
	[switch] $logcapture,
	[switch] $logdraw
)
#endregion ~~< Parameters >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Admin Check >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole] "Administrator")) `
{ `
	Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
	break
}
#endregion ~~< Admin Check >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
#endregion ~~< Constructor >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Pre-PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_PowerCliModule
{
	$PowerCliCheck =  [System.Windows.Forms.MessageBox]::Show( "PowerCLI Module was not found. Would you like to install? Click 'Yes' to install and 'No' cancel.","Warning! Powershell Module is missing.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning )
		switch  ( $PowerCliCheck ) `
		{ `
			'Yes' 
			{ `
				Write-Host "[$DateTime] Installing Module" $VMwareModule.Name -ForegroundColor Green
				set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
				Install-Module -Name VMware.PowerCLI -Scope AllUsers -AllowClobber
				$PowerCliUpdate = Get-Module -Name VMware.PowerCLI -ListAvailable
				Write-Host "[$DateTime] VMware PowerCLI Module" $PowerCliUpdate.Version "is installed." -ForegroundColor Green
			}
			'No'
			{ `
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Unable to proceed without the PowerCLI Module installed. Please run script again and select to install module." -ForegroundColor Red
				exit
			}
		}
}
#endregion ~~< Find_PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Install PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCli = Get-Module -Name VMware.PowerCLI -ListAvailable
$PowerCliLatest = Find-Module -Name VMware.PowerCLI -Repository PSGallery -ErrorAction Stop
if ( ( $PowerCli.Name ) -match ( "VMware.PowerCLI" ) ) `
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] VMware PowerCLI Module(s)" $PowerCli.Version " found on this machine." -ForegroundColor Yellow
	if ( ( $PowerCliLatest.Version ) -gt ( $PowerCli.Version[0] ) ) `
	{ `
		$PowerCliUpgrade =  [System.Windows.Forms.MessageBox]::Show( "PowerCLI Module is not the latest. Would you like to upgrade? Click 'Yes' to upgrade and 'No' cancel.","Warning! PowerCLI Module is not the latest.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Information )
			switch  ( $PowerCliUpgrade ) `
			{ `
				'Yes' 
				{ `
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] You elected to upgrade VMware PowerCLI Module to " $PowerCliLatest.Version -ForegroundColor Yellow
					$Modules = Get-InstalledModule -Name VMware.*
 
					foreach ( $Module in $Modules ) `
					{ `
						$VMwareModules = Get-InstalledModule -Name $Module.Name -AllVersions
						foreach ( $VMwareModule in $VMwareModules )
						{
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Get-Module $VMwareModule.Name -ListAvailable | Uninstall-Module -Force
							Write-Host "[$DateTime] Uninstalling Module" $VMwareModule.Name $VMwareModule.Version -ForegroundColor Yellow
						}
					}
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Installing latest VMware PowerCLI Module" -ForegroundColor Green
					set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
					Install-Module -Name VMware.PowerCLI -Scope AllUsers -AllowClobber
					$PowerCliUpdate = Get-Module -Name VMware.PowerCLI -ListAvailable
					Write-Host "[$DateTime] VMware PowerCLI Module" $PowerCliUpdate.Version "is installed." -ForegroundColor Green
				}
				'No'
				{ `
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] You elected not to upgrade VMware PowerCLI Module. Current version is" $PowerCli.Version[0] -ForegroundColor Yellow
				}
			}
	}
	
} `
else `
{ `
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] VMware PowerCLI Module is not installed." -ForegroundColor Red
	Find_PowerCliModule 
}
#endregion ~~< Install PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Pre-PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Post-Constructor Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< About >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
$MyVer = "2.0.12"
$LastUpdated = "Septemberer 12, 2021"
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
if ( ( Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) -or $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) ) `
{ `
	$TestShapes1 = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes\vDiagram_" + $MyVer + ".vssx"
	if (!(Test-Path $TestShapes1))
	{
		$CurrentLocation = Get-Location
		$UpdatedShapes = "$CurrentLocation" + "\vDiagram_" + "$MyVer" + ".vssx"
		Copy-Item $UpdatedShapes $TestShapes1
		Write-Host "Copying Default Shapes File to My Shapes"
	}
	$TestShapes2 = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes\vDiagram_" + $MyVer + "_VVD" + ".vssx"
	if (!(Test-Path $TestShapes2))
	{
		$CurrentLocation = Get-Location
		$UpdatedShapes = "$CurrentLocation" + "\vDiagram_" + "$MyVer" + "_VVD" + ".vssx"
		Copy-Item $UpdatedShapes $TestShapes2
		Write-Host "Copying VMware VVD Shapes File to My Shapes"
	}
}
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

    $Win32ShowWindowAsync = Add-Type -memberDefinition @"
    [DllImport("user32.dll")] 
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
"@ -name "Win32ShowWindowAsync" -namespace Win32Functions -passThru

    $Win32ShowWindowAsync::ShowWindowAsync($MainWindowHandle, $WindowStates[$Style]) | Out-Null
}
#Set_WindowStyle MINIMIZE
if($debug -eq $true)
{
	$ErrorActionPreference = "Continue"
	$WarningPreference = "Continue"
	Set_WindowStyle SHOWDEFAULT
}
if($debug -eq $false)
{
	$ErrorActionPreference = "SilentlyContinue"
	$WarningPreference = "SilentlyContinue"
	Set_WindowStyle FORCEMINIMIZE
}
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

#region ~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$components = New-Object System.ComponentModel.Container
#endregion ~~< components >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$MainMenu = New-Object System.Windows.Forms.MenuStrip
$MainMenu.Location = New-Object System.Drawing.Point(0, 0)
$MainMenu.Size = New-Object System.Drawing.Size(1010, 24)
$MainMenu.TabIndex = 0
$MainMenu.Text = "MainMenu"

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
$FileToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($ExitToolStripMenuItem)))
#endregion ~~< ExitToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
$HelpToolStripMenuItem.DropDownItems.AddRange([System.Windows.Forms.ToolStripItem[]](@($AboutToolStripMenuItem)))
#endregion ~~< AboutToolStripMenuItem >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Help >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$MainMenu.Items.AddRange([System.Windows.Forms.ToolStripItem[]](@($FileToolStripMenuItem, $HelpToolStripMenuItem)))
$vDiagram.Controls.Add($MainMenu)

#endregion ~~< MainMenu >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UpperTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UpperTabs = New-Object System.Windows.Forms.TabControl
$UpperTabs.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$UpperTabs.ItemSize = New-Object System.Drawing.Size(85, 20)
$UpperTabs.Location = New-Object System.Drawing.Point(10, 30)
$UpperTabs.Size = New-Object System.Drawing.Size(990, 98)
$UpperTabs.TabIndex = 1
$UpperTabs.Text = "UpperTabs"

#region ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Prerequisites = New-Object System.Windows.Forms.TabPage
$Prerequisites.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$Prerequisites.Location = New-Object System.Drawing.Point(4, 24)
$Prerequisites.Padding = New-Object System.Windows.Forms.Padding(3)
$Prerequisites.Size = New-Object System.Drawing.Size(982, 70)
$Prerequisites.TabIndex = 0
$Prerequisites.Text = "Prerequisites"
$Prerequisites.ToolTipText = "Prerequisites: These items are needed in order to run this script."
$Prerequisites.BackColor = [System.Drawing.Color]::LightGray

#region ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellLabel = New-Object System.Windows.Forms.Label
$PowershellLabel.Location = New-Object System.Drawing.Point(10, 15)
$PowershellLabel.Size = New-Object System.Drawing.Size(75, 20)
$PowershellLabel.TabIndex = 0
$PowershellLabel.Text = "Powershell:"
$Prerequisites.Controls.Add($PowershellLabel)
#endregion ~~< PowershellLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellInstalled = New-Object System.Windows.Forms.Label
$PowershellInstalled.Location = New-Object System.Drawing.Point(96, 15)
$PowershellInstalled.Size = New-Object System.Drawing.Size(350, 20)
$PowershellInstalled.TabIndex = 1
$PowershellInstalled.Text = ""
$Prerequisites.Controls.Add($PowershellInstalled)
#endregion ~~< PowershellInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Powershell >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliLabel = New-Object System.Windows.Forms.Label
$PowerCliLabel.Location = New-Object System.Drawing.Point(450, 15)
$PowerCliLabel.Size = New-Object System.Drawing.Size(64, 20)
$PowerCliLabel.TabIndex = 4
$PowerCliLabel.Text = "PowerCLI:"
$Prerequisites.Controls.Add($PowerCliLabel)
#endregion ~~< PowerCliLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliInstalled = New-Object System.Windows.Forms.Label
$PowerCliInstalled.Location = New-Object System.Drawing.Point(520, 15)
$PowerCliInstalled.Size = New-Object System.Drawing.Size(400, 20)
$PowerCliInstalled.TabIndex = 5
$PowerCliInstalled.Text = ""
$Prerequisites.Controls.Add($PowerCliInstalled)
#endregion ~~< PowerCliInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PowerCli >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleLabel = New-Object System.Windows.Forms.Label
$PowerCliModuleLabel.Location = New-Object System.Drawing.Point(10, 40)
$PowerCliModuleLabel.Size = New-Object System.Drawing.Size(110, 20)
$PowerCliModuleLabel.TabIndex = 2
$PowerCliModuleLabel.Text = "PowerCLI Module:"
$Prerequisites.Controls.Add($PowerCliModuleLabel)
#endregion ~~< PowerCliModuleLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleInstalled = New-Object System.Windows.Forms.Label
$PowerCliModuleInstalled.Location = New-Object System.Drawing.Point(128, 40)
$PowerCliModuleInstalled.Size = New-Object System.Drawing.Size(320, 20)
$PowerCliModuleInstalled.TabIndex = 3
$PowerCliModuleInstalled.Text = ""
$Prerequisites.Controls.Add($PowerCliModuleInstalled)
#endregion ~~< PowerCliModuleInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PowerCliModule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioLabel = New-Object System.Windows.Forms.Label
$VisioLabel.Location = New-Object System.Drawing.Point(450, 40)
$VisioLabel.Size = New-Object System.Drawing.Size(40, 20)
$VisioLabel.TabIndex = 6
$VisioLabel.Text = "Visio:"
$Prerequisites.Controls.Add($VisioLabel)
#endregion ~~< VisioLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioInstalled = New-Object System.Windows.Forms.Label
$VisioInstalled.Location = New-Object System.Drawing.Point(490, 40)
$VisioInstalled.Size = New-Object System.Drawing.Size(320, 20)
$VisioInstalled.TabIndex = 7
$VisioInstalled.Text = ""
$Prerequisites.Controls.Add($VisioInstalled)
#endregion ~~< VisioInstalled >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.Controls.Add($Prerequisites)
#endregion ~~< Prerequisites >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfo = New-Object System.Windows.Forms.TabPage
$vCenterInfo.Location = New-Object System.Drawing.Point(4, 24)
$vCenterInfo.Padding = New-Object System.Windows.Forms.Padding(3)
$vCenterInfo.Size = New-Object System.Drawing.Size(982, 70)
$vCenterInfo.TabIndex = 1
$vCenterInfo.Text = "vCenter Info"
$vCenterInfo.BackColor = [System.Drawing.Color]::LightGray

#region ~~< VcenterLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterLabel = New-Object System.Windows.Forms.Label
$VcenterLabel.Location = New-Object System.Drawing.Point(8, 11)
$VcenterLabel.Size = New-Object System.Drawing.Size(70, 20)
$VcenterLabel.TabIndex = 0
$VcenterLabel.Text = "vCenter:"
$vCenterInfo.Controls.Add($VcenterLabel)
#endregion ~~< VcenterLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VcenterTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterTextBox = New-Object System.Windows.Forms.TextBox
$VcenterTextBox.Location = New-Object System.Drawing.Point(78, 8)
$VcenterTextBox.Size = New-Object System.Drawing.Size(238, 21)
$VcenterTextBox.TabIndex = 1
$VcenterTextBox.Text = ""
$vCenterInfo.Controls.Add($VcenterTextBox)
#endregion ~~< VcenterTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VcenterToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VcenterToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VcenterToolTip.AutoPopDelay = 5000
$VcenterToolTip.InitialDelay = 50
$VcenterToolTip.IsBalloon = $true
$VcenterToolTip.ReshowDelay = 100
$VcenterToolTip.SetToolTip($VcenterTextBox, "Enter vCenter name")
#endregion ~~< VcenterToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameLabel = New-Object System.Windows.Forms.Label
$UserNameLabel.Location = New-Object System.Drawing.Point(324, 11)
$UserNameLabel.Size = New-Object System.Drawing.Size(70, 20)
$UserNameLabel.TabIndex = 2
$UserNameLabel.Text = "User Name:"
$vCenterInfo.Controls.Add($UserNameLabel)
#endregion ~~< UserNameLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameTextBox = New-Object System.Windows.Forms.TextBox
$UserNameTextBox.Location = New-Object System.Drawing.Point(402, 8)
$UserNameTextBox.Size = New-Object System.Drawing.Size(238, 21)
$UserNameTextBox.TabIndex = 3
$UserNameTextBox.Text = ""
$vCenterInfo.Controls.Add($UserNameTextBox)
#endregion ~~< UserNameTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< UserNameToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$UserNameToolTip = New-Object System.Windows.Forms.ToolTip($components)
$UserNameToolTip.AutoPopDelay = 5000
$UserNameToolTip.InitialDelay = 50
$UserNameToolTip.IsBalloon = $true
$UserNameToolTip.ReshowDelay = 100
$UserNameToolTip.SetToolTip($UserNameTextBox, "Enter User Name."+[char]13+[char]10+[char]13+[char]10+"Example:"+[char]13+[char]10+"administrator@vsphere.local"+[char]13+[char]10+"Domain\User")
#endregion ~~< UserNameToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordLabel = New-Object System.Windows.Forms.Label
$PasswordLabel.Location = New-Object System.Drawing.Point(656, 11)
$PasswordLabel.Size = New-Object System.Drawing.Size(70, 20)
$PasswordLabel.TabIndex = 4
$PasswordLabel.Text = "Password:"
$vCenterInfo.Controls.Add($PasswordLabel)
#endregion ~~< PasswordLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordTextBox = New-Object System.Windows.Forms.TextBox
$PasswordTextBox.Location = New-Object System.Drawing.Point(734, 8)
$PasswordTextBox.Size = New-Object System.Drawing.Size(238, 21)
$PasswordTextBox.TabIndex = 5
$PasswordTextBox.Text = ""
$PasswordTextBox.UseSystemPasswordChar = $true
$vCenterInfo.Controls.Add($PasswordTextBox)
#endregion ~~< PasswordTextBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PasswordToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PasswordToolTip = New-Object System.Windows.Forms.ToolTip($components)
$PasswordToolTip.AutoPopDelay = 5000
$PasswordToolTip.InitialDelay = 50
$PasswordToolTip.IsBalloon = $true
$PasswordToolTip.ReshowDelay = 100
$PasswordToolTip.SetToolTip($PasswordTextBox, "Enter Passwrd."+[char]13+[char]10+[char]13+[char]10+"Characters will not be seen.")
#endregion ~~< PasswordToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton = New-Object System.Windows.Forms.Button
$ConnectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$ConnectButton.Location = New-Object System.Drawing.Point(8, 37)
$ConnectButton.Size = New-Object System.Drawing.Size(345, 25)
$ConnectButton.TabIndex = 6
$ConnectButton.Text = "Connect to vCenter"
$ConnectButton.UseVisualStyleBackColor = $true
$vCenterInfo.Controls.Add($ConnectButton)
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ConnectButtonToolTip.AutoPopDelay = 5000
$ConnectButtonToolTip.InitialDelay = 50
$ConnectButtonToolTip.IsBalloon = $true
$ConnectButtonToolTip.ReshowDelay = 100
$ConnectButtonToolTip.SetToolTip($ConnectButton, "Click to connect to vCenter."+[char]13+[char]10+[char]13+[char]10+"If connected this button will turn green and show connected to the name entered in the vCenter box."+[char]13+[char]10+[char]13+[char]10+"If disconnected or unable to connect this button will display red text, indicating that you were unable to"+[char]13+[char]10+"connect to vCenter either due to bad creditials, not being on the same network or insufficient access to vCenter.")
#endregion ~~< ConnectButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.Controls.Add($vCenterInfo)

#endregion ~~< vCenterInfo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$UpperTabs.SelectedIndex = 0
$vDiagram.Controls.Add($UpperTabs)
#endregion ~~< UpperTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LowerTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LowerTabs = New-Object System.Windows.Forms.TabControl
$LowerTabs.Font = New-Object System.Drawing.Font("Tahoma", 8.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$LowerTabs.Location = New-Object System.Drawing.Point(10, 136)
$LowerTabs.Size = New-Object System.Drawing.Size(990, 512)
$LowerTabs.TabIndex = 2
$LowerTabs.Text = "LowerTabs"

#region ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDirections = New-Object System.Windows.Forms.TabPage
$TabDirections.Location = New-Object System.Drawing.Point(4, 22)
$TabDirections.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDirections.Size = New-Object System.Drawing.Size(982, 486)
$TabDirections.TabIndex = 0
$TabDirections.Text = "Directions"
$TabDirections.UseVisualStyleBackColor = $true

#region ~~< Prerequisites Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PrerequisitesHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PrerequisitesHeading = New-Object System.Windows.Forms.Label
$PrerequisitesHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
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
$PrerequisitesDirections.Text = "1. Verify that prerequisites are met on the "+[char]34+"Prerequisites"+[char]34+" tab."+[char]34+[char]13+[char]10+"2. If not please install needed requirements."+[char]13+[char]10
$TabDirections.Controls.Add($PrerequisitesDirections)
#endregion ~~< PrerequisitesDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Prerequisites Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterInfo Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterInfoHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterInfoHeading = New-Object System.Windows.Forms.Label
$vCenterInfoHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
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
$vCenterInfoDirections.Text = "1. Click on"+[char]34+"vCenter Info"+[char]34+" tab."+[char]13+[char]10+"2. Enter name of vCenter"+[char]13+[char]10+"3. Enter User Name and Password (password will be hashed and not plain text)."+[char]13+[char]10+"4. Click on "+[char]34+"Connect to vCenter"+[char]34+" button."
$TabDirections.Controls.Add($vCenterInfoDirections)
#endregion ~~< vCenterInfoDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vCenterInfo Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvHeading = New-Object System.Windows.Forms.Label
$CaptureCsvHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
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

#endregion ~~< Capture Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawHeading >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawHeading = New-Object System.Windows.Forms.Label
$DrawHeading.Font = New-Object System.Drawing.Font("Tahoma", 11.25, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
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
$DrawDirections.Text = "1. Click on "+[char]34+"Select Input Folder"+[char]34+" button and select location where CSVs can be found."+[char]13+[char]10+"2. Click on "+[char]34+"Check for CSVs"+[char]34+" button to validate presence of required files."+[char]13+[char]10+"4. Select shape file that you would like to use."+[char]13+[char]10+"4. Click on "+[char]34+"Select Output Folder"+[char]34+" button and select where location where you would like to save the Visio drawing."+[char]13+[char]10+"5. Select drawing that you would like to produce."+[char]13+[char]10+"6. Click on "+[char]34+"Draw Visio"+[char]34+" button."+[char]13+[char]10+"7. Click on "+[char]34+"Open Visio Drawing"+[char]34+" button once "+[char]34+"Draw Visio"+[char]34+" button says it has completed."
$TabDirections.Controls.Add($DrawDirections)
#endregion ~~< DrawDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Draw Lower >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add($TabDirections)
#endregion ~~< TabDirections >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCapture = New-Object System.Windows.Forms.TabPage
$TabCapture.Location = New-Object System.Drawing.Point(4, 22)
$TabCapture.Padding = New-Object System.Windows.Forms.Padding(3)
$TabCapture.Size = New-Object System.Drawing.Size(982, 486)
$TabCapture.TabIndex = 1
$TabCapture.Text = "Capture CSVs for Visio"
$TabCapture.UseVisualStyleBackColor = $true

#region ~~< TabCaptureToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabCaptureToolTip = New-Object System.Windows.Forms.ToolTip($components)
$TabCaptureToolTip.AutoPopDelay = 5000
$TabCaptureToolTip.InitialDelay = 50
$TabCaptureToolTip.IsBalloon = $true
$TabCaptureToolTip.ReshowDelay = 100
$TabCaptureToolTip.SetToolTip($TabCapture, "This must be ran first in order to collect the information"+[char]13+[char]10+"about your environment.")
#endregion ~~< TabCaptureToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Folder Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton = New-Object System.Windows.Forms.Button
$CaptureCsvOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCsvOutputButton.Location = New-Object System.Drawing.Point(220, 10)
$CaptureCsvOutputButton.Size = New-Object System.Drawing.Size(750, 25)
$CaptureCsvOutputButton.TabIndex = 1
$CaptureCsvOutputButton.Text = "Select Output Folder"
$CaptureCsvOutputButton.UseVisualStyleBackColor = $false
$CaptureCsvOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureCsvOutputButton)
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureCsvOutputButtonToolTip.AutoPopDelay = 5000
$CaptureCsvOutputButtonToolTip.InitialDelay = 50
$CaptureCsvOutputButtonToolTip.IsBalloon = $true
$CaptureCsvOutputButtonToolTip.ReshowDelay = 100
$CaptureCsvOutputButtonToolTip.SetToolTip($CaptureCsvOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"CSV"+[char]39+"s."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green."+[char]13+[char]10+[char]13+[char]10+"If the folder has files in it you will be presented with an "+[char]13+[char]10+"option to move or delete the files that are currently there.")
#endregion ~~< CaptureCsvOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputLabel = New-Object System.Windows.Forms.Label
$CaptureCsvOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CaptureCsvOutputLabel.Location = New-Object System.Drawing.Point(10, 10)
$CaptureCsvOutputLabel.Size = New-Object System.Drawing.Size(210, 25)
$CaptureCsvOutputLabel.TabIndex = 0
$CaptureCsvOutputLabel.Text = "CSV Output Folder:"
$TabCapture.Controls.Add($CaptureCsvOutputLabel)
#endregion ~~< CaptureCsvOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$CaptureCsvBrowse.Description = "Select a directory"
$CaptureCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< CaptureCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Capture Folder Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenterCsvCheckBox.Checked = $true
$vCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 40)
$vCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$vCenterCsvCheckBox.TabIndex = 2
$vCenterCsvCheckBox.Text = "Export vCenter Info"
$vCenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($vCenterCsvCheckBox)
#endregion ~~< vCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$vCenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 40)
$vCenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$vCenterCsvValidationComplete.TabIndex = 3
$vCenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($vCenterCsvValidationComplete)
#endregion ~~< vCenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$vCenterCsvToolTip.AutoPopDelay = 5000
$vCenterCsvToolTip.InitialDelay = 50
$vCenterCsvToolTip.IsBalloon = $true
$vCenterCsvToolTip.ReshowDelay = 100
$vCenterCsvToolTip.SetToolTip($vCenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"vCenter.")
#endregion ~~< vCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatacenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatacenterCsvCheckBox.Checked = $true
$DatacenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatacenterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 60)
$DatacenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatacenterCsvCheckBox.TabIndex = 4
$DatacenterCsvCheckBox.Text = "Export Datacenter Info"
$DatacenterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatacenterCsvCheckBox)
#endregion ~~< DatacenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatacenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatacenterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 60)
$DatacenterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatacenterCsvValidationComplete.TabIndex = 5
$DatacenterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatacenterCsvValidationComplete)
#endregion ~~< DatacenterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatacenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatacenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatacenterCsvToolTip.AutoPopDelay = 5000
$DatacenterCsvToolTip.InitialDelay = 50
$DatacenterCsvToolTip.IsBalloon = $true
$DatacenterCsvToolTip.ReshowDelay = 100
$DatacenterCsvToolTip.SetToolTip($DatacenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datacenters in this vCenter.")
#endregion ~~< DatacenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ClusterCsvCheckBox.Checked = $true
$ClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 80)
$ClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ClusterCsvCheckBox.TabIndex = 6
$ClusterCsvCheckBox.Text = "Export Cluster Info"
$ClusterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($ClusterCsvCheckBox)
#endregion ~~< ClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$ClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 80)
$ClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ClusterCsvValidationComplete.TabIndex = 7
$ClusterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($ClusterCsvValidationComplete)
#endregion ~~< ClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ClusterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ClusterCsvToolTip.AutoPopDelay = 5000
$ClusterCsvToolTip.InitialDelay = 50
$ClusterCsvToolTip.IsBalloon = $true
$ClusterCsvToolTip.ReshowDelay = 100
$ClusterCsvToolTip.SetToolTip($ClusterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Clusters in this vCenter.")
#endregion ~~< ClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHostCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmHostCsvCheckBox.Checked = $true
$VmHostCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmHostCsvCheckBox.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmHostCsvCheckBox.TabIndex = 8
$VmHostCsvCheckBox.Text = "Export Host Info"
$VmHostCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VmHostCsvCheckBox)
#endregion ~~< VmHostCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHostCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmHostCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 100)
$VmHostCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationComplete.TabIndex = 9
$VmHostCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VmHostCsvValidationComplete)
#endregion ~~< VmHostCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHostCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VmHostCsvToolTip.AutoPopDelay = 5000
$VmHostCsvToolTip.InitialDelay = 50
$VmHostCsvToolTip.IsBalloon = $true
$VmHostCsvToolTip.ReshowDelay = 100
$VmHostCsvToolTip.SetToolTip($VmHostCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all ESXi Hosts in this vCenter.")
#endregion ~~< VmHostCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Vm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VmCsvCheckBox.Checked = $true
$VmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VmCsvCheckBox.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VmCsvCheckBox.TabIndex = 10
$VmCsvCheckBox.Text = "Export Virtual Machine Info"
$VmCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VmCsvCheckBox)
#endregion ~~< VmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidationComplete = New-Object System.Windows.Forms.Label
$VmCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 120)
$VmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VmCsvValidationComplete.TabIndex = 11
$VmCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VmCsvValidationComplete)
#endregion ~~< VmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VmCsvToolTip.AutoPopDelay = 5000
$VmCsvToolTip.InitialDelay = 50
$VmCsvToolTip.IsBalloon = $true
$VmCsvToolTip.ReshowDelay = 100
$VmCsvToolTip.SetToolTip($VmCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Machines in this vCenter.")
#endregion ~~< VmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Vm >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TemplateCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$TemplateCsvCheckBox.Checked = $true
$TemplateCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$TemplateCsvCheckBox.Location = New-Object System.Drawing.Point(10, 140)
$TemplateCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$TemplateCsvCheckBox.TabIndex = 12
$TemplateCsvCheckBox.Text = "Export Template Info"
$TemplateCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($TemplateCsvCheckBox)
#endregion ~~< TemplateCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TemplateCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvValidationComplete = New-Object System.Windows.Forms.Label
$TemplateCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 140)
$TemplateCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$TemplateCsvValidationComplete.TabIndex = 13
$TemplateCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($TemplateCsvValidationComplete)
#endregion ~~< TemplateCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TemplateCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TemplateCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$TemplateCsvToolTip.AutoPopDelay = 5000
$TemplateCsvToolTip.InitialDelay = 50
$TemplateCsvToolTip.IsBalloon = $true
$TemplateCsvToolTip.ReshowDelay = 100
$TemplateCsvToolTip.SetToolTip($TemplateCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Templates in this vCenter.")
#endregion ~~< TemplateCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreClusterCsvCheckBox.Checked = $true
$DatastoreClusterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreClusterCsvCheckBox.Location = New-Object System.Drawing.Point(10, 160)
$DatastoreClusterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreClusterCsvCheckBox.TabIndex = 14
$DatastoreClusterCsvCheckBox.Text = "Export Datastore Cluster Info"
$DatastoreClusterCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatastoreClusterCsvCheckBox)
#endregion ~~< DatastoreClusterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreClusterCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 160)
$DatastoreClusterCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreClusterCsvValidationComplete.TabIndex = 15
$DatastoreClusterCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatastoreClusterCsvValidationComplete)
#endregion ~~< DatastoreClusterCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreClusterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatastoreClusterCsvToolTip.AutoPopDelay = 5000
$DatastoreClusterCsvToolTip.InitialDelay = 50
$DatastoreClusterCsvToolTip.IsBalloon = $true
$DatastoreClusterCsvToolTip.ReshowDelay = 100
$DatastoreClusterCsvToolTip.SetToolTip($DatastoreClusterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datastore Clusters in this vCenter.")
#endregion ~~< DatastoreClusterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Datastore Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DatastoreCsvCheckBox.Checked = $true
$DatastoreCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DatastoreCsvCheckBox.Location = New-Object System.Drawing.Point(10, 180)
$DatastoreCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DatastoreCsvCheckBox.TabIndex = 16
$DatastoreCsvCheckBox.Text = "Export Datastore Info"
$DatastoreCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DatastoreCsvCheckBox)
#endregion ~~< DatastoreCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvValidationComplete = New-Object System.Windows.Forms.Label
$DatastoreCsvValidationComplete.Location = New-Object System.Drawing.Point(210, 180)
$DatastoreCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DatastoreCsvValidationComplete.TabIndex = 17
$DatastoreCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DatastoreCsvValidationComplete)
#endregion ~~< DatastoreCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DatastoreCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DatastoreCsvToolTip.AutoPopDelay = 5000
$DatastoreCsvToolTip.InitialDelay = 50
$DatastoreCsvToolTip.IsBalloon = $true
$DatastoreCsvToolTip.ReshowDelay = 100
$DatastoreCsvToolTip.SetToolTip($DatastoreCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Datastores in this vCenter.")
#endregion ~~< DatacenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Standard Switch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VsSwitchCsvCheckBox.Checked = $true
$VsSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VsSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 40)
$VsSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VsSwitchCsvCheckBox.TabIndex = 18
$VsSwitchCsvCheckBox.Text = "Export Standard Switch Info"
$VsSwitchCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VsSwitchCsvCheckBox)
#endregion ~~< VsSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 40)
$VsSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VsSwitchCsvValidationComplete.TabIndex = 19
$VsSwitchCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VsSwitchCsvValidationComplete)
#endregion ~~< VsSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VsSwitchCsvToolTip.AutoPopDelay = 5000
$VsSwitchCsvToolTip.InitialDelay = 50
$VsSwitchCsvToolTip.IsBalloon = $true
$VsSwitchCsvToolTip.ReshowDelay = 100
$VsSwitchCsvToolTip.SetToolTip($VsSwitchCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switches in this vCenter.")
#endregion ~~< VsSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Standard Switch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Standard Port Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPortGroupCsvCheckBox.Checked = $true
$VssPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 60)
$VssPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPortGroupCsvCheckBox.TabIndex = 20
$VssPortGroupCsvCheckBox.Text = "Export VSS Port Group Info"
$VssPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssPortGroupCsvCheckBox)
#endregion ~~< VssPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 60)
$VssPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPortGroupCsvValidationComplete.TabIndex = 21
$VssPortGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssPortGroupCsvValidationComplete)
#endregion ~~< VssPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPortGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssPortGroupCsvToolTip.AutoPopDelay = 5000
$VssPortGroupCsvToolTip.InitialDelay = 50
$VssPortGroupCsvToolTip.IsBalloon = $true
$VssPortGroupCsvToolTip.ReshowDelay = 100
$VssPortGroupCsvToolTip.SetToolTip($VssPortGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch Port Groups in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VssPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Standard Port Grouph >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Standard VMKernel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssVmkernelCsvCheckBox.Checked = $true
$VssVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 80)
$VssVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssVmkernelCsvCheckBox.TabIndex = 22
$VssVmkernelCsvCheckBox.Text = "Export VSS VMkernel Info"
$VssVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssVmkernelCsvCheckBox)
#endregion ~~< VssVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 80)
$VssVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssVmkernelCsvValidationComplete.TabIndex = 23
$VssVmkernelCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssVmkernelCsvValidationComplete)
#endregion ~~< VssVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssVmkernelCsvToolTip.AutoPopDelay = 5000
$VssVmkernelCsvToolTip.InitialDelay = 50
$VssVmkernelCsvToolTip.IsBalloon = $true
$VssVmkernelCsvToolTip.ReshowDelay = 100
$VssVmkernelCsvToolTip.SetToolTip($VssVmkernelCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch VMkernels in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VssVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Standard VMKernel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Standard PNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VssPnicCsvCheckBox.Checked = $true
$VssPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VssPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 100)
$VssPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VssPnicCsvCheckBox.TabIndex = 24
$VssPnicCsvCheckBox.Text = "Export VSS pNIC Info"
$VssPnicCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VssPnicCsvCheckBox)
#endregion ~~< VssPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VssPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 100)
$VssPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VssPnicCsvValidationComplete.TabIndex = 25
$VssPnicCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VssPnicCsvValidationComplete)
#endregion ~~< VssPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VssPnicCsvToolTip.AutoPopDelay = 5000
$VssPnicCsvToolTip.InitialDelay = 50
$VssPnicCsvToolTip.IsBalloon = $true
$VssPnicCsvToolTip.ReshowDelay = 100
$VssPnicCsvToolTip.SetToolTip($VssPnicCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Standard Switch Physical NICs in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VssPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Standard PNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Distributed Switch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdSwitchCsvCheckBox.Checked = $true
$VdSwitchCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdSwitchCsvCheckBox.Location = New-Object System.Drawing.Point(310, 120)
$VdSwitchCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdSwitchCsvCheckBox.TabIndex = 26
$VdSwitchCsvCheckBox.Text = "Export Distributed Switch Info"
$VdSwitchCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdSwitchCsvCheckBox)
#endregion ~~< VdSwitchCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 120)
$VdSwitchCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdSwitchCsvValidationComplete.TabIndex = 27
$VdSwitchCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdSwitchCsvValidationComplete)
#endregion ~~< VdSwitchCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdSwitchCsvToolTip.AutoPopDelay = 5000
$VdSwitchCsvToolTip.InitialDelay = 50
$VdSwitchCsvToolTip.IsBalloon = $true
$VdSwitchCsvToolTip.ReshowDelay = 100
$VdSwitchCsvToolTip.SetToolTip($VdSwitchCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switches in this vCenter.")
#endregion ~~< VdSwitchCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Distributed Switch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Distributed Port Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPortGroupCsvCheckBox.Checked = $true
$VdsPortGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPortGroupCsvCheckBox.Location = New-Object System.Drawing.Point(310, 140)
$VdsPortGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPortGroupCsvCheckBox.TabIndex = 28
$VdsPortGroupCsvCheckBox.Text = "Export VDS Port Group Info"
$VdsPortGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsPortGroupCsvCheckBox)
#endregion ~~< VdsPortGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPortGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 140)
$VdsPortGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPortGroupCsvValidationComplete.TabIndex = 29
$VdsPortGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsPortGroupCsvValidationComplete)
#endregion ~~< VdsPortGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPortGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsPortGroupCsvToolTip.AutoPopDelay = 5000
$VdsPortGroupCsvToolTip.InitialDelay = 50
$VdsPortGroupCsvToolTip.IsBalloon = $true
$VdsPortGroupCsvToolTip.ReshowDelay = 100
$VdsPortGroupCsvToolTip.SetToolTip($VdsPortGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch Port Groups in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VdsPortGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Distributed Port Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Distributed VMKernel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsVmkernelCsvCheckBox.Checked = $true
$VdsVmkernelCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsVmkernelCsvCheckBox.Location = New-Object System.Drawing.Point(310, 160)
$VdsVmkernelCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsVmkernelCsvCheckBox.TabIndex = 30
$VdsVmkernelCsvCheckBox.Text = "Export VDS VMkernel Info"
$VdsVmkernelCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsVmkernelCsvCheckBox)
#endregion ~~< VdsVmkernelCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 160)
$VdsVmkernelCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsVmkernelCsvValidationComplete.TabIndex = 31
$VdsVmkernelCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsVmkernelCsvValidationComplete)
#endregion ~~< VdsVmkernelCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsVmkernelCsvToolTip.AutoPopDelay = 5000
$VdsVmkernelCsvToolTip.InitialDelay = 50
$VdsVmkernelCsvToolTip.IsBalloon = $true
$VdsVmkernelCsvToolTip.ReshowDelay = 100
$VdsVmkernelCsvToolTip.SetToolTip($VdsVmkernelCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch VMkernels in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VdsVmkernelCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Distributed VMKernel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Virtual Distributed PNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$VdsPnicCsvCheckBox.Checked = $true
$VdsPnicCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VdsPnicCsvCheckBox.Location = New-Object System.Drawing.Point(310, 180)
$VdsPnicCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$VdsPnicCsvCheckBox.TabIndex = 32
$VdsPnicCsvCheckBox.Text = "Export VDS pNIC Info"
$VdsPnicCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($VdsPnicCsvCheckBox)
#endregion ~~< VdsPnicCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidationComplete = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidationComplete.Location = New-Object System.Drawing.Point(520, 180)
$VdsPnicCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$VdsPnicCsvValidationComplete.TabIndex = 33
$VdsPnicCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($VdsPnicCsvValidationComplete)
#endregion ~~< VdsPnicCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VdsPnicCsvToolTip.AutoPopDelay = 5000
$VdsPnicCsvToolTip.InitialDelay = 50
$VdsPnicCsvToolTip.IsBalloon = $true
$VdsPnicCsvToolTip.ReshowDelay = 100
$VdsPnicCsvToolTip.SetToolTip($VdsPnicCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Virtual Distributed Switch Physical NICs in"+[char]13+[char]10+"this vCenter.")
#endregion ~~< VdsPnicCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Virtual Distributed PNIC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FolderCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$FolderCsvCheckBox.Checked = $true
$FolderCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$FolderCsvCheckBox.Location = New-Object System.Drawing.Point(620, 40)
$FolderCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$FolderCsvCheckBox.TabIndex = 34
$FolderCsvCheckBox.Text = "Export Folder Info"
$FolderCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($FolderCsvCheckBox)
#endregion ~~< FolderCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FolderCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvValidationComplete = New-Object System.Windows.Forms.Label
$FolderCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 40)
$FolderCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$FolderCsvValidationComplete.TabIndex = 35
$FolderCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($FolderCsvValidationComplete)
#endregion ~~< FolderCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FolderCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$FolderCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$FolderCsvToolTip.AutoPopDelay = 5000
$FolderCsvToolTip.InitialDelay = 50
$FolderCsvToolTip.IsBalloon = $true
$FolderCsvToolTip.ReshowDelay = 100
$FolderCsvToolTip.SetToolTip($FolderCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Folders in this vCenter.")
#endregion ~~< FolderCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RdmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$RdmCsvCheckBox.Checked = $true
$RdmCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$RdmCsvCheckBox.Location = New-Object System.Drawing.Point(620, 60)
$RdmCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$RdmCsvCheckBox.TabIndex = 36
$RdmCsvCheckBox.Text = "Export RDM Info"
$RdmCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($RdmCsvCheckBox)
#endregion ~~< RdmCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RdmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvValidationComplete = New-Object System.Windows.Forms.Label
$RdmCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 60)
$RdmCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$RdmCsvValidationComplete.TabIndex = 37
$RdmCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($RdmCsvValidationComplete)
#endregion ~~< RdmCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RdmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RdmCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$RdmCsvToolTip.AutoPopDelay = 5000
$RdmCsvToolTip.InitialDelay = 50
$RdmCsvToolTip.IsBalloon = $true
$RdmCsvToolTip.ReshowDelay = 100
$RdmCsvToolTip.SetToolTip($RdmCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Raw Device Mappings (RDMs) in this vCenter.")
#endregion ~~< RdmCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DRS Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsRuleCsvCheckBox.Checked = $true
$DrsRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 80)
$DrsRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsRuleCsvCheckBox.TabIndex = 38
$DrsRuleCsvCheckBox.Text = "Export DRS Rule Info"
$DrsRuleCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsRuleCsvCheckBox)
#endregion ~~< DrsRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 80)
$DrsRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsRuleCsvValidationComplete.TabIndex = 39
$DrsRuleCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsRuleCsvValidationComplete)
#endregion ~~< DrsRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsRuleCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsRuleCsvToolTip.AutoPopDelay = 5000
$DrsRuleCsvToolTip.InitialDelay = 50
$DrsRuleCsvToolTip.IsBalloon = $true
$DrsRuleCsvToolTip.ReshowDelay = 100
$DrsRuleCsvToolTip.SetToolTip($DrsRuleCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Rules"+[char]13+[char]10+"(DRS Rules) in this vCenter.")
#endregion ~~< DrsRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< DRS Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DRS Cluster Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsClusterGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsClusterGroupCsvCheckBox.Checked = $true
$DrsClusterGroupCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsClusterGroupCsvCheckBox.Location = New-Object System.Drawing.Point(620, 100)
$DrsClusterGroupCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsClusterGroupCsvCheckBox.TabIndex = 40
$DrsClusterGroupCsvCheckBox.Text = "Export DRS Cluster Group Info"
$DrsClusterGroupCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsClusterGroupCsvCheckBox)
#endregion ~~< DrsClusterGroupCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsClusterGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsClusterGroupCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 100)
$DrsClusterGroupCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsClusterGroupCsvValidationComplete.TabIndex = 41
$DrsClusterGroupCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsClusterGroupCsvValidationComplete)
#endregion ~~< DrsClusterGroupCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsClusterGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsClusterGroupCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsClusterGroupCsvToolTip.AutoPopDelay = 5000
$DrsClusterGroupCsvToolTip.InitialDelay = 50
$DrsClusterGroupCsvToolTip.IsBalloon = $true
$DrsClusterGroupCsvToolTip.ReshowDelay = 100
$DrsClusterGroupCsvToolTip.SetToolTip($DrsClusterGroupCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Cluster Rules"+[char]13+[char]10+"(DRS Cluster Rules) in this vCenter.")
#endregion ~~< DrsClusterGroupCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< DRS Cluster Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DRS VMHost Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsVmHostRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$DrsVmHostRuleCsvCheckBox.Checked = $true
$DrsVmHostRuleCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$DrsVmHostRuleCsvCheckBox.Location = New-Object System.Drawing.Point(620, 120)
$DrsVmHostRuleCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$DrsVmHostRuleCsvCheckBox.TabIndex = 42
$DrsVmHostRuleCsvCheckBox.Text = "Export DRS VMHost Rule Info"
$DrsVmHostRuleCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($DrsVmHostRuleCsvCheckBox)
#endregion ~~< DrsVmHostRuleCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsVmHostRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvValidationComplete = New-Object System.Windows.Forms.Label
$DrsVmHostRuleCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 120)
$DrsVmHostRuleCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$DrsVmHostRuleCsvValidationComplete.TabIndex = 43
$DrsVmHostRuleCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($DrsVmHostRuleCsvValidationComplete)
#endregion ~~< DrsVmHostRuleCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsVmHostRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrsVmHostRuleCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrsVmHostRuleCsvToolTip.AutoPopDelay = 5000
$DrsVmHostRuleCsvToolTip.InitialDelay = 50
$DrsVmHostRuleCsvToolTip.IsBalloon = $true
$DrsVmHostRuleCsvToolTip.ReshowDelay = 100
$DrsVmHostRuleCsvToolTip.SetToolTip($DrsVmHostRuleCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Distributed Resource Scheduler Host Rules"+[char]13+[char]10+"(DRS Host Rules) in this vCenter.")
#endregion ~~< DrsVmHostRuleCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< DRS VMHost Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Resource Pool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ResourcePoolCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$ResourcePoolCsvCheckBox.Checked = $true
$ResourcePoolCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$ResourcePoolCsvCheckBox.Location = New-Object System.Drawing.Point(620, 140)
$ResourcePoolCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$ResourcePoolCsvCheckBox.TabIndex = 44
$ResourcePoolCsvCheckBox.Text = "Export Resource Pool Info"
$ResourcePoolCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($ResourcePoolCsvCheckBox)
#endregion ~~< ResourcePoolCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ResourcePoolCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvValidationComplete = New-Object System.Windows.Forms.Label
$ResourcePoolCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 140)
$ResourcePoolCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$ResourcePoolCsvValidationComplete.TabIndex = 45
$ResourcePoolCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($ResourcePoolCsvValidationComplete)
#endregion ~~< ResourcePoolCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ResourcePoolCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ResourcePoolCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$ResourcePoolCsvToolTip.AutoPopDelay = 5000
$ResourcePoolCsvToolTip.InitialDelay = 50
$ResourcePoolCsvToolTip.IsBalloon = $true
$ResourcePoolCsvToolTip.ReshowDelay = 100
$ResourcePoolCsvToolTip.SetToolTip($ResourcePoolCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Resource Pools in this vCenter.")
#endregion ~~< ResourcePoolCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Resource Pool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SnapshotCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$SnapshotCsvCheckBox.Checked = $true
$SnapshotCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SnapshotCsvCheckBox.Location = New-Object System.Drawing.Point(620, 160)
$SnapshotCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$SnapshotCsvCheckBox.TabIndex = 46
$SnapshotCsvCheckBox.Text = "Export Snapshot Info"
$SnapshotCsvCheckBox.UseVisualStyleBackColor = $true
$TabCapture.Controls.Add($SnapshotCsvCheckBox)
#endregion ~~< SnapshotCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SnapshotCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvValidationComplete = New-Object System.Windows.Forms.Label
$SnapshotCsvValidationComplete.Location = New-Object System.Drawing.Point(830, 160)
$SnapshotCsvValidationComplete.Size = New-Object System.Drawing.Size(90, 20)
$SnapshotCsvValidationComplete.TabIndex = 47
$SnapshotCsvValidationComplete.Text = ""
$TabCapture.Controls.Add($SnapshotCsvValidationComplete)
#endregion ~~< SnapshotCsvValidationComplete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SnapshotCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SnapshotCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$SnapshotCsvToolTip.AutoPopDelay = 5000
$SnapshotCsvToolTip.InitialDelay = 50
$SnapshotCsvToolTip.IsBalloon = $true
$SnapshotCsvToolTip.ReshowDelay = 100
$SnapshotCsvToolTip.SetToolTip($SnapshotCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Snapshots in this vCenter.")
#endregion ~~< SnapshotCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Snapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Linked vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LinkedvCenterCsvCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvCheckBox = New-Object System.Windows.Forms.CheckBox
$LinkedvCenterCsvCheckBox.Checked = $true
$LinkedvCenterCsvCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$LinkedvCenterCsvCheckBox.Location = New-Object System.Drawing.Point(620, 180)
$LinkedvCenterCsvCheckBox.Size = New-Object System.Drawing.Size(200, 20)
$LinkedvCenterCsvCheckBox.TabIndex = 48
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

#region ~~< LinkedvCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$LinkedvCenterCsvToolTip = New-Object System.Windows.Forms.ToolTip($components)
$LinkedvCenterCsvToolTip.AutoPopDelay = 5000
$LinkedvCenterCsvToolTip.InitialDelay = 50
$LinkedvCenterCsvToolTip.IsBalloon = $true
$LinkedvCenterCsvToolTip.ReshowDelay = 100
$LinkedvCenterCsvToolTip.SetToolTip($LinkedvCenterCsvCheckBox, "Check this box to collect information about"+[char]13+[char]10+"all Linked vCenters in this vCenter.")
#endregion ~~< LinkedvCenterCsvToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Linked vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton = New-Object System.Windows.Forms.Button
$CaptureUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureUncheckButton.Location = New-Object System.Drawing.Point(8, 215)
$CaptureUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureUncheckButton.TabIndex = 50
$CaptureUncheckButton.Text = "Uncheck All"
$CaptureUncheckButton.UseVisualStyleBackColor = $false
$CaptureUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureUncheckButton)
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureUncheckButtonToolTip.AutoPopDelay = 5000
$CaptureUncheckButtonToolTip.InitialDelay = 50
$CaptureUncheckButtonToolTip.IsBalloon = $true
$CaptureUncheckButtonToolTip.ReshowDelay = 100
$CaptureUncheckButtonToolTip.SetToolTip($CaptureUncheckButton, "Click to clear all check boxes above.")
#endregion ~~< CaptureUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton = New-Object System.Windows.Forms.Button
$CaptureCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureCheckButton.Location = New-Object System.Drawing.Point(228, 215)
$CaptureCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureCheckButton.TabIndex = 51
$CaptureCheckButton.Text = "Check All"
$CaptureCheckButton.UseVisualStyleBackColor = $false
$CaptureCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureCheckButton)
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureCheckButtonToolTip.AutoPopDelay = 5000
$CaptureCheckButtonToolTip.InitialDelay = 50
$CaptureCheckButtonToolTip.IsBalloon = $true
$CaptureCheckButtonToolTip.ReshowDelay = 100
$CaptureCheckButtonToolTip.SetToolTip($CaptureCheckButton, "Click to check all check boxes above.")
#endregion ~~< CaptureCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Capture Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton = New-Object System.Windows.Forms.Button
$CaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CaptureButton.Location = New-Object System.Drawing.Point(448, 215)
$CaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$CaptureButton.TabIndex = 52
$CaptureButton.Text = "Collect CSV Data"
$CaptureButton.UseVisualStyleBackColor = $false
$CaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($CaptureButton)
#endregion ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CaptureButtonToolTip.AutoPopDelay = 5000
$CaptureButtonToolTip.InitialDelay = 50
$CaptureButtonToolTip.IsBalloon = $true
$CaptureButtonToolTip.ReshowDelay = 100
$CaptureButtonToolTip.SetToolTip($CaptureButton, "Click to begin collecting environment information"+[char]13+[char]10+"on options selected above.")
#endregion ~~< CaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Capture Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButton = New-Object System.Windows.Forms.Button
$OpenCaptureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenCaptureButton.Location = New-Object System.Drawing.Point(668, 215)
$OpenCaptureButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenCaptureButton.TabIndex = 53
$OpenCaptureButton.Text = "Open CSV Output Folder"
$OpenCaptureButton.UseVisualStyleBackColor = $false
$OpenCaptureButton.BackColor = [System.Drawing.Color]::LightGray
$TabCapture.Controls.Add($OpenCaptureButton)
#endregion ~~< OpenCaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenCaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenCaptureButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$OpenCaptureButtonToolTip.AutoPopDelay = 5000
$OpenCaptureButtonToolTip.InitialDelay = 50
$OpenCaptureButtonToolTip.IsBalloon = $true
$OpenCaptureButtonToolTip.ReshowDelay = 100
$OpenCaptureButtonToolTip.SetToolTip($OpenCaptureButton, "Click once collection is complete to open output folder"+[char]13+[char]10+"seleted above.")
#endregion ~~< OpenCaptureButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add($TabCapture)
#endregion ~~< TabCapture >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$TabDraw = New-Object System.Windows.Forms.TabPage
$TabDraw.Location = New-Object System.Drawing.Point(4, 22)
$TabDraw.Padding = New-Object System.Windows.Forms.Padding(3)
$TabDraw.Size = New-Object System.Drawing.Size(982, 486)
$TabDraw.TabIndex = 2
$TabDraw.Text = "Draw Visio"
$TabDraw.UseVisualStyleBackColor = $true


#region ~~< Csv Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvInput >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#region ~~< DrawCsvInputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawCsvInputButtonToolTip.AutoPopDelay = 5000
$DrawCsvInputButtonToolTip.InitialDelay = 50
$DrawCsvInputButtonToolTip.IsBalloon = $true
$DrawCsvInputButtonToolTip.ReshowDelay = 100
$DrawCsvInputButtonToolTip.SetToolTip($DrawCsvInputButton, "Click to select the folder where the CSV"+[char]39+"s are located."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green.")
#endregion ~~< DrawCsvInputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputLabel = New-Object System.Windows.Forms.Label
$DrawCsvInputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DrawCsvInputLabel.Location = New-Object System.Drawing.Point(10, 10)
$DrawCsvInputLabel.Size = New-Object System.Drawing.Size(190, 25)
$DrawCsvInputLabel.TabIndex = 0
$DrawCsvInputLabel.Text = "CSV Input Folder:"
$TabDraw.Controls.Add($DrawCsvInputLabel)
#endregion ~~< DrawCsvInputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$DrawCsvBrowse.Description = "Select a directory"
$DrawCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< DrawCsvBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< CsvInput >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
$vCenterCsvValidationCheck.add_Click({VCenterCsvValidationCheckClick($vCenterCsvValidationCheck)})
$TabDraw.Controls.Add($vCenterCsvValidationCheck)
#endregion ~~< vCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatacenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DatacenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ClusterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< ClusterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMHostCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHostCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidation = New-Object System.Windows.Forms.Label
$VmHostCsvValidation.Location = New-Object System.Drawing.Point(10, 100)
$VmHostCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmHostCsvValidation.TabIndex = 8
$VmHostCsvValidation.Text = "Host CSV File:"
$TabDraw.Controls.Add($VmHostCsvValidation)
#endregion ~~< VmHostCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHostCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmHostCsvValidationCheck = New-Object System.Windows.Forms.Label
$VmHostCsvValidationCheck.Location = New-Object System.Drawing.Point(180, 100)
$VmHostCsvValidationCheck.Size = New-Object System.Drawing.Size(90, 20)
$VmHostCsvValidationCheck.TabIndex = 9
$VmHostCsvValidationCheck.Text = ""
$VmHostCsvValidationCheck.add_Click({VmHostCsvValidationCheckClick($VmHostCsvValidationCheck)})
$TabDraw.Controls.Add($VmHostCsvValidationCheck)
#endregion ~~< VmHostCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VMHostCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VmCsvValidation = New-Object System.Windows.Forms.Label
$VmCsvValidation.Location = New-Object System.Drawing.Point(10, 120)
$VmCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VmCsvValidation.TabIndex = 10
$VmCsvValidation.Text = "Virtual Machine CSV File:"
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

#endregion ~~< VMCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< TemplateCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< TemplateCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreClusterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DatastoreClusterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DatastoreCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitchCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VsSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VsSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 40)
$VsSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VsSwitchCsvValidation.TabIndex = 18
$VsSwitchCsvValidation.Text = "Standard Switch CSV File:"
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

#endregion ~~< VsSwitchCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPortGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
$VssPortGroupCsvValidationCheck.add_Click({VssPortGroupCsvValidationCheckClick($VssPortGroupCsvValidationCheck)})
$TabDraw.Controls.Add($VssPortGroupCsvValidationCheck)
#endregion ~~< VssPortGroupCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VssPortGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmkernelCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VssVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 80)
$VssVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssVmkernelCsvValidation.TabIndex = 22
$VssVmkernelCsvValidation.Text = "Vss VMkernel CSV File:"
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

#endregion ~~< VssVmkernelCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnicCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VssPnicCsvValidation = New-Object System.Windows.Forms.Label
$VssPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 100)
$VssPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VssPnicCsvValidation.TabIndex = 24
$VssPnicCsvValidation.Text = "Vss pNIC CSV File:"
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

#endregion ~~< VssPnicCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitchCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitchCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdSwitchCsvValidation = New-Object System.Windows.Forms.Label
$VdSwitchCsvValidation.Location = New-Object System.Drawing.Point(270, 120)
$VdSwitchCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdSwitchCsvValidation.TabIndex = 26
$VdSwitchCsvValidation.Text = "Distributed Switch CSV File:"
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

#endregion ~~< VdSwitchCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPortGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< VdsPortGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmkernelCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmkernelCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsVmkernelCsvValidation = New-Object System.Windows.Forms.Label
$VdsVmkernelCsvValidation.Location = New-Object System.Drawing.Point(270, 160)
$VdsVmkernelCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsVmkernelCsvValidation.TabIndex = 30
$VdsVmkernelCsvValidation.Text = "Vds VMkernel CSV File:"
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

#endregion ~~< VdsVmkernelCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnicCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnicCsvValidation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VdsPnicCsvValidation = New-Object System.Windows.Forms.Label
$VdsPnicCsvValidation.Location = New-Object System.Drawing.Point(270, 180)
$VdsPnicCsvValidation.Size = New-Object System.Drawing.Size(165, 20)
$VdsPnicCsvValidation.TabIndex = 32
$VdsPnicCsvValidation.Text = "Vds pNIC CSV File:"
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

#endregion ~~< VdsPnicCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< FolderCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
$FolderCsvValidationCheck.add_Click({Label1Click($FolderCsvValidationCheck)})
$TabDraw.Controls.Add($FolderCsvValidationCheck)
#endregion ~~< FolderCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< FolderCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< RdmCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< RdmCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsRuleCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DrsRuleCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsClusterGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DrsClusterGroupCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrsVmHostRuleCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< DrsVmHostRuleCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ResourcePoolCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< ResourcePoolCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SnapshotCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#endregion ~~< SnapshotCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< LinkedvCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
$LinkedvCenterCsvValidationCheck.TabIndex = 49
$LinkedvCenterCsvValidationCheck.Text = ""
$TabDraw.Controls.Add($LinkedvCenterCsvValidationCheck)
#endregion ~~< LinkedvCenterCsvValidationCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< LinkedvCenterCsv >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton = New-Object System.Windows.Forms.Button
$CsvValidationButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$CsvValidationButton.Location = New-Object System.Drawing.Point(8, 200)
$CsvValidationButton.Size = New-Object System.Drawing.Size(200, 25)
$CsvValidationButton.TabIndex = 50
$CsvValidationButton.Text = "Check for CSVs"
$CsvValidationButton.UseVisualStyleBackColor = $false
$CsvValidationButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($CsvValidationButton)
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ShapesfileSelection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ShapesfileSelectionText >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ShapesfileSelectionText = New-Object System.Windows.Forms.Label
$ShapesfileSelectionText.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ( [System.Byte] ( 0 ) ) )
$ShapesfileSelectionText.Location = New-Object System.Drawing.Point(300, 200)
$ShapesfileSelectionText.Size = New-Object System.Drawing.Size(300, 25)
$ShapesfileSelectionText.TabIndex = 51
$ShapesfileSelectionText.Text = "Select the Visio Shape File:"
$TabDraw.Controls.Add($ShapesfileSelectionText)
#endregion ~~< ShapesfileSelectionText >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ShapesfileSelectionRadioButton1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ShapesfileSelectionRadioButton1 = New-Object System.Windows.Forms.RadioButton
$ShapesfileSelectionRadioButton1.Location = New-Object System.Drawing.Point(600, 200)
$ShapesfileSelectionRadioButton1.Size = New-Object System.Drawing.Size(190, 25)
$ShapesfileSelectionRadioButton1.Checked = $true
$ShapesfileSelectionRadioButton1.TabIndex = 52
$ShapesfileSelectionRadioButton1.TabStop = $true
$ShapesfileSelectionRadioButton1.Text = "vDiagram Default Shapes"
$ShapesfileSelectionRadioButton1.UseVisualStyleBackColor = $true
#endregion ~~< ShapesfileSelectionRadioButton1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ShapesfileSelectionRadioButton2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ShapesfileSelectionRadioButton2 = New-Object System.Windows.Forms.RadioButton
$ShapesfileSelectionRadioButton2.Location = New-Object System.Drawing.Point(800, 200)
$ShapesfileSelectionRadioButton2.Size = New-Object System.Drawing.Size(190, 25)
$ShapesfileSelectionRadioButton2.TabIndex =53
$ShapesfileSelectionRadioButton2.TabStop = $true
$ShapesfileSelectionRadioButton2.Text = "VMware VVD Shapes"
$ShapesfileSelectionRadioButton2.UseVisualStyleBackColor = $true
$TabDraw.Controls.AddRange(@($ShapesfileSelectionRadioButton1,$ShapesfileSelectionRadioButton2))
#endregion ~~< ShapesfileSelectionRadioButton2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ShapesfileSelection >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$CsvValidationButtonToolTip.IsBalloon = $true
$CsvValidationButtonToolTip.SetToolTip($CsvValidationButton, "Click to validate that the required CSV files are present."+[char]13+[char]10+"You must validate files prior to drawing Visio.")
#endregion ~~< CsvValidationButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< ValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Csv Validation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Output Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOutputLabel = New-Object System.Windows.Forms.Label
$VisioOutputLabel.Font = New-Object System.Drawing.Font("Tahoma", 15.0, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$VisioOutputLabel.Location = New-Object System.Drawing.Point(10, 230)
$VisioOutputLabel.Size = New-Object System.Drawing.Size(215, 25)
$VisioOutputLabel.TabIndex = 51
$VisioOutputLabel.Text = "Visio Output Folder:"
$TabDraw.Controls.Add($VisioOutputLabel)
#endregion ~~< VisioOutputLabel >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton = New-Object System.Windows.Forms.Button
$VisioOpenOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$VisioOpenOutputButton.Location = New-Object System.Drawing.Point(230, 230)
$VisioOpenOutputButton.Size = New-Object System.Drawing.Size(740, 25)
$VisioOpenOutputButton.TabIndex = 52
$VisioOpenOutputButton.Text = "Select Visio Output Folder"
$VisioOpenOutputButton.UseVisualStyleBackColor = $false
$VisioOpenOutputButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($VisioOpenOutputButton)
#endregion ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VisioOpenOutputButtonToolTip.AutoPopDelay = 5000
$VisioOpenOutputButtonToolTip.InitialDelay = 50
$VisioOpenOutputButtonToolTip.IsBalloon = $true
$VisioOpenOutputButtonToolTip.ReshowDelay = 100
$VisioOpenOutputButtonToolTip.SetToolTip($VisioOpenOutputButton, "Click to select the folder where the script will output the"+[char]13+[char]10+"Visio Drawings."+[char]13+[char]10+[char]13+[char]10+"Once selected the button will show the path in green.")
#endregion ~~< VisioOpenOutputButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
$VisioBrowse.Description = "Select a directory"
$VisioBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
#endregion ~~< VisioBrowse >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Output Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenter_to_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenter_to_LinkedvCenter_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_Complete = New-Object System.Windows.Forms.Label
$vCenter_to_LinkedvCenter_Complete.Location = New-Object System.Drawing.Point(315, 260)
$vCenter_to_LinkedvCenter_Complete.Size = New-Object System.Drawing.Size(120, 20)
$vCenter_to_LinkedvCenter_Complete.TabIndex = 54
$vCenter_to_LinkedvCenter_Complete.Text = ""
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)
#endregion ~~< vCenter_to_LinkedvCenter_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenter_to_LinkedvCenter_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$vCenter_to_LinkedvCenter_DrawCheckBox.Checked = $true
$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$vCenter_to_LinkedvCenter_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 260)
$vCenter_to_LinkedvCenter_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$vCenter_to_LinkedvCenter_DrawCheckBox.TabIndex = 53
$vCenter_to_LinkedvCenter_DrawCheckBox.Text = "vCenter to Linked vCenter Visio Drawing"
$vCenter_to_LinkedvCenter_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_DrawCheckBox)
#endregion ~~< vCenter_to_LinkedvCenter_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< vCenter_to_LinkedvCenter_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.AutoPopDelay = 5000
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.InitialDelay = 50
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.IsBalloon = $true
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.ReshowDelay = 100
$vCenter_to_LinkedvCenter_DrawCheckBoxToolTip.SetToolTip($vCenter_to_LinkedvCenter_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to all Linked vCenters. This will also add all"+[char]13+[char]10+"metadata to the Visio shapes."+[char]13+[char]10)
#endregion ~~< vCenter_to_LinkedvCenter_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vCenter_to_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_Complete = New-Object System.Windows.Forms.Label
$VM_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 280)
$VM_to_Host_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VM_to_Host_Complete.TabIndex = 56
$VM_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Host_Complete)
#endregion ~~< VM_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Host_DrawCheckBox.Checked = $true
$VM_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 280)
$VM_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Host_DrawCheckBox.TabIndex = 55
$VM_to_Host_DrawCheckBox.Text = "VM to Host Visio Drawing"
$VM_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Host_DrawCheckBox)
#endregion ~~< VM_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Host_DrawCheckBoxToolTip.SetToolTip($VM_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes."+[char]13+[char]10)
#endregion ~~< VM_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Folder_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_Complete = New-Object System.Windows.Forms.Label
$VM_to_Folder_Complete.Location = New-Object System.Drawing.Point(315, 300)
$VM_to_Folder_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VM_to_Folder_Complete.TabIndex = 58
$VM_to_Folder_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Folder_Complete)
#endregion ~~< VM_to_Folder_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Folder_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Folder_DrawCheckBox.Checked = $true
$VM_to_Folder_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Folder_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 300)
$VM_to_Folder_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Folder_DrawCheckBox.TabIndex = 57
$VM_to_Folder_DrawCheckBox.Text = "VM to Folder Visio Drawing"
$VM_to_Folder_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Folder_DrawCheckBox)
#endregion ~~< VM_to_Folder_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Folder_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Folder_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Folder_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Folder_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Folder_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Folder_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Folder_DrawCheckBoxToolTip.SetToolTip($VM_to_Folder_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Folders to Virtual Machines."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#endregion ~~< VM_to_Folder_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMs_with_RDMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_Complete = New-Object System.Windows.Forms.Label
$VMs_with_RDMs_Complete.Location = New-Object System.Drawing.Point(315, 320)
$VMs_with_RDMs_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VMs_with_RDMs_Complete.TabIndex = 60
$VMs_with_RDMs_Complete.Text = ""
$TabDraw.Controls.Add($VMs_with_RDMs_Complete)
#endregion ~~< VMs_with_RDMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMs_with_RDMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMs_with_RDMs_DrawCheckBox.Checked = $true
$VMs_with_RDMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMs_with_RDMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 320)
$VMs_with_RDMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VMs_with_RDMs_DrawCheckBox.TabIndex = 59
$VMs_with_RDMs_DrawCheckBox.Text = "VMs with RDMs Visio Drawing"
$VMs_with_RDMs_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMs_with_RDMs_DrawCheckBox)
#endregion ~~< VMs_with_RDMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMs_with_RDMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMs_with_RDMs_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMs_with_RDMs_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMs_with_RDMs_DrawCheckBoxToolTip.InitialDelay = 50
$VMs_with_RDMs_DrawCheckBoxToolTip.IsBalloon = $true
$VMs_with_RDMs_DrawCheckBoxToolTip.ReshowDelay = 100
$VMs_with_RDMs_DrawCheckBoxToolTip.SetToolTip($VMs_with_RDMs_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Virtual Machines"+[char]13+[char]10+"to Raw Device Mappings (RDMs). This will also add all"+[char]13+[char]10+"metadata to the Visio shapes.")
#endregion ~~< VMs_with_RDMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SRM_Protected_VMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_Complete = New-Object System.Windows.Forms.Label
$SRM_Protected_VMs_Complete.Location = New-Object System.Drawing.Point(315, 340)
$SRM_Protected_VMs_Complete.Size = New-Object System.Drawing.Size(120, 20)
$SRM_Protected_VMs_Complete.TabIndex = 62
$SRM_Protected_VMs_Complete.Text = ""
$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)
#endregion ~~< SRM_Protected_VMs_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SRM_Protected_VMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$SRM_Protected_VMs_DrawCheckBox.Checked = $true
$SRM_Protected_VMs_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$SRM_Protected_VMs_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 340)
$SRM_Protected_VMs_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$SRM_Protected_VMs_DrawCheckBox.TabIndex = 61
$SRM_Protected_VMs_DrawCheckBox.Text = "SRM Protected VMs Visio Drawing"
$SRM_Protected_VMs_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($SRM_Protected_VMs_DrawCheckBox)
#endregion ~~< SRM_Protected_VMs_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SRM_Protected_VMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$SRM_Protected_VMs_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$SRM_Protected_VMs_DrawCheckBoxToolTip.AutoPopDelay = 5000
$SRM_Protected_VMs_DrawCheckBoxToolTip.InitialDelay = 50
$SRM_Protected_VMs_DrawCheckBoxToolTip.IsBalloon = $true
$SRM_Protected_VMs_DrawCheckBoxToolTip.ReshowDelay = 100
$SRM_Protected_VMs_DrawCheckBoxToolTip.SetToolTip($SRM_Protected_VMs_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#endregion ~~< SRM_Protected_VMs_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Datastore_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_Complete = New-Object System.Windows.Forms.Label
$VM_to_Datastore_Complete.Location = New-Object System.Drawing.Point(315, 360)
$VM_to_Datastore_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VM_to_Datastore_Complete.TabIndex = 64
$VM_to_Datastore_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_Datastore_Complete)
#endregion ~~< VM_to_Datastore_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Datastore_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_Datastore_DrawCheckBox.Checked = $true
$VM_to_Datastore_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_Datastore_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 360)
$VM_to_Datastore_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_Datastore_DrawCheckBox.TabIndex = 63
$VM_to_Datastore_DrawCheckBox.Text = "VM to Datastore Visio Drawing"
$VM_to_Datastore_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_Datastore_DrawCheckBox)
#endregion ~~< VM_to_Datastore_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Datastore_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_Datastore_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_Datastore_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_Datastore_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_Datastore_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_Datastore_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_Datastore_DrawCheckBoxToolTip.SetToolTip($VM_to_Datastore_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Datastore Clusters"+[char]13+[char]10+"to Datastores to Virtual Machines. This will also add all"+[char]13+[char]10+"metadata to the Visio shapes.")
#endregion ~~< VM_to_Datastore_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_ResourcePool_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_Complete = New-Object System.Windows.Forms.Label
$VM_to_ResourcePool_Complete.Location = New-Object System.Drawing.Point(315, 380)
$VM_to_ResourcePool_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VM_to_ResourcePool_Complete.TabIndex = 66
$VM_to_ResourcePool_Complete.Text = ""
$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)
#endregion ~~< VM_to_ResourcePool_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_ResourcePool_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VM_to_ResourcePool_DrawCheckBox.Checked = $true
$VM_to_ResourcePool_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VM_to_ResourcePool_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 380)
$VM_to_ResourcePool_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$VM_to_ResourcePool_DrawCheckBox.TabIndex = 65
$VM_to_ResourcePool_DrawCheckBox.Text = "VM to ResourcePool Visio Drawing"
$VM_to_ResourcePool_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VM_to_ResourcePool_DrawCheckBox)
#endregion ~~< VM_to_ResourcePool_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_ResourcePool_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VM_to_ResourcePool_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VM_to_ResourcePool_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VM_to_ResourcePool_DrawCheckBoxToolTip.InitialDelay = 50
$VM_to_ResourcePool_DrawCheckBoxToolTip.IsBalloon = $true
$VM_to_ResourcePool_DrawCheckBoxToolTip.ReshowDelay = 100
$VM_to_ResourcePool_DrawCheckBoxToolTip.SetToolTip($VM_to_ResourcePool_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Resource Pools  to"+[char]13+[char]10+"Virtual Machines. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#endregion ~~< VM_to_ResourcePool_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_Complete = New-Object System.Windows.Forms.Label
$Datastore_to_Host_Complete.Location = New-Object System.Drawing.Point(315, 400)
$Datastore_to_Host_Complete.Size = New-Object System.Drawing.Size(120, 20)
$Datastore_to_Host_Complete.TabIndex = 68
$Datastore_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($Datastore_to_Host_Complete)
#endregion ~~< Datastore_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Datastore_to_Host_DrawCheckBox.Checked = $true
$Datastore_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Datastore_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 400)
$Datastore_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Datastore_to_Host_DrawCheckBox.TabIndex = 67
$Datastore_to_Host_DrawCheckBox.Text = "Datastore to Host Visio Drawing"
$Datastore_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Datastore_to_Host_DrawCheckBox)
#endregion ~~< Datastore_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Datastore_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Datastore_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Datastore_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$Datastore_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$Datastore_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$Datastore_to_Host_DrawCheckBoxToolTip.SetToolTip($Datastore_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Datastores. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#endregion ~~< Datastore_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_Complete = New-Object System.Windows.Forms.Label
$Snapshot_to_VM_Complete.Location = New-Object System.Drawing.Point(315, 420)
$Snapshot_to_VM_Complete.Size = New-Object System.Drawing.Size(120, 20)
$Snapshot_to_VM_Complete.TabIndex = 70
$Snapshot_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($Snapshot_to_VM_Complete)
#endregion ~~< Snapshot_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Snapshot_to_VM_DrawCheckBox.Checked = $true
$Snapshot_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Snapshot_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(10, 420)
$Snapshot_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(300, 20)
$Snapshot_to_VM_DrawCheckBox.TabIndex = 69
$Snapshot_to_VM_DrawCheckBox.Text = "Snapshot to VM Visio Drawing"
$Snapshot_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Snapshot_to_VM_DrawCheckBox)
#endregion ~~< Snapshot_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Snapshot_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Snapshot_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Snapshot_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$Snapshot_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$Snapshot_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$Snapshot_to_VM_DrawCheckBoxToolTip.SetToolTip($Snapshot_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Virtual Machines"+[char]13+[char]10+"to Snapshot Tree. This will also add all metadata to the"+[char]13+[char]10+"Visio shapes.")
#endregion ~~< Snapshot_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PhysicalNIC_to_vSwitch_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_Complete = New-Object System.Windows.Forms.Label
$PhysicalNIC_to_vSwitch_Complete.Location = New-Object System.Drawing.Point(790, 260)
$PhysicalNIC_to_vSwitch_Complete.Size = New-Object System.Drawing.Size(120, 20)
$PhysicalNIC_to_vSwitch_Complete.TabIndex = 72
$PhysicalNIC_to_vSwitch_Complete.Text = ""
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)
#endregion ~~< PhysicalNIC_to_vSwitch_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PhysicalNIC_to_vSwitch_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$PhysicalNIC_to_vSwitch_DrawCheckBox.Checked = $true
$PhysicalNIC_to_vSwitch_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$PhysicalNIC_to_vSwitch_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 260)
$PhysicalNIC_to_vSwitch_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$PhysicalNIC_to_vSwitch_DrawCheckBox.TabIndex = 71
$PhysicalNIC_to_vSwitch_DrawCheckBox.Text = "PhysicalNIC to vSwitch Visio Drawing"
$PhysicalNIC_to_vSwitch_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_DrawCheckBox)
#endregion ~~< PhysicalNIC_to_vSwitch_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.AutoPopDelay = 5000
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.InitialDelay = 50
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.IsBalloon = $true
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.ReshowDelay = 100
$PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip.SetToolTip($PhysicalNIC_to_vSwitch_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Physical NIC. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#endregion ~~< PhysicalNIC_to_vSwitch_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VSS_to_Host_Complete.Location = New-Object System.Drawing.Point(790, 280)
$VSS_to_Host_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VSS_to_Host_Complete.TabIndex = 74
$VSS_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VSS_to_Host_Complete)
#endregion ~~< VSS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSS_to_Host_DrawCheckBox.Checked = $true
$VSS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 280)
$VSS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSS_to_Host_DrawCheckBox.TabIndex = 73
$VSS_to_Host_DrawCheckBox.Text = "Standard Switch to Host Visio Drawing"
$VSS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VSS_to_Host_DrawCheckBox)
#endregion ~~< VSS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSS_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VSS_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VSS_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VSS_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VSS_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VSS_to_Host_DrawCheckBoxToolTip.SetToolTip($VSS_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Port Groups. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#endregion ~~< VSS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VSS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VSS_Complete.Location = New-Object System.Drawing.Point(790, 300)
$VMK_to_VSS_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VMK_to_VSS_Complete.TabIndex = 76
$VMK_to_VSS_Complete.Text = ""
$TabDraw.Controls.Add($VMK_to_VSS_Complete)
#endregion ~~< VMK_to_VSS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VSS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VSS_DrawCheckBox.Checked = $true
$VMK_to_VSS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VSS_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 300)
$VMK_to_VSS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VSS_DrawCheckBox.TabIndex = 75
$VMK_to_VSS_DrawCheckBox.Text = "VMkernel to Standard Switch Visio Drawing"
$VMK_to_VSS_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMK_to_VSS_DrawCheckBox)
#endregion ~~< VMK_to_VSS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VSS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VSS_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMK_to_VSS_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMK_to_VSS_DrawCheckBoxToolTip.InitialDelay = 50
$VMK_to_VSS_DrawCheckBoxToolTip.IsBalloon = $true
$VMK_to_VSS_DrawCheckBoxToolTip.ReshowDelay = 100
$VMK_to_VSS_DrawCheckBoxToolTip.SetToolTip($VMK_to_VSS_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to VMkernels. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#endregion ~~< VMK_to_VSS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VSSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(790, 320)
$VSSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VSSPortGroup_to_VM_Complete.TabIndex = 78
$VSSPortGroup_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)
#endregion ~~< VSSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VSSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VSSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VSSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 320)
$VSSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VSSPortGroup_to_VM_DrawCheckBox.TabIndex = 77
$VSSPortGroup_to_VM_DrawCheckBox.Text = "Standard Switch Port Group to VM Visio Drawing"
$VSSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VSSPortGroup_to_VM_DrawCheckBox)
#endregion ~~< VSSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VSSPortGroup_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$VSSPortGroup_to_VM_DrawCheckBoxToolTip.SetToolTip($VSSPortGroup_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Standard Switches to Port Groups to VMs."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#endregion ~~< VSSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_Complete = New-Object System.Windows.Forms.Label
$VDS_to_Host_Complete.Location = New-Object System.Drawing.Point(790, 340)
$VDS_to_Host_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VDS_to_Host_Complete.TabIndex = 80
$VDS_to_Host_Complete.Text = ""
$TabDraw.Controls.Add($VDS_to_Host_Complete)
#endregion ~~< VDS_to_Host_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDS_to_Host_DrawCheckBox.Checked = $true
$VDS_to_Host_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDS_to_Host_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 340)
$VDS_to_Host_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDS_to_Host_DrawCheckBox.TabIndex = 79
$VDS_to_Host_DrawCheckBox.Text = "Distributed Switch to Host Visio Drawing"
$VDS_to_Host_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VDS_to_Host_DrawCheckBox)
#endregion ~~< VDS_to_Host_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDS_to_Host_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VDS_to_Host_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VDS_to_Host_DrawCheckBoxToolTip.InitialDelay = 50
$VDS_to_Host_DrawCheckBoxToolTip.IsBalloon = $true
$VDS_to_Host_DrawCheckBoxToolTip.ReshowDelay = 100
$VDS_to_Host_DrawCheckBoxToolTip.SetToolTip($VDS_to_Host_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to Port Groups. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#endregion ~~< VDS_to_Host_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VDS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_Complete = New-Object System.Windows.Forms.Label
$VMK_to_VDS_Complete.Location = New-Object System.Drawing.Point(790, 360)
$VMK_to_VDS_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VMK_to_VDS_Complete.TabIndex = 82
$VMK_to_VDS_Complete.Text = ""
$TabDraw.Controls.Add($VMK_to_VDS_Complete)
#endregion ~~< VMK_to_VDS_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VDS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VMK_to_VDS_DrawCheckBox.Checked = $true
$VMK_to_VDS_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VMK_to_VDS_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 360)
$VMK_to_VDS_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VMK_to_VDS_DrawCheckBox.TabIndex = 81
$VMK_to_VDS_DrawCheckBox.Text = "VMkernel to Distributed Switch Visio Drawing"
$VMK_to_VDS_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VMK_to_VDS_DrawCheckBox)
#endregion ~~< VMK_to_VDS_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VDS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VMK_to_VDS_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VMK_to_VDS_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VMK_to_VDS_DrawCheckBoxToolTip.InitialDelay = 50
$VMK_to_VDS_DrawCheckBoxToolTip.IsBalloon = $true
$VMK_to_VDS_DrawCheckBoxToolTip.ReshowDelay = 100
$VMK_to_VDS_DrawCheckBoxToolTip.SetToolTip($VMK_to_VDS_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to VMkernels. This will"+[char]13+[char]10+"also add all metadata to the Visio shapes.")
#endregion ~~< VMK_to_VDS_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_Complete = New-Object System.Windows.Forms.Label
$VDSPortGroup_to_VM_Complete.Location = New-Object System.Drawing.Point(790, 380)
$VDSPortGroup_to_VM_Complete.Size = New-Object System.Drawing.Size(120, 20)
$VDSPortGroup_to_VM_Complete.TabIndex = 84
$VDSPortGroup_to_VM_Complete.Text = ""
$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)
#endregion ~~< VDSPortGroup_to_VM_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$VDSPortGroup_to_VM_DrawCheckBox.Checked = $true
$VDSPortGroup_to_VM_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$VDSPortGroup_to_VM_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 380)
$VDSPortGroup_to_VM_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$VDSPortGroup_to_VM_DrawCheckBox.TabIndex = 83
$VDSPortGroup_to_VM_DrawCheckBox.Text = "Distributed Switch Port Group to VM Visio Drawing"
$VDSPortGroup_to_VM_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($VDSPortGroup_to_VM_DrawCheckBox)
#endregion ~~< VDSPortGroup_to_VM_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VDSPortGroup_to_VM_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.AutoPopDelay = 5000
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.InitialDelay = 50
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.IsBalloon = $true
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.ReshowDelay = 100
$VDSPortGroup_to_VM_DrawCheckBoxToolTip.SetToolTip($VDSPortGroup_to_VM_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to Hosts to"+[char]13+[char]10+"Virtual Distributed Switches to Port Groups to VMs."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#endregion ~~< VDSPortGroup_to_VM_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_to_DRS_Rule_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_Complete = New-Object System.Windows.Forms.Label
$Cluster_to_DRS_Rule_Complete.Location = New-Object System.Drawing.Point(790, 400)
$Cluster_to_DRS_Rule_Complete.Size = New-Object System.Drawing.Size(120, 20)
$Cluster_to_DRS_Rule_Complete.TabIndex = 86
$Cluster_to_DRS_Rule_Complete.Text = ""
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
#endregion ~~< Cluster_to_DRS_Rule_Complete >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_to_DRS_Rule_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_DrawCheckBox = New-Object System.Windows.Forms.CheckBox
$Cluster_to_DRS_Rule_DrawCheckBox.Checked = $true
$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = [System.Windows.Forms.CheckState]::Checked
$Cluster_to_DRS_Rule_DrawCheckBox.Location = New-Object System.Drawing.Point(455, 400)
$Cluster_to_DRS_Rule_DrawCheckBox.Size = New-Object System.Drawing.Size(330, 20)
$Cluster_to_DRS_Rule_DrawCheckBox.TabIndex = 85
$Cluster_to_DRS_Rule_DrawCheckBox.Text = "Cluster to DRS Rule Visio Drawing"
$Cluster_to_DRS_Rule_DrawCheckBox.UseVisualStyleBackColor = $true
$TabDraw.Controls.Add($Cluster_to_DRS_Rule_DrawCheckBox)
#endregion ~~< Cluster_to_DRS_Rule_DrawCheckBox >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_to_DRS_Rule_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip = New-Object System.Windows.Forms.ToolTip($components)
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.AutoPopDelay = 5000
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.InitialDelay = 50
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.IsBalloon = $true
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.ReshowDelay = 100
$Cluster_to_DRS_Rule_DrawCheckBoxToolTip.SetToolTip($Cluster_to_DRS_Rule_DrawCheckBox, "Check this box to create a drawing that depicts this"+[char]13+[char]10+"vCenter to Datacenters to Clusters to DRS Rules."+[char]13+[char]10+"This will also add all metadata to the Visio shapes.")
#endregion ~~< Cluster_to_DRS_Rule_DrawCheckBoxToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButton = New-Object System.Windows.Forms.Button
$DrawUncheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawUncheckButton.Location = New-Object System.Drawing.Point(8, 450)
$DrawUncheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawUncheckButton.TabIndex = 87
$DrawUncheckButton.Text = "Uncheck All"
$DrawUncheckButton.UseVisualStyleBackColor = $false
$DrawUncheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawUncheckButton)
#endregion ~~< DrawUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawUncheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawUncheckButtonToolTip.AutoPopDelay = 5000
$DrawUncheckButtonToolTip.InitialDelay = 50
$DrawUncheckButtonToolTip.IsBalloon = $true
$DrawUncheckButtonToolTip.ReshowDelay = 100
$DrawUncheckButtonToolTip.SetToolTip($DrawUncheckButton, "Click to clear all check boxes above.")
#endregion ~~< DrawUncheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Uncheck Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButton = New-Object System.Windows.Forms.Button
$DrawCheckButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawCheckButton.Location = New-Object System.Drawing.Point(228, 450)
$DrawCheckButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawCheckButton.TabIndex = 88
$DrawCheckButton.Text = "Check All"
$DrawCheckButton.UseVisualStyleBackColor = $false
$DrawCheckButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawCheckButton)
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCheckButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawCheckButtonToolTip.AutoPopDelay = 5000
$DrawCheckButtonToolTip.InitialDelay = 50
$DrawCheckButtonToolTip.IsBalloon = $true
$DrawCheckButtonToolTip.ReshowDelay = 100
$DrawCheckButtonToolTip.SetToolTip($DrawCheckButton, "Click to check all check boxes above.")
#endregion ~~< DrawCheckButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Check Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton = New-Object System.Windows.Forms.Button
$DrawButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$DrawButton.Location = New-Object System.Drawing.Point(448, 450)
$DrawButton.Size = New-Object System.Drawing.Size(200, 25)
$DrawButton.TabIndex = 89
$DrawButton.Text = "Draw Visio"
$DrawButton.UseVisualStyleBackColor = $false
$DrawButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($DrawButton)
#endregion ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$DrawButtonToolTip.AutoPopDelay = 5000
$DrawButtonToolTip.InitialDelay = 50
$DrawButtonToolTip.IsBalloon = $true
$DrawButtonToolTip.ReshowDelay = 100
$DrawButtonToolTip.SetToolTip($DrawButton, "Click to begin drawing environment based on"+[char]13+[char]10+"options selected above.")
#endregion ~~< DrawButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Draw Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open Visio Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButton = New-Object System.Windows.Forms.Button
$OpenVisioButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$OpenVisioButton.Location = New-Object System.Drawing.Point(668, 450)
$OpenVisioButton.Size = New-Object System.Drawing.Size(200, 25)
$OpenVisioButton.TabIndex = 90
$OpenVisioButton.Text = "Open Visio Drawing"
$OpenVisioButton.UseVisualStyleBackColor = $false
$OpenVisioButton.BackColor = [System.Drawing.Color]::LightGray
$TabDraw.Controls.Add($OpenVisioButton)
#endregion ~~< OpenVisioButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< OpenVisioButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OpenVisioButtonToolTip = New-Object System.Windows.Forms.ToolTip($components)
$OpenVisioButtonToolTip.AutoPopDelay = 5000
$OpenVisioButtonToolTip.InitialDelay = 50
$OpenVisioButtonToolTip.IsBalloon = $true
$OpenVisioButtonToolTip.ReshowDelay = 100
$OpenVisioButtonToolTip.SetToolTip($OpenVisioButton, "Click to open Visio drawing once all above check boxes"+[char]13+[char]10+"are marked as completed.")
#endregion ~~< OpenVisioButtonToolTip >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open Visio Button >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.Controls.Add($TabDraw)
#endregion ~~< TabDraw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

$LowerTabs.SelectedIndex = 0
$vDiagram.Controls.Add($LowerTabs)

#endregion ~~< LowerTabs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vDiagram >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Form Creation >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Custom Code >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowershellCheck = $PSVersionTable.PSVersion
if ( $PowershellCheck.Major -ge 4 ) `
{ `
	$PowershellInstalled.Forecolor = "Green"
	$PowershellInstalled.Text = "Installed Version $PowershellCheck"
}
else `
{ `
	$PowershellInstalled.Forecolor = "Red"
	$PowershellInstalled.Text = "Not installed or Powershell version lower than 4"
}
#endregion ~~< PowershellCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$PowerCliModuleCheck = ( Get-Module VMware.PowerCLI -ListAvailable | Where-Object { $_.Name -eq "VMware.PowerCLI" } | Sort-Object Version -Descending )
if ( $null -ne $PowerCliModuleCheck ) `
{ `
	$PowerCliModuleVersion = ( $PowerCliModuleCheck.Version[0] )
	$PowerCliModuleInstalled.Forecolor = "Green"
	$PowerCliModuleInstalled.Text = "Installed Version $PowerCliModuleVersion"
}
else `
{ `
	$PowerCliModuleInstalled.Forecolor = "Red"
	$PowerCliModuleInstalled.Text = "Not Installed"
}

#endregion ~~< PowerCliModuleCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( $null -ne ( Get-PSSnapin -registered | Where-Object { $_.Name -eq "VMware.VimAutomation.Core" } ) ) `
{ `
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerClI Installed"
}
elseif ( $null -ne $PowerCliModuleCheck ) `
{ `
	$PowerCliInstalled.Forecolor = "Green"
	$PowerCliInstalled.Text = "PowerCLI Module Installed"
}
else `
{ `
	$PowerCliInstalled.Forecolor = "Red"
	$PowerCliInstalled.Text = "PowerCLI or PowerCli Module not installed"
}

#endregion ~~< PowerCliCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if ( ( Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) -or $null -ne (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object { $_.DisplayName -like "*Visio*" -and $_.DisplayName -notlike "*Visio View*" } | Select-Object DisplayName ) ) `
{ `
	$VisioInstalled.Forecolor = "Green"
	$VisioInstalled.Text = "Installed"
}
else `
{ `
	$VisioInstalled.Forecolor = "Red"
	$VisioInstalled.Text = "Visio is Not Installed"
}

#endregion ~~< VisioCheck >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Checks >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Buttons >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ConnectButton.Add_MouseClick( { $Connected = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager ; 
	if ( $Connected -eq $null ) `
	{ `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Red ; 
		$ConnectButton.Text = "Unable to Connect"
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Unable to connect to vCenter" -ForegroundColor Red
		}
		if ( $logcapture -eq $true ) `
		{ `
			$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
			$LogCapturePath = $FileDateTime + " " + $DefaultViserver + " - vDiagram_Capture.log"
			Start-Transcript -Path "$LogCapturePath"
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Powershell Module version installed:" $PSVersionTable.PSVersion -ForegroundColor Green
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] PowerCLI Module versions installed:" $PowerCliModuleCheck.Version -ForegroundColor Green
		}
	}
	else `
	{ `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green ;
		$ConnectButton.Text = "Connected to $DefaultViserver."
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Connected to $DefaultVIServer." -ForegroundColor Green
		}
		if ( $logcapture -eq $true ) `
		{ `
			$FileDateTime = (Get-Date -format "yyyy_MM_dd-HH_mm")
			$LogCapturePath = $FileDateTime + " " + $DefaultViserver + " - vDiagram_Capture.log"
			Start-Transcript -Path "$LogCapturePath"
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Powershell Module version installed:" $PSVersionTable.PSVersion -ForegroundColor Green
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] PowerCLI Module versions installed:" $PowerCliModuleCheck.Version -ForegroundColor Green
		}
	}
} )
$ConnectButton.Add_Click( { Connect_vCenter } )
#endregion ~~< ConnectButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCsvOutputButton.Add_Click( { Find_CaptureCsvFolder ; 
	if ( $CaptureCsvFolder -eq $null ) `
	{ `
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$CaptureCsvOutputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		$CaptureCsvOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$CaptureCsvOutputButton.Text = $CaptureCsvFolder
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected CSV export folder = $CaptureCsvFolder" -ForegroundColor Magenta
		}
	}
	Check_CaptureCsvFolder
} )
#endregion ~~< CaptureCsvOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureUncheckButton.Add_Click( { $vCenterCsvCheckBox.CheckState = "UnChecked" ;
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Capture CSV Uncheck All selected." -ForegroundColor Magenta
	}	
} )
#endregion ~~< CaptureUncheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureCheckButton.Add_Click( { $vCenterCsvCheckBox.CheckState = "Checked" ;
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Capture CSV Check All selected." -ForegroundColor Magenta
	}
} )
#endregion ~~< CaptureCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CaptureButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CaptureButton.Add_Click( {
	if( $CaptureCsvFolder -eq $null ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
		$CaptureButton.Forecolor = [System.Drawing.Color]::Red; 
		$CaptureButton.Text = "Folder Not Selected"
	}
	else `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collect CSV Data button selected." -ForegroundColor Magenta
		}
		Write-Host "[$DateTime] CSV collection started." -ForegroundColor Green
		if ( $vCenterCsvCheckBox.Checked -eq "True" ) `
		{ `
			$vCenterCsvValidationComplete.Forecolor = "Blue"
			$vCenterCsvValidationComplete.Text = "Processing ....."
			vCenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$vCenterExportFileComplete = $CsvCompleteDir + "-vCenterExport.csv"
			$vCenterCsvComplete = Test-Path $vCenterExportFileComplete
			
			if ( $vCenterCsvComplete -eq $True ) `
			{ `
				$vCenterCsvValidationComplete.Forecolor = "Green"
				$vCenterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$vCenterCsvValidationComplete.Forecolor = "Red"
				$vCenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		Connect_vCenter
		$Connected = Get-View $DefaultViserver.ExtensionData.Client.ServiceContent.SessionManager
		
		if ( $Connected -eq $null ) { Connect_vCenter } `
		$ConnectButton.Forecolor = [System.Drawing.Color]::Green
		$ConnectButton.Text = "Connected to $DefaultViserver"
		
		if ( $DatacenterCsvCheckBox.Checked -eq "True" ) `
		{ `
			Datacenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatacenterExportFileComplete = $CsvCompleteDir + "-DatacenterExport.csv"
			$DatacenterCsvComplete = Test-Path $DatacenterExportFileComplete
			
			if ( $DatacenterCsvComplete -eq $True ) `
			{ `
				$DatacenterCsvValidationComplete.Forecolor = "Green"
				$DatacenterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DatacenterCsvValidationComplete.Forecolor = "Red"
				$DatacenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $ClusterCsvCheckBox.Checked -eq "True" ) `
		{ `
			Cluster_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$ClusterExportFileComplete = $CsvCompleteDir + "-ClusterExport.csv"
			$ClusterCsvComplete = Test-Path $ClusterExportFileComplete
			
			if ( $ClusterCsvComplete -eq $True ) `
			{ `
				$ClusterCsvValidationComplete.Forecolor = "Green"
				$ClusterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$ClusterCsvValidationComplete.Forecolor = "Red"
				$ClusterCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VmHostCsvCheckBox.Checked -eq "True" ) `
		{ `
			VmHost_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VmHostExportFileComplete = $CsvCompleteDir + "-VmHostExport.csv"
			$VmHostCsvComplete = Test-Path $VmHostExportFileComplete
			
			if ( $VmHostCsvComplete -eq $True ) `
			{ `
				$VmHostCsvValidationComplete.Forecolor = "Green"
				$VmHostCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VmHostCsvValidationComplete.Forecolor = "Red"
				$VmHostCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VmCsvCheckBox.Checked -eq "True" ) `
		{ `
			Vm_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VmExportFileComplete = $CsvCompleteDir + "-VmExport.csv"
			$VmCsvComplete = Test-Path $VmExportFileComplete
			
			if ( $VmCsvComplete -eq $True ) `
			{ `
				$VmCsvValidationComplete.Forecolor = "Green"
				$VmCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VmCsvValidationComplete.Forecolor = "Red"
				$VmCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $TemplateCsvCheckBox.Checked -eq "True" ) `
		{ `
			Template_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$TemplateExportFileComplete = $CsvCompleteDir + "-TemplateExport.csv"
			$TemplateCsvComplete = Test-Path $TemplateExportFileComplete
			
			if ( $TemplateCsvComplete -eq $True ) `
			{ `
				$TemplateCsvValidationComplete.Forecolor = "Green"
				$TemplateCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$TemplateCsvValidationComplete.Forecolor = "Red"
				$TemplateCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DatastoreClusterCsvCheckBox.Checked -eq "True" ) `
		{ `
			DatastoreCluster_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatastoreClusterExportFileComplete = $CsvCompleteDir + "-DatastoreClusterExport.csv"
			$DatastoreClusterCsvComplete = Test-Path $DatastoreClusterExportFileComplete
			
			if ( $DatastoreClusterCsvComplete -eq $True ) `
			{ `
				$DatastoreClusterCsvValidationComplete.Forecolor = "Green"
				$DatastoreClusterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DatastoreClusterCsvValidationComplete.Forecolor = "Red"
				$DatastoreClusterCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DatastoreCsvCheckBox.Checked -eq "True" ) `
		{ `
			Datastore_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DatastoreExportFileComplete = $CsvCompleteDir + "-DatastoreExport.csv"
			$DatastoreCsvComplete = Test-Path $DatastoreExportFileComplete
			
			if ( $DatastoreCsvComplete -eq $True ) `
			{ `
				$DatastoreCsvValidationComplete.Forecolor = "Green"
				$DatastoreCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DatastoreCsvValidationComplete.Forecolor = "Red"
				$DatastoreCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VsSwitchCsvCheckBox.Checked -eq "True" ) `
		{ `
			VsSwitch_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$vSSwitchExportFileComplete = $CsvCompleteDir + "-vSSwitchExport.csv"
			$vSSwitchCsvComplete = Test-Path $vSSwitchExportFileComplete
			
			if ( $vSSwitchCsvComplete -eq $True ) `
			{ `
				$vSSwitchCsvValidationComplete.Forecolor = "Green"
				$vSSwitchCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$vSSwitchCsvValidationComplete.Forecolor = "Red"
				$vSSwitchCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VssPortGroupCsvCheckBox.Checked -eq "True" ) `
		{ `
			VssPort_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssPortGroupExportFileComplete = $CsvCompleteDir + "-VssPortGroupExport.csv"
			$VssPortGroupCsvComplete = Test-Path $VssPortGroupExportFileComplete
			
			if ( $VssPortGroupCsvComplete -eq $True ) `
			{ `
				$VssPortGroupCsvValidationComplete.Forecolor = "Green"
				$VssPortGroupCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VssPortGroupCsvValidationComplete.Forecolor = "Red"
				$VssPortGroupCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VssVmkernelCsvCheckBox.Checked -eq "True" ) `
		{ `
			VssVmk_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssVmkernelExportFileComplete = $CsvCompleteDir + "-VssVmkernelExport.csv"
			$VssVmkernelCsvComplete = Test-Path $VssVmkernelExportFileComplete
			
			if ( $VssVmkernelCsvComplete -eq $True ) `
			{ `
				$VssVmkernelCsvValidationComplete.Forecolor = "Green"
				$VssVmkernelCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VssVmkernelCsvValidationComplete.Forecolor = "Red"
				$VssVmkernelCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VssPnicCsvCheckBox.Checked -eq "True" ) `
		{ `
			VssPnic_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VssPnicExportFileComplete = $CsvCompleteDir + "-VssPnicExport.csv"
			$VssPnicCsvComplete = Test-Path $VssPnicExportFileComplete
			
			if ( $VssPnicCsvComplete -eq $True ) `
			{ `
				$VssPnicCsvValidationComplete.Forecolor = "Green"
				$VssPnicCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VssPnicCsvValidationComplete.Forecolor = "Red"
				$VssPnicCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VdSwitchCsvCheckBox.Checked -eq "True" ) `
		{ `
			VdSwitch_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdSwitchExportFileComplete = $CsvCompleteDir + "-VdSwitchExport.csv"
			$VdSwitchCsvComplete = Test-Path $VdSwitchExportFileComplete
			
			if ( $VdSwitchCsvComplete -eq $True ) `
			{ `
				$VdSwitchCsvValidationComplete.Forecolor = "Green"
				$VdSwitchCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VdSwitchCsvValidationComplete.Forecolor = "Red"
				$VdSwitchCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VdsPortGroupCsvCheckBox.Checked -eq "True" ) `
		{ `
			VdsPort_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsPortGroupExportFileComplete = $CsvCompleteDir + "-VdsPortGroupExport.csv"
			$VdsPortGroupCsvComplete = Test-Path $VdsPortGroupExportFileComplete
			
			if ( $VdsPortGroupCsvComplete -eq $True ) `
			{ `
				$VdsPortGroupCsvValidationComplete.Forecolor = "Green"
				$VdsPortGroupCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VdsPortGroupCsvValidationComplete.Forecolor = "Red"
				$VdsPortGroupCsvValidationComplete.Text = "Not Complete"
				
			}
		}
		
		if ( $VdsVmkernelCsvCheckBox.Checked -eq "True" ) `
		{ `
			VdsVmk_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsVmkernelExportFileComplete = $CsvCompleteDir + "-VdsVmkernelExport.csv"
			$VdsVmkernelCsvComplete = Test-Path $VdsVmkernelExportFileComplete
			
			if ( $VdsVmkernelCsvComplete -eq $True ) `
			{ `
				$VdsVmkernelCsvValidationComplete.Forecolor = "Green"
				$VdsVmkernelCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VdsVmkernelCsvValidationComplete.Forecolor = "Red"
				$VdsVmkernelCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $VdsPnicCsvCheckBox.Checked -eq "True" ) `
		{ `
			VdsPnic_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$VdsPnicExportFileComplete = $CsvCompleteDir + "-VdsPnicExport.csv"
			$VdsPnicCsvComplete = Test-Path $VdsPnicExportFileComplete
			
			if ( $VdsPnicCsvComplete -eq $True ) `
			{ `
				$VdsPnicCsvValidationComplete.Forecolor = "Green"
				$VdsPnicCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$VdsPnicCsvValidationComplete.Forecolor = "Red"
				$VdsPnicCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $FolderCsvCheckBox.Checked -eq "True" ) `
		{ `
			Folder_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$FolderExportFileComplete = $CsvCompleteDir + "-FolderExport.csv"
			$FolderCsvComplete = Test-Path $FolderExportFileComplete
			
			if ( $FolderCsvComplete -eq $True ) `
			{ `
				$FolderCsvValidationComplete.Forecolor = "Green"
				$FolderCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$FolderCsvValidationComplete.Forecolor = "Red"
				$FolderCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $RdmCsvCheckBox.Checked -eq "True" ) `
		{ `
			Rdm_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$RdmExportFileComplete = $CsvCompleteDir + "-RdmExport.csv"
			$RdmCsvComplete = Test-Path $RdmExportFileComplete
			
			if ($RdmCsvComplete -eq $True) `
			{ `
				$RdmCsvValidationComplete.Forecolor = "Green"
				$RdmCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$RdmCsvValidationComplete.Forecolor = "Red"
				$RdmCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DrsRuleCsvCheckBox.Checked -eq "True" ) `
		{ `
			Drs_Rule_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsRuleExportFileComplete = $CsvCompleteDir + "-DrsRuleExport.csv"
			$DrsRuleCsvComplete = Test-Path $DrsRuleExportFileComplete
			
			if ( $DrsRuleCsvComplete -eq $True ) `
			{ `
				$DrsRuleCsvValidationComplete.Forecolor = "Green"
				$DrsRuleCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DrsRuleCsvValidationComplete.Forecolor = "Red"
				$DrsRuleCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DrsClusterGroupCsvCheckBox.Checked -eq "True" ) `
		{ `
			Drs_Cluster_Group_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsClusterGroupExportFileComplete = $CsvCompleteDir + "-DrsClusterGroupExport.csv"
			$DrsClusterGroupCsvComplete = Test-Path $DrsClusterGroupExportFileComplete
			
			if ( $DrsClusterGroupCsvComplete -eq $True ) `
			{ `
				$DrsClusterGroupCsvValidationComplete.Forecolor = "Green"
				$DrsClusterGroupCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DrsClusterGroupCsvValidationComplete.Forecolor = "Red"
				$DrsClusterGroupCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $DrsVmHostRuleCsvCheckBox.Checked -eq "True" ) `
		{ `
			Drs_VmHost_Rule_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$DrsVmHostRuleExportFileComplete = $CsvCompleteDir + "-DrsVmHostRuleExport.csv"
			$DrsVmHostRuleCsvComplete = Test-Path $DrsVmHostRuleExportFileComplete
			
			if ( $DrsVmHostRuleCsvComplete -eq $True ) `
			{ `
				$DrsVmHostRuleCsvValidationComplete.Forecolor = "Green"
				$DrsVmHostRuleCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$DrsVmHostRuleCsvValidationComplete.Forecolor = "Red"
				$DrsVmHostRuleCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $ResourcePoolCsvCheckBox.Checked -eq "True" ) `
		{ `
			Resource_Pool_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$ResourcePoolExportFileComplete = $CsvCompleteDir + "-ResourcePoolExport.csv"
			$ResourcePoolCsvComplete = Test-Path $ResourcePoolExportFileComplete
			
			if ( $ResourcePoolCsvComplete -eq $True ) `
			{ `
				$ResourcePoolCsvValidationComplete.Forecolor = "Green"
				$ResourcePoolCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$ResourcePoolCsvValidationComplete.Forecolor = "Red"
				$ResourcePoolCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $SnapshotCsvCheckBox.Checked -eq "True" ) `
		{ `
			Snapshot_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$SnapshotExportFileComplete = $CsvCompleteDir + "-SnapshotExport.csv"
			$SnapshotCsvComplete = Test-Path $SnapshotExportFileComplete
			
			if ( $SnapshotCsvComplete -eq $True ) `
			{ `
				$SnapshotCsvValidationComplete.Forecolor = "Green"
				$SnapshotCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$SnapshotCsvValidationComplete.Forecolor = "Red"
				$SnapshotCsvValidationComplete.Text = "Not Complete"
			}
		}
		
		if ( $LinkedvCenterCsvCheckBox.Checked -eq "True" ) `
		{ `
			Linked_vCenter_Export
			$CsvCompleteDir = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
			$LinkedvCenterExportFileComplete = $CsvCompleteDir + "-LinkedvCenterExport.csv"
			$LinkedvCenterCsvComplete = Test-Path $LinkedvCenterExportFileComplete
			
			if ( $LinkedvCenterCsvComplete -eq $True ) `
			{ `
				$LinkedvCenterCsvValidationComplete.Forecolor = "Green"
				$LinkedvCenterCsvValidationComplete.Text = "Complete"
			}
			else `
			{ `
				$LinkedvCenterCsvValidationComplete.Forecolor = "Red"
				$LinkedvCenterCsvValidationComplete.Text = "Not Complete"
			}
		}
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss" ; `
		Write-Host "[$DateTime] CSV Collection Complete." -ForegroundColor Yellow ; `
		Disconnect_vCenter
		if ( $logcapture -eq $true ) `
		{ `
			Stop-Transcript
		}
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
	Write-Host "[$DateTime] Opening CSV folder." -ForegroundColor Green
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Open CSV Output Folder button was selected." -ForegroundColor Magenta
	}
})
#endregion ~~< CaptureOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawCsvInputButton.Add_MouseClick( { Find_DrawCsvFolder ;
	if ( $DrawCsvFolder -eq $null ) `
	{ `
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Red ;
		$DrawCsvInputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
		if ( $logdraw -eq $true ) `
		{ `
			$FileDateTime = ( Get-Date -format "yyyy_MM_dd-HH_mm" )
			$LogDrawPath = $FileDateTime + " " + $vCenter + " - vDiagram_Draw.log"
			Start-Transcript -Path "$LogDrawPath"
		}
	}
	else `
	{ `
		$DrawCsvInputButton.Forecolor = [System.Drawing.Color]::Green ;
		$DrawCsvInputButton.Text = $DrawCsvFolder
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected import folder = $DrawCsvFolder" -ForegroundColor Magenta
		}
		if ( $logdraw -eq $true ) `
		{ `
			$FileDateTime = ( Get-Date -format "yyyy_MM_dd-HH_mm" )
			$LogDrawPath = $FileDateTime + " " + $vCenter + " - vDiagram_Draw.log"
			Start-Transcript -Path "$LogDrawPath"
		}
	}
} )
$TabDraw.Controls.Add($DrawCsvInputButton)
#endregion ~~< DrawCsvInputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CsvValidationButton.Add_Click(
{
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Check for CSVs button was clicked." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Validating CSVs." -ForegroundColor Green
	$CsvInputDir = $DrawCsvFolder+"\"+$VcenterTextBox.Text
	$vCenterExportFile = $CsvInputDir + "-vCenterExport.csv"
	$vCenterCsvExists = Test-Path $vCenterExportFile
	$TabDraw.Controls.Add($vCenterCsvValidationCheck)
	if ($vCenterCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] vCenter CSV file present." -ForegroundColor Green
		}
		$vCenterCsvValidationCheck.Forecolor = "Green"
		$vCenterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] vCenter CSV file not present." -ForegroundColor Red
		}
		$vCenterCsvValidationCheck.Forecolor = "Red"
		$vCenterCsvValidationCheck.Text = "Not Present"
	}
	
	$DatacenterExportFile = $CsvInputDir + "-DatacenterExport.csv"
	$DatacenterCsvExists = Test-Path $DatacenterExportFile
	$TabDraw.Controls.Add($DatacenterCsvValidationCheck)
			
	if ($DatacenterCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datacenter CSV file present." -ForegroundColor Green
		}
		$DatacenterCsvValidationCheck.Forecolor = "Green"
		$DatacenterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datacenter CSV file not present." -ForegroundColor Red
		}
		$DatacenterCsvValidationCheck.Forecolor = "Red"
		$DatacenterCsvValidationCheck.Text = "Not Present"
	}
	
	$ClusterExportFile = $CsvInputDir + "-ClusterExport.csv"
	$ClusterCsvExists = Test-Path $ClusterExportFile
	$TabDraw.Controls.Add($ClusterCsvValidationCheck)
			
	if ($ClusterCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Cluster CSV file present." -ForegroundColor Green
		}
		$ClusterCsvValidationCheck.Forecolor = "Green"
		$ClusterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Cluster CSV file not present." -ForegroundColor Red
		}
		$ClusterCsvValidationCheck.Forecolor = "Red"
		$ClusterCsvValidationCheck.Text = "Not Present"
	}
			
	$VmHostExportFile = $CsvInputDir + "-VmHostExport.csv"
	$VmHostCsvExists = Test-Path $VmHostExportFile
	$TabDraw.Controls.Add($VmHostCsvValidationCheck)
			
	if ($VmHostCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Host CSV file present." -ForegroundColor Green
		}
		$VmHostCsvValidationCheck.Forecolor = "Green"
		$VmHostCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Host CSV file not present." -ForegroundColor Red
		}
		$VmHostCsvValidationCheck.Forecolor = "Red"
		$VmHostCsvValidationCheck.Text = "Not Present"
	}
			
	$VmExportFile = $CsvInputDir + "-VmExport.csv"
	$VmCsvExists = Test-Path $VmExportFile
	$TabDraw.Controls.Add($VmCsvValidationCheck)
			
	if ($VmCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Virtual Machine CSV file present." -ForegroundColor Green
		}
		$VmCsvValidationCheck.Forecolor = "Green"
		$VmCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Virtual Machine CSV file not present." -ForegroundColor Red
		}
		$VmCsvValidationCheck.Forecolor = "Red"
		$VmCsvValidationCheck.Text = "Not Present"
	}
	
	$SrmVMsFile = $CsvInputDir + "-VmExport.csv"
	$SrmVMs = Import-Csv $SrmVMsFile
	$SrmVMCount = ( $SrmVMs | Where-Object { $_.SRM.contains("placeholderVm") } ).Count
	if ($SrmVMCount -eq 0) `
	{ `
		$SRM_Protected_VMs_DrawCheckBox.CheckState = "UnChecked"
	}

	$TemplateExportFile = $CsvInputDir + "-TemplateExport.csv"
	$TemplateCsvExists = Test-Path $TemplateExportFile
	$TabDraw.Controls.Add($TemplateCsvValidationCheck)
			
	if ($TemplateCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Template CSV file present." -ForegroundColor Green
		}
		$TemplateCsvValidationCheck.Forecolor = "Green"
		$TemplateCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Template CSV file not present." -ForegroundColor Red
		}
		$TemplateCsvValidationCheck.Forecolor = "Red"
		$TemplateCsvValidationCheck.Text = "Not Present"
	}
			
	$DatastoreClusterExportFile = $CsvInputDir + "-DatastoreClusterExport.csv"
	$DatastoreClusterCsvExists = Test-Path $DatastoreClusterExportFile
	$TabDraw.Controls.Add($DatastoreClusterCsvValidationCheck)
			
	if ($DatastoreClusterCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datastore Cluster CSV file present." -ForegroundColor Green
		}
		$DatastoreClusterCsvValidationCheck.Forecolor = "Green"
		$DatastoreClusterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datastore Cluster CSV file not present." -ForegroundColor Red
		}
		$DatastoreClusterCsvValidationCheck.Forecolor = "Red"
		$DatastoreClusterCsvValidationCheck.Text = "Not Present"
	}
			
	$DatastoreExportFile = $CsvInputDir + "-DatastoreExport.csv"
	$DatastoreCsvExists = Test-Path $DatastoreExportFile
	$TabDraw.Controls.Add($DatastoreCsvValidationCheck)
			
	if ($DatastoreCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datastore CSV file present." -ForegroundColor Green
		}
		$DatastoreCsvValidationCheck.Forecolor = "Green"
		$DatastoreCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Datastore CSV file not present." -ForegroundColor Red
		}
		$DatastoreCsvValidationCheck.Forecolor = "Red"
		$DatastoreCsvValidationCheck.Text = "Not Present"
	}
			
	$VsSwitchExportFile = $CsvInputDir + "-VsSwitchExport.csv"
	$VsSwitchCsvExists = Test-Path $VsSwitchExportFile
	$TabDraw.Controls.Add($VsSwitchCsvValidationCheck)
			
	if ($VsSwitchCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Standard Switch CSV file present." -ForegroundColor Green
		}
		$VsSwitchCsvValidationCheck.Forecolor = "Green"
		$VsSwitchCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Standard Switch CSV file not present." -ForegroundColor Red
		}
		$VsSwitchCsvValidationCheck.Forecolor = "Red"
		$VsSwitchCsvValidationCheck.Text = "Not Present"
		$VSS_to_Host_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssPortGroupExportFile = $CsvInputDir + "-VssPortGroupExport.csv"
	$VssPortGroupCsvExists = Test-Path $VssPortGroupExportFile
	$TabDraw.Controls.Add($VssPortGroupCsvValidationCheck)
			
	if ($VssPortGroupCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS Port Group CSV file present." -ForegroundColor Green
		}
		$VssPortGroupCsvValidationCheck.Forecolor = "Green"
		$VssPortGroupCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS Port Group CSV file not present." -ForegroundColor Red
		}
		$VssPortGroupCsvValidationCheck.Forecolor = "Red"
		$VssPortGroupCsvValidationCheck.Text = "Not Present"
		$VSSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssVmkernelExportFile = $CsvInputDir + "-VssVmkernelExport.csv"
	$VssVmkernelCsvExists = Test-Path $VssVmkernelExportFile
	$TabDraw.Controls.Add($VssVmkernelCsvValidationCheck)
			
	if ($VssVmkernelCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS VMkernel CSV file present." -ForegroundColor Green
		}
		$VssVmkernelCsvValidationCheck.Forecolor = "Green"
		$VssVmkernelCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS VMkernel CSV file not present." -ForegroundColor Red
		}
		$VssVmkernelCsvValidationCheck.Forecolor = "Red"
		$VssVmkernelCsvValidationCheck.Text = "Not Present"
		$VMK_to_VSS_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VssPnicExportFile = $CsvInputDir + "-VssPnicExport.csv"
	$VssPnicCsvExists = Test-Path $VssPnicExportFile
	$TabDraw.Controls.Add($VssPnicCsvValidationCheck)
			
	if ($VssPnicCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS pNIC CSV file present." -ForegroundColor Green
		}
		$VssPnicCsvValidationCheck.Forecolor = "Green"
		$VssPnicCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VSS pNIC CSV file not present." -ForegroundColor Red
		}
		$VssPnicCsvValidationCheck.Forecolor = "Red"
		$VssPnicCsvValidationCheck.Text = "Not Present"
	}
			
	$VdSwitchExportFile = $CsvInputDir + "-VdSwitchExport.csv"
	$VdSwitchCsvExists = Test-Path $VdSwitchExportFile
	$TabDraw.Controls.Add($VdSwitchCsvValidationCheck)
			
	if ($VdSwitchCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Distributed Switch CSV file present." -ForegroundColor Green
		}
		$VdSwitchCsvValidationCheck.Forecolor = "Green"
		$VdSwitchCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Distributed Switch CSV file not present." -ForegroundColor Red
		}
		$VdSwitchCsvValidationCheck.Forecolor = "Red"
		$VdSwitchCsvValidationCheck.Text = "Not Present"
		$VDS_to_Host_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsPortGroupExportFile = $CsvInputDir + "-VdsPortGroupExport.csv"
	$VdsPortGroupCsvExists = Test-Path $VdsPortGroupExportFile
	$TabDraw.Controls.Add($VdsPortGroupCsvValidationCheck)
			
	if ($VdsPortGroupCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS Port Group CSV file not present." -ForegroundColor Green
		}
		$VdsPortGroupCsvValidationCheck.Forecolor = "Green"
		$VdsPortGroupCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS Port Group CSV file not present." -ForegroundColor Red
		}
		$VdsPortGroupCsvValidationCheck.Forecolor = "Red"
		$VdsPortGroupCsvValidationCheck.Text = "Not Present"
		$VDSPortGroup_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsVmkernelExportFile = $CsvInputDir + "-VdsVmkernelExport.csv"
	$VdsVmkernelCsvExists = Test-Path $VdsVmkernelExportFile
	$TabDraw.Controls.Add($VdsVmkernelCsvValidationCheck)
			
	if ($VdsVmkernelCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS VMkernel CSV file present." -ForegroundColor Green
		}
		$VdsVmkernelCsvValidationCheck.Forecolor = "Green"
		$VdsVmkernelCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS VMkernel CSV file not present." -ForegroundColor Red
		}
		$VdsVmkernelCsvValidationCheck.Forecolor = "Red"
		$VdsVmkernelCsvValidationCheck.Text = "Not Present"
		$VMK_to_VDS_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$VdsPnicExportFile = $CsvInputDir + "-VdsPnicExport.csv"
	$VdsPnicCsvExists = Test-Path $VdsPnicExportFile
	$TabDraw.Controls.Add($VdsPnicCsvValidationCheck)
			
	if ($VdsPnicCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS pNIC CSV file present." -ForegroundColor Green
		}
		$VdsPnicCsvValidationCheck.Forecolor = "Green"
		$VdsPnicCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VDS pNIC CSV file not present." -ForegroundColor Red
		}
		$VdsPnicCsvValidationCheck.Forecolor = "Red"
		$VdsPnicCsvValidationCheck.Text = "Not Present"
	}
			
	$FolderExportFile = $CsvInputDir + "-FolderExport.csv"
	$FolderCsvExists = Test-Path $FolderExportFile
	$TabDraw.Controls.Add($FolderCsvValidationCheck)
			
	if ($FolderCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder CSV file present." -ForegroundColor Green
		}
		$FolderCsvValidationCheck.Forecolor = "Green"
		$FolderCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder CSV file not present." -ForegroundColor Red
		}
		$FolderCsvValidationCheck.Forecolor = "Red"
		$FolderCsvValidationCheck.Text = "Not Present"
	}
			
	$RdmExportFile = $CsvInputDir + "-RdmExport.csv"
	$RdmCsvExists = Test-Path $RdmExportFile
	$TabDraw.Controls.Add($RdmCsvValidationCheck)
			
	if ($RdmCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] RDM CSV file present." -ForegroundColor Green
		}
		$RdmCsvValidationCheck.Forecolor = "Green"
		$RdmCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] RDM CSV file not present." -ForegroundColor Red
		}
		$RdmCsvValidationCheck.Forecolor = "Red"
		$RdmCsvValidationCheck.Text = "Not Present"
		$VMs_with_RDMs_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$DrsRuleExportFile = $CsvInputDir + "-DrsRuleExport.csv"
	$DrsRuleCsvExists = Test-Path $DrsRuleExportFile
	$TabDraw.Controls.Add($DrsRuleCsvValidationCheck)
			
	if ($DrsRuleCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS Rule CSV file present." -ForegroundColor Green
		}
		$DrsRuleCsvValidationCheck.Forecolor = "Green"
		$DrsRuleCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS Rule CSV file not present." -ForegroundColor Red
		}
		$DrsRuleCsvValidationCheck.Forecolor = "Red"
		$DrsRuleCsvValidationCheck.Text = "Not Present"
		$Cluster_to_DRS_Rule_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$DrsClusterGroupExportFile = $CsvInputDir + "-DrsClusterGroupExport.csv"
	$DrsClusterGroupCsvExists = Test-Path $DrsClusterGroupExportFile
	$TabDraw.Controls.Add($DrsClusterGroupCsvValidationCheck)
			
	if ($DrsClusterGroupCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS Cluster Group CSV file present." -ForegroundColor Green
		}
		$DrsClusterGroupCsvValidationCheck.Forecolor = "Green"
		$DrsClusterGroupCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS Cluster Group CSV file not present." -ForegroundColor Red
		}
		$DrsClusterGroupCsvValidationCheck.Forecolor = "Red"
		$DrsClusterGroupCsvValidationCheck.Text = "Not Present"
	}
			
	$DrsVmHostRuleExportFile = $CsvInputDir + "-DrsVmHostRuleExport.csv"
	$DrsVmHostRuleCsvExists = Test-Path $DrsVmHostRuleExportFile
	$TabDraw.Controls.Add($DrsVmHostRuleCsvValidationCheck)
			
	if ($DrsVmHostRuleCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS VmHost Rule CSV file present." -ForegroundColor Green
		}
		$DrsVmHostRuleCsvValidationCheck.Forecolor = "Green"
		$DrsVmHostRuleCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] DRS VmHost Rule CSV file not present." -ForegroundColor Red
		}
		$DrsVmHostRuleCsvValidationCheck.Forecolor = "Red"
		$DrsVmHostRuleCsvValidationCheck.Text = "Not Present"
	}
			
	$ResourcePoolExportFile = $CsvInputDir + "-ResourcePoolExport.csv"
	$ResourcePoolCsvExists = Test-Path $ResourcePoolExportFile
	$TabDraw.Controls.Add($ResourcePoolCsvValidationCheck)
			
	if ($ResourcePoolCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Resource Pool CSV file present." -ForegroundColor Green
		}
		$ResourcePoolCsvValidationCheck.Forecolor = "Green"
		$ResourcePoolCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Resource Pool CSV file not present." -ForegroundColor Red
		}
		$ResourcePoolCsvValidationCheck.Forecolor = "Red"
		$ResourcePoolCsvValidationCheck.Text = "Not Present"
		$VM_to_ResourcePool_DrawCheckBox.CheckState = "UnChecked"
	}
			
	$SnapshotExportFile = $CsvInputDir + "-SnapshotExport.csv"
	$SnapshotCsvExists = Test-Path $SnapshotExportFile
	$TabDraw.Controls.Add($SnapshotCsvValidationCheck)
			
	if ($SnapshotCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Snapshot CSV file present." -ForegroundColor Green
		}
		$SnapshotCsvValidationCheck.Forecolor = "Green"
		$SnapshotCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Snapshot CSV file not present." -ForegroundColor Red
		}
		$SnapshotCsvValidationCheck.Forecolor = "Red"
		$SnapshotCsvValidationCheck.Text = "Not Present"
		$Snapshot_to_VM_DrawCheckBox.CheckState = "UnChecked"
	}
	
	$LinkedvCenterExportFile = $CsvInputDir + "-LinkedvCenterExport.csv"
	$LinkedvCenterCsvExists = Test-Path $LinkedvCenterExportFile
	$TabDraw.Controls.Add($LinkedvCenterCsvValidationCheck)
			
	if ($LinkedvCenterCsvExists -eq $True) `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Linked vCenter CSV file present." -ForegroundColor Green
		}
		$LinkedvCenterCsvValidationCheck.Forecolor = "Green"
		$LinkedvCenterCsvValidationCheck.Text = "Present"
	}
	else `
	{ `
		if ( $logdraw -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Linked vCenter CSV file not present." -ForegroundColor Red
		}
		$LinkedvCenterCsvValidationCheck.Forecolor = "Red"
		$LinkedvCenterCsvValidationCheck.Text = "Not Present"
		$vCenter_to_LinkedvCenter_DrawCheckBox.CheckState = "UnChecked"
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] CSV Validation complete." -ForegroundColor Green
} )
$CsvValidationButton.Add_MouseClick({ $CsvValidationButton.Forecolor = [System.Drawing.Color]::Green ; $CsvValidationButton.Text = "CSV Validation Complete" })
#endregion ~~< CsvValidationButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VisioOpenOutputButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$VisioOpenOutputButton.Add_MouseClick({Find_DrawVisioFolder; 
	if( $VisioFolder -eq $null ) `
	{ `
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Red ;
		$VisioOpenOutputButton.Text = "Folder Not Selected"
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		$VisioOpenOutputButton.Forecolor = [System.Drawing.Color]::Green ;
		$VisioOpenOutputButton.Text = $VisioFolder
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Selected Visio export folder = $VisioFolder" -ForegroundColor Magenta
		}
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Draw - Uncheck All selected." -ForegroundColor Magenta
	}
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Draw - Check All selected." -ForegroundColor Magenta
	}
} )
#endregion ~~< DrawCheckButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DrawButton >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DrawButton.Add_Click({ `
	if( $VisioFolder -eq $null ) `
	{ `
		$DrawButton.Forecolor = [System.Drawing.Color]::Red ;
		$DrawButton.Text = "Folder Not Selected"
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Folder not selected." -ForegroundColor Red
		}
	}
	else `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Draw Visio button was clicked." -ForegroundColor Magenta
		}
		Write-Host "[$DateTime] Starting drawings." -ForegroundColor Green
		$Host.UI.RawUI.WindowTitle = "vDiagram $MyVer Drawing $vCenter"
		$DrawButton.Forecolor = [System.Drawing.Color]::Blue ;
		$DrawButton.Text = "Drawing Please Wait" ;
        Shapefile_Select;
		Create_Visio_Base;
		if ($vCenter_to_LinkedvCenter_DrawCheckBox.Checked -eq "True") `
		{ `
			$vCenter_to_LinkedvCenter_Complete.Forecolor = "Blue"
			$vCenter_to_LinkedvCenter_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)
			vCenter_to_LinkedvCenter
			$vCenter_to_LinkedvCenter_Complete.Forecolor = "Green"
			$vCenter_to_LinkedvCenter_Complete.Text = "Complete"
		};
		if ($VM_to_Host_DrawCheckBox.Checked -eq "True") `
		{ `
			$VM_to_Host_Complete.Forecolor = "Blue"
			$VM_to_Host_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VM_to_Host_Complete)
			VM_to_Host
			$VM_to_Host_Complete.Forecolor = "Green"
			$VM_to_Host_Complete.Text = "Complete"
		}
		if ($VM_to_Folder_DrawCheckBox.Checked -eq "True") `
		{ `
			$VM_to_Folder_Complete.Forecolor = "Blue"
			$VM_to_Folder_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VM_to_Folder_Complete)
			VM_to_Folder
			$VM_to_Folder_Complete.Forecolor = "Green"
			$VM_to_Folder_Complete.Text = "Complete"
		}
		if ($VMs_with_RDMs_DrawCheckBox.Checked -eq "True") `
		{ `
			$VMs_with_RDMs_Complete.Forecolor = "Blue"
			$VMs_with_RDMs_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VMs_with_RDMs_Complete)
			VMs_with_RDMs
			$VMs_with_RDMs_Complete.Forecolor = "Green"
			$VMs_with_RDMs_Complete.Text = "Complete"
		}
		if ($SRM_Protected_VMs_DrawCheckBox.Checked -eq "True") `
		{ `
			$SRM_Protected_VMs_Complete.Forecolor = "Blue"
			$SRM_Protected_VMs_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)
			SRM_Protected_VMs
			$SRM_Protected_VMs_Complete.Forecolor = "Green"
			$SRM_Protected_VMs_Complete.Text = "Complete"
		}
		if ($VM_to_Datastore_DrawCheckBox.Checked -eq "True") `
		{ `
			$VM_to_Datastore_Complete.Forecolor = "Blue"
			$VM_to_Datastore_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VM_to_Datastore_Complete)
			VM_to_Datastore
			$VM_to_Datastore_Complete.Forecolor = "Green"
			$VM_to_Datastore_Complete.Text = "Complete"
		}
		if ($VM_to_ResourcePool_DrawCheckBox.Checked -eq "True") `
		{ `
			$VM_to_ResourcePool_Complete.Forecolor = "Blue"
			$VM_to_ResourcePool_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)
			VM_to_ResourcePool
			$VM_to_ResourcePool_Complete.Forecolor = "Green"
			$VM_to_ResourcePool_Complete.Text = "Complete"
		}
		if ($Datastore_to_Host_DrawCheckBox.Checked -eq "True") `
		{ `
			$Datastore_to_Host_Complete.Forecolor = "Blue"
			$Datastore_to_Host_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($Datastore_to_Host_Complete)
			Datastore_to_Host
			$Datastore_to_Host_Complete.Forecolor = "Green"
			$Datastore_to_Host_Complete.Text = "Complete"
		}
		if ($Snapshot_to_VM_DrawCheckBox.Checked -eq "True") `
		{ `
			$Snapshot_to_VM_Complete.Forecolor = "Blue"
			$Snapshot_to_VM_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($Snapshot_to_VM_Complete)
			Snapshot_to_VM
			$Snapshot_to_VM_Complete.Forecolor = "Green"
			$Snapshot_to_VM_Complete.Text = "Complete"
		};
		if ($PhysicalNIC_to_vSwitch_DrawCheckBox.Checked -eq "True") `
		{ `
			$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
			$PhysicalNIC_to_vSwitch_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)
			PhysicalNIC_to_vSwitch
			$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Green"
			$PhysicalNIC_to_vSwitch_Complete.Text = "Complete"
		}
		if ($VSS_to_Host_DrawCheckBox.Checked -eq "True") `
		{ `
			$VSS_to_Host_Complete.Forecolor = "Blue"
			$VSS_to_Host_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VSS_to_Host_Complete)
			VSS_to_Host
			$VSS_to_Host_Complete.Forecolor = "Green"
			$VSS_to_Host_Complete.Text = "Complete"
		}
		if ($VMK_to_VSS_DrawCheckBox.Checked -eq "True") `
		{ `
			$VMK_to_VSS_Complete.Forecolor = "Blue"
			$VMK_to_VSS_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VMK_to_VSS_Complete)
			VMK_to_VSS
			$VMK_to_VSS_Complete.Forecolor = "Green"
			$VMK_to_VSS_Complete.Text = "Complete"
		}
		if ($VSSPortGroup_to_VM_DrawCheckBox.Checked -eq "True") `
		{ `
			$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
			$VSSPortGroup_to_VM_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)
			VSSPortGroup_to_VM
			$VSSPortGroup_to_VM_Complete.Forecolor = "Green"
			$VSSPortGroup_to_VM_Complete.Text = "Complete"
		}
		if ($VDS_to_Host_DrawCheckBox.Checked -eq "True") `
		{ `
			$VDS_to_Host_Complete.Forecolor = "Blue"
			$VDS_to_Host_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VDS_to_Host_Complete)
			VDS_to_Host
			$VDS_to_Host_Complete.Forecolor = "Green"
			$VDS_to_Host_Complete.Text = "Complete"
		}
		if ($VMK_to_VDS_DrawCheckBox.Checked -eq "True") `
		{ `
			$VMK_to_VDS_Complete.Forecolor = "Blue"
			$VMK_to_VDS_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VMK_to_VDS_Complete)
			VMK_to_VDS
			$VMK_to_VDS_Complete.Forecolor = "Green"
			$VMK_to_VDS_Complete.Text = "Complete"
		}
		if ($VDSPortGroup_to_VM_DrawCheckBox.Checked -eq "True") `
		{ `
			$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
			$VDSPortGroup_to_VM_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)
			VDSPortGroup_to_VM
			$VDSPortGroup_to_VM_Complete.Forecolor = "Green"
			$VDSPortGroup_to_VM_Complete.Text = "Complete"
		}
		if ($Cluster_to_DRS_Rule_DrawCheckBox.Checked -eq "True") `
		{ `
			$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
			$Cluster_to_DRS_Rule_Complete.Text = "Processing ..."
			$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			Cluster_to_DRS_Rule
			$Cluster_to_DRS_Rule_Complete.Forecolor = "Green"
			$Cluster_to_DRS_Rule_Complete.Text = "Complete"
		};
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss" ; `
	Write-Host "[$DateTime] Visio Drawings Complete. Click Open Visio Drawing button to proceed." -ForegroundColor Yellow ; `
	$DrawButton.Forecolor = [System.Drawing.Color]::Green; $DrawButton.Text = "Visio Drawings Complete" `
	} `
	
	# Follow us on Twitter Prompt
	$LikeUs =  [System.Windows.Forms.MessageBox]::Show( "Did you find this script helpful? Click 'Yes' to follow us on Twitter and 'No' cancel.","Follow us on Twitter.",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning )
	switch  ( $LikeUs ) `
		{ `
			'Yes' 
			{ `
				Start-Process 'https://twitter.com/vDiagramProject'
				[System.Windows.Forms.MessageBox]::Show("Your Visio Drawing is now complete. Please validate all drawings to ensure items are not missing.

Please click on the Open Visio Drawing button now to open and compress the file.","vDiagram is now complete!",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
			}
			'No'
			{ `
				[System.Windows.Forms.MessageBox]::Show("Your Visio Drawing is now complete. Please validate all drawings to ensure items are not missing.

Please click on the Open Visio Drawing button now to open and compress the file.","vDiagram is now complete!",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
			}
		}

	if ( $logdraw -eq $true ) `
	{ `
		Stop-Transcript
	}
})
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Opening drawing." -ForegroundColor Green
	}
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
	$Host.UI.RawUI.WindowTitle = "vDiagram $MyVer connected to $vCenter"
}
#endregion ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Disconnect_vCenter
{
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Disconnect-ViServer * -Confirm:$false
	$Host.UI.RawUI.WindowTitle = "vDiagram $MyVer disconnected from $vCenter"
	if ( $debug -eq $true )`
	{ `
		Write-Host "[$DateTime] Disconnected from $Vcenter successfully." -ForegroundColor Green
	}
	Write-Host "[$DateTime] Click Open CSV Output Folder to view CSVs or proceed to drawing Visio." -ForegroundColor Yellow
}
#endregion ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< vCenter Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Folder Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_CaptureCsvFolder
{
	$CaptureCsvBrowseLoop = $True
	while ($CaptureCsvBrowseLoop) `
	{ `
		if ($CaptureCsvBrowse.ShowDialog() -eq "OK") `
		{ `
			$CaptureCsvBrowseLoop = $False
		}
		else `
		{ `
			$CaptureCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show( "You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel, [System.Windows.Forms.MessageBoxIcon]::Question )
			if ($CaptureCsvBrowseRes -eq "Cancel") `
			{ `
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
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	$CheckContentPath = $CaptureCsvFolder + "\" + $VcenterTextBox.Text
	$CheckContentDir = $CheckContentPath + "*.csv"
	$CheckContent = Test-Path $CheckContentDir
	if ($CheckContent -eq "True")
	{
		$CheckContents_CaptureCsvFolder =  [System.Windows.Forms.MessageBox]::Show( "Files where found in the folder. Would you like to delete these files? Click 'Yes' to delete and 'No' move files to a new folder.","Warning!",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question )
		switch  ($CheckContents_CaptureCsvFolder) `
		{ `
			'Yes' 
			{ `
				Remove-Item $CheckContentDir
				if ( $debug -eq $true )`
				{ `
					Write-Host "[$DateTime] Files were present in folder. Deleting files from folder."
				}
			}
			'No'
			{ `
				$CheckContentCsvBrowse = New-Object System.Windows.Forms.FolderBrowserDialog
				$CheckContentCsvBrowse.Description = "Select a directory to copy files to"
				$CheckContentCsvBrowse.RootFolder = [System.Environment+SpecialFolder]::MyComputer
				$CheckContentCsvBrowse.ShowDialog()
				$global:NewContentCsvFolder = $CheckContentCsvBrowse.SelectedPath
				Copy-Item -Path $CheckContentDir -Destination $NewContentCsvFolder
				Remove-Item $CheckContentDir
				if ( $debug -eq $true )`
				{ `
					Write-Host "[$DateTime] Files were present in folder. Moving old files to $NewContentCsvFolder"
				}
			}
		}
	}
}
#endregion ~~< Check_CaptureCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Find_DrawCsvFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Find_DrawCsvFolder
{
	$DrawCsvBrowseLoop = $True
	while ($DrawCsvBrowseLoop) `
	{ `
		if ($DrawCsvBrowse.ShowDialog() -eq "OK") `
		{ `
			$DrawCsvBrowseLoop = $False
		}
		else `
		{
			$DrawCsvBrowseRes = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
			if ($DrawCsvBrowseRes -eq "Cancel") `
			{ `
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
	while($VisioBrowseLoop) `
	{ `
		if ($VisioBrowse.ShowDialog() -eq "OK") `
		{ `
			$VisioBrowseLoop = $False
		}
		else `
		{ `
			$VisioBrowseRes = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
			if($VisioBrowseRes -eq "Cancel") `
			{ `
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
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export vCenter Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting vCenter Info." -ForegroundColor Green
	$vCenterExportFile = "$CaptureCsvFolder\$vCenter-vCenterExport.csv"
	if ( $debug -eq $true )`
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on vCenter object -" $DefaultVIServer.Name
		}
	$global:DefaultVIServer | `
	Select-Object `
		@{ Name = "Name" ; Expression = { $_.Name } }, `
		@{ Name = "Version" ; Expression = { $_.Version } }, `
		@{ Name = "Build" ; Expression = { $_.Build } }, `
		@{ Name = "OsType" ; Expression = { $_.ExtensionData.Content.About.OsType } }, `
		@{ Name = "DatacenterCount" ; Expression = { ( Get-Datacenter ).Count } }, `
		@{ Name = "ClusterCount" ; Expression = { ( Get-Cluster ).Count } }, `
		@{ Name = "HostCount" ; Expression = { ( Get-VMHost ).Count } }, `
		@{ Name = "VMCount" ; Expression = { ( Get-VM ).Count } }, `
		@{ Name = "PoweredOnVMCount" ; Expression = { ( Get-VM | Where-Object { $_.PowerState -eq "PoweredOn" } ).Count } }, `
		@{ Name = "TemplateCount" ; Expression = { ( Get-Template ).Count } }, `
		@{ Name = "IsConnected" ; Expression = { $_.IsConnected } }, `
		@{ Name = "ServiceUri" ; Expression = { $_.ServiceUri } }, `
		@{ Name = "Port" ; Expression = { $_.Port } }, `
		@{ Name = "ProductLine" ; Expression = { $_.ProductLine } }, `
		@{ Name = "InstanceUuid" ; Expression = { $_.InstanceUuid } }, `
		@{ Name = "RefCount" ; Expression = { $_.RefCount } }, `
		@{ Name = "ServerClock" ; Expression = { $_.ExtensionData.ServerClock } }, `
		@{ Name = "ProvisioningSupported" ; Expression = { $_.ExtensionData.Capability.ProvisioningSupported } }, `
		@{ Name = "MultiHostSupported" ; Expression = { $_.ExtensionData.Capability.MultiHostSupported } }, `
		@{ Name = "UserShellAccessSupported" ; Expression = { $_.ExtensionData.Capability.UserShellAccessSupported } }, `
		@{ Name = "NetworkBackupAndRestoreSupported" ; Expression = { $_.ExtensionData.Capability.NetworkBackupAndRestoreSupported } }, `
		@{ Name = "FtDrsWithoutEvcSupported" ; Expression = { $_.ExtensionData.Capability.FtDrsWithoutEvcSupported } }, `
		@{ Name = "HciWorkflowSupported" ; Expression = { $_.ExtensionData.Capability.HciWorkflowSupported } }, `
		@{ Name = "RootFolder" ; Expression = { Get-Folder -Id ( $_.ExtensionData.Content.RootFolder ) } }, `
		@{ Name = "Product" ; Expression = { $_.ExtensionData.Content.About.Name } }, `
		@{ Name = "FullName" ; Expression = { $_.ExtensionData.Content.About.FullName } }, `
		@{ Name = "Vendor" ; Expression = { $_.ExtensionData.Content.About.Vendor } }, `
		@{ Name = "LocaleVersion" ; Expression = { $_.ExtensionData.Content.About.LocaleVersion } }, `
		@{ Name = "LocaleBuild" ; Expression = { $_.ExtensionData.Content.About.LocaleBuild } }, `
		@{ Name = "ProductLineId" ; Expression = { $_.ExtensionData.Content.About.ProductLineId } }, `
		@{ Name = "ApiType" ; Expression = { $_.ExtensionData.Content.About.ApiType } }, `
		@{ Name = "ApiVersion" ; Expression = { $_.ExtensionData.Content.About.ApiVersion } }, `
		@{ Name = "LicenseProductName" ; Expression = { $_.ExtensionData.Content.About.LicenseProductName } }, `
		@{ Name = "LicenseProductVersion" ; Expression = { $_.ExtensionData.Content.About.LicenseProductVersion } } | `
	Export-Csv $vCenterExportFile -Append -NoTypeInformation
}
#endregion ~~< vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datacenter_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Datacenter Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting all Datacenter Info." -ForegroundColor Green
	$DatacenterExportFile = "$CaptureCsvFolder\$vCenter-DatacenterExport.csv"
	$i = 0
	$DatastoreNumber = 0

	foreach( $Datacenter in ( Get-View -ViewType Datacenter | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DatastoreNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Datacenter object $DatastoreNumber of $( ( Get-View -ViewType Datacenter ).Count ) -" $Datacenter.Name
		}
		$i++
		$DatacenterCsvValidationComplete.Forecolor = "Blue"
		$DatacenterCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType Datacenter ).Count )"
		$TabCapture.Controls.Add($DatacenterCsvValidationComplete)		
		$Datacenter | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } }, `
			@{ Name = "VmFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -type VM | Where-Object { $_.MoRef -eq $_.VmFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "HostFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -type HostAndCluster | Where-Object { $_.MoRef -eq $_.HostFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -type Datastore | Where-Object { $_.MoRef -eq $_.DatastoreFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "NetworkFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -type Network | Where-Object { $_.MoRef -eq $_.NetworkFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( Get-VMHost -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( Get-VMHost -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-VM -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( Get-VM -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Template" ; Expression = { [string]::Join(", ", ( Get-Template -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join(", ", ( Get-Template -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Folder" ; Expression = { [string]::Join(", ", ( Get-Folder -type Datacenter | Where-Object { $_.MoRef -eq $_.DatacenterFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "FolderId" ; Expression = { [string]::Join(", ", ( Get-Folder -type Datacenter ).Id ) } }, `
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join(", ", ( Get-DatastoreCluster -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreClusterId" ; Expression = { [string]::Join(", ", ( Get-DatastoreCluster -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join(", ", ( Get-Datastore -Id $_.Datastore | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreId" ; Expression = { [string]::Join(", ", ( Get-Datastore -Id $_.Datastore | Sort-Object Name ).Id ) } }, `
			@{ Name = "vSwitch" ; Expression = { [string]::Join(", ", ( Get-VirtualSwitch -Datacenter $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "vSwitchId" ; Expression = { [string]::Join(", ", ( Get-VirtualSwitch -Datacenter $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Network" ; Expression = { [string]::Join(", ", ( Get-VirtualPortGroup | Where-Object { $_.MoRef -eq $_.Network } | Sort-Object Name ) ) } }, `
			@{ Name = "NetworkId" ; Expression = { [string]::Join(", ", ( Get-VirtualPortGroup | Where-Object { $_.MoRef -eq $_.Network } | Sort-Object Name ).Id  ) } }, `
			@{ Name = "DefaultHardwareVersionKey" ; Expression = { [string]::Join(", ", ( $_.Configuration.DefaultHardwareVersionKey | Sort-Object Name ) ) } }, `
			@{ Name = "LinkedView" ; Expression = { [string]::Join(", ", ( $_.LinkedView | Sort-Object Name ) ) } }, `
			@{ Name = "Parent" ; Expression = { [string]::Join(", ", ( Get-Folder -type Datacenter | Where-Object { $_.MoRef -eq $_.Parent } | Sort-Object Name ) ) } }, `
			@{ Name = "OverallStatus" ; Expression = { [string]::Join(", ", ( $_.OverallStatus ) ) } }, `
			@{ Name = "ConfigStatus" ; Expression = { [string]::Join(", ", ( $_.ConfigStatus ) ) } }, `
			@{ Name = "ConfigIssue" ; Expression = { [string]::Join( ", ", ( $_.ConfigIssue ) ) } }, `
			@{ Name = "EffectiveRole" ; Expression = { [string]::Join( ", ", ( $_.EffectiveRole ) ) } }, `
			@{ Name = "AlarmActionsEnabled" ; Expression = { [string]::Join(", ", ( $_.AlarmActionsEnabled ) ) } }, `
			@{ Name = "Tag" ; Expression = { [string]::Join(", ", ( $_.Tag ) ) } }, `
			@{ Name = "Value" ; Expression = { [string]::Join(", ", ( $_.Value ) ) } }, `
			@{ Name = "AvailableField" ; Expression = { [string]::Join(", ", ( $_.AvailableField ) ) } }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join(", ", ( $_.MoRef ) ) } } | `
		Export-Csv $DatacenterExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Cluster_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Cluster Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting all Cluster Info." -ForegroundColor Green
	$ClusterExportFile = "$CaptureCsvFolder\$vCenter-ClusterExport.csv"
	$i = 0
	$ClusterNumber = 0
	
	foreach( $Cluster in ( Get-View -ViewType ClusterComputeResource | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$ClusterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Cluster object $ClusterNumber of $( ( Get-View -ViewType ClusterComputeResource ).Count ) -" $Cluster.Name 
		}
		$i++
		$ClusterCsvValidationComplete.Forecolor = "Blue"
		$ClusterCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType ClusterComputeResource ).Count )"
		$TabCapture.Controls.Add($ClusterCsvValidationComplete)
		
		$Cluster | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } }, `
			@{ Name = "Datacenter" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -Cluster $_.Name ) ) } }, `
			@{ Name = "DatacenterId" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -Cluster $_.Name ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join( ", ", ( Get-Vm -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join( ", ", ( Get-Vm -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Template" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Location $_.Name | Get-Template | Sort-Object Name ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Location $_.Name | Get-Template | Sort-Object Name ).Id ) } }, `
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -Datastore ( Get-Cluster $_.Name | Get-Datastore ) | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreClusterId" ; Expression = { [string]::Join( ", ", ( ( Get-DatastoreCluster -Datastore ( Get-Cluster $_.Name | Get-Datastore ) ) | Sort-Object Name ).Id ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( Get-Cluster $_.Name | Get-Datastore | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreId" ; Expression = { [string]::Join( ", ", ( Get-Cluster $_.Name | Get-Datastore | Sort-Object Name ).Id ) } }, `
			@{ Name = "ResourcePool" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "ResourcePoolId" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "DRSRule" ; Expression = { [string]::Join( ", ", ( Get-DrsRule -Cluster $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "DrsClusterGroup" ; Expression = { [string]::Join( ", ", ( Get-DrsClusterGroup -Cluster $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "DrsVMHostRule" ; Expression = { [string]::Join( ", ", ( Get-DrsVMHostRule -Cluster $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "HAEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster -Id $_.MoRef ).HAEnabled ) ) } }, `
			@{ Name = "HAAdmissionControlEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).HAAdmissionControlEnabled ) ) } }, `
			@{ Name = "AdmissionControlPolicyCpuFailoverResourcesPercent" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.AdmissionControlPolicy.CpuFailoverResourcesPercent ) ) } }, `
			@{ Name = "AdmissionControlPolicyMemoryFailoverResourcesPercent" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.MemoryFailoverResourcesPercent ) ) } }, `
			@{ Name = "AdmissionControlPolicyFailoverLevel" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.FailoverLevel ) ) } }, `
			@{ Name = "AdmissionControlPolicyAutoComputePercentages" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.AutoComputePercentages ) ) } }, `
			@{ Name = "AdmissionControlPolicyResourceReductionToToleratePercent" ; Expression = { [string]::Join( ", ", ( $_.ConfigurationEx.DasConfig.AdmissionControlPolicy.ResourceReductionToToleratePercent ) ) } }, `
			@{ Name = "DrsEnabled" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).DrsEnabled ) ) } }, `
			@{ Name = "DrsAutomationLevel" ; Expression = { [string]::Join( ", ", ( ( Get-Cluster  -Id $_.MoRef ).DrsAutomationLevel ) ) } }, `
			@{ Name = "VmMonitoring" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.VmMonitoring ) ) } }, `
			@{ Name = "HostMonitoring" ; Expression = { [string]::Join( ", ", ( $_.Configuration.DasConfig.HostMonitoring ) ) } }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } | `
		Export-Csv $ClusterExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Cluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VmHost_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VmHost Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Host Info." -ForegroundColor Green
	$VmHostExportFile = "$CaptureCsvFolder\$vCenter-VmHostExport.csv"
	$ServiceInstance = Get-View ServiceInstance
	$LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
	$LicenseAssignmentManager = Get-View $LicenseManager.LicenseAssignmentManager
	$i = 0
	$VmHostNumber = 0
	
	foreach( $VmHost in ( Get-View -ViewType HostSystem | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$VmHostNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Host object $VmHostNumber of $( ( Get-View -ViewType HostSystem ).Count ) -" $VmHost.Name
		}
		$i++
		$VMHostCsvValidationComplete.Forecolor = "Blue"
		$VMHostCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType HostSystem ).Count )"
		$TabCapture.Controls.Add($VMHostCsvValidationComplete)
		
		$VmHost | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join( ", ", (  $_.Name ) ) } }, `
            @{ Name = "Datacenter" ; Expression = { $Datacenter = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Datacenter -isnot [VMware.Vim.Datacenter] -and $Datacenter.Parent ) `
				{ $Datacenter = Get-View -Id $Datacenter.Parent -Property Name, Parent } `
				if ( $Datacenter -is [VMware.Vim.Datacenter] ) `
				{ $Datacenter.Name } } }, `
			@{ Name = "DatacenterId" ; Expression = { $Datacenter = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Datacenter -isnot [VMware.Vim.Datacenter] -and $Datacenter.Parent ) `
				{ $Datacenter = Get-View -Id $Datacenter.Parent -Property Name, Parent } `
				if ( $Datacenter -is [VMware.Vim.Datacenter] ) `
				{ $Datacenter.MoRef } } }, `
			@{ Name = "Cluster" ; Expression = { $Cluster = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent) `
				{ $Cluster = Get-View -Id $Cluster.Parent -Property Name, Parent }`
				if ( $Cluster -is [VMware.Vim.ClusterComputeResource] ) `
				{ $Cluster.Name } } }, `
			@{ Name = "ClusterId" ; Expression = { $Cluster = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent) `
				{ $Cluster = Get-View -Id $Cluster.Parent -Property Name, Parent }`
				if ( $Cluster -is [VMware.Vim.ClusterComputeResource] ) `
				{ $Cluster.MoRef } } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join( ", ", ( Get-VM -Id $_.Vm | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join( ", ", ( Get-VM -Id $_.Vm | Sort-Object Name ).Id ) } }, `
            @{ Name = "Template" ; Expression = { [string]::Join( ", ", ( Get-Template -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join( ", ", ( Get-Template -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( ( Get-Datastore -Id $_.Datastore | Sort-Object Name ) ) ) } }, `
            @{ Name = "DatastoreId" ; Expression = { [string]::Join( ", ", ( ( Get-Datastore -Id $_.Datastore | Sort-Object Name ).Id ) ) } }, `
			@{ Name = "vSwitch" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualSwitch -VMHost $_.Name | Sort-Object Name ) ) ) } }, `
            @{ Name = "vSwitchId" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualSwitch -VMHost $_.Name | Sort-Object Name ).Id ) ) } }, `
            @{ Name = "Version" ; Expression = { $_.Config.Product.Version } }, `
		    @{ Name = "Build" ; Expression = { $_.Config.Product.Build } }, `
		    @{ Name = "Manufacturer" ; Expression = { $_.Summary.Hardware.Vendor } }, `
		    @{ Name = "Model" ; Expression = { $_.Summary.Hardware.Model } }, `
		    @{ Name = "LicenseType" ; Expression = { $LicenseAssignmentManager.QueryAssignedLicenses($_.Config.Host.Value).AssignedLicense.Name  } }, `
			@{ Name = "BIOSVersion" ; Expression = { ( $_.Hardware.BiosInfo.BiosVersion ) } }, `
		    @{ Name = "BIOSReleaseDate" ; Expression = { ( $_.Hardware.BiosInfo.ReleaseDate ) } }, `
		    @{ Name = "ProcessorType" ; Expression = { $_.Summary.Hardware.CpuModel } }, `
		    @{ Name = "CpuMhz" ; Expression = { $_.Summary.Hardware.CpuMhz } }, `
		    @{ Name = "NumCpuPkgs" ; Expression = { $_.Summary.Hardware.NumCpuPkgs } }, `
		    @{ Name = "NumCpuCores" ; Expression = { $_.Summary.Hardware.NumCpuCores } }, `
		    @{ Name = "NumCpuThreads" ; Expression = { $_.Summary.Hardware.NumCpuThreads } }, `
		    @{ Name = "Memory" ; Expression = { [math]::Round([decimal]$_.Summary.Hardware.MemorySize / 1073741824) } }, `
		    @{ Name = "MaxEVCMode" ; Expression = { $_.Summary.MaxEVCModeKey } }, `
		    @{ Name = "NumNics" ; Expression = { $_.Summary.Hardware.NumNics } }, `
		    @{ Name = "ManagemetIP" ; Expression = { Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty IP } }, `
		    @{ Name = "ManagemetMacAddress" ; Expression = { Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Mac } }, `
		    @{ Name = "ManagemetVMKernel" ; Expression = { Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Name } }, `
		    @{ Name = "ManagemetSubnetMask" ; Expression = { Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.ManagementTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask } }, `
		    @{ Name = "vMotionIP" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } }, `
		    @{ Name = "vMotionMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } }, `
		    @{ Name = "vMotionVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } }, `
		    @{ Name = "vMotionSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VMotionEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } }, `
		    @{ Name = "FtIP" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } }, `
		    @{ Name = "FtMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } }, `
		    @{ Name = "FtVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } }, `
		    @{ Name = "FtSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.FaultToleranceLoggingEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } }, `
		    @{ Name = "VSANIP" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty IP ) ) } }, `
		    @{ Name = "VSANMacAddress" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Mac ) ) } }, `
		    @{ Name = "VSANVMKernel" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty Name ) ) } }, `
		    @{ Name = "VSANSubnetMask" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -VMHost $_.Name | Where-Object { $_.VsanTrafficEnabled -eq 'True' } | Select-Object -ExpandProperty SubnetMask ) ) } }, `
		    @{ Name = "NumHBAs" ; Expression = { $_.Summary.Hardware.NumHBAs } }, `
		    @{ Name = "iSCSIIP" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).IPv4 ) } }, `
		    @{ Name = "iSCSIMac" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).MACAddress ) } }, `
		    @{ Name = "iSCSIVMKernel" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).Vmknic ) } }, `
		    @{ Name = "iSCSISubnetMask" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).IPv4SubnetMask ) } }, `
		    @{ Name = "iSCSIAdapter" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).Adapter ) } }, `
		    @{ Name = "iSCSILinkUp" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).LinkUp ) } }, `
		    @{ Name = "iSCSIMTU" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).MTU ) } }, `
		    @{ Name = "iSCSINICDriver" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICDriver ) } }, `
		    @{ Name = "iSCSINICDriverVersion" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICDriverVersion ) } }, `
		    @{ Name = "iSCSINICFirmwareVersion" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).NICFirmwareVersion ) } }, `
		    @{ Name = "iSCSIPathStatus" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).PathStatus ) } }, `
		    @{ Name = "iSCSIVlanID" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).VlanID ) } }, `
		    @{ Name = "iSCSIVswitch" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).VSwitch ) } }, `
		    @{ Name = "iSCSICompliantStatus" ; Expression = { [string]::Join( ", ", ( ( Get-EsxCli -V2 -VMHost $_.Name ).Iscsi.NetworkPortal.List.Invoke() ).CompliantStatus ) } }, `
		    @{ Name = "IScsiName" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostHBA -type IScsi ).IScsiName ) } }, `
            @{ Name = "PortGroup" ; Expression = { `
				if ( $_.Network -like "DistributedVirtualPortgroup*" ) { [string]::Join( ", ", ( Get-VDPortGroup -Id $_.Network | Sort-Object Name ) ) } `
                elseif ( $_.Network -like "VmwareDistributedVirtualSwitch*" ) { [string]::Join( ", ", ( Get-VDSwitch -Id $_.Network | Sort-Object Name ) ) } `
                elseif ( $_.Network -like "Network*" ) { [string]::Join( ", ", ( Get-VirtualNetwork -Id $_.Network | Sort-Object Name ) ) } } }, `
			@{ Name = "PortGroupId" ; Expression = { `
				if ( $_.Network -like "DistributedVirtualPortgroup*" ) { [string]::Join( ", ", ( Get-VDPortGroup -Id $_.Network | Sort-Object Name ).Id ) } `
                elseif ( $_.Network -like "VmwareDistributedVirtualSwitch*" ) { [string]::Join( ", ", ( Get-VDSwitch -Id $_.Network | Sort-Object Name ).Id ) } `
                elseif ( $_.Network -like "Network*" ) { [string]::Join( ", ", ( Get-VirtualNetwork -Id $_.Network | Sort-Object Name ).Id ) } } }, `
			@{ Name = "CdpLldpInfo" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | `
				ForEach-Object { Get-View $_.Id } | `
					ForEach-Object { Get-View $_.ConfigManager.NetworkSystem} | `
						ForEach-Object { foreach ( $PhysicalNic in $_.NetworkInfo.Pnic ) `
						{ `
							$PnicsInfo = $_.QueryNetworkHint( $PhysicalNic.Device ) 
							foreach ( $PnicInfo in $PnicsInfo ) `
							{ `
								( $PnicInfo.ConnectedSwitchPort | `
								Select-Object `
									@{ Name = "VMNic" ; Expression = { $PhysicalNic.Device } }, `
									@{ Name = "DevId" ; Expression = { $_.DevId } }, `
									@{ Name = "PortId" ; Expression = { $_.PortId } }, `
									@{ Name = "MAC" ; Expression = { $PhysicalNic.Mac } } | Format-Table -HideTableHeaders | Out-String ).Trim() `
							} `
						} } `
				) ) `
			} }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", (  $_.MoRef ) ) } } | `
		Export-Csv $VmHostExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< VmHost_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Vm_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VM Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Machine Info." -ForegroundColor Green
	$VmExportFile = "$CaptureCsvFolder\$vCenter-VmExport.csv"
	$i = 0
	$VmNumber = 0
	
	foreach( $VM in ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $False } | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$VmNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on VM object $VmNumber of $( ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $False } ).Count) -" $VM.Name
		}
		$i++
		$VmCsvValidationComplete.Forecolor = "Blue"
		$VmCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $False } ).Count)"
		$TabCapture.Controls.Add($VmCsvValidationComplete)

		$VM | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VM ( Get-VM -Id $_.MoRef ) } }, `
			@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VM ( Get-VM -Id $_.MoRef ) ).Id } }, `
			@{ Name = "Cluster" ; Expression = { Get-Cluster -VM ( Get-VM -Id $_.MoRef ) } }, `
			@{ Name = "ClusterId" ; Expression = { ( Get-Cluster -VM ( Get-VM -Id $_.MoRef ) ).Id } }, `
			@{ Name = "VmHost" ; Expression = { Get-VmHost -Id ( $_.Runtime.Host ) } }, `
			@{ Name = "VmHostId" ; Expression = { ( Get-VmHost -Id ( $_.Runtime.Host ) ).Id } }, `
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreClusterId" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name ).Id ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( $_.Config.DatastoreUrl.Name | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreId" ; Expression = { [string]::Join( ", ", ( Get-Datastore ( $_.Config.DatastoreUrl.Name | Sort-Object Name ) ).Id ) } }, `
			@{ Name = "ResourcePool" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -VM ( Get-VM -Id $_.MoRef ) | Where-Object { $_ -notlike "Resources" } | Sort-Object Name ) ) } }, `
			@{ Name = "ResourcePoolId" ; Expression = { [string]::Join( ", ", ( Get-ResourcePool -VM ( Get-VM -Id $_.MoRef ) | Where-Object { $_ -notlike "Resources" } | Sort-Object Name ).Id ) } }, `
			@{ Name = "vSwitch" ; Expression = { [string]::Join( ", ", ( Get-VirtualSwitch -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name ) ) } }, `
			@{ Name = "vSwitchId" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualSwitch -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name ).Id ) ) } }, `
			@{ Name = "PortGroup" ; Expression = { [string]::Join( ", ", ( Get-VirtualPortGroup -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name ) ) } }, `
			@{ Name = "PortGroupId" ; Expression = { [string]::Join( ", ", ( Get-VirtualPortGroup -VM ( Get-VM -Id $_.MoRef ) | Sort-Object Name | `
				ForEach-Object `
				{ `
					if ($( $_.key -like "key-vim.host.PortGroup*" ) )
					{ `
						$_.Key
					}
					elseif ($( $_.key -like "dvportgroup-*" ) )
					{ `
						$_.Id
					}
				} ) ) `
			} }, `
			@{ Name = "OS" ; Expression = { [string]::Join( ", ", ( $_.Config.GuestFullName ) ) } }, `
			@{ Name = "Version" ; Expression = { [string]::Join( ", ", ( $_.Config.Version ) ) } }, `
			@{ Name = "VMToolsVersion" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersion ) ) } }, `
			@{ Name = "ToolsVersionStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersionStatus ) ) } }, `
			@{ Name = "ToolsStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsStatus ) ) } }, `
			@{ Name = "ToolsRunningStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsRunningStatus ) ) } }, `
			@{ Name = "Folder" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).Name ) ) } }, `
			@{ Name = "FolderId" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).MoRef ) ) } }, `
			@{ Name = "NumCPU" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCPU ) ) } }, `
			@{ Name = "CoresPerSocket" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCoresPerSocket ) ) } }, `
			@{ Name = "MemoryGB" ; Expression = { [string]::Join( ", ", ( [math]::Round([decimal] ( $_.Config.Hardware.MemoryMB / 1024 ), 0 ) ) ) } }, `
			@{ Name = "IP" ; Expression = { [string]::Join(", ", ( $_.Guest.IpAddress ) ) } }, `
			@{ Name = "MacAddress" ; Expression = { [string]::Join(", ", ( $_.Guest.Net.MacAddress ) ) } }, `
			@{ Name = "NumVirtualDisks" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumVirtualDisks ) ) } }, `
			@{ Name = "VmdkInfo" ; Expression = { [string]::Join(", ", ( Get-HardDisk -Vm $_.Name | Sort-Object Name | `
				ForEach-Object `
				{ `
					if ($($_.DiskType -like "Raw*"))
					{
						"VMDK = $( $_.FileName.split()[1].split('/')[1] ) - Hard Disk = $( $_.Name ) - Storage Format = $( $_.StorageFormat ) - Persistence = $( $_.Persistence ) - Disk Type = $( $_.DiskType ) - ScsiCanonicalName = $( $_.ScsiCanonicalName )" `
					}
					else
					{
						"VMDK = $( $_.FileName.split()[1].split('/')[1] ) - Hard Disk = $( $_.Name ) - Storage Format = $( $_.StorageFormat ) - Persistence = $( $_.Persistence ) - Disk Type = $( $_.DiskType )" `
					}
				} ) ) `
			} }, `
			@{ Name = "Volumes" ; Expression = { [string]::Join(", ", ( $_.Guest.Disk | Sort-Object DiskPath | `
				ForEach-Object `
				{ `
					"$( $_.DiskPath ) - $( [math]::Ceiling( $_.Capacity/1GB ) )GB Total - $( [math]::Ceiling( $_.FreeSpace/1GB ) )GB Free" `
				} ) ) `
			} }, `
			@{ Name = "ProvisionedSpaceGB" ; Expression = { ( [math]::Ceiling( ( Get-VM -Id $_.MoRef ).ProvisionedSpaceGB ) ) } }, `
			@{ Name = "NumEthernetCards" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumEthernetCards ) ) } }, `
			@{ Name = "CpuReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.CpuReservation ) ) } }, `
			@{ Name = "MemoryReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.MemoryReservation ) ) } }, `
			@{ Name = "CpuHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotAddEnabled ) ) } }, `
			@{ Name = "CpuHotRemoveEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotRemoveEnabled ) ) } }, `
			@{ Name = "MemoryHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.MemoryHotAddEnabled ) ) } }, `
			@{ Name = "SRM" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.ManagedBy.Type ) ) } }, `
			@{ Name = "Snapshot" ; Expression = { [string]::Join( ", ", ( Get-Snapshot -VM $_.Name -Id ( $_.Snapshot.CurrentSnapshot ) ) ) } }, `
			@{ Name = "RootSnapshot" ; Expression = { [string]::Join( ", ", ( ( Get-Snapshot -VM $_.Name -Id $_.RootSnapshot ).Name ) ) } }, `
			@{ Name = "SnapshotId" ; Expression = { [string]::Join( ", ", ( Get-Snapshot -VM $_.Name  | Sort-Object Name ).Id ) } }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( ( $_.MoRef ) ) ) } } | `
		Export-Csv $VmExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Template_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Template Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Template Info." -ForegroundColor Green
	$TemplateExportFile = "$CaptureCsvFolder\$vCenter-TemplateExport.csv"
	$i = 0
	$TemplateNumber = 0
	
	foreach( $Template in ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $True } | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$TemplateNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Template object $TemplateNumber of $( ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $True } ).Count ) -" $Template.Name
		}
		$i++
		$TemplateCsvValidationComplete.Forecolor = "Blue"
		$TemplateCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $True } ).Count )"
		$TabCapture.Controls.Add($TemplateCsvValidationComplete)

		$Template | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join( ", ", ( ( [uri]::UnescapeDataString( $_.Name ) ) ) ) } }, `
			@{ Name = "Datacenter" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -VMHost ( Get-VMHost -Id $_.Runtime.Host ) ) ) } }, `
			@{ Name = "DatacenterId" ; Expression = { [string]::Join( ", ", ( Get-Datacenter -VMHost ( Get-VMHost -Id $_.Runtime.Host ) ).Id ) } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VMHost ( Get-VMHost -Id $_.Runtime.Host ) ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VMHost ( Get-VMHost -Id $_.Runtime.Host ) ).Id ) } }, `
			@{ Name = "DatastoreCluster" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -Template $_.Name ) ) } }, `
			@{ Name = "DatastoreClusterId" ; Expression = { [string]::Join( ", ", ( Get-DatastoreCluster -Template $_.Name ).Id ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join( ", ", ( $_.Config.DatastoreUrl.Name ) ) } }, `
			@{ Name = "DatastoreId" ; Expression = { [string]::Join( ", ", ( Get-Datastore ( $_.Config.DatastoreUrl.Name ) ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Runtime.Host ) ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Runtime.Host ) ).Id ) } }, `
			@{ Name = "OS" ; Expression = { [string]::Join( ", ", ( $_.Config.GuestFullName ) ) } }, `
			@{ Name = "Version" ; Expression = { [string]::Join( ", ", ( $_.Config.Version ) ) } }, `
			@{ Name = "ToolsVersion" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersion ) ) } }, `
			@{ Name = "ToolsVersionStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsVersionStatus ) ) } }, `
			@{ Name = "ToolsStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsStatus ) ) } }, `
			@{ Name = "ToolsRunningStatus" ; Expression = { [string]::Join( ", ", ( $_.Guest.ToolsRunningStatus ) ) } }, `
			@{ Name = "Folder" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).Name ) ) } }, `
			@{ Name = "FolderId" ; Expression = { [string]::Join( ", ", ( ( Get-View -Id $_.Parent -Property Name).MoRef ) ) } }, `
			@{ Name = "NumCPU" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCPU ) ) } }, `
			@{ Name = "CoresPerSocket" ; Expression = { [string]::Join( ", ", ( $_.Config.Hardware.NumCoresPerSocket ) ) } }, `
			@{ Name = "MemoryGB" ; Expression = { [string]::Join( ", ", ( [math]::Round([decimal] ( $_.Config.Hardware.MemoryMB / 1024 ), 0 ) ) ) } }, `
			@{ Name = "MacAddress" ; Expression = { [string]::Join(", ", ( $_.Guest.Net.MacAddress ) ) } }, `
			@{ Name = "NumEthernetCards" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumEthernetCards ) ) } }, `
			@{ Name = "NumVirtualDisks" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.NumVirtualDisks ) ) } }, `
			@{ Name = "CpuReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.CpuReservation ) ) } }, `
			@{ Name = "MemoryReservation" ; Expression = { [string]::Join( ", ", ( $_.Summary.Config.MemoryReservation ) ) } }, `
			@{ Name = "CpuHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotAddEnabled ) ) } }, `
			@{ Name = "CpuHotRemoveEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.CpuHotRemoveEnabled ) ) } }, `
			@{ Name = "MemoryHotAddEnabled" ; Expression = { [string]::Join( ", ", ( $_.Config.MemoryHotAddEnabled ) ) } }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( ( $_.MoRef ) ) ) } } | `
		Export-Csv $TemplateExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Template_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function DatastoreCluster_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Datastore Cluster Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Datastore Cluster Info." -ForegroundColor Green
	$DatastoreClusterExportFile = "$CaptureCsvFolder\$vCenter-DatastoreClusterExport.csv"
	$i = 0
	$DatastoreClusterNumber = 0

	foreach( $DatastoreCluster in ( Get-View -ViewType StoragePod | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DatastoreClusterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Datastore Cluster object $DatastoreClusterNumber of $( ( Get-View -ViewType StoragePod ).Count ) -" $DatastoreCluster.Name
		}
		$i++
		$DatastoreClusterCsvValidationComplete.Forecolor = "Blue"
		$DatastoreClusterCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType StoragePod ).Count )"
		$TabCapture.Controls.Add($DatastoreClusterCsvValidationComplete)

		$DatastoreCluster | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { ( Get-DatastoreCluster -Id $_.MoRef | Get-Datastore ).Datacenter | Select-Object -Unique } }, `
			@{ Name = "DatacenterId" ; Expression = { ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Datastore ).Datacenter | Select-Object -Unique ).Id } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost ).Parent | Select-Object -Unique ) ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( ( ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost ).Parent | Select-Object -Unique ).Id ) ) ) ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost | Sort-Object Name ).Name ) ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-VmHost | Sort-Object Name ).Id ) ) ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Vm | Sort-Object Name ).Name ) ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Vm | Sort-Object Name ).Id ) ) ) } }, `
			@{ Name = "Template" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Template | Sort-Object Name ).Name ) ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Template | Sort-Object Name ).Id ) ) ) } }, `
			@{ Name = "Datastore" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Datastore ) | Select-Object -Unique ).Name ) ) } }, `
			@{ Name = "DatastoreId" ; Expression = { [string]::Join(", ", ( ( ( Get-DatastoreCluster -Id $_.MoRef | Get-Datastore ) | Select-Object -Unique ).Id ) ) } }, `
			@{ Name = "SdrsAutomationLevel" ; Expression = { $_.PodStorageDrsEntry.StorageDrsConfig.PodConfig.DefaultVmBehavior } }, `
			@{ Name = "IOLoadBalanceEnabled" ; Expression = { $_.PodStorageDrsEntry.StorageDrsConfig.PodConfig.IoLoadBalanceEnabled } }, `
			@{ Name = "CapacityGB" ; Expression = { [math]::Round( [decimal]$_.Summary.Capacity/1073741824, 0 ) } }, `
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | `
		Export-Csv $DatastoreClusterExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< DatastoreCluster_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datastore_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Datastore Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Datastore Info." -ForegroundColor Green
	$DatastoreExportFile = "$CaptureCsvFolder\$vCenter-DatastoreExport.csv"
	$i = 0
	$DatastoreNumber = 0	
	
	foreach( $Datastore in ( Get-View -ViewType Datastore | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$DatastoreNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Datastore object $DatastoreNumber of $( ( Get-View -ViewType Datastore ).Count ) -" $Datastore.Name
		}
		$i++
		$DatastoreCsvValidationComplete.Forecolor = "Blue"
		$DatastoreCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType Datastore ).Count )"
		$TabCapture.Controls.Add($DatastoreCsvValidationComplete)

		$Datastore | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { ( Get-Datastore -Id $_.MoRef ).Datacenter } }, `
			@{ Name = "DatacenterId" ; Expression = { ( ( Get-Datastore -Id $_.MoRef ).Datacenter ).Id } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster (Get-VmHost -Id $_.Host.Key).Parent.Name ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster (Get-VmHost -Id $_.Host.Key).Parent.Name ).Id ) } }, `
			@{ Name = "DatastoreCluster" ; Expression = { Get-DatastoreCluster -Datastore ( Get-Datastore -Id $_.MoRef ) } }, `
			@{ Name = "DatastoreClusterId" ; Expression = { ( Get-DatastoreCluster -Datastore ( Get-Datastore -Id $_.MoRef ) ).Id } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( Get-VmHost -Id $_.Host.Key | Sort-Object Name ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( Get-VmHost -Id $_.Host.Key | Sort-Object Name ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-Datastore $Datastore.Name | Get-Vm | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( Get-Datastore $Datastore.Name | Get-Vm | Sort-Object Name ).Id ) } }, `
			@{ Name = "Template" ; Expression = { [string]::Join(", ", ( ( ( Get-Datastore -Id $_.MoRef | Get-Template | Sort-Object Name ).Name ) ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join(", ", ( ( ( Get-Datastore -Id $_.MoRef | Get-Template | Sort-Object Name ).Id ) ) ) } }, `
			@{ Name = "Type" ; Expression = { $_.Info.Vmfs.Type } }, `
			@{ Name = "FileSystemVersion" ; Expression = { $_.Info.Vmfs.Version } }, `
			@{ Name = "DiskName" ; Expression = { $_.Info.VMFS.Extent.DiskName } }, `
			@{ Name = "DiskPath" ; Expression = { ($_.Summary.Url).Trim('ds://') } }, `
			@{ Name = "DiskUuid" ; Expression = { $_.Info.Vmfs.Uuid } }, `
			@{ Name = "StorageIOControlEnabled" ; Expression = { $_.IormConfiguration.Enabled } }, `
			@{ Name = "CapacityGB" ; Expression = { [math]::Round( [decimal] $_.Summary.Capacity / 1073741824, 0 ) } }, `
			@{ Name = "FreeSpaceGB" ; Expression = { [math]::Round( [decimal] $_.Summary.FreeSpace / 1073741824, 0 ) } }, `
			@{ Name = "CongestionThresholdMillisecond" ; Expression = { $_.IormConfiguration.CongestionThreshold } }, `
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | `
		Export-Csv $DatastoreExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Datastore_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VsSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VsSwitch_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Virtual Standard Switch Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Standard Switch Info." -ForegroundColor Green
	$VsSwitchExportFile = "$CaptureCsvFolder\$vCenter-VsSwitchExport.csv"
	$i = 0
	$VsSwitchNumber = 0

	foreach( $VsSwitch in ( Get-VirtualSwitch -Standard | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$VsSwitchNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Virtual Standard Switch object $VsSwitchNumber of $( ( Get-VirtualSwitch -Standard ).Count ) -" $VsSwitch.Name "on" $VsSwitch.VmHost
		}
		$i++
		$VsSwitchCsvValidationComplete.Forecolor = "Blue"
		$VsSwitchCsvValidationComplete.Text = "$i of $( ( Get-VirtualSwitch -Standard ).Count )"
		$TabCapture.Controls.Add($VsSwitchCsvValidationComplete)

		$VsSwitch | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $_.VmHost ) ) } }, `
			@{ Name = "DatacenterId" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $_.VmHost ).Id ) } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( $_.VmHost ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( $_.VmHost ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-VirtualSwitch -Standard -Name $_.Name -VMHost $_.VmHost | Get-VM ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( Get-VirtualSwitch -Standard -Name $_.Name -VMHost $_.VmHost | Get-VM ).Id ) } }, `
			@{ Name = "PortGroup" ; Expression = { [string]::Join(", ", ( Get-VirtualPortGroup -Standard -VirtualSwitch $_ ) ) } }, `
			@{ Name = "PortGroupId" ; Expression = { [string]::Join(", ", ( Get-VirtualPortGroup -Standard -VirtualSwitch $_ ).Key ) } }, `
			@{ Name = "Nic" ; Expression = { [string]::Join(", ", ( $_.Nic ) ) } }, `
			@{ Name = "NicId" ; Expression = { [string]::Join(", ", ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $_ -VMHost $_.VmHost -Name $_.Nic ).Id ) } }, 	
			@{ Name = "SpecNumPorts" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.NumPorts ) ) } }, `
			@{ Name = "SpecPolicySecurityAllowPromiscuous" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.Security.AllowPromiscuous ) ) } }, `
			@{ Name = "SpecPolicySecurityMacChanges" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.Security.MacChanges ) ) } }, `
			@{ Name = "SpecPolicySecurityForgedTransmits" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.Security.ForgedTransmits ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingPolicy" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.Policy ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingReversePolicy" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.ReversePolicy ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingNotifySwitches" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.NotifySwitches ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingRollingOrder" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.RollingOrder ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingNicOrderActiveNic" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic ) ) } }, `
			@{ Name = "SpecPolicyNicTeamingNicOrderStandbyNic" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.StandbyNic ) ) } }, `
			@{ Name = "NumPorts" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.NumPorts ) ) } }, `
			@{ Name = "NumPortsAvailable" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.NumPortsAvailable ) ) } }, `
			@{ Name = "Mtu" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Mtu ) ) } }, `
			@{ Name = "SpecBridgeBeacon" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ExtensionData.Spec.Bridge.Beacon ) ) } }, `
			@{ Name = "SpecBridgeLinkDiscoveryProtocolConfig" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.Bridge.LinkDiscoveryProtocolConfig | `
				ForEach-Object `
				{ `
					"$( $_.Protocol ) - $( $_.Operation )" `
				} ) ) `
			} }, `
			@{ Name = "MoRef" ; Expression = { $_.Id } } | `
		Export-Csv $VsSwitchExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< VsSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssPort_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VSS Port Group Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Standard Port Group Info." -ForegroundColor Green
	$VssPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VssPortGroupExport.csv"
	$i = 0
	$VssPortGroupNumber = 0
	
	foreach ( $VMHost in Get-VMHost ) `
	{ `
		foreach ( $VsSwitch in ( Get-VirtualSwitch -Standard -VMHost $VmHost ) ) `
		{ `
			foreach( $VssPortGroup in ( Get-VirtualPortGroup -Standard -VirtualSwitch $VsSwitch | Sort-Object Name ) ) `
			{ `
				if ( $debug -eq $true )`
				{ `
					$VssPortGroupNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Collecting info on Virtual Standard Port Group object $VssPortGroupNumber of $( ( Get-VirtualPortGroup -Standard ).Count ) -" $VssPortGroup.Name "on" $VsSwitch.Name "on" $VMHost.Name
				}
				$i++
				$VssPortGroupCsvValidationComplete.Forecolor = "Blue"
				$VssPortGroupCsvValidationComplete.Text = "$i of $( ( Get-VirtualPortGroup -Standard ).Count )"
				$TabCapture.Controls.Add($VssPortGroupCsvValidationComplete)

				$VssPortGroup | `
				Select-Object `
					@{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } }, `
					@{ Name = "Datacenter" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VMHost $VMHost.Name ) ) } }, `
					@{ Name = "DatacenterId" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VMHost $VMHost.Name ).Id ) } }, `
					@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -VMHost $VMHost.Name ) ) } }, `
					@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -VMHost $VMHost.Name ).Id ) } }, `
					@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( $VMHost.Name ) ) } }, `
					@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( $VMHost.Id ) ) } }, `
					@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( $_ | Get-VM | Sort-Object Name ) ) } }, `
					@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( $_ | Get-VM | Sort-Object Name ).Id ) } }, `
					@{ Name = "VsSwitch" ; Expression = { [string]::Join(", ", ( $VsSwitch.Name ) ) } }, `
					@{ Name = "VsSwitchId" ; Expression = { [string]::Join(", ", ( $VsSwitch.Id ) ) } }, `
					@{ Name = "VLanId" ; Expression = { [string]::Join(", ", ( $_.VLanId ) ) } }, `
					@{ Name = "Security_AllowPromiscuous" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.Security.AllowPromiscuous ) ) } }, `
					@{ Name = "Security_MacChanges" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.Security.MacChanges ) ) } }, `
					@{ Name = "Security_ForgedTransmits" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.Security.ForgedTransmits ) ) } }, `
					@{ Name = "NicTeaming_Policy" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.Policy ) ) } }, `
					@{ Name = "NicTeaming_ReversePolicy" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.ReversePolicy ) ) } }, `
					@{ Name = "NicTeaming_NotifySwitches" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.NotifySwitches ) ) } }, `
					@{ Name = "NicTeaming_RollingOrder" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.RollingOrder ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_CheckSpeed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.CheckSpeed ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_Speed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.Speed ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_CheckDuplex" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.CheckDuplex ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_FullDuplex" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.FullDuplex ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_CheckErrorPercent" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.CheckErrorPercent ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_Percentage" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.Percentage ) ) } }, `
					@{ Name = "NicTeaming_FailureCriteria_CheckBeacon" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.FailureCriteria.CheckBeacon ) ) } }, `
					@{ Name = "NicTeaming_NicOrder_ActiveNic" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.ActiveNic ) ) } }, `
					@{ Name = "NicTeaming_NicOrder_StandbyNic" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.StandbyNic ) ) } }, `
					@{ Name = "OffloadPolicy_CsumOffload" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.OffloadPolicy.CsumOffload ) ) } }, `
					@{ Name = "OffloadPolicy_TcpSegmentation" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.OffloadPolicy.TcpSegmentation ) ) } }, `
					@{ Name = "OffloadPolicy_ZeroCopyXmit" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.OffloadPolicy.ZeroCopyXmit ) ) } }, `
					@{ Name = "ShapingPolicy_Enabled" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.ShapingPolicy.Enabled ) ) } }, `
					@{ Name = "ShapingPolicy_AverageBandwidth" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.ShapingPolicy.AverageBandwidth ) ) } }, `
					@{ Name = "ShapingPolicy_PeakBandwidth" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.ShapingPolicy.PeakBandwidth ) ) } }, `
					@{ Name = "ShapingPolicy_BurstSize" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ComputedPolicy.ShapingPolicy.BurstSize ) ) } }, `
					@{ Name = "MoRef" ; Expression = { $_.Key } } | `
				Export-Csv $VssPortGroupExportFile -Append -NoTypeInformation
			}
		}
	}
}
#endregion ~~< VssPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssVmk_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VSS VMkernel Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Standard VMkernel Info." -ForegroundColor Green
	$VssVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VssVmkernelExport.csv"
	$i = 0
	$VssVmkernelNumber = 0
	
	foreach ( $VMHost in Get-VMHost ) `
	{ `
		foreach ( $VsSwitch in ( Get-VirtualSwitch -VMHost $VmHost -Standard ) ) `
		{ `
			foreach ( $VssPort in ( Get-VirtualPortGroup -Standard -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
				foreach ( $VMHostNetworkAdapterVMKernel in ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VsSwitch -PortGroup $VssPort | Sort-Object Name ) ) `
				{ `
					if ( $debug -eq $true )`
					{ `
						$VssVmkernelNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Collecting info on Virtual Standard VMkernel object $VssVmkernelNumber of $( ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch ( Get-VirtualSwitch -Standard ) ).Count ) -" $VMHostNetworkAdapterVMKernel.Name "on" $VsSwitch.Name
					}
					$i++
					$VssVmkernelCsvValidationComplete.Forecolor = "Blue"
					$VssVmkernelCsvValidationComplete.Text = "$i of $( ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch ( Get-VirtualSwitch -Standard ) ).Count )"
					$TabCapture.Controls.Add($VssVmkernelCsvValidationComplete)

					$VMHostNetworkAdapterVMKernel | `
					Select-Object `
						@{ Name = "Name" ; Expression = { $_.Name } }, `
						@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost $VMHost.Name } }, `
						@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VMHost $VMHost.Name ).Id } }, `
						@{ Name = "Cluster" ; Expression = { Get-Cluster -VMHost $VMHost.Name } }, `
						@{ Name = "ClusterId" ; Expression = { ( Get-Cluster -VMHost $VMHost.Name ).Id } }, `
						@{ Name = "VmHost" ; Expression = { $VMHost.Name } }, `
						@{ Name = "VmHostId" ; Expression = { $VMHost.Id } }, `
						@{ Name = "VSwitch" ; Expression = { $VsSwitch.Name } }, `
						@{ Name = "VSwitchId" ; Expression = { $VsSwitch.Id } }, `
						@{ Name = "PortGroupName" ; Expression = { $_.PortGroupName } }, `
						@{ Name = "PortGroupId" ; Expression = { $_.Id } }, `
						@{ Name = "VMotionEnabled" ; Expression = { $_.VMotionEnabled } }, `
						@{ Name = "FaultToleranceLoggingEnabled" ; Expression = { $_.FaultToleranceLoggingEnabled } }, `
						@{ Name = "ManagementTrafficEnabled" ; Expression = { $_.ManagementTrafficEnabled } }, `
						@{ Name = "IP" ; Expression = { $_.IP } }, `
						@{ Name = "Mac" ; Expression = { $_.Mac } }, `
						@{ Name = "SubnetMask" ; Expression = { $_.SubnetMask } }, `
						@{ Name = "DhcpEnabled" ; Expression = { $_.DhcpEnabled } }, `
						@{ Name = "IPv6" ; Expression = { $_.IPv6 } }, `
						@{ Name = "AutomaticIPv6" ; Expression = { $_.AutomaticIPv6 } }, `
						@{ Name = "IPv6ThroughDhcp" ; Expression = { $_.IPv6ThroughDhcp } }, `
						@{ Name = "IPv6Enabled" ; Expression = { $_.IPv6Enabled } }, `
						@{ Name = "VsanTrafficEnabled" ; Expression = { $_.VsanTrafficEnabled } }, `
						@{ Name = "Mtu" ; Expression = { $_.Mtu } }, `
						@{ Name = "SpecTsoEnabled" ; Expression = { $_.ExtensionData.Spec.TsoEnabled } }, `
						@{ Name = "SpecNetStackInstanceKey" ; Expression = { $_.ExtensionData.Spec.NetStackInstanceKey } }, `
						@{ Name = "SpecOpaqueNetwork" ; Expression = { $_.ExtensionData.Spec.OpaqueNetwork } }, `
						@{ Name = "SpecExternalId" ; Expression = { $_.ExtensionData.Spec.ExternalId } }, `
						@{ Name = "SpecPinnedPnic" ; Expression = { $_.ExtensionData.Spec.PinnedPnic } }, `
						@{ Name = "SpecIpRouteSpec" ; Expression = { $_.ExtensionData.Spec.IpRouteSpec } } | `
					Export-Csv $VssVmkernelExportFile -Append -NoTypeInformation
				}
			}
		}
	}
}
#endregion ~~< VssVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VssPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VssPnic_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VSS pNIC Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Standard Physical NIC Info." -ForegroundColor Green
	$VssPnicExportFile = "$CaptureCsvFolder\$vCenter-VssPnicExport.csv"
	$i = 0
	$VssPnicNumber = 0
	
	foreach ( $VMHost in Get-VMHost ) `
	{ `
		foreach ( $VsSwitch in ( Get-VirtualSwitch -Standard -VMHost $VmHost ) ) `
		{ `
			foreach ( $VMHostNetworkAdapterUplink in ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VsSwitch -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
				if ( $debug -eq $true )`
				{ `
					$VssPnicNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Collecting info on Virtual Standard Uplink object $VssPnicNumber of $( ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch ( Get-VirtualSwitch -Standard ) ).Count ) -" $VMHostNetworkAdapterUplink.Name "on" $VsSwitch.Name
				}
				$i++
				$VssPnicCsvValidationComplete.Forecolor = "Blue"
				$VssPnicCsvValidationComplete.Text = "$i of $( ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch ( Get-VirtualSwitch -Standard ) ).Count )"
				$TabCapture.Controls.Add($VssPnicCsvValidationComplete)

				$VMHostNetworkAdapterUplink | `
				Select-Object `
					@{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } }, `
					@{ Name = "Datacenter" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $VmHost ) ) } }, `
					@{ Name = "DatacenterId" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $VmHost ).Id ) } }, `
					@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ) ) } }, `
					@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ).Id ) } }, `
					@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( $_.VmHost ) ) } }, `
					@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( $_.VmHost ).Id ) } }, `
					@{ Name = "VsSwitch" ; Expression = { [string]::Join(", ", ( $VsSwitch.Name ) ) } }, `
					@{ Name = "VsSwitchId" ; Expression = { [string]::Join(", ", ( $VsSwitch.Id ) ) } }, `
					@{ Name = "Mac" ; Expression = { [string]::Join(", ", ( $_.Mac ) ) } }, `
					@{ Name = "DhcpEnabled" ; Expression = { [string]::Join(", ", ( $_.DhcpEnabled ) ) } }, `
					@{ Name = "IP" ; Expression = { [string]::Join(", ", ( $_.IP ) ) } }, `
					@{ Name = "SubnetMask" ; Expression = { [string]::Join(", ", ( $_.SubnetMask ) ) } }, `
					@{ Name = "BitRatePerSec" ; Expression = { [string]::Join(", ", ( $_.BitRatePerSec ) ) } }, `
					@{ Name = "FullDuplex" ; Expression = { [string]::Join(", ", ( $_.FullDuplex ) ) } }, `
					@{ Name = "PciId" ; Expression = { [string]::Join(", ", ( $_.PciId ) ) } }, `
					@{ Name = "WakeOnLanSupported" ; Expression = { [string]::Join(", ", ( $_.WakeOnLanSupported ) ) } }, `
					@{ Name = "Driver" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Driver ) ) } }, `
					@{ Name = "LinkSpeed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.LinkSpeed | `
						ForEach-Object { "Speed in Mb = $( $_.SpeedMb ) / Duplex = $( $_.Duplex )" } ) ) } }, `
					@{ Name = "SpecEnableEnhancedNetworkingStack" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.EnableEnhancedNetworkingStack ) ) } }, `
					@{ Name = "FcoeConfigurationPriorityClass" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.PriorityClass ) ) } }, `
					@{ Name = "FcoeConfigurationSourceMac" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.SourceMac ) ) } }, `
					@{ Name = "FcoeConfigurationVlanRange" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.VlanRange | `
						ForEach-Object { "VLAN Low = $( $_.VlanLow ) / VLAN High = $( $_.VlanHigh )" } ) ) } }, `
					@{ Name = "FcoeConfigurationCapabilities" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.Capabilities | `
						ForEach-Object { "Priority Class = $( $_.PriorityClass ) / Source MAC Address = $( $_.SourceMacAddress ) / VLAN Range = $( $_.VlanRange )" } ) ) } }, `
					@{ Name = "FcoeConfigurationFcoeActive" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.FcoeActive ) ) } }, `
					@{ Name = "VmDirectPathGen2Supported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.VmDirectPathGen2Supported ) ) } }, `
					@{ Name = "VmDirectPathGen2SupportedMode" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.VmDirectPathGen2SupportedMode ) ) } }, `
					@{ Name = "ResourcePoolSchedulerAllowed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ResourcePoolSchedulerAllowed ) ) } }, `
					@{ Name = "ResourcePoolSchedulerDisallowedReason" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ResourcePoolSchedulerDisallowedReason ) ) } }, `
					@{ Name = "AutoNegotiateSupported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.AutoNegotiateSupported ) ) } }, `
					@{ Name = "EnhancedNetworkingStackSupported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.EnhancedNetworkingStackSupported ) ) } }, `
					@{ Name = "MoRef" ; Expression = { [string]::Join(", ", ( $_.Id ) ) } } | `
				Export-Csv $VssPnicExportFile -Append -NoTypeInformation
			}
		}
	}
}
#endregion ~~< VssPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdSwitch_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Virtual Distributed Switch Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Distributed Switch Info." -ForegroundColor Green
	$VdSwitchExportFile = "$CaptureCsvFolder\$vCenter-VdSwitchExport.csv"
	$i = 0
	$VdSwitchNumber = 0
	
	foreach( $DistributedVirtualSwitch in ( Get-View -ViewType DistributedVirtualSwitch ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$VdSwitchNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Virtual Distributed Switch object $VdSwitchNumber of $( ( Get-View -ViewType DistributedVirtualSwitch ).Count ) -" $DistributedVirtualSwitch.Name
		}
		$i++
		$VdSwitchCsvValidationComplete.Forecolor = "Blue"
		$VdSwitchCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType DistributedVirtualSwitch ).Count )"
		$TabCapture.Controls.Add($VdSwitchCsvValidationComplete)

		$DistributedVirtualSwitch | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) ) } }, `
			@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) ) ).Id } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ) ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join( ", ", ( Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ) ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Summary.HostMember ) | Sort-Object Name ).Id ) } }, `
			@{ Name = "PortgroupName" ; Expression = { [string]::Join( ", ", ( $_.Summary.PortgroupName | Sort-Object ) ) } }, `
			@{ Name = "PortgroupId" ; Expression = { [string]::Join( ", ", ( ( Get-VirtualPortGroup -VirtualSwitch $_.Name -Name $_.Summary.PortGroupName | Sort-Object Name).Id ) ) } }, `
			@{ Name = "Nic" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "NicId" ; Expression = { [string]::Join( ", ", ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "NumHosts" ; Expression = { $_.Summary.NumHosts } }, `
			@{ Name = "NumPorts" ; Expression = { $_.Summary.NumPorts } }, `
			@{ Name = "Vendor" ; Expression = { $_.Summary.ProductInfo.Vendor } }, `
			@{ Name = "Version" ; Expression = { $_.Summary.ProductInfo.Version } }, `
			@{ Name = "ConfigVspanSession" ; Expression = { [string]::Join(", ", ( $_.Config.VspanSession ) ) } }, `
			@{ Name = "ConfigPvlanConfig" ; Expression = { [string]::Join(", ", ( $_.Config.PvlanConfig ) ) } }, `
			@{ Name = "ConfigMaxMtu" ; Expression = { [string]::Join(", ", ( $_.Config.MaxMtu ) ) } }, `
			@{ Name = "ConfigLinkDiscoveryProtocolConfig" ; Expression = { [string]::Join(", ", ( $_.Config.LinkDiscoveryProtocolConfig | `
				ForEach-Object `
				{ `
					"$( $_.Protocol ) - $( $_.Operation )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigIpfixConfigCollectorIpAddress" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.CollectorIpAddress ) ) } },  
			@{ Name = "ConfigIpfixConfigCollectorPort" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.CollectorPort ) ) } }, `
			@{ Name = "ConfigIpfixConfigObservationDomainId" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.ObservationDomainId ) ) } }, `
			@{ Name = "ConfigIpfixConfigActiveFlowTimeout" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.ActiveFlowTimeout ) ) } }, `
			@{ Name = "ConfigIpfixConfigIdleFlowTimeout" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.IdleFlowTimeout ) ) } }, `
			@{ Name = "ConfigIpfixConfigSamplingRate" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.SamplingRate ) ) } }, `
			@{ Name = "ConfigIpfixConfigInternalFlowsOnly" ; Expression = { [string]::Join(", ", ( $_.Config.IpfixConfig.InternalFlowsOnly ) ) } }, `
			@{ Name = "ConfigLacpGroupConfig" ; Expression = { [string]::Join(", ", ( $_.Config.LacpGroupConfig ) ) } }, `
			@{ Name = "ConfigLacpApiVersion" ; Expression = { [string]::Join(", ", ( $_.Config.LacpApiVersion ) ) } }, `
			@{ Name = "ConfigMulticastFilteringMode" ; Expression = { [string]::Join(", ", ( $_.Config.MulticastFilteringMode ) ) } }, `
			@{ Name = "ConfigNumStandalonePorts" ; Expression = { [string]::Join(", ", ( $_.Config.NumStandalonePorts ) ) } }, `
			@{ Name = "ConfigNumPorts" ; Expression = { [string]::Join(", ", ( $_.Config.NumPorts ) ) } }, `
			@{ Name = "ConfigMaxPorts" ; Expression = { [string]::Join(", ", ( $_.Config.MaxPorts ) ) } }, `
			@{ Name = "ConfigNumUplinkPorts" ; Expression = { ($_.Config.UplinkPortPolicy.UplinkPortName).Count } }, `
			@{ Name = "ConfigUplinkPortName" ; Expression = { [string]::Join( ", ", ( $_.Config.UplinkPortPolicy.UplinkPortName | Sort-Object Name ) ) } }, `
			@{ Name = "ConfigUplinkPortgroup"; Expression = { [string]::Join( ", ", ( Get-VirtualPortGroup -Id $_.Config.UplinkPortgroup | Sort-Object Name ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigVlan" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.Vlan | `
				ForEach-Object `
				{ `
					"VLAN Id = $( $_.VlanId ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigQosTag" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.QosTag | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyPolicy" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyReversePolicy" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.ReversePolicy | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyNotifySwitches" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.NotifySwitches | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyRollingOrder" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.RollingOrder | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckSpeed" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.CheckSpeed | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaSpeed" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.Speed | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckDuplex" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.CheckDuplex | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaFullDuplex" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.FullDuplex | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckErrorPercent" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.CheckErrorPercent | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaPercentage" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.Percentage | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckBeacon" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.CheckBeacon | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.FailureCriteria.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigSecurityPolicyAllowPromiscuous" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.SecurityPolicy.AllowPromiscuous | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigSecurityPolicyMacChanges" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.SecurityPolicy.MacChanges | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigSecurityPolicyForgedTransmits" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.SecurityPolicy.ForgedTransmits | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigSecurityPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.SecurityPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigUplinkTeamingPolicyUplinkPortOrder" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder | `
				ForEach-Object `
				{ `
					"Active Uplink Port = $( $_.ActiveUplinkPort ) / Standby Uplink Port = $( $_.StandbyUplinkPort ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigIpfixEnabled" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.IpfixEnabled | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigTxUplink" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.TxUplink | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigLacpPolicyEnable" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.LacpPolicy.Enable | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigLacpPolicyMode" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.LacpPolicy.Mode | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigLacpPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.LacpPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyAllowPromiscuous" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.AllowPromiscuous ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacChanges" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacChanges ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyForgedTransmits" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.ForgedTransmits ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyEnabled" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacLearningPolicy.Enabled ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyAllowUnicastFlooding" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacLearningPolicy.AllowUnicastFlooding ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimit" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacLearningPolicy.Limit ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimitPolicy" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacLearningPolicy.LimitPolicy ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.MacLearningPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigMacManagementPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.MacManagementPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigBlocked" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.Blocked | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigVmDirectPathGen2Allowed" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.VmDirectPathGen2Allowed | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigInShapingPolicyEnabled" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.InShapingPolicy.Enabled | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigInShapingPolicyAverageBandwidth" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.InShapingPolicy.AverageBandwidth | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigInShapingPolicyPeakBandwidth" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.InShapingPolicy.PeakBandwidth | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigInShapingPolicyBurstSize" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.InShapingPolicy.BurstSize | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigInShapingPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.InShapingPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigOutShapingPolicyEnabled" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.OutShapingPolicy.Enabled | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigOutShapingPolicyAverageBandwidth" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.OutShapingPolicy.AverageBandwidth | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigOutShapingPolicyPeakBandwidth" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.OutShapingPolicy.PeakBandwidth | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigOutShapingPolicyBurstSize" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.OutShapingPolicy.BurstSize | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigOutShapingPolicyInherited" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.OutShapingPolicy.Inherited ) ) } }, `
			@{ Name = "ConfigDefaultPortConfigVendorSpecificConfig" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.VendorSpecificConfig | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigNetworkResourcePoolKey" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.NetworkResourcePoolKey | `
				ForEach-Object `
				{ `
					"Value = $( $_.Value ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigDefaultPortConfigFilterPolicy" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.FilterPolicy | `
				ForEach-Object `
				{ `
					"FilterConfig = $( $_.FilterConfig ) / Inherited = $( $_.Inherited )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigPolicyAutoPreInstallAllowed" ; Expression = { [string]::Join(", ", ( $_.Config.Policy.AutoPreInstallAllowed ) ) } }, `
			@{ Name = "ConfigPolicyAutoUpgradeAllowed" ; Expression = { [string]::Join(", ", ( $_.Config.Policy.AutoUpgradeAllowed ) ) } }, `
			@{ Name = "ConfigPolicyPartialUpgradeAllowed" ; Expression = { [string]::Join(", ", ( $_.Config.Policy.PartialUpgradeAllowed ) ) } }, `
			@{ Name = "ConfigSwitchIpAddress" ; Expression = { [string]::Join(", ", ( $_.Config.SwitchIpAddress ) ) } }, `
			@{ Name = "ConfigCreateTime" ; Expression = { [string]::Join(", ", ( $_.Config.CreateTime ) ) } }, `
			@{ Name = "ConfigNetworkResourceManagementEnabled" ; Expression = { [string]::Join(", ", ( $_.Config.NetworkResourceManagementEnabled ) ) } }, `
			@{ Name = "ConfigDefaultProxySwitchMaxNumPorts" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultProxySwitchMaxNumPorts ) ) } }, `
			@{ Name = "ConfigHealthCheckConfig" ; Expression = { [string]::Join(", ", ( $_.Config.HealthCheckConfig | `
				ForEach-Object `
				{ `
					"Enabled = $( $_.Enable ) / Interval = $( $_.Interval )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigInfrastructureTrafficResourceConfig" ; Expression = { [string]::Join(", ", ( $_.Config.InfrastructureTrafficResourceConfig | `
				ForEach-Object `
				{ `
					"$( $_.Key ) - $( $_.Description ) / Limit = $( $_.AllocationInfo.Limit ) / Shares = $( $_.AllocationInfo.Shares.Shares ) / Level = $( $_.AllocationInfo.Shares.Level ) / Reservation = $( $_.AllocationInfo.Reservation )" `
				} ) ) `
			} }, `
			@{ Name = "ConfigNetResourcePoolTrafficResourceConfig" ; Expression = { [string]::Join(", ", ( $_.Config.NetResourcePoolTrafficResourceConfig ) ) } }, `
			@{ Name = "ConfigNetworkResourceControlVersion" ; Expression = { [string]::Join(", ", ( $_.Config.NetworkResourceControlVersion ) ) } }, `
			@{ Name = "ConfigVmVnicNetworkResourcePool" ; Expression = { [string]::Join(", ", ( $_.Config.VmVnicNetworkResourcePool ) ) } }, `
			@{ Name = "ConfigPnicCapacityRatioForReservation" ; Expression = { [string]::Join(", ", ( $_.Config.PnicCapacityRatioForReservation ) ) } }, `
			@{ Name = "RuntimeHostMemberRuntime" ; Expression = { [string]::Join(", ", ( $_.Runtime.HostMemberRuntime | `
				ForEach-Object `
				{ `
					"Host = $( Get-VMHost -Id ($_.Host ) ) / Status = $( $_.Status ) / Status Detail = $( $_.StatusDetail ) / Health Check Result = $( $_.HealthCheckResult )" `
				} ) ) `
			} }, `
			@{ Name = "OverallStatus" ; Expression = { [string]::Join(", ", ( $_.OverallStatus ) ) } }, `
			@{ Name = "ConfigStatus" ; Expression = { [string]::Join(", ", ( $_.ConfigStatus ) ) } }, `
			@{ Name = "AlarmActionsEnabled" ; Expression = { [string]::Join(", ", ( $_.AlarmActionsEnabled ) ) } }, `
			@{ Name = "Mtu" ; Expression = { $_.Config.MaxMtu } }, `
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | `
		Export-Csv $VdSwitchExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< VdSwitch_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsPort_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VDS Port Group Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Distributed Switch Port Group Info." -ForegroundColor Green
	$VdsPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VdsPortGroupExport.csv"
	$i = 0
	$VdsPortGroupNumber = 0
	
	foreach( $DistributedVirtualPortgroup in ( Get-View -ViewType DistributedVirtualPortgroup ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$VdsPortGroupNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Virtual Distributed Port Group object $VdsPortGroupNumber of $( ( Get-View -ViewType DistributedVirtualPortgroup ).Count ) -" $DistributedVirtualPortgroup.Name "on" ( Get-VdSwitch  -Id $DistributedVirtualPortgroup.Config.DistributedVirtualSwitch )
		}
		$i++
		$VdsPortGroupCsvValidationComplete.Forecolor = "Blue"
		$VdsPortGroupCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType DistributedVirtualPortgroup ).Count )"
		$TabCapture.Controls.Add($VdsPortGroupCsvValidationComplete)

		$DistributedVirtualPortgroup | `
		Sort-Object Name | `
		Where-Object { $_.Name -notlike "*DVUplinks*" } | `
		Select-Object `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Host ) ) } }, `
			@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VMHost ( Get-VmHost -Id ( $_.Host ) ) ).Id } }, `
			@{ Name = "Cluster" ; Expression = { Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Host ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { ( Get-Cluster -VMHost ( Get-VmHost -Id ( $_.Host ) ) ).Id } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Host ) | Sort-Object Name  ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join( ", ", ( Get-VmHost -Id ( $_.Host ) | Sort-Object Name  ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join( ", ", ( Get-Vm -Id ( $_.VM ) | Sort-Object Name  ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join( ", ", ( Get-Vm -Id ( $_.VM ) | Sort-Object Name  ).Id ) } }, `
			@{ Name = "VdSwitch" ; Expression = { ( Get-VdSwitch  -Id $_.Config.DistributedVirtualSwitch ) } }, `
			@{ Name = "VdSwitchId" ; Expression = { ( Get-VdSwitch  -Id $_.Config.DistributedVirtualSwitch ).Id } }, `
			@{ Name = "VlanConfiguration" ; Expression = { "VLAN "+ $_.Config.DefaultPortConfig.Vlan.VlanId } }, `
			@{ Name = "NumPorts" ; Expression = { $_.Config.NumPorts } }, `
			@{ Name = "ActiveUplinkPort" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.ActiveUplinkPort)) } }, `
			@{ Name = "StandbyUplinkPort" ; Expression = { [string]::Join(", ", ( $_.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.StandbyUplinkPort)) } }, `
			@{ Name = "Policy" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value } }, `
			@{ Name = "ReversePolicy" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.ReversePolicy.Value } }, `
			@{ Name = "NotifySwitches" ; Expression = { $_.Config.DefaultPortConfig.UplinkTeamingPolicy.NotifySwitches.Value } }, `
			@{ Name = "PortBinding" ; Expression = { ( Get-VDPortgroup  -Id $_.MoRef ).PortBinding } }, `
			@{ Name = "MoRef" ; Expression = { $_.MoRef } } | `
		Export-Csv $VdsPortGroupExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< VdsPort_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsVmk_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VDS VMkernel Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Distributed Switch VMkernel Info." -ForegroundColor Green
	$VdsVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VdsVmkernelExport.csv"
	$i = 0
	$VdsVmkernelNumber = 0
	
	foreach ( $VmHost in Get-VmHost ) `
	{ `
		foreach ( $VdSwitch in ( Get-VdSwitch -VMHost $VmHost ) ) `
		{ `
			foreach ( $VdsVmkernel in ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VdSwitch -VMHost $VmHost | Sort-Object -Property Name -Unique ) ) `
			{ `
				if ( $debug -eq $true )`
				{ `
					$VdsVmkernelNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Collecting info on Virtual Distributed VMkernel object $VdsVmkernelNumber of $( ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch ( Get-VdSwitch ) ).Count ) -" $VdsVmkernel.Name "on" $VMHost.Name
				}
				$i++
				$VdsVmkernelCsvValidationComplete.Forecolor = "Blue"
				$VdsVmkernelCsvValidationComplete.Text = "$i of $( ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch ( Get-VdSwitch ) ).Count )"
				$TabCapture.Controls.Add($VdsVmkernelCsvValidationComplete)

				$VdsVmkernel | `
				Select-Object `
					@{ Name = "Name" ; Expression = { $_.Name } }, `
					@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VMHost $VMHost.Name } }, `
					@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VMHost $VMHost.Name ).Id } }, `
					@{ Name = "Cluster" ; Expression = { Get-Cluster -VMHost $VMHost.Name } }, `
					@{ Name = "ClusterId" ; Expression = { ( Get-Cluster -VMHost $VMHost.Name ).Id } }, `
					@{ Name = "VmHost" ; Expression = { $VMHost.Name } }, `
					@{ Name = "VmHostId" ; Expression = { ( $VMHost.Id ) } }, `
					@{ Name = "VSwitch" ; Expression = { $VdSwitch.Name } }, `
					@{ Name = "VSwitchId" ; Expression = { $VdSwitch.Id } }, `
					@{ Name = "PortGroupName" ; Expression = { $_.PortGroupName } }, `
					@{ Name = "PortGroupId" ; Expression = { $_.Id } }, `
					@{ Name = "DhcpEnabled" ; Expression = { $_.DhcpEnabled } }, `
					@{ Name = "IP" ; Expression = { $_.IP } }, `
					@{ Name = "Mac" ; Expression = { $_.Mac } }, `
					@{ Name = "ManagementTrafficEnabled" ; Expression = { $_.ManagementTrafficEnabled } }, `
					@{ Name = "VMotionEnabled" ; Expression = { $_.VMotionEnabled } }, `
					@{ Name = "FaultToleranceLoggingEnabled" ; Expression = { $_.FaultToleranceLoggingEnabled } }, `
					@{ Name = "VsanTrafficEnabled" ; Expression = { $_.VsanTrafficEnabled } }, `
					@{ Name = "Mtu" ; Expression = { $_.Mtu } } | `
				Export-Csv $VdsVmkernelExportFile -Append -NoTypeInformation
			}
		}
	}
}
#endregion ~~< VdsVmk_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VdsPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VdsPnic_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export VDS pNIC Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Virtual Distributed Switch Physical NIC Info." -ForegroundColor Green
	$VdsPnicExportFile = "$CaptureCsvFolder\$vCenter-VdsPnicExport.csv"
	$i = 0
	$VdsPnicNumber = 0
	
	foreach ( $VmHost in Get-VmHost ) `
	{ `
		foreach ( $VdSwitch in ( Get-VdSwitch -VMHost $VmHost ) ) `
		{ `
			foreach ( $VMHostNetworkAdapterUplink in ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VdSwitch -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
				if ( $debug -eq $true )`
				{ `
					$VdsPnicNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Collecting info on Virtual Distributed Uplink $VdsPnicNumber of $( ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch ( Get-VdSwitch ) ).Count ) -" ( Get-VDPort -VdSwitch $VdSwitch -Uplink -ConnectedOnly | Sort-Object -Property ConnectedEntity | Where-Object { $_.ProxyHost -like $VmHost -and $_.ConnectedEntity -like $VMHostNetworkAdapterUplink.Name } ).Name "on" $VmHost
				}
				$i++
				$VdsPnicCsvValidationComplete.Forecolor = "Blue"
				$VdsPnicCsvValidationComplete.Text = "$i of $( ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch ( Get-VdSwitch ) ).Count )"
				$TabCapture.Controls.Add($VdsPnicCsvValidationComplete)

				$VMHostNetworkAdapterUplink | `
				Select-Object `
					@{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } }, `
					@{ Name = "Datacenter" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $VmHost ) ) } }, `
					@{ Name = "DatacenterId" ; Expression = { [string]::Join(", ", ( Get-Datacenter -VmHost $VmHost ).Id ) } }, `
					@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ) ) } }, `
					@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -VmHost $_.VmHost ).Id ) } }, `
					@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( $_.VmHost ) ) } }, `
					@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( $_.VmHost ).Id ) } }, `
					@{ Name = "VdSwitch" ; Expression = { [string]::Join(", ", ( $VdSwitch.Name ) ) } }, `
					@{ Name = "VdSwitchId" ; Expression = { [string]::Join(", ", ( $VdSwitch.Id ) ) } }, `
					@{ Name = "Mac" ; Expression = { [string]::Join(", ", ( $_.Mac ) ) } }, `
					@{ Name = "DhcpEnabled" ; Expression = { [string]::Join(", ", ( $_.DhcpEnabled ) ) } }, `
					@{ Name = "IP" ; Expression = { [string]::Join(", ", ( $_.IP ) ) } }, `
					@{ Name = "SubnetMask" ; Expression = { [string]::Join(", ", ( $_.SubnetMask ) ) } }, `
					@{ Name = "Portgroup" ; Expression = { ( Get-VDPort -VdSwitch $VdSwitch -Uplink -ConnectedOnly | Sort-Object -Property ConnectedEntity | Where-Object { $_.ProxyHost -like $VmHost -and $_.ConnectedEntity -like $VMHostNetworkAdapterUplink.Name } ).Portgroup } }, `
					@{ Name = "ConnectedEntity" ; Expression = { ( Get-VDPort -VdSwitch $VdSwitch -Uplink -ConnectedOnly | Sort-Object -Property ConnectedEntity | Where-Object { $_.ProxyHost -like $VmHost -and $_.ConnectedEntity -like $VMHostNetworkAdapterUplink.Name } ).Name } }, `
					@{ Name = "VlanConfiguration" ; Expression = { ( Get-VDPort -VdSwitch $VdSwitch -Uplink -ConnectedOnly | Sort-Object -Property ConnectedEntity | Where-Object { $_.ProxyHost -like $VmHost -and $_.ConnectedEntity -like $VMHostNetworkAdapterUplink.Name } ).VlanConfiguration } }, `
					@{ Name = "BitRatePerSec" ; Expression = { [string]::Join(", ", ( $_.BitRatePerSec ) ) } }, `
					@{ Name = "FullDuplex" ; Expression = { [string]::Join(", ", ( $_.FullDuplex ) ) } }, `
					@{ Name = "PciId" ; Expression = { [string]::Join(", ", ( $_.PciId ) ) } }, `
					@{ Name = "WakeOnLanSupported" ; Expression = { [string]::Join(", ", ( $_.WakeOnLanSupported ) ) } }, `
					@{ Name = "Driver" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Driver ) ) } }, `
					@{ Name = "LinkSpeed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.LinkSpeed | `
						ForEach-Object { "Speed in Mb = $( $_.SpeedMb ) / Duplex = $( $_.Duplex )" } ) ) } }, `
					@{ Name = "SpecEnableEnhancedNetworkingStack" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.Spec.EnableEnhancedNetworkingStack ) ) } }, `
					@{ Name = "FcoeConfigurationPriorityClass" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.PriorityClass ) ) } }, `
					@{ Name = "FcoeConfigurationSourceMac" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.SourceMac ) ) } }, `
					@{ Name = "FcoeConfigurationVlanRange" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.VlanRange | `
						ForEach-Object { "VLAN Low = $( $_.VlanLow ) / VLAN High = $( $_.VlanHigh )" } ) ) } }, `
					@{ Name = "FcoeConfigurationCapabilities" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.Capabilities | `
						ForEach-Object { "Priority Class = $( $_.PriorityClass ) / Source MAC Address = $( $_.SourceMacAddress ) / VLAN Range = $( $_.VlanRange )" } ) ) } }, `
					@{ Name = "FcoeConfigurationFcoeActive" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.FcoeConfiguration.FcoeActive ) ) } }, `
					@{ Name = "VmDirectPathGen2Supported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.VmDirectPathGen2Supported ) ) } }, `
					@{ Name = "VmDirectPathGen2SupportedMode" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.VmDirectPathGen2SupportedMode ) ) } }, `
					@{ Name = "ResourcePoolSchedulerAllowed" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ResourcePoolSchedulerAllowed ) ) } }, `
					@{ Name = "ResourcePoolSchedulerDisallowedReason" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.ResourcePoolSchedulerDisallowedReason ) ) } }, `
					@{ Name = "AutoNegotiateSupported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.AutoNegotiateSupported ) ) } }, `
					@{ Name = "EnhancedNetworkingStackSupported" ; Expression = { [string]::Join(", ", ( $_.ExtensionData.EnhancedNetworkingStackSupported ) ) } }, `
					@{ Name = "MoRef" ; Expression = { [string]::Join(", ", ( $_.Id ) ) } } | `
				Export-Csv $VdsPnicExportFile -Append -NoTypeInformation
			}
		}
	}
}
#endregion ~~< VdsPnic_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Folder_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Folder_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Folder Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Folder Info." -ForegroundColor Green
	$FolderExportFile = "$CaptureCsvFolder\$vCenter-FolderExport.csv"
	$i = 0
	$FolderNumber = 0
	
	foreach ( $Datacenter in Get-Datacenter ) `
	{ `
		foreach ( $Folder in ( Get-Datacenter $Datacenter | Get-Folder | Get-View | Sort-Object ) ) `
		{ `
			if ( $debug -eq $true )`
			{ `
				$FolderNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Collecting info on Folder object $FolderNumber of $( ( Get-Datacenter | Get-Folder | Get-View ).Count ) -" $Folder.Name
			}
			$i++
			$FolderCsvValidationComplete.Forecolor = "Blue"
			$FolderCsvValidationComplete.Text = "$i of $( ( Get-Datacenter | Get-Folder | Get-View ).Count )"
			$TabCapture.Controls.Add($FolderCsvValidationComplete)

			$Folder | `
			Sort-Object Name | `
			Select-Object `
				@{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter $Datacenter } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter $Datacenter ).Id } }, `
				@{ Name = "ChildType" ; Expression = { [string]::Join( ", ", ( $_.ChildType | Sort-Object ) ) } }, `
				@{ Name = "ChildEntity" ; Expression = { `
					if ( $_.ChildEntity -like "Datacenter*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datacenter -Id $_.ChildEntity | Sort-Object ) ) `
					} `
					elseif ( $_.ChildEntity -like "ClusterComputeResource*" ) `
					{ `
						[string]::Join( ", ", ( Get-Cluster -Id $_.ChildEntity | Sort-Object ) ) `
					} 
					elseif ( $_.ChildEntity -like "DistributedVirtualPortgroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDPortGroup -Id $_.ChildEntity | Sort-Object ) ) `
					}  
					elseif ( $_.ChildEntity -like "VmwareDistributedVirtualSwitch*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDSwitch -Id $_.ChildEntity | Sort-Object ) ) `
					}  
					elseif ( $_.ChildEntity -like "Network*" ) `
					{ `
						[string]::Join( ", ", ( Get-VirtualSwitch -Id $_.ChildEntity | Sort-Object ) ) `
					} 
					elseif ( $_.ChildEntity -like "Datastore*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datastore -Id $_.ChildEntity | Sort-Object ) ) `
					} 
					elseif ( $_.ChildEntity -like "StoragePod*" ) `
					{ `
						[string]::Join( ", ", ( Get-DatastoreCluster -Id $_.ChildEntity | Sort-Object ) ) `
					} 
					elseif ( $_.ChildEntity -like "Folder*" ) `
					{ `
						[string]::Join( ", ", ( Get-Folder -Id $_.ChildEntity | Sort-Object ) ) `
					} 
					elseif ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-VM -Id $_.ChildEntity | Sort-Object ) ) `
					} `
					elseif ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-Template -Id $_.ChildEntity | Sort-Object ) ) `
					} `
				} }, `
				@{ Name = "Cluster" ; Expression = { `
					if ( $_.ChildEntity -like "ClusterComputeResource*" ) `
					{ `
						[string]::Join( ", ", ( Get-Cluster -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "ClusterId" ; Expression = {  `
					if ( $_.ChildEntity -like "ClusterComputeResource*" ) `
					{ `
						[string]::Join( ", ", ( Get-Cluster -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "Vm" ; Expression = { `
					if ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-VM -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "VmId" ; Expression = { `
					if ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-VM -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "Template" ; Expression = { `
					if ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-Template -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "TemplateId" ; Expression = { `
					if ( $_.ChildEntity -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-Template -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "Datastore" ; Expression = { `
					if ( $_.ChildEntity -like "Datastore*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datastore -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "DatastoreId" ; Expression = { `
					if ( $_.ChildEntity -like "Datastore*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datastore -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "DatastoreCluster" ; Expression = { `
					if ( $_.ChildEntity -like "StoragePod*" ) `
					{ `
						[string]::Join( ", ", ( Get-DatastoreCluster -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "DatastoreClusterId" ; Expression = { `
					if ( $_.ChildEntity -like "StoragePod*" ) `
					{ `
						[string]::Join( ", ", ( Get-DatastoreCluster -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "VsSwitch" ; Expression = { `
					if ( $_.ChildEntity -like "Network*" ) `
					{ `
						[string]::Join( ", ", ( Get-VirtualSwitch -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "VsSwitchId" ; Expression = { `
					if ( $_.ChildEntity -like "Network*" ) `
					{ `
						[string]::Join( ", ", ( Get-VirtualSwitch -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "VdSwitch" ; Expression = { `
					if ( $_.ChildEntity -like "VmwareDistributedVirtualSwitch*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDSwitch -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "VdSwitchId" ; Expression = { `
					if ( $_.ChildEntity -like "VmwareDistributedVirtualSwitch*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDSwitch -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "VdsPortgroup" ; Expression = { `
					if ( $_.ChildEntity -like "DistributedVirtualPortgroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDPortGroup -Id $_.ChildEntity | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "VdsPortgroupId" ; Expression = {  `
					if ( $_.ChildEntity -like "DistributedVirtualPortgroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDPortGroup -Id $_.ChildEntity | Sort-Object ).Id ) `
					}
				} }, `
				@{ Name = "LinkedView" ; Expression = { [string]::Join( ", ", ( $_.LinkedView ) ) } }, `
				@{ Name = "Parent" ; Expression = { `
					if ( $_.Parent -like "Datacenter*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datacenter -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "ClusterComputeResource*" ) `
					{ `
						[string]::Join( ", ", ( Get-Cluster -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "DistributedVirtualPortgroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDPortGroup -Id $_.Parent ) ) `
					}  
					elseif ( $_.Parent -like "VmwareDistributedVirtualSwitch*" ) `
					{ `
						[string]::Join( ", ", ( Get-VDSwitch -Id $_.Parent ) ) `
					}  
					elseif ( $_.Parent -like "Network*" ) `
					{ `
						[string]::Join( ", ", ( Get-VirtualSwitch -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "Datastore*" ) `
					{ `
						[string]::Join( ", ", ( Get-Datastore -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "StoragePod*" ) `
					{ `
						[string]::Join( ", ", ( Get-DatastoreCluster -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "Folder*" ) `
					{ `
						[string]::Join( ", ", ( Get-Folder -Id $_.Parent ) ) `
					} 
					elseif ( $_.Parent -like "VirtualMachine*" ) `
					{ `
						[string]::Join( ", ", ( Get-VM -Id $_.Parent ) ) `
					} `
				} },`
				@{ Name = "ParentId" ; Expression = { $_.Parent } }, `
				@{ Name = "CustomValue" ; Expression = { [string]::Join( ", ", ( $_.CustomValue ) ) } }, `
				@{ Name = "OverallStatus" ; Expression = { [string]::Join( ", ", ( $_.OverallStatus ) ) } }, `
				@{ Name = "ConfigStatus" ; Expression = { [string]::Join( ", ", ( $_.ConfigStatus ) ) } }, `
				@{ Name = "ConfigIssue" ; Expression = { [string]::Join( ", ", ( $_.ConfigIssue ) ) } }, `
				@{ Name = "EffectiveRole" ; Expression = { [string]::Join( ", ", ( $_.EffectiveRole ) ) } }, `
				@{ Name = "DisabledMethod" ; Expression = { [string]::Join( ", ", ( $_.DisabledMethod ) ) } }, `
				@{ Name = "AlarmActionsEnabled" ; Expression = { [string]::Join( ", ", ( $_.AlarmActionsEnabled ) ) } }, `
				@{ Name = "Tag" ; Expression = { [string]::Join( ", ", ( $_.Tag.Key ) ) } }, `
				@{ Name = "Value" ; Expression = { [string]::Join( ", ", ( $_.Value ) ) } }, `
				@{ Name = "AvailableField" ; Expression = { [string]::Join( ", ", ( $_.AvailableField ) ) } }, `
				@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } | `
			Export-Csv $FolderExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Folder_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Rdm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Rdm_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export RDM Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Raw Device Mapping (RDM) Info." -ForegroundColor Green
	$RdmExportFile = "$CaptureCsvFolder\$vCenter-RdmExport.csv"
	$i = 0
	$RdmNumber = 0
	
	foreach( $RDM in ( Get-VM | Get-HardDisk | Where-Object { $_.DiskType -like "Raw*" } | Sort-Object Parent ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$RdmNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on RDM $RdmNumber of $( ( Get-VM | Get-HardDisk | Where-Object { $_.DiskType -like "Raw*" } | Sort-Object Parent ).Count ) -" $RDM.ScsiCanonicalName
		}
		$i++
		$RdmCsvValidationComplete.Forecolor = "Blue"
		$RdmCsvValidationComplete.Text = "$i of $( ( Get-VM | Get-HardDisk | Where-Object { $_.DiskType -like "Raw*" } | Sort-Object Parent ).Count )"
		$TabCapture.Controls.Add($RdmCsvValidationComplete)

		$RDM | `
		Select-Object `
			@{ Name = "ScsiCanonicalName" ; Expression = { $_.ScsiCanonicalName } }, `
			@{ Name = "Datacenter" ; Expression = { Get-Datacenter -VM $_.Parent } }, `
			@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -VM $_.Parent ).Id } }, `
			@{ Name = "Cluster" ; Expression = { Get-Cluster -VM $_.Parent } }, `
			@{ Name = "ClusterId" ; Expression = { ( Get-Cluster -VM $_.Parent ).Id } }, `
			@{ Name = "Vm" ; Expression = { $_.Parent } }, `
			@{ Name = "VmId" ; Expression = { ( $_.Parent ).Id } }, `
			@{ Name = "Label" ; Expression = { $_.Name } }, `
			@{ Name = "CapacityGB" ; Expression = { [math]::Round([decimal]$_.CapacityGB, 2) } }, `
			@{ Name = "DiskType" ; Expression = { $_.DiskType } }, `
			@{ Name = "Persistence" ; Expression = { $_.Persistence } }, `
			@{ Name = "CompatibilityMode" ; Expression = { $_.ExtensionData.Backing.CompatibilityMode } }, `
			@{ Name = "DeviceName" ; Expression = { $_.ExtensionData.Backing.DeviceName } }, `
			@{ Name = "Sharing" ; Expression = { $_.ExtensionData.Backing.Sharing } } | `
		Export-Csv $RdmExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Rdm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Drs_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_Rule_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export DRS Rule Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting DRS Rule Info." -ForegroundColor Green
	$DrsRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsRuleExport.csv"
	$i = 0
	$DrsRuleNumber = 0
	
	foreach ( $Cluster in Get-Cluster ) `
	{ `
		foreach ( $DrsRule in ( Get-Cluster $Cluster | Get-DrsRule | Sort-Object Name) ) `
		{ `
			if ( $debug -eq $true )`
			{ `
				$DrsRuleNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Collecting info on DRS Rule $DrsRuleNumber of $( ( Get-DrsRule -Cluster ( Get-Cluster ) ).Count ) -" $DrsRule.Name
			}
			$i++
			$DrsRuleCsvValidationComplete.Forecolor = "Blue"
			$DrsRuleCsvValidationComplete.Text = "$i of $( ( Get-DrsRule -Cluster ( Get-Cluster ) ).Count )"
			$TabCapture.Controls.Add($DrsRuleCsvValidationComplete)

			$DrsRule | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter -Cluster $Cluster.Name } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -Cluster $Cluster.Name ).Id } }, `
				@{ Name = "Cluster" ; Expression = { $_.Cluster } }, `
				@{ Name = "ClusterId" ; Expression = { ( $_.Cluster ).Id } }, `
				@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-VM -Id $_.VMIDs | Sort-Object Name ) ) } }, `
				@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( $_.VMIDs | Sort-Object Name ) ) } }, `
				@{ Name = "Type" ; Expression = { $_.Type } }, `
				@{ Name = "Enabled" ; Expression = { $_.Enabled } }, `
				@{ Name = "Mandatory" ; Expression = { $_.Mandatory } } | `
			Export-Csv $DrsRuleExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Drs_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Drs_Cluster_Group_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_Cluster_Group_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export DRS Cluster Group Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting DRS Cluster Group Info." -ForegroundColor Green
	$DrsClusterGroupExportFile = "$CaptureCsvFolder\$vCenter-DrsClusterGroupExport.csv"
	$i = 0
	$DrsClusterGroupNumber = 0
	
	foreach ( $Cluster in Get-Cluster ) `
	{ `
		foreach ( $DrsClusterGroup in ( Get-DrsClusterGroup -Cluster $Cluster | Sort-Object Name ) ) `
		{ `
			if ( $debug -eq $true )`
			{ `
				$DrsClusterGroupNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Collecting info on DRS Cluster Group $DrsClusterGroupNumber of $( ( Get-DrsClusterGroup ).Count ) -" $DrsClusterGroup.Name
			}
			$i++
			$DrsClusterGroupCsvValidationComplete.Forecolor = "Blue"
			$DrsClusterGroupCsvValidationComplete.Text = "$i of $( ( Get-DrsClusterGroup -VMHost ( Get-VMHost ) ).Count )"
			$TabCapture.Controls.Add($DrsClusterGroupCsvValidationComplete)

			$DrsClusterGroup | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter -Cluster $_.Cluster } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -Cluster $_.Cluster ).Id } }, `
				@{ Name = "Cluster" ; Expression = { $_.Cluster } }, `
				@{ Name = "ClusterId" ; Expression = { ( $_.Cluster ).Id } }, `
				@{ Name = "GroupType" ; Expression = { $_.GroupType } }, `
				@{ Name = "DrsVMHostRule" ; Expression = { `
					if ( $_.GroupType -like "VMHostGroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-DrsVMHostRule -VMHostGroup $_.Name | Sort-Object ) ) `
					} `
					elseif ( $_.GroupType -like "VMGroup*" ) `
					{ `
						[string]::Join( ", ", ( Get-DrsVMHostRule -VMGroup $_.Name | Sort-Object ) ) `
					}
				} }, `
				@{ Name = "Member" ; Expression = { [string]::Join(", ", ( $_.Member | Sort-Object ) ) } }, `
				@{ Name = "MemberId" ; Expression = { [string]::Join(", ", ( $_.Member | Sort-Object ).Id ) } } | `
			Export-Csv $DrsClusterGroupExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Drs_Cluster_Group_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Drs_VmHost_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_VmHost_Rule_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export DRS VMHost Rule Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting DRS VMHost Rule Info." -ForegroundColor Green
	$DrsVmHostRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsVmHostRuleExport.csv"
	$i = 0
	$DrsVmHostRuleNumber = 0
	
	foreach ( $Cluster in Get-Cluster ) `
	{ `
		foreach ( $DrsVmHostRule in ( Get-Cluster $Cluster | Get-DrsVmHostRule | Sort-Object Name ) ) `
		{ `
			if ( $debug -eq $true )`
			{ `
				$DrsVmHostRuleNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Collecting info on DRS Host Rule $DrsVmHostRuleNumber of $( ( Get-DrsVmHostRule ).Count ) -" $DrsVmHostRule.Name
			}
			$i++
			$DrsVmHostRuleCsvValidationComplete.Forecolor = "Blue"
			$DrsVmHostRuleCsvValidationComplete.Text = "$i of $( ( Get-DrsVmHostRule ).Count )"
			$TabCapture.Controls.Add($DrsVmHostRuleCsvValidationComplete)

			$DrsVmHostRule | `
			Sort-Object Name | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter -Cluster $Cluster.Name } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -Cluster $Cluster.Name ).Id } }, `
				@{ Name = "Cluster" ; Expression = { $_.Cluster } }, `
				@{ Name = "ClusterId" ; Expression = { ( $_.Cluster ).Id } }, `
				@{ Name = "Enabled" ; Expression = { $_.Enabled } }, `
				@{ Name = "Type" ; Expression = { $_.Type } }, `
				@{ Name = "VMGroup" ; Expression = { $_.VMGroup } }, `
				@{ Name = "VMGroupMember" ; Expression = { [string]::Join(", ", ( $_.VMGroup.Member | Sort-Object Name ) ) } }, `
				@{ Name = "VMGroupMemberId" ; Expression = { [string]::Join(", ", ( $_.VMGroup.Member | Sort-Object Name ).Id ) } }, `
				@{ Name = "VMHostGroup" ; Expression = { $_.VMHostGroup } }, `
				@{ Name = "VMHostGroupMember" ; Expression = { [string]::Join(", ", ( $_.VMHostGroup.Member | Sort-Object Name ) ) } }, `
				@{ Name = "VMHostGroupMemberId" ; Expression = { [string]::Join(", ", ( $_.VMHostGroup.Member | Sort-Object Name ).Id ) } }, `
				@{ Name = "AffineHostGroupName" ; Expression = { $_.ExtensionData.AffineHostGroupName } }, `
				@{ Name = "AntiAffineHostGroupName" ; Expression = { $_.ExtensionData.AntiAffineHostGroupName } } | `
			Export-Csv $DrsVmHostRuleExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Drs_VmHost_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Resource_Pool_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Resource_Pool_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Resource Pool Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Resource Pool Info." -ForegroundColor Green
	$ResourcePoolExportFile = "$CaptureCsvFolder\$vCenter-ResourcePoolExport.csv"
	$i = 0
	$ResourcePoolNumber = 0
	
	foreach( $ResourcePool in ( Get-View -ViewType ResourcePool | Sort-Object Name ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$ResourcePoolNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Resource Pool $ResourcePoolNumber of $( ( Get-View -ViewType ResourcePool ).Count ) -" $ResourcePool.Name
		}
		$i++
		$ResourcePoolCsvValidationComplete.Forecolor = "Blue"
		$ResourcePoolCsvValidationComplete.Text = "$i of $( ( Get-View -ViewType ResourcePool ).Count )"
		$TabCapture.Controls.Add($ResourcePoolCsvValidationComplete)

		$ResourcePool | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join( ", ", ( $_.Name ) ) } }, `
			@{ Name = "Cluster" ; Expression = { $Cluster = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent) `
				{ $Cluster = Get-View -Id $Cluster.Parent -Property Name, Parent }`
				if ( $Cluster -is [VMware.Vim.ClusterComputeResource] ) `
				{ $Cluster.Name } } }, `
			@{ Name = "ClusterId" ; Expression = { $Cluster = Get-View -Id $_.Parent -Property Name, Parent
				while ( $Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent) `
				{ $Cluster = Get-View -Id $Cluster.Parent -Property Name, Parent }`
				if ( $Cluster -is [VMware.Vim.ClusterComputeResource] ) `
				{ $Cluster.MoRef } } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join( ", ", ( Get-Vm -Id $_.Vm | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join( ", ", ( Get-Vm -Id $_.Vm | Sort-Object Name ).Id ) } }, `
			@{ Name = "CpuSharesLevel" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuSharesLevel ) ) } }, `
			@{ Name = "NumCpuShares" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).NumCpuShares ) ) } }, `
			@{ Name = "CpuReservationMHz" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuReservationMHz ) ) } }, `
			@{ Name = "CpuExpandableReservation" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuExpandableReservation ) ) } }, `
			@{ Name = "CpuLimitMHz" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).CpuLimitMHz ) ) } }, `
			@{ Name = "MemSharesLevel" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemSharesLevel ) ) } }, `
			@{ Name = "NumMemShares" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).NumMemShares ) ) } }, `
			@{ Name = "MemReservationGB" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemReservationGB ) ) } }, `
			@{ Name = "MemExpandableReservation" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemExpandableReservation ) ) } }, `
			@{ Name = "MemLimitGB" ; Expression = { [string]::Join( ", ", ( ( Get-ResourcePool -Id $_.MoRef ).MemLimitGB ) ) } }, `
			@{ Name = "Parent" ; Expression = { [string]::Join( ", ", ( $_.Parent ) ) } }, `
			@{ Name = "MoRef" ; Expression = { [string]::Join( ", ", ( $_.MoRef ) ) } } | `
		Export-Csv $ResourcePoolExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Resource_Pool_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Snapshot_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Snapshot Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Snapshot Info." -ForegroundColor Green
	$SnapshotExportFile = "$CaptureCsvFolder\$vCenter-SnapshotExport.csv"
	$i = 0
	$SnapshotNumber = 0
	
	foreach( $Snapshot in ( Get-VM | Get-Snapshot | Sort-Object  VM, Created ) ) `
	{ `
		if ( $debug -eq $true )`
		{ `
			$SnapshotNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Collecting info on Snapshot $SnapshotNumber of $( ( Get-VM | Get-Snapshot ).Count ) -" $Snapshot.Name
		}
		$i++
		$SnapshotCsvValidationComplete.Forecolor = "Blue"
		$SnapshotCsvValidationComplete.Text = "$i of $( ( Get-VM | Get-Snapshot ).Count )"
		$TabCapture.Controls.Add($SnapshotCsvValidationComplete)

		$Snapshot | `
		Select-Object `
			@{ Name = "VM" ; Expression = { $_.VM } }, `
			@{ Name = "VMId" ; Expression = { ( $_.VM ).Id } }, `
			@{ Name = "Name" ; Expression = { $_.Name } }, `
			@{ Name = "Created" ; Expression = { $_.Created } }, `
			@{ Name = "Id" ; Expression = { $_.Id } }, `
			@{ Name = "Children" ; Expression = { $_.Children } }, `
			@{ Name = "ParentSnapshot" ; Expression = { $_.ParentSnapshot } }, `
			@{ Name = "ParentSnapshotId" ; Expression = { $_.ParentSnapshotId } }, `
			@{ Name = "IsCurrent" ; Expression = { $_.IsCurrent } } | `
		Export-Csv $SnapshotExportFile -Append -NoTypeInformation
	}
}
#endregion ~~< Snapshot_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Linked_vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Linked_vCenter_Export
{
	if ( $logcapture -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Export Linked vCenter Info selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Collecting Linked vCenter Info." -ForegroundColor Green
	$LinkedvCenterExportFile = "$CaptureCsvFolder\$vCenter-LinkedvCenterExport.csv"
	Disconnect-ViServer * -Confirm:$false
	$global:vCenter = $VcenterTextBox.Text
	$User = $UserNameTextBox.Text
	Connect-VIServer $Vcenter -user $User -password $PasswordTextBox.Text -AllLinked
	$i = 0
	$LinkedVcenterNumber = 0
	
	if ( ( $global:DefaultVIServers ).Count -gt "1" ) `
	{ `
		foreach ( $LinkedvCenter in ( $global:DefaultVIServers | Where-Object { $_.Name -ne "$vCenter" } ) ) `
		{ `
			if ( $debug -eq $true )`
			{ `
				$LinkedVcenterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Collecting info on Linked vCenter object $LinkedVcenterNumber of $( ( $Global:DefaultVIServers ).Count ) - " $LinkedvCenter.Name
			}
			$i++
			$LinkedvCenterCsvValidationComplete.Forecolor = "Blue"
			$LinkedvCenterCsvValidationComplete.Text = "$i of $( ( $Global:DefaultVIServers ).Count )"
			$TabCapture.Controls.Add($LinkedvCenterCsvValidationComplete)

			$LinkedvCenter | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Version" ; Expression = { $_.Version } }, `
				@{ Name = "Build" ; Expression = { $_.Build } }, `
				@{ Name = "OsType" ; Expression = { $_.ExtensionData.Content.About.OsType } }, `
				@{ Name = "vCenter" ; Expression = { $vCenter } } | `
			Export-Csv $LinkedvCenterExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Linked_vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Export Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Shapefile Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Shapefile_Select
{
	if ( $logdraw -eq $true ) `
	{ `
		if ( $ShapesfileSelectionRadioButton1.Checked -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] vDiagram Default Shapes file selected." -ForegroundColor Magenta
		}
		else `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] VMware VVD Shapes file selected." -ForegroundColor Magenta
		}
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	if ( $debug -eq $true )`
	{ `
		if ( $ShapesfileSelectionRadioButton1.Checked -eq $true ) `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Setting vDiagram Default Shapes as shape file." -ForegroundColor Magenta
		}
		else `
		{ `
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Setting VMware VVD Shapes as shape file." -ForegroundColor Magenta
		}
	}
	if ( $ShapesfileSelectionRadioButton1.Checked -eq $true ) `
	{ `
		$global:shpFile = "\vDiagram_" + $MyVer + ".vssx"
	}
	else `
	{ `
		$global:shpFile = "\vDiagram_" + $MyVer + "_VVD" + ".vssx"
	}
}
#endregion ~~< Shapefile Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

#region ~~< Add-VisioObjectVssPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Add-VisioObjectVssPG($mastObj, $item)
{
	$shpObj = $pagObj.Drop($mastObj, $x, $y)
	$shpObj.Text = $item.name
	return $shpObj
}
#endregion ~~< Add-VisioObjectVssPG >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
	$VCObject.Cells("Prop.Name").Formula = '"' + $vCenterImport.Name + '"'
	$VCObject.Cells("Prop.Version").Formula = '"' + $vCenterImport.Version + '"'
	$VCObject.Cells("Prop.Build").Formula = '"' + $vCenterImport.Build + '"'
	$VCObject.Cells("Prop.OsType").Formula = '"' + $vCenterImport.OsType + '"'
	$VCObject.Cells("Prop.DatacenterCount").Formula = '"' + $vCenterImport.DatacenterCount + '"'
	$VCObject.Cells("Prop.ClusterCount").Formula = '"' + $vCenterImport.ClusterCount + '"'
	$VCObject.Cells("Prop.HostCount").Formula = '"' + $vCenterImport.HostCount + '"'
	$VCObject.Cells("Prop.VMCount").Formula = '"' + $vCenterImport.VMCount + '"'
	$VCObject.Cells("Prop.PoweredOnVMCount").Formula = '"' + $vCenterImport.PoweredOnVMCount + '"'
	$VCObject.Cells("Prop.TemplateCount").Formula = '"' + $vCenterImport.TemplateCount + '"'
	$VCObject.Cells("Prop.IsConnected").Formula = '"' + $vCenterImport.IsConnected + '"'
	$VCObject.Cells("Prop.ServiceUri").Formula = '"' + $vCenterImport.ServiceUri + '"'
	$VCObject.Cells("Prop.Port").Formula = '"' + $vCenterImport.Port + '"'
	$VCObject.Cells("Prop.ProductLine").Formula = '"' + $vCenterImport.ProductLine + '"'
	$VCObject.Cells("Prop.InstanceUuid").Formula = '"' + $vCenterImport.InstanceUuid + '"'
	$VCObject.Cells("Prop.RefCount").Formula = '"' + $vCenterImport.RefCount + '"'
	$VCObject.Cells("Prop.ServerClock").Formula = '"' + $vCenterImport.ServerClock + '"'
	$VCObject.Cells("Prop.ProvisioningSupported").Formula = '"' + $vCenterImport.ProvisioningSupported + '"'
	$VCObject.Cells("Prop.MultiHostSupported").Formula = '"' + $vCenterImport.MultiHostSupported + '"'
	$VCObject.Cells("Prop.UserShellAccessSupported").Formula = '"' + $vCenterImport.UserShellAccessSupported + '"'
	$VCObject.Cells("Prop.NetworkBackupAndRestoreSupported").Formula = '"' + $vCenterImport.NetworkBackupAndRestoreSupported + '"'
	$VCObject.Cells("Prop.FtDrsWithoutEvcSupported").Formula = '"' + $vCenterImport.FtDrsWithoutEvcSupported + '"'
	$VCObject.Cells("Prop.HciWorkflowSupported").Formula = '"' + $vCenterImport.HciWorkflowSupported + '"'
	$VCObject.Cells("Prop.RootFolder").Formula = '"' + $vCenterImport.RootFolder + '"'
	$VCObject.Cells("Prop.Product").Formula = '"' + $vCenterImport.Product + '"'
	$VCObject.Cells("Prop.FullName").Formula = '"' + $vCenterImport.FullName + '"'
	$VCObject.Cells("Prop.Vendor").Formula = '"' + $vCenterImport.Vendor + '"'
	$VCObject.Cells("Prop.LocaleVersion").Formula = '"' + $vCenterImport.LocaleVersion + '"'
	$VCObject.Cells("Prop.LocaleBuild").Formula = '"' + $vCenterImport.LocaleBuild + '"'
	$VCObject.Cells("Prop.ProductLineId").Formula = '"' + $vCenterImport.ProductLineId + '"'
	$VCObject.Cells("Prop.ApiType").Formula = '"' + $vCenterImport.ApiType + '"'
	$VCObject.Cells("Prop.ApiVersion").Formula = '"' + $vCenterImport.ApiVersion + '"'
	$VCObject.Cells("Prop.LicenseProductName").Formula = '"' + $vCenterImport.LicenseProductName + '"'
	$VCObject.Cells("Prop.LicenseProductVersion").Formula = '"' + $vCenterImport.LicenseProductVersion + '"'
}
#endregion ~~< Draw_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Datacenter
{
	$DatacenterObject.Cells("Prop.Name").Formula = '"' + $Datacenter.Name + '"'
	$DatacenterObject.Cells("Prop.VmFolder").Formula = '"' + $Datacenter.VmFolder + '"'
	$DatacenterObject.Cells("Prop.HostFolder").Formula = '"' + $Datacenter.HostFolder + '"'
	$DatacenterObject.Cells("Prop.DatastoreFolder").Formula = '"' + $Datacenter.DatastoreFolder + '"'
	$DatacenterObject.Cells("Prop.NetworkFolder").Formula = '"' + $Datacenter.NetworkFolder + '"'
	$DatacenterObject.Cells("Prop.Folder").Formula = '"' + $Datacenter.Folder + '"'
	$DatacenterObject.Cells("Prop.FolderId").Formula = '"' + $Datacenter.FolderId + '"'
	$DatacenterObject.Cells("Prop.Datastore").Formula = '"' + $Datacenter.Datastore + '"'
	$DatacenterObject.Cells("Prop.DatastoreId").Formula = '"' + $Datacenter.DatastoreId + '"'
	$DatacenterObject.Cells("Prop.Network").Formula = '"' + $Datacenter.Network + '"'
	$DatacenterObject.Cells("Prop.NetworkId").Formula = '"' + $Datacenter.NetworkId + '"'
	$DatacenterObject.Cells("Prop.DefaultHardwareVersionKey").Formula = '"' + $Datacenter.DefaultHardwareVersionKey + '"'
	$DatacenterObject.Cells("Prop.LinkedView").Formula = '"' + $Datacenter.LinkedView + '"'
	$DatacenterObject.Cells("Prop.Parent").Formula = '"' + $Datacenter.Parent + '"'
	$DatacenterObject.Cells("Prop.OverallStatus").Formula = '"' + $Datacenter.OverallStatus + '"'
	$DatacenterObject.Cells("Prop.ConfigStatus").Formula = '"' + $Datacenter.ConfigStatus + '"'
	$DatacenterObject.Cells("Prop.ConfigIssue").Formula = '"' + $Datacenter.ConfigIssue + '"'
	$DatacenterObject.Cells("Prop.EffectiveRole").Formula = '"' + $Datacenter.EffectiveRole + '"'
	$DatacenterObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $Datacenter.AlarmActionsEnabled + '"'
	$DatacenterObject.Cells("Prop.Tag").Formula = '"' + $Datacenter.Tag + '"'
	$DatacenterObject.Cells("Prop.Value").Formula = '"' + $Datacenter.Value + '"'
	$DatacenterObject.Cells("Prop.AvailableField").Formula = '"' + $Datacenter.AvailableField + '"'
	$DatacenterObject.Cells("Prop.MoRef").Formula = '"' + $Datacenter.MoRef + '"'
}
#endregion ~~< Draw_Datacenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Cluster
{
	$ClusterObject.Cells("Prop.Name").Formula = '"' + $Cluster.Name + '"'
	$ClusterObject.Cells("Prop.Datacenter").Formula = '"' + $Cluster.Datacenter + '"'
	$ClusterObject.Cells("Prop.DatacenterId").Formula = '"' + $Cluster.DatacenterId + '"'
	$ClusterObject.Cells("Prop.HAEnabled").Formula = '"' + $Cluster.HAEnabled + '"'
	$ClusterObject.Cells("Prop.HAAdmissionControlEnabled").Formula = '"' + $Cluster.HAAdmissionControlEnabled + '"'
	$ClusterObject.Cells("Prop.AdmissionControlPolicyCpuFailoverResourcesPercent").Formula = '"' + $Cluster.AdmissionControlPolicyCpuFailoverResourcesPercent + '"'
	$ClusterObject.Cells("Prop.AdmissionControlPolicyMemoryFailoverResourcesPercent").Formula = '"' + $Cluster.AdmissionControlPolicyMemoryFailoverResourcesPercent + '"'
	$ClusterObject.Cells("Prop.AdmissionControlPolicyFailoverLevel").Formula = '"' + $Cluster.AdmissionControlPolicyFailoverLevel + '"'
	$ClusterObject.Cells("Prop.AdmissionControlPolicyAutoComputePercentages").Formula = '"' + $Cluster.AdmissionControlPolicyAutoComputePercentages + '"'
	$ClusterObject.Cells("Prop.AdmissionControlPolicyResourceReductionToToleratePercent").Formula = '"' + $Cluster.AdmissionControlPolicyResourceReductionToToleratePercent + '"'
	$ClusterObject.Cells("Prop.DrsEnabled").Formula = '"' + $Cluster.DrsEnabled + '"'
	$ClusterObject.Cells("Prop.DrsAutomationLevel").Formula = '"' + $Cluster.DrsAutomationLevel + '"'
	$ClusterObject.Cells("Prop.VmMonitoring").Formula = '"' + $Cluster.VmMonitoring + '"'
	$ClusterObject.Cells("Prop.HostMonitoring").Formula = '"' + $Cluster.HostMonitoring + '"'
	$ClusterObject.Cells("Prop.MoRef").Formula = '"' + $Cluster.MoRef + '"'
}
#endregion ~~< Draw_Cluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VmHost
{
	$HostObject.Cells("Prop.Name").Formula = '"' + $VMHost.Name + '"'
	$HostObject.Cells("Prop.Datacenter").Formula = '"' + $VMHost.Datacenter + '"'
	$HostObject.Cells("Prop.DatacenterId").Formula = '"' + $VMHost.DatacenterId + '"'
	$HostObject.Cells("Prop.Cluster").Formula = '"' + $VMHost.Cluster + '"'
	$HostObject.Cells("Prop.ClusterId").Formula = '"' + $VMHost.ClusterId + '"'
	$HostObject.Cells("Prop.Vm").Formula = '"' + $VMHost.Vm + '"'
	$HostObject.Cells("Prop.VmId").Formula = '"' + $VMHost.VmId + '"'
	$HostObject.Cells("Prop.Datastore").Formula = '"' + $VMHost.Datastore + '"'
	$HostObject.Cells("Prop.DatastoreId").Formula = '"' + $VMHost.DatastoreId + '"'
	$HostObject.Cells("Prop.Version").Formula = '"' + $VMHost.Version + '"'
	$HostObject.Cells("Prop.Build").Formula = '"' + $VMHost.Build + '"'
	$HostObject.Cells("Prop.Manufacturer").Formula = '"' + $VMHost.Manufacturer + '"'
	$HostObject.Cells("Prop.Model").Formula = '"' + $VMHost.Model + '"'
	$HostObject.Cells("Prop.LicenseType").Formula = '"' + $VMHost.LicenseType + '"'
	$HostObject.Cells("Prop.BIOSVersion").Formula = '"' + $VMHost.BIOSVersion + '"'
	$HostObject.Cells("Prop.BIOSReleaseDate").Formula = '"' + $VMHost.BIOSReleaseDate + '"'
	$HostObject.Cells("Prop.ProcessorType").Formula = '"' + $VMHost.ProcessorType + '"'
	$HostObject.Cells("Prop.CpuMhz").Formula = '"' + $VMHost.CpuMhz + '"'
	$HostObject.Cells("Prop.NumCpuPkgs").Formula = '"' + $VMHost.NumCpuPkgs + '"'
	$HostObject.Cells("Prop.NumCpuCores").Formula = '"' + $VMHost.NumCpuCores + '"'
	$HostObject.Cells("Prop.NumCpuThreads").Formula = '"' + $VMHost.NumCpuThreads + '"'
	$HostObject.Cells("Prop.Memory").Formula = '"' + $VMHost.Memory + '"'
	$HostObject.Cells("Prop.MaxEVCMode").Formula = '"' + $VMHost.MaxEVCMode + '"'
	$HostObject.Cells("Prop.NumNics").Formula = '"' + $VMHost.NumNics + '"'
	$HostObject.Cells("Prop.ManagemetIP").Formula = '"' + $VMHost.ManagemetIP + '"'
	$HostObject.Cells("Prop.ManagemetMacAddress").Formula = '"' + $VMHost.ManagemetMacAddress + '"'
	$HostObject.Cells("Prop.ManagemetVMKernel").Formula = '"' + $VMHost.ManagemetVMKernel + '"'
	$HostObject.Cells("Prop.ManagemetSubnetMask").Formula = '"' + $VMHost.ManagemetSubnetMask + '"'
	$HostObject.Cells("Prop.vMotionIP").Formula = '"' + $VMHost.vMotionIP + '"'
	$HostObject.Cells("Prop.vMotionMacAddress").Formula = '"' + $VMHost.vMotionMacAddress + '"'
	$HostObject.Cells("Prop.vMotionVMKernel").Formula = '"' + $VMHost.vMotionVMKernel + '"'
	$HostObject.Cells("Prop.vMotionSubnetMask").Formula = '"' + $VMHost.vMotionSubnetMask + '"'
	$HostObject.Cells("Prop.FtIP").Formula = '"' + $VMHost.FtIP + '"'
	$HostObject.Cells("Prop.FtMacAddress").Formula = '"' + $VMHost.FtMacAddress + '"'
	$HostObject.Cells("Prop.FtVMKernel").Formula = '"' + $VMHost.FtVMKernel + '"'
	$HostObject.Cells("Prop.FtSubnetMask").Formula = '"' + $VMHost.FtSubnetMask + '"'
	$HostObject.Cells("Prop.VSANIP").Formula = '"' + $VMHost.VSANIP + '"'
	$HostObject.Cells("Prop.VSANMacAddress").Formula = '"' + $VMHost.VSANMacAddress + '"'
	$HostObject.Cells("Prop.VSANVMKernel").Formula = '"' + $VMHost.VSANVMKernel + '"'
	$HostObject.Cells("Prop.VSANSubnetMask").Formula = '"' + $VMHost.VSANSubnetMask + '"'
	$HostObject.Cells("Prop.NumHBAs").Formula = '"' + $VMHost.NumHBAs + '"'
	$HostObject.Cells("Prop.iSCSIIP").Formula = '"' + $VMHost.iSCSIIP + '"'
	$HostObject.Cells("Prop.iSCSIMac").Formula = '"' + $VMHost.iSCSIMac + '"'
	$HostObject.Cells("Prop.iSCSIVMKernel").Formula = '"' + $VMHost.iSCSIVMKernel + '"'
	$HostObject.Cells("Prop.iSCSISubnetMask").Formula = '"' + $VMHost.iSCSISubnetMask + '"'
	$HostObject.Cells("Prop.iSCSIAdapter").Formula = '"' + $VMHost.iSCSIAdapter + '"'
	$HostObject.Cells("Prop.iSCSILinkUp").Formula = '"' + $VMHost.iSCSILinkUp + '"'
	$HostObject.Cells("Prop.iSCSIMTU").Formula = '"' + $VMHost.iSCSIMTU + '"'
	$HostObject.Cells("Prop.iSCSINICDriver").Formula = '"' + $VMHost.iSCSINICDriver + '"'
	$HostObject.Cells("Prop.iSCSINICDriverVersion").Formula = '"' + $VMHost.iSCSINICDriverVersion + '"'
	$HostObject.Cells("Prop.iSCSINICFirmwareVersion").Formula = '"' + $VMHost.iSCSINICFirmwareVersion + '"'
	$HostObject.Cells("Prop.iSCSIPathStatus").Formula = '"' + $VMHost.iSCSIPathStatus + '"'
	$HostObject.Cells("Prop.iSCSIVlanID").Formula = '"' + $VMHost.iSCSIVlanID + '"'
	$HostObject.Cells("Prop.iSCSIVswitch").Formula = '"' + $VMHost.iSCSIVswitch + '"'
	$HostObject.Cells("Prop.iSCSICompliantStatus").Formula = '"' + $VMHost.iSCSICompliantStatus + '"'
	$HostObject.Cells("Prop.IScsiName").Formula = '"' + $VMHost.IScsiName + '"'
	$HostObject.Cells("Prop.PortGroup").Formula = '"' + $VMHost.PortGroup + '"'
	$HostObject.Cells("Prop.CdpLldpInfo").Formula = '"' + $VMHost.CdpLldpInfo + '"'
	$HostObject.Cells("Prop.MoRef").Formula = '"' + $VMHost.MoRef + '"'
}
#endregion ~~< Draw_VmHost >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VM
{
	$VMObject.Cells("Prop.Name").Formula = '"' + $VM.Name + '"'
	$VMObject.Cells("Prop.Datacenter").Formula = '"' + $VM.Datacenter + '"'
	$VMObject.Cells("Prop.DatacenterId").Formula = '"' + $VM.DatacenterId + '"'
	$VMObject.Cells("Prop.Cluster").Formula = '"' + $VM.Cluster + '"'
	$VMObject.Cells("Prop.ClusterId").Formula = '"' + $VM.ClusterId + '"'
	$VMObject.Cells("Prop.VmHost").Formula = '"' + $VM.VmHost + '"'
	$VMObject.Cells("Prop.VmHostId").Formula = '"' + $VM.VmHostId + '"'
	$VMObject.Cells("Prop.DatastoreCluster").Formula = '"' + $VM.DatastoreCluster + '"'
	$VMObject.Cells("Prop.DatastoreClusterId").Formula = '"' + $VM.DatastoreClusterId + '"'
	$VMObject.Cells("Prop.Datastore").Formula = '"' + $VM.Datastore + '"'
	$VMObject.Cells("Prop.DatastoreId").Formula = '"' + $VM.DatastoreId + '"'
	$VMObject.Cells("Prop.ResourcePool").Formula = '"' + $VM.ResourcePool + '"'
	$VMObject.Cells("Prop.ResourcePoolId").Formula = '"' + $VM.ResourcePoolId + '"'
	$VMObject.Cells("Prop.vSwitch").Formula = '"' + $VM.vSwitch + '"'
	$VMObject.Cells("Prop.vSwitchId").Formula = '"' + $VM.vSwitchId + '"'
	$VMObject.Cells("Prop.PortGroup").Formula = '"' + $VM.PortGroup + '"'
	$VMObject.Cells("Prop.PortGroupId").Formula = '"' + $VM.PortGroupId + '"'
	$VMObject.Cells("Prop.OS").Formula = '"' + $VM.OS + '"'
	$VMObject.Cells("Prop.Version").Formula = '"' + $VM.Version + '"'
	$VMObject.Cells("Prop.VMToolsVersion").Formula = '"' + $VM.VMToolsVersion + '"'
	$VMObject.Cells("Prop.ToolsVersionStatus").Formula = '"' + $VM.ToolsVersionStatus + '"'
	$VMObject.Cells("Prop.ToolsStatus").Formula = '"' + $VM.ToolsStatus + '"'
	$VMObject.Cells("Prop.ToolsRunningStatus").Formula = '"' + $VM.ToolsRunningStatus + '"'
	$VMObject.Cells("Prop.Folder").Formula = '"' + $VM.Folder + '"'
	$VMObject.Cells("Prop.FolderId").Formula = '"' + $VM.FolderId + '"'
	$VMObject.Cells("Prop.NumCPU").Formula = '"' + $VM.NumCPU + '"'
	$VMObject.Cells("Prop.CoresPerSocket").Formula = '"' + $VM.CoresPerSocket + '"'
	$VMObject.Cells("Prop.MemoryGB").Formula = '"' + $VM.MemoryGB + '"'
	$VMObject.Cells("Prop.IP").Formula = '"' + $VM.IP + '"'
	$VMObject.Cells("Prop.MacAddress").Formula = '"' + $VM.MacAddress + '"'
	$VMObject.Cells("Prop.NumVirtualDisks").Formula = '"' + $VM.NumVirtualDisks + '"'
	$VMObject.Cells("Prop.VmdkInfo").Formula = '"' + $VM.VmdkInfo + '"'
	$VMObject.Cells("Prop.Volumes").Formula = '"' + $VM.Volumes + '"'
	$VMObject.Cells("Prop.ProvisionedSpaceGB").Formula = '"' + $VM.ProvisionedSpaceGB + '"'
	$VMObject.Cells("Prop.NumEthernetCards").Formula = '"' + $VM.NumEthernetCards + '"'
	$VMObject.Cells("Prop.CpuReservation").Formula = '"' + $VM.CpuReservation + '"'
	$VMObject.Cells("Prop.MemoryReservation").Formula = '"' + $VM.MemoryReservation + '"'
	$VMObject.Cells("Prop.CpuHotAddEnabled").Formula = '"' + $VM.CpuHotAddEnabled + '"'
	$VMObject.Cells("Prop.CpuHotRemoveEnabled").Formula = '"' + $VM.CpuHotRemoveEnabled + '"'
	$VMObject.Cells("Prop.MemoryHotAddEnabled").Formula = '"' + $VM.MemoryHotAddEnabled + '"'
	$VMObject.Cells("Prop.SRM").Formula = '"' + $VM.SRM + '"'
	$VMObject.Cells("Prop.Snapshot").Formula = '"' + $VM.Snapshot + '"'
	$VMObject.Cells("Prop.RootSnapshot").Formula = '"' + $VM.RootSnapshot + '"'
	$VMObject.Cells("Prop.MoRef").Formula = '"' + $VM.MoRef + '"'
}
#endregion ~~< Draw_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Template
{
	$TemplateObject.Cells("Prop.Name").Formula = '"' + $Template.Name + '"'
	$TemplateObject.Cells("Prop.Datacenter").Formula = '"' + $Template.Datacenter + '"'
	$TemplateObject.Cells("Prop.DatacenterId").Formula = '"' + $Template.DatacenterId + '"'
	$TemplateObject.Cells("Prop.Cluster").Formula = '"' + $Template.Cluster + '"'
	$TemplateObject.Cells("Prop.ClusterId").Formula = '"' + $Template.ClusterId + '"'
	$TemplateObject.Cells("Prop.DatastoreCluster").Formula = '"' + $Template.DatastoreCluster + '"'
	$TemplateObject.Cells("Prop.DatastoreClusterId").Formula = '"' + $Template.DatastoreClusterId + '"'
	$TemplateObject.Cells("Prop.Datastore").Formula = '"' + $Template.Datastore + '"'
	$TemplateObject.Cells("Prop.DatastoreId").Formula = '"' + $Template.DatastoreId + '"'
	$TemplateObject.Cells("Prop.VmHost").Formula = '"' + $Template.VmHost + '"'
	$TemplateObject.Cells("Prop.VmHostId").Formula = '"' + $Template.VmHostId + '"'
	$TemplateObject.Cells("Prop.OS").Formula = '"' + $Template.OS + '"'
	$TemplateObject.Cells("Prop.Version").Formula = '"' + $Template.Version + '"'
	$TemplateObject.Cells("Prop.ToolsVersion").Formula = '"' + $Template.ToolsVersion + '"'
	$TemplateObject.Cells("Prop.ToolsVersionStatus").Formula = '"' + $Template.ToolsVersionStatus + '"'
	$TemplateObject.Cells("Prop.ToolsStatus").Formula = '"' + $Template.ToolsStatus + '"'
	$TemplateObject.Cells("Prop.ToolsRunningStatus").Formula = '"' + $Template.ToolsRunningStatus + '"'
	$TemplateObject.Cells("Prop.Folder").Formula = '"' + $Template.Folder + '"'
	$TemplateObject.Cells("Prop.FolderId").Formula = '"' + $Template.FolderId + '"'
	$TemplateObject.Cells("Prop.NumCPU").Formula = '"' + $Template.NumCPU + '"'
	$TemplateObject.Cells("Prop.CoresPerSocket").Formula = '"' + $Template.CoresPerSocket + '"'
	$TemplateObject.Cells("Prop.MemoryGB").Formula = '"' + $Template.MemoryGB + '"'
	$TemplateObject.Cells("Prop.MacAddress").Formula = '"' + $Template.MacAddress + '"'
	$TemplateObject.Cells("Prop.NumEthernetCards").Formula = '"' + $Template.NumEthernetCards + '"'
	$TemplateObject.Cells("Prop.NumVirtualDisks").Formula = '"' + $Template.NumVirtualDisks + '"'
	$TemplateObject.Cells("Prop.CpuReservation").Formula = '"' + $Template.CpuReservation + '"'
	$TemplateObject.Cells("Prop.MemoryReservation").Formula = '"' + $Template.MemoryReservation + '"'
	$TemplateObject.Cells("Prop.CpuHotAddEnabled").Formula = '"' + $Template.CpuHotAddEnabled + '"'
	$TemplateObject.Cells("Prop.CpuHotRemoveEnabled").Formula = '"' + $Template.CpuHotRemoveEnabled + '"'
	$TemplateObject.Cells("Prop.MemoryHotAddEnabled").Formula = '"' + $Template.MemoryHotAddEnabled + '"'
	$TemplateObject.Cells("Prop.MoRef").Formula = '"' + $Template.MoRef + '"'
}
#endregion ~~< Draw_Template >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Folder
{
	$FolderObject.Cells("Prop.Name").Formula = '"' + $Folder.Name + '"'
	$FolderObject.Cells("Prop.Datacenter").Formula = '"' + $Folder.Datacenter + '"'
	$FolderObject.Cells("Prop.DatacenterId").Formula = '"' + $Folder.DatacenterId + '"'
	$FolderObject.Cells("Prop.ChildType").Formula = '"' + $Folder.ChildType + '"'
	$FolderObject.Cells("Prop.ChildEntity").Formula = '"' + $Folder.ChildEntity + '"'
	$FolderObject.Cells("Prop.LinkedView").Formula = '"' + $Folder.LinkedView + '"'
	$FolderObject.Cells("Prop.Parent").Formula = '"' + $Folder.Parent + '"'
	$FolderObject.Cells("Prop.ParentId").Formula = '"' + $Folder.ParentId + '"'
	$FolderObject.Cells("Prop.CustomValue").Formula = '"' + $Folder.CustomValue + '"'
	$FolderObject.Cells("Prop.OverallStatus").Formula = '"' + $Folder.OverallStatus + '"'
	$FolderObject.Cells("Prop.ConfigStatus").Formula = '"' + $Folder.ConfigStatus + '"'
	$FolderObject.Cells("Prop.ConfigIssue").Formula = '"' + $Folder.ConfigIssue + '"'
	$FolderObject.Cells("Prop.EffectiveRole").Formula = '"' + $Folder.EffectiveRole + '"'
	$FolderObject.Cells("Prop.Permission").Formula = '"' + $Folder.Permission + '"'
	$FolderObject.Cells("Prop.DisabledMethod").Formula = '"' + $Folder.DisabledMethod + '"'
	$FolderObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $Folder.AlarmActionsEnabled + '"'
	$FolderObject.Cells("Prop.Tag").Formula = '"' + $Folder.Tag + '"'
	$FolderObject.Cells("Prop.Value").Formula = '"' + $Folder.Value + '"'
	$FolderObject.Cells("Prop.AvailableField").Formula = '"' + $Folder.AvailableField + '"'
	$FolderObject.Cells("Prop.MoRef").Formula = '"' + $Folder.MoRef + '"'
}
#endregion ~~< Draw_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SubFolder
{
	$SubFolderObject.Cells("Prop.Name").Formula = '"' + $SubFolder.Name + '"'
	$SubFolderObject.Cells("Prop.Datacenter").Formula = '"' + $SubFolder.Datacenter + '"'
	$SubFolderObject.Cells("Prop.DatacenterId").Formula = '"' + $SubFolder.DatacenterId + '"'
	$SubFolderObject.Cells("Prop.ChildType").Formula = '"' + $SubFolder.ChildType + '"'
	$SubFolderObject.Cells("Prop.ChildEntity").Formula = '"' + $SubFolder.ChildEntity + '"'
	$SubFolderObject.Cells("Prop.LinkedView").Formula = '"' + $SubFolder.LinkedView + '"'
	$SubFolderObject.Cells("Prop.Parent").Formula = '"' + $SubFolder.Parent + '"'
	$SubFolderObject.Cells("Prop.ParentId").Formula = '"' + $SubFolder.ParentId + '"'
	$SubFolderObject.Cells("Prop.CustomValue").Formula = '"' + $SubFolder.CustomValue + '"'
	$SubFolderObject.Cells("Prop.OverallStatus").Formula = '"' + $SubFolder.OverallStatus + '"'
	$SubFolderObject.Cells("Prop.ConfigStatus").Formula = '"' + $SubFolder.ConfigStatus + '"'
	$SubFolderObject.Cells("Prop.ConfigIssue").Formula = '"' + $SubFolder.ConfigIssue + '"'
	$SubFolderObject.Cells("Prop.EffectiveRole").Formula = '"' + $SubFolder.EffectiveRole + '"'
	$SubFolderObject.Cells("Prop.Permission").Formula = '"' + $SubFolder.Permission + '"'
	$SubFolderObject.Cells("Prop.DisabledMethod").Formula = '"' + $SubFolder.DisabledMethod + '"'
	$SubFolderObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $SubFolder.AlarmActionsEnabled + '"'
	$SubFolderObject.Cells("Prop.Tag").Formula = '"' + $SubFolder.Tag + '"'
	$SubFolderObject.Cells("Prop.Value").Formula = '"' + $SubFolder.Value + '"'
	$SubFolderObject.Cells("Prop.AvailableField").Formula = '"' + $SubFolder.AvailableField + '"'
	$SubFolderObject.Cells("Prop.MoRef").Formula = '"' + $SubFolder.MoRef + '"'
}
#endregion ~~< Draw_SubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SubSubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SubSubFolder
{
	$SubSubFolderObject.Cells("Prop.Name").Formula = '"' + $SubSubFolder.Name + '"'
	$SubSubFolderObject.Cells("Prop.Datacenter").Formula = '"' + $SubSubFolder.Datacenter + '"'
	$SubSubFolderObject.Cells("Prop.DatacenterId").Formula = '"' + $SubSubFolder.DatacenterId + '"'
	$SubSubFolderObject.Cells("Prop.ChildType").Formula = '"' + $SubSubFolder.ChildType + '"'
	$SubSubFolderObject.Cells("Prop.ChildEntity").Formula = '"' + $SubSubFolder.ChildEntity + '"'
	$SubSubFolderObject.Cells("Prop.LinkedView").Formula = '"' + $SubSubFolder.LinkedView + '"'
	$SubSubFolderObject.Cells("Prop.Parent").Formula = '"' + $SubSubFolder.Parent + '"'
	$SubSubFolderObject.Cells("Prop.ParentId").Formula = '"' + $SubSubFolder.ParentId + '"'
	$SubSubFolderObject.Cells("Prop.CustomValue").Formula = '"' + $SubSubFolder.CustomValue + '"'
	$SubSubFolderObject.Cells("Prop.OverallStatus").Formula = '"' + $SubSubFolder.OverallStatus + '"'
	$SubSubFolderObject.Cells("Prop.ConfigStatus").Formula = '"' + $SubSubFolder.ConfigStatus + '"'
	$SubSubFolderObject.Cells("Prop.ConfigIssue").Formula = '"' + $SubSubFolder.ConfigIssue + '"'
	$SubSubFolderObject.Cells("Prop.EffectiveRole").Formula = '"' + $SubSubFolder.EffectiveRole + '"'
	$SubSubFolderObject.Cells("Prop.Permission").Formula = '"' + $SubSubFolder.Permission + '"'
	$SubSubFolderObject.Cells("Prop.DisabledMethod").Formula = '"' + $SubSubFolder.DisabledMethod + '"'
	$SubSubFolderObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $SubSubFolder.AlarmActionsEnabled + '"'
	$SubSubFolderObject.Cells("Prop.Tag").Formula = '"' + $SubSubFolder.Tag + '"'
	$SubSubFolderObject.Cells("Prop.Value").Formula = '"' + $SubSubFolder.Value + '"'
	$SubSubFolderObject.Cells("Prop.AvailableField").Formula = '"' + $SubSubFolder.AvailableField + '"'
	$SubSubFolderObject.Cells("Prop.MoRef").Formula = '"' + $SubSubFolder.MoRef + '"'
}
#endregion ~~< Draw_SubSubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SubSubSubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SubSubSubFolder
{
	$SubSubSubFolderObject.Cells("Prop.Name").Formula = '"' + $SubSubSubFolder.Name + '"'
	$SubSubSubFolderObject.Cells("Prop.Datacenter").Formula = '"' + $SubSubSubFolder.Datacenter + '"'
	$SubSubSubFolderObject.Cells("Prop.DatacenterId").Formula = '"' + $SubSubSubFolder.DatacenterId + '"'
	$SubSubSubFolderObject.Cells("Prop.ChildType").Formula = '"' + $SubSubSubFolder.ChildType + '"'
	$SubSubSubFolderObject.Cells("Prop.ChildEntity").Formula = '"' + $SubSubSubFolder.ChildEntity + '"'
	$SubSubSubFolderObject.Cells("Prop.LinkedView").Formula = '"' + $SubSubSubFolder.LinkedView + '"'
	$SubSubSubFolderObject.Cells("Prop.Parent").Formula = '"' + $SubSubSubFolder.Parent + '"'
	$SubSubSubFolderObject.Cells("Prop.ParentId").Formula = '"' + $SubSubSubFolder.ParentId + '"'
	$SubSubSubFolderObject.Cells("Prop.CustomValue").Formula = '"' + $SubSubSubFolder.CustomValue + '"'
	$SubSubSubFolderObject.Cells("Prop.OverallStatus").Formula = '"' + $SubSubSubFolder.OverallStatus + '"'
	$SubSubSubFolderObject.Cells("Prop.ConfigStatus").Formula = '"' + $SubSubSubFolder.ConfigStatus + '"'
	$SubSubSubFolderObject.Cells("Prop.ConfigIssue").Formula = '"' + $SubSubSubFolder.ConfigIssue + '"'
	$SubSubSubFolderObject.Cells("Prop.EffectiveRole").Formula = '"' + $SubSubSubFolder.EffectiveRole + '"'
	$SubSubSubFolderObject.Cells("Prop.Permission").Formula = '"' + $SubSubSubFolder.Permission + '"'
	$SubSubSubFolderObject.Cells("Prop.DisabledMethod").Formula = '"' + $SubSubSubFolder.DisabledMethod + '"'
	$SubSubSubFolderObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $SubSubSubFolder.AlarmActionsEnabled + '"'
	$SubSubSubFolderObject.Cells("Prop.Tag").Formula = '"' + $SubSubSubFolder.Tag + '"'
	$SubSubSubFolderObject.Cells("Prop.Value").Formula = '"' + $SubSubSubFolder.Value + '"'
	$SubSubSubFolderObject.Cells("Prop.AvailableField").Formula = '"' + $SubSubSubFolder.AvailableField + '"'
	$SubSubSubFolderObject.Cells("Prop.MoRef").Formula = '"' + $SubSubSubFolder.MoRef + '"'
}
#endregion ~~< Draw_SubSubSubFolder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_RDM
{
	$RDMObject.Cells("Prop.ScsiCanonicalName").Formula = '"' + $HardDisk.ScsiCanonicalName + '"'
	$RDMObject.Cells("Prop.Cluster").Formula = '"' + $HardDisk.Cluster + '"'
	$RDMObject.Cells("Prop.ClusterId").Formula = '"' + $HardDisk.ClusterId + '"'
	$RDMObject.Cells("Prop.Vm").Formula = '"' + $HardDisk.Vm + '"'
	$RDMObject.Cells("Prop.VmId").Formula = '"' + $HardDisk.VmId + '"'
	$RDMObject.Cells("Prop.Label").Formula = '"' + $HardDisk.Label + '"'
	$RDMObject.Cells("Prop.CapacityGB").Formula = '"' + [math]::Round([decimal]$HardDisk.CapacityGB, 2) + '"'
	$RDMObject.Cells("Prop.DiskType").Formula = '"' + $HardDisk.DiskType + '"'
	$RDMObject.Cells("Prop.Persistence").Formula = '"' + $HardDisk.Persistence + '"'
	$RDMObject.Cells("Prop.CompatibilityMode").Formula = '"' + $HardDisk.CompatibilityMode + '"'
	$RDMObject.Cells("Prop.DeviceName").Formula = '"' + $HardDisk.DeviceName + '"'
	$RDMObject.Cells("Prop.Sharing").Formula = '"' + $HardDisk.Sharing + '"'
}
#endregion ~~< Draw_RDM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SRM
{
	$SrmObject.Cells("Prop.Name").Formula = '"' + $SrmVM.Name + '"'
	$SrmObject.Cells("Prop.OS").Formula = '"' + $SrmVM.ConfigGuestFullName + '"'
	$SrmObject.Cells("Prop.Version").Formula = '"' + $SrmVM.Version + '"'
	$SrmObject.Cells("Prop.Folder").Formula = '"' + $SrmVM.Folder + '"'
	$SrmObject.Cells("Prop.NumCPU").Formula = '"' + $SrmVM.NumCPU + '"'
	$SrmObject.Cells("Prop.CoresPerSocket").Formula = '"' + $SrmVM.CoresPerSocket + '"'
	$SrmObject.Cells("Prop.MemoryGB").Formula = '"' + $SrmVM.MemoryGB + '"'
}
#endregion ~~< Draw_SRM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_DatastoreCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DatastoreCluster
{
	$DatastoreClusObject.Cells("Prop.Name").Formula = '"' + $DatastoreCluster.Name + '"'
	$DatastoreClusObject.Cells("Prop.Datacenter").Formula = '"' + $DatastoreCluster.Datacenter + '"'
	$DatastoreClusObject.Cells("Prop.DatacenterId").Formula = '"' + $DatastoreCluster.DatacenterId + '"'
	$DatastoreClusObject.Cells("Prop.Cluster").Formula = '"' + $DatastoreCluster.Cluster + '"'
	$DatastoreClusObject.Cells("Prop.ClusterId").Formula = '"' + $DatastoreCluster.ClusterId + '"'
	$DatastoreClusObject.Cells("Prop.VmHost").Formula = '"' + $DatastoreCluster.VmHost + '"'
	$DatastoreClusObject.Cells("Prop.VmHostId").Formula = '"' + $DatastoreCluster.VmHostId + '"'
	$DatastoreClusObject.Cells("Prop.SdrsAutomationLevel").Formula = '"' + $DatastoreCluster.SdrsAutomationLevel + '"'
	$DatastoreClusObject.Cells("Prop.IOLoadBalanceEnabled").Formula = '"' + $DatastoreCluster.IOLoadBalanceEnabled + '"'
	$DatastoreClusObject.Cells("Prop.CapacityGB").Formula = '"' + $DatastoreCluster.CapacityGB + '"'
	$DatastoreClusObject.Cells("Prop.MoRef").Formula = '"' + $DatastoreCluster.MoRef + '"'
}
#endregion ~~< Draw_DatastoreCluster >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_Datastore
{
	$DatastoreObject.Cells("Prop.Name").Formula = '"' + $Datastore.Name + '"'
	$DatastoreObject.Cells("Prop.Datacenter").Formula = '"' + $Datastore.Datacenter + '"'
	$DatastoreObject.Cells("Prop.DatacenterId").Formula = '"' + $Datastore.DatacenterId + '"'
	$DatastoreObject.Cells("Prop.Cluster").Formula = '"' + $Datastore.Cluster + '"'
	$DatastoreObject.Cells("Prop.ClusterId").Formula = '"' + $Datastore.ClusterId + '"'
	$DatastoreObject.Cells("Prop.DatastoreCluster").Formula = '"' + $Datastore.DatastoreCluster + '"'
	$DatastoreObject.Cells("Prop.DatastoreClusterId").Formula = '"' + $Datastore.DatastoreClusterId + '"'
	$DatastoreObject.Cells("Prop.VmHost").Formula = '"' + $Datastore.VmHost + '"'
	$DatastoreObject.Cells("Prop.VmHostId").Formula = '"' + $Datastore.VmHostId + '"'
	$DatastoreObject.Cells("Prop.Vm").Formula = '"' + $Datastore.Vm + '"'
	$DatastoreObject.Cells("Prop.VmId").Formula = '"' + $Datastore.VmId + '"'
	$DatastoreObject.Cells("Prop.Type").Formula = '"' + $Datastore.Type + '"'
	$DatastoreObject.Cells("Prop.FileSystemVersion").Formula = '"' + $Datastore.FileSystemVersion + '"'
	$DatastoreObject.Cells("Prop.DiskName").Formula = '"' + $Datastore.DiskName + '"'
	$DatastoreObject.Cells("Prop.DiskPath").Formula = '"' + $Datastore.DiskPath + '"'
	$DatastoreObject.Cells("Prop.DiskUuid").Formula = '"' + $Datastore.DiskUuid + '"'
	$DatastoreObject.Cells("Prop.StorageIOControlEnabled").Formula = '"' + $Datastore.StorageIOControlEnabled + '"'
	$DatastoreObject.Cells("Prop.CapacityGB").Formula = '"' + $Datastore.CapacityGB + '"'
	$DatastoreObject.Cells("Prop.FreeSpaceGB").Formula = '"' + $Datastore.FreeSpaceGB + '"'
	$DatastoreObject.Cells("Prop.CongestionThresholdMillisecond").Formula = '"' + $Datastore.CongestionThresholdMillisecond + '"'
	$DatastoreObject.Cells("Prop.MoRef").Formula = '"' + $Datastore.MoRef + '"'
}
#endregion ~~< Draw_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ResourcePool
{
	$ResourcePoolObject.Cells("Prop.Name").Formula = '"' + $ResourcePool.Name + '"'
	$ResourcePoolObject.Cells("Prop.Cluster").Formula = '"' + $ResourcePool.Cluster + '"'
	$ResourcePoolObject.Cells("Prop.ClusterId").Formula = '"' + $ResourcePool.ClusterId + '"'
	$ResourcePoolObject.Cells("Prop.Vm").Formula = '"' + $ResourcePool.Vm + '"'
	$ResourcePoolObject.Cells("Prop.VmId").Formula = '"' + $ResourcePool.VmId + '"'
	$ResourcePoolObject.Cells("Prop.CpuSharesLevel").Formula = '"' + $ResourcePool.CpuSharesLevel + '"'
	$ResourcePoolObject.Cells("Prop.NumCpuShares").Formula = '"' + $ResourcePool.NumCpuShares + '"'
	$ResourcePoolObject.Cells("Prop.CpuReservationMHz").Formula = '"' + $ResourcePool.CpuReservationMHz + '"'
	$ResourcePoolObject.Cells("Prop.CpuExpandableReservation").Formula = '"' + $ResourcePool.CpuExpandableReservation + '"'
	$ResourcePoolObject.Cells("Prop.CpuLimitMHz").Formula = '"' + $ResourcePool.CpuLimitMHz + '"'
	$ResourcePoolObject.Cells("Prop.MemSharesLevel").Formula = '"' + $ResourcePool.MemSharesLevel + '"'
	$ResourcePoolObject.Cells("Prop.NumMemShares").Formula = '"' + $ResourcePool.NumMemShares + '"'
	$ResourcePoolObject.Cells("Prop.MemReservationGB").Formula = '"' + $ResourcePool.MemReservationGB + '"'
	$ResourcePoolObject.Cells("Prop.MemExpandableReservation").Formula = '"' + $ResourcePool.MemExpandableReservation + '"'
	$ResourcePoolObject.Cells("Prop.MemLimitGB").Formula = '"' + $ResourcePool.MemLimitGB + '"'
	$ResourcePoolObject.Cells("Prop.MoRef").Formula = '"' + $ResourcePool.MoRef + '"'
}
#endregion ~~< Draw_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SubResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SubResourcePool
{
	$SubResourcePoolObject.Cells("Prop.Name").Formula = '"' + $SubResourcePool.Name + '"'
	$SubResourcePoolObject.Cells("Prop.Cluster").Formula = '"' + $SubResourcePool.Cluster + '"'
	$SubResourcePoolObject.Cells("Prop.ClusterId").Formula = '"' + $SubResourcePool.ClusterId + '"'
	$SubResourcePoolObject.Cells("Prop.Vm").Formula = '"' + $SubResourcePool.Vm + '"'
	$SubResourcePoolObject.Cells("Prop.VmId").Formula = '"' + $SubResourcePool.VmId + '"'
	$SubResourcePoolObject.Cells("Prop.CpuSharesLevel").Formula = '"' + $SubResourcePool.CpuSharesLevel + '"'
	$SubResourcePoolObject.Cells("Prop.NumCpuShares").Formula = '"' + $SubResourcePool.NumCpuShares + '"'
	$SubResourcePoolObject.Cells("Prop.CpuReservationMHz").Formula = '"' + $SubResourcePool.CpuReservationMHz + '"'
	$SubResourcePoolObject.Cells("Prop.CpuExpandableReservation").Formula = '"' + $SubResourcePool.CpuExpandableReservation + '"'
	$SubResourcePoolObject.Cells("Prop.CpuLimitMHz").Formula = '"' + $SubResourcePool.CpuLimitMHz + '"'
	$SubResourcePoolObject.Cells("Prop.MemSharesLevel").Formula = '"' + $SubResourcePool.MemSharesLevel + '"'
	$SubResourcePoolObject.Cells("Prop.NumMemShares").Formula = '"' + $SubResourcePool.NumMemShares + '"'
	$SubResourcePoolObject.Cells("Prop.MemReservationGB").Formula = '"' + $SubResourcePool.MemReservationGB + '"'
	$SubResourcePoolObject.Cells("Prop.MemExpandableReservation").Formula = '"' + $SubResourcePool.MemExpandableReservation + '"'
	$SubResourcePoolObject.Cells("Prop.MemLimitGB").Formula = '"' + $SubResourcePool.MemLimitGB + '"'
	$SubResourcePoolObject.Cells("Prop.MoRef").Formula = '"' + $SubResourcePool.MoRef + '"'
}
#endregion ~~< Draw_SubResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_SubSubResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_SubSubResourcePool
{
	$SubSubResourcePoolObject.Cells("Prop.Name").Formula = '"' + $SubSubResourcePool.Name + '"'
	$SubSubResourcePoolObject.Cells("Prop.Cluster").Formula = '"' + $SubSubResourcePool.Cluster + '"'
	$SubSubResourcePoolObject.Cells("Prop.ClusterId").Formula = '"' + $SubSubResourcePool.ClusterId + '"'
	$SubSubResourcePoolObject.Cells("Prop.Vm").Formula = '"' + $SubSubResourcePool.Vm + '"'
	$SubSubResourcePoolObject.Cells("Prop.VmId").Formula = '"' + $SubSubResourcePool.VmId + '"'
	$SubSubResourcePoolObject.Cells("Prop.CpuSharesLevel").Formula = '"' + $SubSubResourcePool.CpuSharesLevel + '"'
	$SubSubResourcePoolObject.Cells("Prop.NumCpuShares").Formula = '"' + $SubSubResourcePool.NumCpuShares + '"'
	$SubSubResourcePoolObject.Cells("Prop.CpuReservationMHz").Formula = '"' + $SubSubResourcePool.CpuReservationMHz + '"'
	$SubSubResourcePoolObject.Cells("Prop.CpuExpandableReservation").Formula = '"' + $SubSubResourcePool.CpuExpandableReservation + '"'
	$SubSubResourcePoolObject.Cells("Prop.CpuLimitMHz").Formula = '"' + $SubSubResourcePool.CpuLimitMHz + '"'
	$SubSubResourcePoolObject.Cells("Prop.MemSharesLevel").Formula = '"' + $SubSubResourcePool.MemSharesLevel + '"'
	$SubSubResourcePoolObject.Cells("Prop.NumMemShares").Formula = '"' + $SubSubResourcePool.NumMemShares + '"'
	$SubSubResourcePoolObject.Cells("Prop.MemReservationGB").Formula = '"' + $SubSubResourcePool.MemReservationGB + '"'
	$SubSubResourcePoolObject.Cells("Prop.MemExpandableReservation").Formula = '"' + $SubSubResourcePool.MemExpandableReservation + '"'
	$SubSubResourcePoolObject.Cells("Prop.MemLimitGB").Formula = '"' + $SubSubResourcePool.MemLimitGB + '"'
	$SubSubResourcePoolObject.Cells("Prop.MoRef").Formula = '"' + $SubSubResourcePool.MoRef + '"'
}
#endregion ~~< Draw_SubSubResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VsSwitch
{
	$VsSwitchObject.Cells("Prop.Name").Formula = '"' + $VsSwitch.Name + '"'
	$VsSwitchObject.Cells("Prop.Datacenter").Formula = '"' + $VsSwitch.Datacenter + '"'
	$VsSwitchObject.Cells("Prop.DatacenterId").Formula = '"' + $VsSwitch.DatacenterId + '"'
	$VsSwitchObject.Cells("Prop.Cluster").Formula = '"' + $VsSwitch.Cluster + '"'
	$VsSwitchObject.Cells("Prop.ClusterId").Formula = '"' + $VsSwitch.ClusterId + '"'
	$VsSwitchObject.Cells("Prop.VmHost").Formula = '"' + $VsSwitch.VmHost + '"'
	$VsSwitchObject.Cells("Prop.VmHostId").Formula = '"' + $VsSwitch.VmHostId + '"'
	$VsSwitchObject.Cells("Prop.Vm").Formula = '"' + $VsSwitch.Vm + '"'
	$VsSwitchObject.Cells("Prop.VmId").Formula = '"' + $VsSwitch.VmId + '"'
	$VsSwitchObject.Cells("Prop.Nic").Formula = '"' + $VsSwitch.Nic + '"'
	$VsSwitchObject.Cells("Prop.SpecNumPorts").Formula = '"' + $VsSwitch.SpecNumPorts + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicySecurityAllowPromiscuous").Formula = '"' + $VsSwitch.SpecPolicySecurityAllowPromiscuous + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicySecurityMacChanges").Formula = '"' + $VsSwitch.SpecPolicySecurityMacChanges + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicySecurityForgedTransmits").Formula = '"' + $VsSwitch.SpecPolicySecurityForgedTransmits + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingPolicy").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingPolicy + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingReversePolicy").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingReversePolicy + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingNotifySwitches").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingNotifySwitches + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingRollingOrder").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingRollingOrder + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingNicOrderActiveNic").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingNicOrderActiveNic + '"'
	$VsSwitchObject.Cells("Prop.SpecPolicyNicTeamingNicOrderStandbyNic").Formula = '"' + $VsSwitch.SpecPolicyNicTeamingNicOrderStandbyNic + '"'
	$VsSwitchObject.Cells("Prop.NumPorts").Formula = '"' + $VsSwitch.NumPorts + '"'
	$VsSwitchObject.Cells("Prop.NumPortsAvailable").Formula = '"' + $VsSwitch.NumPortsAvailable + '"'
	$VsSwitchObject.Cells("Prop.Mtu").Formula = '"' + $VsSwitch.Mtu + '"'
	$VsSwitchObject.Cells("Prop.SpecBridgeBeacon").Formula = '"' + $VsSwitch.SpecBridgeBeacon + '"'
	$VsSwitchObject.Cells("Prop.SpecBridgeLinkDiscoveryProtocolConfig").Formula = '"' + $VsSwitch.SpecBridgeLinkDiscoveryProtocolConfig + '"'
	$VsSwitchObject.Cells("Prop.Id").Formula = '"' + $VsSwitch.Id + '"'
}
#endregion ~~< Draw_VsSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VssPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssPnic
{
	$VssPNICObject.Cells("Prop.Name").Formula = '"' + $VssPnic.Name + '"'
	$VssPNICObject.Cells("Prop.Datacenter").Formula = '"' + $VssPnic.Datacenter + '"'
	$VssPNICObject.Cells("Prop.DatacenterId").Formula = '"' + $VssPnic.DatacenterId + '"'
	$VssPNICObject.Cells("Prop.Cluster").Formula = '"' + $VssPnic.Cluster + '"'
	$VssPNICObject.Cells("Prop.ClusterId").Formula = '"' + $VssPnic.ClusterId + '"'
	$VssPNICObject.Cells("Prop.VmHost").Formula = '"' + $VssPnic.VmHost + '"'
	$VssPNICObject.Cells("Prop.VmHostId").Formula = '"' + $VssPnic.VmHostId + '"'
	$VssPNICObject.Cells("Prop.VsSwitch").Formula = '"' + $VssPnic.VsSwitch + '"'
	$VssPNICObject.Cells("Prop.VsSwitchId").Formula = '"' + $VssPnic.VsSwitchId + '"'
	$VssPNICObject.Cells("Prop.Mac").Formula = '"' + $VssPnic.Mac + '"'
	$VssPNICObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VssPnic.DhcpEnabled + '"'
	$VssPNICObject.Cells("Prop.IP").Formula = '"' + $VssPnic.IP + '"'
	$VssPNICObject.Cells("Prop.SubnetMask").Formula = '"' + $VssPnic.SubnetMask + '"'
	$VssPNICObject.Cells("Prop.BitRatePerSec").Formula = '"' + $VssPnic.BitRatePerSec + '"'
	$VssPNICObject.Cells("Prop.FullDuplex").Formula = '"' + $VssPnic.FullDuplex + '"'
	$VssPNICObject.Cells("Prop.PciId").Formula = '"' + $VssPnic.PciId + '"'
	$VssPNICObject.Cells("Prop.WakeOnLanSupported").Formula = '"' + $VssPnic.WakeOnLanSupported + '"'
	$VssPNICObject.Cells("Prop.Driver").Formula = '"' + $VssPnic.Driver + '"'
	$VssPNICObject.Cells("Prop.LinkSpeed").Formula = '"' + $VssPnic.LinkSpeed + '"'
	$VssPNICObject.Cells("Prop.SpecEnableEnhancedNetworkingStack").Formula = '"' + $VssPnic.SpecEnableEnhancedNetworkingStack + '"'
	$VssPNICObject.Cells("Prop.FcoeConfigurationPriorityClass").Formula = '"' + $VssPnic.FcoeConfigurationPriorityClass + '"'
	$VssPNICObject.Cells("Prop.FcoeConfigurationSourceMac").Formula = '"' + $VssPnic.FcoeConfigurationSourceMac + '"'
	$VssPNICObject.Cells("Prop.FcoeConfigurationVlanRange").Formula = '"' + $VssPnic.FcoeConfigurationVlanRange + '"'
	$VssPNICObject.Cells("Prop.FcoeConfigurationCapabilities").Formula = '"' + $VssPnic.FcoeConfigurationCapabilities + '"'
	$VssPNICObject.Cells("Prop.FcoeConfigurationFcoeActive").Formula = '"' + $VssPnic.FcoeConfigurationFcoeActive + '"'
	$VssPNICObject.Cells("Prop.VmDirectPathGen2Supported").Formula = '"' + $VssPnic.VmDirectPathGen2Supported + '"'
	$VssPNICObject.Cells("Prop.VmDirectPathGen2SupportedMode").Formula = '"' + $VssPnic.VmDirectPathGen2SupportedMode + '"'
	$VssPNICObject.Cells("Prop.ResourcePoolSchedulerAllowed").Formula = '"' + $VssPnic.ResourcePoolSchedulerAllowed + '"'
	$VssPNICObject.Cells("Prop.ResourcePoolSchedulerDisallowedReason").Formula = '"' + $VssPnic.ResourcePoolSchedulerDisallowedReason + '"'
	$VssPNICObject.Cells("Prop.AutoNegotiateSupported").Formula = '"' + $VssPnic.AutoNegotiateSupported + '"'
	$VssPNICObject.Cells("Prop.EnhancedNetworkingStackSupported").Formula = '"' + $VssPnic.EnhancedNetworkingStackSupported + '"'
}
#endregion ~~< Draw_VssPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VssPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssPort
{
	$VssPortObject.Cells("Prop.Name").Formula = '"' + $VssPort.Name + '"'
	$VssPortObject.Cells("Prop.Datacenter").Formula = '"' + $VssPort.Datacenter + '"'
	$VssPortObject.Cells("Prop.DatacenterId").Formula = '"' + $VssPort.DatacenterId + '"'
	$VssPortObject.Cells("Prop.Cluster").Formula = '"' + $VssPort.Cluster + '"'
	$VssPortObject.Cells("Prop.ClusterId").Formula = '"' + $VssPort.ClusterId + '"'
	$VssPortObject.Cells("Prop.VmHost").Formula = '"' + $VssPort.VmHost + '"'
	$VssPortObject.Cells("Prop.VmHostId").Formula = '"' + $VssPort.VmHostId + '"'
	$VssPortObject.Cells("Prop.Vm").Formula = '"' + $VssPort.Vm + '"'
	$VssPortObject.Cells("Prop.VmId").Formula = '"' + $VssPort.VmId + '"'
	$VssPortObject.Cells("Prop.VsSwitch").Formula = '"' + $VssPort.VsSwitch + '"'
	$VssPortObject.Cells("Prop.VsSwitchId").Formula = '"' + $VssPort.VsSwitchId + '"'
	$VssPortObject.Cells("Prop.VLanId").Formula = '"' + $VssPort.VLanId + '"'
	$VssPortObject.Cells("Prop.Security_AllowPromiscuous").Formula = '"' + $VssPort.Security_AllowPromiscuous + '"'
	$VssPortObject.Cells("Prop.Security_MacChanges").Formula = '"' + $VssPort.Security_MacChanges + '"'
	$VssPortObject.Cells("Prop.Security_ForgedTransmits").Formula = '"' + $VssPort.Security_ForgedTransmits + '"'
	$VssPortObject.Cells("Prop.NicTeaming_Policy").Formula = '"' + $VssPort.NicTeaming_Policy + '"'
	$VssPortObject.Cells("Prop.NicTeaming_ReversePolicy").Formula = '"' + $VssPort.NicTeaming_ReversePolicy + '"'
	$VssPortObject.Cells("Prop.NicTeaming_NotifySwitches").Formula = '"' + $VssPort.NicTeaming_NotifySwitches + '"'
	$VssPortObject.Cells("Prop.NicTeaming_RollingOrder").Formula = '"' + $VssPort.NicTeaming_RollingOrder + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_CheckSpeed").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_CheckSpeed + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_Speed").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_Speed + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_CheckDuplex").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_CheckDuplex + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_FullDuplex").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_FullDuplex + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_CheckErrorPercent").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_CheckErrorPercent + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_Percentage").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_Percentage + '"'
	$VssPortObject.Cells("Prop.NicTeaming_FailureCriteria_CheckBeacon").Formula = '"' + $VssPort.NicTeaming_FailureCriteria_CheckBeacon + '"'
	$VssPortObject.Cells("Prop.NicTeaming_NicOrder_ActiveNic").Formula = '"' + $VssPort.NicTeaming_NicOrder_ActiveNic + '"'
	$VssPortObject.Cells("Prop.NicTeaming_NicOrder_StandbyNic").Formula = '"' + $VssPort.NicTeaming_NicOrder_StandbyNic + '"'
	$VssPortObject.Cells("Prop.OffloadPolicy_CsumOffload").Formula = '"' + $VssPort.OffloadPolicy_CsumOffload + '"'
	$VssPortObject.Cells("Prop.OffloadPolicy_TcpSegmentation").Formula = '"' + $VssPort.OffloadPolicy_TcpSegmentation + '"'
	$VssPortObject.Cells("Prop.OffloadPolicy_ZeroCopyXmit").Formula = '"' + $VssPort.OffloadPolicy_ZeroCopyXmit + '"'
	$VssPortObject.Cells("Prop.ShapingPolicy_Enabled").Formula = '"' + $VssPort.ShapingPolicy_Enabled + '"'
	$VssPortObject.Cells("Prop.ShapingPolicy_AverageBandwidth").Formula = '"' + $VssPort.ShapingPolicy_AverageBandwidth + '"'
	$VssPortObject.Cells("Prop.ShapingPolicy_PeakBandwidth").Formula = '"' + $VssPort.ShapingPolicy_PeakBandwidth + '"'
	$VssPortObject.Cells("Prop.ShapingPolicy_BurstSize").Formula = '"' + $VssPort.ShapingPolicy_BurstSize + '"'
}
#endregion ~~< Draw_VssPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VssVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VssVmk
{
	$VssVmkNicObject.Cells("Prop.Name").Formula = '"' + $VssVmk.Name + '"'
	$VssVmkNicObject.Cells("Prop.Datacenter").Formula = '"' + $VssVmk.Datacenter + '"'
	$VssVmkNicObject.Cells("Prop.DatacenterId").Formula = '"' + $VssVmk.DatacenterId + '"'
	$VssVmkNicObject.Cells("Prop.Cluster").Formula = '"' + $VssVmk.Cluster + '"'
	$VssVmkNicObject.Cells("Prop.ClusterId").Formula = '"' + $VssVmk.ClusterId + '"'
	$VssVmkNicObject.Cells("Prop.VmHost").Formula = '"' + $VssVmk.VmHost + '"'
	$VssVmkNicObject.Cells("Prop.VmHostId").Formula = '"' + $VssVmk.VmHostId + '"'
	$VssVmkNicObject.Cells("Prop.VSwitch").Formula = '"' + $VssVmk.VSwitch + '"'
	$VssVmkNicObject.Cells("Prop.VSwitchId").Formula = '"' + $VssVmk.VSwitchId + '"'
	$VssVmkNicObject.Cells("Prop.PortGroupName").Formula = '"' + $VssVmk.PortGroupName + '"'
	$VssVmkNicObject.Cells("Prop.PortGroupId").Formula = '"' + $VssVmk.PortGroupId + '"'
	$VssVmkNicObject.Cells("Prop.VMotionEnabled").Formula = '"' + $VssVmk.VMotionEnabled + '"'
	$VssVmkNicObject.Cells("Prop.FaultToleranceLoggingEnabled").Formula = '"' + $VssVmk.FaultToleranceLoggingEnabled + '"'
	$VssVmkNicObject.Cells("Prop.ManagementTrafficEnabled").Formula = '"' + $VssVmk.ManagementTrafficEnabled + '"'
	$VssVmkNicObject.Cells("Prop.IP").Formula = '"' + $VssVmk.IP + '"'
	$VssVmkNicObject.Cells("Prop.Mac").Formula = '"' + $VssVmk.Mac + '"'
	$VssVmkNicObject.Cells("Prop.SubnetMask").Formula = '"' + $VssVmk.SubnetMask + '"'
	$VssVmkNicObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VssVmk.DhcpEnabled + '"'
	$VssVmkNicObject.Cells("Prop.IPv6").Formula = '"' + $VssVmk.IPv6 + '"'
	$VssVmkNicObject.Cells("Prop.AutomaticIPv6").Formula = '"' + $VssVmk.AutomaticIPv6 + '"'
	$VssVmkNicObject.Cells("Prop.IPv6ThroughDhcp").Formula = '"' + $VssVmk.IPv6ThroughDhcp + '"'
	$VssVmkNicObject.Cells("Prop.IPv6Enabled").Formula = '"' + $VssVmk.IPv6Enabled + '"'
	$VssVmkNicObject.Cells("Prop.VsanTrafficEnabled").Formula = '"' + $VssVmk.VsanTrafficEnabled + '"'
	$VssVmkNicObject.Cells("Prop.Mtu").Formula = '"' + $VssVmk.Mtu + '"'
	$VssVmkNicObject.Cells("Prop.SpecTsoEnabled").Formula = '"' + $VssVmk.SpecTsoEnabled + '"'
	$VssVmkNicObject.Cells("Prop.SpecNetStackInstanceKey").Formula = '"' + $VssVmk.SpecNetStackInstanceKey + '"'
	$VssVmkNicObject.Cells("Prop.SpecOpaqueNetwork").Formula = '"' + $VssVmk.SpecOpaqueNetwork + '"'
	$VssVmkNicObject.Cells("Prop.SpecExternalId").Formula = '"' + $VssVmk.SpecExternalId + '"'
	$VssVmkNicObject.Cells("Prop.SpecPinnedPnic").Formula = '"' + $VssVmk.SpecPinnedPnic + '"'
	$VssVmkNicObject.Cells("Prop.SpecIpRouteSpec").Formula = '"' + $VssVmk.SpecIpRouteSpec + '"'
}
#endregion ~~< Draw_VssVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdSwitch
{
	$VdSwitchObject.Cells("Prop.Name").Formula = '"' + $VdSwitch.Name + '"'
	$VdSwitchObject.Cells("Prop.Datacenter").Formula = '"' + $VdSwitch.Datacenter + '"'
	$VdSwitchObject.Cells("Prop.DatacenterId").Formula = '"' + $VdSwitch.DatacenterId + '"'
	$VdSwitchObject.Cells("Prop.Cluster").Formula = '"' + $VdSwitch.Cluster + '"'
	$VdSwitchObject.Cells("Prop.ClusterId").Formula = '"' + $VdSwitch.ClusterId + '"'
	$VdSwitchObject.Cells("Prop.VmHost").Formula = '"' + $VdSwitch.VmHost + '"'
	$VdSwitchObject.Cells("Prop.VmHostId").Formula = '"' + $VdSwitch.VmHostId + '"'
	$VdSwitchObject.Cells("Prop.PortgroupName").Formula = '"' + $VdSwitch.PortgroupName + '"'
	$VdSwitchObject.Cells("Prop.PortgroupId").Formula = '"' + $VdSwitch.PortgroupId + '"'
	$VdSwitchObject.Cells("Prop.NumHosts").Formula = '"' + $VdSwitch.NumHosts + '"'
	$VdSwitchObject.Cells("Prop.NumPorts").Formula = '"' + $VdSwitch.NumPorts + '"'
	$VdSwitchObject.Cells("Prop.Vendor").Formula = '"' + $VdSwitch.Vendor + '"'
	$VdSwitchObject.Cells("Prop.Version").Formula = '"' + $VdSwitch.Version + '"'
	$VdSwitchObject.Cells("Prop.ConfigVspanSession").Formula = '"' + $VdSwitch.ConfigVspanSession + '"'
	$VdSwitchObject.Cells("Prop.ConfigPvlanConfig").Formula = '"' + $VdSwitch.ConfigPvlanConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigMaxMtu").Formula = '"' + $VdSwitch.ConfigMaxMtu + '"'
	$VdSwitchObject.Cells("Prop.ConfigLinkDiscoveryProtocolConfig").Formula = '"' + $VdSwitch.ConfigLinkDiscoveryProtocolConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigCollectorIpAddress").Formula = '"' + $VdSwitch.ConfigIpfixConfigCollectorIpAddress + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigCollectorPort").Formula = '"' + $VdSwitch.ConfigIpfixConfigCollectorPort + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigObservationDomainId").Formula = '"' + $VdSwitch.ConfigIpfixConfigObservationDomainId + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigActiveFlowTimeout").Formula = '"' + $VdSwitch.ConfigIpfixConfigActiveFlowTimeout + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigIdleFlowTimeout").Formula = '"' + $VdSwitch.ConfigIpfixConfigIdleFlowTimeout + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigSamplingRate").Formula = '"' + $VdSwitch.ConfigIpfixConfigSamplingRate + '"'
	$VdSwitchObject.Cells("Prop.ConfigIpfixConfigInternalFlowsOnly").Formula = '"' + $VdSwitch.ConfigIpfixConfigInternalFlowsOnly + '"'
	$VdSwitchObject.Cells("Prop.ConfigLacpGroupConfig").Formula = '"' + $VdSwitch.ConfigLacpGroupConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigLacpApiVersion").Formula = '"' + $VdSwitch.ConfigLacpApiVersion + '"'
	$VdSwitchObject.Cells("Prop.ConfigMulticastFilteringMode").Formula = '"' + $VdSwitch.ConfigMulticastFilteringMode + '"'
	$VdSwitchObject.Cells("Prop.ConfigNumStandalonePorts").Formula = '"' + $VdSwitch.ConfigNumStandalonePorts + '"'
	$VdSwitchObject.Cells("Prop.ConfigNumPorts").Formula = '"' + $VdSwitch.ConfigNumPorts + '"'
	$VdSwitchObject.Cells("Prop.ConfigMaxPorts").Formula = '"' + $VdSwitch.ConfigMaxPorts + '"'
	$VdSwitchObject.Cells("Prop.ConfigNumUplinkPorts").Formula = '"' + $VdSwitch.ConfigNumUplinkPorts + '"'
	$VdSwitchObject.Cells("Prop.ConfigUplinkPortName").Formula = '"' + $VdSwitch.ConfigUplinkPortName + '"'
	$VdSwitchObject.Cells("Prop.ConfigUplinkPortgroup").Formula = '"' + $VdSwitch.ConfigUplinkPortgroup + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigVlan").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigVlan + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigQosTag").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigQosTag + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyPolicy").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyPolicy + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyReversePolicy").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyReversePolicy + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyNotifySwitches").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyNotifySwitches + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyRollingOrder").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyRollingOrder + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckSpeed").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckSpeed + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaSpeed").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaSpeed + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckDuplex").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckDuplex + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaFullDuplex").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaFullDuplex + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckErrorPercent").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckErrorPercent + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaPercentage").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaPercentage + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckBeacon").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaCheckBeacon + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyFailureCriteriaInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigSecurityPolicyAllowPromiscuous").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigSecurityPolicyAllowPromiscuous + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigSecurityPolicyMacChanges").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigSecurityPolicyMacChanges + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigSecurityPolicyForgedTransmits").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigSecurityPolicyForgedTransmits + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigSecurityPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigSecurityPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigUplinkTeamingPolicyUplinkPortOrder").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigUplinkTeamingPolicyUplinkPortOrder + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigIpfixEnabled").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigIpfixEnabled + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigTxUplink").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigTxUplink + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigLacpPolicyEnable").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigLacpPolicyEnable + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigLacpPolicyMode").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigLacpPolicyMode + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigLacpPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigLacpPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyAllowPromiscuous").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyAllowPromiscuous + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacChanges").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacChanges + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyForgedTransmits").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyForgedTransmits + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyEnabled").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyEnabled + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyAllowUnicastFlooding").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyAllowUnicastFlooding + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimit").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimit + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimitPolicy").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyLimitPolicy + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyMacLearningPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigMacManagementPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigMacManagementPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigBlocked").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigBlocked + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigVmDirectPathGen2Allowed").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigVmDirectPathGen2Allowed + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigInShapingPolicyEnabled").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigInShapingPolicyEnabled + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigInShapingPolicyAverageBandwidth").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigInShapingPolicyAverageBandwidth + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigInShapingPolicyPeakBandwidth").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigInShapingPolicyPeakBandwidth + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigInShapingPolicyBurstSize").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigInShapingPolicyBurstSize + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigInShapingPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigInShapingPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigOutShapingPolicyEnabled").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigOutShapingPolicyEnabled + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigOutShapingPolicyAverageBandwidth").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigOutShapingPolicyAverageBandwidth + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigOutShapingPolicyPeakBandwidth").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigOutShapingPolicyPeakBandwidth + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigOutShapingPolicyBurstSize").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigOutShapingPolicyBurstSize + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigOutShapingPolicyInherited").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigOutShapingPolicyInherited + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigVendorSpecificConfig").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigVendorSpecificConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigNetworkResourcePoolKey").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigNetworkResourcePoolKey + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultPortConfigFilterPolicy").Formula = '"' + $VdSwitch.ConfigDefaultPortConfigFilterPolicy + '"'
	$VdSwitchObject.Cells("Prop.ConfigPolicyAutoPreInstallAllowed").Formula = '"' + $VdSwitch.ConfigPolicyAutoPreInstallAllowed + '"'
	$VdSwitchObject.Cells("Prop.ConfigPolicyAutoUpgradeAllowed").Formula = '"' + $VdSwitch.ConfigPolicyAutoUpgradeAllowed + '"'
	$VdSwitchObject.Cells("Prop.ConfigPolicyPartialUpgradeAllowed").Formula = '"' + $VdSwitch.ConfigPolicyPartialUpgradeAllowed + '"'
	$VdSwitchObject.Cells("Prop.ConfigSwitchIpAddress").Formula = '"' + $VdSwitch.ConfigSwitchIpAddress + '"'
	$VdSwitchObject.Cells("Prop.ConfigCreateTime").Formula = '"' + $VdSwitch.ConfigCreateTime + '"'
	$VdSwitchObject.Cells("Prop.ConfigNetworkResourceManagementEnabled").Formula = '"' + $VdSwitch.ConfigNetworkResourceManagementEnabled + '"'
	$VdSwitchObject.Cells("Prop.ConfigDefaultProxySwitchMaxNumPorts").Formula = '"' + $VdSwitch.ConfigDefaultProxySwitchMaxNumPorts + '"'
	$VdSwitchObject.Cells("Prop.ConfigHealthCheckConfig").Formula = '"' + $VdSwitch.ConfigHealthCheckConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigInfrastructureTrafficResourceConfig").Formula = '"' + $VdSwitch.ConfigInfrastructureTrafficResourceConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigNetResourcePoolTrafficResourceConfig").Formula = '"' + $VdSwitch.ConfigNetResourcePoolTrafficResourceConfig + '"'
	$VdSwitchObject.Cells("Prop.ConfigNetworkResourceControlVersion").Formula = '"' + $VdSwitch.ConfigNetworkResourceControlVersion + '"'
	$VdSwitchObject.Cells("Prop.ConfigVmVnicNetworkResourcePool").Formula = '"' + $VdSwitch.ConfigVmVnicNetworkResourcePool + '"'
	$VdSwitchObject.Cells("Prop.ConfigPnicCapacityRatioForReservation").Formula = '"' + $VdSwitch.ConfigPnicCapacityRatioForReservation + '"'
	$VdSwitchObject.Cells("Prop.RuntimeHostMemberRuntime").Formula = '"' + $VdSwitch.RuntimeHostMemberRuntime + '"'
	$VdSwitchObject.Cells("Prop.OverallStatus").Formula = '"' + $VdSwitch.OverallStatus + '"'
	$VdSwitchObject.Cells("Prop.ConfigStatus").Formula = '"' + $VdSwitch.ConfigStatus + '"'
	$VdSwitchObject.Cells("Prop.AlarmActionsEnabled").Formula = '"' + $VdSwitch.AlarmActionsEnabled + '"'
	$VdSwitchObject.Cells("Prop.Mtu").Formula = '"' + $VdSwitch.Mtu + '"'
	$VdSwitchObject.Cells("Prop.MoRef").Formula = '"' + $VdSwitch.MoRef + '"'
}
#endregion ~~< Draw_VdSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VdsPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsPnic
{
	$VdsPNICObject.Cells("Prop.Name").Formula = '"' + $VdsPnic.Name + '"'
	$VdsPNICObject.Cells("Prop.Datacenter").Formula = '"' + $VdsPnic.Datacenter + '"'
	$VdsPNICObject.Cells("Prop.DatacenterId").Formula = '"' + $VdsPnic.DatacenterId + '"'
	$VdsPNICObject.Cells("Prop.Cluster").Formula = '"' + $VdsPnic.Cluster + '"'
	$VdsPNICObject.Cells("Prop.ClusterId").Formula = '"' + $VdsPnic.ClusterId + '"'
	$VdsPNICObject.Cells("Prop.VmHost").Formula = '"' + $VdsPnic.VmHost + '"'
	$VdsPNICObject.Cells("Prop.VmHostId").Formula = '"' + $VdsPnic.VmHostId + '"'
	$VdsPNICObject.Cells("Prop.VdSwitch").Formula = '"' + $VdsPnic.VdSwitch + '"'
	$VdsPNICObject.Cells("Prop.VdSwitchId").Formula = '"' + $VdsPnic.VdSwitchId + '"'
	$VdsPNICObject.Cells("Prop.Mac").Formula = '"' + $VdsPnic.Mac + '"'
	$VdsPNICObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VdsPnic.DhcpEnabled + '"'
	$VdsPNICObject.Cells("Prop.IP").Formula = '"' + $VdsPnic.IP + '"'
	$VdsPNICObject.Cells("Prop.SubnetMask").Formula = '"' + $VdsPnic.SubnetMask + '"'
	$VdsPNICObject.Cells("Prop.Portgroup").Formula = '"' + $VdsPnic.Portgroup + '"'
	$VdsPNICObject.Cells("Prop.ConnectedEntity").Formula = '"' + $VdsPnic.ConnectedEntity + '"'
	$VdsPNICObject.Cells("Prop.VlanConfiguration").Formula = '"' + $VdsPnic.VlanConfiguration + '"'
	$VdsPNICObject.Cells("Prop.BitRatePerSec").Formula = '"' + $VdsPnic.BitRatePerSec + '"'
	$VdsPNICObject.Cells("Prop.FullDuplex").Formula = '"' + $VdsPnic.FullDuplex + '"'
	$VdsPNICObject.Cells("Prop.PciId").Formula = '"' + $VdsPnic.PciId + '"'
	$VdsPNICObject.Cells("Prop.WakeOnLanSupported").Formula = '"' + $VdsPnic.WakeOnLanSupported + '"'
	$VdsPNICObject.Cells("Prop.Driver").Formula = '"' + $VdsPnic.Driver + '"'
	$VdsPNICObject.Cells("Prop.LinkSpeed").Formula = '"' + $VdsPnic.LinkSpeed + '"'
	$VdsPNICObject.Cells("Prop.SpecEnableEnhancedNetworkingStack").Formula = '"' + $VdsPnic.SpecEnableEnhancedNetworkingStack + '"'
	$VdsPNICObject.Cells("Prop.FcoeConfigurationPriorityClass").Formula = '"' + $VdsPnic.FcoeConfigurationPriorityClass + '"'
	$VdsPNICObject.Cells("Prop.FcoeConfigurationSourceMac").Formula = '"' + $VdsPnic.FcoeConfigurationSourceMac + '"'
	$VdsPNICObject.Cells("Prop.FcoeConfigurationVlanRange").Formula = '"' + $VdsPnic.FcoeConfigurationVlanRange + '"'
	$VdsPNICObject.Cells("Prop.FcoeConfigurationCapabilities").Formula = '"' + $VdsPnic.FcoeConfigurationCapabilities + '"'
	$VdsPNICObject.Cells("Prop.FcoeConfigurationFcoeActive").Formula = '"' + $VdsPnic.FcoeConfigurationFcoeActive + '"'
	$VdsPNICObject.Cells("Prop.VmDirectPathGen2Supported").Formula = '"' + $VdsPnic.VmDirectPathGen2Supported + '"'
	$VdsPNICObject.Cells("Prop.VmDirectPathGen2SupportedMode").Formula = '"' + $VdsPnic.VmDirectPathGen2SupportedMode + '"'
	$VdsPNICObject.Cells("Prop.ResourcePoolSchedulerAllowed").Formula = '"' + $VdsPnic.ResourcePoolSchedulerAllowed + '"'
	$VdsPNICObject.Cells("Prop.ResourcePoolSchedulerDisallowedReason").Formula = '"' + $VdsPnic.ResourcePoolSchedulerDisallowedReason + '"'
	$VdsPNICObject.Cells("Prop.AutoNegotiateSupported").Formula = '"' + $VdsPnic.AutoNegotiateSupported + '"'
	$VdsPNICObject.Cells("Prop.EnhancedNetworkingStackSupported").Formula = '"' + $VdsPnic.EnhancedNetworkingStackSupported + '"'
	$VdsPNICObject.Cells("Prop.MoRef").Formula = '"' + $VdsPnic.MoRef + '"'
}
#endregion ~~< Draw_VdsPnic >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VdsPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsPort
{
	$VdsPortObject.Cells("Prop.Name").Formula = '"' + $VdsPort.Name + '"'
	$VdsPortObject.Cells("Prop.Datacenter").Formula = '"' + $VdsPort.Datacenter + '"'
	$VdsPortObject.Cells("Prop.DatacenterId").Formula = '"' + $VdsPort.DatacenterId + '"'
	$VdsPortObject.Cells("Prop.Cluster").Formula = '"' + $VdsPort.Cluster + '"'
	$VdsPortObject.Cells("Prop.ClusterId").Formula = '"' + $VdsPort.ClusterId + '"'
	$VdsPortObject.Cells("Prop.VmHost").Formula = '"' + $VdsPort.VmHost + '"'
	$VdsPortObject.Cells("Prop.VmHostId").Formula = '"' + $VdsPort.VmHostId + '"'
	$VdsPortObject.Cells("Prop.Vm").Formula = '"' + $VdsPort.Vm + '"'
	$VdsPortObject.Cells("Prop.VmId").Formula = '"' + $VdsPort.VmId + '"'
	$VdsPortObject.Cells("Prop.VdSwitch").Formula = '"' + $VdsPort.VdSwitch + '"'
	$VdsPortObject.Cells("Prop.VdSwitchId").Formula = '"' + $VdsPort.VdSwitchId + '"'
	$VdsPortObject.Cells("Prop.VlanConfiguration").Formula = '"' + $VdsPort.VlanConfiguration + '"'
	$VdsPortObject.Cells("Prop.NumPorts").Formula = '"' + $VdsPort.NumPorts + '"'
	$VdsPortObject.Cells("Prop.ActiveUplinkPort").Formula = '"' + $VdsPort.ActiveUplinkPort + '"'
	$VdsPortObject.Cells("Prop.StandbyUplinkPort").Formula = '"' + $VdsPort.StandbyUplinkPort + '"'
	$VdsPortObject.Cells("Prop.Policy").Formula = '"' + $VdsPort.Policy + '"'
	$VdsPortObject.Cells("Prop.ReversePolicy").Formula = '"' + $VdsPort.ReversePolicy + '"'
	$VdsPortObject.Cells("Prop.NotifySwitches").Formula = '"' + $VdsPort.NotifySwitches + '"'
	$VdsPortObject.Cells("Prop.PortBinding").Formula = '"' + $VdsPort.PortBinding + '"'
	$VdsPortObject.Cells("Prop.MoRef").Formula = '"' + $VdsPort.MoRef + '"'
}
#endregion ~~< Draw_VdsPort >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_VdsVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_VdsVmk
{
	$VdsVmkNicObject.Cells("Prop.Name").Formula = '"' + $VdsVmk.Name + '"'
	$VdsVmkNicObject.Cells("Prop.Datacenter").Formula = '"' + $VdsVmk.Datacenter + '"'
	$VdsVmkNicObject.Cells("Prop.DatacenterId").Formula = '"' + $VdsVmk.DatacenterId + '"'
	$VdsVmkNicObject.Cells("Prop.Cluster").Formula = '"' + $VdsVmk.Cluster + '"'
	$VdsVmkNicObject.Cells("Prop.ClusterId").Formula = '"' + $VdsVmk.ClusterId + '"'
	$VdsVmkNicObject.Cells("Prop.VmHost").Formula = '"' + $VdsVmk.VmHost + '"'
	$VdsVmkNicObject.Cells("Prop.VmHostId").Formula = '"' + $VdsVmk.VmHostId + '"'
	$VdsVmkNicObject.Cells("Prop.VSwitch").Formula = '"' + $VdsVmk.VSwitch + '"'
	$VdsVmkNicObject.Cells("Prop.VSwitchId").Formula = '"' + $VdsVmk.VSwitchId + '"'
	$VdsVmkNicObject.Cells("Prop.PortGroupName").Formula = '"' + $VdsVmk.PortGroupName + '"'
	$VdsVmkNicObject.Cells("Prop.PortGroupId").Formula = '"' + $VdsVmk.PortGroupId + '"'
	$VdsVmkNicObject.Cells("Prop.DhcpEnabled").Formula = '"' + $VdsVmk.DhcpEnabled + '"'
	$VdsVmkNicObject.Cells("Prop.IP").Formula = '"' + $VdsVmk.IP + '"'
	$VdsVmkNicObject.Cells("Prop.Mac").Formula = '"' + $VdsVmk.Mac + '"'
	$VdsVmkNicObject.Cells("Prop.ManagementTrafficEnabled").Formula = '"' + $VdsVmk.ManagementTrafficEnabled + '"'
	$VdsVmkNicObject.Cells("Prop.VMotionEnabled").Formula = '"' + $VdsVmk.VMotionEnabled + '"'
	$VdsVmkNicObject.Cells("Prop.FaultToleranceLoggingEnabled").Formula = '"' + $VdsVmk.FaultToleranceLoggingEnabled + '"'
	$VdsVmkNicObject.Cells("Prop.VsanTrafficEnabled").Formula = '"' + $VdsVmk.VsanTrafficEnabled + '"'
	$VdsVmkNicObject.Cells("Prop.Mtu").Formula = '"' + $VdsVmk.Mtu + '"'
}
#endregion ~~< Draw_VdsVmk >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_DrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsRule
{
	$DRSObject.Cells("Prop.Name").Formula = '"' + $DRSRule.Name + '"'
	$DRSObject.Cells("Prop.Datacenter").Formula = '"' + $DRSRule.Datacenter + '"'
	$DRSObject.Cells("Prop.DatacenterId").Formula = '"' + $DRSRule.DatacenterId + '"'
	$DRSObject.Cells("Prop.Cluster").Formula = '"' + $DRSRule.Cluster + '"'
	$DRSObject.Cells("Prop.ClusterId").Formula = '"' + $DRSRule.ClusterId + '"'
	$DRSObject.Cells("Prop.Vm").Formula = '"' + $DRSRule.Vm + '"'
	$DRSObject.Cells("Prop.VmId").Formula = '"' + $DRSRule.VmId + '"'
	$DRSObject.Cells("Prop.Type").Formula = '"' + $DRSRule.Type + '"'
	$DRSObject.Cells("Prop.Enabled").Formula = '"' + $DRSRule.Enabled + '"'
	$DRSObject.Cells("Prop.Mandatory").Formula = '"' + $DRSRule.Mandatory + '"'
}
#endregion ~~< Draw_DrsRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_DrsVmHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsVmHostRule
{
	$DRSVMHostRuleObject.Cells("Prop.Name").Formula = '"' + $DrsVmHostRule.Name + '"'
	$DRSVMHostRuleObject.Cells("Prop.Datacenter").Formula = '"' + $DrsVmHostRule.Datacenter + '"'
	$DRSVMHostRuleObject.Cells("Prop.DatacenterId").Formula = '"' + $DrsVmHostRule.DatacenterId + '"'
	$DRSVMHostRuleObject.Cells("Prop.Cluster").Formula = '"' + $DrsVmHostRule.Cluster + '"'
	$DRSVMHostRuleObject.Cells("Prop.ClusterId").Formula = '"' + $DrsVmHostRule.ClusterId + '"'
	$DRSVMHostRuleObject.Cells("Prop.Enabled").Formula = '"' + $DrsVmHostRule.Enabled + '"'
	$DRSVMHostRuleObject.Cells("Prop.Type").Formula = '"' + $DrsVmHostRule.Type + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMGroup").Formula = '"' + $DrsVmHostRule.VMGroup + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMGroupMember").Formula = '"' + $DrsVmHostRule.VMGroupMember + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMGroupMemberId").Formula = '"' + $DrsVmHostRule.VMGroupMemberId + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMHostGroup").Formula = '"' + $DrsVmHostRule.VMHostGroup + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMHostGroupMember").Formula = '"' + $DrsVmHostRule.VMHostGroupMember + '"'
	$DRSVMHostRuleObject.Cells("Prop.VMHostGroupMemberId").Formula = '"' + $DrsVmHostRule.VMHostGroupMemberId + '"'
	$DRSVMHostRuleObject.Cells("Prop.AffineHostGroupName").Formula = '"' + $DrsVmHostRule.AffineHostGroupName + '"'
	$DRSVMHostRuleObject.Cells("Prop.AntiAffineHostGroupName").Formula = '"' + $DrsVmHostRule.AntiAffineHostGroupName + '"'
}
#endregion ~~< Draw_DrsVmHostRule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_DrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_DrsClusterGroup
{
	$DrsClusterGroupObject.Cells("Prop.Name").Formula = '"' + $DrsClusterGroup.Name + '"'
	$DrsClusterGroupObject.Cells("Prop.Datacenter").Formula = '"' + $DrsClusterGroup.Datacenter + '"'
	$DrsClusterGroupObject.Cells("Prop.DatacenterId").Formula = '"' + $DrsClusterGroup.DatacenterId + '"'
	$DrsClusterGroupObject.Cells("Prop.Cluster").Formula = '"' + $DrsClusterGroup.Cluster + '"'
	$DrsClusterGroupObject.Cells("Prop.ClusterId").Formula = '"' + $DrsClusterGroup.ClusterId + '"'
	$DrsClusterGroupObject.Cells("Prop.GroupType").Formula = '"' + $DrsClusterGroup.GroupType + '"'
	$DrsClusterGroupObject.Cells("Prop.Member").Formula = '"' + $DrsClusterGroup.Member + '"'
	$DrsClusterGroupObject.Cells("Prop.MemberId").Formula = '"' + $DrsClusterGroup.MemberId + '"'
}
#endregion ~~< Draw_DrsClusterGroup >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_ParentSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ParentSnapshot
{
	$ParentSnapshotObject.Cells("Prop.VM").Formula = '"' + $ParentSnapshot.VM + '"'
	$ParentSnapshotObject.Cells("Prop.VMId").Formula = '"' + $ParentSnapshot.VMId + '"'
	$ParentSnapshotObject.Cells("Prop.Name").Formula = '"' + $ParentSnapshot.Name + '"'
	$ParentSnapshotObject.Cells("Prop.Created").Formula = '"' + $ParentSnapshot.Created + '"'
	$ParentSnapshotObject.Cells("Prop.Id").Formula = '"' + $ParentSnapshot.Id + '"'
	$ParentSnapshotObject.Cells("Prop.Children").Formula = '"' + $ParentSnapshot.Children + '"'
	$ParentSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ParentSnapshot.ParentSnapshot + '"'
	$ParentSnapshotObject.Cells("Prop.ParentSnapshotId").Formula = '"' + $ParentSnapshot.ParentSnapshotId + '"'
	$ParentSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ParentSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ParentSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_ChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildSnapshot
{
	$ChildSnapshotObject.Cells("Prop.VM").Formula = '"' + $ChildSnapshot.VM + '"'
	$ChildSnapshotObject.Cells("Prop.VMId").Formula = '"' + $ChildSnapshot.VMId + '"'
	$ChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildSnapshot.Name + '"'
	$ChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildSnapshot.Created + '"'
	$ChildSnapshotObject.Cells("Prop.Id").Formula = '"' + $ChildSnapshot.Id + '"'
	$ChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildSnapshot.Children + '"'
	$ChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildSnapshot.ParentSnapshot + '"'
	$ChildSnapshotObject.Cells("Prop.ParentSnapshotId").Formula = '"' + $ChildSnapshot.ParentSnapshotId + '"'
	$ChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_ChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildChildSnapshot
{
	$ChildChildSnapshotObject.Cells("Prop.VM").Formula = '"' + $ChildChildSnapshot.VM + '"'
	$ChildChildSnapshotObject.Cells("Prop.VMId").Formula = '"' + $ChildChildSnapshot.VMId + '"'
	$ChildChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildChildSnapshot.Name + '"'
	$ChildChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildChildSnapshot.Created + '"'
	$ChildChildSnapshotObject.Cells("Prop.Id").Formula = '"' + $ChildChildSnapshot.Id + '"'
	$ChildChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildChildSnapshot.Children + '"'
	$ChildChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildChildSnapshot.ParentSnapshot + '"'
	$ChildChildSnapshotObject.Cells("Prop.ParentSnapshotId").Formula = '"' + $ChildChildSnapshot.ParentSnapshotId + '"'
	$ChildChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_ChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_Draw_ChildChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_ChildChildChildSnapshot
{
	$ChildChildChildSnapshotObject.Cells("Prop.VM").Formula = '"' + $ChildChildChildSnapshot.VM + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.VMId").Formula = '"' + $ChildChildChildSnapshot.VMId + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.Name").Formula = '"' + $ChildChildChildSnapshot.Name + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.Created").Formula = '"' + $ChildChildChildSnapshot.Created + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.Id").Formula = '"' + $ChildChildChildSnapshot.Id + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.Children").Formula = '"' + $ChildChildChildSnapshot.Children + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.ParentSnapshot").Formula = '"' + $ChildChildChildSnapshot.ParentSnapshot + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.ParentSnapshotId").Formula = '"' + $ChildChildChildSnapshot.ParentSnapshotId + '"'
	$ChildChildChildSnapshotObject.Cells("Prop.IsCurrent").Formula = '"' + $ChildChildChildSnapshot.IsCurrent + '"'
}
#endregion ~~< Draw_Draw_ChildChildChildSnapshot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Draw_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Draw_LinkedvCenter
{
	$LinkedvCenterObject.Cells("Prop.Name").Formula = '"' + $LinkedvCenter.Name + '"'
	$LinkedvCenterObject.Cells("Prop.Version").Formula = '"' + $LinkedvCenter.Version + '"'
	$LinkedvCenterObject.Cells("Prop.Build").Formula = '"' + $LinkedvCenter.Build + '"'
	$LinkedvCenterObject.Cells("Prop.OsType").Formula = '"' + $LinkedvCenter.OsType + '"'
}
#endregion ~~< Draw_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Draw Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CSV Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function CSV_In_Out
{
	$global:DrawCsvFolder = $DrawCsvBrowse.SelectedPath
	# vCenter
	$global:vCenterExportFile = "$DrawCsvFolder\$vCenter-vCenterExport.csv"
	if(Test-Path $vCenterExportFile) `
	{ `
		$global:vCenterImport = Import-Csv $vCenterExportFile `
	} `
	# Datacenter
	$global:DatacenterExportFile = "$DrawCsvFolder\$vCenter-DatacenterExport.csv"
	if(Test-Path $DatacenterExportFile) `
	{ `
		$global:DatacenterImport = Import-Csv $DatacenterExportFile `
	} `
	# Cluster
	$global:ClusterExportFile = "$DrawCsvFolder\$vCenter-ClusterExport.csv"
	if(Test-Path $ClusterExportFile) `
	{ `
		$global:ClusterImport = Import-Csv $ClusterExportFile `
	} `
	# VmHost
	$global:VmHostExportFile = "$DrawCsvFolder\$vCenter-VmHostExport.csv"
	if(Test-Path $VmHostExportFile) `
	{ `
		$global:VmHostImport = Import-Csv $VmHostExportFile `
	} `
	# Vm
	$global:VmExportFile = "$DrawCsvFolder\$vCenter-VmExport.csv"
	if(Test-Path $VmExportFile) `
	{ `
		$global:VmImport = Import-Csv $VmExportFile `
	} `
	#Template
	$global:TemplateExportFile = "$DrawCsvFolder\$vCenter-TemplateExport.csv"
	if(Test-Path $TemplateExportFile) `
	{ `
		$global:TemplateImport = Import-Csv $TemplateExportFile `
	} `
	# Folder
	$global:FolderExportFile = "$DrawCsvFolder\$vCenter-FolderExport.csv"
	if(Test-Path $FolderExportFile) `
	{ `
		$global:FolderImport = Import-Csv $FolderExportFile `
	} `
	# Datastore Cluster
	$global:DatastoreClusterExportFile = "$DrawCsvFolder\$vCenter-DatastoreClusterExport.csv"
	if(Test-Path $DatastoreClusterExportFile) `
	{ `
		$global:DatastoreClusterImport = Import-Csv $DatastoreClusterExportFile `
	} `
	# Datastore
	$global:DatastoreExportFile = "$DrawCsvFolder\$vCenter-DatastoreExport.csv"
	if(Test-Path $DatastoreExportFile) `
	{ `
		$global:DatastoreImport = Import-Csv $DatastoreExportFile `
	} `
	# RDM's
	$global:RdmExportFile = "$DrawCsvFolder\$vCenter-RdmExport.csv"
	if(Test-Path $RdmExportFile) `
	{ `
		$global:RdmImport = Import-Csv $RdmExportFile `
	} `
	# ResourcePool
	$global:ResourcePoolExportFile = "$DrawCsvFolder\$vCenter-ResourcePoolExport.csv"
	if(Test-Path $ResourcePoolExportFile) `
	{ `
		$global:ResourcePoolImport = Import-Csv $ResourcePoolExportFile `
	} `
	# Vss Switch
	$global:VsSwitchExportFile = "$DrawCsvFolder\$vCenter-VsSwitchExport.csv"
	if(Test-Path $VsSwitchExportFile) `
	{ `
		$global:VsSwitchImport = Import-Csv $VsSwitchExportFile `
	} `
	# Vss Port Group
	$global:VssPortExportFile = "$DrawCsvFolder\$vCenter-VssPortGroupExport.csv"
	if(Test-Path $VssPortExportFile) `
	{ `
		$global:VssPortImport = Import-Csv $VssPortExportFile `
	} `
	# Vss VMKernel
	$global:VssVmkExportFile = "$DrawCsvFolder\$vCenter-VssVmkernelExport.csv"
	if(Test-Path $VssVmkExportFile) `
	{ `
		$global:VssVmkImport = Import-Csv $VssVmkExportFile `
	} `
	# Vss Pnic
	$global:VssPnicExportFile = "$DrawCsvFolder\$vCenter-VssPnicExport.csv"
	if(Test-Path $VssPnicExportFile) `
	{ `
		$global:VssPnicImport = Import-Csv $VssPnicExportFile `
	} `
	# Vds Switch
	$global:VdSwitchExportFile = "$DrawCsvFolder\$vCenter-VdSwitchExport.csv"
	if(Test-Path $VdSwitchExportFile) `
	{ `
		$global:VdSwitchImport = Import-Csv $VdSwitchExportFile `
	} `
	# Vds Port Group
	$global:VdsPortExportFile = "$DrawCsvFolder\$vCenter-VdsPortGroupExport.csv"
	if(Test-Path $VdsPortExportFile) `
	{ `
		$global:VdsPortImport = Import-Csv $VdsPortExportFile `
	} `
	# Vds VMKernel
	$global:VdsVmkExportFile = "$DrawCsvFolder\$vCenter-VdsVmkernelExport.csv"
	if(Test-Path $VdsVmkExportFile) `
	{ `
		$global:VdsVmkImport = Import-Csv $VdsVmkExportFile `
	} `
	# Vds Pnic
	$global:VdsPnicExportFile = "$DrawCsvFolder\$vCenter-VdsPnicExport.csv"
	if(Test-Path $VdsPnicExportFile) `
	{ `
		$global:VdsPnicImport = Import-Csv $VdsPnicExportFile `
	} `
	# DRS Rule
	$global:DrsRuleExportFile = "$DrawCsvFolder\$vCenter-DrsRuleExport.csv"
	if(Test-Path $DrsRuleExportFile) `
	{ `
		$global:DrsRuleImport = Import-Csv $DrsRuleExportFile `
	} `
	# DRS Cluster Group
	$global:DrsClusterGroupExportFile = "$DrawCsvFolder\$vCenter-DrsClusterGroupExport.csv"
	if(Test-Path $DrsClusterGroupExportFile) `
	{ `
		$global:DrsClusterGroupImport = Import-Csv $DrsClusterGroupExportFile `
	} `
	# DRS VmHost Rule
	$global:DrsVmHostRuleExportFile = "$DrawCsvFolder\$vCenter-DrsVmHostRuleExport.csv"
	if(Test-Path $DrsVmHostRuleExportFile) `
	{ `
		$global:DrsVmHostImport = Import-Csv $DrsVmHostRuleExportFile `
	} `
	# Snapshot
	$global:SnapshotExportFile = "$DrawCsvFolder\$vCenter-SnapshotExport.csv"
	if(Test-Path $SnapshotExportFile) `
	{ `
		$global:SnapshotImport = Import-Csv $SnapshotExportFile `
	} `
	# Linked vCenter
	$global:LinkedvCenterExportFile = "$DrawCsvFolder\$vCenter-LinkedvCenterExport.csv"
	if(Test-Path $LinkedvCenterExportFile) `
	{ `
		$global:LinkedvCenterImport = Import-Csv $LinkedvCenterExportFile `
	} `
}
#endregion ~~< CSV_In_Out >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< CSV Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Shapes Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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
	$global:VssPortGroupObj = $stnObj.Masters.Item("VSS Port Group")
	# VSSVMK NIC Object
	$global:VssVmkNicObj = $stnObj.Masters.Item("VSS VMKernel")
	# VDS Object
	$global:VDSObj = $stnObj.Masters.Item("VDS")
	# VDS PNIC Object
	$global:VdsPNICObj = $stnObj.Masters.Item("VDS Physical NIC")
	# VDSNIC Object
	$global:VdsPortGroupObj = $stnObj.Masters.Item("VDS Port Group")
	# VDSVMK NIC Object
	$global:VdsVmkNicObj = $stnObj.Masters.Item("VDS VMKernel")
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

#endregion ~~< Shapes Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Create_Visio_Base >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Create_Visio_Base
{
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Visio default page and loading shapes." -ForegroundColor Green
	$global:vCenter = $VcenterTextBox.Text
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
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
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] vCenter to Linked vCenter Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating vCenter to Linked vCenter Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "vCenter to Linked vCenters"
	$DocsObj.Pages('vCenter to Linked vCenters')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('vCenter to Linked vCenters')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$LinkedvCenterNumber = 0
	$LinkedvCenterTotal = $LinkedvCenterImport.Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $LinkedvCenterTotal + $vCenterImport.Name.Count

	# Draw Objects
	$x = 0
	#$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$vCenter_to_LinkedvCenter_Complete.Forecolor = "Blue"
	$vCenter_to_LinkedvCenter_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $LinkedvCenter in ( $LinkedvCenterImport | Sort-Object Name ) )
	{
		$x += 2.50
		$LinkedvCenterObject = Add-VisioObjectVC $VCObj $LinkedvCenter
		Draw_LinkedvCenter
		$ObjectNumber++
		$vCenter_to_LinkedvCenter_Complete.Forecolor = "Blue"
		$vCenter_to_LinkedvCenter_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($vCenter_to_LinkedvCenter_Complete)

		if ( $debug -eq $true )`
		{ `
			$LinkedvCenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Linked vCenter object " $LinkedvCenterNumber " of " $LinkedvCenterTotal " - " $LinkedvCenter.Name
		}
		Connect-VisioObject $VCObject $LinkedvCenterObject
		$VCObject = $LinkedvCenterObject
	}
		
	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< vCenter_to_LinkedvCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Host
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VM to Host Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Machine to VMHost Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VM to Host"
	$DocsObj.Pages('VM to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
	
	# Load a set of stencils and select one to drop
	Visio_Shapes
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VMNumber = 0
	$VMTotal = ( $VMImport | Where-Object { $_.SRM.contains("placeholderVm") -eq $False } ).Name.Count
	$TemplateNumber = 0
	$TemplateTotal = $TemplateImport.Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VMTotal + $TemplateTotal + $vCenterImport.Name.Count
	
		
	# Draw Objects
	$x = 0
	$y = 1.50

	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VM_to_Host_Complete.Forecolor = "Blue"
	$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VM_to_Host_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VM_to_Host_Complete.Forecolor = "Blue"
		$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VM_to_Host_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VM_to_Host_Complete.Forecolor = "Blue"
			$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VM_to_Host_Complete.Forecolor = "Blue"
				$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50
				
				foreach ( $VM in ( $VmImport | Where-Object { $VmHost.VmId.contains( $_.MoRef ) -and $VmHost.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Host_Complete.Forecolor = "Blue"
						$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Host_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Host_Complete.Forecolor = "Blue"
							$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Host_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Host_Complete.Forecolor = "Blue"
							$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Host_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}	
					Connect-VisioObject $HostObject $VMObject
					$HostObject = $VMObject
				}
				
				foreach ( $Template in ( $TemplateImport | Where-Object { $VmHost.TemplateId.contains( $_.MoRef ) -and $VmHost.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template
					$ObjectNumber++
					$VM_to_Host_Complete.Forecolor = "Blue"
					$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$TemplateNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
					}
					Connect-VisioObject $HostObject $TemplateObject
					$HostObject = $TemplateObject
				}
			}
		}
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VM_to_Host_Complete.Forecolor = "Blue"
			$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_Host_Complete)
			
			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50
			
			foreach ( $VM in ( $VmImport | Where-Object { $VmHost.VmId.contains( $_.MoRef ) -and $VmHost.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
			{ `
				$x += 2.50
				if ( $VM.OS -eq "" ) `
				{ `
					$VMObject = Add-VisioObjectVM $OtherObj $VM
					Draw_VM
					$ObjectNumber++
					$VM_to_Host_Complete.Forecolor = "Blue"
					$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Host_Complete)
			
					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
					}
				}
				else `
				{ `
					if ( $VM.OS.contains("Microsoft") -eq $True ) `
					{ `
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Host_Complete.Forecolor = "Blue"
						$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Host_Complete)
					
						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Host_Complete.Forecolor = "Blue"
						$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Host_Complete)
						
						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
				}
				Connect-VisioObject $HostObject $VMObject
				$HostObject = $VMObject
			}
			foreach ( $Template in ( $TemplateImport | Where-Object { $VmHost.TemplateId.contains( $_.MoRef ) -and $VmHost.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
			{ `
				$x += 2.50
				$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
				Draw_Template
				$ObjectNumber++
				$VM_to_Host_Complete.Forecolor = "Blue"
				$VM_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Host_Complete)
				
				if ( $debug -eq $true )`
				{ `
					$TemplateNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
				}
				Connect-VisioObject $HostObject $TemplateObject
				$HostObject = $TemplateObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Folder
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VM to Folder Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Machine to Folder Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VM to Folder"
	$DocsObj.Pages('VM to Folder')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Folder')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$FolderNumber = 0
	$FolderTotal = ( $FolderImport | Where-Object { $_.ParentId -like "Folder-group-v*" } ).Name.Count	
	$VMNumber = 0
	$VMTotal = ( $VMImport | Where-Object { $_.SRM.contains("placeholderVm") -eq $False } ).Name.Count
	$TemplateNumber = 0
	$TemplateTotal = $TemplateImport.Name.Count	
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $FolderTotal + $VMTotal + $TemplateTotal + $vCenterImport.Name.Count

	# Draw Objects
	$x = 0
	$y = 1.50
	
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VM_to_Folder_Complete.Forecolor = "Blue"
	$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VM_to_Folder_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VM_to_Folder_Complete.Forecolor = "Blue"
		$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VM_to_Folder_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		$y += 1.50
		
		foreach ( $Folder in ( $FolderImport | Where-Object { $_.DatacenterId -eq ( $Datacenter.MoRef ) -and $_.Datacenter -eq ( $Datacenter.Name ) } | Sort-Object Parent, Name -Descending ) ) `
		{ `
			$x = 5.00
			
			if ( $Folder.Parent -like "vm" ) `
			{ `
				$FolderObject = Add-VisioObjectFolder $FolderObj $Folder
				Draw_Folder
				$ObjectNumber++
				$VM_to_Folder_Complete.Forecolor = "Blue"
				$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Folder_Complete)

				if ( $debug -eq $true )`
				{ `
					$FolderNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Folder" $FolderNumber " of " $FolderTotal " - " $Folder.Name
				}
				Connect-VisioObject $DatacenterObject $FolderObject
				
				foreach ( $SubFolder in ( $FolderImport | Where-Object { $_.Parent -notlike "vm" -and $_.ParentId -like $Folder.MoRef } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 7.50
					$y += 1.50
					$SubFolderObject = Add-VisioObjectFolder $FolderObj $SubFolder
					Draw_SubFolder
					$ObjectNumber++
					$VM_to_Folder_Complete.Forecolor = "Blue"
					$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Folder_Complete)

					if ( $debug -eq $true )`
					{ `
						$FolderNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Folder" $FolderNumber " of " $FolderTotal " - " $SubFolder.Name
					}
					Connect-VisioObject $FolderObject $SubFolderObject
					
					foreach ( $SubSubFolder in ( $FolderImport | Where-Object { $_.ParentId -like $SubFolder.MoRef } | Sort-Object Name -Descending ) ) `
					{ `
						$x = 10.00
						$y += 1.50
						$SubSubFolderObject = Add-VisioObjectFolder $FolderObj $SubSubFolder
						Draw_SubSubFolder
						$ObjectNumber++
						$VM_to_Folder_Complete.Forecolor = "Blue"
						$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Folder_Complete)

						if ( $debug -eq $true )`
						{ `
							$FolderNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Folder" $FolderNumber " of " $FolderTotal " - " $SubSubFolder.Name
						}
						Connect-VisioObject $SubFolderObject $SubSubFolderObject
						
						foreach ( $SubSubSubFolder in ( $FolderImport | Where-Object { $_.ParentId -like $SubSubFolder.MoRef } | Sort-Object Name -Descending ) ) `
						{ `
							$x = 12.50
							$y += 1.50
							$SubSubSubFolderObject = Add-VisioObjectFolder $FolderObj $SubSubSubFolder
							Draw_SubSubSubFolder
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$FolderNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Folder" $FolderNumber " of " $FolderTotal " - " $SubSubSubFolder.Name
							}
							Connect-VisioObject $SubSubFolderObject $SubSubSubFolderObject
							$y += 1.50
							
							foreach ( $Template in ( $TemplateImport | Where-Object { $SubSubSubFolder.TemplateId.contains( $_.MoRef ) -and $SubSubSubFolder.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
							{ `
								$x += 2.50
								$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
								Draw_Template
								$ObjectNumber++
								$VM_to_Folder_Complete.Forecolor = "Blue"
								$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Folder_Complete)

								Connect-VisioObject $SubSubSubFolderObject $TemplateObject
								$SubSubSubFolderObject = $TemplateObject
							}
										
							foreach ( $VM in ( $VmImport | Where-Object { $SubSubSubFolder.VmId.contains( $_.MoRef ) -and $SubSubSubFolder.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
							{ `
								$x += 2.50
								if ( $VM.OS -eq "" ) `
								{ `
									$VMObject = Add-VisioObjectVM $OtherObj $VM
									Draw_VM
									$ObjectNumber++
									$VM_to_Folder_Complete.Forecolor = "Blue"
									$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VM_to_Folder_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
									}
								}
								else `
								{ `
									if ( $VM.OS.contains("Microsoft") -eq $True ) `
									{ `
										$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
										Draw_VM
										$ObjectNumber++
										$VM_to_Folder_Complete.Forecolor = "Blue"
										$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
										$TabDraw.Controls.Add($VM_to_Folder_Complete)

										if ( $debug -eq $true )`
										{ `
											$VMNumber++
											$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
											Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
										}
									}
									else `
									{ `
										$VMObject = Add-VisioObjectVM $LinuxObj $VM
										Draw_VM
										$ObjectNumber++
										$VM_to_Folder_Complete.Forecolor = "Blue"
										$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
										$TabDraw.Controls.Add($VM_to_Folder_Complete)

										if ( $debug -eq $true )`
										{ `
											$VMNumber++
											$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
											Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
										}
									}
								}	
								Connect-VisioObject $SubSubSubFolderObject $VMObject
								$SubSubSubFolderObject = $VMObject
							}
						}
						$y += 1.50
						
						foreach ( $Template in ( $TemplateImport | Where-Object { $SubSubFolder.TemplateId.contains( $_.MoRef ) -and $SubSubFolder.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
						{ `
							$x += 2.50
							$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
							Draw_Template
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$TemplateNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
							}
							Connect-VisioObject $SubSubFolderObject $TemplateObject
							$SubSubFolderObject = $TemplateObject
						}
									
						foreach ( $VM in ( $VmImport | Where-Object { $SubSubFolder.VmId.contains( $_.MoRef ) -and $SubSubFolder.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
						{ `
							$x += 2.50
							if ( $VM.OS -eq "" ) `
							{ `
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_Folder_Complete.Forecolor = "Blue"
								$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Folder_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
							else `
							{ `
								if ( $VM.OS.contains("Microsoft") -eq $True ) `
								{ `
									$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
									Draw_VM
									$ObjectNumber++
									$VM_to_Folder_Complete.Forecolor = "Blue"
									$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VM_to_Folder_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
									}
								}
								else `
								{ `
									$VMObject = Add-VisioObjectVM $LinuxObj $VM
									Draw_VM
									$ObjectNumber++
									$VM_to_Folder_Complete.Forecolor = "Blue"
									$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VM_to_Folder_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
									}
								}
							}	
							Connect-VisioObject $SubSubFolderObject $VMObject
							$SubSubFolderObject = $VMObject
						}
					}
					$y += 1.50
					
					foreach ( $Template in ( $TemplateImport | Where-Object { $SubFolder.TemplateId.contains( $_.MoRef ) -and $SubFolder.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
						Draw_Template
						$ObjectNumber++
						$VM_to_Folder_Complete.Forecolor = "Blue"
						$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Folder_Complete)

						if ( $debug -eq $true )`
						{ `
							$TemplateNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
						}
						Connect-VisioObject $SubFolderObject $TemplateObject
						$SubFolderObject = $TemplateObject
					}
								
					foreach ( $VM in ( $VmImport | Where-Object { $SubFolder.VmId.contains( $_.MoRef ) -and $SubFolder.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						if ( $VM.OS -eq "" ) `
						{ `
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							if ( $VM.OS.contains("Microsoft") -eq $True ) `
							{ `
								$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_Folder_Complete.Forecolor = "Blue"
								$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Folder_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
							else `
							{ `
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_Folder_Complete.Forecolor = "Blue"
								$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Folder_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
						}	
						Connect-VisioObject $SubFolderObject $VMObject
						$SubFolderObject = $VMObject
					}
				}
				
				$x = 5.00
				$y += 1.50
				foreach ( $Template in ( $TemplateImport | Where-Object { $Folder.TemplateId.contains( $_.MoRef ) -and $Folder.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template
					$ObjectNumber++
					$VM_to_Folder_Complete.Forecolor = "Blue"
					$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Folder_Complete)

					if ( $debug -eq $true )`
					{ `
						$TemplateNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
					}
					Connect-VisioObject $FolderObject $TemplateObject
					$FolderObject = $TemplateObject
				}
							
				foreach ( $VM in ( $VmImport | Where-Object { $Folder.VmId.contains( $_.MoRef ) -and $Folder.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Folder_Complete.Forecolor = "Blue"
						$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Folder_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}	
					Connect-VisioObject $FolderObject $VMObject
					$FolderObject = $VMObject
				}
				$y += 1.50
			}
		}
		
		foreach ( $Folder in ( $FolderImport | Where-Object { $_.DatacenterId -eq ( $Datacenter.MoRef ) -and $_.Datacenter -eq ( $Datacenter.Name ) } | Sort-Object Parent, Name -Descending ) ) `
		{ `
			$x = 2.50
			
			if  ( $Folder.Name -eq "vm" ) `
			{ `
				foreach ( $Template in ( $TemplateImport | Where-Object { $Datacenter.TemplateId.contains( $_.MoRef ) -and $Datacenter.Template.contains( $_.Name ) -and $_.Folder -eq "vm" } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template
					$ObjectNumber++
					$VM_to_Folder_Complete.Forecolor = "Blue"
					$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Folder_Complete)

					if ( $debug -eq $true )`
					{ `
						$TemplateNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " ( $TemplateTotal ) " - " $Template.Name
					}
					Connect-VisioObject $DatacenterObject $TemplateObject
					$DatacenterObject = $TemplateObject
				}
			
				foreach ( $VM in ( $VmImport | Where-Object { ( $Datacenter.VmId.contains( $_.MoRef ) -and $Datacenter.Vm.contains( $_.Name ) -and$_.SRM.contains("placeholderVm") -eq $False ) -and $_.Folder -eq "vm" } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Folder_Complete.Forecolor = "Blue"
						$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Folder_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Folder_Complete.Forecolor = "Blue"
							$VM_to_Folder_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Folder_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}	
					Connect-VisioObject $DatacenterObject $VMObject
					$DatacenterObject = $VMObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMs_with_RDMs
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VM with RDMs Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Machine with RDMs Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VM w/ RDMs"
	$DocsObj.Pages('VM w/ RDMs')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM w/ RDMs')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMNumber = 0
	$VMTotal = ( ( $RdmImport ).VmId | Select-Object -Unique ).Count
	$RDMNumber = 0
	$RDMTotal = ( ( $RdmImport ).ScsiCanonicalName ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMTotal + $RDMTotal + $vCenterImport.Name.Count
	
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter		
	$ObjectNumber++
	$VMs_with_RDMs_Complete.Forecolor = "Blue"
	$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter - " $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VMs_with_RDMs_Complete.Forecolor = "Blue"
		$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}		
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VMs_with_RDMs_Complete.Forecolor = "Blue"
			$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VM in ( $VmImport | Where-Object { $Cluster.VmId.contains( $_.MoRef ) -and $Cluster.Vm.contains( $_.Name ) -and $RdmImport.VmId -eq ( $_.MoRef ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				if ( $VM.OS -eq "" ) `
				{ `
					$VMObject = Add-VisioObjectVM $OtherObj $VM
					Draw_VM
					$ObjectNumber++
					$VMs_with_RDMs_Complete.Forecolor = "Blue"
					$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
					}
				}
				else `
				{ `
					if ( $VM.OS.contains("Microsoft") -eq $True ) `
					{ `
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
						$ObjectNumber++
						$VMs_with_RDMs_Complete.Forecolor = "Blue"
						$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
						}
					}
					else `
					{ `
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
						$ObjectNumber++
						$VMs_with_RDMs_Complete.Forecolor = "Blue"
						$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
						}
					}
				}
				Connect-VisioObject $ClusterObject $VMObject
				$y += 1.50
				
				foreach ( $HardDisk in ( $RdmImport | Sort-Object Label | Where-Object { $_.DatacenterId -eq ( $Datacenter.MoRef ) -and $_.ClusterId -eq ( $Cluster.MoRef ) -and $_.VmId -eq ( $Vm.MoRef ) } ) ) `
				{ `
					$x += 2.50
					$RDMObject = Add-VisioObjectHardDisk $RDMObj $HardDisk
					Draw_RDM
					$ObjectNumber++
					$VMs_with_RDMs_Complete.Forecolor = "Blue"
					$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

					if ( $debug -eq $true )`
					{ `
						$RDMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing RDM" $RDMNumber " of " ( $RDMImport.ScsiCanonicalName.Count ) " - " $HardDisk.ScsiCanonicalName
					}
					Connect-VisioObject $VMObject $RDMObject
					$VMObject = $RDMObject
				}
			}		
		}	
		foreach ( $VM in ( $VmImport | Where-Object { $_.DatacenterId -eq ( $Datacenter.MoRef ) -and $_.ClusterId -eq "" -and $RdmImport.VmId -eq ( $_.MoRef ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			if ( $VM.OS -eq "" ) `
			{ `
				$VMObject = Add-VisioObjectVM $OtherObj $VM
				Draw_VM
				$ObjectNumber++
				$VMs_with_RDMs_Complete.Forecolor = "Blue"
				$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
				}
			}
			else `
			{ `
				if ( $VM.OS.contains("Microsoft") -eq $True ) `
				{ `
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
					$ObjectNumber++
					$VMs_with_RDMs_Complete.Forecolor = "Blue"
					$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
					}
				}
				else `
				{ `
					$VMObject = Add-VisioObjectVM $LinuxObj $VM
					Draw_VM
					$ObjectNumber++
					$VMs_with_RDMs_Complete.Forecolor = "Blue"
					$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $RDMTotal " - " $VM.Name
					}
				}
			}
			Connect-VisioObject $DatacenterObject $VMObject
			$y += 1.50
			
			foreach ( $HardDisk in ( $RdmImport | Sort-Object Label | Where-Object { $_.DatacenterId -eq ( $Datacenter.MoRef ) -and $_.VmId -eq ( $Vm.MoRef ) } ) ) `
			{ `
				$x += 2.50
				$RDMObject = Add-VisioObjectHardDisk $RDMObj $HardDisk
				Draw_RDM
				$ObjectNumber++
				$VMs_with_RDMs_Complete.Forecolor = "Blue"
				$VMs_with_RDMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMs_with_RDMs_Complete)

				if ( $debug -eq $true )`
				{ `
					$RDMNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing RDM" $RDMNumber " of " ( $RDMImport.ScsiCanonicalName.Count ) " - " $HardDisk.ScsiCanonicalName
				}
				Connect-VisioObject $VMObject $RDMObject
				$VMObject = $RDMObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VMs_with_RDMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function SRM_Protected_VMs
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] SRM Protected VMs Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Site Recovery Manager Protected Virtual Machine Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "SRM VM"
	$DocsObj.Pages('SRM VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('SRM VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes

	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VMNumber = 0
	$VMTotal = ( $VMImport | Where-Object { $_.SRM.contains("placeholderVm") -eq $True } ).Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VMTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$SRM_Protected_VMs_Complete.Forecolor = "Blue"
	$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$SRM_Protected_VMs_Complete.Forecolor = "Blue"
		$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		$y += 1.50
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$SRM_Protected_VMs_Complete.Forecolor = "Blue"
			$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$SRM_Protected_VMs_Complete.Forecolor = "Blue"
				$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50
				
				foreach ( $SrmVM in ( $VmImport | Where-Object { $VmHost.VmId.contains( $_.MoRef ) -and $VmHost.Vm.contains( $_.Name ) -and $_.SRM.contains("placeholderVm") } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$SrmObject = Add-VisioObjectSRM $SRMObj $SrmVM
					Draw_SRM
					$ObjectNumber++
					$SRM_Protected_VMs_Complete.Forecolor = "Blue"
					$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $SrmVM.Name
					}
					Connect-VisioObject $HostObject $SrmObject
					$HostObject = $SrmObject
				}
			}
		}
		
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$SRM_Protected_VMs_Complete.Forecolor = "Blue"
			$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50
			
			foreach ( $SrmVM in ( $VmImport | Where-Object { $_.ClusterId -eq "" -and $VmHost.VmId.contains( $_.MoRef ) -and $VmHost.Vm.contains( $_.Name ) -and $_.SRM.contains("placeholderVm") } | Sort-Object Name -Descending ) ) `
			{ `
				$x += 2.50
				$SrmObject = Add-VisioObjectSRM $SRMObj $SrmVM
				Draw_SRM
				$ObjectNumber++
				$SRM_Protected_VMs_Complete.Forecolor = "Blue"
				$SRM_Protected_VMs_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($SRM_Protected_VMs_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $SrmVM.Name
				}
				Connect-VisioObject $HostObject $SrmObject
				$HostObject = $SrmObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< SRM_Protected_VMs >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VM_to_Datastore
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VM to Datastore Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Machine to Datastore Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VM to Datastore"
	$DocsObj.Pages('VM to Datastore')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Datastore')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMNumber = 0
	$VMTotal = ( ( $VMImport | Where-Object { $_.SRM.contains("placeholderVm") -eq $False } ).DatastoreId -split "," ).Count
	$TemplateNumber = 0
	$TemplateTotal = ( ( $TemplateImport ).DatastoreId -split "," ).Count
	$DatastoreClusterNumber = 0
	$DatastoreClusterTotal = ( ( $DatastoreClusterImport ).ClusterId -split "," ).Count
	$DatastoreNumber = 0
	$DatastoreTotal = ( ( $DatastoreImport ).ClusterId -split "," ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMTotal + $TemplateTotal + $DatastoreClusterTotal + $DatastoreTotal + $vCenterImport.Name.Count
	
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VM_to_Datastore_Complete.Forecolor = "Blue"
	$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VM_to_Datastore_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}		
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VM_to_Datastore_Complete.Forecolor = "Blue"
		$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VM_to_Datastore_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VM_to_Datastore_Complete.Forecolor = "Blue"
			$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_Datastore_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $DatastoreCluster in ( $DatastoreClusterImport | Where-Object { $Cluster.DatastoreClusterId.contains( $_.MoRef ) -and $Cluster.DatastoreCluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DatastoreClusObject = Add-VisioObjectDatastore $DatastoreClusObj $DatastoreCluster
				Draw_DatastoreCluster
				$ObjectNumber++
				$VM_to_Datastore_Complete.Forecolor = "Blue"
				$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Datastore_Complete)

				if ( $debug -eq $true )`
				{ `
					$DatastoreClusterNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Datastore Cluster" $DatastoreClusterNumber " of " $DatastoreClusterTotal " - " $DatastoreCluster.Name
				}
				Connect-VisioObject $ClusterObject $DatastoreClusObject
				
				foreach ( $Datastore in ( $DatastoreImport | Where-Object { $DatastoreCluster.DatastoreId.contains( $_.MoRef ) -and $DatastoreCluster.Datastore.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
					Draw_Datastore
					$ObjectNumber++
					$VM_to_Datastore_Complete.Forecolor = "Blue"
					$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Datastore_Complete)

					if ( $debug -eq $true )`
					{ `
						$DatastoreNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
					}
					Connect-VisioObject $DatastoreClusObject $DatastoreObject
					$y += 1.50
					
					foreach ( $VM in ( $VmImport | Where-Object { $Datastore.VmId.contains( $_.MoRef ) -and $Datastore.Vm.contains( $_.Name ) -and $Cluster.VmId.contains( $_.MoRef ) -and $Cluster.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Datastore_Complete.Forecolor = "Blue"
							$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Datastore_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
						else `
						{ `
							if ( $VM.OS.contains("Linux") -eq $True ) `
							{ `
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_Datastore_Complete.Forecolor = "Blue"
								$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Datastore_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
								}
							}
							else `
							{ `
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_Datastore_Complete.Forecolor = "Blue"
								$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_Datastore_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
								}
							}
						}
						Connect-VisioObject $DatastoreObject $VMObject
						$DatastoreObject = $VMObject
					}
					foreach ( $Template in ( $TemplateImport | Where-Object { $Datastore.TemplateId.contains( $_.MoRef ) -and $Datastore.Template.contains( $_.Name ) -and $Cluster.TemplateId.contains( $_.MoRef ) -and $Cluster.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
						Draw_Template
						$ObjectNumber++
						$VM_to_Datastore_Complete.Forecolor = "Blue"
						$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Datastore_Complete)

						if ( $debug -eq $true )`
						{ `
							$TemplateNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " $TemplateTotal " - " $Template.Name
						}
						Connect-VisioObject $DatastoreObject $TemplateObject
						$DatastoreObject = $TemplateObject
					}
				}
			}
			foreach ( $Datastore in ( $DatastoreImport | Where-Object { $Cluster.DatastoreId.contains( $_.MoRef ) -and $Cluster.Datastore.contains( $_.Name ) -and $_.DatastoreClusterId -like "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				$ObjectNumber++
				$VM_to_Datastore_Complete.Forecolor = "Blue"
				$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Datastore_Complete)

				if ( $debug -eq $true )`
				{ `
					$DatastoreNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
				}
				Connect-VisioObject $ClusterObject $DatastoreObject
				$y += 1.50
				
				foreach ( $VM in ( $VmImport | Where-Object { $Datastore.VmId.contains( $_.MoRef ) -and $Datastore.Vm.contains( $_.Name ) -and $Cluster.VmId.contains( $_.MoRef ) -and $Cluster.Vm.contains( $_.Name ) -and( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS.contains("Microsoft") -eq $True ) `
					{ `
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Datastore_Complete.Forecolor = "Blue"
						$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Datastore_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Linux") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Datastore_Complete.Forecolor = "Blue"
							$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Datastore_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Datastore_Complete.Forecolor = "Blue"
							$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Datastore_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
					}
					Connect-VisioObject $DatastoreObject $VMObject
					$DatastoreObject = $VMObject
				}
				foreach ( $Template in ( $TemplateImport | Where-Object { $Datastore.TemplateId.contains( $_.MoRef ) -and $Datastore.Template.contains( $_.Name ) -and $Cluster.TemplateId.contains( $_.MoRef ) -and $Cluster.Template.contains( $_.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template
					$ObjectNumber++
					$VM_to_Datastore_Complete.Forecolor = "Blue"
					$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Datastore_Complete)

					if ( $debug -eq $true )`
					{ `
						$TemplateNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " $TemplateTotal " - " $Template.Name
					}	
					Connect-VisioObject $DatastoreObject $TemplateObject
					$DatastoreObject = $TemplateObject
				}
			}
		}
		foreach ( $DatastoreCluster in ( $DatastoreClusterImport | Where-Object { $Datacenter.DatastoreClusterId.contains( $_.MoRef ) -and $Datacenter.DatastoreCluster.contains( $_.Name ) -and $_.ClusterId -like "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$DatastoreClusObject = Add-VisioObjectDatastore $DatastoreClusObj $DatastoreCluster
			Draw_DatastoreCluster
			$ObjectNumber++
			$VM_to_Datastore_Complete.Forecolor = "Blue"
			$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_Datastore_Complete)

			if ( $debug -eq $true )`
			{ `
				$DatastoreClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Datastore Cluster" $DatastoreClusterNumber " of " $DatastoreClusterTotal " - " $DatastoreCluster.Name
			}
			Connect-VisioObject $DatacenterObject $DatastoreClusObject
			
			foreach ( $Datastore in ( $DatastoreImport | Where-Object { $DatastoreCluster.DatastoreId.contains( $_.MoRef ) -and $DatastoreCluster.Datastore.contains( $_.Name ) -and $_.ClusterId -like "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				$ObjectNumber++
				$VM_to_Datastore_Complete.Forecolor = "Blue"
				$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Datastore_Complete)

				if ( $debug -eq $true )`
				{ `
					$DatastoreNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
				}
				Connect-VisioObject $DatastoreClusObject $DatastoreObject
				$y += 1.50
				
				foreach ( $VM in ( $VmImport | Where-Object { $Datastore.VmId.contains( $_.MoRef ) -and $Datastore.Vm.contains( $_.Name ) -and $_._ClusterId -like "" -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS.contains("Microsoft") -eq $True ) `
					{ `
						$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Datastore_Complete.Forecolor = "Blue"
						$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Datastore_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Linux") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Datastore_Complete.Forecolor = "Blue"
							$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Datastore_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_Datastore_Complete.Forecolor = "Blue"
							$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_Datastore_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
					}
					Connect-VisioObject $HostObject $VMObject
					$HostObject = $VMObject
				}
				foreach ( $Template in ( $TemplateImport | Where-Object { $Datastore.TemplateId.contains( $_.MoRef ) -and $Datastore.Template.contains( $_.Name ) } -and $_._ClusterId -like "" | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
					Draw_Template
					$ObjectNumber++
					$VM_to_Datastore_Complete.Forecolor = "Blue"
					$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Datastore_Complete)

					if ( $debug -eq $true )`
					{ `
						$TemplateNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " $TemplateTotal " - " $Template.Name
					}	
					Connect-VisioObject $HostObject $TemplateObject
					$HostObject = $TemplateObject
				}
			}
		}
		foreach ( $Datastore in ( $DatastoreImport | Where-Object { $Datacenter.DatastoreId.contains( $_.MoRef ) -and $Datacenter.Datastore.contains( $_.Name ) -and $_.ClusterId -like "" -and $_.DatastoreClusterId -like "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
			Draw_Datastore
			$ObjectNumber++
			$VM_to_Datastore_Complete.Forecolor = "Blue"
			$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_Datastore_Complete)

			if ( $debug -eq $true )`
			{ `
				$DatastoreNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
			}
			Connect-VisioObject $DatacenterObject $DatastoreObject
			$y += 1.50
			
			foreach ( $VM in ( $VmImport | Where-Object { $Datastore.VmId.contains( $_.MoRef ) -and $Datastore.Vm.contains( $_.Name ) -and $_.ClusterId -like "" -and $_.DatastoreClusterId -like "" -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
			{ `
				$x += 2.50
				if ( $VM.OS.contains("Microsoft") -eq $True ) `
				{ `
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
					$ObjectNumber++
					$VM_to_Datastore_Complete.Forecolor = "Blue"
					$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_Datastore_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
					}
				}
				else `
				{ `
					if ( $VM.OS.contains("Linux") -eq $True ) `
					{ `
						$VMObject = Add-VisioObjectVM $LinuxObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Datastore_Complete.Forecolor = "Blue"
						$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Datastore_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
						}
					}
					else `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$VM_to_Datastore_Complete.Forecolor = "Blue"
						$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VM_to_Datastore_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
						}
					}
				}
				Connect-VisioObject $DatastoreObject $VMObject
				$DatastoreObject = $VMObject
			}
			foreach ( $Template in ( $TemplateImport | Where-Object { $Datastore.TemplateId.contains( $_.MoRef ) -and $Datastore.Template.contains( $_.Name ) -and $_.ClusterId -like "" -and $_.DatastoreClusterId -like "" } | Sort-Object Name ) ) `
			{ `
				$x += 2.50
				$TemplateObject = Add-VisioObjectTemplate $TemplateObj $Template
				Draw_Template
				$ObjectNumber++
				$VM_to_Datastore_Complete.Forecolor = "Blue"
				$VM_to_Datastore_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VM_to_Datastore_Complete)

				if ( $debug -eq $true )`
				{ `
					$TemplateNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Template" $TemplateNumber " of " $TemplateTotal " - " $Template.Name
				}
				Connect-VisioObject $DatastoreObject $TemplateObject
				$DatastoreObject = $TemplateObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VM_to_Datastore >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#add non-cluster
function VM_to_ResourcePool
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VM to Resource Pool Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Machine to Resource Pool Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VM to Resource Pool"
	$DocsObj.Pages('VM to Resource Pool')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VM to Resource Pool')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMNumber = 0
	$VMTotal = ( $VMImport | Where-Object { ( $_.SRM.contains("placeholderVm") -eq $False ) -and $_.ResourcePool -ne "" } ).Name.Count
	$ResourcePoolNumber = 0
	$ResourcePoolTotal = ( $ResourcePoolImport | Where-Object { $_.Name -notlike "Resources" } ).Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMTotal + $ResourcePoolTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VM_to_ResourcePool_Complete.Forecolor = "Blue"
	$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending  ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VM_to_ResourcePool_Complete.Forecolor = "Blue"
		$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VM_to_ResourcePool_Complete.Forecolor = "Blue"
			$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			$RootResoucePool = ( $ResourcePoolImport | Where-Object { $_.Name -like "Resources" -and $_.ClusterId -like ( $Cluster.MoRef ) } ).MoRef
			
			foreach ( $ResourcePool in ( $ResourcePoolImport | Where-Object { $Cluster.ResourcePoolId.contains( $_.MoRef ) -and $Cluster.ResourcePool.contains( $_.Name ) -and $_.ClusterId -notlike $_.Parent } | Sort-Object Name -Descending ) ) `
			{ `
				if ( $ResourcePool.Parent -like ("$RootResoucePool") ) `
				{ `
					$x = 7.50
					$y += 1.50
					$ResourcePoolObject = Add-VisioObjectResourcePool $ResourcePoolObj $ResourcePool
					Draw_ResourcePool
					$ObjectNumber++
					$VM_to_ResourcePool_Complete.Forecolor = "Blue"
					$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

					if ( $debug -eq $true )`
					{ `
						$ResourcePoolNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Resource Pool " $ResourcePoolNumber " of " $ResourcePoolTotal " - " $ResourcePool.Name
					}
					Connect-VisioObject $ClusterObject $ResourcePoolObject
					$ResourcePoolMoRef = $ResourcePool.MoRef
					
					foreach ( $SubResourcePool in ( $ResourcePoolImport | Where-Object { $Cluster.ResourcePoolId.contains( $_.MoRef ) -and $Cluster.ResourcePool.contains( $_.Name ) -and ( $_.Parent -like ( $ResourcePool.MoRef ) ) } | Sort-Object Name -Descending ) ) `
					{ `
						if ( $SubResourcePool.Parent.contains("$ResourcePoolMoRef") -eq $True ) `
						{ `
							$x = 10.00
							$y += 1.50
							$SubResourcePoolObject = Add-VisioObjectResourcePool $ResourcePoolObj $SubResourcePool
							Draw_SubResourcePool
							$ObjectNumber++
							$VM_to_ResourcePool_Complete.Forecolor = "Blue"
							$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

							if ( $debug -eq $true )`
							{ `
								$ResourcePoolNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Resource Pool " $ResourcePoolNumber " of " $ResourcePoolTotal " - " $SubResourcePool.Name
							}
							Connect-VisioObject $ResourcePoolObject $SubResourcePoolObject
							$SubResourcePoolMoRef = $SubResourcePool.MoRef
			
							foreach ( $SubSubResourcePool in ( $ResourcePoolImport | Where-Object { $Cluster.ResourcePoolId.contains( $_.MoRef ) -and $Cluster.ResourcePool.contains( $_.Name ) -and ( $_.Parent -like ( $SubResourcePool.MoRef ) ) } | Sort-Object Name -Descending ) ) `
							{ `
								if ( $SubSubResourcePool.Parent.contains("$SubResourcePoolMoRef") -eq $True ) `
								{ `
									$x = 12.50
									$y += 1.50
									$SubSubResourcePoolObject = Add-VisioObjectResourcePool $ResourcePoolObj $SubSubResourcePool
									Draw_SubSubResourcePool
									$ObjectNumber++
									$VM_to_ResourcePool_Complete.Forecolor = "Blue"
									$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

									if ( $debug -eq $true )`
									{ `
										$ResourcePoolNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing Resource Pool " $ResourcePoolNumber " of " $ResourcePoolTotal " - " $SubSubResourcePool.Name
									}
									Connect-VisioObject $SubResourcePoolObject $SubSubResourcePoolObject
									#$SubSubResourcePoolMoRef = $SubSubResourcePool.MoRef
									$y += 1.50
								
									foreach ( $VM in ( $VmImport | Where-Object { $SubSubResourcePool.VmId.contains( $_.MoRef ) -and $SubSubResourcePool.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
									{ `
										$x += 2.50
										if ( $VM.OS.contains("Microsoft") -eq $True ) `
										{ `
											$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
											Draw_VM
											$ObjectNumber++
											$VM_to_ResourcePool_Complete.Forecolor = "Blue"
											$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
											$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

											if ( $debug -eq $true )`
											{ `
												$VMNumber++
												$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
												Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
											}
										}
										else `
										{ `
											if ( $VM.OS.contains("Linux") -eq $True ) `
											{ `
												$VMObject = Add-VisioObjectVM $LinuxObj $VM
												Draw_VM
												$ObjectNumber++
												$VM_to_ResourcePool_Complete.Forecolor = "Blue"
												$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
												$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

												if ( $debug -eq $true )`
												{ `
													$VMNumber++
													$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
													Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
												}
											}
											else `
											{ `
												$VMObject = Add-VisioObjectVM $OtherObj $VM
												Draw_VM
												$ObjectNumber++
												$VM_to_ResourcePool_Complete.Forecolor = "Blue"
												$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
												$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

												if ( $debug -eq $true )`
												{ `
													$VMNumber++
													$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
													Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
												}
											}
										}
										Connect-VisioObject $SubSubResourcePoolObject $VMObject
										$SubSubResourcePoolObject = $VMObject
									}
								}
							}
							
							$x = 10.00
							$y = ( $y + 1.50 )
							foreach ( $VM in ( $VmImport | Where-Object { $SubResourcePool.VmId.contains( $_.MoRef ) -and $SubResourcePool.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
							{ `
								$x += 2.50
								if ( $VM.OS.contains("Microsoft") -eq $True ) `
								{ `
									$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
									Draw_VM
									$ObjectNumber++
									$VM_to_ResourcePool_Complete.Forecolor = "Blue"
									$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
									}
								}
								else `
								{ `
									if ( $VM.OS.contains("Linux") -eq $True ) `
									{ `
										$VMObject = Add-VisioObjectVM $LinuxObj $VM
										Draw_VM
										$ObjectNumber++
										$VM_to_ResourcePool_Complete.Forecolor = "Blue"
										$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
										$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

										if ( $debug -eq $true )`
										{ `
											$VMNumber++
											$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
											Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
										}
									}
									else `
									{ `
										$VMObject = Add-VisioObjectVM $OtherObj $VM
										Draw_VM
										$ObjectNumber++
										$VM_to_ResourcePool_Complete.Forecolor = "Blue"
										$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
										$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

										if ( $debug -eq $true )`
										{ `
											$VMNumber++
											$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
											Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
										}
									}
								}
								Connect-VisioObject $SubResourcePoolObject $VMObject
								$SubResourcePoolObject = $VMObject
							}
						}
					}
					
					$x = 7.50
					$y = ( $y + 1.50 )
					
					foreach ( $VM in ( $VmImport | Where-Object { $ResourcePool.VmId.contains( $_.MoRef ) -and $ResourcePool.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VM_to_ResourcePool_Complete.Forecolor = "Blue"
							$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
							}
						}
						else `
						{ `
							if ( $VM.OS.contains("Linux") -eq $True ) `
							{ `
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_ResourcePool_Complete.Forecolor = "Blue"
								$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
								}
							}
							else `
							{ `
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
								$ObjectNumber++
								$VM_to_ResourcePool_Complete.Forecolor = "Blue"
								$VM_to_ResourcePool_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VM_to_ResourcePool_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
								}
							}
						}
						Connect-VisioObject $ResourcePoolObject $VMObject
						$ResourcePoolObject = $VMObject
					}
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VM_to_ResourcePool >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datastore_to_Host
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Datastore to Host Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Datastore to VMHost Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Datastore to Host"
	$DocsObj.Pages('Datastore to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Datastore to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$DatastoreNumber = 0
	$DatastoreTotal = ( ( $DatastoreImport ).VmhostId -split "," ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $DatastoreTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$Datastore_to_Host_Complete.Forecolor = "Blue"
	$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($Datastore_to_Host_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$Datastore_to_Host_Complete.Forecolor = "Blue"
		$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($Datastore_to_Host_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$Datastore_to_Host_Complete.Forecolor = "Blue"
			$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($Datastore_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$Datastore_to_Host_Complete.Forecolor = "Blue"
				$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Datastore_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				$y += 1.50	
				
				foreach ( $Datastore in ( $DatastoreImport | Where-Object { $VmHost.DatastoreId.contains( $_.MoRef ) -and $VmHost.Datastore.contains( $_.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
					Draw_Datastore
					$ObjectNumber++
					$Datastore_to_Host_Complete.Forecolor = "Blue"
					$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Datastore_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$DatastoreNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
					}
					Connect-VisioObject $HostObject $DatastoreObject
					$HostObject = $DatastoreObject
				}
			}
		}
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$Datastore_to_Host_Complete.Forecolor = "Blue"
			$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($Datastore_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			$y += 1.50

			foreach ( $Datastore in ( $DatastoreImport | Where-Object { $VmHost.DatastoreId.contains( $_.MoRef ) -and $VmHost.Datastore.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x += 2.50
				$DatastoreObject = Add-VisioObjectDatastore $DatastoreObj $Datastore
				Draw_Datastore
				$ObjectNumber++
				$Datastore_to_Host_Complete.Forecolor = "Blue"
				$Datastore_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Datastore_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$DatastoreNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Datastore " $DatastoreNumber " of " $DatastoreTotal " - " $Datastore.Name
				}
				Connect-VisioObject $HostObject $DatastoreObject
				$HostObject = $DatastoreObject
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< Datastore_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Snapshot_to_VM
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Snapshot to VM Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Snapshot to Virtual Machine Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Snapshot to VM"
	$DocsObj.Pages('Snapshot to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Snapshot to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$VMNumber = 0
	$VMTotal = ( $SnapshotImport.VMId | Select-Object -Unique ).Count
	$SnapshotNumber = 0
	$SnapshotTotal = ( $SnapshotImport ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $VMTotal + $SnapshotTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$Snapshot_to_VM_Complete.Forecolor = "Blue"
	$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$Snapshot_to_VM_Complete.Forecolor = "Blue"
		$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $VM in( $VmImport | Where-Object { $Datacenter.VmId.contains( $_.MoRef ) -and $Datacenter.Vm.contains( $_.Name ) -and ( $_.Snapshot -notlike "" ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			if ( $VM.OS -eq "" ) `
			{ `
				$VMObject = Add-VisioObjectVM $OtherObj $VM
				Draw_VM
				$ObjectNumber++
				$Snapshot_to_VM_Complete.Forecolor = "Blue"
				$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
				}
			}
			else `
			{ `
				if ( $VM.OS.contains("Microsoft") -eq $True ) `
				{ `
					$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
					Draw_VM
					$ObjectNumber++
					$Snapshot_to_VM_Complete.Forecolor = "Blue"
					$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
					}
				}
				else `
				{ `
					$VMObject = Add-VisioObjectVM $LinuxObj $VM
					Draw_VM
					$ObjectNumber++
					$Snapshot_to_VM_Complete.Forecolor = "Blue"
					$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$VMNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing VM" $VMNumber " of " $VMTotal " - " $VM.Name
					}
				}
			}
			Connect-VisioObject $DatacenterObject $VMObject
			
			foreach ( $ParentSnapshot in ( $SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains( $VM.Name ) -and $_.VMId.contains( $VM.MoRef) -and ( $_.ParentSnapshot -like $null ) } ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				if ( $ParentSnapshot.IsCurrent -eq "FALSE" ) `
				{ `
					$ParentSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ParentSnapshot
					Draw_ParentSnapshot
					$ObjectNumber++
					$Snapshot_to_VM_Complete.Forecolor = "Blue"
					$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$SnapshotNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ParentSnapshot.Name
					}
				}
				else `
				{ `
					$ParentSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ParentSnapshot
					Draw_ParentSnapshot
					$ObjectNumber++
					$Snapshot_to_VM_Complete.Forecolor = "Blue"
					$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$SnapshotNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ParentSnapshot.Name
					}
				}
				Connect-VisioObject $VMObject $ParentSnapshotObject 
				
				foreach ( $ChildSnapshot in ( $SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains( $VM.Name ) -and $_.VMId.contains( $VM.MoRef) -and ( $_.ParentSnapshot -like $ParentSnapshot.Name ) } ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					if ( $ChildSnapshot.IsCurrent -eq "FALSE" ) `
					{ `
						$ChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildSnapshot
						Draw_ChildSnapshot
						$ObjectNumber++
						$Snapshot_to_VM_Complete.Forecolor = "Blue"
						$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

						if ( $debug -eq $true )`
						{ `
							$SnapshotNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildSnapshot.Name
						}
					}
					else `
					{ `
						$ChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildSnapshot
						Draw_ChildSnapshot
						$ObjectNumber++
						$Snapshot_to_VM_Complete.Forecolor = "Blue"
						$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

						if ( $debug -eq $true )`
						{ `
							$SnapshotNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildSnapshot.Name
						}
					}
					Connect-VisioObject $ParentSnapshotObject $ChildSnapshotObject
					
					foreach ( $ChildChildSnapshot in ( $SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains( $VM.Name ) -and $_.VMId.contains( $VM.MoRef) -and ( $_.ParentSnapshot -like $ChildSnapshot.Name ) } ) ) `
					{ `
						$x = 12.50
						$y += 1.50
						if ( $ChildChildSnapshot.IsCurrent -eq "FALSE" ) `
						{ `
							$ChildChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildChildSnapshot
							Draw_ChildChildSnapshot
							$ObjectNumber++
							$Snapshot_to_VM_Complete.Forecolor = "Blue"
							$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

							if ( $debug -eq $true )`
							{ `
								$SnapshotNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildChildSnapshot.Name
							}
						}
						else `
						{ `
							$ChildChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildChildSnapshot
							Draw_ChildChildSnapshot
							$ObjectNumber++
							$Snapshot_to_VM_Complete.Forecolor = "Blue"
							$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

							if ( $debug -eq $true )`
							{ `
								$SnapshotNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildChildSnapshot.Name
							}
						}
						Connect-VisioObject $ChildSnapshotObject $ChildChildSnapshotObject
						$y += 1.50
						
						foreach ( $ChildChildChildSnapshot in ( $SnapshotImport | Sort-Object Created | Where-Object { $_.VM.contains( $VM.Name ) -and $_.VMId.contains( $VM.MoRef) -and ($_.ParentSnapshot -like $ChildChildSnapshot.Name ) } ) ) `
						{ `
							$x += 2.50
							$y += 1.50
							if ( $ChildChildChildSnapshot.IsCurrent -eq "FALSE" ) `
							{ `
								$ChildChildChildSnapshotObject = Add-VisioObjectSnapshot $SnapshotObj $ChildChildChildSnapshot
								Draw_ChildChildChildSnapshot
								$ObjectNumber++
								$Snapshot_to_VM_Complete.Forecolor = "Blue"
								$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

								if ( $debug -eq $true )`
								{ `
									$SnapshotNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildChildChildSnapshot.Name
								}
							}
							else `
							{ `
								$ChildChildChildSnapshotObject = Add-VisioObjectSnapshot $CurrentSnapshotObj $ChildChildChildSnapshot
								Draw_ChildChildChildSnapshot
								$ObjectNumber++
								$Snapshot_to_VM_Complete.Forecolor = "Blue"
								$Snapshot_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($Snapshot_to_VM_Complete)

								if ( $debug -eq $true )`
								{ `
									$SnapshotNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing Snapshot" $SnapshotNumber " of " $SnapshotTotal " - " $ChildChildChildSnapshot.Name
								}
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
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< Snapshot_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function PhysicalNIC_to_vSwitch
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Physical NIC to vSwitch Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Physical NIC to vSwitch Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
		
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "PNIC to switch"
	$DocsObj.Pages('PNIC to Switch')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('PNIC to Switch')
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VsSwitchNumber = 0
	$VsSwitchTotal = $VsSwitchImport.Name.Count
	$VssPnicNumber = 0
	$VssPnicTotal = $VssPnicImport.Name.Count
	$VdSwitchNumber = 0
	$VdSwitchTotal = ( ( $VdSwitchImport ).VmHostId -split "," ).Count
	$VdsPnicNumber = 0
	$VdsPnicTotal = ( $VdsPnicImport | Where-Object { $_.Name -notlike "" } ).Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VsSwitchTotal + $VssPnicTotal + $VdSwitchTotal + $VdsPnicTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
	$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
		$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
			$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			$ClusterObject.Cells("Prop.HostMonitoring").Formula = '"' + $Cluster.HostMonitoring + '"'
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
				$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					$ObjectNumber++
					$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
					$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

					if ( $debug -eq $true )`
					{ `
						$VsSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
					}
					Connect-VisioObject $HostObject $VsSwitchObject
					$y += 1.50
					
					if ( $null -ne $VssPnicImport ) `
					{ `
						foreach ( $VssPnic in ( $VssPnicImport | Where-Object { $VsSwitch.NicId.contains( $_.MoRef ) -and $VsSwitch.Nic.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
						{ `
							$x += 2.50
							$VssPNICObject = Add-VisioObjectVssPNIC $VssPNICObj $VssPnic
							Draw_VssPnic
							$ObjectNumber++
							$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
							$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

							if ( $debug -eq $true )`
							{ `
								$VssPnicNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing Virtual Standard Switch Uplink object " $VssPnicNumber " of " $VssPnicTotal " - " $VssPnic.Name
							}
							Connect-VisioObject $VsSwitchObject $VssPNICObject
							$VsSwitchObject = $VssPNICObject
						}
					}
				}
				
				foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					$ObjectNumber++
					$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
					$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
					}
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
					
					foreach ( $VdsPnic in ( $VdsPnicImport | Where-Object { $VdSwitch.NicId.contains( $_.MoRef ) -and $VdSwitch.Nic.contains( $_.Name )-and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$VdsPNICObject = Add-VisioObjectVdsPNIC $VdsPNICObj $VdsPnic
						Draw_VdsPnic
						$ObjectNumber++
						$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
						$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

						if ( $debug -eq $true )`
						{ `
							$VdsPnicNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Distributed Switch Uplink object " $VdsPnicNumber " of " $VdsPnicTotal " - " $VdsPnic.Name
						}
						Connect-VisioObject $VdSwitchObject $VdsPNICObject
						$VdSwitchObject = $VdsPNICObject
					}
				}
			}
		}
		
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
			$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $VmHost.ClusterId -eq "" -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				$ObjectNumber++
				$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
				$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

				if ( $debug -eq $true )`
				{ `
					$VsSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
				}
				Connect-VisioObject $HostObject $VsSwitchObject
				$y += 1.50
				
				if ( $null -ne $VssPnicImport ) `
				{ `
					foreach ( $VssPnic in ( $VssPnicImport | Where-Object { $VsSwitch.NicId.contains( $_.MoRef ) -and $VsSwitch.Nic.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$VssPNICObject = Add-VisioObjectVssPNIC $VssPNICObj $VssPnic
						Draw_VssPnic
						$ObjectNumber++
						$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
						$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

						if ( $debug -eq $true )`
						{ `
							$VssPnicNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Standard Switch Uplink object " $VssPnicNumber " of " $VssPnicTotal " - " $VssPnic.Name
						}
						Connect-VisioObject $VsSwitchObject $VssPNICObject
						$VsSwitchObject = $VssPNICObject
					}
				}
			}
			
			foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $VmHost.ClusterId -eq "" -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				$ObjectNumber++
				$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
				$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

				if ( $debug -eq $true )`
				{ `
					$VdSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
				}
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50
				
				foreach ( $VdsPnic in ( $VdsPnicImport | Where-Object { $VdSwitch.NicId.contains( $_.MoRef ) -and $VdSwitch.Nic.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$VdsPNICObject = Add-VisioObjectVdsPNIC $VdsPNICObj $VdsPnic
					Draw_VdsPnic
					$ObjectNumber++
					$PhysicalNIC_to_vSwitch_Complete.Forecolor = "Blue"
					$PhysicalNIC_to_vSwitch_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($PhysicalNIC_to_vSwitch_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdsPnicNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed Switch Uplink object " $VdsPnicNumber " of " $VdsPnicTotal " - " $VdsPnic.Name
					}
					Connect-VisioObject $VdSwitchObject $VdsPNICObject
					$VdSwitchObject = $VdsPNICObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< PhysicalNIC_to_vSwitch >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VSS_to_Host
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Standard Switch to Host Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Standard Switch to VMHost Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
		
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VSS to Host"
	$DocsObj.Pages('VSS to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VSS to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VsSwitchNumber = 0
	$VsSwitchTotal = $VsSwitchImport.Name.Count
	$VssPortGroupNumber = 0
	$VssPortGroupTotal = $VssPortImport.Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VsSwitchTotal + $VssPortGroupTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VSS_to_Host_Complete.Forecolor = "Blue"
	$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VSS_to_Host_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VSS_to_Host_Complete.Forecolor = "Blue"
		$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VSS_to_Host_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VSS_to_Host_Complete.Forecolor = "Blue"
			$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VSS_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VSS_to_Host_Complete.Forecolor = "Blue"
				$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VSS_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					$ObjectNumber++
					$VSS_to_Host_Complete.Forecolor = "Blue"
					$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VSS_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$VsSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
					}
					Connect-VisioObject $HostObject $VsSwitchObject
					$y += 1.50
					
					foreach ( $VssPort in ( $VssPortImport | Where-Object { $VsSwitch.PortGroupId.contains( $_.MoRef ) -and $VsSwitch.PortGroup.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$VssPortObject = Add-VisioObjectVssPG $VssPortGroupObj $VssPort
						Draw_VssPort
						$ObjectNumber++
						$VSS_to_Host_Complete.Forecolor = "Blue"
						$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VSS_to_Host_Complete)

						if ( $debug -eq $true )`
						{ `
							$VssPortGroupNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Standard Port Group object " $VssPortGroupNumber " of " $VssPortGroupTotal " - " $VssPort.Name
						}
						Connect-VisioObject $VsSwitchObject $VssPortObject
						$VsSwitchObject = $VssPortObject
					}
				}
			}
		}
		
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VSS_to_Host_Complete.Forecolor = "Blue"
			$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VSS_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.ClusterId -eq "" -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				$ObjectNumber++
				$VSS_to_Host_Complete.Forecolor = "Blue"
				$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VSS_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VsSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
				}
				Connect-VisioObject $HostObject $VsSwitchObject
				$y += 1.50
				
				foreach ( $VssPort in ( $VssPortImport | Where-Object { $VsSwitch.PortGroupId.contains( $_.MoRef ) -and $VsSwitch.PortGroup.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$VssPortObject = Add-VisioObjectVssPG $VssPortGroupObj $VssPort
					Draw_VssPort
					$ObjectNumber++
					$VSS_to_Host_Complete.Forecolor = "Blue"
					$VSS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VSS_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$VssPortGroupNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Port Group object " $VssPortGroupNumber " of " $VssPortGroupTotal " - " $VssPort.Name
					}
					Connect-VisioObject $VsSwitchObject $VssPortObject
					$VsSwitchObject = $VssPortObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VSS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMK_to_VSS
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VMKernel to Standard Switch Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating VMKernel to Virtual Standard Switch Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VMK to VSS"
	$DocsObj.Pages('VMK to VSS')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VMK to VSS')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VsSwitchNumber = 0
	$VsSwitchTotal = $VsSwitchImport.Name.Count
	$VssVmkernelNumber = 0
	$VssVmkernelTotal = $VssVmkImport.Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VsSwitchTotal + $VssVmkernelTotal + $vCenterImport.Name.Count
	
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VMK_to_VSS_Complete.Forecolor = "Blue"
	$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VMK_to_VSS_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VMK_to_VSS_Complete.Forecolor = "Blue"
		$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VMK_to_VSS_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VMK_to_VSS_Complete.Forecolor = "Blue"
			$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VMK_to_VSS_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VMK_to_VSS_Complete.Forecolor = "Blue"
				$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMK_to_VSS_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					$ObjectNumber++
					$VMK_to_VSS_Complete.Forecolor = "Blue"
					$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMK_to_VSS_Complete)

					if ( $debug -eq $true )`
					{ `
						$VsSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
					}
					Connect-VisioObject $HostObject $VsSwitchObject
					$y += 1.50
					
					foreach ( $VssVmk in ( $VssVmkImport | Sort-Object Name | Where-Object { $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.VSwitchId.contains( $VsSwitch.MoRef ) -and $_.VSwitch.contains( $VsSwitch.Name ) } ) ) `
					{ `
						$x += 2.50
						$VssVmkNicObject = Add-VisioObjectVMK $VssVmkNicObj $VssVmk
						Draw_VssVmk
						$ObjectNumber++
						$VMK_to_VSS_Complete.Forecolor = "Blue"
						$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VMK_to_VSS_Complete)

						if ( $debug -eq $true )`
						{ `
							$VssVmkernelNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Standard VMkernel object " $VssVmkernelNumber " of " $VssVmkernelTotal " - " $VssVmk.Name
						}
						Connect-VisioObject $VsSwitchObject $VssVmkNicObject
						$VsSwitchObject = $VssVmkNicObject
					}
				}
			}
		}

		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" }  | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VMK_to_VSS_Complete.Forecolor = "Blue"
			$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VMK_to_VSS_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				$ObjectNumber++
				$VMK_to_VSS_Complete.Forecolor = "Blue"
				$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMK_to_VSS_Complete)

				if ( $debug -eq $true )`
				{ `
					$VsSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
				}
				Connect-VisioObject $HostObject $VsSwitchObject
				$y += 1.50
				
				foreach ( $VssVmk in ( $VssVmkImport | Sort-Object Name | Where-Object { $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.VSwitchId.contains( $VsSwitch.MoRef ) -and $_.VSwitch.contains( $VsSwitch.Name ) -and $_.ClusterId -eq "" } ) ) `
				{ `
					$x += 1.50
					$VssVmkNicObject = Add-VisioObjectVMK $VssVmkNicObj $VssVmk
					Draw_VssVmk
					$ObjectNumber++
					$VMK_to_VSS_Complete.Forecolor = "Blue"
					$VMK_to_VSS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMK_to_VSS_Complete)

					if ( $debug -eq $true )`
					{ `
						$VssVmkernelNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard VMkernel object " $VssVmkernelNumber " of " $VssVmkernelTotal " - " $VssVmk.Name
					}
					Connect-VisioObject $VsSwitchObject $VssVmkNicObject
					$VsSwitchObject = $VssVmkNicObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VMK_to_VSS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VSSPortGroup_to_VM
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Standard Switch Port Group to VM Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Standard Switch Port Group to Virtual Machine Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VSSPortGroup to VM"
	$DocsObj.Pages('VSSPortGroup to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VSSPortGroup to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
	
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VsSwitchNumber = 0
	$VsSwitchTotal = $VsSwitchImport.Name.Count
	$VssPortGroupNumber = 0
	$VssPortGroupTotal = $VssPortImport.Name.Count
	$VMNumber = 0
	$VMTotal = ( ( $VmImport | Where-Object { $_.PortGroupId -notlike "DistributedVirtualPortgroup*" -and $_.SRM.contains("placeholderVm") -eq $False } ).PortGroupId -split "," ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VsSwitchTotal + $VssPortGroupTotal + $VMTotal + $vCenterImport.Name.Count


	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
	$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
		$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
			$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
				$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VsSwitch in ( $VsSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
					Draw_VsSwitch
					$ObjectNumber++
					$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
					$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$VsSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
					}
					Connect-VisioObject $HostObject $VsSwitchObject
					
					foreach ( $VssPort in ( $VssPortImport | Where-Object { $VsSwitch.PortGroupId.contains( $_.MoRef ) -and $VsSwitch.PortGroup.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
					{ `
						$x = 12.50
						$y += 1.50
						$VssPortObject = Add-VisioObjectVssPG $VssPortGroupObj $VssPort
						Draw_VssPort
						$ObjectNumber++
						$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
						$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

						if ( $debug -eq $true )`
						{ `
							$VssPortGroupNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Standard Port Group object " $VssPortGroupNumber " of " $VssPortGroupTotal " - " $VssPort.Name
						}
						Connect-VisioObject $VsSwitchObject $VssPortObject
						$y += 1.50
						
						foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $VssPort.VmId.contains( $_.MoRef ) -and $VssPort.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } | Sort-Object Name ) ) `
						{ `
							$x += 2.50
							if ( $VM.OS -eq "" ) `
							{ `
								$VMObject = Add-VisioObjectVM $OtherObj $VM
								Draw_VM
								$ObjectNumber++
								$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
								$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
							else `
							{ `
								if ( $VM.OS.contains("Microsoft") -eq $True ) `
								{ `
									$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
									Draw_VM
									$ObjectNumber++
									$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
									$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
									}
								}
								else `
								{ `
									$VMObject = Add-VisioObjectVM $LinuxObj $VM
									Draw_VM
									$ObjectNumber++
									$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
									$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
									$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

									if ( $debug -eq $true )`
									{ `
										$VMNumber++
										$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
										Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
									}
								}
							}
							Connect-VisioObject $VssPortObject $VMObject
							$VssPortObject = $VMObject
						}
					}
				}
			}
		}
		foreach ( $VmHost in ( $VmHostImport | Sort-Object Name | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
			$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VsSwitch in ( $VsSwitchImport | Sort-Object Name | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VsSwitchObject = Add-VisioObjectVsSwitch $VSSObj $VsSwitch
				Draw_VsSwitch
				$ObjectNumber++
				$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
				$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

				if ( $debug -eq $true )`
				{ `
					$VsSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Standard Switch object " $VsSwitchNumber " of " $VsSwitchTotal " - " $VsSwitch.Name
				}
				Connect-VisioObject $HostObject $VsSwitchObject
				
				foreach ( $VssPort in ( $VssPortImport | Sort-Object Name | Where-Object { $VsSwitch.PortGroupId.contains( $_.MoRef ) -and $VsSwitch.PortGroup.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VssPortObject = Add-VisioObjectVssPG $VssPortGroupObj $VssPort
					Draw_VssPort
					$ObjectNumber++
					$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
					$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

					if ( $debug -eq $true )`
					{ `
						$VssPortGroupNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Standard Port Group object " $VssPortGroupNumber " of " $VssPortGroupTotal " - " $VssPort.Name
					}
					Connect-VisioObject $VsSwitchObject $VssPortObject
					$y += 1.50
					
					foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $VssPort.VmId.contains( $_.MoRef ) -and $VssPort.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) -and $_.ClusterId -eq "" } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						if ( $VM.OS -eq "" ) `
						{ `
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
							$ObjectNumber++
							$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
							$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							if ( $VM.OS.contains("Microsoft") -eq $True ) `
							{ `
								$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
								Draw_VM
								$ObjectNumber++
								$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
								$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
							else `
							{ `
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
								$ObjectNumber++
								$VSSPortGroup_to_VM_Complete.Forecolor = "Blue"
								$VSSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($VSSPortGroup_to_VM_Complete)

								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
						}
						Connect-VisioObject $VssPortObject $VMObject
						$VssPortObject = $VMObject
					}
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VSSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VDS_to_Host
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Distributed Switch to Host Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Distributed Switch to VMHost Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VDS to Host"
	$DocsObj.Pages('VDS to Host')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VDS to Host')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes

	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VdSwitchNumber = 0
	$VdSwitchTotal = ( ( $VdSwitchImport ).VmHostId -split "," ).Count
	$VdsPortGroupNumber = 0
	$VdsPortGroupTotal = ( ( $VdsPortImport ).VmHostId -split "," ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VdSwitchTotal + $VdsPortGroupTotal + $vCenterImport.Name.Count
		
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VDS_to_Host_Complete.Forecolor = "Blue"
	$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VDS_to_Host_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VDS_to_Host_Complete.Forecolor = "Blue"
		$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VDS_to_Host_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VDS_to_Host_Complete.Forecolor = "Blue"
			$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VDS_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VDS_to_Host_Complete.Forecolor = "Blue"
				$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VDS_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					$ObjectNumber++
					$VDS_to_Host_Complete.Forecolor = "Blue"
					$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VDS_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
					}
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
					
					foreach ( $VdsPort in ( $VdsPortImport | Where-Object { $VdSwitch.PortGroupId.contains( $_.MoRef ) -and $VdSwitch.PortGroupName.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
					{ `
						$x += 2.50
						$VdsPortObject = Add-VisioObjectVdsPG $VdsPortGroupObj $VdsPort
						Draw_VdsPort
						$ObjectNumber++
						$VDS_to_Host_Complete.Forecolor = "Blue"
						$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VDS_to_Host_Complete)

						if ( $debug -eq $true )`
						{ `
							$VdsPortGroupNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Distributed Port Group object " $VdsPortGroupNumber " of " $VdsPortGroupTotal " - " $VdsPort.Name
						}
						Connect-VisioObject $VdSwitchObject $VdsPortObject
						$VdSwitchObject = $VdsPortObject
					}
				}
			}
		}

		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VDS_to_Host_Complete.Forecolor = "Blue"
			$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VDS_to_Host_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $VmHost.ClusterId -eq "" -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				$ObjectNumber++
				$VDS_to_Host_Complete.Forecolor = "Blue"
				$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VDS_to_Host_Complete)

				if ( $debug -eq $true )`
				{ `
					$VdSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
				}
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50

				foreach ( $VdsPort in ( $VdsPortImport | Where-Object { $VdSwitch.PortGroupId.contains( $_.MoRef ) -and $VdSwitch.PortGroupName.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name ) ) `
				{ `
					$x += 2.50
					$VdsPortObject = Add-VisioObjectVdsPG $VdsPortGroupObj $VdsPort
					Draw_VdsPort
					$ObjectNumber++
					$VDS_to_Host_Complete.Forecolor = "Blue"
					$VDS_to_Host_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VDS_to_Host_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdsPortGroupNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed Port Group object " $VdsPortGroupNumber " of " $VdsPortGroupTotal " - " $VdsPort.Name
					}
					Connect-VisioObject $VdSwitchObject $VdsPortObject
					$VdSwitchObject = $VdsPortObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VDS_to_Host >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VMK_to_VDS
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] VMkernel to Distributed Switch Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating VMkernel Virtual Distributed Switch Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VMK to VDS"
	$DocsObj.Pages('VMK to VDS')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VMK to VDS')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$VMHostNumber = 0
	$VMHostTotal = $VMHostImport.Name.Count
	$VdSwitchNumber = 0
	$VdSwitchTotal = ( ( $VdSwitchImport ).VmHostId -split "," ).Count
	$VdsVmkernelNumber = 0
	$VdsVmkernelTotal = $VdsVmkImport.Name.Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $VMHostTotal + $VdSwitchTotal + $VdsVmkernelTotal + $vCenterImport.Name.Count
	
	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VMK_to_VDS_Complete.Forecolor = "Blue"
	$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VMK_to_VDS_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VMK_to_VDS_Complete.Forecolor = "Blue"
		$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VMK_to_VDS_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$VMK_to_VDS_Complete.Forecolor = "Blue"
			$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VMK_to_VDS_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $VmHost in ( $VmHostImport | Where-Object { $Cluster.VmHostId.contains( $_.MoRef ) -and $Cluster.VmHost.contains( $_.Name ) -and $_.Cluster -notlike $null } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$HostObject = Add-VisioObjectHost $HostObj $VMHost
				Draw_VmHost
				$ObjectNumber++
				$VMK_to_VDS_Complete.Forecolor = "Blue"
				$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMK_to_VDS_Complete)

				if ( $debug -eq $true )`
				{ `
					$VMHostNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
				}
				Connect-VisioObject $ClusterObject $HostObject
				
				foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
					Draw_VdSwitch
					$ObjectNumber++
					$VMK_to_VDS_Complete.Forecolor = "Blue"
					$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMK_to_VDS_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdSwitchNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
					}
					Connect-VisioObject $HostObject $VdSwitchObject
					$y += 1.50
					
					foreach ( $VdsVmk in ( $VdsVmkImport | Sort-Object Name | Where-Object { $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.VSwitchId.contains( $VdSwitch.MoRef ) -and $_.VSwitch.contains( $VdSwitch.Name ) } ) ) `
					{ `
						$x += 2.50
						$VdsVmkNicObject = Add-VisioObjectVMK $VdsVmkNicObj $VdsVmk
						Draw_VdsVmk
						$ObjectNumber++
						$VMK_to_VDS_Complete.Forecolor = "Blue"
						$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VMK_to_VDS_Complete)

						if ( $debug -eq $true )`
						{ `
							$VdsVmkernelNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Virtual Distributed VMkernel object " $VdsVmkernelNumber " of " $VdsVmkernelTotal " - " $VdsVmk.Name
						}
						Connect-VisioObject $VdSwitchObject $VdsVmkNicObject
						$VdSwitchObject = $VdsVmkNicObject
					}
				}
			}
		}
		
		foreach ( $VmHost in ( $VmHostImport | Where-Object { $Datacenter.VmHostId.contains( $_.MoRef ) -and $Datacenter.VmHost.contains( $_.Name ) -and $_.ClusterId -eq "" }  | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$HostObject = Add-VisioObjectHost $HostObj $VMHost
			Draw_VmHost
			$ObjectNumber++
			$VMK_to_VDS_Complete.Forecolor = "Blue"
			$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VMK_to_VDS_Complete)

			if ( $debug -eq $true )`
			{ `
				$VMHostNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
			}
			Connect-VisioObject $DatacenterObject $HostObject
			
			foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $VmHost.vSwitchId.contains( $_.MoRef ) -and $VmHost.vSwitch.contains( $_.Name ) -and $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
				Draw_VdSwitch
				$ObjectNumber++
				$VMK_to_VDS_Complete.Forecolor = "Blue"
				$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VMK_to_VDS_Complete)

				if ( $debug -eq $true )`
				{ `
					$VdSwitchNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
				}
				Connect-VisioObject $HostObject $VdSwitchObject
				$y += 1.50
				
				foreach ( $VdsVmk in ( $VdsVmkImport | Sort-Object Name | Where-Object { $_.VmHostId.contains( $VmHost.MoRef ) -and $_.VmHost.contains( $VmHost.Name ) -and $_.VSwitchId.contains( $VdSwitch.MoRef ) -and $_.VSwitch.contains( $VdSwitch.Name ) -and $_.ClusterId -eq "" } ) ) `
				{ `
					$x += 1.50
					$VdsVmkNicObject = Add-VisioObjectVMK $VdsVmkNicObj $VdsVmk
					Draw_VdsVmk
					$ObjectNumber++
					$VMK_to_VDS_Complete.Forecolor = "Blue"
					$VMK_to_VDS_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($VMK_to_VDS_Complete)

					if ( $debug -eq $true )`
					{ `
						$VdsVmkernelNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Virtual Distributed VMkernel object " $VdsVmkernelNumber " of " $VdsVmkernelTotal " - " $VdsVmk.Name
					}
					Connect-VisioObject $VdSwitchObject $VdsVmkNicObject
					$VdSwitchObject = $VdsVmkNicObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VMK_to_VDS >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VDSPortGroup_to_VM
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Distributed Switch Port Group to VM Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Virtual Distributed Switch Port Group to Virtual Machine Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "VDSPortGroup to VM"
	$DocsObj.Pages('VDSPortGroup to VM')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('VDSPortGroup to VM')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$VdSwitchNumber = 0
	$VdSwitchTotal = $VdSwitchImport.Name.Count
	$VdsPortGroupNumber = 0
	$VdsPortGroupTotal = $VdsPortImport.Name.Count
	$VMNumber = 0
	$VMTotal = ( ( $VmImport | Where-Object { $_.PortGroupId -like "DistributedVirtualPortgroup*" -and $_.SRM.contains("placeholderVm") -eq $False } ).PortGroupId -split "," ).Count
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $VdSwitchTotal + $VdsPortGroupTotal + $VMTotal + $vCenterImport.Name.Count

	# Draw Objects
	$x = 0
	$y = 1.50
		
	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
	$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
		
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
		$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $VdSwitch in ( $VdSwitchImport | Where-Object { $Datacenter.vSwitchId.contains( $_.MoRef ) -and $Datacenter.vSwitch.contains( $_.Name ) } | Sort-Object Name -Descending -Unique ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$VdSwitchObject = Add-VisioObjectVdSwitch $VDSObj $VdSwitch
			Draw_VdSwitch
			$ObjectNumber++
			$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
			$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

			if ( $debug -eq $true )`
			{ `
				$VdSwitchNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Virtual Distributed Switch object " $VdSwitchNumber " of " $VdSwitchTotal " - " $VdSwitch.Name
			}
			Connect-VisioObject $DatacenterObject $VdSwitchObject

			foreach ( $VdsPort in ( $VdsPortImport | Where-Object { $VdSwitch.PortgroupId.contains( $_.MoRef ) -and $VdSwitch.PortgroupName.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$VdsPortObject = Add-VisioObjectVdsPG $VdsPortGroupObj $VdsPort
				Draw_VdsPort
				$ObjectNumber++
				$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
				$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

				if ( $debug -eq $true )`
				{ `
					$VdsPortGroupNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing Virtual Distributed Port Group object " $VdsPortGroupNumber " of " $VdsPortGroupTotal " - " $VdsPort.Name
				}
				Connect-VisioObject $VdSwitchObject $VdsPortObject
				$y += 1.50

				foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $VdsPort.VmId.contains( $_.MoRef ) -and $VdsPort.Vm.contains( $_.Name ) -and ( $_.SRM.contains("placeholderVm") -eq $False ) } ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
						$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
							$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$VDSPortGroup_to_VM_Complete.Forecolor = "Blue"
							$VDSPortGroup_to_VM_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($VDSPortGroup_to_VM_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}
					Connect-VisioObject $VdsPortObject $VMObject
					$VdsPortObject = $VMObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< VDSPortGroup_to_VM >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Cluster_to_DRS_Rule
{
	if ( $logdraw -eq $true ) `
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Cluster to DRS Rule Drawing selected." -ForegroundColor Magenta
	}
	$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	Write-Host "[$DateTime] Creating Cluster to DRS Rule Drawing." -ForegroundColor Green
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	CSV_In_Out
	
	$AppVisio = New-Object -ComObject Visio.InvisibleApp
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Add() | Out-Null
	$AppVisio.ActivePage.Name = "Cluster to DRS Rule"
	$DocsObj.Pages('Cluster to DRS Rule')
	$pagsObj = $AppVisio.ActiveDocument.Pages
	$pagObj = $pagsObj.Item('Cluster to DRS Rule')
	$AppVisio.ScreenUpdating = $False
	$AppVisio.EventsEnabled = $False
		
	# Load a set of stencils and select one to drop
	Visio_Shapes
		
	$vCenterTotal = $vCenterImport.Name.Count
	$DatacenterNumber = 0
	$DatacenterTotal = $DatacenterImport.Name.Count
	$ClusterNumber = 0
	$ClusterTotal = $ClusterImport.Name.Count	
	$DRSRuleNumber = 0
	$DRSRuleTotal = $DRSRuleImport.Name.Count
	$DrsVmHostRuleNumber = 0
	$DrsVmHostRuleTotal = $DrsVmHostImport.Name.Count
	$DrsClusterGroupNumber = 0
	$DrsClusterGroupTotal = $DrsClusterGroupImport.Name.Count
	$VMHostNumber = 0
	$VMHostGroup = $DrsClusterGroupImport | Where-Object { $_.GroupType -like "VMHostGroup" } 
	if ($null -eq $VMHostGroup)`
	{ `
		$VMHostTotal = 0
	} `
	else `
	{ `
		$VMHostTotal = ( $VMHostGroup.MemberId -split ", " ).Count
	}

	$VMNumber = 0
	$VMGroup = $DrsClusterGroupImport | Where-Object { $_.GroupType -like "VMGroup" } 
	if ($null -eq $VMGroup)`
	{ `
		$VMGroupTotal = 0
	} `
	else `
	{ `
		$VMGroupTotal = ( $VMGroup.MemberId -split ", " ).Count
	}
	$DRSRuleVM = ( $DRSRuleImport.VmId -split ", " ).Count 
	$VMTotal = $DRSRuleVM + $VMGroupTotal
	$ObjectNumber = 0
	$ObjectsTotal = $DatacenterTotal + $ClusterTotal + $DRSRuleTotal + $DrsVmHostRuleTotal + $DrsClusterGroupTotal + $VMHostTotal + $VMTotal + $vCenterTotal

	# Draw Objects
	$x = 0
	$y = 1.50

	$VCObject = Add-VisioObjectVC $VCObj $vCenterImport
	Draw_vCenter
	$ObjectNumber++
	$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
	$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
	$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Drawing vCenter object -" $vCenterImport.Name
	}
	
	foreach ( $Datacenter in ( $DatacenterImport | Sort-Object Name -Descending ) ) `
	{ `
		$x = 2.50
		$y += 1.50
		$DatacenterObject = Add-VisioObjectDC $DatacenterObj $Datacenter
		Draw_Datacenter
		$ObjectNumber++
		$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
		$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
		$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

		if ( $debug -eq $true )`
		{ `
			$DatacenterNumber++
			$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
			Write-Host "[$DateTime] Drawing Datacenter object " $DatacenterNumber " of " $DatacenterTotal " - " $Datacenter.Name
		}
		Connect-VisioObject $VCObject $DatacenterObject
		
		foreach ( $Cluster in ( $ClusterImport | Where-Object { $Datacenter.ClusterId.contains( $_.MoRef ) -and $Datacenter.Cluster.contains( $_.Name ) } | Sort-Object Name -Descending ) ) `
		{ `
			$x = 5.00
			$y += 1.50
			$ClusterObject = Add-VisioObjectCluster $ClusterObj $Cluster
			Draw_Cluster
			$ObjectNumber++
			$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
			$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
			$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

			if ( $debug -eq $true )`
			{ `
				$ClusterNumber++
				$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
				Write-Host "[$DateTime] Drawing Cluster" $ClusterNumber " of " $ClusterTotal " - " $Cluster.Name
			}
			Connect-VisioObject $DatacenterObject $ClusterObject
			
			foreach ( $DRSRule in ( $DrsRuleImport | Where-Object { $_.Cluster -eq $Cluster.Name -and $_.ClusterId -eq $Cluster.Moref } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DRSObject = Add-VisioObjectDrsRule $DRSRuleObj $DRSRule
				Draw_DrsRule
				$ObjectNumber++
				$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
				$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

				if ( $debug -eq $true )`
				{ `
					$DRSRuleNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing DRS Rule" $DRSRuleNumber " of " ( $DRSRuleTotal ) " - " $DRSRule.Name
				}
				Connect-VisioObject $ClusterObject $DRSObject
				$y += 1.50
				
				foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $DRSRule.VmId.contains( $_.MoRef ) -and $DRSRule.Vm.contains( $_.Name )} ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
						$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
							$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
							$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)

							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}	
					Connect-VisioObject $DRSObject $VMObject
					$DRSObject = $VMObject
				}
			}	

			foreach ( $DrsVmHostRule in ( $DrsVmHostImport | Where-Object { $_.Cluster -eq $Cluster.Name -and $_.ClusterId -eq $Cluster.Moref } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DRSVMHostRuleObject = Add-VisioObjectDRSVMHostRule $DRSVMHostRuleObj $DrsVmHostRule
				Draw_DrsVmHostRule
				$ObjectNumber++
				$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
				$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
				if ( $debug -eq $true )`
				{ `
					$DrsVmHostRuleNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing DRS Host Rule" $DrsVmHostRuleNumber " of " ( $DrsVmHostRuleTotal ) " - " $DrsVmHostRule.Name
				}
				Connect-VisioObject $ClusterObject $DRSVMHostRuleObject
				
				foreach ( $DrsClusterGroup in ( $DrsClusterGroupImport | Where-Object { $_.Name.contains( $DrsVmHostRule.VMHostGroup ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$DrsClusterGroupObject = Add-VisioObjectDrsClusterGroup $DRSClusterGroupObj $DrsClusterGroup
					Draw_DrsClusterGroup
					$ObjectNumber++
					$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
					$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
					if ( $debug -eq $true )`
					{ `
						$DrsClusterGroupNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing DRS Cluster Group" $DrsClusterGroupNumber " of " ( $DrsClusterGroupTotal ) " - " $DrsClusterGroup.Name
					}
					Connect-VisioObject $DRSVMHostRuleObject $DrsClusterGroupObject
					$y += 1.50
			
					foreach ( $VmHost in ( $VmHostImport | Sort-Object Name | Where-Object { $DrsClusterGroup.MemberId.contains( $_.MoRef ) -and $DrsClusterGroup.Member.contains( $_.Name ) } ) ) `
					{ `
						$x += 2.50
						$HostObject = Add-VisioObjectHost $HostObj $VMHost
						Draw_VmHost
						$ObjectNumber++
						$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
						$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
						if ( $debug -eq $true )`
						{ `
							$VMHostNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
						}
						Connect-VisioObject $DrsClusterGroupObject $HostObject
					}
					
				}
				
				foreach ( $DrsClusterGroup in ( $DrsClusterGroupImport | Where-Object { $_.Name.contains( $DrsVmHostRule.VMGroup ) } | Sort-Object Name -Descending ) ) `
				{ `
					$x = 10.00
					$y += 1.50
					$DrsClusterGroupObject = Add-VisioObjectDrsClusterGroup $DRSClusterGroupObj $DrsClusterGroup
					Draw_DrsClusterGroup
					$ObjectNumber++
					$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
					$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
					if ( $debug -eq $true )`
					{ `
						$DrsClusterGroupNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing DRS Cluster Group" $DrsClusterGroupNumber " of " ( $DrsClusterGroupTotal ) " - " $DrsClusterGroup.Name
					}
					Connect-VisioObject $DRSVMHostRuleObject $DrsClusterGroupObject
					$y += 1.50
					
					foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $DrsClusterGroup.MemberId.contains( $_.MoRef ) -and $DrsClusterGroup.Member.contains( $_.Name ) } ) ) `
					{ `
						$x += 2.50
						if ( $VM.OS -eq "" ) `
						{ `
							$VMObject = Add-VisioObjectVM $OtherObj $VM
							Draw_VM
							$ObjectNumber++
							$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
							$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							if ( $VM.OS.contains("Microsoft") -eq $True ) `
							{ `
								$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
								Draw_VM
								$ObjectNumber++
								$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
								$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
							else `
							{ `
								$VMObject = Add-VisioObjectVM $LinuxObj $VM
								Draw_VM
								$ObjectNumber++
								$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
								$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
								$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
								if ( $debug -eq $true )`
								{ `
									$VMNumber++
									$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
									Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
								}
							}
						}	
						Connect-VisioObject $DrsClusterGroupObject $VMObject
						$DrsClusterGroupObject = $VMObject
					}
				}
			}
			
			foreach ( $DrsClusterGroup in ( $DrsClusterGroupImport | Where-Object { $_.Cluster -eq $Cluster.Name -and $_.ClusterId -eq $Cluster.Moref -and $_.DrsVMHostRule -eq "" } | Sort-Object Name -Descending ) ) `
			{ `
				$x = 7.50
				$y += 1.50
				$DrsClusterGroupObject = Add-VisioObjectDrsClusterGroup $DRSClusterGroupObj $DrsClusterGroup
				Draw_DrsClusterGroup
				$ObjectNumber++
				$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
				$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
				$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
				if ( $debug -eq $true )`
				{ `
					$DrsClusterGroupNumber++
					$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
					Write-Host "[$DateTime] Drawing DRS Cluster Group" $DrsClusterGroupNumber " of " ( $DrsClusterGroupTotal ) " - " $DrsClusterGroup.Name
				}
				Connect-VisioObject $ClusterObject $DrsClusterGroupObject
				$y += 1.50
			
				foreach ( $VmHost in ( $VmHostImport | Sort-Object Name | Where-Object { $DrsClusterGroup.MemberId.contains( $_.MoRef ) -and $DrsClusterGroup.Member.contains( $_.Name ) } ) ) `
				{ `
					$x += 2.50
					$HostObject = Add-VisioObjectHost $HostObj $VMHost
					Draw_VmHost
					$ObjectNumber++
					$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
					$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
					$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
					if ( $debug -eq $true )`
					{ `
						$VMHostNumber++
						$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
						Write-Host "[$DateTime] Drawing Host" $VMHostNumber " of " $VMHostTotal " - " $VMHost.Name
					}
					Connect-VisioObject $DrsClusterGroupObject $HostObject
					$DrsClusterGroupObject  = $HostObject
				}
						
				foreach ( $VM in ( $VmImport | Sort-Object Name | Where-Object { $DrsClusterGroup.MemberId.contains( $_.MoRef ) -and $DrsClusterGroup.Member.contains( $_.Name ) } ) ) `
				{ `
					$x += 2.50
					if ( $VM.OS -eq "" ) `
					{ `
						$VMObject = Add-VisioObjectVM $OtherObj $VM
						Draw_VM
						$ObjectNumber++
						$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
						$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
						$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
						if ( $debug -eq $true )`
						{ `
							$VMNumber++
							$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
							Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
						}
					}
					else `
					{ `
						if ( $VM.OS.contains("Microsoft") -eq $True ) `
						{ `
							$VMObject = Add-VisioObjectVM $MicrosoftObj $VM
							Draw_VM
							$ObjectNumber++
							$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
							$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
						else `
						{ `
							$VMObject = Add-VisioObjectVM $LinuxObj $VM
							Draw_VM
							$ObjectNumber++
							$Cluster_to_DRS_Rule_Complete.Forecolor = "Blue"
							$Cluster_to_DRS_Rule_Complete.Text = "Object $ObjectNumber of $ObjectsTotal"
							$TabDraw.Controls.Add($Cluster_to_DRS_Rule_Complete)
			
							if ( $debug -eq $true )`
							{ `
								$VMNumber++
								$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
								Write-Host "[$DateTime] Drawing VM" $VMNumber " of " ( $VMTotal ) " - " $VM.Name
							}
						}
					}	
					Connect-VisioObject $DrsClusterGroupObject $VMObject
					$DrsClusterGroupObject = $VMObject
				}
			}
		}
	}

	# Resize to fit page
	$pagObj.ResizeToFitContents()
	if ( $debug -eq $true )`
	{ `
		$DateTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
		Write-Host "[$DateTime] Saving Visio drawing. Please disregard error " -ForegroundColor Yellow -NoNewLine; Write-Host '"There is a file sharing conflict.  The file cannot be accessed as requested."' -ForegroundColor Red -NoNewLine; Write-Host " below, this is an issue with Visio commands through PowerShell." -ForegroundColor Yellow
	}
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Quit()
}
#endregion ~~< Cluster_to_DRS_Rule >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Visio Pages Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Capture_Folder
{
	explorer.exe $CaptureCsvFolder
	$Host.UI.RawUI.WindowTitle = "vDiagram $MyVer"
}
#endregion ~~< Open_Capture_Folder >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Open_Final_Visio
{
	$SaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsd"
	$ConvertSaveFile = "$VisioFolder" + "\" + "$vCenter" + " VMware vDiagram - " + "$FileDateTime" + ".vsdx"
	$AppVisio = New-Object -ComObject Visio.Application
	$docsObj = $AppVisio.Documents
	$docsObj.Open($SaveFile) | Out-Null
	$AppVisio.ActiveDocument.Pages.Item(1).Delete(1) | Out-Null
	$AppVisio.Documents.SaveAs($SaveFile)
	$AppVisio.Documents.SaveAs($ConvertSaveFile) | Out-Null
	Remove-Item $SaveFile
	$Host.UI.RawUI.WindowTitle = "vDiagram $MyVer"
}
#endregion ~~< Open_Final_Visio >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Open Functions >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#endregion ~~< Event Handlers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$False | Out-Null

Main
<# 
.SYNOPSIS 
   vDiagram Scheduled Export

.DESCRIPTION
   vDiagram Scheduled Export

.NOTES 
   File Name	: vDiagram_Scheduled_Task_2.0.10.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 2.0.11

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data

.CHANGE LOG
	- 02/15/2021 - v2.0.11
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
		New drawing added for linked vCenters.
		
	- 04/06/2019 - v2.0.6
		New drawing added for VMs with snapshots.

	- 10/22/2018 - v2.0.5
		Dupliacte Resource Pools for same cluster were being drawn in Visio.

	- 10/22/2018 - v2.0.4
		Slight changes post presenting at Orlando VMUG UserCon
		Cleaned up global variables for CSVs & vCenter
		Changed date format of Visio file from yyyy_MM_dd-HH_MM to yyyy-MM-dd_HH-MM
		
	- 10/17/2018 - v2.0.3
		Fixed IP and MAC address capture on VMHost and VMs, not listing all IPs and MACs
	
	- 04/12/2018 - v2.0.1
		Added MAC Addresses to VMs & Templates

	- 04/11/2018 - v2.0.0
		Presented as a Community Theater Session at South Florida VMUG
		Feature enhancement requests collected
#>

#region Variables
$vCenter = "Replace with vCenter name."
$CaptureCsvFolder = "C:\vDiagram\Capture"
$SMTPSRV = "SMTP Server"
$EmailFrom = "outbound@email.com"
$EmailTo = "you@email.com"
$Subject = "vDiagram 2.0 Files"

# !!!!!!!!!!!! Comment out Line 1698 once .xml has been created !!!!!!!!!!!!

# Variables (no need to edit)
$Date = (Get-Date -format "yyyy-MM-dd")
$ScriptPath = (Get-Item (Get-Location)).FullName
$XMLFile = $ScriptPath + "\credentials.xml"
$ZipFile = "$ScriptPath" + "\vDiagram Files" + " " + "$Date.zip"
$AttachmentFile = $ZipFile
$EmailSubject = "vDiagram Files"
$Body = $EmailSubject
#endregion

#region Functions

#region PsCreds

#region Export-PSCredential
Function Export-PSCredential {
        param ( $Credential = (Get-Credential), $Path = "credentials.xml" )
 
        # Look at the object type of the $Credential parameter to determine how to handle it
        switch ( $Credential.GetType().Name ) {
                # It is a credential, so continue
                PSCredential            { continue }
                # It is a string, so use that as the username and prompt for the password
                String                          { $Credential = Get-Credential -credential $Credential }
                # In all other caess, throw an error and exit
                default                         { Throw "You must specify a credential object to export to disk." }
        }
       
        # Create temporary object to be serialized to disk
        $export = "" | Select-Object Username, EncryptedPassword
       
        # Give object a type name which can be identified later
        #$export.PSObject.TypeNames.Insert(0,’ExportedPSCredential’)
       
        $export.Username = $Credential.Username
 
        # Encrypt SecureString password using Data Protection API
        # Only the current user account can decrypt this cipher
        $export.EncryptedPassword = $Credential.Password | ConvertFrom-SecureString
 
        # Export using the Export-Clixml cmdlet
        $export | Export-Clixml $Path
        Write-Host -foregroundcolor Green "Credentials saved to: " -noNewLine
 
        # Return FileInfo object referring to saved credentials
        Get-Item $Path
}
#endregion Export-PSCredential

#region Import-PSCredential 
Function Import-PSCredential {
        param ( $Path = "credentials.xml" )
 
        # Import credential file
        $import = Import-Clixml $Path
       
        # Test for valid import
        if ( !$import.UserName -or !$import.EncryptedPassword ) {
                Throw "Input is not a valid ExportedPSCredential object, exiting."
        }
        $Username = $import.Username
       
        # Decrypt the password and store as a SecureString object for safekeeping
        $SecurePass = $import.EncryptedPassword | ConvertTo-SecureString
       
        # Build the new credential object
        $Credential = New-Object System.Management.Automation.PSCredential $Username, $SecurePass
        Write-Output $Credential
}
#endregion Import-PSCredential 

#endregion PsCreds

#region vCenterFunctions

#region ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Connect_vCenter
{
	Connect-VIServer $vCenter -Credential (Import-PSCredential -path $XMLFile)
}
#endregion

#region ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Disconnect_vCenter
{
	Disconnect-ViServer * -Confirm:$False
}
#endregion

#endregion

#region CsvExportFunctions

#region ~~< vCenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function vCenter_Export
{
	$vCenterExportFile = "$CaptureCsvFolder\$vCenter-vCenterExport.csv"
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
	$DatacenterExportFile = "$CaptureCsvFolder\$vCenter-DatacenterExport.csv"
	
	ForEach( $Datacenter in ( Get-View -ViewType Datacenter | Sort-Object Name ) ) `
	{ `
		$Datacenter | `
		Select-Object `
			@{ Name = "Name" ; Expression = { [string]::Join(", ", ( $_.Name ) ) } }, `
			@{ Name = "VmFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -Type VM | Where-Object { $_.MoRef -eq $_.VmFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "HostFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -Type HostAndCluster | Where-Object { $_.MoRef -eq $_.HostFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "DatastoreFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -Type Datastore | Where-Object { $_.MoRef -eq $_.DatastoreFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "NetworkFolder" ; Expression = { [string]::Join(", ", ( Get-Folder -Location $_.Name -Type Network | Where-Object { $_.MoRef -eq $_.NetworkFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "Cluster" ; Expression = { [string]::Join(", ", ( Get-Cluster -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "ClusterId" ; Expression = { [string]::Join(", ", ( Get-Cluster -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "VmHost" ; Expression = { [string]::Join(", ", ( Get-VMHost -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmHostId" ; Expression = { [string]::Join(", ", ( Get-VMHost -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-VM -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( Get-VM -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Template" ; Expression = { [string]::Join(", ", ( Get-Template -Location $_.Name | Sort-Object Name ) ) } }, `
			@{ Name = "TemplateId" ; Expression = { [string]::Join(", ", ( Get-Template -Location $_.Name | Sort-Object Name ).Id ) } }, `
			@{ Name = "Folder" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Datacenter | Where-Object { $_.MoRef -eq $_.DatacenterFolder } | Sort-Object Name ) ) } }, `
			@{ Name = "FolderId" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Datacenter ).Id ) } }, `
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
			@{ Name = "Parent" ; Expression = { [string]::Join(", ", ( Get-Folder -Type Datacenter | Where-Object { $_.MoRef -eq $_.Parent } | Sort-Object Name ) ) } }, `
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
	$ClusterExportFile = "$CaptureCsvFolder\$vCenter-ClusterExport.csv"
	
	ForEach( $Cluster in ( Get-View -ViewType ClusterComputeResource | Sort-Object Name ) ) `
	{ `
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
	$VmHostExportFile = "$CaptureCsvFolder\$vCenter-VmHostExport.csv"
	$ServiceInstance = Get-View ServiceInstance
	$LicenseManager = Get-View $ServiceInstance.Content.LicenseManager
	$LicenseAssignmentManager = Get-View $LicenseManager.LicenseAssignmentManager
	
	ForEach( $VmHost in ( Get-View -ViewType HostSystem | Sort-Object Name ) ) `
	{ `
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
		    @{ Name = "IScsiName" ; Expression = { [string]::Join( ", ", ( Get-VMHost $_.Name | Get-VMHostHBA -Type IScsi ).IScsiName ) } }, `
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
						ForEach-Object { ForEach ( $PhysicalNic in $_.NetworkInfo.Pnic ) `
						{ `
							$PnicsInfo = $_.QueryNetworkHint( $PhysicalNic.Device ) 
							ForEach ( $PnicInfo in $PnicsInfo ) `
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
	$VmExportFile = "$CaptureCsvFolder\$vCenter-VmExport.csv"
	
	ForEach( $VM in ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $False } | Sort-Object Name ) ) `
	{ `
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
	$TemplateExportFile = "$CaptureCsvFolder\$vCenter-TemplateExport.csv"
	
	ForEach( $Template in ( Get-View -ViewType VirtualMachine | Where-Object { $_.Config.Template -eq $True } | Sort-Object Name ) ) `
	{ `
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
	$DatastoreClusterExportFile = "$CaptureCsvFolder\$vCenter-DatastoreClusterExport.csv"

	ForEach( $DatastoreCluster in ( Get-View -ViewType StoragePod | Sort-Object Name ) ) `
	{ `
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
	$DatastoreExportFile = "$CaptureCsvFolder\$vCenter-DatastoreExport.csv"
	
	ForEach( $Datastore in ( Get-View -ViewType Datastore | Sort-Object Name ) ) `
	{ `
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
	$VsSwitchExportFile = "$CaptureCsvFolder\$vCenter-VsSwitchExport.csv"

	ForEach( $VsSwitch in ( Get-VirtualSwitch -Standard | Sort-Object Name ) ) `
	{ `
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
	$VssPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VssPortGroupExport.csv"
	
	ForEach ( $VMHost in Get-VMHost ) `
	{ `
		ForEach ( $VsSwitch in ( Get-VirtualSwitch -Standard -VMHost $VmHost ) ) `
		{ `
			ForEach( $VssPortGroup in ( Get-VirtualPortGroup -Standard -VirtualSwitch $VsSwitch | Sort-Object Name ) ) `
			{ `
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
	$VssVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VssVmkernelExport.csv"
	
	ForEach ( $VMHost in Get-VMHost ) `
	{ `
		ForEach ( $VsSwitch in ( Get-VirtualSwitch -VMHost $VmHost -Standard ) ) `
		{ `
			ForEach ( $VssPort in ( Get-VirtualPortGroup -Standard -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
				ForEach ( $VMHostNetworkAdapterVMKernel in ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VsSwitch -PortGroup $VssPort | Sort-Object Name ) ) `
				{ `
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
	$VssPnicExportFile = "$CaptureCsvFolder\$vCenter-VssPnicExport.csv"
	
	ForEach ( $VMHost in Get-VMHost ) `
	{ `
		ForEach ( $VsSwitch in ( Get-VirtualSwitch -Standard -VMHost $VmHost ) ) `
		{ `
			ForEach ( $VMHostNetworkAdapterUplink in ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VsSwitch -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
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
	$VdSwitchExportFile = "$CaptureCsvFolder\$vCenter-VdSwitchExport.csv"
	
	ForEach( $DistributedVirtualSwitch in ( Get-View -ViewType DistributedVirtualSwitch ) ) `
	{ `
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
	$VdsPortGroupExportFile = "$CaptureCsvFolder\$vCenter-VdsPortGroupExport.csv"
	
	ForEach( $DistributedVirtualPortgroup in ( Get-View -ViewType DistributedVirtualPortgroup ) ) `
	{ `
		$DistributedVirtualPortgroup | Sort-Object Name | Where-Object { $_.Name -notlike "*DVUplinks*" } | `
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
	$VdsVmkernelExportFile = "$CaptureCsvFolder\$vCenter-VdsVmkernelExport.csv"
	
	ForEach ( $VmHost in Get-VmHost ) `
	{ `
		ForEach ( $VdSwitch in ( Get-VdSwitch -VMHost $VmHost ) ) `
		{ `
			ForEach ( $VdsVmkernel in ( Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VdSwitch -VMHost $VmHost | Sort-Object -Property Name -Unique ) ) `
			{ `
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
	$VdsPnicExportFile = "$CaptureCsvFolder\$vCenter-VdsPnicExport.csv"
	
	ForEach ( $VmHost in Get-VmHost ) `
	{ `
		ForEach ( $VdSwitch in ( Get-VdSwitch -VMHost $VmHost ) ) `
		{ `
			ForEach ( $VMHostNetworkAdapterUplink in ( Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VdSwitch -VMHost $VmHost | Sort-Object Name ) ) `
			{ `
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
	$FolderExportFile = "$CaptureCsvFolder\$vCenter-FolderExport.csv"
	
	ForEach ( $Datacenter in Get-Datacenter ) `
	{ `
		ForEach ( $Folder in ( Get-Datacenter $Datacenter | Get-Folder | Get-View | Sort-Object ) ) `
		{ `
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
	$RdmExportFile = "$CaptureCsvFolder\$vCenter-RdmExport.csv"
	
	ForEach( $RDM in ( Get-VM | Get-HardDisk | Where-Object { $_.DiskType -like "Raw*" } | Sort-Object Parent ) ) `
	{ `
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
	$DrsRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsRuleExport.csv"
	
	ForEach ( $Cluster in Get-Cluster ) `
	{ `
		ForEach ( $DrsRule in ( Get-Cluster $Cluster | Get-DrsRule | Sort-Object Name) ) `
		{ `
			$DrsRule | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter -Cluster $Cluster.Name } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -Cluster $Cluster.Name ).Id } }, `
				@{ Name = "Cluster" ; Expression = { $_.Cluster } }, `
				@{ Name = "ClusterId" ; Expression = { ( $_.Cluster ).Id } }, `
				@{ Name = "Vm" ; Expression = { [string]::Join(", ", ( Get-VM -Id $_.VMIDs | Sort-Object Name ) ) } }, `
				@{ Name = "VmId" ; Expression = { [string]::Join(", ", ( $_.VMIDs | Sort Name ) ) } }, `
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
	$DrsClusterGroupExportFile = "$CaptureCsvFolder\$vCenter-DrsClusterGroupExport.csv"
	
	ForEach ( $Cluster in Get-Cluster ) `
	{ `
		ForEach ( $DrsClusterGroup in ( Get-DrsClusterGroup -Cluster $Cluster | Sort-Object Name ) ) `
		{ `
			$DrsClusterGroup | `
			Select-Object `
				@{ Name = "Name" ; Expression = { $_.Name } }, `
				@{ Name = "Datacenter" ; Expression = { Get-Datacenter -Cluster $Cluster.Name } }, `
				@{ Name = "DatacenterId" ; Expression = { ( Get-Datacenter -Cluster $Cluster.Name ).Id } }, `
				@{ Name = "Cluster" ; Expression = { $_.Cluster } }, `
				@{ Name = "ClusterId" ; Expression = { ( $_.Cluster ).Id } }, `
				@{ Name = "GroupType" ; Expression = { $_.GroupType } }, `
				@{ Name = "Member" ; Expression = { [string]::Join(", ", ( $_.Member ) ) } }, `
				@{ Name = "MemberId" ; Expression = { [string]::Join(", ", ( $_.Member ).Id ) } } | `
			Export-Csv $DrsClusterGroupExportFile -Append -NoTypeInformation
		}
	}
}
#endregion ~~< Drs_Cluster_Group_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#region ~~< Drs_VmHost_Rule_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Drs_VmHost_Rule_Export
{
	$DrsVmHostRuleExportFile = "$CaptureCsvFolder\$vCenter-DrsVmHostRuleExport.csv"
	
	ForEach ( $Cluster in Get-Cluster ) `
	{ `
		ForEach ( $DrsVmHostRule in ( Get-Cluster $Cluster | Get-DrsVmHostRule | Sort-Object Name ) ) `
		{ `
			$DrsVmHostRule | Sort-Object Name | `
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
	$ResourcePoolExportFile = "$CaptureCsvFolder\$vCenter-ResourcePoolExport.csv"
	
	ForEach( $ResourcePool in ( Get-View -ViewType ResourcePool | Sort-Object Name ) ) `
	{ `
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
	$SnapshotExportFile = "$CaptureCsvFolder\$vCenter-SnapshotExport.csv"
	
	ForEach( $Snapshot in ( Get-VM | Get-Snapshot | Sort-Object  VM, Created ) ) `
	{ `
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
	$LinkedvCenterExportFile = "$CaptureCsvFolder\$vCenter-LinkedvCenterExport.csv"
	Disconnect-ViServer * -Confirm:$false
	$global:vCenter = $VcenterTextBox.Text
	$User = $UserNameTextBox.Text
	Connect-VIServer $Vcenter -user $User -password $PasswordTextBox.Text -AllLinked
	
	if ( ( $global:DefaultVIServers ).Count -gt "1" ) `
	{ `
		ForEach ( $LinkedvCenter in ( $global:DefaultVIServers | Where-Object { $_.Name -ne "$vCenter" } ) ) `
		{ `
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

#endregion

#endregion

#region Export-PSCredential
Export-PSCredential
#endregion

#region Tasks
Connect_vCenter; vCenter_Export; Datacenter_Export; Cluster_Export; VmHost_Export; Vm_Export; Template_Export; DatastoreCluster_Export; Datastore_Export; VsSwitch_Export; VssPort_Export; VssVmk_Export; VssPnic_Export; VdSwitch_Export; VdsPort_Export; VdsVmk_Export; VdsPnic_Export; Folder_Export; Rdm_Export; Drs_Rule_Export; Drs_Cluster_Group_Export; Drs_VmHost_Rule_Export; Resource_Pool_Export; Snapshot_Export; Linked_vCenter_Export; Disconnect_vCenter
#endregion

#region Zip Files
Compress-Archive -U -Path $CaptureCsvFolder -DestinationPath $ZipFile
#endregion

#region Send E-mail
$msg = new-object Net.Mail.MailMessage
$att = new-object Net.Mail.Attachment($AttachmentFile)
$smtp = new-object Net.Mail.SmtpClient($SMTPSRV) 
$msg.From = $EmailFrom
$msg.To.Add($EmailTo)
$msg.Subject = $EmailSubject
$msg.Body = $Body
$msg.Attachments.Add($AttachmentFile) 
$smtp.Send($msg)
#endregion
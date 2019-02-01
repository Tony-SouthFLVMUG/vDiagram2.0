<# 
.SYNOPSIS 
   vDiagram Scheduled Export

.DESCRIPTION
   vDiagram Scheduled Export

.NOTES 
   File Name	: vDiagram_Scheduled_Task_2.0.4.ps1 
   Author		: Tony Gonzalez
   Author		: Jason Hopkins
   Based on		: vDiagram by Alan Renouf
   Version		: 2.0.4

.USAGE NOTES
	Ensure to unblock files before unzipping
	Ensure to run as administrator
	Required Files:
		PowerCLI or PowerShell 5.0 with PowerCLI Modules installed
		Active connection to vCenter to capture data

.CHANGE LOG
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

# !!!!!!!!!!!! Comment out Line 608 once .xml has been created !!!!!!!!!!!!

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
#region ~~< Connect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function Connect_vCenter
{
	Connect-VIServer $vCenter -Credential (Import-PSCredential -path $XMLFile)
}
#endregion
#region ~~< Disconnect_vCenter >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
	$global:DefaultVIServer | 
	Select-Object @{ N = "Name" ; E = { $_.Name } }, 
	@{ N = "Version" ; E = { $_.Version } }, 
	@{ N = "Build" ; E = { $_.Build } },
	@{ N = "OsType" ; E = { $_.ExtensionData.Content.About.OsType } } | Export-Csv $vCenterExportFile -Append -NoTypeInformation
}
#endregion
#region ~~< Datacenter_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Datacenter_Export
{
	$DatacenterExportFile = "$CaptureCsvFolder\$vCenter-DatacenterExport.csv"
	Get-Datacenter | Sort-Object Name | 
	Select-Object Name | Export-Csv $DatacenterExportFile -Append -NoTypeInformation
}
#endregion
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
#endregion
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
#endregion
#region ~~< Vm_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function Vm_Export
{
	$VmExportFile = "$CaptureCsvFolder\$vCenter-VmExport.csv"
	foreach ($Vm in(Get-View -ViewType VirtualMachine -Property Name, Config, Config.Tools, Guest, Guest.Net, Config.Hardware, Summary.Config, Config.DatastoreUrl, Parent, Runtime.Host -Server $vCenter | Sort-Object Name))
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
		@{ N = "SRM" ; E = { $_.Summary.Config.ManagedBy.Type } } | Export-Csv $VmExportFile -Append -NoTypeInformation
	}
}
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
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
#endregion
#endregion
#endregion

#region Export-PSCredential
Export-PSCredential
#endregion

#region Tasks
Connect_vCenter; vCenter_Export; Datacenter_Export; Cluster_Export; VmHost_Export; Vm_Export; Template_Export; DatastoreCluster_Export; Datastore_Export; VsSwitch_Export; VssPort_Export; VssVmk_Export; VssPnic_Export; VdSwitch_Export; VdsPort_Export; VdsVmk_Export; VdsPnic_Export; Folder_Export; Rdm_Export; Drs_Rule_Export; Drs_Cluster_Group_Export; Drs_VmHost_Rule_Export; Resource_Pool_Export; Disconnect_vCenter
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
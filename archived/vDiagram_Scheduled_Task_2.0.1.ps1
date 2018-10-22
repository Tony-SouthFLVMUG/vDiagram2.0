$Date = (Get-Date -format "yyyy_MM_dd")
$7zip = "C:\Program Files\7-Zip\7z.exe"

# Variables
$MainVC = "Replace with vCenter name where the vCenter you want to collect from is located."
$vCenterShortName = "Replace with vCenter name."
$CsvDir = "C:\vDiagram\Capture"
$SMTPserver = "SMTP Server"
$Mailfrom = "outbound@email.com"
$Mailto = "you@email.com"
$Subject = "vDiagram 2.0 Files"
$ReportFile = "C:\vDiagram\Zip"
$ZipFile = "$ReportFile"+"\vDiagram Files"+" "+"$Date.zip"
$AttachmentFile = $ZipFile
$VC_XMLFile = "C:\vDiagram\XML\credentials.xml"


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
Export-PSCredential
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
#region Connect_vCenter_Main
Function Connect_vCenter_Main
{
	$MainvCenter = Connect-VIServer $MainVC -Credential (Import-PSCredential -path $VC_XMLFile)
}
#endregion Connect_vCenter_Main

#region Connect_vCenter
Function Connect_vCenter
{
	$vCenter = Connect-VIServer $vCenterShortName -Credential (Import-PSCredential -path $VC_XMLFile)
}
#endregion Connect_vCenter

#region Disconnect_vCenter
Function Disconnect_vCenter
{
	Disconnect-ViServer * -Confirm:$False
}
#endregion Disconnect_vCenter
#endregion vCenterFunctions

#region CsvExportFunctions
#region vCenter_Export
Function vCenter_Export
{
	$vCenterExportFile = "$CsvDir\$vCenterShortName-vCenterExport.csv"
	Get-VM $vCenterShortName | Select @{N="Name";E={$_.Name}},
		@{N="Version";E={$global:DefaultVIServer.ExtensionData.Content.About.Version}},
		@{N="Build";E={$global:DefaultVIServer.ExtensionData.Content.About.Build}} | Export-CSV $vCenterExportFile -Append -NoTypeInfo
	Disconnect_vCenter
}
#endregion vCenter_Export

#region Datacenter_Export
Function Datacenter_Export
{
	$DatacenterExportFile = "$CsvDir\$vCenterShortName-DatacenterExport.csv"
	Get-Datacenter | Sort Name | Select Name | Export-CSV $DatacenterExportFile -Append -NoTypeInfo
}
#endregion Datacenter_Export

#region Cluster_Export
Function Cluster_Export
{
	$ClusterExportFile = "$CsvDir\$vCenterShortName-ClusterExport.csv"
	Get-Cluster | Sort Name | Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={Get-Cluster $_.Name | Get-Datacenter}},
		@{N="HAEnabled";E={$_.HAEnabled}},
		@{N="HAAdmissionControlEnabled";E={$_.HAAdmissionControlEnabled}},
		@{N="AdmissionControlPolicyCpuFailoverResourcesPercent";E={$_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.CpuFailoverResourcesPercent}},
		@{N="AdmissionControlPolicyMemoryFailoverResourcesPercent";E={$_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.MemoryFailoverResourcesPercent}},
		@{N="AdmissionControlPolicyFailoverLevel";E={$_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.FailoverLevel}},
		@{N="AdmissionControlPolicyAutoComputePercentages";E={$_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.AutoComputePercentages}},
		@{N="AdmissionControlPolicyResourceDarkCyanuctionToToleratePercent";E={$_.ExtensionData.configuration.dasconfig.AdmissionControlPolicy.ResourceDarkCyanuctionToToleratePercent}},
		@{N="DrsEnabled";E={$_.DrsEnabled}},
		@{N="DrsAutomationLevel";E={$_.DrsAutomationLevel}},
		@{N="VmMonitoring";E={$_.ExtensionData.configuration.dasconfig.VmMonitoring}},
		@{N="HostMonitoring";E={$_.ExtensionData.configuration.dasconfig.HostMonitoring}} | Export-CSV $ClusterExportFile -Append -NoTypeInfo
}
#endregion Cluster_Export

#region VmHost_Export
Function VmHost_Export
{
	$VmHostExportFile = "$CsvDir\$vCenterShortName-VmHostExport.csv"
	Get-View -ViewType HostSystem -Property Name, Config.Product, Summary.Hardware, Summary, Parent |
		Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={
		$Datacenter = Get-View -Id $_.Parent -Property Name,Parent
		While ($Datacenter -isnot [VMware.Vim.Datacenter] -and $Datacenter.Parent)
		{
			$Datacenter = Get-View -Id $Datacenter.Parent -Property Name,Parent
		}
		If($Datacenter -is [VMware.Vim.Datacenter])
		{
			$Datacenter.Name}}},
			@{N="Cluster";E={$Cluster = Get-View -Id $_.Parent -Property Name,Parent
		While ($Cluster -isnot [VMware.Vim.ClusterComputeResource] -and $Cluster.Parent)
		{
			$Cluster = Get-View -Id $Cluster.Parent -Property Name,Parent
		}
		If($Cluster -is [VMware.Vim.ClusterComputeResource]){$Cluster.Name}}},
		@{N="Version";E={$_.Config.Product.Version}},
		@{N="Build";E={$_.Config.Product.Build}},
		@{N="Manufacturer";E={$_.Summary.Hardware.Vendor}},
		@{N="Model";E={$_.Summary.Hardware.Model}},
		@{N="ProcessorType";E={$_.Summary.Hardware.CpuModel}},
		@{N="CpuMhz";E={$_.Summary.Hardware.CpuMhz}},
		@{N="NumCpuPkgs";E={$_.Summary.Hardware.NumCpuPkgs}},
		@{N="NumCpuCores";E={$_.Summary.Hardware.NumCpuCores}},
		@{N="NumCpuThreads";E={$_.Summary.Hardware.NumCpuThreads}},
		@{N="Memory";E={[math]::Round([decimal]$_.Summary.Hardware.MemorySize/1073741824)}},
		@{N="MaxEVCMode";E={$_.Summary.MaxEVCModeKey}},
		@{N="NumNics";E={$_.Summary.Hardware.NumNics}},
		@{N="IP";E={$_.Config.Network.Vnic.Spec.Ip.IpAddress}},
		@{N="NumHBAs";E={$_.Summary.Hardware.NumHBAs}} | Export-CSV $VmHostExportFile -Append -NoTypeInfo
}
#endregion VmHost_Export
	
#region Vm_Export
Function Vm_Export
{
	$VmExportFile = "$CsvDir\$vCenterShortName-VmExport.csv"
	ForEach($Vm in (Get-View -ViewType VirtualMachine -Property Name, Config, Config.Tools, Guest, Config.Hardware, Summary.Config, Config.DatastoreUrl, Parent, Runtime.Host -Server $vCenter | Sort Name))
	{
		$Folder = Get-View -Id $Vm.Parent -Property Name
		$Vm |
		Select Name,
		@{N="Datacenter";E={Get-Datacenter -VM $_.Name -Server $vCenter}},
		@{N="Cluster";E={Get-Cluster -VM $_.Name -Server $vCenter}},
		@{N="VmHost";E={Get-VmHost -VM $_.Name -Server $vCenter}},
		@{N="DatastoreCluster";E={Get-DatastoreCluster -VM $_.Name}},
		@{N="Datastore";E={$_.Config.DatastoreUrl.Name}},
		@{N="ResourcePool";E={Get-Vm $_.Name | Get-ResourcePool | ? {$_ -NotLike "Resources"}}},
		@{N="VsSwitch";E={Get-VirtualSwitch -VM $_.Name -Server $vCenter}},
		@{N="PortGroup";E={Get-VirtualPortGroup -VM $_.Name -Server $vCenter}},
		@{N="OS";E={$_.Config.GuestFullName}},
		@{N="Version";E={$_.Config.Version}},
		@{N="VMToolsVersion";E={$_.Guest.ToolsVersion}},
		@{N="ToolsVersionStatus";E={$_.Guest.ToolsVersionStatus}},
		@{N="ToolsStatus";E={$_.Guest.ToolsStatus}},
		@{N="ToolsRunningStatus";E={$_.Guest.ToolsRunningStatus}},
		@{N='Folder';E={$Folder.Name}},
		@{N="NumCPU";E={$_.Config.Hardware.NumCPU}},
		@{N="CoresPerSocket";E={$_.Config.Hardware.NumCoresPerSocket}},
		@{N="MemoryGB";E={[math]::Round([decimal]($_.Config.Hardware.MemoryMB/1024),0)}},
		@{N="IP";E={$_.Guest.IpAddress}},
		@{N="MacAddress";E={$_.Config.Hardware.Device.MacAddress}},
		@{N="ProvisionedSpaceGB";E={[math]::Round([decimal]($_.ProvisionedSpaceGB - $_.MemoryGB),0)}},
		@{N="NumEthernetCards";E={$_.Summary.Config.NumEthernetCards}},
		@{N="NumVirtualDisks";E={$_.Summary.Config.NumVirtualDisks}},
		@{N="CpuReservation";E={$_.Summary.Config.CpuReservation}},
		@{N="MemoryReservation";E={$_.Summary.Config.MemoryReservation}},
		@{N="SRM";E={$_.Summary.Config.ManagedBy.Type}} | Export-CSV $VmExportFile -Append -NoTypeInfo
	}
}
#endregion Vm_Export

#region  Template_Export
Function Template_Export
{
	$TemplateExportFile = "$CsvDir\$vCenterShortName-TemplateExport.csv"
	ForEach($VmHost in Get-Cluster | Get-VmHost)
	{
		Get-Template -Location $VmHost | Select @{N="Name";E={$_.Name}},
			@{N="Datacenter";E={$VmHost | Get-Datacenter}},
			@{N="Cluster";E={$VmHost | Get-Cluster}},
			@{N="VmHost";E={$VmHost.name}},
			@{N="Datastore";E={Get-Datastore -Id $_.DatastoreIdList}},
			@{N="Folder";E={Get-Folder -Id $_.FolderId}},
			@{N="OS";E={$_.ExtensionData.Config.GuestFullName}},
			@{N="Version";E={$_.ExtensionData.Config.Version}},
			@{N="ToolsVersion";E={$_.ExtensionData.Guest.ToolsVersion}},
			@{N="ToolsVersionStatus";E={$_.ExtensionData.Guest.ToolsVersionStatus}},
			@{N="ToolsStatus";E={$_.ExtensionData.Guest.ToolsStatus}},
			@{N="ToolsRunningStatus";E={$_.ExtensionData.Guest.ToolsRunningStatus}},
			@{N="NumCPU";E={$_.ExtensionData.Config.Hardware.NumCPU}},
			@{N="NumCoresPerSocket";E={$_.ExtensionData.Config.Hardware.NumCoresPerSocket}},
			@{N="MemoryGB";E={[math]::Round([decimal]$_.ExtensionData.Config.Hardware.MemoryMB/1024,0)}},
			@{N="MacAddress";E={$_.ExtensionData.Config.Hardware.Device.MacAddress}},
			@{N="NumEthernetCards";E={$_.ExtensionData.Summary.Config.NumEthernetCards}},
			@{N="NumVirtualDisks";E={$_.ExtensionData.Summary.Config.NumVirtualDisks}},
			@{N="CpuReservation";E={$_.ExtensionData.Summary.Config.CpuReservation}},
			@{N="MemoryReservation";E={$_.ExtensionData.Summary.Config.MemoryReservation}} | Export-CSV $TemplateExportFile -Append -NoTypeInfo
	}
}
#endregion Template_Export

#region DatastoreCluster_Export
Function DatastoreCluster_Export
{
	$DatastoreClusterExportFile = "$CsvDir\$vCenterShortName-DatastoreClusterExport.csv"
	Get-DatastoreCluster | Sort Name | Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={Get-DatastoreCluster $_.Name | Get-VmHost | Get-Datacenter}},
		@{N="Cluster";E={Get-DatastoreCluster $_.Name | Get-VmHost | Get-Cluster}},
		@{N="VmHost";E={Get-DatastoreCluster $_.Name | Get-VmHost}},
		@{N="SdrsAutomationLevel";E={$_.SdrsAutomationLevel}},
		@{N="IOLoadBalanceEnabled";E={$_.IoLoadBalanceEnabled}},
		@{N="CapacityGB";E={[math]::Round([decimal]$_.CapacityGB,0)}} | Export-CSV $DatastoreClusterExportFile -Append -NoTypeInfo
}
#endregion DatastoreCluster_Export

#region Datastore_Export
Function Datastore_Export
{
	$DatastoreExportFile = "$CsvDir\$vCenterShortName-DatastoreExport.csv"
	Get-Datastore | Select @{N="Name";E={$_.Name}},
	@{N="Datacenter";E={$_.Datacenter}},
	@{N="Cluster";E={Get-Datastore $_.Name | Get-VmHost | Get-Cluster}},
	@{N="DatastoreCluster";E={Get-DatastoreCluster -Datastore $_.Name}},
	@{N="VmHost";E={Get-VmHost -Datastore $_.Name}},
	@{N="Vm";E={Get-Datastore $_.Name | Get-Vm}},
	@{N="Type";E={$_.Type}},
	@{N="FileSystemVersion";E={$_.FileSystemVersion}},
	@{N="DiskName";E={$_.ExtensionData.Info.VMFS.Extent.DiskName}},
	@{N="StorageIOControlEnabled";E={$_.StorageIOControlEnabled}},
	@{N="CapacityGB";E={[math]::Round([decimal]$_.CapacityGB,0)}},
	@{N="FreeSpaceGB";E={[math]::Round([decimal]$_.FreeSpaceGB,0)}},
	@{N="Accessible";E={$_.State}},
	@{N="CongestionThresholdMillisecond";E={$_.CongestionThresholdMillisecond}}	| Export-CSV $DatastoreExportFile -Append -NoTypeInfo
}
#endregion Datastore_Export

#region VsSwitch_Export
Function VsSwitch_Export
{
	$VsSwitchExportFile = "$CsvDir\$vCenterShortName-VsSwitchExport.csv"
	Get-VirtualSwitch -Standard | Sort Name | Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={Get-Datacenter -VmHost $_.VmHost}},
		@{N="Cluster";E={Get-Cluster -VmHost $_.VmHost}},
		@{N="VmHost";E={$_.VmHost}},
		@{N="Nic";E={$_.Nic}},
		@{N="NumPorts";E={$_.ExtensionData.Spec.NumPorts}},
		@{N="AllowPromiscuous";E={$_.ExtensionData.Spec.Policy.Security.AllowPromiscuous}},
		@{N="MacChanges";E={$_.ExtensionData.Spec.Policy.Security.MacChanges}},
		@{N="ForgedTransmits";E={$_.ExtensionData.Spec.Policy.Security.ForgedTransmits}},
		@{N="Policy";E={$_.ExtensionData.Spec.Policy.NicTeaming.Policy}},
		@{N="ReversePolicy";E={$_.ExtensionData.Spec.Policy.NicTeaming.ReversePolicy}},
		@{N="NotifySwitches";E={$_.ExtensionData.Spec.Policy.NicTeaming.NotifySwitches}},
		@{N="RollingOrder";E={$_.ExtensionData.Spec.Policy.NicTeaming.RollingOrder}},
		@{N="ActiveNic";E={$_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic}},
		@{N="StandbyNic";E={$_.ExtensionData.Spec.Policy.NicTeaming.NicOrder.StandbyNic}} | Export-CSV $VsSwitchExportFile -Append -NoTypeInfo
}
#endregion VsSwitch_Export

#region VssPort_Export
Function VssPort_Export
{
	$VssPortGroupExportFile = "$CsvDir\$vCenterShortName-VssPortGroupExport.csv"
	ForEach ($VMHost in Get-VMHost)
	{
		ForEach($VsSwitch in (Get-VirtualSwitch -Standard -VMHost $VmHost))
		{
			Get-VirtualPortGroup -Standard -VirtualSwitch $VsSwitch | Sort Name | Select @{N="Name";E={$_.Name}},
				@{N="Datacenter";E={Get-Datacenter -VMHost $VMHost.Name}},
				@{N="Cluster";E={Get-Cluster -VMHost $VMHost.Name}},
				@{N="VmHost";E={$VMHost.Name}},
				@{N="VsSwitch";E={$VsSwitch.Name}},
				@{N="VLanId";E={$_.VLanId}},
				@{N="ActiveNic";E={$_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.ActiveNic}},
				@{N="StandbyNic";E={$_.ExtensionData.ComputedPolicy.NicTeaming.NicOrder.StandbyNic}} | Export-CSV $VssPortGroupExportFile -Append -NoTypeInfo
		}
	}
}
#endregion VssPort_Export

#region VssVmk_Export
Function VssVmk_Export
{
	$VssVmkernelExportFile = "$CsvDir\$vCenterShortName-VssVmkernelExport.csv"
	ForEach ($VMHost in Get-VMHost)
	{
		ForEach($VsSwitch in (Get-VirtualSwitch -VMHost $VmHost -Standard))
		{
			ForEach($VssPort in (Get-VirtualPortGroup -Standard -VMHost $VmHost | Sort Name))
			{
				Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VsSwitch -PortGroup $VssPort | Sort Name | Select @{N="Name";E={$_.Name}},
					@{N="Datacenter";E={Get-Datacenter -VMHost $VMHost.Name}},
					@{N="Cluster";E={Get-Cluster -VMHost $VMHost.Name}},
					@{N="VmHost";E={$VMHost.Name}},
					@{N="VsSwitch";E={$VsSwitch.Name}},
					@{N="PortGroupName";E={$_.PortGroupName}},
					@{N="DhcpEnabled";E={$_.DhcpEnabled}},
					@{N="IP";E={$_.IP}},
					@{N="Mac";E={$_.Mac}},
					@{N="ManagementTrafficEnabled";E={$_.ManagementTrafficEnabled}},
					@{N="VMotionEnabled";E={$_.VMotionEnabled}},
					@{N="FaultToleranceLoggingEnabled";E={$_.FaultToleranceLoggingEnabled}},
					@{N="VsanTrafficEnabled";E={$_.VsanTrafficEnabled}},
					@{N="Mtu";E={$_.Mtu}} | Export-CSV $VssVmkernelExportFile -Append -NoTypeInfo
			}
		}
	}
}
#endregion VssVmk_Export

#region VssPnic_Export
Function VssPnic_Export
{
	$VssPnicExportFile = "$CsvDir\$vCenterShortName-VssPnicExport.csv"
	ForEach ($VMHost in Get-VMHost)
	{
		ForEach($VsSwitch in (Get-VirtualSwitch -Standard -VMHost $VmHost))
		{
			Get-VMHostNetworkAdapter -Physical -VirtualSwitch $VsSwitch -VMHost $VmHost | Sort Name | Select @{N="Name";E={$_.Name}},
				@{N="Datacenter";E={Get-Datacenter -VmHost $VmHost}},
				@{N="Cluster";E={Get-Cluster -VmHost $_.VmHost}},
				@{N="VmHost";E={$_.VmHost}},
				@{N="VsSwitch";E={$VsSwitch.Name}},
				@{N="Mac";E={$_.Mac}} | Export-CSV $VssPnicExportFile -Append -NoTypeInfo
		}
	}
}
#endregion VssPnic_Export

#region VdSwitch_Export
Function VdSwitch_Export
{
	$VdSwitchExportFile = "$CsvDir\$vCenterShortName-VdSwitchExport.csv"
	ForEach ($VmHost in Get-VmHost)
	{
		Get-VdSwitch -VMHost $VmHost | Select @{N="Name";E={$_.Name}},
			@{N="Datacenter";E={$_.Datacenter}},
			@{N="Cluster";E={Get-Cluster -VMHost $VMHost.name}},
			@{N="VmHost";E={$VMHost.Name}},
			@{N="Vendor";E={$_.Vendor}},
			@{N="Version";E={$_.Version}},
			@{N="NumUplinkPorts";E={$_.NumUplinkPorts}},
			@{N="UplinkPortName";E={$_.ExtensionData.Config.UplinkPortPolicy.UplinkPortName}},
			@{N="Mtu";E={$_.Mtu}} | Export-CSV $VdSwitchExportFile -Append -NoTypeInfo
	}
}
#endregion VdSwitch_Export

#region VdsPort_Export
Function VdsPort_Export
{
	$VdsPortGroupExportFile = "$CsvDir\$vCenterShortName-VdsPortGroupExport.csv"
	ForEach ($VmHost in Get-VmHost)
	{
		ForEach ($VdSwitch in (Get-VdSwitch -VMHost $VmHost | Sort -Property ConnectedEntity -Unique))
		{
			Get-VDPortGroup | Sort Name | ? {$_.Name -NotLike "*DVUplinks*"} | Select @{N="Name";E={$_.Name}},
				@{N="Datacenter";E={Get-Datacenter -VMHost $VMHost.name}},
				@{N="Cluster";E={Get-Cluster -VMHost $VMHost.name}},
				@{N="VmHost";E={$VMHost.Name}},
				@{N="VlanConfiguration";E={$_.VlanConfiguration}},
				@{N="VdSwitch";E={$_.VdSwitch}},
				@{N="NumPorts";E={$_.NumPorts}},
				@{N="ActiveUplinkPort";E={$_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.ActiveUplinkPort}},
				@{N="StandbyUplinkPort";E={$_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.UplinkPortOrder.StandbyUplinkPort}},
				@{N="Policy";E={$_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value}},
				@{N="ReversePolicy";E={$_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.ReversePolicy.Value}},
				@{N="NotifySwitches";E={$_.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.NotifySwitches.Value}},
				@{N="PortBinding";E={$_.PortBinding}} | Export-CSV $VdsPortGroupExportFile -Append -NoTypeInfo
		}
	}
}
#endregion VdsPort_Export

#region VdsVmk_Export
Function VdsVmk_Export
{
	$VdsVmkernelExportFile = "$CsvDir\$vCenterShortName-VdsVmkernelExport.csv"
	ForEach ($VmHost in Get-VmHost)
	{
		ForEach ($VdSwitch in (Get-VdSwitch -VMHost $VmHost))
		{
			Get-VMHostNetworkAdapter -VMKernel -VirtualSwitch $VdSwitch -VMHost $VmHost   | Sort -Property Name -Unique | Select @{N="Name";E={$_.Name}},
				@{N="Datacenter";E={Get-Datacenter -VMHost $VMHost.name}},
				@{N="Cluster";E={Get-Cluster -VMHost $VMHost.name}},
				@{N="VmHost";E={$VMHost.Name}},
				@{N="VdSwitch";E={$VdSwitch.Name}},
				@{N="PortGroupName";E={$_.PortGroupName}},
				@{N="DhcpEnabled";E={$_.DhcpEnabled}},
				@{N="IP";E={$_.IP}},
				@{N="Mac";E={$_.Mac}},
				@{N="ManagementTrafficEnabled";E={$_.ManagementTrafficEnabled}},
				@{N="VMotionEnabled";E={$_.VMotionEnabled}},
				@{N="FaultToleranceLoggingEnabled";E={$_.FaultToleranceLoggingEnabled}},
				@{N="VsanTrafficEnabled";E={$_.VsanTrafficEnabled}},
				@{N="Mtu";E={$_.Mtu}} | Export-CSV $VdsVmkernelExportFile -Append -NoTypeInfo
		
		}
	}
}
#endregion VdsVmk_Export

#region VdsPnic_Export
Function VdsPnic_Export
{
	$VdsPnicExportFile = "$CsvDir\$vCenterShortName-VdsPnicExport.csv"
	ForEach ($VmHost in Get-VmHost)
	{
		ForEach ($VdSwitch in (Get-VdSwitch -VMHost $VmHost))
		{
			Get-VDPort -VdSwitch $VdSwitch -Uplink | Sort -Property ConnectedEntity -Unique | Select @{N="Name";E={$_.ConnectedEntity}},
				@{N="Datacenter";E={Get-Datacenter -VMHost $VMHost.name}},
				@{N="Cluster";E={Get-Cluster -VMHost $VMHost.name}},
				@{N="VmHost";E={$VMHost.Name}},
				@{N="VdSwitch";E={$VdSwitch}},
				@{N="Portgroup";E={$_.Portgroup}},
				@{N="ConnectedEntity";E={$_.Name}},
				@{N="VlanConfiguration";E={$_.VlanConfiguration}} | Export-CSV $VdsPnicExportFile -Append -NoTypeInfo
		}
	}
}
#endregion VdsPnic_Export

#region Folder_Export
Function Folder_Export
{
	$FolderExportFile = "$CsvDir\$vCenterShortName-FolderExport.csv"
	ForEach ($Datacenter in Get-Datacenter)
	{
		Get-Folder -Location $Datacenter -Type VM | Sort Name | Select @{N="Name";E={$_.Name}},
			@{N="Datacenter";E={$Datacenter.Name}} | Export-CSV $FolderExportFile -Append -NoTypeInfo
	}
}
#endregion Folder_Export

#region Rdm_Export
Function Rdm_Export
{
	$RdmExportFile = "$CsvDir\$vCenterShortName-RdmExport.csv"
	Get-VM | Get-HardDisk | ? {$_.DiskType -Like "Raw*"}| Sort Parent | Select @{N="ScsiCanonicalName";E={$_.ScsiCanonicalName}},
		@{N="Cluster";E={Get-Cluster -VM $_.Parent}},
		@{N="Vm";E={$_.Parent}},
		@{N="Label";E={$_.Name}},
		@{N="CapacityGB";E={[math]::Round([decimal]$_.CapacityGB,2)}},
		@{N="DiskType";E={$_.DiskType}},
		@{N="Persistence";E={$_.Persistence}},
		@{N="CompatibilityMode";E={$_.ExtensionData.Backing.CompatibilityMode}},
		@{N="DeviceName";E={$_.ExtensionData.Backing.DeviceName}},
		@{N="Sharing";E={$_.ExtensionData.Backing.Sharing}} | Export-CSV $RdmExportFile -Append -NoTypeInfo
}
#endregion Rdm_Export

#region Drs_Rule_Export
Function Drs_Rule_Export
{
	$DrsRuleExportFile = "$CsvDir\$vCenterShortName-DrsRuleExport.csv"
	ForEach ($Cluster in Get-Cluster)
	{
		Get-Cluster $Cluster | Get-DrsRule | Sort Name | Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={Get-Datacenter -Cluster $Cluster.Name}},
		@{N="Cluster";E={$_.Cluster}},
		@{N="Type";E={$_.Type}},
		@{N="Enabled";E={$_.Enabled}},
		@{N="Mandatory";E={$_.Mandatory}} | Export-CSV $DrsRuleExportFile -Append -NoTypeInfo
	}
}
#endregion Drs_Rule_Export

#region Drs_Cluster_Group_Export
Function Drs_Cluster_Group_Export
{
	$DrsClusterGroupExportFile = "$CsvDir\$vCenterShortName-DrsClusterGroupExport.csv"
	ForEach ($Cluster in Get-Cluster)
	{
		Get-Cluster $Cluster | Get-DrsClusterGroup | Sort Name | Select @{N="Name";E={$_.Name}},
		@{N="Datacenter";E={Get-Datacenter -Cluster $Cluster.Name}},
		@{N="Cluster";E={$_.Cluster}},
		@{N="GroupType";E={$_.GroupType}},
		@{N="Member";E={$_.Member}} | Export-CSV $DrsClusterGroupExportFile -Append -NoTypeInfo
	}
}
#endregion Drs_Cluster_Group_Export

#region Drs_VmHost_Rule_Export
Function Drs_VmHost_Rule_Export
{
	$DrsVmHostRuleExportFile = "$CsvDir\$vCenterShortName-DrsVmHostRuleExport.csv"
	ForEach ($Cluster in Get-Cluster)
	{
		ForEach ($DrsClusterGroup in (Get-Cluster $Cluster | Get-DrsClusterGroup | Sort Name))
		{
			Get-Cluster $Cluster | Get-DrsVmHostRule | Sort Name | Select @{N="Name";E={$_.Name}},
				@{N="Datacenter";E={Get-Datacenter -Cluster $Cluster.Name}},
				@{N="Cluster";E={$_.Cluster}},
				@{N="Enabled";E={$_.Enabled}},
				@{N="Type";E={$_.Type}},
				@{N="VMGroup";E={$_.VMGroup}},
				@{N="VMHostGroup";E={$_.VMHostGroup}},
				@{N="AffineHostGroupName";E={$_.ExtensionData.AffineHostGroupName}},
				@{N="AntiAffineHostGroupName";E={$_.ExtensionData.AntiAffineHostGroupName}}	| Export-CSV $DrsVmHostRuleExportFile -Append -NoTypeInfo
		}
	}
}
#endregion Drs_VmHost_Rule_Export

#region Resource_Pool_Export
Function Resource_Pool_Export
{
	$ResourcePoolExportFile = "$CsvDir\$vCenterShortName-ResourcePoolExport.csv"
	ForEach ($Cluster in Get-Cluster)
	{
		ForEach ($ResourcePool in (Get-Cluster $Cluster | Get-ResourcePool | ?{$_.Name -ne "Resources"} | Sort Name))
		{
			Get-ResourcePool  $ResourcePool | Sort Name | Select @{N="Name";E={$_.Name}},
				@{N="Cluster";E={$Cluster.Name}},
				@{N="CpuSharesLevel";E={$_.CpuSharesLevel}},
				@{N="NumCpuShares";E={$_.NumCpuShares}},
				@{N="CpuReservationMHz";E={$_.CpuReservationMHz}},
				@{N="CpuExpandableReservation";E={$_.CpuExpandableReservation}},
				@{N="CpuLimitMHz";E={$_.CpuLimitMHz}},
				@{N="MemSharesLevel";E={$_.MemSharesLevel}},
				@{N="NumMemShares";E={$_.NumMemShares}},
				@{N="MemReservationGB";E={$_.MemReservationGB}},
				@{N="MemExpandableReservation";E={$_.MemExpandableReservation}},
				@{N="MemLimitGB";E={$_.MemLimitGB}}	| Export-CSV $ResourcePoolExportFile -Append -NoTypeInfo
		}
	}
}
#endregion Resource_Pool_Export

#endregion CsvExportFunctions

#endregion Functions

Connect_vCenter_Main; vCenter_Export; Connect_vCenter; Datacenter_Export; Cluster_Export; VmHost_Export; Vm_Export; Template_Export; DatastoreCluster_Export; Datastore_Export; VsSwitch_Export; VssPort_Export; VssVmk_Export; VssPnic_Export; VdSwitch_Export; VdsPort_Export; VdsVmk_Export; VdsPnic_Export; Folder_Export; Rdm_Export; Drs_Rule_Export; Drs_Cluster_Group_Export; Drs_VmHost_Rule_Export; Resource_Pool_Export; Disconnect_vCenter

#Zip Files
cd $CsvDir
dir *.csv | ForEach-Object { & $7zip a -tzip ($ZipFile) $CsvDir\*.csv }

#Send E-mail
Send-MailMessage -To $Mailto -Subject $Subject -SmtpServer $SMTPserver -From $Mailfrom -Attachments $AttachmentFile

#Clear CSV Folder
cd $CsvDir
del *.csv
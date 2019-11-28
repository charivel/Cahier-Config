<#
    .NOTES
    ===========================================================================
     Créé par:    Christophe HARIVEL
     Blog:          www.vrun.fr
     Twitter:       @harivchr
        ===========================================================================
    .DESCRIPTION
        L'objectif de ce script est de créer un cahier de configuration VMware au format excel
    .NOTICE
        Penser a modifier les variables dans la rubriques "VARIABLES A MODIFIER"
		Puis lancer ./Cahier_Config.ps1
#>

#################### VARIABLES A MODFIER ##################################################

# Nom du vCenter
$VCENTER = "nom_ou_IP_du_vCenter"

# Chemin du répertoire de destination
$path = "C:\Export\Cahier_Config"

# Nom du fichier final
$XLSXfilename = "Cahier_Tech_" + $VCENTER + "_" + (Get-Date -Format "yyyy-MM-dd") + ".xlsx"

# Nom de chaque fichier temporaire CSV (un fichier CSV pour chaque onglet du fichier XLSX final)
$VCENTERfilename = "01_VCENTER_Infos.csv"
$CLUSfilename = "02_Esxi_Clusters_Infos.csv"
$ESXifilename = "03_ESXi_Infos.csv"
$VMKfilename = "04_VMKernel_Infos.csv"
$DSCLUSfilename = "05_Datastores_Clusters_Infos.csv"
$DSfilename = "06_Datastores_Infos.csv"
$SWfilename = "07_vSwitches_Infos.csv"
$PERMfilename = "08_Permissions_Infos.csv"
$ROLESfilename = "09_Roles_Infos.csv"



#################### DECLARATION DES FONCTIONS ############################################

# Fonction pour récupérer toutes les informations sur le vCenter
Function Get-VCenter-Infos {
	write-host "Collecte des infos vCenter en cours" -NoNewLine
	
	$VC = $global:DefaultVIServers
	$VCip = ([System.Net.Dns]::GetHostEntry($VC.Name)).AddressList
	$VCip = ($VCip -join " - ")
	
	$tableau = @()
	$Object = new-object PSObject
	$Object | add-member -name "vCenter Name" -membertype Noteproperty -value $VC.Name
	$Object | add-member -name "Version" -membertype Noteproperty -value $VC.Version
	$Object | add-member -name "Build" -membertype Noteproperty -value $VC.Build
	$Object | add-member -name "OS" -membertype Noteproperty -value $VC.ExtensionData.Content.About.OsType
	$Object | add-member -name "Mgmt IP" -membertype Noteproperty -value $VCip
	
    $vcHAClusterConfig = Get-View failoverClusterConfigurator
    $vcHAConfig = $vcHAClusterConfig.getVchaConfig()

    $vcHAState = $vcHAConfig.State
    switch($vcHAState) {
        configured {
            $vcHAClusterManager = Get-View failoverClusterManager
            $vcHAMode = $vcHAClusterManager.getClusterMode()
			$gw = $vcHAConfig.FailoverNodeInfo1.FailoverIp.Gateway
			$gw = ($gw -join " - ")
			
			$Object | add-member -name "Mgmt Mask" -membertype Noteproperty -value $vcHAConfig.FailoverNodeInfo1.FailoverIp.SubnetMask
			$Object | add-member -name "Mgmt GW" -membertype Noteproperty -value $gw
			$Object | add-member -name "VCHA State" -membertype Noteproperty -value $vcHAState
			$Object | add-member -name "VCHA Mode" -membertype Noteproperty -value $vcHAMode	
			$Object | add-member -name "VCHA ActiveIP" -membertype Noteproperty -value $vcHAConfig.FailoverNodeInfo1.ClusterIpSettings.Ip.IpAddress
			$Object | add-member -name "VCHA PassiveIP" -membertype Noteproperty -value $vcHAConfig.FailoverNodeInfo2.ClusterIpSettings.Ip.IpAddress
			$Object | add-member -name "VCHA WitnessIP" -membertype Noteproperty -value $vcHAConfig.WitnessNodeInfo.IpSettings.Ip.IpAddress
            
			;break
        }
        invalid { Write-Host -ForegroundColor Red "VCHA State is in invalid state ...";break}
        notConfigured { Write-Host "VCHA is not configured";break}
        prepared { Write-Host "VCHA is being prepared, please try again in a little bit ...";break}
    }
	
	$tableau += $Object
	return $tableau
}

# Fonction pour récupérer toutes les informations sur les clusters ESXi
function Get-Clusters-Infos{
	write-host "Collecte des infos clusters ESXi en cours" -NoNewLine
	
	$CLUS = get-cluster

	$tableau = @()
	
	foreach	($clu in $CLUS){
		write-host "." -NoNewLine
		$clusname = $clu.Name
		$HAEnabled = $clu.HAEnabled
		if($HAEnabled -eq "true"){
			$HAEnabled = "Enabled"
		}else{
			$HAEnabled = "Disabled"
		}
		$HAAdmissionControlEnabled = $clu.HAAdmissionControlEnabled
		if($HAAdmissionControlEnabled -eq "true"){
			$HAAdmissionControlEnabled = "Enabled"
		}else{
			$HAAdmissionControlEnabled = "Disabled"
		}
		$HAFailoverLevel = $clu.HAFailoverLevel
		$HAAutoComputePercentages = ($clu | Get-view).Configuration.DasConfig.AdmissionControlPolicy.AutoComputePercentages
		if($HAAutoComputePercentages -eq "true"){
			$HAAutoComputePercentages = "Enabled"
		}else{
			$HAAutoComputePercentages = "Disabled"
		}
		$HAResourceReductionToToleratePercent = ($clu | Get-view).Configuration.DasConfig.AdmissionControlPolicy.ResourceReductionToToleratePercent
		$HAFailoverLevel = $clu.HAFailoverLevel
		$HAIsolationResponse = $clu.HAIsolationResponse
		$HARestartPriority = $clu.HARestartPriority
		$HAHBCandidateDatastorePolicy = ($clu | Get-view).Configuration.DasConfig.HBDatastoreCandidatePolicy
		$HAAPDConfig =  ($clu | Get-view).Configuration.DasConfig.DefaultVmSettings.VmComponentProtectionSettings.VmStorageProtectionForAPD
		$HAPDLConfig =  ($clu | Get-view).Configuration.DasConfig.DefaultVmSettings.VmComponentProtectionSettings.VmStorageProtectionForPDL
		$HAVmMonitoring =  ($clu | Get-view).Configuration.DasConfig.VmMonitoring
		$DRSEnabled = $clu.DRSEnabled
		if($DRSEnabled -eq "true"){
			$DRSEnabled = "Enabled"
		}else{
			$DRSEnabled = "Disabled"
		}
		$DrsAutomationLevel = $clu.DrsAutomationLevel
		$DRSLevel = ($clu | Get-view).Configuration.DrsConfig.VmotionRate
		$EVCMode = $clu.EVCMode
		$VMSwapfilePolicy = $clu.VMSwapfilePolicy
		
				
		$Object = new-object PSObject
		$Object | add-member -name "Cluster Name" -membertype Noteproperty -value $clusname
		$Object | add-member -name "HA Enabled" -membertype Noteproperty -value $HAEnabled
		$Object | add-member -name "HA Admission Control Enabled" -membertype Noteproperty -value $HAAdmissionControlEnabled		
		$Object | add-member -name "HA Failover Level" -membertype Noteproperty -value $HAFailoverLevel		
		$Object | add-member -name "HA Auto Compute Percentages" -membertype Noteproperty -value $HAAutoComputePercentages		
		$Object | add-member -name "HA Resource Reduction To Tolerate Percent" -membertype Noteproperty -value $HAResourceReductionToToleratePercent		
		$Object | add-member -name "HA Isolation Response" -membertype Noteproperty -value $HAIsolationResponse
		$Object | add-member -name "HA Restart Priority" -membertype Noteproperty -value $HARestartPriority	 		
		$Object | add-member -name "HA HB Candidate DS Policy" -membertype Noteproperty -value $HAHBCandidateDatastorePolicy
		$Object | add-member -name "HA APD Response" -membertype Noteproperty -value $HAAPDConfig
		$Object | add-member -name "HA PDL Response" -membertype Noteproperty -value $HAPDLConfig 
		$Object | add-member -name "HA VM Monitoring" -membertype Noteproperty -value $HAVmMonitoring
		$Object | add-member -name "DRS Enabled" -membertype Noteproperty -value $DRSEnabled
		$Object | add-member -name "DRS Mode" -membertype Noteproperty -value $DrsAutomationLevel
		$Object | add-member -name "DRS Level" -membertype Noteproperty -value $DRSLevel		
		$Object | add-member -name "EVC Mode" -membertype Noteproperty -value $EVCMode
		$Object | add-member -name "VM Swapfile Policy" -membertype Noteproperty -value $VMSwapfilePolicy
		
		$tableau += $Object
	}
	return $tableau
}

# Fonction pour récupérer toutes les informations sur les serveurs ESXi
function Get-ESXi-Infos{
	write-host "Collecte des infos ESXi en cours" -NoNewLine
	
	$CLUS = get-cluster

	$tableau = @()
	
	foreach	($clu in $CLUS){
		$ESXS = $clu | get-vmhost
		foreach ($esx in $ESXS){
			write-host "." -NoNewLine
			$esxname = $esx.Name
			$pos = $esxname.IndexOf(".")
			$esxname = $esxname.Substring(0, $pos)
			
			$gw = $esx.ExtensionData.Config.Network.IpRouteConfig.DefaultGateway
			$ntp = Get-VMHostNtpServer -VMHost $esx
			$ntp = ($ntp -join " - ")
			$mgmtVMK = get-vmhost -Name $esx.Name | Get-VMHostNetworkAdapter | Where-Object {$_.ManagementTrafficEnabled -eq $true}
			$mgmtPG = Get-VirtualPortGroup -Name $mgmtVMK.PortGroupName -VMhost $esx.Name
			$mgmtVS = Get-VirtualSwitch -Name $mgmtPG.VirtualSwitch.Name -VMhost $esx.Name
			
			# On teste s'il s'agit d'un VSS ou d'un VDS
			if($mgmtVS.ExtensionData.Capability.DvsOperationSupported -eq $true){
				# Il s'agit d'un VDS
				$vds = Get-VDSwitch $mgmtVS.Name
				$vdpg = Get-VDPortGroup -Name $mgmtPG.Name -VDSwitch $vds
				
				$mgmtvlanid = $vdpg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
				$mgmtActiveNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).ActiveUplinkPort
				$mgmtActiveNIC = ($mgmtActiveNIC -join " - ")
				$mgmtStandbyNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).StandbyUplinkPort
				$mgmtStandbyNIC = ($mgmtStandbyNIC -join " - ")
				$mgmtUnusedNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).UnusedUplinkPort
				$mgmtUnusedNIC = ($mgmtUnusedNIC -join " - ")
				
			}else{
				# Il s'agit d'un VSS
				$mgmtvlanid = $mgmtPG.VLanId
				$mgmtActiveNIC = ($mgmtPG | Get-NicTeamingPolicy).ActiveNic
				$mgmtActiveNIC = ($mgmtActiveNIC -join " - ")
				$mgmtStandbyNIC = ($mgmtPG | Get-NicTeamingPolicy).StandbyNic
				$mgmtStandbyNIC = ($mgmtStandbyNIC -join " - ")
				$mgmtUnusedNIC = ($mgmtPG | Get-NicTeamingPolicy).UnusedNic
				$mgmtUnusedNIC = ($mgmtUnusedNIC -join " - ")
			}
			
			$vmotionVMK = get-vmhost -Name $esx.Name | Get-VMHostNetworkAdapter | Where-Object {$_.VMotionEnabled -eq $true}
			if($vmotionVMK -ne $null){
				$vmotionPG = Get-VirtualPortGroup -Name $vmotionVMK.PortGroupName -VMhost $esx.Name
				$vmotionVS = Get-VirtualSwitch -Name $vmotionPG.VirtualSwitch.Name -VMhost $esx.Name
				
				# On teste s'il s'agit d'un VSS ou d'un VDS
				if($vmotionVS.ExtensionData.Capability.DvsOperationSupported -eq $true){
					# Il s'agit d'un VDS
					$vds = Get-VDSwitch $vmotionVS.Name
					$vdpg = Get-VDPortGroup -Name $vmotionPG.Name -VDSwitch $vds
					
					$vmotionvlanid = $vmotionPG.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
					$vmotionActiveNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).ActiveUplinkPort
					$vmotionActiveNIC = ($vmotionActiveNIC -join " - ")
					$vmotionStandbyNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).StandbyUplinkPort
					$vmotionStandbyNIC = ($vmotionStandbyNIC -join " - ")
					$vmotionUnusedNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $vdpg).UnusedUplinkPort
					$vmotionUnusedNIC = ($vmotionUnusedNIC -join " - ")
				
				}else{
					# Il s'agit d'un VSS
					$vmotionActiveNIC = ($vmotionPG | Get-NicTeamingPolicy).ActiveNic
					$vmotionActiveNIC = ($vmotionActiveNIC -join " - ")
					$vmotionStandbyNIC = ($vmotionPG | Get-NicTeamingPolicy).StandbyNic
					$vmotionStandbyNIC = ($vmotionStandbyNIC -join " - ")
					$vmotionUnusedNIC = ($vmotionPG | Get-NicTeamingPolicy).UnusedNic
					$vmotionUnusedNIC = ($vmotionUnusedNIC -join " - ")
				}
				
			}
			
			$dns = ((Get-EsxCli -VMHost $esx.Name).network.ip.dns.server.list.Invoke()).DNSServers
			$dns = ($dns -join " - ")
			
			$acpi = (get-view ($esx | get-view).configManager.powersystem).Info.CurrentPolicy.ShortName
			if($acpi -eq "static"){
				$acpi = "High Performance"
			}elseif($acpi -eq "dynamic"){
				$acpi = "Balanced"
			}elseif($acpi -eq "low"){
				$acpi = "Low Power"
			}
			$syslog = Get-AdvancedSetting -Entity $esx -Name Syslog.global.logDir
			$scratch = Get-AdvancedSetting -Entity $esx -Name ScratchConfig.ConfiguredScratchLocation
			$VAAIATS = ($esx | Get-AdvancedSetting -Name VMFS3.UseATSForHBOnVMFS5).value
			$VAAIXCOPY = ($esx | Get-AdvancedSetting -Name DataMover.HardwareAcceleratedMove).value
			$VAAILOCKING = ($esx | Get-AdvancedSetting -Name VMFS3.HardwareAcceleratedLocking).value
			$VAAIINIT = ($esx | Get-AdvancedSetting -Name DataMover.HardwareAcceleratedInit).value
			if($VAAILOCKING -eq 0){
				$VAAILOCKING = "Disabled"
			}else{
				$VAAILOCKING = "Enabled"
			}	
			if($VAAIXCOPY -eq 0){
				$VAAIXCOPY = "Disabled"
			}else{
				$VAAIXCOPY = "Enabled"
			}
			if($VAAIATS -eq 0){
				$VAAIATS = "Disabled"
			}else{
				$VAAIATS = "Enabled"
			}
			if($VAAIINIT -eq 0){
				$VAAIINIT = "Disabled"
			}else{
				$VAAIINIT = "Enabled"
			}
			
			$vms = $esx | get-vm | where { $_.PowerState -like "PoweredOn"}
			$totalvCPU = 0
			foreach($vm in $vms){
				$totalvCPU += $vm.NumCpu
			}
			
			# on calcule le ratio vCPU par pCPU actuel
			$actualRatio = 0
			if($totalvCPU -ne 0){
				$actualRatio = $totalvCPU / $esx.ExtensionData.Hardware.CpuInfo.NumCpuCores
				$actualRatio = [MATH]::Round($actualRatio,2)
				$actualRatio = $actualRatio -replace ',','.'
				
			}
			
			$Object = new-object PSObject
			$Object | add-member -name "ESXi" -membertype Noteproperty -value $esxname
			$Object | add-member -name "Cluster" -membertype Noteproperty -value $clu.Name
			$Object | add-member -name "DC" -membertype Noteproperty -value ($clu.ParentFolder).Parent
			$Object | add-member -name "Model" -membertype Noteproperty -value $esx.Model
			$Object | add-member -name "Socket" -membertype Noteproperty -value $esx.ExtensionData.Hardware.CpuInfo.NumCpuPackages
			$Object | add-member -name "Cores per Socket" -membertype Noteproperty -value ($esx.ExtensionData.Hardware.CpuInfo.NumCpuCores / $esx.ExtensionData.Hardware.CpuInfo.NumCpuPackages)
			$Object | add-member -name "Total Cores" -membertype Noteproperty -value $esx.ExtensionData.Hardware.CpuInfo.NumCpuCores
			$Object | add-member -name "Total PoweredOn VMs" -membertype Noteproperty -value $vms.Count
			$Object | add-member -name "Total vCPUs" -membertype Noteproperty -value $totalvCPU
			$Object | add-member -name "Ratio vCPUs per Core" -membertype Noteproperty -value $actualRatio
			$Object | add-member -name "RAM" -membertype Noteproperty -value ([MATH]::Round($esx.MemoryTotalGB))
			$Object | add-member -name "Version" -membertype Noteproperty -value $esx.Version
			$Object | add-member -name "Build" -membertype Noteproperty -value $esx.Build
			
			$Object | add-member -name "NTP" -membertype Noteproperty -value $ntp
			$Object | add-member -name "DNS" -membertype Noteproperty -value $dns
			
			$Object | add-member -name "Mgmt IP" -membertype Noteproperty -value $mgmtVMK.IP
			$Object | add-member -name "Mgmt Mask" -membertype Noteproperty -value $mgmtVMK.SubnetMask
			$Object | add-member -name "Mgmt Gateway" -membertype Noteproperty -value $gw
			$Object | add-member -name "Mgmt vSwitch" -membertype Noteproperty -value $mgmtVS.Name
			$Object | add-member -name "Mgmt Portgroup" -membertype Noteproperty -value $mgmtPG.Name
			$Object | add-member -name "Mgmt VlanId" -membertype Noteproperty -value $mgmtvlanid
			$Object | add-member -name "Mgmt ActiveNIC" -membertype Noteproperty -value $mgmtActiveNIC
			$Object | add-member -name "Mgmt StandbyNIC" -membertype Noteproperty -value $mgmtStandbyNIC
			$Object | add-member -name "Mgmt UnusedNIC" -membertype Noteproperty -value $mgmtUnusedNIC
			
			$Object | add-member -name "vMotion IP" -membertype Noteproperty -value $vmotionVMK.IP
			$Object | add-member -name "vMotion Mask" -membertype Noteproperty -value $vmotionVMK.SubnetMask
			$Object | add-member -name "vMotion vSwitch" -membertype Noteproperty -value $vmotionVS.Name
			$Object | add-member -name "vMotion Portgroup" -membertype Noteproperty -value $vmotionPG.Name
			$Object | add-member -name "vMotion VlanId" -membertype Noteproperty -value $vmotionvlanid
			$Object | add-member -name "vMotion ActiveNIC" -membertype Noteproperty -value $vmotionActiveNIC
			$Object | add-member -name "vMotion StandbyNIC" -membertype Noteproperty -value $vmotionStandbyNIC
			$Object | add-member -name "vMotion UnusedNIC" -membertype Noteproperty -value $vmotionUnusedNIC
			
			$Object | add-member -name "CPU Power Mgmt" -membertype Noteproperty -value $acpi
			$Object | add-member -name "Syslog Location" -membertype Noteproperty -value $syslog
			$Object | add-member -name "ScratchConfig Location" -membertype Noteproperty -value $scratch
			$Object | add-member -name "VAAI XCOPY" -membertype Noteproperty -value $VAAIXCOPY
			$Object | add-member -name "VAAI ATS" -membertype Noteproperty -value $VAAIATS
			$Object | add-member -name "VAAI LOCKING" -membertype Noteproperty -value $VAAILOCKING
			$Object | add-member -name "VAAI INIT" -membertype Noteproperty -value $VAAIINIT
			
			$Object | add-member -name "BiosVersion" -membertype Noteproperty -value $esx.ExtensionData.Hardware.BiosInfo.BiosVersion
			$Object | add-member -name "BiosReleaseDate" -membertype Noteproperty -value $esx.ExtensionData.Hardware.BiosInfo.ReleaseDate
			
			$tableau += $Object
		}
	}
	return $tableau
}

# Fonction pour récupérer toutes les informations IP des vmkernel de tous les ESXi
function Get-VMKernel-Infos {
	write-host "Collecte des infos VMKernel en cours" -NoNewLine
	$vmkernels = Get-VMHostNetworkAdapter | where {$_.Name -like"vmk*"} #| Select VMHost,Name,IP,SubnetMask,Mac,PortGroupName
	
	$tableau = @()
	
	foreach($vmk in $vmkernels){
		write-host "." -NoNewLine
		$Object = new-object PSObject
		$Object | add-member -name "ESXi" -membertype Noteproperty -value $vmk.VMHost
		$Object | add-member -name "Name" -membertype Noteproperty -value $vmk.Name
		$Object | add-member -name "IP" -membertype Noteproperty -value $vmk.IP
		$Object | add-member -name "Mask" -membertype Noteproperty -value $vmk.SubnetMask
		$Object | add-member -name "MAC" -membertype Noteproperty -value $vmk.Mac
		$Object | add-member -name "PortGroupName" -membertype Noteproperty -value $vmk.PortGroupName
		$tableau += $Object
	}
		
	return $tableau
}

# Fonction pour récupérer toutes les informations sur les clusters de datastores
function Get-Datastores-Clusters-Infos{
	$CLUsDS = Get-DatastoreCluster
	$tableau = @()
	write-host "Collecte des infos Datastores Clusters en cours" -NoNewLine
	foreach	($clu in $CLUsDS){
		write-host "." -NoNewLine
		$Object = new-object PSObject
		$Object | add-member -name "Datastore Cluster" -membertype Noteproperty -value $clu.Name
		$Object | add-member -name "Storage DRS Mode" -membertype Noteproperty -value $clu.SdrsAutomationLevel
		$Object | add-member -name "IOLatencyThresholdMillisecond" -membertype Noteproperty -value $clu.IOLatencyThresholdMillisecond
		$Object | add-member -name "IOLoadBalanceEnabled" -membertype Noteproperty -value $clu.IOLoadBalanceEnabled
		$Object | add-member -name "SpaceUtilizationThresholdPercent" -membertype Noteproperty -value $clu.SpaceUtilizationThresholdPercent
		
		$tableau += $Object
	}
	return $tableau
}

# Fonction pour récupérer toutes les informations sur les datastores
function Get-Datastores-Infos{
	$DSs = get-datastore

	$tableau = @()
	write-host "Collecte des infos Datastores en cours" -NoNewLine
	foreach	($ds in $DSs){
		write-host "." -NoNewLine
		$dsname = $ds.Name
		$cluname = $ds | Get-DatastoreCluster
		$capacity = [int]$ds.CapacityGB
		$type = $ds.Type
		$version = $ds.FileSystemVersion
		$blocksize = $ds.ExtensionData.Info.vmfs.BlockSizeMb
		$sioc = $ds.StorageIOControlEnabled
		if($sioc -eq "true"){
			$sioc = "Enabled"
		}else{
			$sioc = "Disabled"
		}
		$siocThresholdMode = $ds.extensiondata.IormConfiguration.CongestionThresholdMode
		$siocThreshold = $ds.CongestionThresholdMillisecond
		$siocThresholdPercentPeak = $ds.extensiondata.IormConfiguration.PercentOfPeakThroughput
		$siocStatsCollectionEnabled = $ds.extensiondata.IormConfiguration.StatsCollectionEnabled
		if($siocStatsCollectionEnabled -eq "true"){
			$siocStatsCollectionEnabled = "Enabled"
		}else{
			$siocStatsCollectionEnabled = "Disabled"
		}
		$Object = new-object PSObject
		$Object | add-member -name "Datastore Cluster" -membertype Noteproperty -value $cluname
		$Object | add-member -name "Datastore Name" -membertype Noteproperty -value $dsname
		$Object | add-member -name "Type" -membertype Noteproperty -value $type
		$Object | add-member -name "Version" -membertype Noteproperty -value $version
		$Object | add-member -name "Block Size (MB)" -membertype Noteproperty -value $blocksize
		$Object | add-member -name "Capacity (GB)" -membertype Noteproperty -value $capacity
		$Object | add-member -name "SIOC Enabled" -membertype Noteproperty -value $sioc
		$Object | add-member -name "SIOC Threshold Mode" -membertype Noteproperty -value $siocThresholdMode
		$Object | add-member -name "SIOC Threshold (ms)" -membertype Noteproperty -value $siocThreshold
		$Object | add-member -name "SIOC Threshold Peak (%)" -membertype Noteproperty -value $siocThresholdPercentPeak
		$Object | add-member -name "SIOC Stat Collection Enabled" -membertype Noteproperty -value $siocStatsCollectionEnabled
				
		$tableau += $Object
	}
	return $tableau
}

# Fonction pour récupérer toutes les infos sur les vSwitches
function Get-vSwitch-Infos {
	write-host "Collecte des infos vSwitches en cours" -NoNewLine
		
	# On initialise le tableau pour stocker les résultats
	$tableau = @()
		
	# On récupère la liste des vSwitches
	$vSwitches = get-datacenter | get-virtualswitch
	
	foreach($vss in $vSwitches){
		write-host "." -NoNewLine
		# On teste s'il s'agit d'un VSS ou d'un VDS
		if($vss.ExtensionData.Capability.DvsOperationSupported -eq $true){
			# Pour chaque VDS, on récupère la liste des Portgroups
			$vds = Get-VDSwitch $vss.Name
			$PGS = Get-VDPortGroup -VDSwitch $vds
			
			foreach($pg in $PGS){
				# on enregistre les résultats dans un objet
				$Object = new-object PSObject
				$Object | add-member -name "Host" -membertype Noteproperty -value "N/A"
				$Object | add-member -name "vSwitch Type" -membertype Noteproperty -value "VDS"
				$Object | add-member -name "vSwitch Name" -membertype Noteproperty -value $vds.Name
				$Object | add-member -name "PG Name" -membertype Noteproperty -value $pg.Name
				
				#On teste s'il s'agit d'un VLAN unique ou d'un range de VLANs
				if($pg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId.Start -ne $null){
					$vlan = ($pg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId.Start).ToString()
					$vlan += "-"
					$vlan +=($pg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId.End).ToString()
					
					$Object | add-member -name "VLAN" -membertype Noteproperty -value $vlan
				
				}else{
					$Object | add-member -name "VLAN" -membertype Noteproperty -value $pg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
				}
				$Object | add-member -name "MTU" -membertype Noteproperty -value $vss.Mtu
				$Object | add-member -name "Policy" -membertype Noteproperty -value $pg.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value
									
				$PGActiveNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $pg).ActiveUplinkPort
				$PGActiveNIC = ($PGActiveNIC -join " - ")
				$PGStandbyNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $pg).StandbyUplinkPort
				$PGStandbyNIC = ($PGStandbyNIC -join " - ")
				$PGUnusedNIC = (Get-VDUplinkTeamingPolicy -VDPortgroup $pg).UnusedUplinkPort
				$PGUnusedNIC = ($PGUnusedNIC -join " - ")
						
				$Object | add-member -name "Active NIC" -membertype Noteproperty -value $PGActiveNIC
				$Object | add-member -name "Standby NIC" -membertype Noteproperty -value $PGStandbyNIC
				$Object | add-member -name "Unused NIC" -membertype Noteproperty -value $PGUnusedNIC
				
				# on ajoute l'objet au tableau des résultat
				$tableau += $Object
			}
		}else{
			# Pour chaque VSS, on récupère la liste des Portgroups
			$PGS = Get-VirtualPortGroup -VirtualSwitch $vss
			
			foreach($pg in $PGS){
				# on enregistre les résultats dans un objet
				$Object = new-object PSObject
				$Object | add-member -name "Host" -membertype Noteproperty -value $vss.VMHost.Name
				$Object | add-member -name "vSwitch Type" -membertype Noteproperty -value "VSS"
				$Object | add-member -name "vSwitch Name" -membertype Noteproperty -value $vss.Name
				$Object | add-member -name "PG Name" -membertype Noteproperty -value $pg.Name
				$Object | add-member -name "VLAN" -membertype Noteproperty -value $pg.VLanId
				$Object | add-member -name "MTU" -membertype Noteproperty -value $vss.Mtu
				$Object | add-member -name "Policy" -membertype Noteproperty -value $pg.ExtensionData.ComputedPolicy.NicTeaming.Policy
				
				$PGActiveNIC = ($pg | Get-NicTeamingPolicy).ActiveNic
				$PGActiveNIC = ($PGActiveNIC -join " - ")
				$PGStandbyNIC = ($pg | Get-NicTeamingPolicy).StandbyNic
				$PGStandbyNIC = ($PGStandbyNIC -join " - ")
				$PGUnusedNIC = ($pg | Get-NicTeamingPolicy).UnusedNic
				$PGUnusedNIC = ($PGUnusedNIC -join " - ")
				
				$Object | add-member -name "Active NIC" -membertype Noteproperty -value $PGActiveNIC
				$Object | add-member -name "Standby NIC" -membertype Noteproperty -value $PGStandbyNIC
				$Object | add-member -name "Unused NIC" -membertype Noteproperty -value $PGUnusedNIC
							
				# on ajoute l'objet au tableau des résultat
				$tableau += $Object
			}
		}
			

	}
	
	return $tableau
	
}

# Fonction pour récupérer toutes les informations sur les permissions
function Get-Permissions-Infos {
	write-host "Collecte des infos sur les Permissions en cours" -NoNewLine
	
	$tableau = @()
	$authMgr = Get-View AuthorizationManager
	$roleHash = @{}
	$authMgr.RoleList | %{
		$roleHash[$_.RoleId] = $_.Name
	}
  
    $perms = $authMgr.RetrieveAllPermissions()
    foreach($perm in $perms){
		if($perm.group -eq $true){
			$obj = "Groupe"
		}else{
			$obj = "Utilisateur"
		}
		$Object = New-Object PSObject
		$entity = Get-View $perm.Entity
		$Object | Add-Member -Type noteproperty -Name "Objet" -Value $entity.Name
		$Object | Add-Member -Type noteproperty -Name "Type" -Value $entity.gettype().Name
		$Object | Add-Member -Type noteproperty -Name "Utilisateur ou groupe" -Value $obj
		$Object | Add-Member -Type noteproperty -Name "Entité" -Value $perm.Principal
		$Object | Add-Member -Type noteproperty -Name "Propagation" -Value $perm.Propagate
		$Object | Add-Member -Type noteproperty -Name "Role" -Value $roleHash[$perm.RoleId]
		$tableau += $Object
    }
  
	return $tableau
}

# Fonction pour récupérer toutes les informations sur les roles
function Get-Roles-Infos {
	write-host "Collecte des infos sur les roles en cours" -NoNewLine
	
	$authMgr = Get-View AuthorizationManager
    $tableau = @()
  
    foreach($role in $authMgr.roleList){
		$listPrivileges = $role.privilege -join " ; "
		
		$Object = New-Object PSObject
		$Object | Add-Member -Type noteproperty -Name "Nom" -Value $role.name
		$Object | Add-Member -Type noteproperty -Name "Label" -Value $role.info.label
		$Object | Add-Member -Type noteproperty -Name "Role Systeme" -Value $role.system
		$Object | Add-Member -Type noteproperty -Name "Privileges" -Value $listPrivileges
		$tableau += $Object
    }
  
    return $tableau
}

# Fonction pour convertir les fichiers CSV temporaire en un fichier final XLSX avec plusieurs onglets
function Convert-CSV {
	Write-Host "Conversion des fichiers CSV en un seul fichier XLSX" -foregroundcolor "green"
	$p = $path + "\*"
	$csvs = Get-ChildItem $p -Include *.csv
	$y=$csvs.Count
	if($y -eq '0'){
		write-host "Aucun fichier a convertir"
	}else{
		foreach ($csv in $csvs)
		{
			Write-Host $csv.Name
		}
		
		$excelapp = new-object -comobject Excel.Application
		$excelapp.sheetsInNewWorkbook = $csvs.Count
		$xlsx = $excelapp.Workbooks.Add()
		$sheet=1

		foreach ($csv in $csvs)
		{
			$row=1
			$column=1
			$worksheet = $xlsx.Worksheets.Item($sheet)
			$worksheet.Name = $csv.Name -replace '_Infos.csv',''
			$file = Get-Content $csv
			foreach($line in $file)
			{
				$linecontents=$line -split ',(?!\s*\w+”)'
				foreach($cell in $linecontents)
				{
					$cell = $cell -replace '"',''
					$worksheet.Cells.Item($row,$column) = $cell
					$column++
				}
				$column=1
				$row++
			}
			$sheet++
		}
		
		$output = "$path\$XLSXfilename"
		Write-Host "Enregistrement du fichier XLSX => $XLSXfilename"
		$xlsx.SaveAs($output)
		$excelapp.quit()
	}
}




#################### DEBUT DU SCRIPT ########################################################

############## CONNEXION AU VCENTER
write-host "Connexion au vCenter: $VCENTER" -foregroundcolor "green"
Connect-VIServer -server $VCENTER
write-host

############## COLLECTE DES INFOS
write-host "Collectes des infos:" -foregroundcolor "green"

# On lance la collecte d'informations du vCenter et on récupère le résultat
$resultVCENTER = Get-VCenter-Infos
write-host

# On lance la collecte d'informations sur les clusters ESXi et on récupère le résultat
$resultClusters = Get-Clusters-Infos
write-host

# On lance la collecte d'informations sur les ESXi et on récupère le résultat
$resultESXi = Get-ESXi-Infos
write-host

# On lance la collecte d'informations sur les VMkernels et on récupère le résultat
$resultVMK = Get-VMKernel-Infos
write-host

# On lance la collecte d'informations sur les clusters de datastores et on récupère le résultat
$resultCLUSDatastores = Get-Datastores-Clusters-Infos
write-host

# On lance la collecte d'informations sur les datastores et on récupère le résultat
$resultDatastores = Get-Datastores-Infos
write-host

# On lance la collecte d'informations sur les vSwitches et on récupère le résultat
$resultvSwitches = Get-vSwitch-Infos
write-host

# On lance la collecte d'informations sur les permissions et on récupère le résultat
$resultPermissions = Get-Permissions-Infos
write-host

# On lance la collecte d'informations sur les roles et on récupère le résultat
$resultRoles = Get-Roles-Infos
write-host

############## AFFICHAGE DES RESULTATS
write-host
write-host "Bilan des objets analyses:" -foregroundcolor "green"
write-host $resultClusters.Count "Clusters ESXi trouves"
write-host $resultESXi.Count "ESXi trouves"
write-host $resultCLUSDatastores.Count "Clusters de Datastores trouves"
write-host $resultDatastores.Count "Datastores trouves"
write-host $resultvSwitches.Count "vSwitches trouves"
write-host $resultPermissions.Count "Permissions trouvees"
write-host $resultRoles.Count "Roles trouves"
write-host


############## EXPORTS DES FICHIERS CSV TEMPORAIRES
write-host "Export des fichiers CSV:" -foregroundcolor "green"

# EXPORT VCENTER
write-host "Export du fichier: $VCENTERfilename"
$out = $path + "\" + $VCENTERfilename
$resultVCENTER | Export-csv -Path $out -NoTypeInformation

# EXPORT CLUSTERS
write-host "Export du fichier: $CLUSfilename"
$out = $path + "\" + $CLUSfilename
$resultClusters | Export-csv -Path $out -NoTypeInformation

# EXPORT ESXI
write-host "Export du fichier: $ESXifilename"
$out = $path + "\" + $ESXifilename
$resultESXi | Export-csv -Path $out -NoTypeInformation

# EXPORT VMKERNEL
write-host "Export du fichier: $VMKfilename"
$out = $path + "\" + $VMKfilename
$resultVMK | Export-csv -Path $out -NoTypeInformation

# EXPORT CLUSTERS DE DATASTORES
write-host "Export du fichier: $DSCLUSfilename"
$out = $path + "\" + $DSCLUSfilename
$resultCLUSDatastores | Export-csv -Path $out -NoTypeInformation

# EXPORT DATASTORES
write-host "Export du fichier: $DSfilename"
$out = $path + "\" + $DSfilename
$resultDatastores | Export-csv -Path $out -NoTypeInformation

# EXPORT VSWITCHES
write-host "Export du fichier: $SWfilename"
$out = $path + "\" + $SWfilename
$resultvSwitches | Export-csv -Path $out -NoTypeInformation

# EXPORT PERMISSIONS
write-host "Export du fichier: $PERMfilename"
$out = $path + "\" + $PERMfilename
$resultPermissions | Export-csv -Path $out -NoTypeInformation

# EXPORT ROLES
write-host "Export du fichier: $ROLESfilename"
$out = $path + "\" + $ROLESfilename
$resultRoles | Export-csv -Path $out -NoTypeInformation
write-host

############## CONVERSION DES FICHIERS CSV TEMPORAIRES EN XLSX
Convert-CSV
$p = $path + "\*" 
Remove-Item -path $p -include *.csv
write-host

############## DECONNEXION DU VCENTER
write-host "Deconnexion du vCenter: $VCENTER" -foregroundcolor "green"
Disconnect-VIServer * -confirm:$false


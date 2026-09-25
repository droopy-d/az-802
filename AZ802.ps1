#Requires -Version 7.2
<################################################################################
.SYNOPSIS
Deploys Microsoft AZ-802 Azure lab environments for 1-8 students.

.DESCRIPTION
Wraps the current MicrosoftLearning AZ-802 Azure deployment scripts. Each
student receives a separate resource group and VNet, which safely permits the
official VM names, IP addresses, and contoso.com domain to be reused. The
official deployment prepares one selected lab (1-8) for each student.

This preserves the Microsoft lab's shared CONTOSO\Administrator credential.
Resource-group/VNet separation blocks private east-west traffic, but every
source address allowed by the RDP rule can reach every student's public RDP
endpoint. This is not per-student identity or strong classroom access isolation.

Labs 1-3 and 5-8 use the official direct-VM topology during provisioning, then
replace each VM public IP with one Standard Load Balancer public IP per student.
Each VM receives its own RDP frontend port, and an outbound rule preserves guest
internet access. Lab 4's nested host is published the same way. These deployments
create real billable Azure resources; delete the generated resource groups after
the training.

Use -OsDiskSku to select Premium_LRS, StandardSSD_LRS, or Standard_LRS for VM OS
disks. The default remains Premium_LRS to match Microsoft's templates. The
deployment runs against a temporary copy of the lab assets, so the upstream
repository files are not modified.

By default, a one-time Azure Automation runbook deletes all lab resource groups
created by this invocation at 17:00 using W. Europe Standard Time (CET/CEST
daylight-saving aware). If launched after 17:00, cleanup is scheduled for 17:00
the next day. It creates/reuses an Automation Account outside the lab groups and
grants its managed identity Contributor on only the target groups. This requires
Az.Automation and permission to create role assignments (for example Owner or
User Access Administrator). The separate Automation resource group/account is
retained. Disable with -DisableAutoDelete, or customize using -AutoDeleteTime,
-AutoDeleteTimeZoneId, and Automation* parameters. If scheduling fails, manual
cleanup commands are printed.

The Microsoft source repository can be supplied with -RepositoryPath. If it is
omitted, the script downloads the current main branch archive to a temporary
folder and removes it when finished. For reproducible deliveries, use a local
clone/release that you have reviewed.

BUDGET ESTIMATE (planning only; USD retail prices converted at 1 USD = EUR 0.87696)
Estimate dated 2026-09-26 for Denmark East or Austria East, 8 students, all 8
labs, up to 2 hours of student lab use per lab, and a 1-hour powered-on buffer
before each lab. Assumes the current serial deployment flow (including typical
setup/wait time), deallocation after deployment and overnight, D4s_v5 VMs for
Labs 1-3 and 5-8, a D8s_v5 nested host for Lab 4, one Standard Load Balancer/
public IP per student environment, and deletion of each resource group after use.
The 1-hour pre-lab startup buffer adds approximately EUR 9 per student (EUR 72
for 8 students) to the prior 2-hour-use estimate.

	OS disk SKU          Approx. per student, all 8 labs    Approx. class total (8)
	Premium_LRS          EUR 51.50-63.40                    EUR 412-507
	StandardSSD_LRS      EUR 50-62                          EUR 400-493
	Standard_LRS (HDD)   EUR 49.60-60.95                    EUR 397-488

These are estimates, not Azure quotes. Actual costs vary with region, agreement,
deployment duration, VM availability, and resource cleanup timing. The estimates
exclude Azure File Sync resources students create in Lab 5, variable data transfer,
tax, discounts, and Azure Automation usage. The separate Automation Account
resource group is retained after lab cleanup. Disks, public IPs, and load balancers
can continue billing while their resource groups exist, even after VMs are
deallocated.

.EXAMPLE
./AZ802.ps1 -LabNumber 1 -StudentCount 8 -Location denmarkeast -OsDiskSku StandardSSD_LRS

Uses the default '*' RDP source rule. The Load Balancer public IP and RDP ports
are reachable from any internet address; the script prints a security warning.

.EXAMPLE
./AZ802.ps1 -LabNumber 4 -StudentCount 2 -Location austriaeast -OsDiskSku StandardSSD_LRS

Uses the nested Hyper-V host required for Lab 4 in Austria East.

.EXAMPLE
./AZ802.ps1 -LabNumber 5 -StudentCount 1 -RepositoryPath C:\Labs\AZ-802-Windows-Server-Administrator-Associate

.EXAMPLE
./AZ802.ps1 -LabNumber 1 -StudentCount 1 -DisableAutoDelete

Deploys without creating an automatic cleanup schedule.
################################################################################>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
	[Parameter(Mandatory)]
	[ValidateRange(1, 8)]
	[int]$LabNumber,

	# Choose the class size explicitly. One student uses five VMs in the full
	# baseline for most labs; Lab 4 uses one nested-virtualization host.
	[ValidateRange(1, 8)]
	[int]$StudentCount = 1,

	[string]$SubscriptionId,

	[string]$Location = 'denmarkeast',

	[ValidatePattern('^[A-Za-z0-9._()-]{1,50}$')]
	[string]$ResourceGroupPrefix = 'az802',

	[string]$RepositoryPath,

	[string]$AdminUsername = 'labadmin',

	# This is the disposable credential documented by the official lab.
	# Change it only if all lab instructions/guest setup are updated together.
	[securestring]$AdminPassword = (ConvertTo-SecureString 'PA55w.rd1234' -AsPlainText -Force),

	[string]$VmSize = 'Standard_D4s_v5',

	[string]$Lab04HostVmSize = 'Standard_D8s_v5',

	[ValidateSet('Premium_LRS', 'StandardSSD_LRS', 'Standard_LRS')]
	[string]$OsDiskSku = 'Premium_LRS',

	[ValidateRange(1024, 65530)]
	[int]$RdpFrontendPortStart = 50001,

	# '*' permits RDP from anywhere (convenient for students at different homes).
	# To restrict access instead, pass a trainer/classroom public IPv4 or CIDR.
	# Pass an empty string explicitly to discover the operator's public IPv4 /32.
	[string]$AllowedRdpSourceIP = '*',

	[uri]$Lab04VhdUri = 'https://go.microsoft.com/fwlink/p/?linkid=2195166&clcid=0x409&culture=en-us&country=us',

	[string]$OutputCsvPath = '.\az802-lab-rdp-targets.csv',

	# Existing generated groups are never replaced unless this switch is set.
	# It deletes only the student resource groups generated by this invocation.
	[switch]$CleanStart,

	# By default, schedule deletion of this invocation's lab resource groups.
	[switch]$DisableAutoDelete,

	[ValidatePattern('^(?:[01][0-9]|2[0-3])[0-5][0-9]$')]
	[string]$AutoDeleteTime = '1700',

	[string]$AutoDeleteTimeZoneId = 'W. Europe Standard Time',

	[string]$AutomationResourceGroupName = 'rg-az802-automation',

	[string]$AutomationAccountName = 'az802-lab-cleanup',

	[string]$AutomationLocation = 'northeurope'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$scriptStartedAt = Get-Date
$temporaryRepositoryRoot = $null
$originalProgressPreference = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'

$labVmNames = @{
	1 = @('SEA-DC1', 'SEA-ADM1', 'SEA-SVR1')
	2 = @('SEA-DC1', 'SEA-ADM1')
	3 = @('SEA-DC1', 'SEA-ADM1', 'SEA-SVR1')
	4 = @('AZ802-L4-HOST')
	5 = @('SEA-DC1', 'SEA-ADM1', 'SEA-SVR1', 'SEA-SVR2')
	6 = @('SEA-DC1', 'SEA-ADM1', 'SEA-SVR1', 'SEA-SVR2', 'SEA-SVR3')
	7 = @('SEA-DC1', 'SEA-SVR1', 'SEA-SVR2')
	8 = @('SEA-DC1', 'SEA-SVR2')
}

$primaryVmNames = @{
	1 = 'SEA-ADM1'
	2 = 'SEA-ADM1'
	3 = 'SEA-ADM1'
	4 = 'AZ802-L4-HOST'
	5 = 'SEA-ADM1'
	6 = 'SEA-ADM1'
	7 = 'SEA-SVR2'
	8 = 'SEA-SVR2'
}

function Resolve-Az802DeploymentFolder {
	if ($RepositoryPath) {
		$resolvedRepositoryPath = (Resolve-Path -LiteralPath $RepositoryPath -ErrorAction Stop).Path
		$allFilesSource = Join-Path $resolvedRepositoryPath 'Allfiles'
		if (-not (Test-Path -LiteralPath (Join-Path $allFilesSource 'AZ802-Lab00\VM-Specs\Deploy') -PathType Container)) {
			if ((Split-Path -Leaf $resolvedRepositoryPath) -eq 'Allfiles') {
				$allFilesSource = $resolvedRepositoryPath
			}
			elseif ((Split-Path -Leaf $resolvedRepositoryPath) -eq 'Deploy') {
				$allFilesSource = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $resolvedRepositoryPath))
			}
			else {
				throw "RepositoryPath must point to the Microsoft repository root, its Allfiles folder, or VM-Specs\Deploy: $resolvedRepositoryPath"
			}
		}

		if (-not (Test-Path -LiteralPath (Join-Path $allFilesSource 'AZ802-Lab00\VM-Specs\Deploy') -PathType Container)) {
			throw "The AZ-802 deployment folder was not found under $allFilesSource."
		}

		$workingRoot = Join-Path ([IO.Path]::GetTempPath()) "az802-working-$([guid]::NewGuid().ToString('N'))"
		$script:temporaryRepositoryRoot = $workingRoot
		New-Item -Path $workingRoot -ItemType Directory -Force | Out-Null
		Copy-Item -LiteralPath $allFilesSource -Destination (Join-Path $workingRoot 'Allfiles') -Recurse -Force
		return (Join-Path $workingRoot 'Allfiles\AZ802-Lab00\VM-Specs\Deploy')
	}

	$temporaryRepositoryRoot = Join-Path ([IO.Path]::GetTempPath()) "az802-source-$([guid]::NewGuid().ToString('N'))"
	$script:temporaryRepositoryRoot = $temporaryRepositoryRoot
	$archivePath = Join-Path $temporaryRepositoryRoot 'az802-main.zip'
	$extractPath = Join-Path $temporaryRepositoryRoot 'source'
	New-Item -Path $extractPath -ItemType Directory -Force | Out-Null

	$archiveUri = 'https://github.com/MicrosoftLearning/AZ-802-Windows-Server-Administrator-Associate/archive/refs/heads/main.zip'
	Write-Host 'Downloading the current MicrosoftLearning AZ-802 lab files...'
	Invoke-WebRequest -Uri $archiveUri -OutFile $archivePath -ErrorAction Stop
	Expand-Archive -LiteralPath $archivePath -DestinationPath $extractPath -Force

	$repositoryRoot = Get-ChildItem -LiteralPath $extractPath -Directory | Select-Object -First 1
	if (-not $repositoryRoot) {
		throw 'The AZ-802 repository archive did not contain the expected root folder.'
	}

	$deploymentFolder = Join-Path $repositoryRoot.FullName 'Allfiles\AZ802-Lab00\VM-Specs\Deploy'
	return $deploymentFolder
}

function Set-Az802TemplateOsDiskSku {
	param([Parameter(Mandatory)][string]$TemplatePath)

	$template = Get-Content -LiteralPath $TemplatePath -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 100
	$changedDiskCount = 0
	foreach ($resource in $template.resources) {
		if ($resource.type -eq 'Microsoft.Compute/virtualMachines' -and
			$resource.properties.storageProfile.osDisk.managedDisk) {
			$resource.properties.storageProfile.osDisk.managedDisk.storageAccountType = $OsDiskSku
			$changedDiskCount++
		}
	}

	# Lab 0 explicitly specifies one 128-GB raw data disk on each of these VMs
	# for Lab 5. The current upstream ARM template has 8 GB in these two arrays.
	$changedLab05DiskCount = 0
	if ($LabNumber -eq 5) {
		foreach ($diskVariableName in @('lab05Svr1DataDisks', 'lab05Svr2DataDisks')) {
			$diskVariable = $template.variables.PSObject.Properties[$diskVariableName]
			if (-not $diskVariable -or @($diskVariable.Value).Count -ne 1) {
				throw "Expected exactly one Lab 5 data disk in template variable '$diskVariableName' in '$TemplatePath'."
			}
			$disk = @($diskVariable.Value)[0]
			if ($disk.createOption -ne 'Empty' -or $disk.managedDisk.storageAccountType -ne 'Standard_LRS') {
				throw "Unexpected Lab 5 data-disk configuration in '$diskVariableName'; refusing to alter an unrecognized template."
			}
			$disk.diskSizeGB = 128
			$changedLab05DiskCount++
		}
	}

	if ($changedDiskCount -eq 0) {
		throw "No VM OS disk definitions were found in template '$TemplatePath'; refusing to deploy with an unverified disk type."
	}
	if ($LabNumber -eq 5 -and $changedLab05DiskCount -ne 2) {
		throw "Expected to correct two Lab 5 data-disk definitions in '$TemplatePath'; corrected $changedLab05DiskCount."
	}

	$template | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $TemplatePath -Encoding utf8
	Write-Host "Configured $changedDiskCount VM OS disk(s) in $(Split-Path -Leaf $TemplatePath) as $OsDiskSku, and aligned Lab 5 data disks to 128 GB each."
}

function Set-Az802ScheduledCleanup {
	param(
		[Parameter(Mandatory)][string[]]$ResourceGroupNames,
		[Parameter(Mandatory)][string]$SubscriptionId
	)

	if (-not (Get-Module -ListAvailable -Name Az.Automation)) {
		throw "Az.Automation is required for automatic cleanup. Install it with Install-Module Az.Automation -Scope CurrentUser, or rerun with -DisableAutoDelete."
	}
	Import-Module Az.Automation -ErrorAction Stop

	$automationProvider = Get-AzResourceProvider -ProviderNamespace Microsoft.Automation -ErrorAction Stop
	if ($automationProvider.RegistrationState -ne 'Registered') {
		Register-AzResourceProvider -ProviderNamespace Microsoft.Automation -ErrorAction Stop | Out-Null
		for ($attempt = 1; $attempt -le 30; $attempt++) {
			$automationProvider = Get-AzResourceProvider -ProviderNamespace Microsoft.Automation -ErrorAction Stop
			if ($automationProvider.RegistrationState -eq 'Registered') { break }
			Start-Sleep -Seconds 10
		}
		if ($automationProvider.RegistrationState -ne 'Registered') {
			throw 'Microsoft.Automation provider registration did not complete. Register it and rerun, or use -DisableAutoDelete.'
		}
	}

	$automationRg = Get-AzResourceGroup -Name $AutomationResourceGroupName -ErrorAction SilentlyContinue
	if (-not $automationRg) {
		New-AzResourceGroup -Name $AutomationResourceGroupName -Location $AutomationLocation -ErrorAction Stop | Out-Null
	}

	$automationAccount = Get-AzAutomationAccount `
		-ResourceGroupName $AutomationResourceGroupName `
		-Name $AutomationAccountName `
		-ErrorAction SilentlyContinue
	if (-not $automationAccount) {
		$automationAccount = New-AzAutomationAccount `
			-ResourceGroupName $AutomationResourceGroupName `
			-Name $AutomationAccountName `
			-Location $AutomationLocation `
			-AssignSystemIdentity `
			-ErrorAction Stop
	}

	for ($attempt = 1; $attempt -le 18; $attempt++) {
		$automationAccount = Get-AzAutomationAccount `
			-ResourceGroupName $AutomationResourceGroupName `
			-Name $AutomationAccountName `
			-ErrorAction SilentlyContinue
		if ($automationAccount -and $automationAccount.Identity.PrincipalId) { break }
		if ($attempt -lt 18) { Start-Sleep -Seconds 10 }
	}
	if (-not $automationAccount) {
		throw "Automation Account '$AutomationAccountName' could not be read after creation."
	}
	$principalId = $automationAccount.Identity.PrincipalId
	if (-not $principalId) {
		throw "Automation Account '$AutomationAccountName' did not receive a system-assigned managed identity."
	}

	foreach ($resourceGroupName in $ResourceGroupNames) {
		$resourceGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction Stop
		$assignment = Get-AzRoleAssignment `
			-ObjectId $principalId `
			-Scope $resourceGroup.ResourceId `
			-RoleDefinitionName Contributor `
			-ErrorAction SilentlyContinue
		if ($assignment) { continue }

		$assigned = $false
		for ($attempt = 1; $attempt -le 12; $attempt++) {
			try {
				New-AzRoleAssignment `
					-ObjectId $principalId `
					-ObjectType ServicePrincipal `
					-RoleDefinitionName Contributor `
					-Scope $resourceGroup.ResourceId `
					-ErrorAction Stop | Out-Null
				$assigned = $true
				break
			}
			catch {
				if ($_.Exception.Message -notmatch 'PrincipalNotFound' -or $attempt -eq 12) { throw }
				Start-Sleep -Seconds 10
			}
		}
		if (-not $assigned) { throw "Could not assign cleanup rights on '$resourceGroupName'." }
	}

	$cleanupToken = [guid]::NewGuid().ToString('N')
	foreach ($resourceGroupName in $ResourceGroupNames) {
		$resourceGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction Stop
		$tags = @{}
		if ($resourceGroup.Tags) {
			foreach ($tagName in $resourceGroup.Tags.Keys) { $tags[$tagName] = $resourceGroup.Tags[$tagName] }
		}
		$tags['AZ802AutoDeleteToken'] = $cleanupToken
		Set-AzResourceGroup -Name $resourceGroupName -Tag $tags -ErrorAction Stop | Out-Null
	}

	$runbookName = 'Remove-AZ802LabResourceGroups'
	$runbookScript = @'
param(
    [Parameter(Mandatory = $true)][string]$SubscriptionId,
    [Parameter(Mandatory = $true)][string]$ResourceGroupNamesJson,
    [Parameter(Mandatory = $true)][string]$ResourceGroupPrefix,
	[Parameter(Mandatory = $true)][int]$LabNumber,
	[Parameter(Mandatory = $true)][string]$CleanupToken
)

$ErrorActionPreference = 'Stop'
Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Disable-AzContextAutosave -Scope Process | Out-Null
Connect-AzAccount -Identity | Out-Null
Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

$expectedPattern = '^' + [regex]::Escape($ResourceGroupPrefix) + '-L' + $LabNumber.ToString('00') + '-S\d{2}-ENV$'
$targetNames = @($ResourceGroupNamesJson | ConvertFrom-Json)
foreach ($name in $targetNames) {
    if ($name -notmatch $expectedPattern) {
        throw "Refusing to delete unexpected resource group name '$name'."
    }
	$resourceGroup = Get-AzResourceGroup -Name $name -ErrorAction SilentlyContinue
	if ($resourceGroup -and $resourceGroup.Tags['AZ802AutoDeleteToken'] -eq $CleanupToken) {
        Write-Output "Deleting AZ-802 lab resource group '$name'."
        Remove-AzResourceGroup -Name $name -Force -ErrorAction Stop | Out-Null
    }
	elseif ($resourceGroup) {
		Write-Warning "Skipping '$name': its cleanup token changed after this schedule was created."
	}
}
'@
	$runbookPath = Join-Path ([IO.Path]::GetTempPath()) "$runbookName-$([guid]::NewGuid().ToString('N')).ps1"
	Set-Content -LiteralPath $runbookPath -Value $runbookScript -Encoding utf8
	try {
		Import-AzAutomationRunbook `
			-ResourceGroupName $AutomationResourceGroupName `
			-AutomationAccountName $AutomationAccountName `
			-Name $runbookName `
			-Path $runbookPath `
			-Type PowerShell `
			-Force `
			-ErrorAction Stop | Out-Null
		Publish-AzAutomationRunbook `
			-ResourceGroupName $AutomationResourceGroupName `
			-AutomationAccountName $AutomationAccountName `
			-Name $runbookName `
			-ErrorAction Stop | Out-Null
	}
	finally {
		Remove-Item -LiteralPath $runbookPath -Force -ErrorAction SilentlyContinue
	}

	$timeZone = [TimeZoneInfo]::FindSystemTimeZoneById($AutoDeleteTimeZoneId)
	$nowLocal = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $timeZone)
	$hour = [int]$AutoDeleteTime.Substring(0, 2)
	$minute = [int]$AutoDeleteTime.Substring(2, 2)
	$deleteLocal = [DateTime]::SpecifyKind($nowLocal.Date.AddHours($hour).AddMinutes($minute), [DateTimeKind]::Unspecified)
	if ($deleteLocal -le $nowLocal.AddMinutes(10)) { $deleteLocal = $deleteLocal.AddDays(1) }
	$deleteUtc = [TimeZoneInfo]::ConvertTimeToUtc($deleteLocal, $timeZone)
	$scheduleName = 'az802-delete-' + $deleteUtc.ToString('yyyyMMddHHmm') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6)

	$automationSchedule = New-AzAutomationSchedule `
		-ResourceGroupName $AutomationResourceGroupName `
		-AutomationAccountName $AutomationAccountName `
		-Name $scheduleName `
		-StartTime ([DateTimeOffset]$deleteUtc) `
		-TimeZone $AutoDeleteTimeZoneId `
		-OneTime `
		-ErrorAction Stop

	$runbookParameters = @{
		SubscriptionId = $SubscriptionId
		ResourceGroupNamesJson = (ConvertTo-Json -InputObject @($ResourceGroupNames) -Compress)
		ResourceGroupPrefix = $ResourceGroupPrefix
		LabNumber = [string]$LabNumber
		CleanupToken = $cleanupToken
	}
	Register-AzAutomationScheduledRunbook `
		-ResourceGroupName $AutomationResourceGroupName `
		-AutomationAccountName $AutomationAccountName `
		-RunbookName $runbookName `
		-ScheduleName $scheduleName `
		-Parameters $runbookParameters `
		-ErrorAction Stop | Out-Null

	return [pscustomobject]@{
		DeleteLocal = $deleteLocal
		TimeZoneId = $AutoDeleteTimeZoneId
		ScheduleName = $scheduleName
		RunbookName = $runbookName
		AutomationResourceGroupName = $AutomationResourceGroupName
		AutomationAccountName = $AutomationAccountName
	}
}

function Wait-Az802ResourceGroupDeleted {
	param([Parameter(Mandatory)][string]$Name)

	for ($attempt = 1; $attempt -le 180; $attempt++) {
		$group = Get-AzResourceGroup -Name $Name -ErrorAction SilentlyContinue
		if (-not $group) { return }
		Write-Host "Waiting for resource group '$Name' deletion to finish ($attempt/180)..."
		Start-Sleep -Seconds 10
	}

	throw "Timed out waiting for resource group '$Name' to be deleted."
}

function Set-Az802StudentLoadBalancer {
	param(
		[Parameter(Mandatory)][string]$Name,
		[Parameter(Mandatory)][int]$StudentNumber
	)

	$studentVmNames = @($labVmNames[$LabNumber])
	if (($RdpFrontendPortStart + $studentVmNames.Count - 1) -gt 65535) {
		throw "RdpFrontendPortStart $RdpFrontendPortStart leaves too few ports for $($studentVmNames.Count) VMs."
	}

	$loadBalancerName = "$Name-lb"
	$publicIpName = "$Name-lb-pip"
	$publicIp = New-AzPublicIpAddress `
		-Name $publicIpName `
		-ResourceGroupName $Name `
		-Location $Location `
		-Sku Standard `
		-AllocationMethod Static `
		-ErrorAction Stop

	$frontend = New-AzLoadBalancerFrontendIpConfig `
		-Name 'az802-frontend' `
		-PublicIpAddress $publicIp
	$backendPool = New-AzLoadBalancerBackendAddressPoolConfig -Name 'az802-backend'
	$natRules = [Collections.Generic.List[object]]::new()
	for ($index = 0; $index -lt $studentVmNames.Count; $index++) {
		$vmName = $studentVmNames[$index]
		$natRule = New-AzLoadBalancerInboundNatRuleConfig `
			-Name "rdp-$vmName" `
			-FrontendIpConfiguration $frontend `
			-Protocol Tcp `
			-FrontendPort ($RdpFrontendPortStart + $index) `
			-BackendPort 3389 `
			-IdleTimeoutInMinutes 15 `
			-EnableTcpReset
		$natRules.Add($natRule)
	}

	# Keep guest internet access after removing the VM-level public IPs.
	$outboundRule = New-AzLoadBalancerOutboundRuleConfig `
		-Name 'az802-outbound' `
		-Protocol All `
		-FrontendIpConfiguration $frontend `
		-BackendAddressPool $backendPool `
		-AllocatedOutboundPort 1024 `
		-IdleTimeoutInMinutes 15 `
		-EnableTcpReset

	$null = New-AzLoadBalancer `
		-Name $loadBalancerName `
		-ResourceGroupName $Name `
		-Location $Location `
		-Sku Standard `
		-FrontendIpConfiguration $frontend `
		-BackendAddressPool $backendPool `
		-InboundNatRule $natRules.ToArray() `
		-OutboundRule $outboundRule `
		-ErrorAction Stop
	$loadBalancer = Get-AzLoadBalancer -Name $loadBalancerName -ResourceGroupName $Name -ErrorAction Stop
	$publicIp = Get-AzPublicIpAddress -Name $publicIpName -ResourceGroupName $Name -ErrorAction Stop

	for ($index = 0; $index -lt $studentVmNames.Count; $index++) {
		$vmName = $studentVmNames[$index]
		$vm = Get-AzVM -ResourceGroupName $Name -Name $vmName -ErrorAction Stop
		$nicId = $vm.NetworkProfile.NetworkInterfaces[0].Id
		$nic = Get-AzNetworkInterface -ResourceId $nicId -ErrorAction Stop
		$ipConfiguration = $nic.IpConfigurations | Select-Object -First 1
		$oldPublicIpId = if ($ipConfiguration.PublicIpAddress) { $ipConfiguration.PublicIpAddress.Id } else { $null }
		$natRule = $loadBalancer.InboundNatRules | Where-Object Name -eq "rdp-$vmName" | Select-Object -First 1
		if (-not $natRule) {
			throw "Load Balancer NAT rule for $vmName was not found."
		}

		$ipConfiguration.PublicIpAddress = $null
		$ipConfiguration.LoadBalancerBackendAddressPools = @($loadBalancer.BackendAddressPools[0])
		$ipConfiguration.LoadBalancerInboundNatRules = @($natRule)
		Set-AzNetworkInterface -NetworkInterface $nic -ErrorAction Stop | Out-Null

		if ($oldPublicIpId) {
			$oldPublicIp = Get-AzPublicIpAddress -ResourceId $oldPublicIpId -ErrorAction SilentlyContinue
			if ($oldPublicIp) {
				Remove-AzPublicIpAddress -Name $oldPublicIp.Name -ResourceGroupName $Name -Force -ErrorAction Stop
			}
		}
	}

	Write-Host "Student $StudentNumber RDP is mapped through $($publicIp.IpAddress):$RdpFrontendPortStart-$($RdpFrontendPortStart + $studentVmNames.Count - 1)."
}

function Get-Az802StudentEndpoints {
	param(
		[Parameter(Mandatory)][string]$Name,
		[Parameter(Mandatory)][int]$StudentNumber
	)

	$endpointRecords = [Collections.Generic.List[object]]::new()
	$publicIp = Get-AzPublicIpAddress -ResourceGroupName $Name -Name "$Name-lb-pip" -ErrorAction Stop
	for ($index = 0; $index -lt $labVmNames[$LabNumber].Count; $index++) {
		$vmName = $labVmNames[$LabNumber][$index]
		$vm = Get-AzVM -ResourceGroupName $Name -Name $vmName -ErrorAction Stop
		$nicId = $vm.NetworkProfile.NetworkInterfaces[0].Id
		$nic = Get-AzNetworkInterface -ResourceId $nicId -ErrorAction Stop
		$ipConfiguration = $nic.IpConfigurations | Select-Object -First 1
		$rdpPort = $RdpFrontendPortStart + $index

		$endpointRecords.Add([pscustomobject]@{
			StudentNumber = $StudentNumber
			LabNumber = $LabNumber
			ResourceGroupName = $Name
			VmName = $vmName
			IsPrimary = ($vmName -eq $primaryVmNames[$LabNumber])
			PrivateIpAddress = $ipConfiguration.PrivateIpAddress
			PublicIpAddress = $publicIp.IpAddress
			RdpPort = $rdpPort
			RdpTarget = ('{0}:{1}' -f $publicIp.IpAddress, $rdpPort)
			Status = 'Ready'
			Error = ''
		})
	}

	return $endpointRecords.ToArray()
}

try {
	$deploymentFolder = Resolve-Az802DeploymentFolder

	$officialDeploymentScript = Join-Path $deploymentFolder 'deploy-lab-environment.ps1'
	$standardTemplate = Join-Path $deploymentFolder 'azuredeploy.json'
	$lab04Template = Join-Path $deploymentFolder 'azuredeploy-lab04-nested.json'
	$requiredAssets = @(
		$officialDeploymentScript,
		$standardTemplate,
		(Join-Path $deploymentFolder 'deploy-lab04-nested-environment.ps1'),
		$lab04Template,
		(Join-Path $deploymentFolder 'scripts\Initialize-Lab04NestedHost.ps1')
	)
	foreach ($asset in $requiredAssets) {
		if (-not (Test-Path -LiteralPath $asset -PathType Leaf)) {
			throw "Required AZ-802 deployment asset was not found: $asset. Supply a complete, current MicrosoftLearning repository with -RepositoryPath."
		}
	}
	Set-Az802TemplateOsDiskSku -TemplatePath $standardTemplate
	Set-Az802TemplateOsDiskSku -TemplatePath $lab04Template

	foreach ($moduleName in @('Az.Accounts', 'Az.Resources', 'Az.Compute', 'Az.Network')) {
		if (-not (Get-Module -ListAvailable -Name $moduleName)) {
			throw "Required Azure PowerShell module '$moduleName' is missing. Install the Az module with Install-Module Az -Scope CurrentUser, then rerun."
		}
		Import-Module $moduleName -ErrorAction Stop
	}

	if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
		Connect-AzAccount | Out-Null
	}
	if ($SubscriptionId) {
		Set-AzContext -SubscriptionId $SubscriptionId -ErrorAction Stop | Out-Null
	}
	$activeContext = Get-AzContext -ErrorAction Stop
	if (-not $activeContext -or -not $activeContext.Subscription) {
		throw 'No active Azure subscription context is available.'
	}

	if ([string]::IsNullOrWhiteSpace($AllowedRdpSourceIP)) {
		try {
			$publicAddress = (Invoke-RestMethod -Uri 'https://api4.ipify.org' -TimeoutSec 15 -ErrorAction Stop).ToString().Trim()
		}
		catch {
			throw 'Could not discover the current public IPv4 address. Pass -AllowedRdpSourceIP with the trainer/classroom egress IP or CIDR.'
		}
		if ($publicAddress -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
			throw "Public IP discovery returned an invalid IPv4 address '$publicAddress'. Pass -AllowedRdpSourceIP explicitly."
		}
		$AllowedRdpSourceIP = "$publicAddress/32"
	}
	if ($AllowedRdpSourceIP -eq '*') {
		Write-Warning 'RDP (and the template WAC inbound rule) is being exposed to all IPv4 sources. This is intentionally unrestricted; use only a disposable lab and delete it promptly.'
	}
	elseif ($AllowedRdpSourceIP -notmatch '^(?:(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)\.){3}(?:25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(?:/(?:[0-9]|[12][0-9]|3[0-2]))?$') {
		throw "AllowedRdpSourceIP '$AllowedRdpSourceIP' is not a valid IPv4 address/CIDR or '*'."
	}

	$labResourceGroups = @(
		for ($studentNumber = 1; $studentNumber -le $StudentCount; $studentNumber++) {
			'{0}-L{1:D2}-S{2:D2}-ENV' -f $ResourceGroupPrefix, $LabNumber, $studentNumber
		}
	)

	if (-not $CleanStart) {
		$existingGroups = @($labResourceGroups | Where-Object {
			Get-AzResourceGroup -Name $_ -ErrorAction SilentlyContinue
		})
		if ($existingGroups.Count -gt 0) {
			throw "These generated lab resource groups already exist: $($existingGroups -join ', '). Nothing was changed. Use -CleanStart only if you intend to delete and recreate those exact groups."
		}
	}

	$requiredVmCount = if ($LabNumber -eq 4) { $StudentCount } else { $labVmNames[$LabNumber].Count * $StudentCount }
	Write-Host ''
	Write-Host "AZ-802 Lab $LabNumber deployment: $StudentCount student(s), $requiredVmCount Azure VM(s)."
	Write-Host "Subscription: $($activeContext.Subscription.Name) [$($activeContext.Subscription.Id)]"
	Write-Host "Region: $Location | RDP source: $AllowedRdpSourceIP"
	Write-Host "Each student gets a separate resource group/VNet. Deployments run sequentially to reduce ARM throttling."
	Write-Host 'The Microsoft deployment scripts may take 30-45 minutes per student; Lab 4 can take 60-90 minutes per student.'
	Write-Warning 'All environments use the Microsoft lab CONTOSO\Administrator password. The allowed RDP source rule applies equally to every student environment.'
	Write-Warning 'Azure compute, managed disks, and public IPs incur charges. Plan regional vCPU quotas for all students before deployment.'
	if (-not $DisableAutoDelete -and -not $WhatIfPreference) {
		Write-Host "Automatic deletion is enabled for the next $AutoDeleteTime ($AutoDeleteTimeZoneId); scheduling requires Az.Automation and permission to create role assignments."
	}

	$deploymentRecords = [Collections.Generic.List[object]]::new()
	for ($studentNumber = 1; $studentNumber -le $StudentCount; $studentNumber++) {
		$resourceGroupName = $labResourceGroups[$studentNumber - 1]
		Write-Host "`n--- Student $studentNumber of $StudentCount | $resourceGroupName ---" -ForegroundColor Cyan

		if (-not $PSCmdlet.ShouldProcess($resourceGroupName, "Deploy AZ-802 Lab $LabNumber for Student $studentNumber")) {
			$deploymentRecords.Add([pscustomobject]@{
				StudentNumber = $studentNumber; LabNumber = $LabNumber; ResourceGroupName = $resourceGroupName
				VmName = ''; IsPrimary = $false; PrivateIpAddress = ''; PublicIpAddress = ''; RdpPort = ''; RdpTarget = ''
				Status = 'Skipped'; Error = 'WhatIf or confirmation declined.'
			})
			continue
		}

		try {
			$existingGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
			if ($existingGroup) {
				if (-not $CleanStart) {
					throw "Resource group '$resourceGroupName' already exists. Use -CleanStart to explicitly replace it."
				}
				if ($PSCmdlet.ShouldProcess($resourceGroupName, 'Delete existing AZ-802 student resource group')) {
					Remove-AzResourceGroup -Name $resourceGroupName -Force -ErrorAction Stop | Out-Null
					Wait-Az802ResourceGroupDeleted -Name $resourceGroupName
				}
				else {
					throw "Replacement of '$resourceGroupName' was declined; deployment skipped."
				}
			}

			$deployParameters = @{
				LabNumber = $LabNumber
				ResourceGroupName = $resourceGroupName
				Location = $Location
				AdminUsername = $AdminUsername
				AdminPassword = $AdminPassword
				VmSize = $VmSize
				Lab04HostVmSize = $Lab04HostVmSize
				AllowedRdpSourceIP = $AllowedRdpSourceIP
				Lab04VhdUri = $Lab04VhdUri
				ErrorAction = 'Stop'
			}

			Write-Host 'The upstream deployer will initially report temporary VM public IPs; the script removes them after Load Balancer setup. Use the final CSV/summary targets.'
			& $officialDeploymentScript @deployParameters
			Set-Az802StudentLoadBalancer -Name $resourceGroupName -StudentNumber $studentNumber

			$studentEndpoints = @(Get-Az802StudentEndpoints -Name $resourceGroupName -StudentNumber $studentNumber)
			foreach ($endpoint in $studentEndpoints) {
				$deploymentRecords.Add($endpoint)
			}
		}
		catch {
			$failureMessage = $_.Exception.Message
			Write-Warning "Student $studentNumber deployment failed: $failureMessage"
			$deploymentRecords.Add([pscustomobject]@{
				StudentNumber = $studentNumber; LabNumber = $LabNumber; ResourceGroupName = $resourceGroupName
				VmName = ''; IsPrimary = $false; PrivateIpAddress = ''; PublicIpAddress = ''; RdpPort = ''; RdpTarget = ''
				Status = 'Failed'; Error = $failureMessage
			})
		}
	}

	$autoDeleteSchedule = $null
	$autoDeleteScheduleError = $null
	if (-not $DisableAutoDelete -and -not $WhatIfPreference) {
		$resourceGroupsToDelete = @($labResourceGroups | Where-Object {
			Get-AzResourceGroup -Name $_ -ErrorAction SilentlyContinue
		})
		if ($resourceGroupsToDelete.Count -gt 0) {
			Write-Host "Scheduling cleanup of $($resourceGroupsToDelete.Count) generated lab resource group(s) at $AutoDeleteTime ($AutoDeleteTimeZoneId)..."
			try {
				$autoDeleteSchedule = Set-Az802ScheduledCleanup `
					-ResourceGroupNames $resourceGroupsToDelete `
					-SubscriptionId $activeContext.Subscription.Id
			}
			catch {
				$autoDeleteScheduleError = $_.Exception.Message
				Write-Warning "Could not schedule automatic lab deletion: $autoDeleteScheduleError. Delete the lab resource groups manually after class."
			}
		}
	}

	$resolvedOutputCsvPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputCsvPath)
	$deploymentRecords |
		Sort-Object StudentNumber, VmName |
		Export-Csv -LiteralPath $resolvedOutputCsvPath -NoTypeInformation -Encoding utf8

	Write-Host ''
	Write-Host "AZ-802 Lab $LabNumber deployment summary (CSV: $resolvedOutputCsvPath)"
	$deploymentRecords |
		Sort-Object StudentNumber, VmName |
		Format-Table StudentNumber, ResourceGroupName, VmName, IsPrimary, PrivateIpAddress, RdpPort, RdpTarget, Status -AutoSize |
		Out-String |
		Write-Host

	$failedStudents = @($deploymentRecords | Where-Object Status -eq 'Failed')
	if ($failedStudents.Count -gt 0) {
		throw "$($failedStudents.Count) student deployment(s) failed. See the warnings and CSV; successful student environments have been left intact."
	}

	$elapsed = (Get-Date) - $scriptStartedAt
	Write-Host "Completed in $([math]::Floor($elapsed.TotalHours).ToString('00')):$($elapsed.Minutes.ToString('00')):$($elapsed.Seconds.ToString('00'))."
	if ($autoDeleteSchedule) {
		Write-Host "Automatic cleanup: scheduled for $($autoDeleteSchedule.DeleteLocal.ToString('dd-MMM-yyyy HH:mm')) ($($autoDeleteSchedule.TimeZoneId))."
		Write-Host "Automation Account: $($autoDeleteSchedule.AutomationAccountName) in '$($autoDeleteSchedule.AutomationResourceGroupName)'."
	}
	elseif ($DisableAutoDelete) {
		Write-Host 'Automatic cleanup: disabled by -DisableAutoDelete.'
	}
	elseif ($WhatIfPreference) {
		Write-Host 'Automatic cleanup: not scheduled during -WhatIf.'
	}
	else {
		Write-Warning "Automatic cleanup: NOT scheduled. $autoDeleteScheduleError"
	}
	Write-Host 'Manual cleanup (also removes lab resources if the scheduled runbook fails):'
	$labResourceGroups | ForEach-Object { Write-Host "Remove-AzResourceGroup -Name '$_' -Force" }
}
finally {
	if ($temporaryRepositoryRoot -and (Test-Path -LiteralPath $temporaryRepositoryRoot)) {
		Remove-Item -LiteralPath $temporaryRepositoryRoot -Recurse -Force -ErrorAction SilentlyContinue
	}
	$ProgressPreference = $originalProgressPreference
}

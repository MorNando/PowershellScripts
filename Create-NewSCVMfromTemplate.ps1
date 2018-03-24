Param(
	[Parameter(Mandatory=$true)]
	[string]$VMMServer,

	[string]$VMHostGroup,

	[string]$TemplateName,
	[string]$VMName
)
#Functions
Function Provision-SCLocalTemplatedVM {
    param(
        [string]$VMMServer,
        [string]$VMHostGroup,
        [string]$TemplateName,
        [string]$VMName
        
    )

    import-module virtualmachinemanager
    $jobid = [System.Guid]::NewGuid().ToString()

    $Template = Get-SCVMTemplate -VMMServer $VMMServer | where {$_.Name -eq $templatename}
    $VMHost = Get-SCVMHostGroup -Name "All Hosts"
    $virtualMachineConfiguration = New-SCVMConfiguration -VMTemplate $Template -Name $vmname -VMHostGroup $VMHostGroup
    $VHDConfiguration = Get-SCVirtualHardDiskConfiguration -VMConfiguration $virtualMachineConfiguration

    $StorageClassification = Get-SCStorageClassification -Name "Local Storage" 
    Set-SCVirtualHardDiskConfiguration -VHDConfiguration $VHDConfiguration -PinSourceLocation $false -PinDestinationLocation $false -FileName "$vmname - Diff"  -DeploymentOption "UseDifferencing" -StorageClassification $StorageClassification | Out-Null

    Update-SCVMConfiguration -VMConfiguration $virtualMachineConfiguration | Out-Null

    $vm = New-SCVirtualMachine -Name $vmname -ComputerName $VMName -StartVM -VMConfiguration $virtualMachineConfiguration -Description "" -JobGroup $jobid -UseDiffDiskOptimization -StartAction "TurnOnVMIfRunningWhenVSStopped" -StopAction "SaveVM"

}

#Start of tasks
Provision-SCLocalTemplatedVM -templatename $TemplateName -vmname $VMName -VMMServer $VMMServer -VMHostGroup $VMHostGroup
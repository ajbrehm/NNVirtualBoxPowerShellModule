class VirtualBoxVM
{
    [ValidateNotNullOrEmpty()]
    [string]$Name
    [ValidateNotNullOrEmpty()]
    [string]$Uuid
    [string]$State
    [bool]$Running
    [string]$Info
    [string]$GuestOS
}

Update-TypeData -TypeName VirtualBoxVM -DefaultDisplayPropertySet @("Name","UUID","State") -Force

$ptIsolateUuid = "^(.*{)|}$"
$ptVMNameLineTag = "^Name:    "
$ptVMUuidLineTag = "^UUID:    "
$ptVMStateLineTag = "^State:     "
$ptGuestOSLineTag = "^Guest OS:     "

function New-VirtualBoxVMObject
{
    param(
        $aVMInfo
    )

    $sVMName = ([string](($aVMInfo | Where-Object {$_ -match $ptVMNameLineTag}) -replace $ptVMNameLineTag)).Trim()
    $sVMUuid = ([string](($aVMInfo | Where-Object {$_ -match $ptVMUuidLineTag}) -replace $ptVMUuidLineTag)).Trim()
    $sVMState = ([string](($aVMInfo | Where-Object {$_ -match $ptVMStateLineTag}) -replace $ptVMStateLineTag)).Trim()
    $sGuestOS = ([string](($aVMInfo | Where-Object {$_ -match $ptGuestOSLineTag}) -replace $ptGuestOSLineTag)).Trim()
    $vm = New-Object VirtualBoxVM
    $vm.Name = $sVMName
    $vm.Uuid = $sVMUuid
    $vm.State = $sVMState
    $vm.GuestOS = $sGuestOS
    if ($vm.State -like "Running*") {
        $vm.Running = $true
    } else {
        $vm.Running = $false
    }#if
    $vm.Info = $aVMInfo

    return $vm

}

function Get-VirtualBoxVM
{

    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Mandatory=$false)]
        [Switch]$PowerOn,
        [Parameter(Mandatory=$false)]
        [Switch]$ShowConsole,
        [Parameter(Mandatory=$false)]
        [Switch]$PowerOff,
        [Parameter(Mandatory=$false)]
        [Switch]$ShutDown,
        [Parameter(Mandatory=$false)]
        [Switch]$GetExtraData,
        [Parameter(Mandatory=$false)]
        [Switch]$SetExtraData,
        [Parameter(Mandatory=$false)]
        [string]$ExtraDataKey,
        [Parameter(Mandatory=$false)]
        [string]$ExtraDataValue,
        [Parameter(Mandatory=$false)]
        [string]$ModifyVM
    )

    if (!(VBoxManage.exe)) {
        Write-Host "Set path to VBoxManage.exe."
        return 3
    }#if

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }
    if (!$Name) { $Name = "*" }

    if ($Name.Contains("*")){

        $aVirtualBoxVM = @()
        $aVMList = [array](VBoxManage.exe list vms | Where-Object {($_ -replace """" -replace " {.*") -like $Name})
        $aVMList | ForEach-Object {

            $sVMEntry = $_
            $sVMUuid = $sVMEntry -replace $ptIsolateUuid
            $vm = Get-VirtualBoxVM $sVMUuid
            $aVirtualBoxVM += $vm

        }#foreach

        return $aVirtualBoxVM

    } else {

        if ($PowerOn) {
            Write-Host (VBoxManage.exe startvm $Name --type headless)
        } elseif ($ShowConsole) {
            Write-Host (VBoxManage.exe startvm $Name --type separate)
        } elseif ($PowerOff) {
            Write-Host "Powering off VM ""$Name""."
            VBoxManage.exe controlvm $Name poweroff
        } elseif ($ShutDown) {
            Write-Host "Shutting down VM ""$Name""."
            VBoxManage.exe controlvm $Name acpipowerbutton
        } elseif ($GetExtraData) {
            return VBoxManage.exe getextradata $Name $ExtraDataKey
        } elseif ($SetExtraData) {
            return VBoxManage.exe setextradata $Name $ExtraDataKey $ExtraDataValue
        } elseif ($ModifyVM) {
            $vboxmanage = New-Object System.Diagnostics.Process
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "VBoxManage.exe"
            $psi.Arguments = $ModifyVM
            $psi.CreateNoWindow = $true
            $vboxmanage.StartInfo = $psi
            $vboxmanage.Start()
            $vboxmanage.WaitForExit()
        }#if

        $aVMInfo = VBoxManage.exe showvminfo $Name
        $vm = New-VirtualBoxVMObject $aVMInfo
        
        return $vm

    }#if
    
}

function Start-VirtualBoxVM
{
    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Mandatory=$false)]
        [Switch]$ShowConsole
    )

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }

    if ($ShowConsole) {
        return Get-VirtualBoxVM -Name $Name -ShowConsole
    } else {
        return Get-VirtualBoxVM -Name $Name -PowerOn
    }#if
    
}

function Stop-VirtualBoxVM
{
    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Mandatory=$false)]
        [Switch]$PowerOff
    )

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }
    
    if ($PowerOff) {
        return Get-VirtualBoxVM -Name $Name -PowerOff
    } else {
        return Get-VirtualBoxVM -Name $Name -ShutDown
    }#if
    
}

function Open-VirtualBoxVMConsole
{
    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Mandatory=$false)]
        [Switch]$PowerOff
    )

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }

    return Start-VirtualBoxVM -Name $Name -ShowConsole
}

function Submit-VirtualBoxVMProcess
{
    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Position=1,Mandatory=$true)]
        [string]$PathToExecutable,
        [Parameter(Position=2)]
        [string[]]$Arguments,
        [Parameter(Position=3)]
        [pscredential]$Credential
    )

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }
    
    if (!$Credential) {
        $secPassword = Read-Host "Password" -AsSecureString
        $Credential = New-Object pscredential($env:USERNAME,$secPassword)
    }#if

    $secPassword = $Credential.Password
    $usPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secPassword))
    
    return VBoxManage.exe guestcontrol $Name --username $Credential.UserName --password $usPassword run --exe $PathToExecutable -- $PathToExecutable $Arguments    
}

function Submit-VirtualBoxVMPowerShellScript
{
    param(
        [Parameter(Position=0,ParameterSetName="Name")]
        [string]$Name,
        [Parameter(ParameterSetName="Uuid")]
        [string]$Uuid,
        [Parameter(Position=0,ParameterSetName="VirtualBoxVM",ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [VirtualBoxVM]$VirtualBoxVM,
        [Parameter(Position=1,Mandatory=$true)]
        [string]$ScriptBlock,
        [Parameter(Position=3)]
        [pscredential]$Credential
    )

    if ($VirtualBoxVM) { $Name = $VirtualBoxVM.Name }
    if ($Uuid) { $Name = $Uuid }

    return Submit-VirtualBoxVMProcess -Name $Name -PathToExecutable "cmd.exe" -Arguments "/c","powershell","-command",$ScriptBlock -Credential $Credential
    
}


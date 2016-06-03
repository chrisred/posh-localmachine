# LocalMachine

LocalMachine provides some simple functions for managing accounts and settings on a Windows client running PowerShell 2.0 and greater.

Currently the following cmdlets are available:

 > New-LocalUser Get-LocalUser Set-LocalUser Remove-LocalUser
 > 
 > Set-LocalAccountPassword Test-LocalAccountPassword
 > 
 > Add-LocalGroupMember Get-LocalGroupMember Remove-LocalGroupMember New-LocalGroup Get-LocalGroup Set-LocalGroup Remove-LocalGroup
 > 
 > Set-PowerStandbyOptions Set-RemoteDesktopOptions
 > 
 > Import-RegistryHive Remove-RegistryHive

## Installation

From [powershellgallery.com](https://www.powershellgallery.com/GettingStarted) using `Install-Module` from the new PackageManagement module.

`Install-Module -Name LocalMachine`

[https://www.powershellgallery.com/packages/LocalMachine](https://www.powershellgallery.com/packages/LocalMachine)

By placing `LocalMachine.psm1`, `LocalMachine.psd1` and `LocalMachine.format.ps1xml` into one of the default locations PowerShell will search for modules.

1. `$home\Documents\WindowsPowerShell\Modules\LocalMachine` (per user)
2. `C:\Program Files\WindowsPowerShell\Modules\LocalMachine` (all users, PowerShell 4.0 and greater)
3. `%windir%\System32\WindowsPowerShell\v1.0\Modules\LocalMachine` (all users, PowerShell 3.0 and less)

See [about_Modules](https://technet.microsoft.com/en-us/library/hh847804%28v=wps.640%29.aspx) for more information.

## Remote Connections

In a Workgroup the following options are available to get the `-ComputerName` parameter working.

* Using the built-in `Administrator` account with identical credentials on the executing and target machines.
* A user created administrator account with identical credentials and UAC disabled on any target machines.
* A user created administrator account with identical credentials, UAC enabled and the following registry edit to disable [Remote UAC](https://msdn.microsoft.com/en-us/library/windows/desktop/aa826699%28v=vs.85%29.aspx) on any target machines.
`New-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System' -Name 'LocalAccountTokenFilterPolicy' -Value 1 -PropertyType 'DWORD'`

In a Domain remote connections using **domain accounts** are not subject to UAC token filtering and [Remote UAC](https://msdn.microsoft.com/en-us/library/windows/desktop/aa826699%28v=vs.85%29.aspx) is therefore not enforced.

For **both** Workgroup and Domain (check this!) joined machines the Remote Registry service must be started and the Windows Firewall rule groups below must be enabled, most simply using the "Allow and app or feature through Windows Firewall" wizard.

1. Windows Management Instrumentation (WMI)
2. Remote Event Log Management

## Changelog
	1.2
	
	
    1.1
    Fix PowerShell 2.0 compatibility in the module manifest
    Fix errors when Get-LocalGroup and Get-LocalGroupMember returned domain objects under a local user context
    Fix detection of bound parameters to improve compatibility across PowerShell versions
    
    1.0
    Initial release

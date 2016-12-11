# LocalMachine ##

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

## Installation ##

From [powershellgallery.com](https://www.powershellgallery.com/GettingStarted) using `Install-Module` from the new `PackageManagement` module.

`Install-Module -Name LocalMachine`

[https://www.powershellgallery.com/packages/LocalMachine](https://www.powershellgallery.com/packages/LocalMachine)

By placing `LocalMachine.psm1`, `LocalMachine.psd1` and `LocalMachine.format.ps1xml` into one of the default locations PowerShell will search for modules.

1. `$home\Documents\WindowsPowerShell\Modules\LocalMachine` (per user)
2. `C:\Program Files\WindowsPowerShell\Modules\LocalMachine` (all users, PowerShell 4.0 and greater)
3. `%windir%\System32\WindowsPowerShell\v1.0\Modules\LocalMachine` (all users, PowerShell 3.0 and less)

See [about_Modules](https://technet.microsoft.com/en-us/library/hh847804%28v=wps.640%29.aspx) for more information.

## Remote Connections ##

In a Workgroup WinRM it is disabled by default, to enable and test run the following commands.

    Enable-PSRemoting -Force
    Set-Item wsman:\localhost\client\trustedhosts *
    Test-WsMan

The `trustedhosts` key can be a wildcard or a comma separated list of ip addresses or hostnames.

In a Domain WinRM will need to be enabled for clients, but without the need for the `trustedhosts` key modification. For servers WinRM should be enabled by default.

### Remoting without WinRM ###

In a Workgroup the following options are available to get the `-ComputerName` parameter working.

* Using the built-in `Administrator` account with identical credentials on the executing and target machines.
* A user created administrator account with identical credentials and UAC disabled on any target machines.
* A user created administrator account with identical credentials, UAC enabled and the following registry edit to disable [Remote UAC](https://msdn.microsoft.com/en-gb/library/windows/desktop/aa826699%28v=vs.85%29.aspx) on any target machines.
`New-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System' -Name 'LocalAccountTokenFilterPolicy' -Value 1 -PropertyType 'DWORD'`

In a Domain remote connections using **domain accounts** are not subject to UAC token filtering and [Remote UAC](https://msdn.microsoft.com/en-gb/library/windows/desktop/aa826699%28v=vs.85%29.aspx) is therefore not enforced.

For **both** Workgroup and Domain joined clients the `Remote Registry` service must be started and the Windows Firewall rule groups below must be enabled, most simply using the "Allow an app or feature through Windows Firewall" wizard.

1. Windows Management Instrumentation (WMI)
2. Remote Event Log Management

Servers may have some or all of these enabled by default.

## Changelog
    1.3 (2016-12-11)
    Add Power button and lid actions added to Set-PowerStandbyOptions 
    Fix Add-LocalGroupMember can now add members from a domain context with DOMAIN\Member format
    Fix Password parameters now only accept "Security.SecureString" as an input type (in compliance with PSScriptAnalyzer)
	
    1.2 (2016-06-08)
    Add Set-PowerStandbyOptions Set-RemoteDesktopOptions Import-RegistryHive Remove-RegistryHive cmdlets
    Fix Remove-LocalGroupMember can now remove members from a domain context with DOMAIN\Member format
    
    1.1 (2016-02-08)
    Fix PowerShell 2.0 compatibility in the module manifest
    Fix errors when Get-LocalGroup and Get-LocalGroupMember returned domain objects under a local user context
    Fix detection of bound parameters to improve compatibility across PowerShell versions
    
    1.0 (2016-02-02)
    Initial release

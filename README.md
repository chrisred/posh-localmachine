# LocalMachine

LocalMachine provides some simple functions for managing accounts and settings on a Windows client running PowerShell 2.0 and greater.

Currently the following cmdlets are available:

 > New-LocalUser Get-LocalUser Set-LocalUser Remove-LocalUser
 > 
 > Set-LocalAccountPassword Test-LocalAccountPassword
 > 
 > Add-LocalGroupMember Get-LocalGroupMember Remove-LocalGroupMember New-LocalGroup Get-LocalGroup Set-LocalGroup Remove-LocalGroup

## Installation

From [powershellgallery.com](https://www.powershellgallery.com/GettingStarted) using `Install-Module` from the new PackageManagement module.

`Install-Module -Name LocalMachine`

[https://www.powershellgallery.com/packages/LocalMachine](https://www.powershellgallery.com/packages/LocalMachine)

By placing `LocalMachine.psm1` and `LocalMachine.psd1` into one of the default locations PowerShell will search for modules.

1. `$home\Documents\WindowsPowerShell\Modules\LocalMachine` (per user)
2. `C:\Program Files\WindowsPowerShell\Modules\LocalMachine` (all users, PowerShell 4.0 and greater)
3. `%windir%\System32\WindowsPowerShell\v1.0\Modules\LocalMachine` (all users, PowerShell 3.0 and less)

See [about_Modules](https://technet.microsoft.com/en-us/library/hh847804%28v=wps.640%29.aspx) for more information.

## Changelog

    1.0
    Initial release


Add-Type -AssemblyName System.DirectoryServices.AccountManagement

Function New-LocalUser
{
    <#
    .SYNOPSIS
        Create a new local user.
    .DESCRIPTION
        The New-LocalUser cmdlet creates a new local user. Parameters reflect names in the GUI where possible. Two parameter sets control usage of the home folder path, a local path using "HomeFolderLocalPath" and a remote path with "HomeFolderPath" and "HomeFolderDrive".
    .PARAMETER SamAccountName
        Alias Name
        Specifies the user name for the new local user.
    .PARAMETER AccountPassword
        Alias Password
    .PARAMETER ProfilePath
        Specifies a path for the user profile, e.g. 'C:\Profiles\John'.
    .PARAMETER LogonScript
        Specifies a relative path against a share named NETLOGON, e.g. 'Script\Startup.bat'. A share must exist named NETLOGON, e.g. '\\ComputerName\NETLOGON\'.
    .PARAMETER HomeFolderLocalPath
        Specifies a local home folder path.
    .PARAMETER HomeFolderPath
        Specifies a remote home folder path. This must be a resolvable UNC path e.g. '\\SERVER01\Folders\John'.
    .PARAMETER HomeFolderDrive
        Specifies a drive letter to map to the home folder path. This must be a string declaring the drive e.g. 'H:'.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object already exists.
        A terminating error if invalid data is provided, user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        New-LocalUser -SamAccountName John -AccountPassword Password01 -FullName 'John Smith' -UserMustChangePasswordOnNextLogin $true
    .EXAMPLE
        New-LocalUser -SamAccountName John -AccountPassword Password01 -HomeFolderLocalPath 'C:\Folders\John'
    .EXAMPLE
        New-LocalUser -SamAccountName John -AccountPassword Password01 -HomeFolderDrive 'H:' -HomeFolderPath '\\SERVER01\Folders\John'
    #>
    [CmdletBinding(DefaultParametersetName='LocalPath')]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [ValidateLength(1,20)]
        [Alias('Name')]
        [String]$SamAccountName,
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(Position=1,Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [ValidateLength(6,127)]
        [Alias('Password')]
        [String]$AccountPassword,
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$FullName = '',
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$Description = '',
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$ProfilePath = '',
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$LogonScript ='',
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [String]$HomeFolderLocalPath = '',
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$HomeFolderPath = '',
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [String]$HomeFolderDrive = '',
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [Bool]$PasswordNeverExpires = $false,
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [Bool]$UserCannotChangePassword = $false,
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [Bool]$UserMustChangePasswordOnNextLogin = $false,
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='LocalPath')]
        [Parameter(ValueFromPipelineByPropertyName=$True,ParameterSetName='RemotePath')]
        [Bool]$AccountIsDisabled = $false,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Process
    {
        try
        {
            $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
            
            $User = New-Object DirectoryServices.AccountManagement.UserPrincipal($Context)
            $User.SAMAccountName = $SAMAccountName
            $User.SetPassword($AccountPassword)
            $User.DisplayName = $FullName
            $User.Description = $Description
            $User.ScriptPath = $LogonScript
            # choose which parameter sets HomeDirectory
            if ($HomeFolderLocalPath.Length -gt 0) {$User.HomeDirectory = $HomeFolderLocalPath} else {$User.HomeDirectory = $HomeFolderPath}
            # if HomeDrive is set a share path is needed for HomeDirectory, the ParameterSet forces the correct value
            $User.HomeDrive = $HomeFolderDrive
            $User.PasswordNeverExpires = $PasswordNeverExpires
            $User.UserCannotChangePassword = $UserCannotChangePassword
            if ($UserMustChangePasswordOnNextLogin) {$User.ExpirePasswordNow()}
            # matching the parameter name to the GUI is the opposite of the AccountManagement object value
            if ($AccountIsDisabled) {$User.Enabled = $false} else {$User.Enabled = $true}

            # principle object needs to be initialized before calling GetUnderlyingObject()
            $User.Save()
            $User.GetUnderlyingObject().Profile = $ProfilePath
            $User.Save()

            # add all new accounts to the Users group
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, 'Users')
            $Group.GetUnderlyingObject().Add($User.GetUnderlyingObject().Path)
            $Group.Save()
            $Group.Dispose()

            # tidy up principle objects
            $User.Dispose()
            $Context.Dispose()
        }
        catch [DirectoryServices.AccountManagement.PrincipalExistsException]
        {
            # catch specific object exists exception and make it non-terminating
            Write-Error "Error creating the object '$SAMAccountName' on '$ComputerName. $($_.Exception.Message)"
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            # catches all remaning exceptions that might be generated around object modification or access (from .NET and COM sources)
            throw "Error accessing or creating the object '$SAMAccountName' on '$ComputerName'. $($_.Exception.Message)"
        }
    }
}

Function Get-LocalUser
{
    <#
    .SYNOPSIS
        Get a local user or get all local users.
    .DESCRIPTION
        The Get-LocalUser cmdlet gets a local user or gets all local users if no user is defined.

        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None or DirectoryServices.AccountManagement.UserPrincipal on success.
        A terminating error if the SAM database cannot be accessed.
    .EXAMPLE
        Get-LocalUser -Identity John
    #>
    [CmdletBinding()]
    Param(
        [String][Parameter(Position=0,ValueFromPipeline=$True)]$Identity,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            if ($Identity.Length -gt 0)
            {
                [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)
            }
            else
            {
                $Searcher = New-Object DirectoryServices.AccountManagement.PrincipalSearcher(New-Object DirectoryServices.AccountManagement.UserPrincipal($Context))
                $Searcher.FindAll()
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing the object store on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        # for cmdlets that can return an object don't dispose the Context so it is usable
    }
}

Function Set-LocalUser
{
    <#
    .SYNOPSIS
        Modify a local user.
    .DESCRIPTION
        The Set-LocalUser cmdlet modifies the properties of a local user. Parameters that are not selected will not be changed.

        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER SamAccountName
        Alias Name
        Specifies the account name for the user. This can be used to rename a user account.
    .PARAMETER ProfilePath
        Specifies a path for the user profile, e.g. 'C:\Profiles\John'.
    .PARAMETER LogonScript
        Specifies a relative path against a share named NETLOGON, e.g. 'Script\Startup.bat'. A share must exist named NETLOGON, e.g. '\\ComputerName\NETLOGON\'.
    .PARAMETER HomeFolderLocalPath
        Specifies a local home folder path.
    .PARAMETER HomeFolderPath
        Specifies a remote home folder path. This must be a resolvable UNC path e.g. '\\SERVER01\Folders\John'.
    .PARAMETER HomeFolderDrive
        Specifies a drive letter to map to the home folder path. This must be a string declaring the drive e.g. 'H:'.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error if invalid data is provided, user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Set-LocalUser -Identity John -FullName 'John Smith-Roberts' -Description ''
    #>
    [CmdletBinding(DefaultParametersetName='LocalPath')]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='LocalPath')]
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='RemotePath')]
        [String]$Identity,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [ValidateLength(1,20)]
        [Alias('Name')]
        [String]$SamAccountName,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$FullName,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$Description,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$ProfilePath,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$LogonScript,
        [Parameter(ParameterSetName='LocalPath')]
        [String]$HomeFolderLocalPath,
        [Parameter(Mandatory=$true,ParameterSetName='RemotePath')]
        [String]$HomeFolderPath,
        [Parameter(Mandatory=$true,ParameterSetName='RemotePath')]
        [String]$HomeFolderDrive,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [Bool]$PasswordNeverExpires,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [Bool]$UserCannotChangePassword,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [Bool]$UserMustChangePasswordOnNextLogin,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [Bool]$AccountIsDisabled,
        [Parameter(ParameterSetName='LocalPath')]
        [Parameter(ParameterSetName='RemotePath')]
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)
            
            if ($User -ne $null)
            {
                if ($PSBoundParameters.ContainsKey('SamAccountName')) {$User.GetUnderlyingObject().Rename($SamAccountName)}
                if ($PSBoundParameters.ContainsKey('FullName')) {$User.DisplayName = $FullName}
                if ($PSBoundParameters.ContainsKey('Description')) {$User.Description = $Description}
                if ($PSBoundParameters.ContainsKey('ProfilePath')) {$User.GetUnderlyingObject().Profile = $ProfilePath}
                if ($PSBoundParameters.ContainsKey('ScriptPath')) {$User.ScriptPath = $LogonScript}
                if ($PSBoundParameters.ContainsKey('HomeFolderLocalPath'))
                {
                    $User.HomeDirectory = $HomeFolderLocalPath
                    # HomeDrive must be empty if a local home folder path is used
                    $User.HomeDrive = ''
                }

                if ($PSBoundParameters.ContainsKey('HomeFolderPath')) {$User.HomeDirectory = $HomeFolderPath}
                if ($PSBoundParameters.ContainsKey('HomeFolderDrive')) {$User.HomeDrive = $HomeFolderDrive}
                if ($PSBoundParameters.ContainsKey('PasswordNeverExpires')) {$User.PasswordNeverExpires = $PasswordNeverExpires}
                if ($PSBoundParameters.ContainsKey('UserCannotChangePassword')) {$User.UserCannotChangePassword = $UserCannotChangePassword}
                if ($PSBoundParameters.ContainsKey('UserMustChangePasswordOnNextLogin'))
                {
                    if ($UserMustChangePasswordOnNextLogin -eq $true) {$User.ExpirePasswordNow()} else {$User.RefreshExpiredPassword()}
                }

                if ($PSBoundParameters.ContainsKey('AccountIsDisabled'))
                {
                    # matching the parameter name to the GUI
                    if ($AccountIsDisabled) {$User.Enabled = $false} else {$User.Enabled = $true}
                }

                $User.Save()
                $User.Dispose()
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Remove-LocalUser
{
    <#
    .SYNOPSIS
        Remove a local user.
    .DESCRIPTION
        The Remove-LocalUser cmdlet removes a local user.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Remove-LocalUser -Identity John
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)

        try
        {
            if ($User -ne $null)
            {
                $User.Delete()
                $User.Dispose()
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Set-LocalAccountPassword
{
    <#
    .SYNOPSIS
        Change a user account password.
    .DESCRIPTION
        The Set-LocalAccountPassword modifies the password for a local user account.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Remove-LocalUser -Identity John
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [Alias('Name')]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true)][ValidateLength(6,127)]
        [Alias('Password')]
        [String]$AccountPassword,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)

            if ($User -ne $null)
            {
                $User.SetPassword($AccountPassword)
                $User.Save()
                $User.Dispose()
            }
            else
            {
                Write-Error "Cannot find and object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Test-LocalAccountPassword
{
    <#
    .SYNOPSIS
        Test a local account password.
    .DESCRIPTION
        The Remove-LocalUser cmdlet compares a the current password to a given password. This can be useful for identifying insecure common passwords.
        
        The AccountPassword parameter specifies the given password. This is then used to try and reset the account password, success shows the password to be correct.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        True when a password matches.
        False when a password does not match.
        A non-terminating error if the object cannot be found.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Test-LocaUserPassword -Identity John -AccountPassword Password01
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true)]
        [ValidateLength(1,20)]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true)]
        [AllowEmptyString()]
        [Alias('Password')]
        [String]$AccountPassword,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)

    try
    {
        $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)
    
        if ($User -ne $null)
        {
            # if an exception is raised then the password was incorrect or violates a password policy
            $User.ChangePassword($AccountPassword,$AccountPassword)
            return $true
        }
        else
        {
            Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
        }
    }
    catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
    {
        # have to search the Message string as FullyQualifiedErrorId is always 'PasswordException'
        if ($_.Exception.Message.Contains('The password does not meet the password policy requirements.'))
        {
            # we got the password right, but cant change due to local or domain password policy
            return $true
        }
        elseif ($_.Exception.Message.Contains('The specified network password is not correct.'))
        {
            # message here would contain 'The specified network password is not correct.'
            return $false
        }

        # throw an exception for any other error
        throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
    }
}

Function Add-LocalGroupMember
{
    <#
    .SYNOPSIS
        Add one or more members to a local group.
    .DESCRIPTION
        The Add-LocalGroupMember cmdlet adds one or more members to a local group.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a group object by using the SAMAccountName or the SID.
    .PARAMETER Members
        Specifies a set of user objects in a comma-separated list to add to a group.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found or the object is already a memeber of the group.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Add-LocalGroupMember -Identity Administrators -Members John,Paul,Simon
    .EXAMPLE
        'Backup Operators','Remote Desktop Users' | Add-LocalGroupMember -Members John,Paul
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true)]
        [Array]$Members,
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)
            
            if ($Group -ne $null)
            {
                foreach ($Member in $Members)
                {
                    $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Member)
                    if ($User -ne $null)
                    {
                        try
                        {
                            # the DirectoryServices.AccountManagement object raises: Exception "The network path was not found." on domain objects as a local user
                            $Group.GetUnderlyingObject().Add($User.GetUnderlyingObject().Path)
                        }
                        catch [Runtime.InteropServices.COMException]
                        {
                            if ($_.Exception.Message.Contains('The specified account name is already a member of the group.'))
                            {
                                # create a non-terminating error if a user is already a memeber of the group
                                Write-Error "Cannot add object $Member to group '$Identity'. $($_.Exception.Message)"
                            }
                            else
                            {
                                throw
                            }
                        }
                    }
                    else
                    {
                        Write-Error "Cannot find an object with identity '$Member' on '$ComputerName'. This object was not added."
                    }
                }

                $Group.Save()
                $Group.Dispose()
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Get-LocalGroupMember
{
    <#
    .SYNOPSIS
        Get the members of a local group.
    .DESCRIPTION
        The Get-LocalGroupMember cmdlet gets all the members of a local group.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a group object by using the SAMAccountName or the SID.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None or DirectoryServices.AccountManagement.UserPrincipal or GroupPrincipal on success.
        A terminating error if the SAM database cannot be accessed.
    .EXAMPLE
        Get-LocalUser -Identity John
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity = '',
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)
            
            if ($Group -ne $null)
            {
                try
                {
                    # all the members need to be read once to check if any domain objects can't be accessed under a local user
                    $Group.GetMembers() | Out-Null
                    # no error so get the members again and return
                    $Group.GetMembers()
                }
                catch [Management.Automation.RuntimeException]
                {
                    if ($_.Exception.Message.Contains('The network path was not found.'))
                    {
                        Write-Warning "This group contains objects from a Domain context. To return the members of this group as AccountManagement.Principal objects a user with read permissions in the domain is required."
                        $Group.GetUnderlyingObject().Members() | ForEach-Object { ([ADSI]$_).InvokeGet('Name') }
                    }
                    else
                    {
                        throw
                    }
                }
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing the object store on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        # for cmdlets that can return an object don't dispose the Context so it is usable
    }
}

Function Remove-LocalGroupMember
{
    <#
    .SYNOPSIS
        Remove one or more users from a local group.
    .DESCRIPTION
        The Remove-LocalGroupMember cmdlet removes one or more users from a local group.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a group object by using the SAMAccountName or the SID.
    .PARAMETER Members
        Specifies a set of user objects in a comma-separated list to remove from the group.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Remove-LocalGroupMember -Identity Administrators -Members John,Paul,Simon
    .EXAMPLE
        'Backup Operators','Remote Desktop Users' | Remove-LocalGroupMember -Members John,Paul
    #>
    [CmdletBinding()]
    Param(
        [String][Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]$Identity = '',
        [Array]$Members = @{},
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)
        
            if ($Group -ne $null)
            {
                foreach ($Member in $Members)
                {
                    $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Member)
                    if ($User -ne $null)
                    {
                        try
                        {
                            # the DirectoryServices.AccountManagement object raises: Exception "The network path was not found." on domain objects as a local user
                            $Group.GetUnderlyingObject().Remove($User.GetUnderlyingObject().Path)
                        }
                        catch [Runtime.InteropServices.COMException]
                        {
                            if ($_.Exception.Message.Contains('The specified account name is already a member of the group.'))
                            {
                                # create a non-terminating error if a user is already a memeber of the group
                                Write-Error "Cannot remove object $Member from group '$Identity'. $($_.Exception.Message)"
                            }
                            else
                            {
                                throw
                            }
                        }
                    }
                    else
                    {
                        Write-Error "Cannot find an object with identity '$Member' on '$ComputerName'. This object was not removed."
                    }
                }

                $Group.Save()
                $Group.Dispose()
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$Identity' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function New-LocalGroup
{
    <#
    .SYNOPSIS
        Create a local group.
    .DESCRIPTION
        The New-LocalGroup cmdlet creates a new local group. Parameters reflect the GUI names where possible.
    .PARAMETER SAMAccountName
        Alias Name.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if there is an invalid SAMAccountName.
        A terminating error if invalid data is provided, user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        New-LocalGroup -SAMAccountName MyGroup -Description 'My group'
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$True)]
        [Alias('Name')]
        [String]$SAMAccountName,
        [Parameter(Position=1,ValueFromPipelineByPropertyName=$True)]
        [String]$Description = '',
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = New-Object DirectoryServices.AccountManagement.GroupPrincipal($Context)
            $Group.SAMAccountName = $SAMAccountName
            $Group.Save()
            # DirectoryServices.AccountManagement raises the following error under PowerShell 2.0 'Exception setting "Description": "Property is not valid for this store type."'
            $Group.GetUnderlyingObject().Description = $Description
            $Group.Save()
            $Group.Dispose()
        }
        catch [DirectoryServices.AccountManagement.PrincipalExistsException]
        {
            Write-Error "Error creating the object '$SAMAccountName' on '$ComputerName'. $($_.Exception.Message)"
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or creating the object '$SAMAccountName' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Get-LocalGroup
{
    <#
    .SYNOPSIS
        Get a local group or all local groups.
    .DESCRIPTION
        The Get-LocalGroup cmdlet gets a defined local group or gets all local groups if no group identity is defined.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a local group object by using the SAMAccountName or the SID.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None or DirectoryServices.AccountManagement.GroupPrincipal on success.
        A terminating error if the SAM database cannot be accessed.
    .EXAMPLE
        Get-LocalGroup -Identity MyGroup
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,ValueFromPipeline=$True)]
        [String]$Identity,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            if ($Identity.Length -gt 0)
            {
                [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)
            }
            else
            {
                $Searcher = New-Object DirectoryServices.AccountManagement.PrincipalSearcher(New-Object DirectoryServices.AccountManagement.GroupPrincipal($Context))
                $Searcher.FindAll()
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing the object store on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        # for cmdlets that can return an object don't dispose the Context so it is usable
    }
}

Function Set-LocalGroup
{
    <#
    .SYNOPSIS
        Modify a local groups.
    .DESCRIPTION
        The Set-LocalGroup cmdlet modifies the properties of a local group. Parameters that are not selected will not be changed.

        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a local user account by using the SAMAccountName or the SID.
    .PARAMETER SamAccountName
        Alias Name
        Specifies the account name for the group. This can be used to rename a group.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error if invalid data is provided, user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Set-LocalGroup -Identity MyNewGroup -Description 'My new local group'
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [Alias('Name')]
        [String]$SAMAccountName,
        [String]$Description,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)

            if ($Group -ne $null)
            {
                if ($PSBoundParameters.ContainsKey('SamAccountName')) {$Group.GetUnderlyingObject().Rename($SAMAccountName)}
                if ($PSBoundParameters.ContainsKey('Description')) {$Group.GetUnderlyingObject().Description = $Description}
                $Group.Save()
                $Group.Dispose()
            }
            else
            {
                Write-Error "Cannot find and object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$SAMAccountName' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Remove-LocalGroup
{
    <#
    .SYNOPSIS
        Removes a local group.
    .DESCRIPTION
        The Remove-LocalGroup cmdlet removes a local group.
        
        The Identity parameter specifies the account using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a local group by using the SAMAccountName or the SID.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Remove-LocalGroup -Identity MyGroup
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [String]$ComputerName = $env:COMPUTERNAME
    )

    Begin
    {
        $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)
    }

    Process
    {
        try
        {
            $Group = [DirectoryServices.AccountManagement.GroupPrincipal]::FindByIdentity($Context, $Identity)

            if ($Group -ne $null)
            {
                $Group.Delete()
                $Group.Dispose()
            }
            else
            {
                Write-Error "Cannot find an object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException]
        {
            throw "Error accessing or updating the object '$SAMAccountName' on '$ComputerName'. $($_.Exception.Message)"
        }
    }

    End
    {
        $Context.Dispose()
    }
}

Function Import-RegistryHive
{
    <#
    .SYNOPSIS
        Import a registry hive from a file.
    .DESCRIPTION
        The Import-RegistryHive cmdlet imports a registry hive from a file.
        
        An imported hive is loaded into a registry key and then the key is mapped to a PSDrive using the registry provider. The PSDrive is available globally in the current session and must be unloaded using Remove-RegistryHive for it to be fully removed from the session.
    .PARAMETER File
        Specifies the registry hive file to load.
    .PARAMETER Key
        Specifies the registry key to load the hive into, in the format HKLM\MY_KEY or HKCU\MY_KEY
    .PARAMETER Name
        Specifies the name of the PSDrive used to access the hive, excluding the characters ;~/\.:
    .OUTPUTS
        None on success.
        A terminating error if the PSDrive name already exists, the registry hive cannot be loaded or the key cannot be created.
    .EXAMPLE
        Import-RegistryHive -File 'C:\Users\Default\NTUSER.DAT' -Key 'HKLM\TEMP_HIVE' -Name TempHive
        Get-ChildItem TempHive:\
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [String]$File,
        # check the registry key name is not an invalid format
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^(HKLM\\|HKCU\\)[a-zA-Z0-9- _\\]+$')]
        [String]$Key,
        # check the PSDrive name does not include invalid characters
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^[^;~/\\\.\:]+$')]
        [String]$Name
    )

    # check whether the drive name is available
    $TestDrive = Get-PSDrive -Name $Name -EA SilentlyContinue
    if ($TestDrive -ne $null)
    {
        throw [Management.Automation.SessionStateException] "A PSDrive with the name '$Name' already exists."
    }
    
    $Process = Start-Process -FilePath "$env:WINDIR\system32\reg.exe" -ArgumentList "load $Key $File" -WindowStyle Hidden -PassThru -Wait
    
    if ($Process.ExitCode > 0)
    {
        throw [Management.Automation.PSInvalidOperationException] "The registry hive '$File' failed to load. Verify the source path or target registry key."
    }
    
    try
    {
        # validate patten on $Name in the Params and the drive name check at the start make it very unlikely New-PSDrive will fail
        New-PSDrive -Name $Name -PSProvider Registry -Root $Key -Scope Global -EA Stop | Out-Null
    }
    catch
    {
        throw [Management.Automation.PSInvalidOperationException] "A critical error creating PSDrive '$Name' has caused the registy key '$Key' to be left loaded, this must be unloaded manually."
    }
}

Function Remove-RegistryHive
{
    <#
    .SYNOPSIS
        Remove a registry hive loaded via Import-RegistryHive.
    .DESCRIPTION
        The Remove-RegistryHive cmdlet removes a registry hive loaded via Import-RegistryHive.

        The associated PSDrive will be removed and the registry key created during the import will be unloaded.
    .PARAMETER Name
        Specifies the name of the PSDrive used to access the hive.
    .OUTPUTS
        None on success.
        A terminating error if the PSDrive or registry key reasources are still in use.
    .EXAMPLE
        Remove-RegistryHive -Name TempHive
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidatePattern('^[^;~/\\\.\:]+$')]
        [String]$Name
    )
    
    # set -ErrorAction Stop as we never want to proceed if the drive doesnt exist
    $Drive = Get-PSDrive -Name $Name -EA Stop
    # $Drive.Root is the path to the registry key, save this before the drive is removed
    $Key = $Drive.Root
    
    # remove the drive, the only reason this should fail is if the reasource is busy
    Remove-PSDrive $Name -EA Stop
    
    $Process = Start-Process -FilePath "$env:WINDIR\system32\reg.exe" -ArgumentList "unload $Key" -WindowStyle Hidden -PassThru -Wait
    if ($Process.ExitCode)
    {
        # if "reg unload" fails due to the resource being busy, the drive gets added back to keep the original state
        New-PSDrive -Name $Name -PSProvider Registry -Root $Key -Scope Global -EA Stop | Out-Null
        throw [Management.Automation.PSInvalidOperationException] "The registry key '$Key' could not be unloaded, the key may still be in use."
    }
}

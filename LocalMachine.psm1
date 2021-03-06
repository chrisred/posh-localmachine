# LocalMachine v1.3 (c) Chris Redit

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
        Specifies the password to set as a secure string, this cannot be an empty string. ConvertTo-SecureString and Get-Credential can create a secure string object. Set-LocalAccountPassword can set a blank password on an existing account.
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
        New-LocalUser -SamAccountName John -AccountPassword (ConvertTo-SecureString 'Password01' -AsPlainText -Force) -FullName 'John Smith' -UserMustChangePasswordOnNextLogin $true
    .EXAMPLE
        New-LocalUser -SamAccountName John -AccountPassword (ConvertTo-SecureString 'Password01' -AsPlainText -Force) -HomeFolderLocalPath 'C:\Folders\John' -PasswordNeverExpires $true
    .EXAMPLE
        New-LocalUser -SamAccountName John -AccountPassword (ConvertTo-SecureString 'Password01' -AsPlainText -Force) -HomeFolderDrive 'H:' -HomeFolderPath '\\SERVER01\Folders\John'
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
        [Alias('Password')]
        [Security.SecureString]$AccountPassword,
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

            # SetPassword() only accepts plain text input
            $BinaryString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($AccountPassword)
            $User.SetPassword([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BinaryString))
            Remove-Variable -Name BinaryString

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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
        {
            # catches all remaning exceptions that might be generated around object modification or access
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        The Set-LocalAccountPassword modifies the password for a local user account. The password must be defined as a secure string object.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER AccountPassword
        Specifies the password to set as a secure string, this cannot be an empty string.
    .PARAMETER NoPassword
        Indicates that the password for this user account will be blank.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error if the object cannot be found.
        A terminating error user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Set-LocalAccountPassword -Identity John -AccountPassword (ConvertTo-SecureString 'Password01' -AsPlainText -Force)
    .EXAMPLE
        Set-LocalAccountPassword -Identity John -NoPassword
    #>
    [CmdletBinding(DefaultParametersetName='Password')]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='Password')]
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='NoPassword')]
        [Alias('Name')]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true,ParameterSetName='Password')]
        [Alias('Password')]
        [Security.SecureString]$AccountPassword,
        [Parameter(Mandatory=$true,ParameterSetName='NoPassword')]
        [Switch]$NoPassword,
        [Parameter(ParameterSetName='Password')]
        [Parameter(ParameterSetName='NoPassword')]
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
                if ($PSBoundParameters.ContainsKey('AccountPassword'))
                {
                    # SetPassword() only accepts plain text input
                    $BinaryString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($AccountPassword)
                    $User.SetPassword([System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BinaryString))
                    Remove-Variable -Name BinaryString
                }

                if ($PSBoundParameters.ContainsKey('NoPassword'))
                {
                    $User.SetPassword('')
                }
                
                $User.Save()
                $User.Dispose()
            }
            else
            {
                Write-Error "Cannot find and object with identity '$Identity' on '$ComputerName'."
            }
        }
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        The Test-LocalAccountPassword cmdlet compares a the current password to a given password. This can be useful for identifying insecure common passwords. The password must be defined as a secure string object.
        
        The AccountPassword parameter specifies the given password. This is then used to try and reset the account password, success shows the password to be correct.
    .PARAMETER Identity
        Specifies a user object by using the SAMAccountName or the SID.
    .PARAMETER AccountPassword
        Specifies the password to test as a secure string, this cannot be an empty string.
    .PARAMETER NoPassword
        Indicates that the password to test will be blank.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        True when a password matches.
        False when a password does not match.
        A non-terminating error if the object cannot be found.
        A terminating error if user permissions are incorrect or the SAM database cannot be accessed.
    .EXAMPLE
        Test-LocaUserPassword -Identity John -AccountPassword (ConvertTo-SecureString 'Password01' -AsPlainText -Force)
    .EXAMPLE
        Test-LocaUserPassword -Identity John -NoPassword
    #>
    [CmdletBinding(DefaultParametersetName='Password')]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='Password')]
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='NoPassword')]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true,ParameterSetName='Password')]
        [Alias('Password')]
        [Security.SecureString]$AccountPassword,
        [Parameter(Mandatory=$true,ParameterSetName='NoPassword')]
        [Switch]$NoPassword,
        [Parameter(ParameterSetName='Password')]
        [Parameter(ParameterSetName='NoPassword')]
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    $Context = New-Object DirectoryServices.AccountManagement.PrincipalContext('Machine',$ComputerName)

    try
    {
        $User = [DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($Context, $Identity)
    
        if ($User -ne $null)
        {
            if ($PSBoundParameters.ContainsKey('AccountPassword'))
            {
                # ChangePassword() only accepts plain text input
                $BinaryString = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($AccountPassword)
                $User.ChangePassword(
                    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BinaryString),
                    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BinaryString)
                )
                Remove-Variable -Name BinaryString
            }

            if ($PSBoundParameters.ContainsKey('NoPassword'))
            {
                $User.ChangePassword('','')
            }

            # if an exception is raised then the password was incorrect or violates a password policy
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
        The Add-LocalGroupMember cmdlet adds one or more members to a local group. Use DOMAIN\Member to add domain context users or groups.
        
        The Identity parameter specifies the object using the SAMAccountName or the SID.
    .PARAMETER Identity
        Specifies a group object by using the SAMAccountName or the SID.
    .PARAMETER Members
        Specifies a set of user objects in a comma-separated list to add to a group. The DOMAIN\Member format can be used to add members from a domain context, this requires the machine to be a member of the specified domain.
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
    .EXAMPLE
        Add-LocalGroupMember -Identity Administrators -Members John,Paul,'EXAMPLE\Domain Users'
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true)]
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
                    # check for valid DOMAIN\User logon name format, matches a valid NETBIOS domain name and SAMAccountName
                    # https://msdn.microsoft.com/en-us/library/bb726984.aspx https://support.microsoft.com/en-us/kb/909264
                    if ($Member -match '\A([^"/\\\:\|\*\?<>]+)\\([^"/\\\[\]\:;\|=,\+\*\?<>]+)\z')
                    {
                        $ContextString = $Matches[1]
                        $ObjectString = $Matches[2]
                    }
                    else
                    {
                        $ContextString = $ComputerName
                        $ObjectString = $Member
                    }

                    try
                    {
                        # create a string reference for the object to be added as this is the simplest way to handle adding domain context objects
                        $Group.GetUnderlyingObject().Add("WinNT://$ContextString/$ObjectString")
                    }
                    catch [Runtime.InteropServices.COMException]
                    {
                        if ($_.Exception.Message.Contains('The specified account name is already a member of the group.') -or `
                            $_.Exception.Message.Contains('A member could not be added to or removed from the local group because the member does not exist.'))
                        {
                            # create a non-terminating error if the object does not exist or is already a member of the group
                            Write-Error "Cannot add object $Member to group '$Identity'. $($_.Exception.Message)"
                        }
                        else
                        {
                            throw
                        }
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        Specifies a group member by using the SAMAccountName or the SID.
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
                        Write-Warning "This group contains objects from a domain context. To return the members of this group as AccountManagement.Principal objects a user with read permissions in the domain is required."
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        Remove one or more members from a local group.
    .DESCRIPTION
        The Remove-LocalGroupMember cmdlet removes one or more members from a local group. Use DOMAIN\Member to remove domain context users or groups.
        
        The Identity parameter specifies the object using the SAMAccountName or SID.
    .PARAMETER Identity
        Specifies a group object by using the SAMAccountName or the SID.
    .PARAMETER Members
        Specifies a set of members in a comma-separated list to remove from the group. The DOMAIN\Member format can be used to remove members from a domain context.
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
    .EXAMPLE
        Remove-LocalGroupMember -Identity Administrators -Members John,Paul,'EXAMPLE\Domain Users'
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True)]
        [String]$Identity,
        [Parameter(Position=1,Mandatory=$true)]
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
                    if ($Member -match '\A([^"/\\\:\|\*\?<>]+)\\([^"/\\\[\]\:;\|=,\+\*\?<>]+)\z')
                    {
                        $ContextString = $Matches[1]
                        $ObjectString = $Matches[2]
                    }
                    else
                    {
                        $ContextString = $ComputerName
                        $ObjectString = $Member
                    }

                    try
                    {
                        $Group.GetUnderlyingObject().Remove("WinNT://$ContextString/$ObjectString")
                    }
                    catch [Runtime.InteropServices.COMException]
                    {
                        if ($_.Exception.Message.Contains('A member could not be added to or removed from the local group because the member does not exist.'))
                        {
                            # create a non-terminating error if the object is not a member of the group
                            Write-Error "Cannot remove object '$Member' from group '$Identity'. $($_.Exception.Message)"
                        }
                        else
                        {
                            throw
                        }
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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
        catch [DirectoryServices.AccountManagement.PrincipalException], [Runtime.InteropServices.COMException], [UnauthorizedAccessException], [Management.Automation.RuntimeException]
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

Function Set-PowerStandbyOptions
{
    <#
    .SYNOPSIS
        Set power standby options for the machine and display.
    .DESCRIPTION
        The Set-PowerStandbyOptions cmdlet changes the standby timeout for sleep, hibernate and display power saving options.

        By default settings are applied to the AC power plan type (plugged in).
    .PARAMETER Name
        Specifies the power plan by name.
    .PARAMETER Active
        Specifies the active power plan.
    .PARAMETER SleepAfter
        Specifies the idle time in minutes before the machine will sleep, 0 will set "Never".
    .PARAMETER HibernateAfter
        Specifies the idle time in minutes before the machine will hibernate, 0 will set "Never".
    .PARAMETER TurnOffDisplayAfter
        Specifies the idle time in minutes before the display will sleep, 0 will set "Never".
    .PARAMETER LidCloseAction
        Specifies the action when closing a laptop lid, options are: "Do nothing", "Sleep", "Hibernate", "Shut down"
    .PARAMETER PowerButtonAction
        Specifies the action when pressing the power button, options are: "Do nothing", "Sleep", "Hibernate", "Shut down"
    .PARAMETER Battery
        Applies the settings to the DC power plan type (on battery).
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A non-terminating error the power plan does not exist.
        A terminating error if the options cannot be modified or are not supported.
    .EXAMPLE
        Set-PowerStandbyOptions -Active -SleepAfter 30 -TurnOffDisplayAfter 5 -Battery
    .EXAMPLE
        Set-PowerStandbyOptions -Name 'Balanced' -SleepAfter 90 -TurnOffDisplayAfter 30
    #>
    [CmdletBinding(DefaultParametersetName='Active')]
    Param(
        [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$True,ParameterSetName='Named')]
        [Alias('PowerPlan')]
        [String]$Name,
        [Parameter(Mandatory=$true,ParameterSetName='Active')]
        [Switch]$Active,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [Int]$SleepAfter,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [Int]$HibernateAfter,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [Int]$TurnOffDisplayAfter,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [ValidateSet('Do nothing','Sleep','Hibernate','Shut down')]
        [String]$LidCloseAction,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [ValidateSet('Do nothing','Sleep','Hibernate','Shut down')]
        [String]$PowerButtonAction,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [Switch]$Battery,
        [Parameter(ParameterSetName='Named')]
        [Parameter(ParameterSetName='Active')]
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    Process
    {
        try
        {
            if ($Active)
            {
                # -EA Stop is required as Get-WmiObject raises a number of non-terminating errors
                $InstanceId = (Get-WmiObject -Namespace "root\cimv2\power" -Query "SELECT * FROM Win32_PowerPlan WHERE IsActive = True" -ComputerName $ComputerName -EA Stop).InstanceID 
            }
            
            if ($Name.Length -gt 0)
            {
                $InstanceId = (Get-WmiObject -Namespace "root\cimv2\power" -Query "SELECT * FROM Win32_PowerPlan WHERE ElementName = '$Name'" -ComputerName $ComputerName -EA Stop).InstanceID
            }
            
            if ($InstanceId -match '{(.*)}\z')
            {
                $PowerPlanGuid = $Matches[1]
            }
            else
            {
                Write-Error "Cannot find a power plan named '$Name' on '$ComputerName'."
                # stop function execution here
                return $null
            }
            
            # choose power plan type
            $PowerPlanType = 'AC'
            if ($Battery) {$PowerPlanType = 'DC'}
            
            # map action friendly name to the index value
            $PowerActions = @{'Do nothing' = 0; 'Sleep' = 1; 'Hibernate' = 2; 'Shut down' = 3}
            
            if ($PSBoundParameters.ContainsKey('SleepAfter'))
            {
                $PowerSettingGuid = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da' # Sleep after
                $PowerSettingDataIndex = Get-WmiObject `
                    -Namespace "root\cimv2\power" `
                    -Query "SELECT * FROM Win32_PowerSettingDataIndex WHERE InstanceID = 'Microsoft:PowerSettingDataIndex\\{$PowerPlanGuid}\\$PowerPlanType\\{$PowerSettingGuid}'" `
                    -ComputerName $ComputerName

                $PowerSettingValue = ($SleepAfter * 60)
                Set-WmiInstance -InputObject $PowerSettingDataIndex -Arguments @{SettingIndexValue=$PowerSettingValue} | Out-Null
            }
            
            if ($PSBoundParameters.ContainsKey('HibernateAfter'))
            {
                $PowerSettingGuid = '9d7815a6-7ee4-497e-8888-515a05f02364' # Hibernate after
                $PowerSettingDataIndex = Get-WmiObject `
                    -Namespace "root\cimv2\power" `
                    -Query "SELECT * FROM Win32_PowerSettingDataIndex WHERE InstanceID = 'Microsoft:PowerSettingDataIndex\\{$PowerPlanGuid}\\$PowerPlanType\\{$PowerSettingGuid}'" `
                    -ComputerName $ComputerName

                $PowerSettingValue = ($HibernateAfter * 60)
                Set-WmiInstance -InputObject $PowerSettingDataIndex -Arguments @{SettingIndexValue=$PowerSettingValue} | Out-Null
            }
        
            if ($PSBoundParameters.ContainsKey('TurnOffDisplayAfter'))
            {
                $PowerSettingGuid = '3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e' # Turn off display after
                $PowerSettingDataIndex = Get-WmiObject `
                    -Namespace "root\cimv2\power" `
                    -Query "SELECT * FROM Win32_PowerSettingDataIndex WHERE InstanceID = 'Microsoft:PowerSettingDataIndex\\{$PowerPlanGuid}\\$PowerPlanType\\{$PowerSettingGuid}'" `
                    -ComputerName $ComputerName
            
                $PowerSettingValue = ($TurnOffDisplayAfter * 60)
                Set-WmiInstance -InputObject $PowerSettingDataIndex -Arguments @{SettingIndexValue=$PowerSettingValue} | Out-Null
            }

            if ($PSBoundParameters.ContainsKey('LidCloseAction'))
            {
                $PowerSettingGuid = '5ca83367-6e45-459f-a27b-476b1d01c936' # Lid close action
                $PowerSettingDataIndex = Get-WmiObject `
                    -Namespace "root\cimv2\power" `
                    -Query "SELECT * FROM Win32_PowerSettingDataIndex WHERE InstanceID = 'Microsoft:PowerSettingDataIndex\\{$PowerPlanGuid}\\$PowerPlanType\\{$PowerSettingGuid}'" `
                    -ComputerName $ComputerName
            
                $PowerSettingValue = $PowerActions.Item($LidCloseAction)
                Set-WmiInstance -InputObject $PowerSettingDataIndex -Arguments @{SettingIndexValue=$PowerSettingValue} | Out-Null
            }

            if ($PSBoundParameters.ContainsKey('PowerButtonAction'))
            {
                $PowerSettingGuid = '7648efa3-dd9c-4e3e-b566-50f929386280' # Power button action
                $PowerSettingDataIndex = Get-WmiObject `
                    -Namespace "root\cimv2\power" `
                    -Query "SELECT * FROM Win32_PowerSettingDataIndex WHERE InstanceID = 'Microsoft:PowerSettingDataIndex\\{$PowerPlanGuid}\\$PowerPlanType\\{$PowerSettingGuid}'" `
                    -ComputerName $ComputerName
            
                $PowerSettingValue = $PowerActions.Item($PowerButtonAction)
                Set-WmiInstance -InputObject $PowerSettingDataIndex -Arguments @{SettingIndexValue=$PowerSettingValue} | Out-Null
            }
        }
        catch [Management.Automation.RuntimeException]
        {
            # a group policy resulting in settings that cannot be modified could raise this error
            throw "Error accessing or updating the options on '$ComputerName'. $($_.Exception.Message)"
        }
    }
}

Function Set-RemoteDesktopOptions
{
    <#
    .SYNOPSIS
        Set remote desktop options.
    .DESCRIPTION
        The Set-RemoteDesktopOptions cmdlet modifies the remote desktop options.
    .PARAMETER AllowRDP
        Specifies whether Remote Desktop connections will be accepted and the firewall exceptions modified accordingly.
    .PARAMETER AllowNLAOnly
        Specifies whether only connections from computers running Network Level Authentication are accepted.
    .PARAMETER ComputerName
        Runs the cmdlet on the specified computer. The default is the local computer. To successfully run on a remote computer the account executing the cmdlet must have permissions on both machines.
    .OUTPUTS
        None on success.
        A terminating error if the options cannot be modified or are not supported.
    .EXAMPLE
        Set-RemoteDesktopOptions -AllowRDP $true -AllowNLAOnly $false
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [Bool]$AllowRDP,
        [Parameter(Mandatory=$true)]
        [Bool]$AllowNLAOnly,
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    try
    {
        # -EA Stop is required as Get-WmiObject raises a number of non-terminating errors
        $TSSetting = Get-WmiObject -Namespace 'root\cimv2\TerminalServices' -Query "SELECT * FROM Win32_TerminalServiceSetting" -ComputerName $ComputerName -EA Stop
        # TSGeneralSetting returns two objects on Win7 (RDP-Tcp and EH-Tcp) we only want RDP, EH-Tcp appers to be for Media Center Extender compatability
        $TSGeneralSetting = Get-WmiObject -Namespace 'root\cimv2\TerminalServices' -Query "SELECT * FROM Win32_TSGeneralSetting WHERE TerminalName = 'RDP-Tcp'" -ComputerName $ComputerName -EA Stop

        if ($AllowRDP)
        {
            # enable RDP and modify firewall excpetions
            $TSSetting.SetAllowTSConnections(1,1) | Out-Null
        }
        else
        {
            $TSSetting.SetAllowTSConnections(0,1) | Out-Null
        }

        if ($AllowNLAOnly)
        {
            $TSGeneralSetting.SetUserAuthenticationRequired(1) | Out-Null
        }
        else
        {
            $TSGeneralSetting.SetUserAuthenticationRequired(0) | Out-Null
        }
    }
    catch [Management.Automation.RuntimeException]
    {
        if ($_.Exception.Message.Contains('Invalid namespace'))
        {
            throw "Error accessing or updating the options on '$ComputerName'. It is likely this version of Windows does not support Remote Desktop Services."
        }
        else
        {
            throw "Error accessing or updating the options on '$ComputerName'. $($_.Exception.Message)"
        }
    }
}

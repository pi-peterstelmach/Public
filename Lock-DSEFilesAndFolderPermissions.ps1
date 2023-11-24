function Lock-DSEFilesAndFolderPermissions {
    <#
        .SYNOPSIS
        This command set all permissions on the specified path to read only. The ServiceUser will be the owner.
        
        .DESCRIPTION
        Long description
        
        .PARAMETER Path
        Is the file / folder where the permissions will be modified.
        
        .PARAMETER ServiceUser
        The Owner of the file or folder
                
        .NOTES
        General notes
    #>
    [CmdletBinding()]
    param (
        # Path
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [STRING]
        $Path,

        # ServiceUser
        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false
        )]
        [STRING]
        $ServiceUser = "$($env:USERDOMAIN)\$($env:USERNAME)"
    )
    Begin{
        Write-Verbose -Message "Overwrite owner is '$($ServiceUser)'."
        [Object]$OwnerObject = New-Object System.Security.Principal.NTAccount($ServiceUser)
    }
    Process{
        Try{
            Write-Verbose -Message "Fetch access control list from $($Path)"
            [Object]$ItemACL = Get-ACL -Path $Path
            #BackUp the existing access control list. This list will be manipulated with the read only permissions.
            [Array]$BackupACL = $ItemACL.Access
            Write-Verbose -Message "Disable permission inheritance from parent folder and remove all permissions."
            $ItemACL.SetAccessRuleProtection($true, $false)
            Write-Verbose -Message "(Re)add 'read' permissions to access control list."
            $BackupACL.ForEach{
                $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$($_.IdentityReference.Value)","Read", "$($_.InheritanceFlags)", "None",  "$($_.AccessControlType)")
                $ItemACL.SetAccessRule($AccessRule)
            }
            Write-Verbose -Message "Set owner of to service user '$($ServiceUser)'."
            $ItemACL.SetOwner($OwnerObject)
            Write-Verbose -Message "Apply new access rule to $($Path)."
            $ItemACL | Set-ACL -Path $Path
        }
        Catch{
            Write-Host "$($_)"
            Write-Host "Press any key to continue."
            Read-Host
        }
    }
    End{

    }
}
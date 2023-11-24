function Restore-DSEFilesAndFolderPermissions {
    <#
        .SYNOPSIS
        This command restores the permission on a file or folder
        
        .DESCRIPTION
        
        .PARAMETER Object
        The SQL object with the path and permissions. 
        
        .EXAMPLE
        An example
        
        .NOTES
        General notes
    #>
    [CmdletBinding()]
    param (

        # Object
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [Object]
        $Object
    )
    Begin{
        #nothing here.
    }
    Process{
        Try{
            Write-Verbose -Message "Restore acl on '$($Object.FullName)'."
            [Object]$ItemACL = Get-Acl -Path $Object.FullName
            #Maybe this line is not mandatory.
            #In production no inherited permissions should be applied, and so nothing can be removed. 
            #But in Test this is not the case. This line should not crash anything.
            Write-Verbose -Message "Disable permission inheritance from parent folder and remove all permissions."
            $ItemACL.SetAccessRuleProtection($true, $false)
            #At first all rule will be delete and then the new rules are applied.    
            $ItemACL.Access.foreach{
                $ItemACL.RemoveAccessRuleAll($_)
            }
            Write-Verbose -Message "Add access rule."
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$($Object.IdentityReference)","$($Object.FileSystemRights)", "$($Object.InheritanceFlags)", "$($Object.PropagationFlags)",  "$($Object.AccessControlType)")
            $ItemACL.AddAccessRule($AccessRule)
            Write-Verbose -Message "(Re)set owner."
            [Object]$OwnerObject = New-Object System.Security.Principal.NTAccount($Object.Owner)
            $ItemAcl.SetOwner($OwnerObject)
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
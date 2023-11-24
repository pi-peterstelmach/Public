<#
    .SYNOPSIS
    This script retrieves permissions for files and folders at a specified path and creates custom PSObjects.

    .DESCRIPTION
    The Get-FilesAndFoldersPermissions function retrieves permissions for files and folders at the specified path and returns custom PSObjects with various properties, including FullName, PSParentPath, PSChildName, Owner, Type, and more.

    .PARAMETER Path
    The path for which to retrieve permissions.

    .EXAMPLE
    Get-FilesAndFoldersPermissions -Path "C:\ExampleFolder"
#>



##### TODO
# WriteToDB entfernen

function Get-DSEFilesAndFoldersPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $false)]
        [ValidateNotNull()]
        [string]$Path,

        [switch]$WriteToDB
    )

    $objects = @()  # Initialize the $objects array to store the custom objects.

    Get-Childitem -Path $Path -Recurse | ForEach-Object {
        $item = $_
        $acl = Get-Acl -Path $item.FullName
        $owner = $acl.Owner
        $access = $acl.Access

        Foreach ($entry in $access) {
            $object = New-Object PSCustomObject
            Add-Member -InputObject $object -NotePropertyMembers @{
                FullName = $acl.Path
                PSParentPath = $item.PSParentPath
                PSChildName = $item.PSChildName
                CentralAccessPolicyId = $acl.CentralAccessPolicyId
                CentralAccessPolicyName = $acl.CentralAccessPolicyName
                Path = $acl.Path
                Owner = $acl.Owner
                UserEnabled = $entry.UserEnabled
                FileSystemRights = $entry.FileSystemRights
                AccessControlType = $entry.AccessControlType
                IdentityReference = $entry.IdentityReference
                IsInherited = $entry.IsInherited
                InheritanceFlags = $entry.InheritanceFlags
                PropagationFlags = $entry.PropagationFlags
            } # Add-Member

            # Determine the 'Type' (User/Group) based on the Owner (local user or Active Directory)
            
            # Initialize empty ownerType before the Check
            $ownerType = $null

            $ownerObject = New-Object System.Security.Principal.NTAccount($owner)
            $sid = $ownerObject.Translate([System.Security.Principal.SecurityIdentifier])
            $principal = $sid.Translate([System.Security.Principal.NTAccount])
            $ownerName = $principal.Value

            # Check if Owner exists in local system users or groups
            if (Get-WmiObject -Class Win32_UserAccount | Where-Object { $_.SID -eq $sid }) {$ownerType = "User"}
            elseif (Get-WmiObject -Class Win32_Group | Where-Object { $_.SID -eq $sid }) {$ownerType = "Group"}

            # Else check in Active Directory
            elseif ($null -eq $ownerTyp) {               
                
                $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
                $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $domain.Name)
                $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher($context)
                $directorySearcher.Filter = "(&(objectSid=$sid))"
                $result = $directorySearcher.FindOne()

                if ($result.Properties['objectCategory'] -eq 'group') {$ownerType = "Group"}
                elseif ($result.Properties['objectCategory'] -eq 'user') {$ownerType = "User"}
                
                # If ownerType Check in AD not User or Group set it to $null
                else {$ownerType = $null}

            } #elseif

            $object | Add-Member -MemberType NoteProperty -Name "OwnerType" -Value $ownerType
            $objects += $object
            $object = $null
        } #Foreach
    }


    if (-not $WriteToDB) {
        Write-Host -ForegroundColor Green "Output Objects from $($Path):"
        Write-Host -ForegroundColor Green "Objects count: $($objects.count)"
        return $objects
        
    }

    if ($WriteToDB) {

        try {
            $serverInstance = 'localhost\sqlexpress'
            $database = 'project'
            $tableName = '[GetFilesAndFoldersPermissions]'
        
            # Using SQL authentication
            $sqlUsername = 'scriptuser'
            $sqlPassword = 'script'
            $connectionString = "Data Source=$serverInstance;Database=$database;User ID=$sqlUsername;Password=$sqlPassword;TrustServerCertificate=True"


            ForEach ($object in $objects) {

                $valuesList = @(
                    $object.Owner, $object.FullName, $object.PropagationFlags, $object.CentralAccessPolicyName,
                    $object.InheritanceFlags, $object.CentralAccessPolicyId, $object.PSChildName, $object.PSParentPath,
                    $object.Path, $object.IsInherited, $object.FileSystemRights, $object.UserEnabled, $object.AccessControlType,
                    $object.IdentityReference, $object.OwnerType
                )

            # Create values list for SQL Insert Command, e.g. ('value1','value2',...)          
            $valuesListJoined = "'" + $($valuesList -join "','") + "'"

            ##DEBUG
            #Write-Host "valuesList: $($valuesList)"

            $insertQuery = "INSERT INTO $tableName (Owner, FullName, PropagationFlags, CentralAccessPolicyName, InheritanceFlags, CentralAccessPolicyId, PSChildName, PSParentPath, Path, IsInherited, FileSystemRights, UserEnabled, AccessControlType, IdentityReference, OwnerType) VALUES ($($valuesListJoined))"
            Invoke-Sqlcmd -ConnectionString $connectionString -Query $insertQuery
            

            ##VERIFY INSERT
            $verifyQuery = "SELECT * FROM $tableName
                                WHERE
	                                Owner = '$($valuesList[0])'
	                                AND FullName = '$($valuesList[1])'
	                                AND PropagationFlags = '$($valuesList[2])'
	                                AND CentralAccessPolicyName = '$($valuesList[3])'
	                                AND InheritanceFlags = '$($valuesList[4])'
	                                AND CentralAccessPolicyId = '$($valuesList[5])'
	                                AND PSChildName = '$($valuesList[6])'
	                                AND PSParentPath = '$($valuesList[7])'
	                                AND Path = '$($valuesList[8])'
	                                AND IsInherited = '$($valuesList[9])'
	                                AND FileSystemRights = '$($valuesList[10])'
	                                AND UserEnabled = '$($valuesList[11])'
	                                AND AccessControlType = '$($valuesList[12])'
	                                AND IdentityReference = '$($valuesList[13])'
                                    AND OwnerType = '$($valuesList[14])'"

                Write-Host -ForegroundColor Cyan "Checking SQL INSERT Operation for: $($valuesListJoined)."
                $verifyCurrentSQLInsert = Invoke-Sqlcmd -ConnectionString $connectionString -Query $verifyQuery | Format-List
                
                if ($verifyCurrentSQLInsert) {
                    Write-Host -ForegroundColor Green "SQL INSERT: OK"
                }
                else {
                    Write-Host -ForegroundColor Red "SQL INSERT: NOT OK"
                }

            } #Foreach
        } #try
        catch {
            Write-Host "Error: $_"
        }
    } #if
} #function
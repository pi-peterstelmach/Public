<#
.SYNOPSIS
   This script imports files and folders from a specified path and returns custom PSObjects with file and folder details.
.DESCRIPTION
   The script searches for files and folders in the specified path and returns custom PSObjects for each item found.
   The custom PSObject includes properties such as Name, FullName, PathLength, Extension (for files), FileSize (for files), Type, Hash (for files), etc.
.PARAMETER Path
   The path to search for files and folders.
.PARAMETER Recurse
   If specified, includes all child objects in all subfolders within the provided directory.
.PARAMETER Filter
   A filter to specify file extensions for filtering the results.
.EXAMPLE
   Get-AfilitateFilesAndFolder -Path "C:\ExampleFolder"
   Get-AfilitateFilesAndFolder -Path "C:\ExampleFolder" -Filter ".txt"
#>

function Get-DSEAfilitateFilesAndFolder {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline=$false, 
            HelpMessage="Specify the path")]
        [string]$Path,

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$True,
            HelpMessage="Specify file extention(s)")]
        [string]$Filter
    )

    if (-not (Test-Path -Path $Path -PathType Container)) {
        Write-Error "The specified path '$Path' does not exist or is not a valid directory."
        return
    }

    Get-ChildItem -Path $Path -Recurse | Where-Object {
        if (-not $Filter) {
            Write-Verbose -Message "No Filter specified."
            $true  # If no filter is specified, return all items
        } else {
            $item = $_
            if ($item.PSIsContainer) {
                $true  # Always include folders
            } else {
                $item.Extension -eq $Filter
            }
        }
    } | ForEach-Object {
        Write-Verbose -Message "Creating PSCustomObject for: $($item.Name)"
        $item = $_
        $object = New-Object PSObject -Property @{
            
            Name = $item.Name
            FileHash = if (-not $item.PSIsContainer) {
                $hash = (Get-FileHash -Path $item.FullName -Algorithm SHA256).Hash
            }
            else { $null }

            FullName = $item.FullName
            Exists = $item.Exists
            Extension = if ($item.PSIsContainer) { $null } else { $item.Extension }
            CreationTime = $item.CreationTime 
            LastAccessTime = $item.LastAccessTime
            LastWriteTime = $item.LastWriteTime
            LastWriteTimeUtc = $item.LastWriteTimeUtc
            Attributes = $item.Attributes
            PSPath = $item.PSPath
            PSParentPath = $item.PSParentPath
            PSChildName = $item.PSChildName
            PSIsContainer = $item.PSIsContainer
            Mode = $item.Mode
            BaseName = $item.BaseName
            LinkType = $item.LinkType
            PathLength = $item.FullName.Length
            FileSize = if ($item.PSIsContainer) { 0 } else { "{0:N2} MB" -f ($item.Length / 1MB) }
            #Type = if ($item.PSIsContainer) { "Folder" } else { "File" }
        }
        return $object
    }
}

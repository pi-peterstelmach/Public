<#
.SYNOPSIS
    Retrieves specific attributes of Microsoft Office files (XLS, XLSX, DOC, DOCX, PPT, PPTX) and returns them as [PSCustomObject].

.DESCRIPTION
    This PowerShell function identifies specified file types (XLS, XLSX, DOC, DOCX, PPT, PPTX) within the specified path and retrieves the properties "PasswordProtection," "HyperLinks," "Author," and "FullName."

.PARAMETER Path
    Specifies the path where the files are located.

.EXAMPLE
    Get-DSEOfficeFileAttributes -Path "C:\Files"

.NOTES
    - Requires Microsoft Office applications to be installed for accessing file properties.
    - The function uses COM objects for Microsoft Office applications.
#>


function Get-DSEOfficeFileAttributes {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false, 
                   HelpMessage="Specify the path")]
        [string]$Path
    )

    BEGIN {
        Write-Verbose -Message "Starting to process function 'Get-DSEOfficeFileAttributes' for Path: $($Path)..."
    }

    PROCESS {
        # Check the specified path and collect file properties
        if (Test-Path $Path -PathType Container) {

            Write-Verbose -Message "Getting all office files in $($Path)."
            $files = Get-ChildItem -Path $Path -Include *.xls, *.xlsx, *.doc, *.docx, *.ppt, *.pptx -Recurse -File
                    
            $fileProperties = foreach ($file in $files) {
                Write-Verbose -Message "Getting file properties for file: $($file.FullName)."
                Get-FileProperty -file $file
            }
            Write-Verbose -Message "Returning all file property results."
            return $fileProperties  # Return the result
        } else {
            Write-Verbose -Message "Path not found or is not a directory: $($Path)"
        }
    } #PROCESS

    END {}

} #Get-DSEOfficeFileAttributes
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
    param (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$false, 
                   HelpMessage="Specify the path")]
        [string]$Path
    )

    # Function to check file properties
    function Get-FileProperty {
        param($File)

        $fileObj = New-Object -TypeName PSObject
        $fileObj | Add-Member -MemberType NoteProperty -Name "FullName" -Value $file.FullName
        
        Write-Verbose -Message "Processing file: $($file.FullName)..."

        try {

            # Start a background job to open the document
            $job = Start-Job -ScriptBlock {
                param ($filePath)
                Write-Verbose -Message "Opening document in background job..."
            
                # Create ComObject for Office Apps
                if ($file.Extension -eq ".doc" -or $file.Extension -eq ".docx") {
                    $ComObjectApp = "Word.Application"
                    $app = New-Object -ComObject $ComObjectApp
                    $officeDoc = $ComObjectApp.Documents.Open($filePath, $null, $null, $null, "")
                }
                elseif ($file.Extension -eq ".xls" -or $file.Extension -eq ".xlsx") {
                    $ComObjectApp = "Excel.Application"
                    $app = New-Object -ComObject $ComObjectApp
                    $officeDoc = $ComObjectApp.Workbooks.Open($filePath, $null, $null, $null, "")
                }
                elseif ($file.Extension -eq ".ppt" -or $file.Extension -eq ".pptx") {
                    $ComObjectApp = "PowerPoint.Application"
                    $app = New-Object -ComObject $ComObjectApp
                    $officeDoc = $ComObjectApp.Presentations.Open($filePath, $null, $null, $null, "")
                }
            
                $app.Visible = $false
            
                return $officeDoc
            } -ArgumentList $file.FullName

            # Wait for the job to complete or time out after 10 seconds
            $result = Wait-Job $job -Timeout 10

            if ($result.State -eq 'Completed') {
                # Retrieve the job output

                $hyperlinksList = $officeDoc.HyperLinks | Foreach {
                    $hyperlinksList += $($_.Address)
                }

                $fileObj | Add-Member -MemberType NoteProperty -Name "PasswordProtection" -Value $($officeDoc.ProtectionType)
                $fileObj | Add-Member -MemberType NoteProperty -Name "HyperLinks" -Value $($hyperlinksList)
                $fileObj | Add-Member -MemberType NoteProperty -Name "Author" -Value $($officeDoc.Author)

                Write-Verbose -Message "Document opened successfully."
            }
            else {
                # If $job State not 'Completed' and still running

                Write-Verbose -Message "Force quitting Job"
                Remove-Job $job -Force

                # Quit Job after 30 sec of retries
                $timeout = 0
                while ($result.State -eq 'Running' -and $timeout -lt 30) {
                    Write-Verbose -Message "Job still running, waiting for force quit to finish..."
                    Start-Sleep 10
                    $timeout += 10
                }
            }
        }
        catch {
            Write-Error -Message "Error processing file: $($file.FullName)"
            Write-Verbose -Message "File is protected or an error occurred: $($file.FullName)."

            # Handle the situation when the file is protected or an error occurred during the process
            $fileObj | Add-Member -MemberType NoteProperty -Name "PasswordProtection" -Value 'File protected or error occurred'
            $fileObj | Add-Member -MemberType NoteProperty -Name "HyperLinks" -Value 'File protected or error occurred'
            $fileObj | Add-Member -MemberType NoteProperty -Name "Author" -Value 'File protected or error occurred'
        }
        finally {
            $officeDoc.Close()
            Write-Verbose -Message "Processing completed for file: $($file.FullName)"
        }

    return $fileObj
    } #Get-FileProperty

    # Check the specified path and collect file properties
    if (Test-Path $Path -PathType Container) {
        $files = Get-ChildItem -Path $Path -Include *.xls, *.xlsx, *.doc, *.docx, *.ppt, *.pptx -Recurse -File
        $fileProperties = foreach ($file in $files) {
            Get-FileProperty -file $file
        }

        $fileProperties  # Return the result
    } else {
        Write-Verbose -Message "Path not found or is not a directory: $($Path)"
    }

} #Get-DSEOfficeFileAttributes
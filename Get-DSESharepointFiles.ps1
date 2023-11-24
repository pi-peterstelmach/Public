#Function to get all files of a folder
Function Get-DSESharepointFiles(){
    [CmdletBinding()]
    param (

        # OnPremisePath
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True
        )]
        [STRING]
        $OnPremisePath,

        # URL
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True
        )]
        [STRING]
        $URL,

        # Credentials
        [Parameter(
            Mandatory=$True,
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True
        )]
        [System.Management.Automation.PSCredential]
        $Credentials
    )
    Begin{
        Try{
            #This command depends on 'Microsoft.Online.SharePoint.PowerShell'
            Write-Verbose -Message "Import sharepoint pnp powershell module."
            Import-Module PNP.Powershell
            #
            Write-Verbose -Message "Login to sharepoint online."
            Connect-PnPOnline -Credentials $Credentials -Url $URL
        }
        Catch{
            Write-Host $_
        }
    }
    Process{
        #Get the root folder of the Library
        $Folder = Get-PnPFolder -Url "/Shared Documents"
        #Call the function to download the document library
        $SharepointHashes = Get-DSESharepointFileHash -Folder $Folder -WorkFolder $env:TEMP
        foreach($SharepointHash in $SharepointHashes){
            Write-Verbose -Message "Add fullname to calculated sharepoint hash."
            $SharepointHash["FullName"] = $OnPremisePath + $SharepointHash.SharepointURl.Substring($SharepointHash.SharepointURL.IndexOf("Shared Documents") +17)
        }
    }
    End{
        return $SharepointHashes
    }
}
function Get-SQLObject {
    <#
        .SYNOPSIS
        Short description
        
        .DESCRIPTION
        Long description
        
        .PARAMETER Object
        Das Object welches angelegt / aktualisert werden soll.
        
        .PARAMETER Table
        Die Tabelle in der das Objekt gespeichert werden soll.
        
        .PARAMETER PrimaryKey
        Wenn der PrimaryKey angegeben wird und das Objekt mit dem PrimaryKey bereits in der SQL DB existiert, dann wird dieses Objekt updaten.
        
        .EXAMPLE
        An example
        
        .NOTES
        General notes
    #>
    [CmdletBinding()]
    param (
        
        # Table
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [STRING]
        $Table,

        # ConfigurationFile
        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false
        )]
        [ValidateScript({Test-Path -Path $_})]
        [STRING]
        $ConfigurationFile = ".\bin\config.json"
    )
    Begin{
        Try{
            #import sql connection configuration from configuration file.
            Write-Verbose -Message "Iniztialize configuration file $($ConfigurationFile)."
            [Object]$Configuration = (Get-Content -Path $ConfigurationFile) | ConvertFrom-Json

            Write-Verbose -Message "Initialize sql connection."
            $SQLConnection = New-Object System.Data.SQLClient.SQLConnection
            $SQLConnection.ConnectionString = "server='$($Configuration.SQLConnection.ServerName)';database='$($Configuration.SQLConnection.DatabaseName)';trusted_connection=true;"
            $SQLConnection.Open()

            $SQLCommand = New-Object System.Data.SQLClient.SQLCommand
            $SQLCommand.Connection = $SQLConnection
        }
        Catch{
            Write-Host ("Error occured: $($_)") -ForegroundColor Red
            Wait-Event
        }
    }
    Process{
        $QUERY = "Select * from $($Table)."
        Write-Host "SQLCommand: $($Query)"
        $SQLCommand.CommandText = $QUERY
        $results = $SQLCommand.ExecuteReader()
    }
    End{
        return $results
        $SQLConnection.Close();
    }
}
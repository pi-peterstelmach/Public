function Add-DSESQLObject {
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
        # Object
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [PSCustomObject]
        $Object,
        
        # Table
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [STRING]
        $Table,

        # PrimaryKey
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [STRING]
        $PrimaryKey,

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
            if($PrimaryKey){
                [Object[]]$ExistingObjects = .\Get-SQLObject.ps1 -Table $Table -ConfigurationFile $ConfigurationFile   
            }
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
        if($ExistingObjects.$PrimaryKey -contains $Object.$PrimaryKey){
            #Update
            Write-Host "Update exsting object with primary key value $($Object.$PrimaryKey)."
            $QUERY = "UPDATE $($Table) SET "
            $QUERY += $Object.psobject.BaseObject.Keys.foreach{"$_ = '$($Object.$_)',"} -join " "
            $QUERY = $QUERY.TrimEnd(",")
            $QUERY += " Where $($PrimaryKey) = $($Object.$PrimaryKey)."
        }
        Else{
            #Add
            $QUERY = "INSERT INTO $($Table) ('$($Object.psobject.BaseObject.Keys -join "','")') VALUES ('$($Object.psobject.BaseObject.Values -join "','"))"
            Write-Host "SQLCommand: $($Query)"
        }
        $SQLCommand.CommandText = $QUERY
        $SQLCommand.ExecuteNonQuery()
    }
    End{
        $SQLConnection.Close();
    }
}
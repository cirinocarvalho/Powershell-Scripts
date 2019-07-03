<###
Created by: Cirino Carvalho
Date: 07/01/2019
Purpose: Import Data from Rest API and Insert into SQL Server Database
###>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#-----------------------------------------------------------[Declarations]----------------------------------------------------------------
#Script Version
$sScriptVersion = "1.0"

#Create the folder Logs in inside the folder where the script is located
$sLogPath = ("$PSScriptRoot\Logs\")
If (!(test-path $sLogPath)) {
    New-Item -ItemType Directory -Force -Path $sLogPath
}

#Log File Info
$sLogName = ("BigBelly_Clean_Assets_" + (Get-Date -Format 'MMddyyyy') + ".log")
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName


#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function JsonData {
    Param()
    
   
    Begin {
        Log-Write -LogPath $sLogFile -LineValue ("Start Get-JsonData From BigBelly API")
    }
    
    Process {
        Try {

            $uri = 'https://api'
            $headers = @{
                'X-Token'       = ''
                'Cache-Control' = 'no-cache'
            }
            
            $body = @{ }

            $JsonResult = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -Body $body

            $dataTable = $JsonResult.assets | Select-Object -Property `
                                                            accountId,
                                                            latestFullness, 
                                                            reason, 
                                                            serialNumber,
                                                            description, 
                                                            position, 
                                                            stationSerialNumber, 
                                                            ageThreshold, 
                                                            fullnessThreshold,
                                                            latitude,
                                                            longitude | Out-DataTable
            #Export Data to SQL Server
            ExportJsonDataToSQL($dataTable)
        }
        Catch {
            Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
            Break
        }
    }

    End {
        If ($?) {
            Log-Write -LogPath $sLogFile -LineValue "Get-JsonData Completed Successfully."
            Log-Write -LogPath $sLogFile -LineValue " "
        }
    }

}

Function ExportJsonDataToSQL {
    Param($dataTable)
       
   
    Begin {
        Log-Write -LogPath $sLogFile -LineValue ("Start ExportJsonDataToSQL")
    }
    
    Process {
        Try {
            $conn = New-Object System.Data.SqlClient.SqlConnection
            #$conn.ConnectionString = "Data Source=.x.local; Integrated Security=True;Initial Catalog=db;"
            $conn.ConnectionString = "Data Source=x.local; User Id=user; password=pass;Database=db;"

            $sqlBulkCopy = New-Object ("Data.SqlClient.SqlBulkCopy") -Args $conn
            $sqlBulkCopy.DestinationTableName = "dbo.Clean_Assets"

            <#
            $ColumnMap1 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(0, 0)
            $ColumnMap2 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(1, 1)
            $ColumnMap3 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(2, 2)
            $ColumnMap4 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(3, 3)
            $ColumnMap5 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(4, 4)
            $ColumnMap6 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(5, 5)
            $ColumnMap7 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(6, 6)
            $ColumnMap8 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(7, 7)
            $ColumnMap9 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(8, 8)
            $ColumnMap10 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(9, 9)
            $ColumnMap11 = New-Object System.Data.SqlClient.SqlBulkCopyColumnMapping(10, 10)

            $sqlBulkCopy.ColumnMappings.Add($ColumnMap1)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap2)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap3)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap4)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap5)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap6)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap7)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap8)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap9)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap10)
            $sqlBulkCopy.ColumnMappings.Add($ColumnMap11)
            #>
            $sqlText = "delete from [dbo].[Clean_Assets]"

            $conn.Open()
            $cmd = New-Object Data.SqlClient.SqlCommand $sqlText, $conn;            
            $cmd.ExecuteNonQuery(); 

            $sqlBulkCopy.WriteToServer($dataTable)
            $conn.Close()

            $cmd.Dispose()
            $conn.Dispose()
           
        }
        Catch {
            Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
            Break
        }
    }

    End {
        If ($?) {
            Log-Write -LogPath $sLogFile -LineValue "ExportJsonDataToSQL Completed Successfully."
            Log-Write -LogPath $sLogFile -LineValue " "
        }
    }

}


function Get-Type 
{ 
    param($type) 
 
$types = @( 
'System.Boolean', 
'System.Byte[]', 
'System.Byte', 
'System.Char', 
'System.Datetime', 
'System.Decimal', 
'System.Double', 
'System.Guid', 
'System.Int16', 
'System.Int32', 
'System.Int64', 
'System.Single', 
'System.UInt16', 
'System.UInt32', 
'System.UInt64') 
 
    if ( $types -contains $type ) { 
        Write-Output "$type" 
    } 
    else { 
        Write-Output 'System.String' 
         
    } 
} 
 

function Out-DataTable 
{ 
    [CmdletBinding()] 
    param([Parameter(Position=0, Mandatory=$true, ValueFromPipeline = $true)] [PSObject[]]$InputObject) 
 
    Begin 
    { 
        $dt = new-object Data.datatable   
        $First = $true  
    } 
    Process 
    { 
        foreach ($object in $InputObject) 
        { 
            $DR = $DT.NewRow()   
            foreach($property in $object.PsObject.get_properties()) 
            {   
                if ($first) 
                {   
                    $Col =  new-object Data.DataColumn   
                    $Col.ColumnName = $property.Name.ToString()   
                    if ($property.value) 
                    { 
                        if ($property.value -isnot [System.DBNull]) { 
                            $Col.DataType = [System.Type]::GetType("$(Get-Type $property.TypeNameOfValue)") 
                         } 
                    } 
                    $DT.Columns.Add($Col) 
                }   
                if ($property.Gettype().IsArray) { 
                    $DR.Item($property.Name) =$property.value | ConvertTo-XML -AS String -NoTypeInformation -Depth 1 
                }   
               else { 
                    $DR.Item($property.Name) = $property.value 
                } 
            }   
            $DT.Rows.Add($DR)   
            $First = $false 
        } 
    }  
      
    End 
    { 
        Write-Output @(,($dt)) 
    } 
 
} 

Function Log-Start {
  
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $true)][string]$LogPath, [Parameter(Mandatory = $true)][string]$LogName, [Parameter(Mandatory = $true)][string]$ScriptVersion)
    
    Process {
        $sFullPath = $LogPath + "\" + $LogName
      
        #Check if file exists and delete if it does
    <#
      If((Test-Path -Path $sFullPath)){
        Remove-Item -Path $sFullPath -Force
      }
    #>
        #Create file and start logging
        New-Item -Path $LogPath -Value $LogName -ItemType File
      
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value "Started processing at [$([DateTime]::Now)]."
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value ""
        Add-Content -Path $sFullPath -Value "Running script version [$ScriptVersion]."
        Add-Content -Path $sFullPath -Value ""
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value ""
    
        #Write to screen for debug mode
        Write-Debug "***************************************************************************************************"
        Write-Debug "Started processing at [$([DateTime]::Now)]."
        Write-Debug "***************************************************************************************************"
        Write-Debug ""
        Write-Debug "Running script version [$ScriptVersion]."
        Write-Debug ""
        Write-Debug "***************************************************************************************************"
        Write-Debug ""
    }
}
  
Function Log-Write {
    
    [CmdletBinding()] 
    Param ([Parameter(Mandatory = $true)][string]$LogPath, [Parameter(Mandatory = $true)][string]$LineValue)
    
    Process {
        Add-Content -Path $LogPath -Value $LineValue
    
        #Write to screen for debug mode
        Write-Debug $LineValue
    }
}
  
Function Log-Error {
    
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $true)][string]$LogPath, [Parameter(Mandatory = $true)][string]$ErrorDesc, [Parameter(Mandatory = $true)][boolean]$ExitGracefully)
    
    Process {
        Add-Content -Path $LogPath -Value "Error: An error has occurred [$ErrorDesc]."
    
        #Write to screen for debug mode
        Write-Debug "Error: An error has occurred [$ErrorDesc]."
      
        #If $ExitGracefully = True then run Log-Finish and exit script
        If ($ExitGracefully -eq $True) {
            Log-Finish -LogPath $LogPath
            Break
        }
    }
}
  
Function Log-Finish {
    
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $true)][string]$LogPath, [Parameter(Mandatory = $false)][string]$NoExit)
    
    Process {
        Add-Content -Path $LogPath -Value ""
        Add-Content -Path $LogPath -Value "***************************************************************************************************"
        Add-Content -Path $LogPath -Value "Finished processing at [$([DateTime]::Now)]."
        Add-Content -Path $LogPath -Value "***************************************************************************************************"
    
        #Write to screen for debug mode
        Write-Debug ""
        Write-Debug "***************************************************************************************************"
        Write-Debug "Finished processing at [$([DateTime]::Now)]."
        Write-Debug "***************************************************************************************************"
    
        #Exit calling script if NoExit has not been specified or is set to False
        If (!($NoExit) -or ($NoExit -eq $False)) {
            Exit
        }    
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
JsonData
Log-Finish -LogPath $sLogFile





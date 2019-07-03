<###
Created by: Cirino Carvalho
Date: 06/19/2019
Purpose: Create one tiff from pdf with 2 pages
###>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#-----------------------------------------------------------[Declarations]----------------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#MasterFolder
$pathMaster = "C:\!Tools\ghostscript"
#Ghostscript path exe
$tool = 'C:\Program Files\gs\gs9.27\bin\gswin64c.exe'

#1-Vertical/0-Horizontal
[bool]$vertical = 0

#Source of pdf's to be converted
$pdfSource = "$pathMaster\"

#Dest of pdf's processed
$pdfDest = "$pathMaster\pdf\"
If(!(test-path $pdfDest))
{
      New-Item -ItemType Directory -Force -Path $pdfDest
}

#Source of tiff's converted from pdf's
$tiffSource = "$pathMaster\"

#Dest of tiff's after appeding
$tiffDest = "$pathMaster\tiff\"
If(!(test-path $tiffDest))
{
      New-Item -ItemType Directory -Force -Path $tiffDest
}

#Create the folder Logs in inside the folder where the script is located
$sLogPath = ("$PSScriptRoot\Logs\")
If(!(test-path $sLogPath))
{
      New-Item -ItemType Directory -Force -Path $sLogPath
}

#Log File Info
$sLogName = ("pdf2tiff_" + (Get-Date -Format 'MMddyyyy') + ".log")
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#-----------------------------------------------------------[Functions]------------------------------------------------------------
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
Function AppendTiff {

    Param($filename)
    
    Begin {
        Log-Write -LogPath $sLogFile -LineValue ("Start Appending the files..." + $filename)
    }
    
    Process {
        Try {
            $firstPath = ($tiffSource + $filename + "1.tif")
            $secondPath = ($tiffSource + $filename + "2.tif")
    
            if ((test-path $firstPath) -and (test-path $secondPath)) {
                [System.Drawing.Image]$firstImage = [System.Drawing.Image]::FromFile($firstPath)
                [System.Drawing.Image]$secondImage = [System.Drawing.Image]::FromFile($secondPath)
    
                [int] $outputImageWidth
                [int] $outputImageHeight
    
                If ($vertical) {
                    $outputImageWidth = If ($firstImage.Width -gt $secondImage.Width) { $firstImage.Width } Else { $secondImage.Width }
                    $outputImageHeight = $firstImage.Height + $secondImage.Height
                }
                Else {
                    $outputImageWidth = $firstImage.Width + $secondImage.Width
                    $outputImageHeight = If ($firstImage.Height -gt $secondImage.Height) { $firstImage.Height } Else { $secondImage.Height }
                }
        
                [System.Drawing.Bitmap] $outputImage = New-Object System.Drawing.Bitmap -Args $outputImageWidth, $outputImageHeight, Format24bppRgb
                [System.Drawing.Graphics] $graphics = $null
        
                $graphics = [System.Drawing.Graphics]::FromImage($outputImage)
    
                If ($vertical) {
                    $graphics.DrawImage($firstImage, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $firstImage.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $firstImage.Size), "Pixel")
                    $graphics.DrawImage($secondImage, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point -Args 0, $firstImage.Height), $secondImage.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $secondImage.Size), "Pixel")
                }
                Else {
                    $graphics.DrawImage($firstImage, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $firstImage.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $firstImage.Size), "Pixel")
                    $graphics.DrawImage($secondImage, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point -Args $firstImage.Width, 0), $secondImage.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $secondImage.Size), "Pixel")
                }
            
                #Write the new image file
                $outputImage.Save("$tiffDest\$filename.tif", "tiff")

                $firstImage.Dispose()
                $secondImage.Dispose()

                #Delete the tiffs temp file
                Remove-Item -Path $firstPath -Force
                Remove-Item -Path $secondPath -Force
            
                #Move the pdf to another folder
                Move-Item -Path ($pdfSource + $filename + ".pdf") -Destination ($pdfDest + $filename + ".pdf") -Force
    
                Log-Write -LogPath $sLogFile -LineValue ("Finishing Appending the files..." + $filename)

                $graphics.Dispose()

            }
            else {
                Log-Write -LogPath $sLogFile -LineValue ("Temp tiff file not exist!")
            }
        }
      
        Catch {
            Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
            Break
        }
    }
    
    End {
        If ($?) {
            Log-Write -LogPath $sLogFile -LineValue "File Appending Completed Successfully."
            Log-Write -LogPath $sLogFile -LineValue " "
        }
    }

}

Function ConvertPDF2Tiff {

    Param()
    
    Begin {
        Log-Write -LogPath $sLogFile -LineValue "Start Converting the Files..."
    }
    
    Process {
        Try {
            $pdfs = get-childitem $pdfSource | Where-Object { $_.Extension -match "pdf" }
            If ($pdfs.Count -gt 0) {
                foreach ($pdf in $pdfs) {
      
                    $tiff = $pdf.FullName.split('.')[0] + '%01d' + '.tif'
                    if (test-path $tiff) {
                        Log-Write -LogPath $sLogFile -LineValue ("tiff file already exists " + $tiff)
                    }
                    else { 
                        Log-Write -LogPath $sLogFile -LineValue ('Start Convert the file ' + $pdf.Name)         
                        $param = "-sOutputFile=$tiff"
                        & $tool -q -dNOPAUSE -sDEVICE=tiff24nc -sCompression=lzw $param -r200  $pdf.FullName -c quit
                        Log-Write -LogPath $sLogFile -LineValue ('File Converted Sucessfull ' + $pdf.Name)  
                        AppendTiff($pdf.Name.split('.')[0])
                    }
                }
            }
            Else {
                Log-Write -LogPath $sLogFile -LineValue "There is no file to be converted..."
            }
        }
      
        Catch {
            Log-Error -LogPath $sLogFile -ErrorDesc $_.Exception -ExitGracefully $True
            Break
        }
    }
    
    End {
        If ($?) {
            Log-Write -LogPath $sLogFile -LineValue "End function for convert completed Successfully."
            Log-Write -LogPath $sLogFile -LineValue " "
        }
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
  
Function Log-Email {
  
    [CmdletBinding()]
    Param ([Parameter(Mandatory = $true)][string]$LogPath, [Parameter(Mandatory = $true)][string]$EmailFrom, [Parameter(Mandatory = $true)][string]$EmailTo, [Parameter(Mandatory = $true)][string]$EmailSubject)
    
    Process {
        Try {
            $sBody = (Get-Content $LogPath | out-string)
        
            #Create SMTP object and send email
            $sSmtpServer = "smtp.yourserver"
            $oSmtp = new-object Net.Mail.SmtpClient($sSmtpServer)
            $oSmtp.Send($EmailFrom, $EmailTo, $EmailSubject, $sBody)
            Exit 0
        }
      
        Catch {
            Exit 1
        } 
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
ConvertPDF2Tiff
Log-Finish -LogPath $sLogFile
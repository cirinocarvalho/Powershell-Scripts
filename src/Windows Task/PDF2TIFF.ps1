<###
Created by: Cirino Carvalho
Date: 06/19/2019
Purpose: Create one tiff from pdf
###>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#-----------------------------------------------------------[Declarations]----------------------------------------------------------------

#Script Version
$sScriptVersion = "1.1"

#Get-PSDrive B | Remove-PSDrive
#New-PSDrive –Name "B" –PSProvider FileSystem –Root "\\file\folder\"

#MasterFolder
$pathMaster = "C:\!Tools\ghostscript\"
#Ghostscript path exe
$tool = 'C:\Program Files\gs\gs9.27\bin\gswin64c.exe'

#1-Vertical/0-Horizontal
[bool]$vertical = 0

#Source of pdf's to be converted
$pdfSource = "$pathMaster"

#Dest of pdf's processed
$pdfDest = "$pathMaster\pdf\"
If(!(test-path $pdfDest))
{
      New-Item -ItemType Directory -Force -Path $pdfDest
}

#Source of tiff's converted from pdf's
$tiffSource = $pdfSource

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
                $filepath = ($tiffSource + $filename + "*.tif")

                $files = Get-ChildItem $filepath

                If ($files.Count -gt 0)
                {
                    foreach ($file in $files) 
                    {

                        $Index = If ($files.Count -gt 1) {($files.IndexOf($file)+1)} Else { 1 }
                        
                        [System.Drawing.Image]$imageTemp = [System.Drawing.Image]::FromFile($file.Fullname)
                        New-Variable -Name "image${Index}" -Value $imageTemp -Force

                    }

                    [int] $outputImageWidth = 0
                    [int] $outputImageHeight = 0

                
                    If ($vertical)
                    {
                        for ($i=1; $i -le $files.Count; $i++)
                            {

                               $temp = Get-Variable -Name "image$i" -ValueOnly
                               $outputImageWidth = If ($outputImageWidth.Width -gt $temp.Width) { $outputImageWidth.Width } Else { $temp.Width }
                               $outputImageHeight += $temp.Height
                            }
                    }
                    Else
                    {
                        for ($i=1; $i -le $files.Count; $i++)
                            {

                               $temp = Get-Variable -Name "image$i" -ValueOnly
                               $outputImageWidth += $temp.Width
                               $outputImageHeight =  If ($outputImageWidth.Height -gt $temp.Height) { $outputImageWidth.Height } Else { $temp.Height }
                            }
                    }

                    [System.Drawing.Bitmap] $outputImage = New-Object System.Drawing.Bitmap -Args $outputImageWidth, $outputImageHeight, Format24bppRgb
                    [System.Drawing.Graphics] $graphics = $null

                    $graphics = [System.Drawing.Graphics]::FromImage($outputImage)

                    If ($vertical)
                    {
                        $graphics.DrawImage($image1, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $image1.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $image1.Size), "Pixel")
                        $size = 0

                        If($files.Count -gt 1)
                        {
                            for ($i=2; $i -le $files.Count; $i++)
                                {
                                    $j = $i - 1
                                    $size += (Get-Variable -Name "image$j" -ValueOnly).Height
                                    $graphics.DrawImage((Get-Variable -Name "image$i" -ValueOnly), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point -Args 0, $size), (Get-Variable -Name "image$i" -ValueOnly).Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), (Get-Variable -Name "image$i" -ValueOnly).Size), "Pixel")

                                }
                        }
    
                    }
                    Else
                    {
                        $graphics.DrawImage($image1, (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $image1.Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), $image1.Size), "Pixel")
                        $size = 0

                        If($files.Count -gt 1)
                        {
                            for ($i=2; $i -le $files.Count; $i++)
                                {
                                    $j = $i - 1
                                    $size += (Get-Variable -Name "image$j" -ValueOnly).Width
                                    $graphics.DrawImage((Get-Variable -Name "image$i" -ValueOnly), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point -Args $size, 0), (Get-Variable -Name "image$i" -ValueOnly).Size), (New-Object System.Drawing.Rectangle -Args (New-Object System.Drawing.Point), (Get-Variable -Name "image$i" -ValueOnly).Size), "Pixel")

                                }
                        }
                    }

                    #Write the new image file
                    $lastFile = "$tiffDest\$filename.tif"
                    
                    If(test-path $lastFile)
                    {
                        Remove-Item -Path $lastFile -Force
                    }

                    $outputImage.Save($lastFile, "tiff")

                    $temp.Dispose()
                    $imageTemp.Dispose()

                    for ($i=1; $i -le $files.Count; $i++)
                        {

                           (Get-Variable -Name "image$i" -ValueOnly).Dispose()

                        }

                    #Delete the tiffs temp file
                    foreach ($file in $files) {
 
                        Remove-Item -Path $file.FullName -Force

                    } 

                    #Move the pdf to another folder
                    Move-Item -Path ($pdfSource + $filename + ".pdf") -Destination ($pdfDest + $filename + ".pdf") -Force
    
                    Log-Write -LogPath $sLogFile -LineValue ("Finishing Appending the files..." + $filename)

                    $graphics.Dispose()
                
                }
                

            Else {
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

                        Write-Output($pdf.FullName)
                        Log-Write -LogPath $sLogFile -LineValue ('Start Convert the file ' + $pdf.Name)         
                        $param = "-sOutputFile=$tiff"
                        & $tool -q `
                        -dNOPAUSE `
                        -sDEVICE=tiff24nc `
                        -sCompression=lzw $param `
                        -r200  `
                        $pdf.FullName `
                        -c quit `

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

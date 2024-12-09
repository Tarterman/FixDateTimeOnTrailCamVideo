using namespace System.Drawing
function Set-OCRDetectedDateTime {
    <#
    .SYNOPSIS
        Uses OCR to read the timestamp on trail cam video and set the CreationDate and LastWriteTime to that value
    .DESCRIPTION
        When copying videos off of some (maybe all trail cameras) to a phone, it sets the CreationDate and LastWriteTime of
        the video to the date/time they were copied. The videos have a timestamp at the bottom of the video. This script
        uses ffmpeg to grab the first frame of video and save it. It then crops the image to just get the bar at the bottom
        with the timestamp.

        It then uses Tesseract OCR to read the date/time off of the cropped image and sets the CreationDate and LastWriteTime
        to the value detected.
    .NOTES
        This script has been tested on Wosoda G300 and Vikeri A1 trail cameras.

        This script requires .NET 4 and is not supported in Linux.
    .EXAMPLE
        Set-OCRDetectedDateTime -CamerasPlacedDate "2024/10/23" -CamerasCheckedDate "2024/12/01" -VideoFolder "C:\Videos" -ffmpegExe "C:\ffmpeg\bin\ffmpeg.exe"

        Uses Tesseract OCR (using the default installed path) to analyze video files in C:\Videos and set the detected date/time on the file
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [datetime]$CamerasPlacedDate,
        [Parameter(Mandatory)]
        [datetime]$CamerasCheckedDate,
        [Parameter(Mandatory)]
        [string]$VideoFolder,
        [Parameter(Mandatory)]
        [string]$ffmpegExe,
        [Parameter(Mandatory=$false)]
        [string]$Tesseract = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    )
    
    begin {
        $firstFramePath = "$($env:TEMP)\firstframe.jpg"
        $croppedImagePath = "$($env:TEMP)\croppedimage.jpg"
        $hiddenConsoleOutput = "$($env:TEMP)\hiddenConsoleOutput.txt"
    }
    
    process {
        $currentFile = 1
        $files = Get-ChildItem -Path $VideoFolder | Where-Object {$_.extension -in ".mp4",".avi",".mkv"}
        foreach ($file in $files) {
            Write-Progress -Activity "Analyzing video $($currentFile) of $($files.Count)." -PercentComplete (($currentFile / $files.Count) * 100)
            if (Test-Path $firstFramePath) {
                Remove-Item $firstFramePath
            }
        
            if (Test-Path $croppedImagePath) {
                Remove-Item $croppedImagePath
            }
        
            # Get first frame of video and save as image
            & $ffmpegExe -i $file.FullName -vf "select=eq(n\,0)" -q:v 3 -frames:v 1 $firstFramePath > $hiddenConsoleOutput 2>&1
        
            # Crop image to just get the bottom where the time and date are
            $image = [Bitmap]::new($firstFramePath)
            $cloneRect = [Rectangle]::new(0,$image.Height * .95,$image.Width,$image.Height * .05)
            $format = $image.PixelFormat
            $croppedImage = $image.Clone($cloneRect, $format)
            $croppedImage.Save($croppedImagePath)
            
            $ocr = (& $Tesseract --dpi $croppedImage.HorizontalResolution $croppedImagePath stdout -c textord_heavy_nr=1).Split(" ")
        
            # Set created time and last modified time to timestamp on image
            $ocr | ForEach-Object {
                $item = $_
                if ($item -match '\d{4}/\d{2}/\d{2}') {
                    $date = $item
                } elseif ($item -match '\d{2}:\d{2}:\d{2}') {
                    $time = $item
                }
            }
        
            # Set the CreationTime and LastWriteTime based off of OCR detected time
            if (-not [string]::IsNullOrEmpty($date) -and -not [string]::IsNullOrEmpty($time)) {
                try {
                    $dateTime = [datetime]("{0} {1}" -f $date, $time)
        
                    # Adjust for DST if necessary
                    $dstInfo = Get-DSTInfo -Year ([datetime]$date).Year
        
                    if ($datetime -gt $dstInfo.EndDate -and $CamerasPlacedDate -lt $dstInfo.EndDate -and $CamerasCheckedDate -gt $dstInfo.EndDate) {
                        $file.CreationTime = $dateTime.AddHours(-1)
                        $file.LastWriteTime = $dateTime.AddHours(-1)
                    } elseif ($datetime -gt $dstInfo.BeginDate -and $CamerasPlacedDate -lt $dstInfo.BeginDate -and $CamerasCheckedDate -gt $dstInfo.BeginDate) {
                        $file.CreationTime = $dateTime.AddHours(+1)
                        $file.LastWriteTime = $dateTime.AddHours(+1)
                    } else {
                        $file.CreationTime = $dateTime
                        $file.LastWriteTime = $dateTime
                    }
                } catch [System.Management.Automation.RuntimeException] {
                    # Sometimes the date/time is detected incorrectly
                    Write-Warning "The date/time $($date) $($time) is not valid. The file $($file.Name) was skipped."
                }
            }
        
            # Cleanup
            $croppedImage.Dispose()
            $image.Dispose()
            Remove-Item $firstFramePath, $croppedImagePath
            $currentFile++
        }
        Write-Progress -Activity "Analyzing video $($currentFile) of $($files.Count)." -Completed
        Remove-Item -Path $hiddenConsoleOutput
    }
    
    end {
        
    }
}

function Get-DSTInfo {
    param(
        [int]$Year = (Get-Date).Year
    )

    $beginDate = [datetime]"March 1, $Year"
    while ($beginDate.DayOfWeek -ne 'Sunday') {
        $beginDate = $beginDate.AddDays(1)
    }

    $endDate = [datetime]"November 1, $Year"
    while ($endDate.DayOfWeek -ne 'Sunday') {
        $endDate = $endDate.AddDays(1)
    }

    [PSCustomObject]@{
        Year = $Year
        BeginDate = $($beginDate.AddDays(7).AddHours(2))
        EndDate = $($endDate.AddHours(2))
    }
}
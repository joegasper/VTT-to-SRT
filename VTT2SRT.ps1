<#
.SYNOPSIS
Fast conversion of Microsoft Stream VTT subtitle file to SRT format.
.DESCRIPTION
Uses select-string instead of get-content to improve speed 2 magnitudes.
.NOTES
Inspiration from: https://gist.github.com/brianonn/455bce106bd86c9587d223acfbbe9751/
Takes a 3.5 minute process for 17K row VTT to 3.5 seconds

This script is specifically for generating an EXE of the script for easier end-user usage.

Download Win-PS2EXE
https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5
https://github.com/MScholtes/TechNet-Gallery

Run Win-PS2EXE.exe
Source: path to this file
Uncheck the box for "Compile a graphical windows program"
Click the Compile button.

Now you can drag-n-drop .vtt|.txt files onto the EXE and it will convert the files,
placing them in the same folder as the VTT, but with an SRT file extension.

https://github.com/joegasper
#>

foreach ($File in $Args) {

    $Item += 1
    Write-Output "Processing $Item of $($Args.Count) --> $($File)"

    $Lines = @()

    $OutFile = $File -replace '(\.vtt$|\.txt$)', '.srt'
    if ( $OutFile.split('.')[-1] -ne 'srt' ) {
        $OutFile = $OutFile + '.srt'
    }

    New-Item -Path $OutFile -ItemType File -Force | Out-Null
    $Subtitles = Select-String -Path $File -Pattern '(^|\s)(\d\d):(\d\d):(\d\d)\.(\d{1,3})' -Context 0, 2

    for ($i = 0; $i -lt $Subtitles.count; $i++) {
        $Lines += $i + 1
        $Lines += $Subtitles[$i].Line -replace '\.', ','
        $Lines += $Subtitles[$i].Context.DisplayPostContext
        $Lines += ''
    }
    $Lines | Out-File -FilePath $OutFile -Append -Force
}

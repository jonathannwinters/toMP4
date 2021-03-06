<#
.SYNOPSIS
  toMP4 converts existing video files to MP4 and recreates same
  folder structure as source in chosen destination folder.
.DESCRIPTION
  This script recursively navigates a given folder
  to find video files (*.avi,*.mkv,*.ogm,*.wmv) and converts 
  them to .mp4  placing the newly  generated .mp4 files into 
  the given destination folder with the same folder strucure.
.NOTES
  File Name : toMP4.ps1
  Author  :  Jonathan N. Winters, jonathan@winters.im
  Requires  : Powershell V2, Handbrake CLI <https://handbrake.fr/downloads2.php>
  Creation Date : 20170531
  Version : 0.3
  Usage : powershell .\toMP4.ps1 
.LINK

#>


# Test to see if HandbrakeCLI is installed
# locate Handbrake CLI exe path 
$handbrakeCLIpath = (Get-ItemProperty -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Handbrake" -erroraction silentlycontinue).'(Default)' -replace '.exe','cli.exe'

# if not installed, open browser to Handbrake's page     
# statement credit, Pat Richard,  http://www.ehloworld.com/643  
if ($handbrakeCLIpath -eq $null){
    Write-Host "Handbrake not found on this system. Please install Handbrake and try again." -foregroundcolor red;
    # Speak -phrase "Handbrake not found on this system. Please install Handbrake and try again."
    $ie = new-object -comobject "InternetExplorer.Application"
    $ie.visible = $true
    $ie.navigate("https://handbrake.fr/downloads2.php")
    exit
}

# Read-FolderBrowserDialog function credit, Daniel Schroeder
# http://blog.danskingdom.com/powershell-multi-line-input-box-dialog-open-file-dialog-folder-browser-dialog-input-box-and-message-box/
function Read-FolderBrowserDialog([string]$Message, [string]$InitialDirectory, [switch]$NoNewFolderButton){
    $browseForFolderOptions = 0
    if ($NoNewFolderButton) { $browseForFolderOptions += 512 }
 
    $app = New-Object -ComObject Shell.Application
    $folder = $app.BrowseForFolder(0, $Message, $browseForFolderOptions, $InitialDirectory)
    if ($folder) { $selectedDirectory = $folder.Self.Path } else { $selectedDirectory = '' }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) > $null
    return $selectedDirectory
}



#get root folder path
$sourceFolder = Read-FolderBrowserDialog -Message "Please select a source folder" -InitialDirectory '%USERPROFILE%\Desktop' -NoNewFolderButton 
Write-Host "Source folder      ="  $sourceFolder



#get root folder path
$destFolder =  Read-FolderBrowserDialog -Message "Please select a destination folder" -InitialDirectory '%USERPROFILE%\Desktop' -NoNewFolderButton 
Write-Host "Destination folder ="  $destFolder



#only run if both a source and destination folder have been selected
if(    (![string]::IsNullOrEmpty($sourceFolder)) -and  (![string]::IsNullOrEmpty($destFolder))  ) {
  
    #fetch all video file objects recursively
    $fileList = Get-ChildItem $sourceFolder -include *.avi,*.mkv,*.ogm,*.wmv -recurse
    
    $fileCount = $fileList.count
    $i = 0;
    ForEach ($file in $fileList){
        $currentSubFolder = $file.DirectoryName.Substring($sourceFolder.length)
        $i++;
        $oldFile = $file.DirectoryName + "\" + $file.BaseName + $file.Extension;
        
        #provide user visual feedback as to status
        Write-Host "********************************************************************************"
        Write-Host "toMP4 using HandbrakeCLI is converting file $i of $fileCount." 
        Write-Host "Processing Source      - $oldfile"
        
        #create destination folder path if it does not exist
        $destPath = $destFolder + $currentSubFolder
        New-Item -ItemType Directory -Force -Path $destPath | Out-Null
        
        #define new file path and name
        $newFile =  $destPath + "\" + $file.BaseName + ".mp4";
        
        Write-Host "Processing Destination - $newfile"
        Write-Host "Conversion in progress ..."
        Write-Host "********************************************************************************"
        
        #perform the conversion
        Start-Process $handbrakeCLIpath -ArgumentList "-i `"$oldFile`" -t 1 --angle 1 -c 1 -o `"$newFile`" -e x264 -b 1500 -a 1 -E faac -B 160 -R Auto -6 dpl2 -f mp4 -m -2 -T -x ref=2:bframes=2:me=umh -n eng"  -Wait -NoNewWindow
  
        Write-Host "Thank you for using toMP4 by Jonathan N. Winters, http://Jonathan.Winters.im"
    }
}
else {
    Write-Host "Both a source and destination folder must be selected. Try again."
}

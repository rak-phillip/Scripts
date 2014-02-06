Param([string]$folderPath, [string]$destination)

#$folderPath = "\\Public Folders\All Public Folders\Co:Training Records - (Unprocessed)"
#$destination = "D:\test\"

.{.\CopyAllAttachments $folderPath $destination}

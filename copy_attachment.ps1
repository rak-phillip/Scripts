Function GetFolder([string]$folderPath, $session)
{
    $backslash = "\"
    
    if ($folderPath.StartsWith("\\")) {
        $folderPath = $folderPath.Remove(0, 2)
    }
    
    $folders = $folderPath.Split($backslash.ToCharArray())
    $folder = $session.Session.Folders.item($folders[0])
    
    if ($folder -ne $null) {
        for ($i = 1; $i -le $folders.GetUpperBound(0); $i++) {
            $subFolders = $folder.Folders
            $folder = $subFolders.item($folders[$i])
            
            if ($folder -eq $null) {
                return $null
            }            
        }
    }
    return $folder    
}

Function CopyAllAttachments([string]$folderPath)
{
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"

    $session = New-Object -comobject Outlook.Application

    $session.Version

    if (!($session.Version -like "12.*" -or $session.Version -like "14.*")){
        write-host "Requires 2007 or 2010"
        return
    }

    $trainingRecords = GetFolder $folderPath $session

    $folderItems = $trainingRecords.Items
    $currentMail = $null

    foreach ($collectionitem in $folderItems) {
        $currentMail = $collectionitem
        if ($currentMail -ne $null) {
            if ($currentMail.Attachments.Count -gt 0) {
                for ($i = 1; $i -le $currentMail.Attachments.Count; $i++) {
                    $currentMail.Attachments.Item($i).SaveAsFile("D:\test\" + $currentMail.Attachments.Item($i).FileName);
                }
            }
        }
    }
}

$folderPath = "\\Public Folders\All Public Folders\Co:Training Records - (Unprocessed)"

CopyAllAttachments $folderPath
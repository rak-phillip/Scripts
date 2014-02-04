Add-Type -assembly "Microsoft.Office.Interop.Outlook"

$folderPath = "\\Public Folders\All Public Folders\Co:Training Records - (Unprocessed)"
$session = New-Object -comobject Outlook.Application

$session.Version
$session.Session.Folders | ft Name

if (!($session.Version -like "12.*" -or $session.Version -like "14.*")){
    write-host "Requires 2007 or 2010"
    return
}

$backslash = "\"

if ($folderPath.StartsWith("\\")){
    $folderPath = $folderPath.Remove(0, 2)
}

$folders = $folderPath.Split($backslash.ToCharArray())
$folder = $session.Session.Folders
$name = "Public Folders"

foreach ($item in $folder) {
    if ($item.Name -eq $name -or $item.DisplayName -eq $name) {
        $publicFolder = $item
    }
}

$publicFolder.Folders | ft Name
$folder = $publicFolder.Folders
$name = "All Public Folders"

foreach ($item in $folder) {
    if ($item.Name -eq $name -or $item.DisplayName -eq $name) {
        $allPublicFolders = $item
    }
}

$allPublicFolders.Folders | ft Name
$folder = $allPublicFolders.Folders
$name = "Co:Training Records - (Unprocessed)"

foreach ($item in $folder) {
    if ($item.Name -eq $name -or $item.DisplayName -eq $name) {
        $trainingRecords = $item
    }
}

$trainingRecords.Folders | ft Name
$folderItems = $trainingRecords.Items
$folderItems.Count
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

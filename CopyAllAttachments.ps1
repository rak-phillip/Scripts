Param ([string]$folderPath, [string]$destination)

Function CopyAllAttachments([string]$folderPath, [string]$destination)
{
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"

    $session = New-Object -comobject Outlook.Application

    $session.Version

    if (!($session.Version -like "12.*" -or $session.Version -like "14.*")){
        write-host "Requires 2007 or 2010"
        return
    }

    $trainingRecords = .{.\GetFolderByPath $folderPath $session}

    $folderItems = $trainingRecords.Items
    $currentMail = $null

    foreach ($collectionitem in $folderItems) {
        $currentMail = $collectionitem
        if ($currentMail -ne $null) {
            if ($currentMail.Attachments.Count -gt 0) {
                for ($i = 1; $i -le $currentMail.Attachments.Count; $i++) {
                    $currentMail.Attachments.Item($i).SaveAsFile($destination + $currentMail.Attachments.Item($i).FileName);
                }
            }
        }
    }
}

CopyAllAttachments $folderPath $destination

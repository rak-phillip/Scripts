Param($outlookFolder, [string]$destination)

Function CopyAllAttachments($outlookFolder, [string]$destination)
{
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"

    $folderItems = $outlookFolder.Items
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

CopyAllAttachments $outlookFolder $destination

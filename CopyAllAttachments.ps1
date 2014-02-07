#------------------------------------------------------------------------------
#      ####            ####
#   ############   ############      
# ################################    Author: Phillip Rak
#|  ##############  ##############    Description: Copy all attachments in an
#|      ##########      ##########      Outlook public folder to a directory on
#|        ########        ########      the filesystem.
#|        ########        ########    
#|        ########.       ########   
#|        ######    * .   ########   
#|        ##.           ##########       
#|            * .   ##############                               
#|                ################   
#|                ################   
#|                ################       
#|        ##      ######  ########   
#|        ######  ##      ########       
#|        ########        ########   
#|        ########        ########       
# .       ########.       ########   
#    .    ######     .    ###### 
#       . ##            . ##
#------------------------------------------------------------------------------

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
                    $filePath = $destination + $currentMail.Attachments.Item($i).FileName
                    if (Test-Path ($filePath)) {
                        write-host "The file" $currentMail.Attachments.Item($i).FileName "already exists"
                    }
                    else {
                        write-host "Writing file" $currentMail.Attachments.Item($i).FileName "to" $destination
                        $currentMail.Attachments.Item($i).SaveAsFile($filePath);
                    }
                }
            }
        }
    }
}

CopyAllAttachments $outlookFolder $destination

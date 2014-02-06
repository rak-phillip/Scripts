#------------------------------------------------------------------------------
#      ####            ####
#   ############   ############      
# ################################    Author: Phillip Rak
#|  ##############  ##############    Description: 
#|      ##########      ##########   
#|        ########        ########   
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

Param([string]$folderPath, [string]$destination, $session=0)

#$folderPath = "\\Public Folders\All Public Folders\Co:Training Records - (Unprocessed)"
#$destination = "D:\test\"

if($session -eq 0){
    #Start a new Outlook Session
    $session = New-Object -comobject Outlook.Application

    $session.Version

    if (!($session.Version -like "12.*" -or $session.Version -like "14.*")){
        write-host "Requires 2007 or 2010"
        return
    }
}

#Get the folder that has items with attachments
$outlookFolder = .{.\GetFolderByPath.ps1 $folderPath $session}

#Copy all attachments from all items that contain attachments to a directory
.{.\CopyAllAttachments.ps1 $outlookFolder $destination}

#------------------------------------------------------------------------------
#      ####            ####
#   ############   ############      
# ################################    Author: Phillip Rak
#|  ##############  ##############    Description: Get a public foder from 
#|      ##########      ##########      Outlook that matches the path entered.
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

#*[string] outlookFolderPath the path to an Outlook folder 
# e.g. \\Public Folders\Corporate Mail
# [object] session the outlook application object
# returns folder object on success
Param([string]$outlookFolderPath, $session=0)

Function GetFolderByPath([string]$outlookFolderPath, $session)
{
    $backslash = "\"
    
    if ($outlookFolderPath.StartsWith("\\")) {
        $outlookFolderPath = $outlookFolderPath.Remove(0, 2)
    }

    $folders = $outlookFolderPath.Split($backslash.ToCharArray())
    $folder = $session.Session.Folders.item($folders[0])
    
    if ($folder -ne $null) {
        for ($i = 1; $i -le $folders.GetUpperBound(0); $i++) {
            if (!($folders[$i] -eq "")) {
                $subFolders = $folder.Folders
                $folder = $subFolders.item($folders[$i])

                if ($folder -eq $null) {
                    return $null
                }            
            }
        }
    }
    return $folder    
}

#if a session was not passed, start a new Outlook session
if($session -eq 0){
    #Start a new Outlook Session
    $session = New-Object -comobject Outlook.Application

    $version = $session.Version

    if (!($version.Substring(0, 2) -ge 12)) {
        write-host "Requires Outlook Version 2007 or greater"
        return
    }
}

GetFolderByPath $outlookFolderPath $session

#------------------------------------------------------------------------------
#      ####            ####
#   ############   ############      
# ################################    Author: Phillip Rak
#|  ##############  ##############    Description: A script that searches for
#|      ##########      ##########      a public folder in Outlook and writes
#|        ########        ########      all of the attachments associated with
#|        ########        ########      emails in the folder to a directory.
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
# [string] destination the directory for all attachments to be copied to
Param([string]$outlookFolderPath="", [string]$destination="", $session=0)

#get the directory of the running script
function GetScriptDirectory
{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

#if the outlookFolderPath was not entered, cancel the script
if($outlookFolderPath -eq "") {
    write-host "A Folder Path is required"
    return
}

#if the destination was not entered, cancel the script
if($destination -eq "") {
    write-host "A destination is required"
    return
}

#if the session is not assigned, create a new Outlook session
if($session -eq 0){
    #Start a new Outlook Session
    $session = New-Object -comobject Outlook.Application

    $session.Version

    if (!($session.Version -like "12.*" -or $session.Version -like "14.*")){
        write-host "Requires 2007 or 2010"
        return
    }
}

#Get the folder that has items with attachments by running the GetFolderByPath
#script
$subScriptName = "GetFolderByPath.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptName

if (Test-Path $subScriptPath)
{
    # use file from local folder
    $outlookFolder = . $subScriptPath $outlookFolderPath $session
}
else
{
    # use central file (via PATH-Variable)
    $outlookFolder = . $subScriptName $outlookFolderPath $session
}

#Copy all attachments from all items that contain attachments to a directory by
#running the CopyAllAttachments script
$subScriptName = "CopyAllAttachments.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptName

if (Test-Path $subScriptPath)
{
    # use file from local folder
    . $subScriptPath $outlookFolder $destination
}
else
{
    # use central file (via PATH-Variable)
    . $subScriptName $outlookFolder $destination
}

write-host "Copied attachments to $destination"

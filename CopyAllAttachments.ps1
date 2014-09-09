#* [string] Account - the email account with permissions to access the desired
#  public folder
#  [string] Displayname - the display name of the desired public folder
#  [integer] FolderCount - the number of folders to search (deafault = 100)
#  [string] WebServicesDllLocation - the location of the Exchange Web Services
#  dll (default = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Param([string]$Account="", $ArchiveDestination="", $Destination="", $FolderCount=0, $FolderId="", [string]$WebServicesDllLocation="")

#if an account was not entered, cancel the script
if($Account -eq "")
{
    write-host "An email address is required. Usage '-Account bjones@acme.com'"
    return
}

if($Destination -eq "")
{
    write-host "`nA destination is required.`n`nUsage: -Destination D:\attachments\`n"
    return
}

if($FolderId -eq "")
{
    write-host "`nA valid folder id is required. Usage '-FolderId `$folderId'`n`nSee script 'GetFolderIdByName.ps1'`n"
    return
}

#use the default web services directory if one is not specified
if($WebServicesDllLocation -eq "")
{
    $WebServicesDllLocation = “C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll”
}

#default to folder count to 100 if no value was entered
if($FolderCount -eq 0){
    $FolderCount = 100
}

Function CopyAllAttachments($FolderId, $Destination, $ArchiveDestination)
{
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $FolderId)
    $view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000)

    $count = 0

    $item = [Microsoft.Exchange.WebServices.Data.Item]

    foreach($item in $rootFolder.FindItems($view))
    {
        $copy = $false
        
        if($item.HasAttachments)
        {
            $item.Load()
            
            $attachment = [Microsoft.Exchange.WebServices.Data.Attachment]
            
            foreach($attachment in $item.Attachments)
            {
                if (($attachment -is [Microsoft.Exchange.WebServices.Data.FileAttachment]) -and ($attachment.Name.EndsWith(".xlsm")))
                {
                    $fileAttachment = [Microsoft.Exchange.WebServices.Data.FileAttachment]
                    $fileAttachment = $attachment
                    $fileAttachment.Load($Destination+"\"+$fileAttachment.Name)
                    
                    write-host $attachment.Name
                    
                    $copy = $true
                }
            }
            
            if(($ArchiveDestination -ne "") -and ($copy -eq $true))
            {
                .\MoveItemToPublicFolder.ps1 $Account $ArchiveDestination $item
            }
        }
    }
}

#import the Exchange Web Services module
Import-Module -Name $WebServicesDllLocation

#set the version of the Exchange Version
$version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

#Instantiate and configure a new service object
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($version)
$service.UseDefaultCredentials = $true
$service.AutodiscoverUrl($Account);

CopyAllAttachments $FolderId $Destination $ArchiveDestination
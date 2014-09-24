#* [string] Account - the email account with permissions to access the desired
#  public folder
#  [string] Displayname - the display name of the desired public folder
#  [integer] FolderCount - the number of folders to search (deafault = 100)
#  [string] WebServicesDllLocation - the location of the Exchange Web Services
#  dll (default = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Param([string]$Account="", [string]$DisplayName="", $FolderCount=0, [string]$WebServicesDllLocation="")

#if an account was not entered, cancel the script
if($Account -eq "")
{
    write-host "An email address is required. Usage '-Account bjones@acme.com'"
    return
}

#if a display was not entered, cancel the script
if($DisplayName -eq "")
{
    write-host "A public folder name is required. Usage '-DisplayName PublicFolderName'"
    return
}

#use the default web services directory if one is not specified
if($WebServicesDllLocation -eq "")
{
    $WebServicesDllLocation = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
}

#default to folder count to 100 if no value was entered
if($FolderCount -eq 0){
    $FolderCount = 100
}

Function GetFolderIdByName([string]$DisplayName, $FolderCount)
{
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]
    $searchFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot
    $folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView($FolderCount)

    $rootFolder = $rootFolder::Bind($service, $searchFolder)

    foreach($folder in $rootFolder.FindFolders($folderView))
    {    
        if ($folder.DisplayName -eq $DisplayName)
        {
            $folderId = $folder.Id
        }
    }

    if($folderId -ne $null)
    {
        write-host "The unique identifier of the" $DisplayName "folder (in the public folder) is: " + $folderId.ToString()
        return $folderId
    }
    else
    {
        write-host "The" $DisplayName "folder was not found in the Inbox folder"
        return
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

GetFolderIdByName $DisplayName $FolderCount


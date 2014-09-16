Param([string]$Account="", [string]$ArchiveDestination="", [string]$Destination="", $DisplayName="", $FolderCount=0, [string]$WebServicesDllLocation="")

#if an account was not entered, cancel the script
if($Account -eq "")
{
    write-host "`nAn email address is required.`n`nUsage: -Account bjones@acme.com`n"
    return
}

#if a display was not entered, cancel the script
if($DisplayName -eq "")
{
    write-host "`nA public folder name is required.`n`nUsage: -DisplayName PublicFolderName`n"
    return
}

#use the default web services directory if one is not specified
if($WebServicesDllLocation -eq "")
{
    $WebServicesDllLocation = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
    #test if the file exists
}

#default to folder count to 100 if no value was entered
if($FolderCount -eq 0){
    $FolderCount = 100
}

#get the directory of the running script
function GetScriptDirectory
{
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

#get the folder id of the public folder
$subScriptName = "GetFolderIdByName.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptname
$currentPath = GetScriptDirectory

if (Test-Path $subScriptPath)
{
    #use the script from the local folder
    $FolderId = . $subscriptPath $Account $DisplayName $FolderCount $WebServicesDllLocation
}
else
{
    #use central file (via PATH-Variable)
    $FolderId = . $subscriptName $Account $DisplayName $FolderCount $WebServicesDllLocation
}

#cancel the script if no folder id is returned
if ($FolderId -eq $null)
{
    write-host "`nUnable to get the public folder id for :" $DisplayName `n
    return $null
}

#copy all item attechments in the public folder and archive the item if copy successful
$subScriptName = "CopyAllAttachments.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptName

$itemArray = @()

if (Test-Path $subScriptPath)
{
    #use the script from the local folder
    $itemArray = . $subscriptPath $Account $Destination $FolderCount $FolderId $WebServicesDllLocation
}
else
{
    #use central file (via PATH-Variable)
    $itemArray = . $subscriptName $Account $Destination $FolderCount $FolderId $WebServicesDllLocation
}

#get the folder id for the Archive Public Folder
$subScriptName = "GetFolderIdByName.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptname
$FolderCount = 100

if (Test-Path $subScriptPath)
{
    #use the script from the local folder
    $archiveId = . $subscriptPath $Account $ArchiveDestination $FolderCount $WebServicesDllLocation
}
else
{
    #use central file (via PATH-Variable)
    $archiveId = . $subscriptName $Account $ArchiveDestination $FolderCount $WebServicesDllLocation
}

#archive all items that had attachments copied
$subScriptName = "MoveItemToPublicFolder.ps1"
$subScriptPath = Join-Path (GetScriptDirectory) $subScriptName

if ($itemArray.count -lt 0)
{
    write-host "No items to move in the public folder" $DisplayName
    return
}

foreach ($item in $itemArray)
{
    if (Test-Path $subScriptPath)
    {
        #use the script from the local folder
        . $subScriptPath $Account $archiveId $item $WebServicesDllLocation
    }
    else
    {
        #use the script from the local folder
        . $subScriptName $Account $archiveId $item $WebServicesDllLocation
    }
}


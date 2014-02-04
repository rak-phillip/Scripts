Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll”

$exchVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

$exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($exchVersion)
$exchService.UseDefaultCredentials = $true
$exchService.AutodiscoverURL("training@hoopercorp.com")

$pfs = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot)

$tinyView = new-object Microsoft.Exchange.WebServices.Data.FolderView(2)
$displayNameProperty = [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName
$filter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($displayNameProperty, "Co:Training Records - (Unprocessed)")

$results = $pfs.FindFolders($filter, $tinyView)
if ($results.TotalCount -gt 1) {
    "Ambiguous Name"
}
elseif ($results.TotalCount -lt 1) {
    "Folder not found"
}
$folder = $results.Folders[0]

$folderClassProperty = [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass
$folderClassValue = $null
$succeeded = $folder.TryGetProperty($folderClassProperty, [ref]$folderClassValue)

if(!($succeeded)) {
    "Couldn't get folder class"
}
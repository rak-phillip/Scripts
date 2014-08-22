Import-Module -Name “C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll”

$version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

#Instantiate a new service object
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($version)
$service.UseDefaultCredentials = $true
$service.AutodiscoverUrl("prak@hoopercorp.com");

$folder = [Microsoft.Exchange.WebServices.Data.Folder]
$rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]
$searchFolder = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot
$folderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
$displayName = "HC:EPD – Xcel Colorado Timesheets (Review)"

$rootFolder = $rootFolder::Bind($service, $searchFolder)

foreach($folder in $rootFolder.FindFolders($folderView))
{
        
    if ($folder.DisplayName -eq $displayName)
    {
        $folderId = $folder.Id
    }
}

if($folderId -ne $null)
{
    write-host "The unique identifier of the 'HC:EPD – Xcel Colorado Timesheets (Review)' folder (in the public folder) is: " + $folderId.ToString()
}
else
{
    write-host "The 'HC:EPD – Xcel Colorado Timesheets (Review)' folder was not found in the Inbox folder"
}

$rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId)
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000)

$count = 0

$item = [Microsoft.Exchange.WebServices.Data.Item]

foreach($item in $rootFolder.FindItems($view))
{
    #write-host $item.Subject
    
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
                $fileAttachment.Load("C:\\HooperIT\\" + $fileAttachment.Name)
                
                write-host $attachment.Name
            }
        }
    }
}
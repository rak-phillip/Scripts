Param([string]$folderPath, $session)

Function GetFolderByPath([string]$folderPath, $session)
{
    $backslash = "\"
    
    if ($folderPath.StartsWith("\\")) {
        $folderPath = $folderPath.Remove(0, 2)
    }
    
    $folders = $folderPath.Split($backslash.ToCharArray())
    $folder = $session.Session.Folders.item($folders[0])
    
    if ($folder -ne $null) {
        for ($i = 1; $i -le $folders.GetUpperBound(0); $i++) {
            $subFolders = $folder.Folders
            $folder = $subFolders.item($folders[$i])
            
            if ($folder -eq $null) {
                return $null
            }            
        }
    }
    return $folder    
}

GetFolderByPath $folderPath $session

# Exchange Web Services Powershell Scripts #

PowerShell scripts that take advantage of Exchange Web Services to manage public
 folder operations.

##Scripts

###copy_attachment.ps1

This script depends on all of the other scripts. Make sure that all scripts are
 in the same folder before running.

A script that looks for all message attachments in an Exchange public folder
 and copies them to a directory. Send all messages to an "Archive" public
 folder if directed.

Example: 

    PS D:\> .\copy_attachment.ps1 -Account gvakarian@csec.com -ArchiveDestination "Council Email Archive" -Destination "D:\destination\directory" -DisplayName "Council Email"
    The unique identifier of the Council Email folder (in the public folder) is:  + W61kSx1LcFYDwkzXKRExdhv3Yy/CYGXNAUpifZLRfgaEFehT09VeQFD3emjqwo8FAA==
    Item Count: 1
    The unique identifier of the Council Email Archive folder (in the public folder) is:  + P2MDv6JF0qg6yPwDGYUU9rJQaQ/R8F98E1bxdgRHGJaC3Yjgthh6vqrYPPbSRRUDAA==
    1


    ToRecipients                 : {}
    BccRecipients                : {}
    ...


###GetFolderIdByName.ps1

A script that finds a public folder by name and returns the public folder id.

Example:

    PS D:\> .\GetFolderIdByName.ps1 -Account gvakarian@csec.com -DisplayName "Council Email"
    The unique identifier of the Council Email folder (in the public folder) is:  + W61kSx1LcFYDwkzXKRExdhv3Yy/CYGXNAUpifZLRfgaEFehT09VeQFD3emjqwo8FAA==

    FolderName          Mailbox             UniqueId            ChangeKey
    ----------          -------             --------            ---------
                                            W61kSx1LcFYDwkzX... P2MDv6JF0qg6yPwD...


###CopyAllAttachments.ps1

A script that copies all attachments in an Exchange public folder. The folder 
is located by ID. Returns an array of items that had attachments copied.

Example:

    PS D:\> $folderId = .\GetFolderIdByName.ps1 -Account gvakarian@csec.com -DisplayName "Council Email"
    The unique identifier of the Council Email folder (in the public folder) is:  + W61kSx1LcFYDwkzXKRExdhv3Yy/CYGXNAUpifZLRfgaEFehT09VeQFD3emjqwo8FAA==
    PS D:\> .\CopyAllAttachments.ps1 -Account gvakarian@csec.com -Destination "D:\destination\directory" -FolderId $folderId


    ToRecipients                 : {jmoreau@normandy.net}
    BccRecipients                : {cshepard@citadel.com}
    CcRecipients                 : {uwrex@urdnot-clan.com}
    ...


###MoveItemToPublicFolder.ps1

A script that moves an array of items from one public folder to another.

Example:

    PS D:\> $folderId = .\GetFolderIdByName.ps1 -Account gvakarian@csec.com -DisplayName "Council Email Archive"
    The unique identifier of the Council Email Archive folder (in the public folder) is:  +  P2MDv6JF0qg6yPwDGYUU9rJQaQ/R8F98E1bxdgRHGJaC3Yjgthh6vqrYPPbSRRUDAA==
    PS D:\> $itemArray = .\CopyAllAttachments.ps1 -Account gvakarian@csec.com -Destination "D:\destination\directory" -FolderId $folderId
    PS D:\> foreach ($item in $itemArray)
    >> {
    >>     .\MoveItemToPublicFolder.ps1 -Account gvakarian@csec.com -DisplayName $folderId -Item $item
    >> }
    >>

    ToRecipients                 : {}
    BccRecipients                : {}
    CcRecipients                 : {}
    ...

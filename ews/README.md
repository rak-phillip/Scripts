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

    PS D:\> .\copy_attachment.ps1 -Account gvakarian@csec.com -ArchiveDisplayName "Council Email Archive" -Destination "D:\destination\directory" -SourceDisplayName "Council Email"
    The unique identifier of the Council Email folder (in the public folder) is:  + W61kSx1LcFYDwkzXKRExdhv3Yy/CYGXNAUpifZLRfgaEFehT09VeQFD3emjqwo8FAA==
    Item Count: 1
    The unique identifier of the Council Email Archive folder (in the public folder) is:  + P2MDv6JF0qg6yPwDGYUU9rJQaQ/R8F98E1bxdgRHGJaC3Yjgthh6vqrYPPbSRRUDAA==
    1


    ToRecipients                 : {}
    BccRecipients                : {}
    ...


###GetFolderIdByName.ps1


###CopyAllAttachments.ps1


###MoveItemToPublicFolder.ps1



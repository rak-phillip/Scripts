# General Purpose Scripts #

This repository contains general purpose scripts that I am able to work on
inbetween major projects. The list of scripts available is incredibly short at 
the moment, but I hope to add more in the future.

##Scripts

###copy_attachment.ps1

CopyAllAttachments.ps1 and GetFolderByPath.ps1 are required to be in the same
directory as this script in order to run successfully.

A setup script that runs both CopyAllAttachments.ps1 and GetFolderByPath.ps1
The main purpost of the script is to check that all parameters were entered and
to create a new Outlook session if one has not been provided to the script.

The script gets the desired Outlook folder by running the script 
GetFolderByPath.ps1 and copies all attachments to a directory by calling
CopyAllAttachments.ps1.

Example:

    PS D:\> $session = new-Object -comobject Outlook.Application
    PS D:\> .\copy_attachment.ps1 -outlookFolderPath "\\Public Folders\All Public Folders\Group Mail" -destination D:\group_attachments\ -session $session
    > Writing file 3_3_2013_12_26_40_22_5140.sig to D:\group_attachments\
    > Copied attachments to D:\group_attachments\

###GetFolderByPath.ps1

This script is able to run independently.

Script that returns an Outlook folder based on the known Outlook folder path.

Example:

    PS D:\> $session = new-Object -comobject Outlook.Application
    PS D:\> $folder = .\GetFolderByPath.ps1 -outlookFolderPath "\\Public Folders\All Public Folders\Group Mail" -session $session

###CopyAllAttachments.ps1

This script is able to run independently.

Script that iterates through every email in an Outlook folder. If an email 
contains attachments, copy them to an existing directory.

Example:
The Outlook Folder was retrieved from GetFolderByPath.ps1

    PS D:\> .\CopyAllAttachments.ps1 -outlookFolder $folder -destination D:\group_attachments\
    > Writing file 3_3_2013_12_26_40_22_5140.sig to D:\group_attachments\

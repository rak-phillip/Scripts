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

###GetFolderBypath.ps1

This script is able to run independently.

Script that returns an Outlook folder based on the known Outlook folder path.

###CopyAllAttachments.ps1

This script is able to run independently.

Script that iterates through every email in an Outlook folder. If an email 
contains attachments, copy them to an existing directory.

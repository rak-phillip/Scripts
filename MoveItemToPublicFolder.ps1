Param([string]$Account="", $DisplayName="", $Item="", [string]$WebServicesDllLocation="")

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

if($Item -eq "")
{
    write-host "An item is required. Usage: -Item `$item"
    return
}

#use the default web services directory if one is not specified
if($WebServicesDllLocation -eq "")
{
    $WebServicesDllLocation = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
}

#import the Exchange Web Services module
Import-Module -Name $WebServicesDllLocation

#set the version of the Exchange Version
$version = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

#Instantiate and configure a new service object
$service = new-object Microsoft.Exchange.WebServices.Data.ExchangeService($version)
$service.UseDefaultCredentials = $true
$service.AutodiscoverUrl($Account);

$Item.Move($DisplayName)
using namespace System.Net
using module ConfigurationHandler
using module SharePointConnectHandler

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$title = $Request.Query.WebAddress
if (-not $title) {
    $title = $Request.Body.WebAddress
}

$fullTitle = $title

if ($title.Length -gt 255)
{
    $title = $title.Substring(0,252) + "..."
}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}

$siteUrl = [ConfigurationHandler]::GetSiteUrl()

$clientId = [ConfigurationHandler]::GetKVSharePointConnectIdName()
$certificate = [Configurationhandler]::GetKVSharePointConnectCertName()

$spch = [SharePointConnectHandler]::new($siteUrl, [ConnectType]::KeyVaultCertificate, $clientId, $certificate, "", "")

<#
$spConnect = "AppExtensionClientId"
$spCertificate = "AppExtensionCertificate"

[SharePointConnectHandler] $spch = [SharePointConnectHandler]::new($siteUrl, [ConnectType]::KeyVaultCertificate, $spConnect, $spCertificate, "", "")
#>

if ($null -ne $spch)
{
    $connection = $spch.GetConnection()

    $item = Add-PnPListItem -List "Tracking" -Values @{ Title = $title; FullTitle = $fullTitle } -Connection $connection
    
    Write-Host "Item added with item id $($item.Id)"

    $body = "Item added with item id $($item.Id)"
}
else
{
    Write-Host "Unable to connect to site $siteUrl"

    $body = "Unable to connect to site $siteUrl"
}

Write-Host "Processed request with Name=$name"

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})

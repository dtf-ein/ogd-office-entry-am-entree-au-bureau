# Delete all AccessRequest list items

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL
)

$ListName = "AccessRequest"

# Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Get all items and delete
$items = Get-PnPListItem -List $ListName -PageSize 1000

Write-Host "Deleting $($items.count) items from $ListName"
foreach ($item in $items)
{
    Write-Host "Removing $($item.Id)"
    Remove-PnPListItem -List $ListName -Identity $item.Id -Force
}

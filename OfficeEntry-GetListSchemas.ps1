#Config Variables

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the lists: https://someurl.ca")]
    [string]$SiteURL
)

$ListNames = @("AccessRequest", "Building", "Floor", "UserSetting", "VisitorLog")
$ListsOutputFile = ".\OfficeEntryListSchemas.xml"

#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin

#Get the List schemas from the Site Templates and export to XML file
$Templates = Get-PnPProvisioningTemplate -OutputInstance -Handlers Lists -ListsToExtract $ListNames
Save-PnPProvisioningTemplate -InputInstance $Templates -Out ($ListsOutputFile)


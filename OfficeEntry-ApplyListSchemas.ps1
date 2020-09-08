Import-Module SharePointPnPPowerShellOnline

#Config Variables

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="Full path to the .xml schema template: .\OfficeEntryListSchemas.xml")]
    [string]$TemplateFile
)

#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL -UseWebLogin

Write-Host "Creating List(s) from Template File..."
Apply-PnPProvisioningTemplate -Path $TemplateFile


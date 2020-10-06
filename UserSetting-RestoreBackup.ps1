# Load an UserSetting.csv that was created by the Power Automate flow
# These backups should be in a /Backups document library on the site

# IMPORTANT!!
# Before running the script, you must remove all .csv colums that aren't part of the Import-Csv call below.

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="CSV file to import")]
    [string]$CSVFile
)

$ListName = "UserSetting"

###!!! CSV Fields need to be hard coded in script below.

#Get the CSV file contents
$CSVData = Import-Csv -Path ($CSVFile) -header "Title","AgreedToPrivacy","AgreedToHealthSafety","ManagerLogin","Modified","Created" -delimiter "," -Encoding UTF8

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
foreach ($Row in $CSVData)
{
    if ($Row.Title -ne "Title")
	{
		Add-PnPListItem -List $ListName -Values @{
            "Title" = $($Row.Title)
            "AgreedToPrivacy" = $($Row.AgreedToPrivacy)
            "AgreedToHealthSafety" = $($Row.AgreedToHealthSafety)
            "ManagerLogin" = $($Row.ManagerLogin)
            "Modified" = $($Row.Modified)
            "Created" = $($Row.Created)}
    }
}


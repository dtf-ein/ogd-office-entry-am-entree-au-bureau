# Load a Building.csv that was created by ./Building-GetListData.ps1

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="Path to the .csv folder: .\")]
    [string]$CSVFolder
)

$ListName = "Building"

###!!! CSV Fields need to be hard coded in script below.
###To Do: investigate using for loop to iterate through columns

#Get the CSV file contents
$CSVData = Import-Csv -Path ($CSVFolder + $ListName + ".csv") -header "Title","BuildingEnglishText","BuildingFrenchText","BuildingShortText","RegionEnglishText","RegionFrenchText","RegionSortOrder","Address","City","TimeZone","TimeZoneOffset" -delimiter "|" -Encoding Unicode

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
foreach ($Row in $CSVData)
{
	if ($Row.Title -ne "Title")
	{

		Add-PnPListItem -List $ListName -Values @{"Title" = $($Row.Title)
                            "BuildingEnglishText" = $($Row.BuildingEnglishText)
                            "BuildingFrenchText" = $($Row.BuildingFrenchText)
                            "BuildingShortText" = $($Row.BuildingShortText)
                            "RegionEnglishText" = $($Row.RegionEnglishText)
                            "RegionFrenchText" = $($Row.RegionFrenchText)
                            "RegionSortOrder" = $($Row.RegionSortOrder)
                            "Address" = $($Row.Address)
                            "City" = $($Row.City)
                            "TimeZone" = $($Row.TimeZone)
                            "TimeZoneOffset" = $($Row.TimeZoneOffset)}
    }
}


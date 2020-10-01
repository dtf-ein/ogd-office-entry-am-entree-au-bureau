# Load a Floor.csv that was created by ./Floor-GetListData.ps1

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="Path to the .csv folder: .\")]
    [string]$CSVFolder
)

$ListName = "Floor"

###!!! CSV Fields need to be hard coded in script below.
###To Do: investigate using for loop to iterate through columns

#Get the CSV file contents
$CSVData = Import-Csv -Path ($CSVFolder + $ListName + ".csv") -header "Title","BuildingID","FloorEnglishText","FloorFrenchText","FloorSortOrder","Capacity","CurrentCapacity","DefaultCapacityPortion","ModifiedCapacityFlag","ChangeReason" -delimiter "|" -Encoding Unicode

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
foreach ($Row in $CSVData)
{
	if ($Row.Title -ne "Title")
	{
		Add-PnPListItem -List $ListName -Values @{"Title" = $($Row.Title)
                            "BuildingID" = $($Row.BuildingID)
                            "FloorEnglishText" = $($Row.FloorEnglishText)
                            "FloorFrenchText" = $($Row.FloorFrenchText)
                            "FloorSortOrder" = $($Row.FloorSortOrder)
                            "Capacity" = $($Row.Capacity)
                            "CurrentCapacity" = $($Row.CurrentCapacity)
                            "DefaultCapacityPortion" = $($Row.DefaultCapacityPortion)
                            "ModifiedCapacityFlag" = $($Row.ModifiedCapacityFlag)
                            "ChangeReason" = $($Row.ChangeReason)}
    }
}


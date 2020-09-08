#Config Variables

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="Path to the .csv folder: .\")]
    [string]$CSVFolder
)

$ListName = "AccessRequest"

###!!! CSV Fields need to be hard coded in script below.
###To Do: investigate using for loop to iterate through columns

#Get the CSV file contents
$CSVData = Import-Csv -Path ($CSVFolder + $ListName + ".csv") -header "Title","StartHour","EndHour","FloorID","Status","VisitorCount","EntryDate","EntryDateID","EmployeeName","EmployeeEmail","ManagerLogin","ManagerName","ManagerEmail","EquipmentEnglishText","EquipmentFrenchText","ReasonEnglishText","ReasonFrenchText","ReasonDetail","BuildingID" -delimiter "|" -Encoding Unicode

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
foreach ($Row in $CSVData)
{
	if ($Row.Title -ne "Title")
	{

		Add-PnPListItem -List $ListName -Values @{"Title" = $([guid]::NewGuid().toString())
                            "StartHour" = $($Row.StartHour)
                            "EndHour" = $($Row.EndHour)
                            "FloorID" = $($Row.FloorID)
                            "Status" = $($Row.Status)
                            "VisitorCount" = $($Row.VisitorCount)
                            "EntryDate" = $($Row.EntryDate)
                            "EntryDateID" = $($Row.EntryDateID)
                            "EmployeeName" = $($Row.EmployeeName)
                            "EmployeeEmail" = $($Row.EmployeeEmail)
                            "ManagerLogin" = $($Row.ManagerLogin)
                            "ManagerName" = $($Row.ManagerName)
                            "ManagerEmail" = $($Row.ManagerEmail)
                            "EquipmentEnglishText" = $($Row.EquipmentEnglishText)
                            "EquipmentFrenchText" = $($Row.EquipmentFrenchText)
                            "ReasonEnglishText" = $($Row.ReasonEnglishText)
                            "ReasonFrenchText" = $($Row.ReasonFrenchText)
                            "ReasonDetail" = $($Row.ReasonDetail)
                            "BuildingID" = $($Row.BuildingID)}
    }
}


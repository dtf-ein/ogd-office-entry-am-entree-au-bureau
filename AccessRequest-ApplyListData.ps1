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
$CSVData = Import-Csv -Path ($CSVFolder + $ListName + ".csv") -header "Title","StartHour","EndHour","FloorID","Status","VisitorCount","EntryDate","EntryDateID","EmployeeLogin","EmployeeName","EmployeeEmail","ManagerLogin","ManagerName","ManagerEmail","EquipmentEnglishText","EquipmentFrenchText","ReasonEnglishText","ReasonFrenchText","ReasonDetail","BuildingID" -delimiter "|" -Encoding Unicode

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
$count = 0
foreach ($Row in $CSVData)
{
	if ($Row.Title -ne "Title")
	{

        # Randomize the request details
        $StartHour = Get-Random -Minimum 9 -Maximum 12
        $EndHour = Get-Random -Minimum 13 -Maximum 17
        $EntryDate = (Get-Date).AddDays(++$count)

		Add-PnPListItem -List $ListName -Values @{"Title" = $([guid]::NewGuid().toString())
                            "StartHour" = $StartHour
                            "EndHour" = $EndHour
                            "EntryDate" = $EntryDate.toString("yyyy-MM-dd hh:mm:ss tt")
                            "EntryDateID" = $EntryDate.toString("yyyyMMdd")
                            "FloorID" = $($Row.FloorID)
                            "Status" = $($Row.Status)
                            "VisitorCount" = $($Row.VisitorCount)
                            "EmployeeLogin" = $($Row.EmployeeLogin)
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


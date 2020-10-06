# Load an AccessRequest.csv that was created by the Power Automate flow

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL,
    [Parameter(Mandatory=$True, HelpMessage="CSV file to import")]
    [string]$CSVFile
)

$ListName = "AccessRequest"

###!!! CSV Fields need to be hard coded in script below.

#Get the CSV file contents
$CSVData = Import-Csv -Path ($CSVFile) -header "Title","StartHour","EndHour","FloorID","Status","VisitorCount","EntryDate","EntryDateID","EmployeeLogin","EmployeeName","EmployeeEmail","ManagerLogin","ManagerName","ManagerEmail","ReasonEnglishText","ReasonFrenchText","ReasonCode","BuildingID","Modified","Created","ReasonDetail" -delimiter "," -Encoding UTF8

#Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

#Iterate through each Row in the CSV and import data to SharePoint Online List
foreach ($Row in $CSVData)
{
    if ($Row.Title -ne "Title")
    {
        Add-PnPListItem -List $ListName -Values @{
            "Title" = $($Row.Title)
            "StartHour" = $($Row.StartHour)
            "EndHour" = $($Row.EndHour)
            "FloorID" = $($Row.FloorID)
            "Status" = $($Row.Status)
            "VisitorCount" = $($Row.VisitorCount)
            "EntryDate" = $($Row.EntryDate)
            "EntryDateID" = $($Row.EntryDateID)
            "EmployeeLogin" = $($Row.EmployeeLogin)
            "EmployeeName" = $($Row.EmployeeName)
            "EmployeeEmail" = $($Row.EmployeeEmail)
            "ManagerLogin" = $($Row.ManagerLogin)
            "ManagerName" = $($Row.ManagerName)
            "ManagerEmail" = $($Row.ManagerEmail)
            "ReasonEnglishText" = $($Row.ReasonEnglishText)
            "ReasonFrenchText" = $($Row.ReasonFrenchText)
            "ReasonCode" = $($Row.ReasonCode)
            "BuildingID" = $($Row.BuildingID)
            "Modified" = $($Row.Modified)
            "Created" = $($Row.Created)
            "ReasonDetail" = $($Row.ReasonDetail)}
    }
}


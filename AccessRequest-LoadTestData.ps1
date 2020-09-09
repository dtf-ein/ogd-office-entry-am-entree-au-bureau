# Load random performance test data.
# An ./Employees.csv should exist with the following data:
#
# Employee.Name@domain.ca,Manager.Name@domain.ca
#
# The ./UserSettings-ApplyListData.ps1 can then be used to create the employee/manager
# relationships using the same ./Employees.csv.

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]
    $SiteURL
    ,
    [Parameter(Mandatory=$True, HelpMessage="Date to start creating requests on: YYYY-MM-DD")]
    [ValidatePattern("[0-9]{4}-[0-9]{2}-[0-9]{2}")]
    [string]
    $StartDateStr
    ,
    [Parameter(Mandatory=$True, HelpMessage="Number of days of access requests to create")]
    [ValidateRange(1,100)]
    [Int]
    $numberDays
)

# Load testing data
$ListName = "AccessRequest"
$Floors = @(
    [pscustomobject]@{ID="360 Lisgar - 2";BuildingID="360 Lisgar"},
    [pscustomobject]@{ID="360 Lisgar - 8";BuildingID="360 Lisgar"},
    [pscustomobject]@{ID="100 Metcalfe - 3";BuildingID="100 Metcalfe"}
    [pscustomobject]@{ID="100 Metcalfe - 4";BuildingID="100 Metcalfe"}
    [pscustomobject]@{ID="100 Metcalfe - 5";BuildingID="100 Metcalfe"}
    [pscustomobject]@{ID="100 Metcalfe - 6";BuildingID="100 Metcalfe"}
    [pscustomobject]@{ID="100 Metcalfe - 7";BuildingID="100 Metcalfe"}
    [pscustomobject]@{ID="101 Colonel By Drive - 1";BuildingID="101 Colonel By Drive"}
    [pscustomobject]@{ID="101 Colonel By Drive - 6";BuildingID="101 Colonel By Drive"}
)
$Reasons = @(
    [pscustomobject]@{English="Critical work";French="Travail critique"},
    [pscustomobject]@{English="Regular work";French="Travail regulier"},
    [pscustomobject]@{English="Pick up a document";French="Recuperer un document"}
)

$Employees = Import-Csv -Path ("./Employees.csv") -header "Email","ManagerEmail" -delimiter "," -Encoding UTF8
$TextInfo = (Get-Culture).TextInfo
$StartDate = [datetime]::parseexact($StartDateStr + " 12:00:00 PM", "yyyy-MM-dd hh:mm:ss tt", $null) # Set time to noon to avoid timezone conversion problems with date

Write-Host("Creating $numberDays day(s) of requests starting $StartDateStr for $($Employees.count) employees")

# Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Create the list items
For ($i=0; $i -lt $numberDays; $i++) {

    # Get the date
    $EntryDate = $StartDate.AddDays($i)

    # Foreach employee in the .csv, create a request
    foreach ($Employee in $Employees){

        $Floor = Get-Random -InputObject $Floors
        $Reason = Get-Random -InputObject $Reasons

        $EmployeeName = $TextInfo.ToTitleCase(($Employee.Email.Split("@")[0]).Split("."))
        $EmployeeEmail = $EmployeeName.Replace(" ", ".") + "@ssc-spc.gc.ca"

        $ManagerName = $TextInfo.ToTitleCase(($Employee.ManagerEmail.Split("@")[0]).Split("."))
        $ManagerEmail = $ManagerName.Replace(" ", ".") + "@ssc-spc.gc.ca"

        Add-PnPListItem -List $ListName -Values @{
            "Title" = $([guid]::NewGuid().toString())
            "StartHour" = $(Get-Random -Minimum 9 -Maximum 12)
            "EndHour" = $(Get-Random -Minimum 13 -Maximum 17)
            "EntryDate" = $EntryDate.toString("yyyy-MM-dd hh:mm:ss tt")
            "EntryDateID" = $EntryDate.toString("yyyyMMdd")
            "FloorID" = $Floor.ID
            "BuildingID" = $Floor.BuildingID
            "Status" = "A"
            "VisitorCount" = "0"
            "EmployeeLogin" = $EmployeeEmail
            "EmployeeName" = $EmployeeName
            "EmployeeEmail" = $EmployeeEmail.ToLower()
            "ManagerLogin" = $ManagerEmail
            "ManagerName" = $ManagerName
            "ManagerEmail" = $ManagerEmail.ToLower()
            "EquipmentEnglishText" = ""
            "EquipmentFrenchText" = ""
            "ReasonEnglishText" = $Reason.English
            "ReasonFrenchText" = $Reason.French
            "ReasonDetail" = "Load testing data"
        }
    }
}


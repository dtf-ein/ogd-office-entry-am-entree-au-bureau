#Config Variables

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]
    $SiteURL
    ,
    [Parameter(Mandatory=$True, HelpMessage="Number of access requests to create")]
    [ValidateRange(1,8000)]
    [Int]
    $numberRequests
)

# Load testing data
$ListName = "AccessRequest"
$States = @("A", "C")
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
    [pscustomobject]@{English="Regular work";French="Travail régulier"},
    [pscustomobject]@{English="Pick up a document";French="Récupérer un document"}
)

# Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Create the list items
For ($i=0; $i -lt $numberRequests; $i++) {

    # Randomize the request details
    $EntryDate = (Get-Date).AddDays($i)
    $Floor = Get-Random -InputObject $Floors
    $Reason = Get-Random -InputObject $Reasons

    Add-PnPListItem -List $ListName -Values @{
        "Title" = $([guid]::NewGuid().toString())
        "StartHour" = $(Get-Random -Minimum 9 -Maximum 12)
        "EndHour" = $(Get-Random -Minimum 13 -Maximum 17)
        "EntryDate" = $EntryDate.toString("yyyy-MM-dd hh:mm:ss tt")
        "EntryDateID" = $EntryDate.toString("yyyyMMdd")
        "FloorID" = $Floor.ID
        "BuildingID" = $Floor.BuildingID
        "Status" = $(Get-Random -InputObject $States)
        "VisitorCount" = "0"
        "EmployeeLogin" = "Patrick.Heard@ssc-spc.gc.ca"
        "EmployeeName" = "Patrick Heard"
        "EmployeeEmail" = "Patrick.Heard@ssc-spc.gc.ca".ToLower()
        "ManagerLogin" = "Brittany.Hurley@ssc-spc.gc.ca"
        "ManagerName" = "Brittany Hurley"
        "ManagerEmail" = "Brittany.Hurley@ssc-spc.gc.ca".ToLower()
        "EquipmentEnglishText" = ""
        "EquipmentFrenchText" = ""
        "ReasonEnglishText" = $Reason.English
        "ReasonFrenchText" = $Reason.French
        "ReasonDetail" = "Load testing data"
    }
}


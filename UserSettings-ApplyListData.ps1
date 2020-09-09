# Load the UserSettings table with the employee/manager relationships
# from the ./Employees.csv file

Param(
    [Parameter(Mandatory=$True, HelpMessage="URL of the SharePoint site that contains the list: https://someurl.ca")]
    [string]$SiteURL
)

$AgreedDate = Get-Date
$ListName = "UserSetting"
$TextInfo = (Get-Culture).TextInfo
$Employees = Import-Csv -Path ("./Employees.csv") -header "Email","ManagerEmail" -delimiter "," -Encoding UTF8

# Connect to site
Connect-PnPOnline -Url $SiteUrl -UseWebLogin

# Iterate through each Employee in the CSV and import data
foreach ($Employee in $Employees)
{

    $EmployeeName = $TextInfo.ToTitleCase(($Employee.Email.Split("@")[0]).Split("."))
    $EmployeeEmail = $EmployeeName.Replace(" ", ".") + "@ssc-spc.gc.ca"

    $ManagerName = $TextInfo.ToTitleCase(($Employee.ManagerEmail.Split("@")[0]).Split("."))
    $ManagerEmail = $ManagerName.Replace(" ", ".") + "@ssc-spc.gc.ca"

    Write-Host("Adding $EmployeeName with manager $ManagerName")

    Add-PnPListItem -List $ListName -Values @{
        "Title" = $EmployeeEmail
        "AgreedToPrivacy" = $AgreedDate.toString("dd/MM/yyyy hh:mm:ss tt")
        "AgreedToHealthSafety" = $AgreedDate.toString("dd/MM/yyyy hh:mm:ss tt")
        "ManagerLogin" = $ManagerEmail
    }
}


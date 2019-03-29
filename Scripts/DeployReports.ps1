param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]
    $path,
    [Parameter(Mandatory = $true)]
    [ValidateSet('D', 'T', 'A', 'P')]
    [String]
    $environment,
    [Parameter(Mandatory = $true)]
    [String]
    $applicationId,
    [Parameter(Mandatory = $true)]
    [String]
    $clientSecret
)

# Install dependencies
# Install-PackageProvider -Name NuGet -Force -Scope CurrentUser
Write-Host "Installing necesarry modules"
Install-Module -Name MicrosoftPowerBIMgmt -Verbose -Scope CurrentUser

# Create credential of Application ID and Client Secret
$username = $applicationId
$password = $clientSecret | ConvertTo-SecureString -asPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($username, $password)

# Connect to the Power BI service
Write-Host "Connecting to PowerBI"
Connect-PowerBIServiceAccount -Credential $credential

# Loop through all Reports (.pbix files) in the supplied path
Write-Host "Finding all wokspaces in $path"
$reports = Get-ChildItem -Path $path -Recurse -Include *.pbix
$reports

foreach($report in $reports)
{
    Write-Host "Deploying report $($report.BaseName)"

    # Get publish settings from .json file
    $PublishFile = "$($report.DirectoryName)/$($report.BaseName).Publish.json"
    Write-Host "Getting publish settings from publish file $PublishFile"
    $publishSettings = (Get-Content $PublishFile | ConvertFrom-Json -Verbose).$environment
    $publishSettings

    # Get workpsace (and leave the deleted workspaces alone)
    Write-Host "Retrieving workspace information from PowerBI"
    $workspace = Get-PowerBIWorkspace -Scope Organization  -Filter "name eq '$($publishSettings.WorkspaceName)' and not(state eq 'Deleted')" -First 1 -Verbose
    $workspace       

    # Publish report to workspace defined in publish settings with the name defined in publish settings
    Write-Host "Deploying report"
    New-PowerBIReport -Path $report.FullName -Name $publishSettings.ReportName -Workspace $workspace -ConflictAction CreateOrOverwrite
}





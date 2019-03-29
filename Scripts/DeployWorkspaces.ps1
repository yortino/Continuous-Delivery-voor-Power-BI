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
Install-Module -Name MicrosoftPowerBIMgmt -Force -Scope CurrentUser

# Create credential of Application ID and Client Secret
$username = $applicationId
$password = $clientSecret | ConvertTo-SecureString -asPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($username, $password)

# Connect to the Power BI service
Write-Host "Connecting to PowerBI"
Connect-PowerBIServiceAccount -Credential $credential -Verbose

# Loop through all Workspaces (.json files) in the supplied path
Write-Host "Finding all wokspaces in $path"
$workspaces = Get-ChildItem -Path $path -Recurse -Include *.json -Verbose
$workspaces

foreach($workspace in $workspaces)
{
    Write-Host "Deploying workspace $($workspace.BaseName)"

    # Get publish settings from .json file
    Write-Host "Getting publish settings from publish file"
    $publishSettings = (Get-Content $workspace.FullName -Verbose | ConvertFrom-Json -Verbose).$environment
    $publishSettings

    # Get workpsace (and leave the deleted workspaces alone)
    Write-Host "Retrieving workspace information from PowerBI"
    $workspace = Get-PowerBIWorkspace -Scope Organization  -Filter "name eq '$($publishSettings.Name)' and not(state eq 'Deleted')" -First 1 -Verbose
    if(!$workspace)
    {
        # Create workspace and load
        Write-Host "Creating workspace on PowerBI since it does not exist"
        Invoke-PowerBIRestMethod -Url "groups?workspaceV2=true" -Method Post -Body "{`"name`":`"$($publishSettings.Name)`"}" -Verbose

        # Get newly created workspace
        Write-Host "Retrieving newly created workspace information from PowerBI"
        $workspace = Get-PowerBIWorkspace -Scope Organization  -Filter "name eq '$($publishSettings.Name)' and not(state eq 'Deleted')" -First 1 -Verbose
    }
    $workspace

    # Loop through all users
    foreach($user in $workspace.users)
    {
        # Check if user should be there
        if($publishSettings.Users.emailAddress -notcontains $user.UserPrincipalName)
        {
            # Remove user
            Write-Host "Removing user $($user.UserPrincipalName)"
            Remove-PowerBIWorkspaceUser -Scope Organization -Workspace $workspace -UserPrincipalName $user.UserPrincipalName -Verbose
        }
    }

    # Loop through all users in publish file
    foreach($user in $publishSettings.Users)
    {
        # Create and update user
        Write-Host "Adding or updating user $($user.emailAddress) as $($user.groupUserAccessRight)"
        Add-PowerBIWorkspaceUser -Scope Organization -Workspace $workspace -UserPrincipalName $user.emailAddress -AccessRight $user.groupUserAccessRight -Verbose
    }

    # Update desciption of workspace
    Write-Host "Updating the description of the workspace"
    $workspace.Description = $publishSettings.Description
    Set-PowerBIWorkspace -Scope Organization -Workspace $workspace -Verbose
}
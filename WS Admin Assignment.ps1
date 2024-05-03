$workspace = Import-Csv C:\PowerBIExport\workspace.csv
# Connect to Power BI - Power BI Connection
Connect-PowerBIServiceAccount

Foreach ($i in $workspace) {

    Write-Host "-----------This is a New Cycle--------------------" -f Green

    # Get Details
    $WorkspaceName = $i.Workspace

    # Get workspace by workspace name
    $Workspace = Get-PowerBIWorkspace -Scope Organization -Name $WorkspaceName
    $workspaceId = $Workspace.id
    Write-Host "$workspaceName-$workspaceId"
    Add-PowerBIWorkspaceUser -Scope Organization -Id $workspaceId -UserEmailAddress shmandal@suncor.com -AccessRight Admin
    
    }
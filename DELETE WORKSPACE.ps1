## DELETE WORKSPACE
# Set Parameters


$data = Import-Csv C:\PowerBIExport\WorksapceDeletePhase2\data_batch11.csv
# Add the current date and time to the log entry
$currentTime = Get-Date -Format "yyyy-MM-dd HH_mm"
# Set Local File Paths
$logFilePath = "C:\PowerBIExport\WorksapceDeletePhase2\Logs\DeletedWorkspace-$currentTime.csv"  # Define the log file path
$logEntry = "Workspacename`tWorkspaceId`tDeleteFlag`tMessage`tDate"  # Use tab character as separator
# Add-Content -Path $logFilePath -Value "-------------------------------"
Add-Content -Path $logFilePath -Value $logEntry

# Connect to Power BI - Power BI Connection
Connect-PowerBIServiceAccount

# For all records (reports) in the CSV file
Foreach ($i in $data) {

    Write-Host "-----------This is a New Cycle--------------------" -f Green

    # Get Details
    $WorkspaceName = $i.Workspace
    $WorkspaceID = $i.WorkspaceID
    #$DatasetName = $i.DatasetName
    #$DatasetId = $i.DatasetId
    #$User = $i.User


    # Get workspace by workspace name
    $Workspace = Get-PowerBIWorkspace -Scope Organization -Name $WorkspaceName

    if ($Workspace) {

        # Worksapce Id
        $workspaceId = $Workspace.id
        Write-Host "$workspaceName-$workspaceId" -f Green
        Add-PowerBIWorkspaceUser -Scope Organization -Id $workspaceId -UserEmailAddress youremail@address.com -AccessRight Admin
        $WorksapceURL = 'groups/' + $workspaceId
        $ErrorActionPreference = 'SilentlyContinue'
        Invoke-PowerBIRestMethod -Url $WorksapceURL -Method Delete
        if($?){
            Write-Host "Deleted..." -f Green
            Add-Content -Path $logFilePath -Value "$WorkspaceName`t$WorkspaceID`tTRUE`tSuccess: Worksapce- $workspaceName Deleted`t$currentTime"
        }
        else{
            Write-Host "Not Deleted..." -f Red
            Add-Content -Path $logFilePath -Value "$WorkspaceName`t$WorkspaceID`tFALSE`tFailure: Worksapce- $workspaceName Not Deleted`t$currentTime"
        }
        
    }

    else {
        Write-Host ("Worksapce Not found in workspace") -f Red
    }
    


}  

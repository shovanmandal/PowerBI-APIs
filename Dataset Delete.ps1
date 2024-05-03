
### DELETE DATASET IN WORKSPACE ###

# Add the current date and time to the log entry
$currentTime = Get-Date -Format "yyyy-MM-dd HH_mm"
# Set Local File Paths
$folder = "C:\PowerBIExport"
$logFilePath = "C:\PowerBIExport\Logs\DatasetDelete-$currentTime.csv"  # Define the log file path

$data = Import-Csv C:\PowerBIExport\data.csv
# Connect to Power BI - Power BI Connection
Connect-PowerBIServiceAccount

# Initialize the log file (create or clear)
# $logdate = "$currentTime - Report Export Log"

$logEntry = "Workspacename`tDatasetName`tWorkspaceId`tDatasetID`tStatusCode`tMessage`tDate"  # Use tab character as separator

# Add-Content -Path $logFilePath -Value $logdate
# Add-Content -Path $logFilePath -Value "-------------------------------"

Add-Content -Path $logFilePath -Value $logEntry


# For all records (reports) in the CSV file
Foreach ($i in $data) {

    Write-Host "-----------This is a New Cycle--------------------" -f Green

    # Get Details
    $WorkspaceName = $i.Workspace
    $WorkspaceID = $i.WorkspaceID
    $DatasetName = $i.DatasetName
    $DatasetId = $i.DatasetId
    #$User = $i.User


    # Get workspace by workspace name
    $Workspace = Get-PowerBIWorkspace -Scope Organization -Name $WorkspaceName
    $workspaceId = $Workspace.id
    Write-Host "$workspaceName-$workspaceId"
    Add-PowerBIWorkspaceUser -Scope Organization -Id $workspaceId -UserEmailAddress c1.shmandal@suncor.onmicrosoft.com -AccessRight Admin

    # Get the report
    $dataset = Get-PowerBIDataset -Name $DatasetName -Workspace $Workspace
    if ($dataset) {

        # Dataset Id
        $datasetID = $dataset.id
        Write-Host "$datasetID/$DatasetName" -f Green
        $DatasetURL = 'groups/' + $workspaceId + '/datasets/' + $datasetID
        $ErrorActionPreference = 'SilentlyContinue'
        
        Invoke-PowerBIRestMethod -Url $DatasetURL -Method Delete
        if($?){
            Write-Host "Deleted" -f Green

            Add-Content -Path $logFilePath -Value "$WorkspaceName`t$DatasetName`t$WorkspaceID`t$datasetID`t200`tSuccess: Dataset- $DatasetName Deleted`t$currentTime"
        }
        else {
            Write-Host "Not Deleted" -f Red

            Add-Content -Path $logFilePath -Value "$WorkspaceName`t$DatasetName`t$WorkspaceID`t$datasetID`t404`tError: Dataset- $DatasetName Could not be deleted`t$currentTime"
        }
        
    }

    else {
        Write-Host ("Dataset Not found in workspace") -f Red

        Add-Content -Path $logFilePath -Value "$WorkspaceName`t$DatasetName`t$WorkspaceID`t$datasetID`t404`tError: Dataset- $DatasetName Not found in workspace`t$currentTime"
    }
    


}  

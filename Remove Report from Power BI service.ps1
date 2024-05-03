# Set Parameters
$folder = "C:\PowerBIExport\Logs\ExportLog-2023-11-17 16_23.csv"
$logfilepath = Import-Csv $folder -Delimiter `t

# Add the current date and time to the log entry
$currentTime = Get-Date -Format "yyyy-MM-dd HH_mm"
# Set Local File Paths
$deletelogPath = "C:\PowerBIExport\Logs\DeleteLog-$currentTime.csv"  # Define the log file path
$deletelogEntry = "Workspacename`tReportName`tWorkspaceId`tReportID`tStatusCode`tMessage`tDate"  # Use tab character as separator

# Add-Content -Path $logFilePath -Value $logdate
# Add-Content -Path $logFilePath -Value "-------------------------------"

Add-Content -Path $deletelogPath -Value $deletelogEntry


ForEach($i in $logfilepath){

    
    $ReportName = $i.ReportName
    $WorkspaceName = $i.Workspacename
    $status  = $i.StatusCode
    $WorkspaceID = $i.WorkspaceID
    $ReportID = $i.ReportID

    Write-Host $ReportName $status

    if($status -eq "200"){
        $statusMessage = "$ReportName- $ReportId will be deleted from $WorkspaceName- $WorkspaceID"
        Write-Host "$ReportName- $ReportId will be deleted from $WorkspaceName- $WorkspaceID" -f Green
        Remove-PowerBIReport -Id $ReportID -WorkspaceId $WorkspaceID
        Add-Content -Path $deletelogPath -Value "$WorkspaceName`t$ReportName`t$WorkspaceID`t$ReportID`t200`t$statusMessage`t$currentTime"
    }
    else{
        $statusMessage = "$ReportName- $ReportId can not be deleted from $WorkspaceName- $WorkspaceID"
        Add-Content -Path $deletelogPath -Value "$WorkspaceName`t$ReportName`t$WorkspaceID`t$ReportID`t400`t$statusMessage`t$currentTime"
        Write-Host "$ReportName- $ReportId can not be deleted from $WorkspaceName- $WorkspaceID" -f Red
    }

}
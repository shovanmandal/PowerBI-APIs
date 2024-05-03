# Set Parameters

# Set Sharepoint paths
$WebURL = "https://abc.sharepoint.com/sites/SiteName/"
$rootFolder = "Shared Documents"

# Add the current date and time to the log entry
$currentTime = Get-Date -Format "yyyy-MM-dd HH_mm"
$currentDate = Get-Date -Format "yyyy-MM-dd"
# Set Local File Paths
$folder = "C:\PowerBIExport"
$logFilePath = "C:\PowerBIExport\Report_Download\Logs\ExportLog-$currentTime.csv"  # Define the log file path
$reports = Import-Csv C:\PowerBIExport\Report_Download\data_batch74.csv

# Connect to PnP Online - Sharepoint Connection
Connect-PnPOnline -Url $WebURL -Interactive

# Connect to Power BI - Power BI Connection
Connect-PowerBIServiceAccount

# Initialize the log file (create or clear)
# $logdate = "$currentTime - Report Export Log"

$logEntry = "Workspacename`tReportName`tWorkspaceId`tReportID`tStatusCode`tMessage`tDate"  # Use tab character as separator

# Add-Content -Path $logFilePath -Value $logdate
# Add-Content -Path $logFilePath -Value "-------------------------------"

Add-Content -Path $logFilePath -Value $logEntry





# For all records (reports) in the CSV file
Foreach ($i in $reports) {

    Write-Host "-----------This is a New Cycle--------------------" -f Green

    # Get Details
    $WorkspaceName = $i.WorkspaceName
    $WorkspaceID = $i.WorkspaceID
    $ReportName = $i.ReportName
    $ReportID = $i.ReportID
    #$User = $i.User

    # Get workspace by workspace name
    $Workspace = Get-PowerBIWorkspace -Scope Organization -Name $WorkspaceName

    if (!($Workspace)) {
        Write-Host "Workspace $workspaceName not found"
        Add-Content -Path $logFilePath -Value "$WorkspaceName`t$ReportName`t$WorkspaceID`t$ReportID`t404`tError: Workspace- $WorkspaceName not found`t$currentTime"
        continue
    }

    # Get Workspace Id
    $workspaceId = $Workspace.id
    # Write-Host "$workspaceId/$workspaceName" -f Green-- printing the workspaceID and Name

    # Add C1.Shovan Mandal as Admin of the Workspaces
    Add-PowerBIWorkspaceUser -Scope Organization -Id $workspaceId -UserEmailAddress youremail@address.com -AccessRight Admin

    # Get the report
    $report = Get-PowerBIReport -Id $ReportID -WorkspaceId $WorkspaceID

    if ($report) {

        # Report Id
        $reportId = $report.id

        Write-Host $report.name $reportID

        # Set SharePoint Folder path for Workspace
        $LibraryName = "$rootFolder/$workspaceName"

        if (Test-Path "C:\PowerBIExport\Report_Download\Worksapce\$WorkspaceName") {
            # Folder exists!! Try Export Report
            Export-PowerBIReport -Id $reportId -WorkspaceId $workspaceId -OutFile "C:\PowerBIExport\Report_Download\Worksapce\$WorkspaceName\$ReportName.pbix"
        }
        else {
            # Folder does not exist - Create Folder
            New-Item -Path "c:\PowerBIExport\Report_Download\Worksapce" -Name "$WorkspaceName" -ItemType "directory"

            # Try Exporting Power BI report
            Export-PowerBIReport -Id $reportId -WorkspaceId $workspaceId -OutFile "C:\PowerBIExport\Report_Download\Worksapce\$WorkspaceName\$ReportName.pbix"

        }

        # Create Folder if it doesn't exist
        Resolve-PnPFolder -SiteRelativePath $LibraryName

        # Add file to SharePoint Folder
        # Set File path
        $filepath = "C:\PowerBIExport\Report_Download\Worksapce\$WorkspaceName\$ReportName.pbix"

        if (Test-Path $filepath) {
            $statusCode = 200  # Set the status code for success
            $statusMessage = "Report Exported Successfully" # Set the status message for success
            Add-PnPFile -Path $filepath -Folder $LibraryName
        }
        else {
            $statusCode = 400  # Set the status code for Failure
            $statusMessage = "Report Export Failed" # Set the status message for success
            Write-Host "File do not Exist: $filepath" -foregroundcolor Red
        }

        # Log the report export status
        Add-Content -Path $logFilePath -Value "$WorkspaceName`t$ReportName`t$WorkspaceID`t$ReportID`t$statusCode`t$statusMessage`t$currentTime"


    }
    else {
        Write-Host "Report not found in the workspace"
        Add-Content -Path $logFilePath -Value "$WorkspaceName`t$ReportName`t$WorkspaceID`t$ReportID`t404`tError: Report- $ReportName not found in $WorkspaceName`t$currentTime"
    }

}

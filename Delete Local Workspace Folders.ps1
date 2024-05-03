# Set Parameters
$folder = "C:\PowerBIExport\Logs\ExportLog-2023-11-17 16_23.csv"
$logfilepath = Import-Csv $folder -Delimiter `t
$path = $logfilepath | Select Workspacename -Unique

ForEach($i in $path){
    Write-Host "----------this is for deleting local folders---------------------" -f Green

    $WorkspaceName = $i.Workspacename
    Write-Host "$WorkspaceName" -f Green
    $folderpath = "C:\PowerBIExport\"+$WorkspaceName
    Write-Host "$folderpath" -f Green
    Remove-Item -Path $folderpath -Recurse -Force

    Write-Host
    Write-Host

}
#This script will generate a report on the usage of the subscription over 2024, summarize and save it to a json file

#Before you begin, ensure you have:
# The Az module installed
# Az.CostManagement module installed 
# You registered the Microsoft.CostManagementExports resource provider Register-AzResourceProvider -ProviderNamespace Microsoft.CostManagementExports
# Created a storage account to hold the reports and set variables for the subscription, resource group, and container name

#Set the subscription you want to report on
$reportSubscriptionId = "[Your Subscription ID]"
$storageSubscriptionId = "[Your Subscription ID]"
$storageResourceGroup = "[Your Resource Group]"
$storageContainerName = "dataexport"
$storageAccountName = "[Your Storage Account Name]"

#Set and check the context
Set-AzContext -SubscriptionId $storageSubscriptionId
$accountInfo = Get-AzContext
if ($accountInfo.Subscription.Id -ne $storageSubscriptionId) {
    Write-Error "Failed to set the subscription. Please check the subscription ID."
    exit
}

# Check if the storage container exists
$containerExists = Get-AzStorageContainer -Name $storageContainerName -Context $storageContext -ErrorAction SilentlyContinue
if ($containerExists) {
    Write-Error "The storage container '$storageContainerName' already exists. Please choose a different name."
    exit
}

# Create a new storage container
New-AzStorageContainer -Name $storageContainerName -Context $storageContext
Write-Output "The storage container '$storageContainerName' has been created successfully."

# Query the billing data for the last 2024
# We can only check 3 months at a time so split into quarters
$quarters = @(
    @{ start = "2024-01-01T00:00:00Z"; end = "2024-03-31T23:59:59Z"; name = "Q1" },
    @{ start = "2024-04-01T00:00:00Z"; end = "2024-06-30T23:59:59Z"; name = "Q2" },
    @{ start = "2024-07-01T00:00:00Z"; end = "2024-09-30T23:59:59Z"; name = "Q3" },
    @{ start = "2024-10-01T00:00:00Z"; end = "2024-12-31T23:59:59Z"; name = "Q4" }
)

# Loop through each quarter and generate the cost details report
foreach ($quarter in $quarters) {

    # Check if the end date is in the future
    $currentDate = (Get-Date).ToUniversalTime()
    if ([DateTime]::Parse($quarter.end) -gt $currentDate) {
        $quarter.end = $currentDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

    New-AzCostManagementExport `
    -Name "Export$($quarter.name)2" `
    -DefinitionType  "Usage" `
    -Scope "/subscriptions/$reportSubscriptionId" `
    -DestinationResourceId  "/subscriptions/$storageSubscriptionId/resourceGroups/$storageResourceGroup/providers/Microsoft.Storage/storageAccounts/$storageAccountName" `
    -DestinationContainer  "$storageContainerName" `
    -DefinitionTimeframe "Custom" `
    -DestinationRootFolderPath "exports" `
    -TimePeriodFrom "$($quarter.start)" `
    -TimePeriodTo "$($quarter.end)" `
    -DataSetGranularity "Daily" `
    -Format "csv" 
}

#Now request the export 
foreach ($quarter in $quarters) {
    Invoke-AzCostManagementExecuteExport `
        -ExportName "Export$($quarter.name)2" `
       -Scope "/subscriptions/$reportSubscriptionId"
}

# Find the generated files
$storageAccountKey = (Get-AzStorageAccountKey -ResourceGroupName $storageResourceGroup -Name $storageAccountName)[0].Value
$storageContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey
$blobs = Get-AzStorageBlob -Container $storageContainerName -Context $storageContext

#Wait for Blobs to be available
$maxRetries = 12
$retryCount = 0
$expectedBlobCount = 4

while ($retryCount -lt $maxRetries) {
    $blobs = Get-AzStorageBlob -Container $storageContainerName -Context $storageContext
    if ($blobs.Count -eq $expectedBlobCount) {
        Write-Output "All $expectedBlobCount blobs are available."
        break
    } else {
        Write-Output "Expected $expectedBlobCount blobs, but found $($blobs.Count). Retrying in 5 seconds..."
        Start-Sleep -Seconds 5
        $retryCount++
    }
}

if ($retryCount -eq $maxRetries) {
    Write-Error "Failed to find all $expectedBlobCount blobs after $($maxRetries * 5) seconds."
    exit
}

# Download each file to local disk
foreach ($blob in $blobs) {
    $fileName = $blob.Name
    $fileNameParts = $fileName -split '/'
    $actualFileName = $fileNameParts[-1]
    Get-AzStorageBlobContent -Blob $fileName -Container $storageContainerName -Destination "./$actualFileName" -Context $storageContext
}

#Build our own report from the data
$report = [PSCustomObject]@{
    Name                   = "Unknown"
    RG_TotalResourceGroupsYear = 0
    RES_TotalResourcesYear      = 0
    SER_TotalConsumedServiceTypesYear = 0
}

$report.Name = (Get-AzADUser).GivenName

#Merge all csv files 
$usageDetailsLines = @()
$csvFiles = Get-ChildItem -Path . -Filter *.csv
foreach ($csvFile in $csvFiles) {
    $csvContent = Import-Csv -Path $csvFile.FullName
    $usageDetailsLines += $csvContent
}

#Get the total number of resource groups and resources (that have been billed)
$report.RG_TotalResourceGroupsYear = $usageDetailsLines  | ForEach-Object { $_.ResourceGroup } | Sort-Object -Unique | Measure-Object | Select-Object -ExpandProperty Count
$report.RES_TotalResourcesYear = $usageDetailsLines  | ForEach-Object { $_.InstanceId } | Sort-Object  -Unique | Measure-Object | Select-Object -ExpandProperty Count
$report.SER_TotalConsumedServiceTypesYear = $usageDetailsLines |  ForEach-Object { $_.ResourceType } | Sort-Object  -Unique | Measure-Object | Select-Object -ExpandProperty Count

#Save Results to a json file
$report | ConvertTo-Json | Out-File -FilePath "sourceData.json" -Force -Encoding ascii

write-host "Report has been generated and saved to sourceData.json now run buildVideo.py to create the video"
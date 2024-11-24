# About 

This is a sample project to demonstrate gathering data from Azure and generating a video from powerpoint using python to show the data.

It is split into two parts:
1) `fetchData.ps1` - Powershell script to gather data from Azure and generate a powerpoint presentation.
2) `buildVideo.py` - Python script to generate a video from the powerpoint presentation.

Note this example only works on Windows as it uses COM to access powerpoint.

This was created as part of the festive tech calendar were you can find the original video post explain this in more detail.

## Requirements

At a hight level you will need the following:

* Python 3.7
* Powershell
* Powerpoint
* Azure Account
* Azure Storage Account

This project uses powershell and the AZ powershell module. It also uses the Az.CostManagement module. You can install the required packages by running:

```
Install-Module -Name Az -AllowClobber -Scope CurrentUser
Install-Module -Name Az.CostManagement -AllowClobber -Scope CurrentUser
```

You must ensure your Azure account has Microsoft.CostManagementExports resource provider registered. You can do this by running:

`Register-AzResourceProvider -ProviderNamespace Microsoft.CostManagementExports`

Cost reports are generated to an Azure storage account. You must create a storage account and container to hold the reports. You can do this by running:

`New-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -Location $location -SkuName Standard_LRS -Kind StorageV2`

Once these are done update the variables in the powershell script to match your environment.

As it uses powerpoint, you must have Microsoft Powerpoint installed on the machine running the python script. This project uses Python 3.7. You can install the required packages by running:

`pip install pywin32`

## Usage
Run the powershell script `fetchData.ps1` to generate the data into a local file `sourceData.json`. Then run the `buildVideo.py` python script to generate the video.
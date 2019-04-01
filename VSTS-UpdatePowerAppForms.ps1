
#requires
<#
.SYNOPSIS
  This script is used to provide commandlets to update PowerApps data sets to be used in other tenants

.DESCRIPTION
	Uses Office Dev PnP to provisiong all IA artefacts

.PARAMETER
		-FilePath
.INPUTS
  <Inputs if any, otherwise state None>


.OUTPUTS
  <Log file stored in ProvisionArtifacts.log>

.NOTES
  Version:        1.0
  Author:         Ramin Ahmadi
  Creation Date:  11/11/2017
  Purpose/Change: First version

.EXAMPLE
.\VSTS-UpdatePowerAppsData.ps1 -SourceSiteUrl https://contoso.sharepoint.com/sites/dcratdev -TargetSiteUrl https://raminahmadi.sharepoint.com/sites/rat

#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Parameters
Param(    
    #Source information
    [Parameter(Mandatory = $True)]
    [string]$SourceSiteUrl,
    #Target information
    [Parameter(Mandatory = $True)]
    [string]$TargetSiteUrl
)

#region Set Global Variables--------------------------------------------------------------------------------------------------

#endregion Set Global Variables----------------------------------------------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Get app from configuration file
#[xml]$ConfigurationContent = Get-Content "$($AppsLocation)\configuration.xml"
$CurrFolderName = $(Get-Location).Path
$PowerAppsDirectory = "$currFolderName\PowerApps"
$TempDirectory = "$currFolderName\temp"

#New-Item -ItemType directory -Path $NewPackagesDirectory
#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Update-AppData() {
    $files = Get-ChildItem $PowerAppsDirectory -Filter *.zip 
    $files | ForEach-Object {
        $Name = $_.Name
        $Path = $_.FullName
        $TempFolder = "$TempDirectory\$Name"
        Write-Host "App Name: $Name"
        
        # Extract the package
        Expand-Archive -LiteralPath $Path -DestinationPath $TempFolder -Force
        
        # Rename .msapp file to zip so we can extract it
        Get-ChildItem -Path $TempFolder -Filter *.msapp  -Recurse | Select-Object -First 1 | Rename-Item -NewName { $_.name -Replace '\.msapp', '.zip' }                

        # Extract app contents
        $MsappFile = Get-ChildItem -Path $TempFolder -Filter *.zip  -Recurse | Select-Object -First 1
        $MsappDirectory = (Get-Item $MsappFile.FullName).Directory.FullName
        $MsappExpandDirectoryName = $MsappFile.BaseName
        $MsappExpandDirectoryPath = "$MsappDirectory\$MsappExpandDirectoryName"
        Expand-Archive -LiteralPath $MsappFile.FullName -DestinationPath $MsappExpandDirectoryPath -Force

        # Find data sources
        $EntitiesFile = (Get-Content "$MsappExpandDirectoryPath\entities.json" -Raw) | ConvertFrom-Json
        $SourceDataSources = $EntitiesFile.Entities | Where-Object { $_.type -eq "ConnectedDataSourceInfo" } | Select-Object -Property Name, TableName
        $TargetDataSources = Get-DataSources -SourceSiteUrl $SourceDataSources
        # Update App confige file
        $AppId = (Get-Item $MsappDirectory ).BaseName
        $ConfigFilePath = "$MsappDirectory\$($AppId).json"
        Update-JsonFile -Path $ConfigFilePath -SourceDataSources $SourceDataSources -TargetDataSources $TargetDataSources
        $PropertiesFilePath = "$MsappExpandDirectoryPath\roperties.json"
        Update-JsonFile -Path $PropertiesFilePath -SourceDataSources $SourceDataSources -TargetDataSources $TargetDataSources
        $EntitiesFilePath = "$MsappExpandDirectoryPath\Entities.json"
        Update-JsonFile -Path $EntitiesFilePath -SourceDataSources $SourceDataSources -TargetDataSources $TargetDataSources
        # Update App Flows
        # Update-Flows $_
        # Rename .zip file to .msapp
        Get-ChildItem -Path "$MsappDirectory" -Filter *.zip  -Recurse | Select-Object -First 1 | Rename-Item -NewName { $_.name -Replace '\.zip', '.msapp' }
        # Insert updated files directly to the zip file
        Insert-UpdatedFilesToPackage -AppDirectory "$MsappDirectory" -ExpandedPath "$MsappExpandDirectoryPath"

    }
    # Compress all packages into one zip file
    # Compress-Archive -Path "$($NewPackagesDirectory)\*.zip" -DestinationPath "$($NewPackagesDirectory)\PowerAppsForms"
}

function Get-DataSources($SourceDataSources){
    $TargetDataSources = @()
    try {                
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin  
        # Replace data sources
        DataSources | ForEach-Object {
            $NewDataSource = New-Object System.Object
            $SourceListName = $_.ListName
            $targetListId = Get-PnPList -Identity "$SourceListName"
            $NewDataSource | Add-Member -type NoteProperty -name Name -Value "$SourceListName"
            $NewDataSource | Add-Member -type NoteProperty -name Id -Value "$($targetListId.Id)"
            $TargetDataSources += $NewDataSource
        }
        $json = (ConvertFrom-Json $json)
        $json | ConvertTo-Json -Depth 60 | Set-Content $Path
    }
    catch {
        Write-Host $_.Exception.Message
    }
    return $TargetDataSources
}
function Update-JsonFile($Path, $SourceDataSources, $TargetDataSources) {
    try {
        # Get file content
        $json = Get-Content $Path | Out-String
        # Replace all source url to target url
        $json = $json.Replace($SourceSiteUrl, $TargetSiteUrl)
        
        # Replace data sources
        DataSources | ForEach-Object {
            $SourceListName = $_.Name
            $SourceListId = $_.TableName
            $TargetItem =  $TargetDataSources | ?{$_.Name -eq "$($SourceListName)"}
            $targetListId = $TargetItem.Id
            $json = $json.Replace($SourceListId, $targetListId)
        }
        $json = (ConvertFrom-Json $json)
        $json | ConvertTo-Json -Depth 60 | Set-Content $Path
    }
    catch {
        Write-Host $_.Exception.Message
    }
}

function Update-Flows($App) {
    if ($App.Flows) {
        $App.Flows.Flow | ForEach-Object {
            try {
                $FlowName = $_.Name
                $FlowId = $_.Id
                $Path = $AppsDirectory + "$($AppName)\Microsoft.Flow\flows\$($FlowId)\definition.json"
                Write-Host "Flow : $($FlowName)"
                # Get file content
                $json = Get-Content $Path | Out-String
                # Replace all source url to target url
                $json = $json.Replace($SourceSiteUrl, $TargetSiteUrl)
                # Replace data sources
                $_.DataSources.DataSource | ForEach-Object {
                    $sourceListName = $_.ListName
                    $sourceListId = Get-ListId $SourceSiteUrl $SourceUserName $SourcePassword $sourceListName
                    $targetListId = Get-ListId $TargetSiteUrl $TargetUserName $TargetPassword $sourceListName
                    $json = $json.Replace($sourceListId, $targetListId)
                }
                $json = (ConvertFrom-Json $json)
                $json | ConvertTo-Json -Depth 40 | Set-Content $Path
            }
            catch {
                Write-Host $_.Exception.Message
            }
        }
    }
}

function Insert-UpdatedFilesToPackage($AppsDirectory,$ExpandedPath) {
    $MsappFilePath = Get-ChildItem -Path "$AppsDirectory\" -Filter *.msapp  -Recurse | Select-Object -First 1    
    $PropertiesFilePath = "$ExpandedPath\Properties.json"
    $EntitiesFilePath = "$ExpandedPath\Entities.json"
    # Open the .msapp file for updating
    $zip = [System.IO.Compression.ZipFile]::Open($MsappFilePath.FullName, "Update")
    # Remove existing properties.json and Entities.json
    $PropertiesEntry = $zip.GetEntry("Properties.json")
    $EntitiesEntry = $zip.GetEntry("Entities.json")
    if ($PropertiesEntry) {
        $PropertiesEntry.Delete()
    }
    if ($EntitiesEntry) {
        $EntitiesEntry.Delete()
    }
    # Add updated files to the .msapp file
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $PropertiesFilePath, "Properties.json", "optimal")
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $EntitiesFilePath, "Entities.json", "optimal")
    # Dispose the package
    $zip.Dispose()
}

function Pre-Actions() {
    if (Test-Path $TempDirectory) {
        Write-Host "Folder temp already exists."
    }
    else {
        Write-Host "Creating temp folder..."
        New-Item -ItemType directory -Path $TempDirectory
    }

}

function Post-Actions {
    Remove-Item –Path $TempDirectory –recurse
}
Pre-Actions
Update-AppData
#Post-Actions
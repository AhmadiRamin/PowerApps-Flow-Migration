
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
.\VSTS-UpdatePowerAppsData.ps1 -AppsLocation C:\Development\PowerShell\powerapps -SourceSiteUrl https://contoso.sharepoint.com/sites/dcratdev -SourceUserName ramin.ahmadi@contoso.com -SourcePassword *** -TargetSiteUrl https://raminahmadi.sharepoint.com/sites/rat -TargetUserName ramin.ahmadi@raminahmadi.com -TargetPassword ***

#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Parameters
Param(
    #Location which app packages are located
    [Parameter(Mandatory = $True)]
    [string]$AppsLocation,
    #Source information
    [Parameter(Mandatory = $True)]
    [string]$SourceSiteUrl,
    [Parameter(Mandatory = $True)]
    [string]$SourceUserName,
    [Parameter(Mandatory = $True)]
    [string]$SourcePassword,
    #Target information
    [Parameter(Mandatory = $True)]
    [string]$TargetSiteUrl,
    [Parameter(Mandatory = $True)]
    [string]$TargetUserName,
    [Parameter(Mandatory = $True)]
    [string]$TargetPassword
)

#region Set Global Variables--------------------------------------------------------------------------------------------------

#endregion Set Global Variables----------------------------------------------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Get app from configuration file
[xml]$ConfigurationContent = Get-Content "$($AppsLocation)\configuration.xml"
$AppsDirectory = $Env:SYSTEM_DEFAULTWORKINGDIRECTORY + "\apps\"
$MsappDirectory = $Env:SYSTEM_DEFAULTWORKINGDIRECTORY + "\msapp\"
$NewPackagesDirectory = $Env:SYSTEM_DEFAULTWORKINGDIRECTORY + "\newpackages\"
New-Item -ItemType directory -Path $NewPackagesDirectory
#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Update-AppData() {
    $ConfigurationContent.Apps.App | ForEach-Object {
        $AppName = $_.Name
        $AppId = $_.Id
        Write-Host "App Name: $($AppName)"
        # Extract the package
        Expand-Archive -LiteralPath "$($AppsLocation)\$($AppName).zip" -DestinationPath "$($AppsDirectory)$($AppName)" -Force
        # Rename .msapp file to zip so we can extract it
        Get-ChildItem -Path "$($AppsDirectory)\$($AppName)" -Filter *.msapp  -Recurse| Select-Object -First 1 | Rename-Item -NewName {$_.name -Replace '\.msapp', '.zip'}
        # Extract app contents
        $MsappFile = Get-ChildItem -Path "$($AppsDirectory)$($AppName)" -Filter *.zip  -Recurse| Select-Object -First 1
        Expand-Archive -LiteralPath $MsappFile.FullName -DestinationPath "$($MsappDirectory)\$($AppName)" -Force
        # Update App confige file
        $ConfigFilePath = $AppsDirectory + "$($AppName)\Microsoft.PowerApps\apps\$($AppId)\$($AppId).json"
        Update-JsonFile $ConfigFilePath $_
        $PropertiesFilePath = $MsappDirectory + "$($AppName)\Properties.json"
        Update-JsonFile $PropertiesFilePath $_
        $EntitiesFilePath = $MsappDirectory + "$($AppName)\Entities.json"
        Update-JsonFile $EntitiesFilePath $_
        # Update App Flows
        Update-Flows $_
        # Rename .zip file to .msapp
        Get-ChildItem -Path "$($AppsDirectory)\$($AppName)" -Filter *.zip  -Recurse| Select-Object -First 1 | Rename-Item -NewName {$_.name -Replace '\.zip', '.msapp'}
        # Insert updated files directly to the zip file
        Insert-UpdatedFilesToPackage $AppName
        # Compress the package
        Compress-Archive -Path "$($AppsDirectory)\$($AppName)\*" -DestinationPath "$($NewPackagesDirectory)\$($AppName)"
    }
    # Compress all packages into one zip file
    Compress-Archive -Path "$($NewPackagesDirectory)\*.zip" -DestinationPath "$($NewPackagesDirectory)\PowerAppsForms"
}

function Update-JsonFile($Path, $App) {
    try {
        # Get file content
        $json = Get-Content $Path | Out-String
        # Replace all source url to target url
        $json = $json.Replace($SourceSiteUrl, $TargetSiteUrl)
        # Replace data sources
        $App.DataSources.DataSource | ForEach-Object {
            $sourceListName = $_.ListName
            $sourceListId = Get-ListId $SourceSiteUrl $SourceUserName $SourcePassword $sourceListName
            $targetListId = Get-ListId $TargetSiteUrl $TargetUserName $TargetPassword $sourceListName
            $json = $json.Replace($sourceListId, $targetListId)
        }
        $json = (Convertfrom-Json $json)
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
                $json = (Convertfrom-Json $json)
                $json | ConvertTo-Json -Depth 40 | Set-Content $Path
            }
            catch {
                Write-Host $_.Exception.Message
            }
        }
    }
}

function Insert-UpdatedFilesToPackage($AppName) {
    $MsappFilePath = Get-ChildItem -Path "$($AppsDirectory)$($AppName)\" -Filter *.msapp  -Recurse| Select-Object -First 1    
    $PropertiesFilePath = $MsappDirectory + "$($AppName)\Properties.json"
    $EntitiesFilePath = $MsappDirectory + "$($AppName)\Entities.json"
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

function Get-ListId($SiteUrl, $Username, $Password, $ListName) {
    $pw = $Password
    $securePassword = $pw | ConvertTo-SecureString -AsPlainText -Force
    $SPCred = New-Object system.management.automation.pscredential -ArgumentList $Username, $securePassword
    Connect-PnPOnline -Url $SiteUrl -credential $SPCred
    $List = Get-PnPList -Identity "$($ListName)"
    return $List.Id
}

Update-AppData

#requires
<#
.SYNOPSIS
  This script is used to provide commandlets to update Flow data sets to be used in other tenants

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
.\VSTS-UpdateFlows.ps1 -SourceSiteUrl https://contoso.sharepoint.com/sites/myapps -SourceUserName ramin.ahmadi@contoso.com -SourcePassword **** -TargetSiteUrl https://raminahmadi.sharepoint.com/sites/myapps -TargetUserName ramin@raminahmadi.onmicrosoft.com -TargetPassword **** -FormsLocation "C:\Development\PowerShell\FlowsPackages"

#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Parameters
Param(
    #Location which app packages are located
    [Parameter(Mandatory=$True)]
    [string]$FlowsLocation,
    #Source information
    [Parameter(Mandatory=$True)]
    [string]$SourceSiteUrl,
    [Parameter(Mandatory=$True)]
    [string]$SourceUserName,
    [Parameter(Mandatory=$True)]
    [string]$SourcePassword,
    #Target information
    [Parameter(Mandatory=$True)]
    [string]$TargetSiteUrl,
    [Parameter(Mandatory=$True)]
    [string]$TargetUserName,
    [Parameter(Mandatory=$True)]
    [string]$TargetPassword
)

#region Set Global Variables--------------------------------------------------------------------------------------------------

#endregion Set Global Variables----------------------------------------------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Get app from configuration file
[xml]$ConfigurationContent = Get-Content "$($FlowsLocation)\configuration.xml"
$FlowsDirectory = $Env:SYSTEM_DEFAULTWORKINGDIRECTORY + "\flows\"
$NewPackagesDirectory = $Env:SYSTEM_DEFAULTWORKINGDIRECTORY + "\newflowpackages\"
New-Item -ItemType directory -Path $NewPackagesDirectory
#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Update-AppData(){
    $ConfigurationContent.Flows.Flow | ForEach-Object{
        try{
            $FlowName= $_.Name
            Write-Host "Flow : $($FlowName)"
            # Extract the package
            Expand-Archive -LiteralPath "$($FlowsLocation)\$($FlowName).zip" -DestinationPath "$($FlowsDirectory)$($FlowName)" -Force
    
            # Update App confige file
            $DefinitionFilePath = Get-ChildItem -Path "$($FlowsDirectory)$($FlowName)\" -Filter definition.json -Recurse| Select-Object -First 1
            Update-JsonFile $DefinitionFilePath.FullName $_
            # Compress the package
            Compress-Archive -Path "$($FlowsDirectory)\$($FlowName)\*" -DestinationPath "$($NewPackagesDirectory)\$($FlowName)"
  
        }
       catch{
            Write-Host $_.Exception.Message
       }
    }
    # Compress all packages into one zip file
    Compress-Archive -Path "$($NewPackagesDirectory)\*.zip" -DestinationPath "$($NewPackagesDirectory)\Flows"

}

function Update-JsonFile($Path,$App){
  try{
        # Get file content
        $json = Get-Content $Path | Out-String
        # Replace all source url to target url
        $json=$json.Replace($SourceSiteUrl,$TargetSiteUrl)
        # Replace data sources
        $App.DataSources.DataSource | ForEach-Object{
          $sourceListName =  $_.ListName
          $sourceListId = Get-ListId $SourceSiteUrl $SourceUserName $SourcePassword $sourceListName
          $targetListId = Get-ListId $TargetSiteUrl $TargetUserName $TargetPassword $sourceListName              
          $json=$json.Replace($sourceListId,$targetListId)
        }
        $json=(Convertfrom-Json $json)
        $json | ConvertTo-Json -Depth 40 | Set-Content $Path
    }
    catch{
        Write-Host $_.Exception.Message
    }
}

function Get-ListId($SiteUrl,$Username,$Password,$ListName)
{
  $pw =$Password
  $securePassword = $pw | ConvertTo-SecureString -AsPlainText -Force
  $SPCred = New-Object system.management.automation.pscredential -ArgumentList $Username, $securePassword
  Connect-PnPOnline -Url $SiteUrl -credential $SPCred
  $List = Get-PnPList -Identity "$($ListName)"
  return $List.Id
}
Update-AppData
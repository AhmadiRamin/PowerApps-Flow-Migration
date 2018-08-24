
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
.\VSTS-InsertUpdatedFilesToPackage.ps1 -MsappFilePath package.msapp -EntitiesFilePath entitiesFilePath -PropertiesFilePath propertiesFilePath

#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Parameters
Param(
    [Parameter(Mandatory = $True)]
	[string]$SiteCollectionUrl,
    [Parameter(Mandatory = $True)]
	[string]$AdminAccount,
    [Parameter(Mandatory=$True)]
    [string]$Password
)

#region Set Global Variables--------------------------------------------------------------------------------------------------


#endregion Set Global Variables----------------------------------------------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------

$pw =$Password
$sp = $pw | ConvertTo-SecureString -AsPlainText -Force
$SPCred = New-Object system.management.automation.pscredential -ArgumentList $AdminAccount, $sp

function CheckSiteCollection(){
    Write-Host "Register $($SiteCollectionUrl) as a hub site."
    try{
        Connect-PnPOnline -Url $SiteCollectionUrl -credential $SPCred
        Register-PnPHubSite -Site $SiteCollectionUrl
    }
    catch{
        Write-Host "Site collection doesn't exist in this tenant!"
    } 
    Write-Host "Finished checking site collection"
}

CheckSiteCollection
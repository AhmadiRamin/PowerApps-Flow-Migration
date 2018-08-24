
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
.\VSTS-CheckSiteCollection.ps1 -TenantUrl cielocosta -Prefix rdp -Title RDP -Description "Retail Design Portal" -AdminAccount ramin.ahmadi@cielocosta.com -Password ****

#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Parameters
Param(
    [Parameter(Mandatory = $True)]
    [string]$Prefix,
    [Parameter(Mandatory = $True)]
    [string]$TenantPrefix,
    [Parameter(Mandatory = $True)]
    [string]$Title,
    [Parameter(Mandatory = $False)]
    [string]$Description,
    [Parameter(Mandatory = $True)]
    [string]$AdminAccount,
    [Parameter(Mandatory = $True)]
    [string]$Password
)

#region Set Global Variables--------------------------------------------------------------------------------------------------


#endregion Set Global Variables----------------------------------------------------------------------------------------------

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------

$pw = $Password
$sp = $pw | ConvertTo-SecureString -AsPlainText -Force
$SPCred = New-Object system.management.automation.pscredential -ArgumentList $AdminAccount, $sp

function CheckSiteCollection() {
    $hubParams = @{
        Type        = "CommunicationSite";
        Url         = "https://$TenantPrefix.sharepoint.com/sites/$Prefix";
        Title       = $Title;
        Description = $Description;
        Classification = "classification";
        SiteDesign = "Showcase"
    }

    Write-Host "Check if the site collection exists"
    $site = $null
    try {
        $connection = Connect-PnPOnline -Url $hubParams.Url -Credentials $SPCred -ErrorAction SilentlyContinue -ReturnConnection
        $site = Get-PnPSite -ErrorAction SilentlyContinue -Connection $connection
    }
    catch {
        Write-Host "Site collection doesn't exist in this tenant."
    }
    if ($site -eq $null) {
        # Site does not exist
        Write-Host "Creating $($hubParams.Url)" -ForegroundColor Cyan
        $connection = Connect-PnPOnline -Url "https://$TenantPrefix-admin.sharepoint.com" -Credentials $SPCred -ReturnConnection
        $hubParams.Connection = $connection        
        New-PnPSite @hubParams
    }
    else {
        $connection = Connect-PnPOnline -Url $TenantUrl -Credentials $SPCred -ReturnConnection
    }
    Write-Host "Finished checking site collection"
}

CheckSiteCollection
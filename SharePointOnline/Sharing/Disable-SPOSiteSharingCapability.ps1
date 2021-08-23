#requires -Modules @{ ModuleName="PnP.PowerShell"; ModuleVersion="1.7.0" }

<#
    .Synopsis
        Sets all sites in a SharePoint Online tenant to not allow sharing with guest/external users.
    .DESCRIPTION
        Sets all sites in a SharePoint Online tenant to not allow sharing with guest/external users.  Populate the $exclusions array with a list of URLs that should not have their
        sharing capability disabled.  Skips all sites that have any type of site lock applied.
    .NOTES
        Azure AD App principal requires SharePoint > Application > Sites.FullControl rights
#>

[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12   

$tenant     = $env:O365_TENANT     # "contoso"
$clientId   = $env:O365_CLIENTID   # "2643912e-2b58-4807-8950-5cedd8ee3e8e"
$thumbprint = $env:O365_THUMBPRINT # "90e7de311c18da348c17419dba63b42c2198699a"
$exclusions = @( "https://$tenant.sharepoint.com/sites/teamsite" )

$tenantConnection = Connect-PnPOnline -Url "https://$tenant-admin.sharepoint.com" -ClientId $clientId -Thumbprint $thumbprint -Tenant "$tenant.onmicrosoft.com" -ReturnConnection

$sites = Get-PnPTenantSite -Connection $tenantConnection

foreach( $site in $sites )
{
    if( $exclusions -contains $site.Url )
    {
        Write-Warning "$(Get-Date) - Skipping excluded site $($site.Url)"
        continue
    }

    if( $site.LockState -ne "Unlock" )
    {
        Write-Warning "$(Get-Date) - Skipping locked site $($site.Url)"
        continue
    }

    if( $site.SharingCapability -ne "Disabled" )
    {
        Write-Host "$(Get-Date) - Disabling sharing for $($site.Url)"
        Set-PnPTenantSite -Identity $site.Url -SharingCapability Disabled -Connection $tenantConnection
    }
}

Disconnect-PnPOnline -Connection $tenantConnection
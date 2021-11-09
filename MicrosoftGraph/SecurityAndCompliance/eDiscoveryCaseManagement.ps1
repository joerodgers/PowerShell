[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12   

function Get-GraphAccessToken
{
    [CmdletBinding()]
    Param 
    (
	    [Parameter(Mandatory=$true)]
	    [string]$TenantId,

	    [Parameter(Mandatory=$true)]
	    [string]$ClientId,

	    [Parameter(Mandatory=$true)]
	    [PSCredential]$Credential
    )

    begin
    {
        $scopes = "https://graph.microsoft.com/.default"

        $body = "client_id={0}&client_info=1&scope={1}&grant_type=password&username={2}&password={3}" -f $ClientId, $scopes, $Credential.UserName, $Credential.GetNetworkCredential().Password
    }
    process
    {
        Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body 
    }
    end
    {
    }
}

function New-AdvancedeDiscoveryCase
{
    [CmdletBinding()]
    Param 
    (
	    [Parameter(Mandatory=$true)]
	    [string]$DisplayName,

	    [Parameter(Mandatory=$true)]
	    [string]$AccessToken
    )

    begin
    {
        $uri = "https://graph.microsoft.com/beta/compliance/ediscovery/cases"

        $body = [PSCustomObject] @{ displayName = $DisplayName } | ConvertTo-Json

        $headers = @{ "Authorization" = "Bearer $AccessToken" }  
    }
    process
    {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -Headers $headers     
    }
    end
    {
    }
}

function New-AdvancedeDiscoveryCaseLegalHold
{
    [CmdletBinding()]
    Param 
    (
	    [Parameter(Mandatory=$true)]
	    [string]$CaseId,

	    [Parameter(Mandatory=$true)]
	    [string]$DisplayName,

	    [Parameter(Mandatory=$true)]
	    [string]$AccessToken
    )

    begin
    {
        $uri = "https://graph.microsoft.com/beta/compliance/ediscovery/cases/$CaseId/legalHolds"

        $body = [PSCustomObject] @{ displayName = $DisplayName } | ConvertTo-Json

        $headers = @{ "Authorization" = "Bearer $AccessToken" }  
    }
    process
    {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -Headers $headers     
    }
    end
    {
    }
}

function Add-AdvancedeDiscoveryCaseLegalHoldSiteSource
{
    [CmdletBinding()]
    Param 
    (
	    [Parameter(Mandatory=$true)]
	    [string]$SiteId,

	    [Parameter(Mandatory=$true)]
	    [string]$CaseId,

	    [Parameter(Mandatory=$true)]
	    [string]$LegalHoldId,

	    [Parameter(Mandatory=$true)]
	    [string]$AccessToken
    )

    begin
    {
        $uri = "https://graph.microsoft.com/beta/compliance/ediscovery/cases/$CaseId/legalHolds/$LegalHoldId/siteSources"

        $body = [PSCustomObject] @{ "site@odata.bind" = "https://graph.microsoft.com/v1.0/sites/$SiteId" } | ConvertTo-Json

        $headers = @{ "Authorization" = "Bearer $AccessToken" }  
    }
    process
    {
        Invoke-RestMethod -Method Post -Uri $uri -Body $body -Headers $headers     
    }
    end
    {
    }
}

function Get-AdvancedeDiscoveryCaseLegalHoldSiteSource
{
    [CmdletBinding()]
    Param 
    (
	    [Parameter(Mandatory=$true)]
	    [string]$CaseId,

	    [Parameter(Mandatory=$true)]
	    [string]$LegalHoldId,

	    [Parameter(Mandatory=$true)]
	    [string]$AccessToken
    )

    begin
    {
        $uri = "https://graph.microsoft.com/beta/compliance/ediscovery/cases/$CaseId/legalHolds/$LegalHoldId/siteSources"

        $headers = @{ "Authorization" = "Bearer $AccessToken" }  
    }
    process
    {
        Invoke-RestMethod -Method Get -Uri $uri -Headers $headers | Select-Object -ExpandProperty value
    }
    end
    {
    }
}

$caseId = "55aa312a-aa00-4e37-9791-5997fc7aef0d"

$legalHoldId = "f682c6c9-6ed1-4d73-9047-2b4d672bab39"

$clientId = "3bc20efb-2096-416a-aaf8-04394bd3de01"

$credential = [PSCredential]::new( $env:O365_USERNAME, ($env:O365_SECURE_PASSWORD | ConvertTo-SecureString))

$accessToken = Get-GraphAccessToken -TenantId $env:O365_TENANTID -ClientId $clientId -Credential $credential

Get-AdvancedeDiscoveryCaseLegalHoldSiteSource -CaseId $caseId -LegalHoldId $legalHoldId -AccessToken $accessToken.access_token | FL *
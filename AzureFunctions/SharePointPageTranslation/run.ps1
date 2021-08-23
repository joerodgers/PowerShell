using namespace System.Net

<# 
    Expected POST body JSON

    {
        "SiteUrl"   : "https://contoso.sharepoint.com/sites/teamsite",
        "PageTitle" : "ExamplePage.aspx",
        "Language"  : "es"

    }

#>

param($Request, $TriggerMetadata)

Import-Module -Name "PnP.PowerShell"

function Start-AzureTranslation
{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$Text,

        [Parameter(Mandatory=$true)]
        [string]$Language,

        [Parameter(Mandatory=$true)]
        [string]$TranlatorKey
    )

    $uri = "https://api.cognitive.microsofttranslator.com/translate?api-version=3.0&to={0}&textType=html" -f $Language

    $headers = @{
        'Ocp-Apim-Subscription-Key' = $TranlatorKey
        'Content-type'              = 'application/json'
    }

    # Create JSON array with 1 object for request body
    $textJson = @{ "Text" = $Text } | ConvertTo-Json

    $body = "[$textJson]"

    # Uri for the request includes language code and text type, which is always html for SharePoint text web parts
    
    # Send request for translation and extract translated text
    $results = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body
    $translatedText = $results[0].translations[0].text

    return $translatedText
}

# environment varibles created in function app's config section
$clientId     = $env:O365_CLIENTID
$thumbprint   = $env:O365_THUMBPRINT
$tenantId     = $env:O365_TENANTID
$tenantName   = $env:O365_TENANTNAME
$tranlatorKey = $env:AZURE_TRANSLATORKEY

# validate the request was a POST
if( $Request.Method -ne "POST" )
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode = [HttpStatusCode]::BadRequest })
    return
}

# validate the POST body has the three properties we require for translation
if( [string]::IsNullOrWhiteSpace($Request.Body.SiteURL)  -or 
    [string]::IsNullOrWhiteSpace($Request.Body.Language) -or 
    [string]::IsNullOrWhiteSpace($Request.Body.PageTitle))
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode = [HttpStatusCode]::BadRequest; Body = "Invalid POST body" })
    return
}

$siteUrl   = $Request.Body.SiteURL
$language  = $Request.Body.Language
$pageTitle = $Request.Body.PageTitle

Write-Host "Connecting to $($SiteUrl)"

# connect to tenant site
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId

# stop if the connection to spo failed
if( -not $?)
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode = [HttpStatusCode]::InternalServerError; Body = "Failed to connect to SharePoint Online service." })
    return
}

Write-Host "Reading Page: $language/$pageTitle.aspx"

if( $page = Get-PnPPage -Identity "$language/$pageTitle.aspx" )
{
    Write-Host "Starting translation of content on $($page.Name) to language '$language'"

    $textControls = $page.Controls | Where-Object { $_.Type.Name -eq "PageText" }

    # translate each page text control
    foreach ( $textControl in $textControls )
    {
        # skip any controls with no text
        if( [string]::IsNullOrWhiteSpace($textControl.Text) )
        {
            continue
        }

        $translatedText = Start-AzureTranslation -text $textControl.Text -Language $language -TranlatorKey $tranlatorKey

        if( $translatedText -ne $textControl.Text -and -not [string]::IsNullOrWhiteSpace($translatedText) )
        {
            Write-Host "Completed translation of content in control $($textControl.InstanceId)"

            Set-PnPPageTextPart -Page "$language/$pageTitle.aspx" -InstanceId $textControl.InstanceId -Text $translatedText

            Write-Host "Updated control $($textControl.InstanceId) with translated text"
        }
        else
        {
            Write-Warning "Translation returned no text or no changed text for control $($textControl.InstanceId)"
        }
    }

    # translate the page title, if exists
    if( -not [string]::IsNullOrWhiteSpace($page.PageTitle) )
    {
        $translatedPageTitle = Start-AzureTranslation -Text $page.PageTitle -Language $language -TranlatorKey $tranlatorKey

        if( -not [string]::IsNullOrWhiteSpace($translatedPageTitle) -and $translatedPageTitle -ne $page.PageTitle )
        {
            Write-Host "Updating page title with translated title"
            Set-PnPPage -Identity "$language/$pageTitle.aspx" -Title $translatedPageTitle
        }
    }

    Push-OutputBinding -Name Response -Value ([HttpResponseContext] @{StatusCode = [HttpStatusCode]::OK })
}
else
{
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{ StatusCode = [HttpStatusCode]::InternalServerError; Body = "Could not read page $language/$pageTitle.aspx" })
}

Disconnect-PnPOnline
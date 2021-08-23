<#
    .Synopsis
        Example of using the REST API to read project data and project server data
    .DESCRIPTION
        Users the ROPC credential flow to create a access token with the provided username/password.  Uses that access token to 
        read _api/projectserver and _api/projectdata endpoints
    .NOTES
        Azure AD App principal requires SharePoint > Delegated > Project.Write
                                        SharePoint > Delegated > ProjectWebAppReporting.Read
#>

$tenantId  = '00000000-0000-0000-0000-000000000000'
$tenant    = 'contoso'
$client_id = '00000000-0000-0000-0000-000000000000'
$username  = 'username'
$password  = 'password'
$scopes    = "https://$tenant.sharepoint.com/Project.Write https://$tenant.sharepoint.com/ProjectWebAppReporting.Read"

$body = "client_id=$client_id&client_info=1&scope=$scopes&grant_type=password&username=$username&password=$password"

$token = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body 

Invoke-RestMethod -Uri https://$tenant.sharepoint.com/sites/projectserver1/_api/projectserver/projects -Method Get -Headers @{ Authorization = "Bearer $($token.access_token)"; Accept = "application/json" }

Invoke-RestMethod -Uri https://$tenant.sharepoint.com/sites/projectserver1/_api/projectdata -Method Get -Headers @{ Authorization = "Bearer $($token.access_token)"; Accept = "application/json" }
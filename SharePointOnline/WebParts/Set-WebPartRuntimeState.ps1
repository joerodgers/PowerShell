function Get-SitePage
{
    [CmdletBinding()]
    param 
    (
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [Microsoft.SharePoint.Client.List]
        $List,

        [Parameter(Mandatory=$true)]
        [PnP.PowerShell.Commands.Base.PnPConnection]
        $Connection
    )


    begin
    {
    }
    process
    { 
        $pages = $List | Get-PnPListItem -PageSize 500 -Connection $Connection

        foreach( $page in $pages )
        {
            Get-PnPProperty -ClientObject $page -Property ContentType, File -Connection $Connection | Out-Null

            if( $page.ContentType.Name -eq "Site Page" )
            {
                Get-PnPClientSidePage -Identity $page.File.Name -Connection $Connection
            }
        }
    }
    end
    {
    }    
}

$timestamp = Get-Date -Format FileDateTime # used to timestamp output file name

# Azure AD App Principal credentials
# Requires Permission: SharePoint > Application Permission > Sites.ReadWrite.All or higher  
$clientId   = $env:O365_CLIENTID
$thumbprint = $env:O365_THUMBPRINT
$tenantId   = $env:O365_TENANTID
$tenant     = $env:O365_TENANT

# target web parts
$webPartIds = "544dd15b-cf3c-441b-96da-004d5a8cea1d", # universal YouTube web part id 
              "f6fdf4f8-4a24-437b-a127-32e66a5dd9b4", # universal Twitter web part id
              "46698648-fcd5-41fc-9526-c7f7b2ace919"  # universal Amazon Kindle web part id

# connect to the tenant admin center
$tenantConnection = Connect-PnPOnline -Url "https://$tenant-admin.sharepoint.com" -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -ReturnConnection

# pull all sites in the tenant (excludes onedrive)
$sites = Get-PnPTenantSite -Connection $tenantConnection | Where-Object -Property LockState -eq "Unlock"

# enumerate sites
foreach( $site in $sites )
{
    Write-Host "$(Get-Date) - Processing Site: $($site.Url)"

    # connect to the site collection
    $connection = Connect-PnPOnline -Url $site.Url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -ReturnConnection

    $webs = Get-PnPSubWeb -Recurse -IncludeRootWeb

    foreach( $web in $webs )
    {
        Write-Host "$(Get-Date) - Processing Web: $($web.Url)"

        # connect to the web
        $webConnection = Connect-PnPOnline -Url $web.Url -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId -ReturnConnection

        # get the sites pages library
        $sitePagesLibrary  = Get-PnPList -Identity "SitePages" -Connection $webConnection 

        if( $null -eq $sitePagesLibrary ) { continue }

        # get all the sites pages in the Site Pages library
        $clientSidePages = $sitePagesLibrary | Get-SitePage -Connection $webConnection

        foreach( $clientSidePage in $clientSidePages )
        {
            Write-Host "$(Get-Date) - Processing Page: $($clientSidePage.Name)"

            # pull the web parts from the page
            $webParts = $clientSidePage | Get-PnPPageComponent -Connection $Connection | Where-Object -Property WebPartId -in $WebPartIds

            # enumerate web parts
            foreach( $webpart in $webParts )
            {
                switch( $webpart.WebPartId )
                {
                    "544dd15b-cf3c-441b-96da-004d5a8cea1d"
                    {
                        $webPartType = "YouTube"
                    }
                    "f6fdf4f8-4a24-437b-a127-32e66a5dd9b4"
                    {
                        $webPartType = "Twitter"
                    }
                    "46698648-fcd5-41fc-9526-c7f7b2ace919"
                    {
                        $webPartType = "Kindle"
                    }
                    default
                    {
                        $webPartType = "Unknown"
                    }
                }
                
                $result = [PSCustomObject] @{
                    SiteUrl     = $site.Url
                    WebUrl      = $web.Url
                    FileName    = $sitePage.File.Name
                    Title       = $webpart.Title
                    Description = $webpart.Description
                    WebPartType = $webPartType
                    WebPartId   = $webpart.WebPartId
                    InstanceId  = $webpart.InstanceId
                } 

                # append results to log file to the same folder as the script
                $result | Export-Csv -Path "$PSScriptRoot\WebPartInventory_$timestamp.csv" -NoTypeInformation -Append

                if( $webpart.WebPartId -eq "544dd15b-cf3c-441b-96da-004d5a8cea1d" )
                {
                    # set web part privacy enhanced mode to true

                    <#

                    $object = $webpart.PropertiesJson | ConvertFrom-Json 

                    if( $object.runtimeState.isPrivate -ne $false )
                    {
                        Write-Host "$(Get-Date) - Updating YouTube web part"

                        $object.runtimeState.isPrivate = $true

                        $webpart.PropertiesJson = $object | ConvertTo-Json -Compress
    
                        $null = $clientSidePage.Save()
    
                        $clientSidePage.Publish( "Administrator update to YouTube web part properties." )
                    }
                    else 
                    {
                        Write-Host "$(Get-Date) - Skipping YouTube web part"
                    }

                    #>
                }
            }
        }
    }
}

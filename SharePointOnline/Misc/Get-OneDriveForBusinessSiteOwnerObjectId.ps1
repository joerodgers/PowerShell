#Requires -Module PnP.PowerShell

function Get-OneDriveForBusinessSiteOwnerObjectId
{
    [CmdletBinding()]
    param
    (
        # OneDrive Site Url
        [Parameter(Mandatory=$false)]
        [string]
        $Identity,

        # AAD App Principal Client/Application Id
        [Parameter(Mandatory=$true)]
        [string]
        $ClientId,

        # SharePoint Tenant Name (contoso)
        [Parameter(Mandatory=$true)]
        [string]
        $Tenant,

        # AAD App Principal Client/Application certificate thumbprint
        [Parameter(Mandatory=$true)]
        [string]
        $Thumbprint
    )
    
    begin 
    {
        $Tenant = $Tenant -replace ".onmicrosoft.com", ""

        if( $PSBoundParameters.ContainsKey("Identity") )
        {
            $urls = @($Identity)
        }
        else 
        {
            Write-Verbose "$(Get-Date) - Querying tenant for all OD4B URLs"

            $connection = Connect-PnPOnline -Url "https://$Tenant-admin.sharepoint.com" -ClientId $ClientId -Thumbprint $Thumbprint -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection -Verbose:$false

            $sites = Get-PnPTenantSite -IncludeOneDriveSites -Connection $connection | Where-Object -Property Template -match "SPSPERS" 
        
            Disconnect-PnPOnline -Connection $connection
        }
    }
    process 
    {
        $counter = 1

        foreach( $site in $sites )
        {
            Write-Verbose "$(Get-Date) - $counter/$($sites.Count) - Processing $($site.Url)"

            if( $site.LockState -ne "Unlock" )
            {
                # can't pull the AadObjectId of the owner on a locked site
                [PSCustomObject] @{
                    SiteUrl            = $site.Url
                    LockState          = $site.LockState 
                    UserName           = $site.Owner
                    SharePointObjectId = ""
                    AzureAdObjectId    = ""
                    ObjectIdMismatch   = ""
                }

                continue
            }

            try 
            {
                # connect to the OD4B site
                $connection = Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Thumbprint $Thumbprint -Tenant "$Tenant.onmicrosoft.com" -ReturnConnection -Verbose:$false

                # get the owner and the AadObjectId value from SPO
                $site = Get-PnPSite -Includes Owner, Owner.AadObjectId, LockState
    
                # remove the claims prefix from the login name
                $ownerUserPrincipalName = $site.Owner.LoginName -replace "i\:0\#\.f\|membership\|", ""
    
                $azureAdObjectId  = ""
                $objectIdMismatch = $false

                if( -not [string]::IsNullOrWhiteSpace($ownerUserPrincipalName) -and $null -ne $site )
                {
                    try
                    {
                        # get the owner AadObjectId value from Azure AD
                        $azureAdUser = Get-PnPAzureADUser -Identity $ownerUserPrincipalName -Connection $connection 

                        $azureAdObjectId = $azureAdUser.Id
                        
                        $objectIdMismatch = $azureAdUser.Id -ne $site.Owner.AadObjectId.NameId
                    }
                    catch
                    {
                        $azureAdObjectId  = "User not found"
                        $objectIdMismatch = "Unknown"
                    }
                }

                # result
                [PSCustomObject] @{
                    SiteUrl            = $site.Url
                    LockState          = $site.LockState 
                    UserName           = $ownerUserPrincipalName
                    SharePointObjectId = $site.Owner.AadObjectId.NameId
                    AzureAdObjectId    = $azureAdObjectId
                    ObjectIdMismatch   = $objectIdMismatch
                }
            }
            catch
            {
                Write-Error "Error processing $($site.Url). Error: $_"
            }
            finally
            {
                if( $connection )
                {
                    Disconnect-PnPOnline -Connection $connection
                }
            }

            $counter++
        }
    }
    end
    {
        
    }
}

# requries Azure AD App Principal Permissions
# Application > SharePoint > Sites.FullControl.All
# Application > Graph > User.Read.All

Get-OneDriveForBusinessSiteOwnerObjectId -ClientId $env:O365_CLIENTID -Tenant $env:O365_TENANT -Thumbprint $env:O365_THUMBPRINT | Export-Csv -Path "OneDriveForBusinessPUIDs.csv" -NoTypeInformation
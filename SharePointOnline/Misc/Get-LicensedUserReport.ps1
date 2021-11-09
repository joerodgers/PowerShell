#requires -Modules @{ ModuleName="PnP.PowerShell";         ModuleVersion="1.7.0"  }
#requires -Modules @{ ModuleName="Microsoft.Graph.Groups"; ModuleVersion="1.0.1"  }
#requires -Modules @{ ModuleName="Microsoft.Graph.Users";  ModuleVersion="1.0.1"  }
#requires -Modules @{ ModuleName="ImportExcel";            ModuleVersion="7.3.0"  }

function Get-UserDetail
{
<#
    .Synopsis
    Reports organizational details about the supplied user account 

    .DESCRIPTION
    Reports Active Directory, Exchange, Azure Active Directory and license details for any user provided with a Microsoft F1, E3 or E5 license assigned to their account. 

    .EXAMPLE
    Get-UserDetail -UserPrincipalName "john.doe@contoso.com"

    .EXAMPLE
    "john.doe@contoso.com" | Get-UserDetail

    .EXAMPLE
    Get-UserDetail -UserPrincipalName "john.doe@contoso.com" -IncludeActiveDirectoryProperties -IncludeExchangeProperties

    .EXAMPLE
    "john.doe@contoso.com" | Get-UserDetail -IncludeActiveDirectoryProperties -IncludeExchangeProperties
#>
    [CmdletBinding()]
    param 
    (
        # UserPrincipalName of the user to report on
        [parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [string]
        $UserPrincipalName,

        # Indicates the function should make additional Active Directory calls for additional user properties
        [parameter(Mandatory=$false)]
        [switch]
        $IncludeActiveDirectoryProperties,

        # Indicates the function should make additional Exchange calls for additional user properties
        [parameter(Mandatory=$false)]
        [switch]
        $IncludeExchangeProperties
    )

    begin
    {
        # property reference: https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties
        $properties =   "City",
                        "CompanyName",
                        "Country",
                        "CreatedDateTime",
                        "Department",
                        "DisplayName",
                        "EmployeeType",
                        "ImAddresses",
                        "JobTitle",
                        "LicenseAssignmentStates",
                        "LicenseDetails",
                        "Mail",
                        "MySite",
                        "officeLocation",
                        "onPremisesExtensionAttributes",
                        "onPremisesLastSyncDateTime",
                        "State",
                        "StreetAddress",
                        "UserPrincipalName",
                        "UserType"
        
        # get all the skus and their names
        $subscribedSkus = Get-MgSubscribedSku

        $groupCache = @()
    }    
    process
    {
        Write-Verbose "$(Get-Date) - Procesing user: '$UserPrincipalName'"

        Write-Verbose "$(Get-Date) - Querying Microsoft Graph API"
        $graphProperties = Get-MgUser -UserId $UserPrincipalName -Property $properties

        if( $IncludeActiveDirectoryProperties.IsPresent )
        {
            Write-Verbose "$(Get-Date) - Querying Active Directory"
            $activeDirectoryUserProperties = Get-ADUser -Filter { UserPrincipalName -eq $UserPrincipalName } -Properties * 
        }

        if( $IncludeExchangeProperties.IsPresent )
        {
            Write-Verbose "$(Get-Date) - Querying Exchange"
            $exchangeMailboxProperties = Get-Mailbox -Identity $UserPrincipalName 

            Write-Verbose "$(Get-Date) - Querying Exchange Statistics"
            $exchangeMailboxStatsProperties = Get-MailboxStatistics -Identity $UserPrincipalName 
        }

        if( -not [string]::IsNullOrWhiteSpace($graphProperties.MySite) )
        {
            Write-Verbose "$(Get-Date) - Querying O365 Tenant"
            $onedriveProperties = Get-PnPTenantSite -Identity $graphProperties.MySite
        }

        # pull display name of the license skuId values
        $assignedSkus = $subscribedSkus | Where-Object -Property SkuId -in $graphProperties.LicenseAssignmentStates.SkuId 

        foreach( $assignedSku in $assignedSkus )
        {
            # skip any licenses that are not "SPE_F1", "SPE_E3", or "SPE_E5" 
            if( @( "SPE_F1", "SPE_E3", "SPE_E5" ) -notcontains $assignedSku.SkuPartNumber )
            {
                Write-Warning "User: $UserPrincipalName - Skipping detailed entry for license type: $($assignedSku.SkuPartNumber)"
                continue
            }

            Write-Verbose "$(Get-Date) - Processing user license: $($assignedSku.SkuPartNumber)"

            # pull how this license was applied to this user
            $licenseAssignment = $graphProperties.LicenseAssignmentStates | Where-Object -Property SkuId -eq $assignedSku.SkuId | Select-Object AssignedByGroup, Error, SkuId, State

            # default values for non-group based licensing
            $licenseGroupDisplayName = "None"
            $licenseAssignmentType   = "Direct"

            if ( $licenseAssignment.AssignedByGroup )
            {
                $licenseAssignmentType = "Inherited"

                if( $cachedEntry = $groupCache | Where-Object -Property Id -eq $licenseAssignment.AssignedByGroup )
                {
                    Write-Verbose "$(Get-Date) - Found group $($cachedEntry.Id) in cache"

                    $licenseGroupDisplayName = $cachedEntry.DisplayName
                }
                else 
                {
                    Write-Verbose "$(Get-Date) - Querying Azure AD group id; '$($licenseAssignment.AssignedByGroup)'"

                    $group = Get-MgGroup -GroupId $licenseAssignment.AssignedByGroup | Select-Object Id, DisplayName

                    $licenseGroupDisplayName = $group.DisplayName

                    $groupCache += $group
                }
            }

            # core properties
            $result = [PSCustomObject] @{
                UserPrincipalName     = $graphProperties.UserPrincipalName
                DisplayName           = $graphProperties.DisplayName
                Office                = $graphProperties.OfficeLocation
                JobTitle              = $graphProperties.JobTitle
                WhenCreated           = $graphProperties.CreatedDateTime
                Department            = $graphProperties.Department
                SIPAddress            = $graphProperties.imAddresses -join ", "
                EmployeeType          = $graphProperties.onPremisesExtensionAttributes["extension_e0ce4f3735d249d7a1042e5f6fedc958_employeeType"]
                LastDirSyncTime       = $graphProperties.onPremisesLastSyncDateTime
                LicenseGroup          = $assignedSku.SkuPartNumber # SPE_F1, SPE_E3, SPE_E5
                GroupsAssignedFrom    = $licenseGroupDisplayName
                LicenseAssignment     = $licenseAssignmentType
                AllLicenses           = $assignedSkus.SkuPartNumber -join ", "
                OD4BSiteUrl           = $onedriveProperties.Url
                OD4BSiteStatus        = $onedriveProperties.Status
                OD4BSiteState         = $onedriveProperties.LockState
                OD4BSiteStorageUsed   = $onedriveProperties.StorageUsageCurrent # MB
            }

            # add all active directory properties
            if( $IncludeActiveDirectoryProperties.IsPresent )
            {
                $result | Add-Member -MemberType NoteProperty -Name "Description"   -Value $activeDirectoryUserProperties.Description
                $result | Add-Member -MemberType NoteProperty -Name "GradeLevel"    -Value $activeDirectoryUserProperties."contoso-comGradeLevel"
                $result | Add-Member -MemberType NoteProperty -Name "GroupName"     -Value $activeDirectoryUserProperties."contoso-comGrpName"
                $result | Add-Member -MemberType NoteProperty -Name "SubGroupName"  -Value $activeDirectoryUserProperties."contoso-comSubGrpName"
                $result | Add-Member -MemberType NoteProperty -Name "VendorName"    -Value $activeDirectoryUserProperties."contoso-comVendorName"
                $result | Add-Member -MemberType NoteProperty -Name "WorkerType"    -Value $activeDirectoryUserProperties."contoso-comWorkerType"
                $result | Add-Member -MemberType NoteProperty -Name "Country"       -Value $activeDirectoryUserProperties."contoso-comWorkCityName"
                $result | Add-Member -MemberType NoteProperty -Name "State"         -Value $activeDirectoryUserProperties."contoso-comWorkStateName"
                $result | Add-Member -MemberType NoteProperty -Name "City"          -Value $activeDirectoryUserProperties."contoso-comWorkCityName"
                $result | Add-Member -MemberType NoteProperty -Name "StreetAddress" -Value $activeDirectoryUserProperties.StreetAddress
                $result | Add-Member -MemberType NoteProperty -Name "OfficeName"    -Value $activeDirectoryUserProperties.PhysicalDeliveryOfficeName
            }

            # add all exchange properties
            if( $IncludeExchangeProperties.IsPresent )
            {
                $result | Add-Member -MemberType NoteProperty -Name "RecipientTypeDetails"  -Value $exchangeMailboxProperties.RecipientTypeDetails
                $result | Add-Member -MemberType NoteProperty -Name "ProhibitSendQuota"     -Value $exchangeMailboxProperties.LitigationHoldEnabled
                $result | Add-Member -MemberType NoteProperty -Name "LitigationHoldEnabled" -Value $exchangeMailboxProperties.LitigationHoldEnabled
                $result | Add-Member -MemberType NoteProperty -Name "LitigationHoldDate"    -Value $exchangeMailboxProperties.LitigationHoldDate
                $result | Add-Member -MemberType NoteProperty -Name "MailboxSize"           -Value $exchangeMailboxStatsProperties.TotalItemSize
                $result | Add-Member -MemberType NoteProperty -Name "LastLogon"             -Value $exchangeMailboxStatsProperties.LastLogonTime
            }
            
            $result
        }

    }
    end
    {
    }   
}

<#
  ***  Required Azure AD App Principal Permissions ***
    
    Microsoft Graph > Application Permissions > 
        User.Read.All
        Files.Read.All
        Group.Read.All
        Directory.Read.All

    SharePoint > Application Permissions > 
        Sites.ReadWrite.All
#>
    
$clientId     = $env:O365_CLIENTID
$thumbprint   = $env:O365_THUMBPRINT
$tenantId     = $env:O365_TENANTID
$tenant       = "contoso"

$timestamp = Get-Date -Format FileDateTime

$tempFilePath = Join-Path -Path $PSScriptRoot -ChildPath "dlur-$timestamp.xlsx"

# connect to services 

    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint -ForceRefresh | Out-Null

    Connect-PnPOnline -Url "https://$tenant-admin.sharepoint.com" -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId

<#
    # https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

    # example filter to get all user with any license applied
    Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/$count ne 0' -CountVariable "licenseUserCount" -ConsistencyLevel "eventual"

    # example filter to get any user with Microsoft 365 E5 (SPE_E5) license applied
    Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)'

    # example filter to get any user with Microsoft 365 E3 (SPE_E3) license applied
    Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 05e9a617-0261-4cee-bb44-138d3ef5d965)'

    # example filter to get any user with Microsoft 365 F1 (SPE_F1) license applied
    Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 6fd2c87f-b296-42f0-b197-1e91e994b900)'
#>

# get all users that match the filter

    # $licensedUsers = Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 06ebc4ee-1bb5-47dd-8120-11324bc54e06)'
    # $licensedUsers = Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 05e9a617-0261-4cee-bb44-138d3ef5d965)'
    # $licensedUsers = Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/any(x:x/skuId eq 66b55226-6b4f-492c-910c-a3b7a3c9d993)' 
    $licensedUsers = Get-MgUser -All -Property UserPrincipalName, AssignedLicenses -Filter 'assignedLicenses/$count ne 0' -CountVariable "licenseUserCount" -ConsistencyLevel "eventual"


# pull the details for all matching users

    $licensedUsersDetails = @($licensedUsers.UserPrincipalName | Get-UserDetail -Verbose)

    if( $licensedUsersDetails.Count -eq 0 ) { }

# generate an Excel workbook with a tab foreach LicenseGroup (SPE_F1, SPE_E3, SPE_E5)

    $licensedUsersDetails | Where-Object -Property LicenseGroup -eq "SPE_F1"  | Export-Excel -Path $tempFilePath -Worksheet "F1"
    $licensedUsersDetails | Where-Object -Property LicenseGroup -eq "SPE_E3"  | Export-Excel -Path $tempFilePath -Worksheet "E3"
    $licensedUsersDetails | Where-Object -Property LicenseGroup -eq "SPE_E5"  | Export-Excel -Path $tempFilePath -Worksheet "E5"


# upload Excel workbook to SPO

    if( Test-Path -Path $tempFilePath -PathType Leaf )
    {
        Connect-PnPOnline -Url "https://tenant-my.sharepoint.com/personal/user_tenant_com" -ClientId $clientId -Thumbprint $thumbprint -Tenant $tenantId

        Add-PnPFile -Path $tempFilePath -Folder "Documents"
    
        if( $? ) { Remove-Item -Path $tempFilePath -Force -ErrorAction SilentlyContinue }
    }


# disconnect from services

    Disconnect-MgGraph
    Disconnect-PnPOnline


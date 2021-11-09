Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null


function Invoke-FakeProfileUpdate
{
    [CmdletBinding(DefaultParameterSetName='Identity')]
    param
    (
        [Parameter(Mandatory=$true,ParameterSetName="Identity")]
        [string]$UserName,

        [Parameter(Mandatory=$true,ParameterSetName="Domain")]
        [string]$EmailDomain
    )
    
    begin
    {
        $context    = Get-SPSite -Limit 1 -WarningAction SilentlyContinue -Verbose:$false | Get-SPServiceContext -Verbose:$false
        
        $profileMgr = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($context,$true)
        
        $profiles = @()
    }
    process
    {
        if( $PSCmdlet.ParameterSetName -eq "Domain" )
        {
            foreach( $profile in $profileMgr.GetEnumerator() )
            {
                if( $profile["WorkEmail"] -match $EmailDomain)
                {
                    $profiles += $profile
                }
            }
        }
        else
        {
            if( $profile = $profileMgr.GetUserProfile( $UserName ) )
            {
                $profiles += $profile
            }
        }

        foreach( $profile in $profiles )
        {
            Write-Verbose "Updating User Profile: $($profile["WorkEmail"])"

            $profile.DisplayName = $profile.DisplayName + "_TEMP"
            $profile.Commit()

            $profile.DisplayName = $profile.DisplayName.TrimEnd("_TEMP")
            $profile.Commit()
        }
    }
    end
    {
    }
}

Invoke-FakeProfileUpdate -UserName "contoso\alans" -Verbose

Invoke-FakeProfileUpdate -EmailDomain "contoso.com"  -Verbose

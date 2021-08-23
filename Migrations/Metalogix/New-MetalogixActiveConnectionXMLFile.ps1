$assemblyPath = "C:\Program Files (x86)\Metalogix\Content Matrix Consoles\SharePoint Edition\Metalogix.Core.dll"

$tenantAdminSiteUrl = "https://contoso-admin.sharepoint.com"

$migrationAccounts = @{
    'john.doe@contoso.com' = 'pass@word1'
    'jane.doe@contoso.com' = 'pass@word2'
}

Add-Type -Path $assemblyPath

$assembly = [System.Reflection.Assembly]::LoadFile($assemblyPath)

$lineFormat = '<Connection NodeType="Metalogix.SharePoint.SPTenant, Metalogix.SharePoint, Version=9.1.0.1, Culture=neutral, PublicKeyToken=3b240fac3e39fc03" SharePointVersion="" UnderlyingAdapterType="" ShowAllSites="True" AdapterType="CSOM" Url="{0}" UserName="{1}" SavePassword="True" Password="{2}" ReadOnly="False" AuthenticationType="Metalogix.SharePoint.Adapters.CSOM2013.Authentication.Office365StandADFSInitializer" IsOAuthAuthentication="False"><Proxy Url="" IsProxyImportedFromBrowser="True"/></Connection>'

$xml = New-Object System.Text.StringBuilder

$null = $xml.AppendLine("<ConnectionCollection>")

foreach( $migrationAccount in $migrationAccounts.GetEnumerator() )
{
    $username = $migrationAccount.Key
    $password = $migrationAccount.Value

    $securepassword = ConvertTo-SecureString -String $password -AsPlainText -Force

    if( $assembly.GetName().Version -lt "9.2" )
    {
        $encryptedPassword = [Metalogix.Cryptography]::EncryptText( $securepassword, [Metalogix.Cryptography+ProtectionScope]::CurrentUser, $null )
    }
    elseif( $assembly.GetName().Version -ge "9.2" )
    {
        $cryptographyService = New-Object Metalogix.CryptographyService 
        $encryptedPassword = $cryptographyService.EncryptText( $securepassword, [Metalogix.ProtectionScope]::CurrentUser, $null )
    }

    $null = $xml.AppendFormat( $lineFormat, $tenantAdminSiteUrl, $username, $encryptedPassword )
    $null = $xml.AppendLine()
}

$null = $xml.AppendLine("</ConnectionCollection>")

$xml.ToString() | Set-Content "ActiveConnections.xml"

using module ..\ConfigurationHandler\ConfigurationHandler.psm1
using module ..\KeyVaultHandler\KeyVaultHandler.psm1

enum ConnectType
{
    ClientSecret
    Certificate
    CertificateUserPassword
    CertificatePath
    Password
    KeyVaultSecret
    KeyVaultCertificate
    KeyVaultCertificateEx
}

class SharePointConnectHandler
{
    <########## public members ##########>

    <########## hidden/private members ##########>

    hidden [KeyVaultHandler] $_keyVault
    hidden [object] $_connection

    <########## constructors ##########>

    SharePointConnectHandler([string] $Url)
    {
        Write-Host "Connect to SharePoint site with predefined client id and certificate from Key Vault"

        $tenantId = [ConfigurationHandler]::GetTenantId()
        $keyVaultName = [ConfigurationHandler]::GetKeyVaultName()

        $this._keyVault = [KeyVaultHandler]::new($keyVaultName)

        $clientId = $this._keyVault.GetKey("SP-Connect")
        $certificate = $this._keyVault.GetCertificateAsBase64("SP-Certificate")

        try 
        {
            $this._connection = Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId -CertificateBase64Encoded $certificate -ReturnConnection -ErrorAction SilentlyContinue
        }
        catch 
        {
            Write-Host "Error occured in SharePointHandler constructor"
            Write-Host $_
        }
    }

    SharePointConnectHandler([string] $Url, [string] $AccessToken)
    {
        Write-Host "Connect to SharePoint site with access token"
        
        try 
        {
            $this._connection = Connect-PnPOnline -Url $Url -AccessToken $accessToken -ReturnConnection -ErrorAction SilentlyContinue
        }
        catch 
        {
            Write-Host "Error occured in SharePointHandler constructor"
            Write-Host $_
        }
    }

    SharePointConnectHandler([string] $Url, [PSCredential] $Credentials)
    {
        Write-Host "Connect to SharePoint site with credentials"
        
        try 
        {
            $this._connection = Connect-PnPOnline -Url $Url -Credential $Credentials -ReturnConnection -ErrorAction SilentlyContinue
        }
        catch 
        {
            Write-Host "Error occured in SharePointHandler constructor"
            Write-Host $_
        }
    }

    SharePointConnectHandler([string] $Url, [ConnectType] $ConnectType, [string] $UserClient, [string] $SecretValue, [string] $User = $null, [string] $Password = $null)
    {
        switch ($ConnectType)
        {
            ClientSecret
            {
                Write-Host "Connect to SharePoint site with client id and secret"
        
                try 
                {
                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $UserClient -ClientSecret $SecretValue -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - Client/Secret Auth"
                    Write-Host $_
                }
            }

            Certificate
            {
                Write-Host "Connect to SharePoint site with client id and certificate from Key Vault"
        
                $tenantId = [ConfigurationHandler]::GetTenantId()
                $keyVaultName = [ConfigurationHandler]::GetKeyVaultName()
        
                $this._keyVault = [KeyVaultHandler]::new($keyVaultName)
        
                $clientId = $this._keyVault.GetSecret($UserClient)
                $certificate = $this._keyVault.GetCertificateAsBase64($SecretValue)
        
                try 
                {
                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId -CertificateBase64Encoded $certificate -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - Certificate Auth"
                    Write-Host $_
                }
            }

            CertificateUserPassword
            {
                Write-Host "Connect to SharePoint site with client id and certificate from Key Vault (connect to Key Vault with username and password)"

                $tenantId = [ConfigurationHandler]::GetTenantId()
                $keyVaultName = [ConfigurationHandler]::GetKeyVaultName()
        
                $this._keyVault = [KeyVaultHandler]::new($keyVaultName, $User, $Password)
        
                $clientId = $this._keyVault.GetSecret($UserClient)
                $certificate = $this._keyVault.GetCertificateAsBase64($SecretValue)
        
                try 
                {
                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId -CertificateBase64Encoded $certificate -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - Certificate Auth"
                    Write-Host $_
                }
            }

            CertificatePath
            {
                Write-Host "Connect to SharePoint site with client id, certificate and password for certificate"

                $tenantId = [ConfigurationHandler]::GetTenantId()
        
                $securePassword = ConvertTo-SecureString -String $Password -AsPlainText

                try 
                {
                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $UserClient -Tenant $tenantId -CertificatePath $SecretValue -CertificatePassword $securePassword -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - Certificate Auth"
                    Write-Host $_
                }
            }

            Password
            {
                Write-Host "Connect to SharePoint site with username and password"

                try
                {
                    $securePassword = ConvertTo-SecureString -String $SecretValue -AsPlainText

                    $credentials = New-Object System.Management.Automation.PSCredential($UserClient, $securePassword)

                    $this._connection = Connect-PnPOnline -Url $Url -Credentials $credentials -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch
                {
                    Write-Host "Error occured in SharePointHandler constructor - Password Auth"
                    Write-Host $_
                }
            }

            KeyVaultSecret
            {

            }

            KeyVaultCertificate
            {
                Write-Host "Connect to SharePoint site with client id and certificate from Key Vault"

                $tenantId = [ConfigurationHandler]::GetTenantId()
                $keyVaultName = [ConfigurationHandler]::GetKeyVaultName()
        
                $this._keyVault = [KeyVaultHandler]::new($keyVaultName)
        
                $clientId = $this._keyVault.GetSecret($UserClient)
                $certificate = $this._keyVault.GetCertificateAsBase64($SecretValue)
        
                try 
                {
                    Write-Host "Connect to SharePoint via KeyVaultCertificate Auth"
                    # Write-Host "Connect-PnPOnline -Url '$Url' -ClientId '$clientId' -Tenant '$tenantId' -CertificateBase64Encoded '$certificate' -ReturnConnection -ErrorAction SilentlyContinue"

                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId -CertificateBase64Encoded $certificate -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - KeyVaultCertificate Auth"
                    Write-Host $_
                }
            }

            KeyVaultCertificateEx
            {
                Write-Host "Connect to SharePoint site with client id and certificate from Key Vault (connect to Key Vault with username and password)"

                $tenantId = [ConfigurationHandler]::GetTenantId()
                $keyVaultName = [ConfigurationHandler]::GetKeyVaultName()
        
                $this._keyVault = [KeyVaultHandler]::new($keyVaultName, $User, $Password)
        
                $clientId = $this._keyVault.GetSecret($UserClient)
                $certificate = $this._keyVault.GetCertificateAsBase64($SecretValue)
        
                try 
                {
                    Write-Host "Connect to SharePoint via KeyVaultCertificate Auth extended"
                    
                    $this._connection = Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId -CertificateBase64Encoded $certificate -ReturnConnection -ErrorAction SilentlyContinue
                }
                catch 
                {
                    Write-Host "Error occured in SharePointHandler constructor - KeyVaultCertificate Auth extended"
                    Write-Host $_
                }
            }
        }
    }

    <########## public methods ##########>

    [bool] IsConnected()
    {
        $result = $false

        try 
        {
            $web = Get-PnPWeb -Connection $this._connection
            
            $result = ($null -ne $web)
        }
        catch 
        {
            <#Do this if a terminating exception happens#>
        }

        return $result
    }

    [bool] IsRootWeb()
    {
        $site = Get-PnPSite -Includes ServerRelativeUrl -Connection $this._connection
        $web = Get-PnPWeb -Includes ServerRelativeUrl -Connection $this._connection

        return ($site.ServerRelativeUrl -eq $web.ServerRelativeUrl)
    }

    [object] GetConnection()
    {
        return $this._connection
    }

    [bool] HasKeyVault()
    {
        return ($null -ne $this._keyVault)
    }

    [KeyVaultHandler] GetKeyVault()
    {
        return $this._keyVault
    }

    <########## hidden/private methods ##########>

    [Hashtable] _GetListItems([string] $Listname, [string] $PagingInfo)
    {
        return $this._GetListItems($ListName, $PagingInfo, "", "100")
    }
    
    [Hashtable] _GetListItems([string] $Listname, [string] $PagingInfo, [string] $CamlWhere)
    {
        return $this._GetListItems($ListName, $PagingInfo, $CamlWhere, "100")
    }
    
    [Hashtable] _GetListItems([string] $Listname, [string] $PagingInfo, [string] $CamlWhere = "", [string] $PageSize = "100")
    {
        <#
        param (
            [Parameter(Mandatory)]
            [string] $Listname,
            [string] $CamlWhere = "",
            [string] $PagingInfo,
            [string] $PageSize = "100",
            $Connection
        )
        #>
        $result = @{
            "Items" = $null
            "PagingInfo" = ""
        }
    
        if ($CamlWhere -eq "")
        {
            $camlQuery = @"
                <View Scope='RecursiveAll'>
                    <Query>
                        <Where>
                            <Gt>
                                <FieldRef Name='ID' />
                                <Value Type='Counter'>0</Value>
                            </Gt>
                        </Where>
                    </Query>
                    <RowLimit Paged='TRUE'>$pageSize</RowLimit>
                </View>
"@
        }
        else 
        {
            $camlQuery = @"
                <View Scope='RecursiveAll'>
                    <Query>$CamlWhere</Query>
                    <RowLimit Paged='TRUE'>$pageSize</RowLimit>
                </View>
"@
        }
        
        $ctx = Get-PnPContext -Connection $this._Connection
    
        $list = Get-PnPList -Identity $listname -Connection $this._Connection
    
        [Microsoft.SharePoint.Client.ListItemCollectionPosition] $collectionPosition = $null
    
        if ($PagingInfo -ne "")
        {
            [Microsoft.SharePoint.Client.ListItemCollectionPosition] $collectionPosition = [Microsoft.SharePoint.Client.ListItemCollectionPosition]::new()
    
            $collectionPosition.PagingInfo = $PagingInfo
        }
    
        [Microsoft.SharePoint.Client.CamlQuery] $query = [Microsoft.SharePoint.Client.CamlQuery]::new()
    
        $query.AllowIncrementalResults = $true
        $query.ViewXml = $camlQuery
    
        $query.ListItemCollectionPosition = $collectionPosition
    
        $items = $list.GetItems($query)
        $ctx.Load($items)
        Invoke-PnPQuery -Connection $this._Connection
    
        if ($null -ne $items)
        {
            $collectionPosition = $items.ListItemCollectionPosition
    
            if ($items.ServerObjectIsNull -eq $false)
            {
                $result = @{
                    "Items" = $items
                    "PagingInfo" = $collectionPosition.PagingInfo
                }
            }
        }
    
        return $result
    }
   
}

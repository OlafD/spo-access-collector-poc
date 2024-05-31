using module ..\ConfigurationHandler\ConfigurationHandler.psm1

class PowerShellAzConnector
{
    <########## public members ##########>

    static [object] $AzConnection

    <########## hidden/private members ##########>

    static hidden [PowerShellAzConnector] $_instance

    <########## constructors ##########>

    PowerShellAzConnector()
    {
        # don't remove this constructor, even when it's empty
    }

    <########## public methods ##########>

    static [PowerShellAzConnector] GetInstance()
    {
        if ($null -eq [PowerShellAzConnector]::_instance)
        {
            $subscriptionId = [ConfigurationHandler]::GetSubscriptionId()

            if ([ConfigurationHandler]::IsDevelopmentMode() -eq $false)
            # if ($UseManagedIdentity -eq $true)
            {
                Write-Host "Connect to Az with managed identity"

                $tenantId = [ConfigurationHandler]::GetTenantId()
                
                [PowerShellAzConnector]::AzConnection = Connect-AzAccount -Identity -TenantId $tenantId -Subscription $subscriptionId
            }
            else 
            {
                Write-Host "Connect to Az with service principal"
    
                $tenantId = [ConfigurationHandler]::GetTenantId()
                $username = [ConfigurationHandler]::GetServicePrincipal()
                $password = [ConfigurationHandler]::GetServicePrincipalSecret()
    
                $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
                $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePassword
    
                [PowerShellAzConnector]::AzConnection = Connect-AzAccount -ServicePrincipal -Credential $Credentials -TenantId $tenantId -Subscription $subscriptionId
            }

            [PowerShellAzConnector]::_instance = [PowerShellAzConnector]::new()
        }

        return [PowerShellAzConnector]::_instance
    }

    static [PowerShellAzConnector] GetInstance($Credentials)
    {
        if ($null -eq [PowerShellAzConnector]::_instance)
        {
            Write-Host "Connect to Az with credentials"

            $subscriptionId = [ConfigurationHandler]::GetSubscriptionId()

            [PowerShellAzConnector]::AzConnection = Connect-AzAccount -Credential $Credentials -Subscription $subscriptionId
            # [PowerShellAzConnector]::_instance = Connect-AzAccount -Credential $Credentials -Subscription $subscriptionId

            [PowerShellAzConnector]::_instance = [PowerShellAzConnector]::new()
        }

        return [PowerShellAzConnector]::_instance
    }

    static [PowerShellAzConnector] GetInstance([string] $Username, [string] $Password)
    {
        if ($null -eq [PowerShellAzConnector]::_instance)
        {
            Write-Host "Connect to Az with username and password"

            $subscriptionId = [ConfigurationHandler]::GetSubscriptionId()
    
            $securePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    
            $credentials = New-Object System.Management.Automation.PSCredential -ArgumentList ($Username, $securePassword)
    
            [PowerShellAzConnector]::AzConnection = Connect-AzAccount -Credential $credentials -Subscription $subscriptionId
            # [PowerShellAzConnector]::_instance = Connect-AzAccount -Credential $credentials -Subscription $subscriptionId

            [PowerShellAzConnector]::_instance = [PowerShellAzConnector]::new()
        }

        return [PowerShellAzConnector]::_instance
    }

    <########## hidden/private methods ##########>

}

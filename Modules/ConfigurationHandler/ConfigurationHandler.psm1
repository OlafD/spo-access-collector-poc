
class ConfigurationHandler
{
    <########## public members ##########>

    <########## hidden/private members ##########>

    <########## constructors ##########>

    ConfigurationHandler()
    {
        # don't remove this constructor, even when it's empty
    }

    <########## public methods ##########>

    # methods to read from the configuration of the function app (environment variables in a dev environment)

    static [string] GetEmailReplyToAddress()
    {
        $value = (Get-Item Env:EmailReplyToAddress -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetDefaultRecipient()
    {
        $value = (Get-Item Env:DefaultRecipient -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetRecipients()
    {
        $value = (Get-Item Env:Recipients -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetTenant()
    {
        $value = (Get-Item Env:Tenant -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetTenantFqdn()
    {
        $value = (Get-Item Env:TenantFqdn -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetTenantId()
    {
        $value = (Get-Item Env:TenantId -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetSubscriptionId()
    {
        $value = (Get-Item Env:SubscriptionId -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKeyVaultName()
    {
        $value = (Get-Item Env:KeyVaultName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVGraphConnectIdName()
    {
        $value = (Get-Item Env:KVGraphConnectIdName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVGraphConnectCertName()
    {
        $value = (Get-Item Env:KVGraphConnectCertName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVSharePointConnectIdName()
    {
        $value = (Get-Item Env:KVSharePointConnectIdName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVSharePointConnectCertName()
    {
        $value = (Get-Item Env:KVSharePointConnectCertName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetSiteUrl()
    {
        $value = (Get-Item Env:SiteUrl -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetServicePrincipal()
    {
        $value = (Get-Item Env:ServicePrincipal -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetServicePrincipalSecret()
    {
        $value = (Get-Item Env:ServicePrincipalSecret -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetSmtpMailRelay()
    {
        $value = (Get-Item Env:SmtpMailRelay -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [int] GetSmtpMailRelayPort()
    {
        $result = 25

        $value = [int](Get-Item Env:SmtpMailRelayPort -ErrorAction SilentlyContinue).Value

        if ($value -ne 0)
        {
            $result = $value
        }

        return $result

    }

    static [string] GetSmtpMailRelaySender()
    {
        $value = (Get-Item Env:SmtpMailRelaySender -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVSmtpMailRelayUserName()
    {
        $value = (Get-Item Env:KVSmtpMailRelayUserName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [string] GetKVSmtpMailRelayPasswordName()
    {
        $value = (Get-Item Env:KVSmtpMailRelayPasswordName -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $value = ""
        }

        return $value
    }

    static [bool] IsDevelopmentMode()
    {
        $result = $false

        $value = (Get-Item Env:DevelopmentMode -ErrorAction SilentlyContinue).Value

        if ($null -eq $value)
        {
            $result = $false
        }
        elseif (($value.ToLower() -eq "true") -or ($value.ToLower() -eq "yes") -or ($value -eq "1"))
        {
            $result = $true
        }

        return $result
    }

    <########## hidden/private methods ##########>

}

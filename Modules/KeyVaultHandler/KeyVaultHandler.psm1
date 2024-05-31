using module ..\PowerShellAzConnector\PowerShellAzConnector.psm1

class KeyVaultHandler
{
    <########## public members ##########>

    <########## hidden/private members ##########>
    hidden [string] $_vaultName

    <########## constructors ##########>

    KeyVaultHandler([string] $VaultName)
    {
        $this._vaultName = $VaultName

        [PowerShellAzConnector]::GetInstance()
    }

    KeyVaultHandler([string] $VaultName, [string] $Username, [string] $Password)
    {
        $this._vaultName = $VaultName

        [PowerShellAzConnector]::GetInstance($Username, $Password)
    }

    <########## public methods ##########>

    [object] GetKey([string] $KeyId)
    {
        $value = Get-AzKeyVaultKey -VaultName $this._vaultName -Name $KeyId

        return $value
    }

    [object] GetSecret([string] $SecretId)
    {
        Write-Host "Get secret for $SecretId"

        $valueRead = Get-AzKeyVaultSecret -VaultName $this._vaultName -Name $SecretId

        $value = ConvertFrom-SecureString -SecureString $valueRead.SecretValue -AsPlainText

        return $value
    }

    [string] GetCertificateAsBase64([string] $SecretId)
    {
        Write-Host "Get certificate as base64 for $SecretId"

        $certificate = ""

        try
        {
            $certificate = Get-AzKeyVaultSecret -VaultName $this._vaultName -Name $SecretId -AsPlainText
        }
        catch 
        {
            Write-Host "Exception in GetCertificateAsBase64"
            Write-Host $_
        }

        return $certificate
    }

    [System.Security.Cryptography.X509Certificates.X509Certificate2] GetCertificateAsX509([string] $SecretId)
    {
        Write-Host "Get certificate as X509 for $SecretId"

        $x509Certificate = $null

        try 
        {
            $certificate = Get-AzKeyVaultSecret -VaultName $this._vaultName -Name $SecretId -AsPlainText
            $nativeCertificate = [System.Convert]::FromBase64String($certificate)
            # $x509Certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($nativeCertificate)
            $x509Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $nativeCertificate, "", "Exportable,PersistKeySet"

            <#
            # with the below code an exception "The certificate certificate does not have a private key" is thrown
            $certificateFromKeyVault = Get-AzKeyVaultCertificate -VaultName $this._vaultName -Name $SecretId
            $x509Certificate = $certificateFromKeyVault.Certificate
            #>
        }
        catch 
        {
            Write-Host "Exception in GetCertificateAsX509"
            Write-Host $_
        }

        return $x509Certificate
    }

    <########## hidden/private methods ##########>

}
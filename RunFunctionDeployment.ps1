param (
    [Parameter(Mandatory=$true)]
    [string] $SubscriptionId,
    [Parameter(Mandatory=$true)]
    [string] $ResourceGroup,
    [Parameter(Mandatory=$true)]
    [string] $FunctionApp
)

$gitStatus = git status -s

if ($gitStatus -ne $null)
{
    Write-Host -ForegroundColor Red "There are uncommited changes, cannot get a correct commit id. Abort action."
}
else 
{
    Connect-AzAccount -Subscription $SubscriptionId

    $subscription = Get-AzSubscription -SubscriptionId $SubscriptionId | Select-AzSubscription

    if ($null -eq $subscription)
    {
        Write-Host -ForegroundColor Red "Cannot get and connect to the subscription. Abort action."

        Exit
    }
    
    # pre check

    <#
        0 => must be available in SAP tenant
        1 => must be available
        2 => can be available, solution uses default value, when missing
    #>
    
    $settingList = @{
        "Tenant" = 1
        "TenantId" = 1
        "SubscriptionId" = 1
        "KeyVaultName" = 1
        "KVSharePointConnectIdName" = 2
        "KVSharePointConnectCertName" = 2
    }
    
    $issuesFound = $false
    $issuesForProdFound = $false
    
    $settings = Get-AzFunctionAppSetting -ResourceGroupName $ResourceGroup -Name $FunctionApp
    
    foreach ($key in $settingList.Keys)
    {
        Write-Host -NoNewline $key + " ... "
    
        if ($settings.ContainsKey($key) -eq $true)
        {
            Write-Host -ForegroundColor Green "available"
        }
        else 
        {
            switch ($settingList[$key])
            {
                0 {
                    Write-Host -NoNewline -BackgroundColor Yellow -ForegroundColor Black "not available"
                    Write-Host " must be set in SAP tenant"
    
                    $issuesForProdFound = $true
                }
                1 {
                    Write-Host -ForegroundColor Red "not available"
    
                    $issuesFound = $true
                }
                2 {
                    Write-Host -NoNewline -ForegroundColor Green "not available"
                    Write-Host " default value will be used"
                }
            }
        }
    }
    
    Write-Host
    
    if ($issuesFound -eq $true)
    {
        Write-Host -NoNewline "There were issues found in the configuration of the Function App. The deployment should be "
        Write-Host -ForegroundColor Red "stopped."
        Write-Host
    }
    
    if ($issuesForProdFound -eq $true)
    {
        Write-Host "There were missing settings found in the configuration of the Function App that must be set in the SAP tenant."
        Write-Host
    }
    
    $answer = "Y"
    
    if (($issuesFound -eq $true) -or ($issuesForProdFound -eq $true))
    {
        $answer = Read-Host "Should the solution be deployed? (Y/N)"
    }
    
    if ($answer.ToUpper() -eq "Y")
    {
        # deployment of the solution

        git config --get remote.origin.url | Tee-Object -Variable originUrl

        if ([string]::IsNullOrEmpty($originUrl) -eq $true)
        {
            Write-Host -ForegroundColor Red "Cannot get the origin url of the repository. Abort action."
        }
        else
        {
            git branch --show-current | Tee-Object -Variable currentBranch

            if ([String]::IsNullOrEmpty($currentBranch) -eq $true)
            {
                Write-Host -ForegroundColor Red "Cannot get the current branch. Abort action."
            }
            else
            {
                git log -1 | Tee-Object -Variable gitCommit
        
                if ($gitCommit.Length -ge 2)
                {
                    $commitId = ($gitCommit[0].Split(" ")[1])
                    $author = $gitCommit[1].TrimStart("Author:").Trim()
                    $date = $gitCommit[2].TrimStart("Date:").Trim()

                    $commitUrl = $originUrl + "/commit/" + $commitId
        
                    $versionInfo = @"
{
    'CommitId': '$commitId', 
    'Branch': '$currentBranch', 
    'Author': '$author', 
    'Date': '$date', 
    'CommitUrl': '$commitUrl'
}
"@

                    $versionInfo | Out-File -FilePath .\version-info.json
        
                    # deploy repository
                    $subscription = Get-AzSubscription -SubscriptionId $SubscriptionId | Select-AzSubscription

                    if ($null -eq $subscription)
                    {
                        Write-Host -ForegroundColor Red "Cannot get and connect to the subscription. Abort action."
                    }
                    else 
                    {
                        $app = Get-AzFunctionApp -SubscriptionId $SubscriptionId -ResourceGroupName $ResourceGroup -Name $FunctionApp

                        if ($null -eq $app)
                        {
                            Write-Host -ForegroundColor Red "Cannot get and connect to the function app. Abort action."
                        }
                        else 
                        {
                            func azure functionapp publish $FunctionApp --powershell
                        }
                    }

                    Remove-Item -Path .\version-info.json
                }
                else 
                {
                    Write-Host -ForegroundColor Red "Incomplete commit data received. Abort action."
                }
            }
        }
    }
}


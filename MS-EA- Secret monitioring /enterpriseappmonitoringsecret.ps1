#Created by Tim Hjort, 2024
#Requirments#
# 1. Create a new Automation Account in Azure
# 2. Create a new Runbook in the Automation Account
# 3. Add the script to the Runbook
# 4. Create the following variables in the Automation Account:
#    - TenantName
#    - msgraph-clientcred-appid
#    - msgraph-clientcred-appsecret
#    - TeamsChannelUri
# 5. Create a new schedule in the Automation Account
# 6. Add the Runbook to the schedule
# 7. Run the schedule
# 8. Monitor the Teams channel for alerts
# 9. Enjoy

#Enterprise app requirmentens and permissions needed:
# 1. The script requires the following permissions:
# Application.Read.All
# Directory.Read.All
# User.Read




$TenantID = Get-AutomationVariable -Name 'TenantID'
$AppID = Get-AutomationVariable -Name "msgraph-clientcred-appid"
$AppSecret = Get-AutomationVariable -Name "msgraph-clientcred-appsecret"
$Uri = Get-AutomationVariable -Name "TeamsChannelUri"

Function Connect-MSGraphAPI {
    param (
        [system.string]$AppID,
        [system.string]$TenantID,
        [system.string]$AppSecret
    )
    begin {
        $URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
        $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            client_Id     = $AppID
            Client_Secret = $AppSecret
        } 
    }
    Process {
        Write-Host "Connecting to the Graph API"
        $Response = Invoke-RestMethod -Uri $URI -Method POST -Body $ReqTokenBody
    }
    End{
        $Response
    }
}


$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret
Function Get-MSGraphRequest {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken
    )
    begin {
        [System.Array]$allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $Uri
        }
    }
    process {
        write-verbose "GET request at endpoint: $Uri"
        $data = Invoke-RestMethod @ReqTokenBody
        while ($data.'@odata.nextLink') {
            $allPages += $data.value
            $ReqTokenBody.Uri = $data.'@odata.nextLink'
            $Data = Invoke-RestMethod @ReqTokenBody
            # to avoid throttling, the loop will sleep for 3 seconds
            Start-Sleep -Seconds 3
        }
        $allPages += $data.value
    }
    end {
        Write-Verbose "Returning all results"
        $allPages
    }
}

# Assuming you have a list of applications in the variable "$applications"
# Modify the code to only show applications expiring inside the defined threshold
$results = @()
$applications = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/"
$expirationThresholdInDays = 30
$timeZoneId = 'Central Standard Time'
$currentTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, $timeZoneId)

foreach ($app in $applications | Sort-Object displayName) {
    $passwordCreds = $app.passwordCredentials | Where-Object { $_.endDateTime }
    if (-not $passwordCreds) {
        continue
    }

    $expiringSecrets = @()
    foreach ($cred in $passwordCreds) {
        $endDate = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]$cred.endDateTime, $timeZoneId)
        if ($endDate -is [system.array]) {
            $endDate = $endDate[0]
        }

        $timeSpan = New-TimeSpan -Start $currentTime -End $endDate
        $daysUntilExpiration = [int][Math]::Floor($timeSpan.TotalDays)
        if ($daysUntilExpiration -le $expirationThresholdInDays) {
            $expiringSecrets += [PSCustomObject]@{
                KeyId     = $cred.KeyId
                EndDate   = $endDate
                DaysUntil = $daysUntilExpiration
            }
        }
    }

    if (-not $expiringSecrets) {
        continue
    }

    $ownerUri = "https://graph.microsoft.com/v1.0/applications/$($app.id)/owners?$select=displayName,userPrincipalName,mail"
    $owners = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri $ownerUri
    $ownerDisplay = if ($owners) {
        ($owners | ForEach-Object {
            if ($_.userPrincipalName) {
                "{0} ({1})" -f $_.displayName, $_.userPrincipalName
            }
            elseif ($_.mail) {
                "{0} ({1})" -f $_.displayName, $_.mail
            }
            else {
                $_.displayName
            }
        }) -join ', '
    }
    else {
        'No owners assigned'
    }

    foreach ($secret in $expiringSecrets | Sort-Object DaysUntil, EndDate) {
        $daysUntilValue = if ($secret.DaysUntil -is [system.array]) { [int]$secret.DaysUntil[0] } else { [int]$secret.DaysUntil }

        $results += [PSCustomObject]@{
            AppId       = $app.id
            DisplayName = $app.displayName
            SecretKeyId = $secret.KeyId
            DaysUntil   = $daysUntilValue
            Expiration  = $secret.EndDate.ToString('yyyy-MM-dd HH:mm')
            Owners      = $ownerDisplay
        }
    }
}

$textTable = $results | Sort-Object DisplayName, DaysUntil | Select-Object DisplayName, SecretKeyId, @{Name='DaysUntil';Expression={[string]($_.DaysUntil)}}, Expiration, Owners | ConvertTo-Html
$JSONBody = [PSCustomObject][Ordered]@{
    "@type"      = "MessageCard"
    "@context"   = "<http://schema.org/extensions>"
    "themeColor" = 'c13d29'
    "title"      = "$($Array.count) App Secrets are expiring Soon"
    "text"       = "$textTable"
}

$TeamMessageBody = ConvertTo-Json $JSONBody

$parameters = @{
    "URI"         = "$Uri"
    "Method"      = 'POST'
    "Body"        = $TeamMessageBody
    "ContentType" = 'application/json'
}

Invoke-RestMethod @parameters
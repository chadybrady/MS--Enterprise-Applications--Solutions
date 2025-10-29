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
# User.Reada




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

$script:ServicePrincipalCache = @{}
Function Get-ServicePrincipalByAppId {
    param (
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    if ([string]::IsNullOrWhiteSpace($AppId)) {
        return $null
    }

    if ($script:ServicePrincipalCache.ContainsKey($AppId)) {
        return $script:ServicePrincipalCache[$AppId]
    }

    $uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq ''{0}''&$select=id,displayName,appId,appRoles,oauth2PermissionScopes' -f $AppId
    $servicePrincipal = Get-MSGraphRequest -AccessToken $AccessToken -Uri $uri | Select-Object -First 1
    $script:ServicePrincipalCache[$AppId] = $servicePrincipal
    return $servicePrincipal
}

# Assuming you have a list of applications in the variable "$applications"
# Modify the code to only show applications expiring inside the defined threshold
$results = @()
$permissionRows = @()
$assignmentRows = @()
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

    $ownerUri = 'https://graph.microsoft.com/v1.0/applications/{0}/owners?$select=displayName,userPrincipalName,mail' -f $app.id
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

    $servicePrincipal = Get-ServicePrincipalByAppId -AppId $app.appId -AccessToken $tokenResponse.access_token
    $localDelegatedPermissions = @()
    $localApplicationPermissions = @()

    if ($app.requiredResourceAccess) {
        foreach ($resource in $app.requiredResourceAccess) {
            if (-not $resource.resourceAccess) {
                continue
            }

            $resourceSp = Get-ServicePrincipalByAppId -AppId $resource.resourceAppId -AccessToken $tokenResponse.access_token
            $resourceName = if ($resourceSp -and $resourceSp.displayName) { $resourceSp.displayName } else { $resource.resourceAppId }

            foreach ($permission in $resource.resourceAccess) {
                switch ($permission.type) {
                    'Scope' {
                        $permissionName = $permission.id
                        if ($resourceSp -and $resourceSp.oauth2PermissionScopes) {
                            $scope = $resourceSp.oauth2PermissionScopes | Where-Object { $_.id -eq $permission.id }
                            if ($scope) {
                                $permissionName = $scope.adminConsentDisplayName
                                if (-not $permissionName) { $permissionName = $scope.value }
                            }
                        }
                        if (-not $permissionName) { $permissionName = $permission.id }

                        $localDelegatedPermissions += "{0}: {1}" -f $resourceName, $permissionName
                        $permissionRows += [PSCustomObject]@{
                            AppDisplayName  = $app.displayName
                            Resource        = $resourceName
                            PermissionType  = 'Delegated'
                            PermissionName  = $permissionName
                        }
                    }
                    'Role' {
                        $permissionName = $permission.id
                        if ($resourceSp -and $resourceSp.appRoles) {
                            $role = $resourceSp.appRoles | Where-Object { $_.id -eq $permission.id }
                            if ($role) {
                                $permissionName = $role.displayName
                                if (-not $permissionName) { $permissionName = $role.value }
                            }
                        }
                        if (-not $permissionName) { $permissionName = $permission.id }

                        $localApplicationPermissions += "{0}: {1}" -f $resourceName, $permissionName
                        $permissionRows += [PSCustomObject]@{
                            AppDisplayName  = $app.displayName
                            Resource        = $resourceName
                            PermissionType  = 'Application'
                            PermissionName  = $permissionName
                        }
                    }
                }
            }
        }
    }

    $localAssignments = @()
    if ($servicePrincipal) {
        $assignmentUri = 'https://graph.microsoft.com/v1.0/servicePrincipals/{0}/appRoleAssignedTo?$select=principalId,principalDisplayName,principalType,appRoleId' -f $servicePrincipal.id
        $assignments = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri $assignmentUri
        foreach ($assignment in $assignments) {
            $principalName = if ($assignment.principalDisplayName) { $assignment.principalDisplayName } else { $assignment.principalId }
            $roleLabel = $null
            if ($assignment.appRoleId -and $servicePrincipal.appRoles) {
                $roleMatch = $servicePrincipal.appRoles | Where-Object { $_.id -eq $assignment.appRoleId }
                if ($roleMatch) {
                    $roleLabel = $roleMatch.displayName
                    if (-not $roleLabel) { $roleLabel = $roleMatch.value }
                }
            }
            if (-not $roleLabel) {
                $roleLabel = 'Default'
            }

            $localAssignments += "{0} ({1}) - {2}" -f $principalName, $assignment.principalType, $roleLabel
            $assignmentRows += [PSCustomObject]@{
                AppDisplayName        = $app.displayName
                PrincipalType         = $assignment.principalType
                PrincipalDisplayName  = $principalName
                Role                  = $roleLabel
            }
        }
    }

    $delegatedSummary = if ($localDelegatedPermissions) { ($localDelegatedPermissions | Sort-Object -Unique) -join '; ' } else { 'None' }
    $applicationSummary = if ($localApplicationPermissions) { ($localApplicationPermissions | Sort-Object -Unique) -join '; ' } else { 'None' }
    $assignmentSummary = if ($localAssignments) { ($localAssignments | Sort-Object -Unique) -join '; ' } else { 'None' }

    foreach ($secret in $expiringSecrets | Sort-Object DaysUntil, EndDate) {
        $daysUntilValue = if ($secret.DaysUntil -is [system.array]) { [int]$secret.DaysUntil[0] } else { [int]$secret.DaysUntil }

        $results += [PSCustomObject]@{
            AppId                  = $app.id
            DisplayName            = $app.displayName
            SecretKeyId            = $secret.KeyId
            DaysUntil              = $daysUntilValue
            Expiration             = $secret.EndDate.ToString('yyyy-MM-dd HH:mm')
            Owners                 = $ownerDisplay
            DelegatedPermissions   = $delegatedSummary
            ApplicationPermissions = $applicationSummary
            Assignments            = $assignmentSummary
        }
    }
}

# Comment out any of the sections below to exclude them from the Teams message payload.
$sections = @()

if ($results) {
    $secretTable = ($results |
        Sort-Object DisplayName, DaysUntil |
        Select-Object DisplayName, SecretKeyId, @{Name = 'DaysUntil'; Expression = { [string]$_.DaysUntil }}, Expiration, Owners, DelegatedPermissions, ApplicationPermissions, Assignments |
        ConvertTo-Html -Fragment | Out-String).Trim()

    $sections += @{
        Title   = 'Expiring Secrets'
        Content = $secretTable
    }
}

if ($permissionRows) {
    $permissionTable = ($permissionRows |
        Sort-Object AppDisplayName, PermissionType, Resource, PermissionName |
        ConvertTo-Html -Fragment | Out-String).Trim()

    $sections += @{
        Title   = 'API Permissions'
        Content = $permissionTable
    }
}

if ($assignmentRows) {
    $assignmentTable = ($assignmentRows |
        Sort-Object AppDisplayName, PrincipalType, PrincipalDisplayName |
        ConvertTo-Html -Fragment | Out-String).Trim()

    $sections += @{
        Title   = 'User and Group Assignments'
        Content = $assignmentTable
    }
}

$textTable = if ($sections) {
    ($sections | ForEach-Object { "<h3>{0}</h3>{1}" -f $_.Title, $_.Content }) -join '<br />'
} else {
    'No expiring secrets within the configured window.'
}
$secretCount = $results.Count
$uniqueApps = @($results | Select-Object -ExpandProperty DisplayName -Unique)
$appCount = $uniqueApps.Count
$JSONBody = [PSCustomObject][Ordered]@{
    "@type"      = "MessageCard"
    "@context"   = "<http://schema.org/extensions>"
    "themeColor" = 'c13d29'
    "title"      = ('{0} secret(s) across {1} app(s) expiring within {2} days' -f $secretCount, $appCount, $expirationThresholdInDays)
    "text"       = $textTable
}

$TeamMessageBody = ConvertTo-Json $JSONBody

$parameters = @{
    "URI"         = "$Uri"
    "Method"      = 'POST'
    "Body"        = $TeamMessageBody
    "ContentType" = 'application/json'
}

Invoke-RestMethod @parameters
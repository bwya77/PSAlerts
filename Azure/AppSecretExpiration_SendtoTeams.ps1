$AppID = Get-AzAutomationVariable -Name 'appID' 
$TenantID = Get-AzAutomationVariable -Name 'tenantID'
$AppSecret = Get-AzAutomationVariable -Name 'appSecret' 

[string]$teamsWebhookURI = ''
[int32]$expirationDays = 30

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
        $allPages = @()
        $ReqTokenBody = @{
            Headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $($AccessToken)"
            }
            Method  = "Get"
            Uri     = $URI
        }
    }
    process {
        $data = Invoke-RestMethod @ReqTokenBody
        if ($data.'@odata.nextLink') {
            do {
                $ReqTokenBody = @{
                    Headers = @{
                        "Content-Type"  = "application/json"
                        "Authorization" = "Bearer $($AccessToken)"
                    }
                    Method  = "Get"
                    Uri     = $URI
                }
                $Data = Invoke-RestMethod @ReqTokenBody
                $allPages += $Data
            } until (
                !$data.'@odata.nextLink'
            )
        }
        else {
            $allPages += $Data
        }
    }
    end {
        $allPages
    }
}

$array = @()
$applications = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/"
$Applications.value | Sort-Object displayName | Foreach-Object {
    #If there are more than one password credentials, we need to get the expiration of each one
    if ($_.passwordCredentials.endDateTime.count -gt 1) {
        $endDates = $_.passwordCredentials.endDateTime
        [int[]]$daysUntilExpiration = @()
        foreach ($Date in $endDates) {
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Date, 'Central Standard Time')
            $daysUntilExpiration += (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Central Standard Time")) -End $Date).Days
        }
    }
    Elseif ($_.passwordCredentials.endDateTime.count -eq 1) {
        $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.passwordCredentials.endDateTime, 'Central Standard Time')
        $daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Central Standard Time")) -End $Date).Days 
    }

    if ($daysUntilExpiration -le $expirationDays) {
        $array += $_ | Select-Object id, displayName, @{
            name = "daysUntil"; 
            expr = { $daysUntilExpiration } 
        }
    }
}

$textTable = $array | Sort-Object daysUntil | select-object displayName, daysUntil | ConvertTo-Html
$JSONBody = [PSCustomObject][Ordered]@{
    "@type"      = "MessageCard"
    "@context"   = "<http://schema.org/extensions>"
    "themeColor" = '0078D7'
    "title"      = "$($Array.count) App Secrets areExpiring Soon"
    "text"       = "$textTable"
}

$TeamMessageBody = ConvertTo-Json $JSONBody

$parameters = @{
    "URI"         = $teamsWebhookURI
    "Method"      = 'POST'
    "Body"        = $TeamMessageBody
    "ContentType" = 'application/json'
}

Invoke-RestMethod @parameters
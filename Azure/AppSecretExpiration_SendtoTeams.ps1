$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret' 

[string]$teamsWebhookURI = '[ENTER WEBHOOK URL HERE]'
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

Function Get-MSGraphRequest {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken
    )
    begin {
        [array]$allPages = @()
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
        do {
            $data = Invoke-RestMethod @ReqTokenBody
            $allpages += $data.value
            if ($data.'@odata.nextLink') {
                $ReqTokenBody.Uri = $data.'@odata.nextLink'
            }
        } until (!$data.'@odata.nextLink')
    }
    end {
        $allPages
    }
}

$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret
$array = @()
$apps = Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/" 
foreach ($app in $apps) {
    $app.passwordCredentials | foreach-object {
        #If there is a secret with a enddatetime, we need to get the expiration of each one
        if ($_.endDateTime -ne $null) {
            [system.string]$secretdisplayName = $_.displayName
            [system.string]$id = $app.id
            [system.string]$displayname = $app.displayName
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.endDateTime, 'Central Standard Time')
            [int32]$daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Central Standard Time")) -End $Date).Days
            
            if (($daysUntilExpiration -ne $null) -and ($daysUntilExpiration -le $expirationDays)) {
                $array += $_ | Select-Object @{
                    name = "id"; 
                    expr = { $id } 
                }, 
                @{
                    name = "displayName"; 
                    expr = { $displayName } 
                }, 
                @{
                    name = "secretName"; 
                    expr = { $secretdisplayName } 
                },
                @{
                    name = "daysUntil"; 
                    expr = { $daysUntilExpiration } 
                }
            }
            $daysUntilExpiration = $null
            $secretdisplayName = $null
        }
    }
}

if ($array.count -ne 0) {
    Write-output "Sending Teams Message"
    $textTable = $array | Sort-Object daysUntil | select-object displayName, secretName, daysUntil | ConvertTo-Html
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
}
else {
    write-output "No App Secrets are expiring soon"
}
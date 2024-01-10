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
Get-MSGraphRequest -AccessToken $tokenResponse.access_token -Uri "https://graph.microsoft.com/v1.0/applications/" |  Foreach-Object {
    [string]$secretdisplayName = $_.passwordCredentials.displayName
    [string]$id = $_.id
    [string]$displayname = $_.displayName

    #If there are more than one password credentials, we need to get the expiration of each one
    if ($_.passwordCredentials.endDateTime.count -gt 1) {
        $endDates = $_.passwordCredentials.endDateTime
        [int[]]$daysUntilExpiration = @()
        foreach ($Date in $endDates) {
            $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($Date, 'Central Standard Time')
            $daysUntilExpiration += (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Central Standard Time")) -End $Date).Days
        }
    }
    ElseIf ($_.passwordCredentials.endDateTime.count -eq 1) {
        $Date = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($_.passwordCredentials.endDateTime, 'Central Standard Time')
        $daysUntilExpiration = (New-TimeSpan -Start ([System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::Now, "Central Standard Time")) -End $Date).Days 
    }

    $daysUntilExpiration | foreach-object { 
        if (($_ -ne $null) -and ($_ -le $expirationDays)) {
            $array += $_ | Select-Object @{
                name = "id"; 
                expr = { $id } }, 
                @{
                name = "displayName"; 
                expr = { $displayName } }, 
                @{
                name = "secretdisplayName"; 
                expr = { $secretdisplayName } },
                @{
                name = "daysUntil"; 
                expr = { $_ } }
        }
    }
    $daysUntilExpiration = $null
}

if ($array.count -ne 0) {
    Write-output "Sending Teams Message"
    $textTable = $array | Sort-Object daysUntil | select-object displayName, secretdisplayname, daysUntil | ConvertTo-Html
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
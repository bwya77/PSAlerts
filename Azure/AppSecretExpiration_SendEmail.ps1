$AppID = Get-AutomationVariable -Name 'appID' 
$TenantID = Get-AutomationVariable -Name 'tenantID'
$AppSecret = Get-AutomationVariable -Name 'appSecret' 

[int32]$expirationDays = 30
[string]$emailSender = "[ENTER THE UPN OF THE SENDER]"
[string] $emailTo = "[ENTER THE EMAIL RECIPIENT]"

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
    End {
        $Response
    }
}

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

Function Send-MSGraphEmail {
    param (
        [system.string]$Uri,
        [system.string]$AccessToken,
        [system.string]$To,
        [system.string]$Subject = "App Secret Expiration Notice",
        [system.string]$Body
    )
    begin {
        $headers = @{
            "Authorization" = "Bearer $($AccessToken)"
            "Content-type"  = "application/json"
        }

        $BodyJsonsend = @"
{
   "message": {
   "subject": "$Subject",
   "body": {
      "contentType": "HTML",
      "content": "$($Body)"
   },
   "toRecipients": [
      {
      "emailAddress": {
      "address": "$to"
          }
      }
   ]
   },
   "saveToSentItems": "true"
}
"@
    }
    process {
        $data = Invoke-RestMethod -Method POST -Uri $Uri -Headers $headers -Body $BodyJsonsend
    }
    end {
        $data
    }
}




$tokenResponse = Connect-MSGraphAPI -AppID $AppID -TenantID $TenantID -AppSecret $AppSecret

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

$textTable = $array | Sort-Object daysUntil | select-object displayName, daysUntil | ConvertTo-Html -Fragment
Send-MSGraphEmail -Uri "https://graph.microsoft.com/v1.0/users/$emailSender/sendMail" -AccessToken $tokenResponse.access_token -To $emailTo  -Body $textTable 
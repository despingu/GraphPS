[string]$Script:tenantName = [string]::Empty
[string]$Script:clientID = [string]::Empty
[string]$Script:clientSecret = [string]::Empty
[string]$Script:resourceAppIdURI = [string]::Empty
[string]$Script:token = [string]::Empty
[datetime]$Script:tokenRenewTime = [datetime]::MinValue
[bool]$Script:connected = $false
[string]$Script:graphUri = "https://graph.microsoft.com"
[string]$Script:graphVersion = [string]::Empty

Function Merge-Hashtables([ScriptBlock]$Operator) {
    <#
        .SYNOPSIS
            A function to merge two or more hash tables.
            By default, all values from duplicated hash table entries will added to an array
            SYNTAX:
            HashTable[] <Hashtables> | Merge-Hashtables [-Operator <ScriptBlock>]
            From: https://stackoverflow.com/questions/8800375/merging-hashtables-in-powershell-how

        .EXAMPLE
            Collecting all values in array
            PS C:> $h1, $h2, $h3 | Merge-Hashtables

            Overwrite using last value:
            PS C:> $h1, $h2, $h3 | Merge-Hashtables {$_[-1]}
            
            Overwrite using first value:
            PS C:> $h1, $h2, $h3 | Merge-Hashtables {$_[0]}

            Overwrite using max value:
            PS C:> $h1, $h2, $h3 | Merge-Hashtables {($_ | Measure-Object -Maximum).Maximum}

            Take the average values
            PS C:\> $h1, $h2, $h3 | Merge-Hashtables {($_ | Measure-Object -Average).Average}

            Join the values together
            PS C:\> $h1, $h2, $h3 | Merge-Hashtables {$_ -Join ""}

            Sort the values list
            PS C:\> $h1, $h2, $h3 | Merge-Hashtables {$_ | Sort-Object}
    #>
    $Output = @{}
    ForEach ($Hashtable in $Input) {
        If ($Hashtable -is [Hashtable]) {
            ForEach ($Key in $Hashtable.Keys) {$Output.$Key = If ($Output.ContainsKey($Key)) {@($Output.$Key) + $Hashtable.$Key} Else  {$Hashtable.$Key}}
        }
    }
    If ($Operator) {ForEach ($Key in @($Output.Keys)) {$_ = @($Output.$Key); $Output.$Key = Invoke-Command $Operator}}
    $Output
}

function Set-AccessToken {
    $logonUrl = "https://login.microsoftonline.com/$Script:tenantName/oauth2/v2.0/token"
    $body = @{
        "client_id" = $Script:clientID;
        "client_secret" = $Script:clientSecret;
        "scope" = $Script:resourceAppIdURI;
        "grant_type" = "client_credentials"
    }
    $tokenInfo = Invoke-RestMethod -Method Post -Uri $logonUrl -Body $body
    $Script:tokenRenewTime = (Get-Date).AddSeconds($tokenInfo.expires_in)
    $Script:token = "$($tokenInfo.token_type) $($tokenInfo.access_token)"
}

function Get-GraphUri ($Endpoint, $FilterExpression, $SelectExpression) {
    $uri = "$($Script:graphUri)/$($Script:graphVersion)/$Endpoint"
    $expressions = @()
    if(-not [string]::IsNullOrEmpty($FilterExpression)) {
        $expressions += "`$filter=$FilterExpression"
    }
    if(-not [string]::IsNullOrEmpty($SelectExpression)) {
        $expressions += "`$select=$SelectExpression"
    }
    if($expressions.count -gt 0) {
        $expressionStr = $expressions -join "&"
        $uri += "?$expressionStr" 
    }

    return $uri
}

function Invoke-MSGraphQuery {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Endpoint,
        [Parameter(Mandatory=$true)]
        [ValidateSet('GET','POST','PATCH','PUT','DELETE')]
        [string]$Method, 
        [Parameter(Mandatory=$false)]
        [string]$FilterExpression,
        [Parameter(Mandatory=$false)]
        [string] $SelectExpression, 
        [Parameter(Mandatory=$false)]
        [string]$Body, 
        [Parameter(Mandatory=$false)]
        [hashtable]$customHeader
    )
    $Uri = Get-GraphUri -Endpoint $Endpoint -FilterExpression $FilterExpression -SelectExpression $SelectExpression

    if ((Get-Date) -ge $Script:tokenRenewTime) {
        $Script:token = Set-AccessToken
    }
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -CurrentOperation "Invoking MS Graph API"
    $defaultHeader = @{
        'Content-Type'  = 'application\json'
        'Authorization' = $Script:token
    }

    $Header = @{} 
    if($null -eq $customHeader) {
        $Header = $defaultHeader, $customHeader | Merge-Hashtables {$_[-1]}
    }
    else {
        $Header = $defaultHeader
    }

    $QueryResults = @()
    if ($Method -eq "Get") {
        do {
            $Results = Invoke-RestMethod -Headers $Header -Uri $Uri -Method $Method -ContentType "application/json"
            if ($null -ne $Results.value) {
                $QueryResults += $Results.value
            }
            else {
                $QueryResults += $Results
            }
            $uri = $Results.'@odata.nextlink'
        } until ($null -eq $uri)
    }
    else  {
        $Results = Invoke-RestMethod -Headers $Header -Uri $Uri -Method $Method -ContentType "application/json" -Body $Body
    }
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -Completed
    Return $QueryResults
}

function Test-Connection {
    $isValid = $true
    if (-not $Script:connected) {
        Write-Error "Please run Connect-GraphPS cmdlet before running this command."
        $isValid = $false
    }
    return $isValid
}

function Connect-GraphPS {
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantName,
        [Parameter(Mandatory=$true)]
        [string]$AppID,
        [Parameter(Mandatory=$true)]
        [string]$AppSecret,
        [Parameter(Mandatory=$false)]
        [ValidateSet('beta','v1.0')]
        [string]$Version = 'v1.0',
        [Parameter(Mandatory=$false)]
        [string]$ResourceAppIdURI = "https://graph.microsoft.com/.default"
    )
    
    $Script:tenantName = $TenantName
    $Script:clientID = $AppID
    $Script:clientSecret = $AppSecret
    $Script:resourceAppIdURI = $ResourceAppIdURI
    $Script:graphVersion = $Version
    Set-AccessToken -TenantName $TenantName -ClientID $AppID -ClientSecret $AppSecret -ResourceAppIdURI $ResourceAppIdURI
    $Script:connected = $true
}

function Disconnect-GraphPS {
    $Script:tenantName = [string]::Empty
    $Script:clientID = [string]::Empty
    $Script:clientSecret = [string]::Empty
    $Script:resourceAppIdURI = [string]::Empty
    $Script:token = [string]::Empty
    $Script:tokenRenewTime = [datetime]::MinValue
    $Script:connected = $false
    $Script:graphUri = "https://graph.microsoft.com"
    $Script:graphVersion = [string]::Empty
}

function Get-GraphPSConnectionInfo {
    Write-Host "Tenant name: $Script:tenantName"
    Write-Host "App ID: $Script:clientID"
    Write-Host "Resource AppID URI: $Script:resourceAppIdURI"
    Write-Host "Token renew time: $Script:tokenRenewTime"
    Write-Host "Connected: $Script:connected"
    Write-Host "Graph base uri: $Script:graphUri"
    Write-Host "Graph version: $Script:graphVersion"
}

function Get-GraphPSUsers {
    param (
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not $Script:connected) {
        Write-Error "Please run Connect-GraphPS cmdlet before running this command."
        return
    }

    $endpoint = "users"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSUser {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users/$identity"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSAuditLogsSignIns {
    param (
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "auditLogs/signIns"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSUserEvents {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false, ParameterSetName='specificCalendar')]
        [string]$calendarId,
        [Parameter(Mandatory=$false, ParameterSetName='defaultCalendar')]
        [switch]$default,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = ""

    if($default) {
        $endpoint = "users/$identity/calendar/events" #default calendar
    }
    if(-not [string]::IsNullOrEmpty($calendarId)) {
        $endpoint = "users/$identity/calendars/$calendarId/events" #specific calendar
    }
    else {
        $endpoint = "users/$identity/events" # all events
    }
    

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSUserCalendarGroups {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = "users/$identity/calendarGroups"
    

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSUserCalendars {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false, ParameterSetName='specificGroup')]
        [string]$groupId,
        [Parameter(Mandatory=$false, ParameterSetName='defaultGroup')]
        [switch]$default,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = ""
    if($default) {
        $endpoint = "users/$identity/calendarGroup/calendars" # calendars from default group
    }
    elseif(-not ([string]::IsNullOrEmpty($groupId))) {
        $endpoint = "users/$identity/calendarGroups/$groupId/calendars" #calendars from specific group
    }
    else {
        $endpoint = "users/$identity/calendars" # all calendars 
    }
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}


<#
function Get-GraphPSSample {
    param (
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}
#>

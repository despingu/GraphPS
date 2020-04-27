function Remove-GraphPSUser {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    if (-not (Test-Connection)) {
        return
    }

    if(-not $Force) {
        $question = "Are you sure you want to delete user $($identity)?"
        $choices  = '&Yes', '&No'
        $decision = $Host.UI.PromptForChoice($null, $question, $choices, 1)
        if ($decision -eq 1) {
            return
        }
    }

    $endpoint = "users/$($identity)"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method DELETE
    return $graphResult
}

function Get-GraphPSUsers {
    param (
        [Parameter(Mandatory=$false)]
        [string]$FilterExpression,
        [Parameter(Mandatory=$false)]
        [string]$SelectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
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

function Get-GraphPSUserEvents {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false, ParameterSetName='specificCalendar')]
        [string]$calendarId,
        [Parameter(Mandatory=$false, ParameterSetName='defaultCalendar')]
        [switch]$default,
        [Parameter(Mandatory=$false)]
        [string]$FilterExpression,
        [Parameter(Mandatory=$false)]
        [string]$SelectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
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

function Get-GraphPSUserCalendars {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false, ParameterSetName='specificGroup')]
        [string]$groupId,
        [Parameter(Mandatory=$false, ParameterSetName='defaultGroup')]
        [switch]$default,
        [Parameter(Mandatory=$false)]
        [string]$FilterExpression,
        [Parameter(Mandatory=$false)]
        [string]$SelectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
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
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Test-GraphPSUserGroupMembership {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$true, Position=1)]
        [string[]]$GroupIds
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users/$identity/checkMemberGroups"
    $payload = @{groupIds = $GroupIds}
    $body = ConvertTo-JSON $payload
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method POST -Body $body
    return $graphResult
}
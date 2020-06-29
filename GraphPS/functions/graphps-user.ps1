function Remove-GraphPSUser {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    if (-not (Test-Connection)) {
        return
    }

    if (-not $Force) {
        $question = "Are you sure you want to delete user $($identity)?"
        $choices = '&Yes', '&No'
        $decision = $Host.UI.PromptForChoice($null, $question, $choices, 1)
        if ($decision -eq 1) {
            return
        }
    }

    $endpoint = "users/$($identity)"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method DELETE
    return $graphResult
}

function Update-GraphPSUser {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $true)]
        [Boolean] $accountEnabled,
        [Parameter(Mandatory = $false)]
        [String[]] $businessPhones,
        [Parameter(Mandatory = $false)]
        [String] $city,
        [Parameter(Mandatory = $false)]
        [String] $country,
        [Parameter(Mandatory = $false)]
        [String] $department,
        [Parameter(Mandatory = $false)]
        [String] $displayName,
        [Parameter(Mandatory = $false)]
        [String] $givenName,
        [Parameter(Mandatory = $false)]
        [String] $jobTitle,
        [Parameter(Mandatory = $false)]
        [String] $mailNickname,
        [Parameter(Mandatory = $false)]
        [String] $mobilePhone,
        [Parameter(Mandatory = $false)]
        [String] $officeLocation,
        [Parameter(Mandatory = $false)]
        [String] $onPremisesImmutableId,
        [Parameter(Mandatory = $false)]
        [String] $otherMails,
        [Parameter(Mandatory = $false)]
        [String] $postalCode,
        [Parameter(Mandatory = $false)]
        [String] $preferredLanguage,
        [Parameter(Mandatory = $false)]
        [String] $state,
        [Parameter(Mandatory = $false)]
        [String] $streetAddress,
        [Parameter(Mandatory = $false)]
        [String] $surname,
        [Parameter(Mandatory = $false)]
        [String] $usageLocation,
        [Parameter(Mandatory = $false)]
        [String] $userPrincipalName,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Member', 'Guest')]
        [String] $userType,
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    if (-not (Test-Connection)) {
        return
    }

    if (-not $Force) {
        $question = "Are you sure you want to update user $($identity)?"
        $choices = '&Yes', '&No'
        $decision = $Host.UI.PromptForChoice($null, $question, $choices, 1)
        if ($decision -eq 1) {
            return
        }
    }

    $endpoint = "users/$identity"

    $userObject = @{ "accountEnabled" = $accountEnabled }
    if ($null -ne $businessPhones) { $userObject.Add("businessPhones", $businessPhones); }
    if ([string]::Empty -ne $city) { $userObject.Add("city", $city); }
    if ([string]::Empty -ne $country) { $userObject.Add("country", $country); }
    if ([string]::Empty -ne $department) { $userObject.Add("department", $department); }
    if ([string]::Empty -ne $displayName) { $userObject.Add("displayName", $displayName); }
    if ([string]::Empty -ne $givenName) { $userObject.Add("givenName", $givenName); }
    if ([string]::Empty -ne $jobTitle) { $userObject.Add("jobTitle", $jobTitle); }
    if ([string]::Empty -ne $mailNickname) { $userObject.Add("mailNickname", $mailNickname); }
    if ([string]::Empty -ne $mobilePhone) { $userObject.Add("mobilePhone", $mobilePhone); }
    if ([string]::Empty -ne $officeLocation) { $userObject.Add("officeLocation", $officeLocation); }
    if ([string]::Empty -ne $onPremisesImmutableId) { $userObject.Add("onPremisesImmutableId", $onPremisesImmutableId); }
    if ([string]::Empty -ne $otherMails) { $userObject.Add("otherMails", $otherMails); }
    if ([string]::Empty -ne $postalCode) { $userObject.Add("postalCode", $postalCode); }
    if ([string]::Empty -ne $preferredLanguage) { $userObject.Add("preferredLanguage", $preferredLanguage); }
    if ([string]::Empty -ne $state) { $userObject.Add("state", $state); }
    if ([string]::Empty -ne $streetAddress) { $userObject.Add("streetAddress", $streetAddress); }
    if ([string]::Empty -ne $surname) { $userObject.Add("surname", $surname); }
    if ([string]::Empty -ne $usageLocation) { $userObject.Add("usageLocation", $usageLocation); }
    if ([string]::Empty -ne $userPrincipalName) { $userObject.Add("userPrincipalName", $userPrincipalName); }
    if ([string]::Empty -ne $userType) { $userObject.Add("userType", $userType) }

    $body = ConvertTo-Json $userObject

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Body $body -Method PATCH
    return $graphResult
}

function Get-GraphPSUsers {
    param (
        [Parameter(Mandatory = $false)]
        [string]$FilterExpression,
        [Parameter(Mandatory = $false)]
        [string]$SelectExpression,
        [Parameter(Mandatory = $false)]
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
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $false)]
        [string]$filterExpression,
        [Parameter(Mandatory = $false)]
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
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $false, ParameterSetName = 'specificCalendar')]
        [string]$calendarId,
        [Parameter(Mandatory = $false, ParameterSetName = 'defaultCalendar')]
        [switch]$default,
        [Parameter(Mandatory = $false)]
        [string]$FilterExpression,
        [Parameter(Mandatory = $false)]
        [string]$SelectExpression,
        [Parameter(Mandatory = $false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = ""

    if ($default) {
        $endpoint = "users/$identity/calendar/events" #default calendar
    }
    if (-not [string]::IsNullOrEmpty($calendarId)) {
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
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $false, ParameterSetName = 'specificGroup')]
        [string]$groupId,
        [Parameter(Mandatory = $false, ParameterSetName = 'defaultGroup')]
        [switch]$default,
        [Parameter(Mandatory = $false)]
        [string]$FilterExpression,
        [Parameter(Mandatory = $false)]
        [string]$SelectExpression,
        [Parameter(Mandatory = $false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = ""
    if ($default) {
        $endpoint = "users/$identity/calendarGroup/calendars" # calendars from default group
    }
    elseif (-not ([string]::IsNullOrEmpty($groupId))) {
        $endpoint = "users/$identity/calendarGroups/$groupId/calendars" #calendars from specific group
    }
    else {
        $endpoint = "users/$identity/calendars" # all calendars 
    }
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSUserCalendarGroups {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $false)]
        [string]$FilterExpression,
        [Parameter(Mandatory = $false)]
        [string]$SelectExpression,
        [Parameter(Mandatory = $false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = "users/$identity/calendarGroups"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Test-GraphPSUserGroupMembership {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $true, Position = 1)]
        [string[]]$GroupIds
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users/$identity/checkMemberGroups"
    $payload = @{groupIds = $GroupIds }
    $body = ConvertTo-JSON $payload
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method POST -Body $body
    return $graphResult
}
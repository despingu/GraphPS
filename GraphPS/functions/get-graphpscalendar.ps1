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

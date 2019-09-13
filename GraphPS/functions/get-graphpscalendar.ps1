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

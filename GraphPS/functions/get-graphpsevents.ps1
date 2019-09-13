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

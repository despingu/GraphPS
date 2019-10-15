
function Get-GraphPSGroupOwners {
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

    $endpoint = "groups/$identity/owners"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}
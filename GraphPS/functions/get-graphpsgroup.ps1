function Get-GraphPSGroups {
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

    $endpoint = "groups"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}


function Get-GraphPSGroup {
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

    $endpoint = "groups/$identity"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}
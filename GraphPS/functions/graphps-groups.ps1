function Get-GraphPSGroups {
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

    $endpoint = "groups"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSGroup {
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

    $endpoint = "groups/$identity"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSGroupMembers {
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
    $endpoint = "groups/$identity/members"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Get-GraphPSGroupOwners {
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

    $endpoint = "groups/$identity/owners"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}

function Add-GraphPSGroupMember {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$ObjectId
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = "groups/$identity/members/`$ref"
    $memberJson = ConvertTo-Json @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$ObjectId" }

    if (-not [string]::IsNullOrEmpty($memberJson)) {
        $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method POST -Body $memberJson
    }
    return $graphResult
}
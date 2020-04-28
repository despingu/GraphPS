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
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'User')]
        [string]$UserToAdd,
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'Group')]
        [string]$GroupToAdd,
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'OrgContact')]
        [string]$OrganizationContactToAdd,
        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'DirObj')]
        [string]$DirectoryObjectToAdd
    )
    if (-not (Test-Connection)) {
        return
    }
    $endpoint = "/groups/$identity/members/`$ref"
    $memberJson = $null
    $graphResult = $null

    if ($null -ne $UserToAdd -and $Member.'@odata.type' -eq "#microsoft.graph.user") {
        $memberJson = ConvertTo-Json $UserToAdd
    }
    elseif ($null -ne $GroupToAdd) {
        Write-Error "Adding groups is not yet implemented."
        return $graphResult
    }
    elseif ($null -ne $OrganizationContactToAdd) {
        Write-Error "Adding organization contacts is not yet implemented."
        return $graphResult
    }
    elseif ($null -ne $DirectoryObjectToAdd) {
        Write-Error "Adding directory objects is not yet implemented."
        return $graphResult
    }

    if (-not [string]::IsNullOrEmpty($memberJson)) {
        $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method POST -Body $memberJson
    }
    return $graphResult
}
function Get-GraphPSO365GroupActivityDetail {
    param (
        [Parameter(Mandatory=$TRUE)]
        [ValidateSet("D7", "D30", "D90", "D180")]
        [string]$Period,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "reports/getOffice365GroupsActivityDetail(period='$Period')"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSSPSiteActivityUsageDetail {
    param (
        [Parameter(Mandatory=$TRUE)]
        [ValidateSet("D7", "D30", "D90", "D180")]
        [string]$Period,
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

    $endpoint = "reports/getSharePointSiteUsageDetail(period='$Period')"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSO365ActiveUserDetail {
    param (
        [Parameter(Mandatory=$TRUE)]
        [ValidateSet("D7", "D30", "D90", "D180")]
        [string]$Period,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "reports/getOffice365ActiveUserDetail(period='$Period')"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

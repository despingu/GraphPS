[string]$Script:tenantName = [string]::Empty
[string]$Script:clientID = [string]::Empty
[string]$Script:clientSecret = [string]::Empty
[string]$Script:resourceAppIdURI = [string]::Empty
[string]$Script:token = [string]::Empty
[datetime]$Script:tokenRenewTime = [datetime]::MinValue
[bool]$Script:connected = $false
[string]$Script:graphUri = "https://graph.microsoft.com"
[string]$Script:graphVersion = [string]::Empty

<#
function Get-GraphPSSample {
    param (
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method GET
    return $graphResult
}
#>

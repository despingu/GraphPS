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


function Get-GraphPSCredentialUserRegistrationDetails {
    param (
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression,
        [Parameter(Mandatory=$false)]
        [string]$FormatExpression
    )
    <#
        .PARAMETER $FilterExpression
        This function supports the optional OData query parameter $filter. You can apply $filter on one or more of the following properties of the credentialUserRegistrationDetails resource.

        Properties          Description and example
        userDisplayName	    Filter by user name. For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "userDisplayName eq 'Contoso'". Supported filter operators: eq, and startswith(). Supports case insensitive.
        userPrincipalName	Filter by user principal name. For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "userPrincipalName eq 'Contoso'". Supported filter operators: eq and startswith(). Supports case insensitive.
        authMethods	        Filter by the authentication methods used during registration. For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "authMethods/any(t:t eq microsoft.graph.registrationAuthMethod'email')". Supported filter operators: eq.
        isRegistered	    Filter for users who have registered for self-service password reset (SSPR). For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "isRegistered eq true". Supported filter operators: eq.
        isEnabled	        Filter for users who have been enabled for SSPR. For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "isEnabled eq true". Supported filtter operators: eq.
        isCapable	        Filter for users who are ready to perform password reset or multi-factor authentication (MFA). For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "isCapable eq true". Supported filter operators: eq
        isMfaRegistered	    Filter for users who are registered for MFA. For example: Get-GraphPSCredentialUserRegistrationDetails -filterExpression "isMfaRegistered eq true". Supported filter operators: eq.
    #>

    if (-not (Test-Connection -RequiredVersion beta)) {
        return
    }

    $endpoint = "reports/credentialUserRegistrationDetails"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSGetMailboxUsageDetail {
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

    $endpoint = "reports/getMailboxUsageDetail(period='$Period')"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}

function Get-GraphPSGetSharePointSiteUsageDetail {
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

    $endpoint = "reports/getSharePointSiteUsageDetail(period='$Period')"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -FormatExpression $FormatExpression -Method GET
    return $graphResult
}
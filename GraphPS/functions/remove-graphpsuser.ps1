function Remove-GraphPSUser {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory=$false)]
        [string]$selectExpression,
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    if (-not (Test-Connection)) {
        return
    }

    if(-not $Force) {
        $question = "Are you sure you want to delete user $identity?"
        $choices  = '&Yes', '&No'
        $decision = $Host.UI.PromptForChoice($null, $question, $choices, 1)
        if ($decision -eq 0) {
            return
        }
    }

    $endpoint = "users/$identity"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -Method DELETE
    return $graphResult
}
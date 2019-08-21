function Remove-GraphPSUser {
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    if (-not (Test-Connection)) {
        return
    }

    if(-not $Force) {
        $question = "Are you sure you want to delete user $($identity)?"
        $choices  = '&Yes', '&No'
        $decision = $Host.UI.PromptForChoice($null, $question, $choices, 1)
        if ($decision -eq 1) {
            return
        }
    }

    $endpoint = "users/$($identity)"

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method DELETE
    return $graphResult
}
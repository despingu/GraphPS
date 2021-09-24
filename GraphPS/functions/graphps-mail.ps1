function Get-GraphPSMessages {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory=$false)]
        [int]$top,
        [Parameter(Mandatory=$false)]
        [string]$filterExpression,
        [Parameter(Mandatory = $false)]
        [string]$selectExpression
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users/$identity/messages"
    $CustomExpressions = $null
    if ($null -ne $top) {
        $CustomExpressions = @("`$top=$top")
    }

    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -FilterExpression $filterExpression -SelectExpression $selectExpression -CustomExpressions $CustomExpressions -Method GET
    return $graphResult
}

function Send-GraphPSMail {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$identity,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$subject,
        [Parameter(Mandatory = $true, Position = 2)]
        [ValidateSet('Text','HTML')]
        [string]$contentType,
        [Parameter(Mandatory = $true, Position = 3)]
        [string]$body,
        [Parameter(Mandatory = $true, Position = 4)]
        [string[]]$toRecipients,
        [Parameter(Mandatory = $false)]
        [string[]]$ccRecipients,
        [Parameter(Mandatory = $false)]
        [switch]$saveToSentItems
    )
    if (-not (Test-Connection)) {
        return
    }

    $endpoint = "users/$($identity)/sendMail"
    [array]$toRecipientArray = Get-RecipientArray $toRecipients
    
   
    $messageObject = @{
        "subject" = $subject;
        "body" = @{
            "contentType" = $contentType;
            "content" = $body
        };
        "toRecipients" = $toRecipientArray;
    }

    if ($ccRecipients.Count -gt 0) {
        [array]$ccRecipientArray = Get-RecipientArray $ccRecipients
        $messageObject.Add("ccRecipients", $ccRecipientArray)
    }

    $bodyObject = @{
        "message" = $messageObject
        "saveToSentItems" = $saveToSentItems.IsPresent
    }

    $bodyJson = ConvertTo-Json -InputObject $bodyObject -Depth 5
    $graphResult = Invoke-MSGraphQuery -Endpoint $endpoint -Method Post -body $bodyJson
    return $graphResult
}

function Get-RecipientArray {
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string[]] $recipients
    )
    
    [array]$array = @()
    foreach ($recipient in $recipients) {
        $recipientObject = @{
            "emailAddress" = @{ 
                "address" = $recipient
            }
        }
        $array += $recipientObject
    }

    return $array
}
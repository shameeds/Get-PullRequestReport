<#
    This script is used to fetch the summary of pull requests that opened, closed and in-progress in the last week
    It requires the following parameters as input from user.
        1. patToken - PAT token must created in Github with repository read permission.
                      This PAT token should be saves as Github secrets if Github action is used to run this script.
        2. emailFrom - An Outlook email id which will be used to send the email. This is the username for smtp server - "smtp-mail.outlook.com"
        3. emailPassword - Password of the Outlook email Id
        4. emailTo - Receiver's email address.

    Note: We can create a pipeline in Github actions and schedule it to run weekly.
#>
param(
    [Parameter(Mandatory = $true, HelpMessage = "Github PAT token")]
    [String]
    $patToken,

    [Parameter(Mandatory = $true, HelpMessage = "Email sender Address - Outlook email id")]
    [String]
    $emailFrom,

    [Parameter(Mandatory = $true, HelpMessage = "Outlook email's password")]
    [String]
    $emailPassword,

    [Parameter(Mandatory = $true, HelpMessage = "Email receiver Address")]
    [String]
    $emailTo

)

$endpoint = "https://api.github.com/repos/github/codeql-variant-analysis-action/pulls?state=all"

$value = Invoke-WebRequest -Uri $endpoint -Headers @{"Authorization"="Basic $patToken"}

$jsonObjects = $value | ConvertFrom-Json

$prCreatedInLast7Days = $jsonObjects | Where-Object { [math]::Ceiling(((Get-Date)-([DateTime]$($_.created_at))).TotalDays) -le 7 }

$prClosedInLast7Days = $jsonObjects | Where-Object { if($_.state -eq "closed") { [math]::Ceiling(((Get-Date)-([DateTime]$($_.closed_at))).TotalDays) -le 7 }}

$prInOpenState = $jsonObjects | Where-Object { $_.state -eq "open" }

$prInDraftState = $prInOpenState | Where-Object { if($_.draft) { [math]::Ceiling(((Get-Date)-([DateTime]$($_.updated_at))).TotalDays) -le 7 }}

$summaryResults = @()
$summaryHeading = "<H2>Summary of pull requests for last week</H2>"
$summaryResults += @{
    "Status" = "Newly opened"
    "No. of pull request" = $prCreatedInLast7Days.Count
}
$summaryResults += @{
    "Status" = "Closed"
    "No. of pull request" = $prClosedInLast7Days.Count
}
$summaryResults += @{
    "Status" = "In-progress"
    "No. of pull request" = $prInDraftState.Count
}

$summary = $summaryResults | Select-Object "Status", "No. of pull request"
$summaryBody = $summary | ConvertTo-Html -PreContent $summaryHeading -Fragment | Out-String

$prDetails = @()
foreach($obj in $jsonObjects) {
    if([math]::Ceiling(((Get-Date)-([DateTime]$($obj.updated_at))).TotalDays) -le 7 ) {
        $prDetails += @{
            "Pull request Number" = $obj.number
            "Pull Request Name" = $obj.title
            "Current Status" = if($obj.draft) { "Work-In-Progress" } else { $obj.state }
            "Created By" = $obj.user.login
            "Source Branch Name" = $obj.head.ref
            "Creation Time" = $obj.created_at
            "Closed at" = $obj.closed_at
        }
    }
}
$prs = $prDetails | Select-Object "Pull request Number", "Pull Request Name", "Current Status", "Created By", "Source Branch Name", "Creation Time", "Closed at"
$prBody = $prs | ConvertTo-Html -Fragment | Out-String

$openPrResults = @()
$openPRHeading = "<H3>Pull requests that are in open state - details</H3>"
foreach($openPR in $prInOpenState) {
    $openPrResults += @{
        "Pull request Number" = $openPR.number
        "Pull Request Name" = $openPR.title
        "Current Status" = if($openPR.draft) { "Work-In-Progress" } else { $openPR.state }
        "Created By" = $openPR.user.login
        "Source Branch Name" = $openPR.head.ref
        "Creation Time" = $openPR.created_at
    }
}
$openPRs = $openPrResults | Select-Object "Pull request Number", "Pull Request Name", "Current Status", "Created By", "Source Branch Name", "Creation Time"
$openPrBody = $openPRs | ConvertTo-Html -PreContent $openPRHeading -Fragment | Out-String

$header = @"
<html>
<head>
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
</head>
<body>
"@

$content = $header + $summaryBody + "<br>" + $prBody + "<br>" + $openPrBody + "</body></html>"

$subject = "Pull request summary report for last week"
$password = ConvertTo-SecureString -force -AsPlainText -string $emailPassword
$credential = New-Object System.Management.Automation.PSCredential $emailFrom, $password
$emailParameters = @{
    From        = $emailFrom
    To          = $emailTo
    Subject     = $subject
    SMTPServer  = "smtp-mail.outlook.com"
    Port        = 587
    credential  = $credential
    body        = $content
}
Send-MailMessage @emailParameters -UseSsl -BodyAsHtml -WarningAction Ignore


Write-Output "`nFrom: $($emailFrom)"
Write-Output "To: $($emailTo)"
Write-Output "Subject: $($subject)`n"

# Email content
Write-Output "Summary of pull requests for last week."

$summaryResults | Select-Object "Status", "No. of pull request" | Format-Table
$prs | Format-Table
Write-Output "`nPull requests that are in open state - details"
$openPRs | Format-Table

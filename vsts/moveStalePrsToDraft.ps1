# script to move pull requests older than 60 days to draft

param (
    [Parameter(Mandatory=$true)]
    [string]$OrganizationUrl,       # URL of the organization

    [Parameter(Mandatory=$true)]
    [string]$ProjectName,           # Name of the project

    [Parameter(Mandatory=$true)]
    [string]$RepositoryName,        # Name of the repository

    [Parameter(Mandatory=$true)]
    [string]$PersonalAccessToken    # Personal access token for authentication
)

$headers=@{
    'Authorization' = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PersonalAccessToken)"))
    'Content-Type' = "application/json"
}

# Get repository ID
$uri = "$OrganizationUrl/$ProjectName/_apis/git/repositories?api-version=6.0"
$repositories = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get
$repositoryId = ($repositories.value | Where-Object { $_.name -eq $RepositoryName }).id

Write-Host "Repository ID: $repositoryId"

# Get all active pull requests using Azure REST API
$uri = "$OrganizationUrl/$ProjectName/_apis/git/repositories/$repositoryId/pullrequests?searchCriteria.status=active&api-version=6.0"
$pullRequests = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get

# Iterate through each pull request
foreach ($pr in $pullRequests.value) {
    Write-Host $pr.title - $pr.creationDate

    $pullRequestId = $pr.pullRequestId
    $creationDate = $pr.creationDate
    $daysSinceCreation = (Get-Date) - [DateTime]$creationDate

    # Move pull requests older than 60 days to draft
    if ($daysSinceCreation.TotalDays -gt 60) {
        $uri = "$OrganizationUrl/$ProjectName/_apis/git/repositories/$repositoryId/pullrequests/$pullRequestId" + "?api-version=7.1"
        $body = @{
            isDraft = "true"
        } | ConvertTo-Json

        Invoke-RestMethod -Uri $uri -Headers $headers -Method Patch -Body $body

        Write-Host "Pull request $pullRequestId has been moved to draft."
    }
}

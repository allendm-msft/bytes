[CmdletBinding()]
param (
    [Parameter()]
    [Int32]
    $DurationInDays = 7,
    [Parameter()]
    [Boolean]
    $SendEmail = $false
)

$personalAccessToken = "<your pat token>"

$headers=@{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($personalAccessToken)"))
}

# get team members
$getTeamMembersUrl = 'https://dev.azure.com/dynamicscrm/_apis' + '/projects/<your project guid>/teams/<your team guid>/members?api-version=5.1'
$teamMembersResponse = Invoke-RestMethod -Uri $getTeamMembersUrl -Method Get -Headers $headers

# get team member individual descriptors
$teamMemberIdentities = @();
foreach ($teamMember in $teamMembersResponse.value.identity)
{
    if ($teamMember.isContainer)
    {
        $membersInGroupUrl = 'https://vssps.dev.azure.com/dynamicscrm/_apis/graph/Memberships/' + $teamMember.descriptor + '?direction=Down&api-version=5.1-preview.1'
        $membersInGroup = Invoke-RestMethod -Uri $membersInGroupUrl -Method Get -Headers $headers
        foreach ($memberInGroup in $membersInGroup.value) 
        {
            Write-Debug $memberInGroup.memberDescriptor
            $teamMemberIdentities += $memberInGroup.memberDescriptor
        }
        continue;
    }
    Write-Debug $teamMember.descriptor
    $teamMemberIdentities += $teamMember.descriptor
}

$teamMemberIdentities = $teamMemberIdentities | Sort-Object -Unique -CaseSensitive
# Write-Debug $teamMemberIdentities
Write-Debug $teamMemberIdentities.Count

# get team member aad ids
$members = @{}
# $teamMemberIds = @()
# $teamMemberNames = @()
foreach ($teamMemberIdentity in $teamMemberIdentities)
{
    $member = New-Object System.Object
    $userGraphUrl = 'https://vssps.dev.azure.com/dynamicscrm/_apis/graph/users/' + $teamMemberIdentity + '?api-version=5.1-preview.1'
    $user = Invoke-RestMethod -Uri $userGraphUrl -Method Get -Headers $headers

    $userStorage = Invoke-RestMethod -Uri $user._links.storageKey.href -Method Get -Headers $headers


    $member | Add-Member -type NoteProperty -name Name -value $user.displayName
    $member | Add-Member -type NoteProperty -name id -value $userStorage.value
    $member | Add-Member -type NoteProperty -name descriptor -value $teamMemberIdentity
    $member | Add-Member -type NoteProperty -name email -value $user.mailAddress
    $member | Add-Member -type NoteProperty -name PRsStarted -value 0
    $member | Add-Member -type NoteProperty -name PRsCompleted -value 0
    $member | Add-Member -type NoteProperty -name ReviewsDone -value 0
    $member | Add-Member -type NoteProperty -name PRsStarted7days -value 0
    $member | Add-Member -type NoteProperty -name PRsCompleted7days -value 0
    $member | Add-Member -type NoteProperty -name ReviewsDone7days -value 0


    $members.Add($userStorage.value, $member)
}
# Write-Debug $teamMemberIds
Write-Debug $members.Count

$lastDateTime = [System.DateTime]::UtcNow.AddDays($DurationInDays * -1)
$lastDateTime7days = [System.DateTime]::UtcNow.AddDays(-7)
$totalCompletedPRs = 0
$totalCompletedPRs7days = 0
# get completed pull requests by teammemberid
foreach ($memberId in $members.Keys)
{
    $prsUrl = 'https://dev.azure.com/dynamicscrm/onecrm/_apis/git/pullrequests?searchCriteria.status=completed&searchCriteria.creatorId=' + $memberId + '&api-version=5.1'
    $prs = Invoke-RestMethod -Uri $prsUrl -Method Get -Headers $headers

    foreach ($pr in $prs.value)
    {
        if((Get-Date $pr.closedDate) -lt $lastDateTime)
        {
            continue
        }
        
        Write-Debug $pr.closedDate

        $members[$memberId].PRsCompleted += 1
        $totalCompletedPRs += 1

        if((Get-Date $pr.closedDate) -ge $lastDateTime7days)
        {
            $totalCompletedPRs7days += 1
            $members[$memberId].PRsCompleted7days += 1
        }

        foreach ($reviewer in $pr.reviewers)
        {
            # ignore self reviews
            if ($memberId -eq $reviewer.id)
            {
                continue
            }

            if($reviewer.vote -ne 0) # 0 is no review done
            {
                if($members.ContainsKey($reviewer.id))
                {
                    $members[$reviewer.id].ReviewsDone += 1

                    if((Get-Date $pr.closedDate) -ge $lastDateTime7days)
                    {
                        $members[$reviewer.id].ReviewsDone7days += 1
                    }
                }
            }
        }
    }

}

foreach ($memberId in $members.Keys)
{
    $name = $members[$memberId].Name 
    if ($name -contains "Allen" -or $name -contains "Natalie" -or $name -contains "Daniel")
    {
        continue
    }
    Write-Host $members[$memberId].Name
    Write-Host $members[$memberId].id
    Write-Host $members[$memberId].email
    Write-Host $members[$memberId].descriptor
    Write-Host $members[$memberId].PRsStarted
    Write-Host $members[$memberId].PRsCompleted
    Write-Host $members[$memberId].ReviewsDone
    Write-Host $members[$memberId].PRsStarted7days
    Write-Host $members[$memberId].PRsCompleted7days
    Write-Host $members[$memberId].ReviewsDone7days

    if($SendEmail)
    {
        # send email to engg and cc me in
        Write-Host Hi $members[$memberId].Name here is PR summary for past 30 and 7 days
        $members[$memberId] | Format-Table -Property Name, PRsCompleted, ReviewsDone, PRsCompleted7days, ReviewsDone7days

    }
}

# send full summary to me
Write-Host This is a summary of work done fo the past $DurationInDays days
Write-Host Total PRs completed in the team $totalCompletedPRs
$members.Values | Format-Table -Property Name, PRsCompleted, ReviewsDone, PRsCompleted7days, ReviewsDone7days

Write-Host Self reviews are ignored





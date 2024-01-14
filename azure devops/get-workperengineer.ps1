
$personalAccessToken = "<your pat token>"
$tfsUrl= "https://<org>.visualstudio.com/<project>/_apis"
$tfsUrl2= "https://<org>.visualstudio.com/_apis"
$headers=@{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($personalAccessToken)"))
}

# Get user stories for current sprint
# create a flat query "currentsprinttasksflat" for getting all tasks
$activeBugsQuery=Invoke-RestMethod -Uri $tfsUrl'/wit/queries/My%20Queries/CurrentSprintTasksFlat?$depth=1&api-version=2.2' -Method Get -Headers $headers
$activeBugsUri=$tfsUrl + '/wit/wiql/' + $activeBugsQuery.id + '?api-version=1.0'
$activeBugs=Invoke-RestMethod -Uri $activeBugsUri -Method Get -Headers $headers

# Get details of all user stories
$activeBugsCommaSep=$activeBugs.workItems.id -join ","
$activeBugsUriFull=$tfsUrl2 + '/wit/workitems?ids=' + $activeBugsCommaSep + '&$expand=all&api-version=1.0'
$activeBugsFull=Invoke-RestMethod -Uri $activeBugsUriFull -Method Get -Headers $headers

$assigneds= $activeBugsFull.value.fields | Select-Object System.AssignedTo, Microsoft.VSTS.Scheduling.OriginalEstimate, Microsoft.VSTS.Scheduling.CompletedWork, Microsoft.VSTS.Scheduling.RemainingWork
$assignedTos=$assigneds.'System.AssignedTo' | sort-object| Get-Unique


$teamData=@()
foreach($assigned in $assignedTos)
{
    if($assigned)
    {
        $nameOriginal = ($assigneds | where { ($_.'System.AssignedTo' -eq $assigned) } | Measure-Object -Property 'Microsoft.VSTS.Scheduling.OriginalEstimate' -Sum).Sum
        $nameCompleted = ($assigneds | where { ($_.'System.AssignedTo' -eq $assigned) } | Measure-Object -Property 'Microsoft.VSTS.Scheduling.CompletedWork' -Sum).Sum
        $nameRemaining = ($assigneds | where { ($_.'System.AssignedTo' -eq $assigned) } | Measure-Object -Property 'Microsoft.VSTS.Scheduling.RemainingWork' -Sum).Sum

        
        Write-Host "Engineer $assigned"
        Write-Host "Original Estimate: $nameOriginal"
        Write-Host "Completed: $nameCompleted"
        Write-Host "Remaining: $nameRemaining"
        
        $prop = New-Object System.Object
        $prop | Add-Member -type NoteProperty -name Engineer -value $assigned 
        $prop | Add-Member -type NoteProperty -name EffortEstimated -value $nameOriginal 
        $prop | Add-Member -type NoteProperty -name EffortCompleted -value $nameCompleted
        $prop | Add-Member -type NoteProperty -name EffortRemaining -value $nameRemaining
        $teamData+=$prop
    }
}

# Given a summary of work completed
# Name : Total effort completed: Total effort Completed since last email

$teamData | Sort-Object -Property Engineer -Descending | ConvertTo-Html > .\body.html
Write-Output "<br><br><br>"  >> .\body.html
Write-Output "* All effort data are in days <br>"  >> .\body.html
Write-Output "** If your name is not here you dont have tasks assigned to you. Fix it or reach out to me"  >> .\body.html

$body=Get-Content .\body.html

$username="me@domain.com"
$pwd = Get-Content e:\scripts\.mypwd | ConvertTo-SecureString
$mycred=New-Object System.Management.Automation.PSCredential -ArgumentList $username, $pwd
$dateOnly=[System.DateTime]::Today.Date.ToString("d")
Send-MailMessage -From $username -To my@email.com -Subject "Weekly Stats Email [$dateOnly]" -SmtpServer "smtp.office365.com" -Credential $mycred -Port 587 -UseSsl -Body "$body" -BodyAsHtml

Write-Host "Sent email successfully!"
rm body*.html


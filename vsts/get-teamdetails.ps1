
$personalAccessToken = "<pat token>"
$tfsUrl= "https://mseng.visualstudio.com/VSOnline/_apis"
$tfsUrl2= "https://mseng.visualstudio.com/_apis"
$headers=@{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($personalAccessToken)"))
}

# create the active bugs query
$activeBugsQuery=Invoke-RestMethod -Uri $tfsUrl'/wit/queries/Shared%20Queries/VSDT%20TT/DTA/Active%20Bugs?$depth=1&api-version=2.2' -Method Get -Headers $headers
$activeBugsUri=$tfsUrl + '/wit/wiql/' + $activeBugsQuery.id + '?api-version=1.0'
$activeBugs=Invoke-RestMethod -Uri $activeBugsUri -Method Get -Headers $headers

$activeBugsCommaSep=$activeBugs.workItems.id -join ","
$activeBugsUriFull=$tfsUrl2 + '/wit/workitems?ids=' + $activeBugsCommaSep + "&api-version=1.0"
$activeBugsFull=Invoke-RestMethod -Uri $activeBugsUriFull -Method Get -Headers $headers



$activeBugsDetails=$activeBugsFull.value.fields

$assigneds= $activeBugsDetails | Select-Object System.AssignedTo, System.Title, System.CreatedDate, Microsoft.VSTS.Common.Priority
$assignedTos=$assigneds.'System.AssignedTo' | Unique

$nameCounts=@{}
$nameStales=@{}
$nameP0P1s=@{}

foreach($assigned in $assignedTos)
{
    if($assigned)
    {
    $nameCounts.Add($assigned, ($assigneds | where { $_.'System.AssignedTo' -eq $assigned } | measure).Count)
    $nameStales.Add($assigned, ($assigneds | where { ($_.'System.AssignedTo' -eq $assigned) -and ([System.DateTime]::Parse($_.'System.CreatedDate') -lt [System.DateTime]::Now.AddDays(-21)) } | measure).Count)
    $nameP0P1s.Add($assigned, ($assigneds | where { ($_.'System.AssignedTo' -eq $assigned) -and (($_.'Microsoft.VSTS.Common.Priority' -eq "0") -or ($_.'Microsoft.VSTS.Common.Priority' -eq "1")) } | measure).Count)
    }
}

Write-Output "Total Bugs: " $activeBugsFull.count "<br>" > .\body1.html
Write-Output "Unassigned: " ($assigneds | where { [System.String]::IsNullOrWhiteSpace($_.'System.AssignedTo') }).Count "<br>"  >> .\body1.html
Write-Output "Total P0P1 Bugs: " ($assigneds | where { ($_.'Microsoft.VSTS.Common.Priority' -eq 0) -or ($_.'Microsoft.VSTS.Common.Priority' -eq 1) }).Count "<br>"  >> .\body1.html
Write-Output "Stale Bugs: " ($assigneds |  where { [System.DateTime]::Parse($_.'System.CreatedDate') -lt [System.DateTime]::Now.AddDays(-21) }).Count "<br>"  >> .\body1.html
Write-Output "<br>" >> .\body1.html
#Write-Host "Engineer    ActiveBugCount  P0P1BugCount    StaleBugCount"
$name=@()
foreach($assigned in $assignedTos)
{
    if($assigned)
    {
        $prop = New-Object System.Object
        $prop | Add-Member -type NoteProperty -name Engineer -value $assigned 
        $prop | Add-Member -type NoteProperty -name ActiveBugCount -value $nameCounts[$assigned] 
        $prop | Add-Member -type NoteProperty -name P0P1BugCount -value $nameP0P1s[$assigned] 
        $prop | Add-Member -type NoteProperty -name StaleBugCount -value $nameStales[$assigned]
        $name+=$prop
        #$name+=@{"Engineer"=$assigned;"ActiveBugCount"=$nameCounts[$assigned];"P0P1Count"=$nameP0P1s[$assigned];"StaleBugCount"=$nameStales[$assigned]}
        #$name.Add($properties)
        #$name.Add($properties)
        # Write-Host $assigned    $nameCounts[$assigned]  $nameP0P1s[$assigned]   $nameStales[$assigned]
    }
}
$name | Sort-Object -Property ActiveBugCount -Descending | ConvertTo-Html > .\body2.html
$assigneds | ConvertTo-Html > .\body3.html

$body1=Get-Content .\body1.html
$body2=Get-Content .\body2.html
$body3=Get-Content .\body3.html

$username="me@domain.com"
$pwd = Get-Content d:\ps1\.mypwd | ConvertTo-SecureString
$mycred=New-Object System.Management.Automation.PSCredential -ArgumentList $username, $pwd
Send-MailMessage -From allendm@microsoft.com -To vsttdta@microsoft.com -Subject "Bug stats" -SmtpServer "smtp.office365.com" -Credential $mycred -Port 587 -UseSsl -Body "$body1 </br> $body2" -BodyAsHtml


rm body*.html

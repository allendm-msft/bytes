# set-executionpolicy -UnRestricted
# "P@ssword1" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File ".\.mypwd"

$token='<pat token>'
$b64token=[System.Convert]::ToBase64String([char[]]$token)

$headers=@{
 Authorization = 'Basic {0}' -f $b64token
}

$otissresponse=Invoke-WebRequest -Headers $headers -Uri https://api.github.com/repos/Microsoft/vsts-tasks/issues?state=open"&"labels=Area%3A%20Test -Method Get
$otiss=$otissresponse.Content | ConvertFrom-Json

Write-Host "we currently have " $otiss.Count " issues"
$otiss | Sort-Object created_at -descending | Format-List -Property title,created_at,@{Label="AssignedTo"; Expression={$_.assignee.login} }


$otiss | Sort-Object created_at -descending | Select-Object title,created_at,@{Label="AssignedTo"; Expression={$_.assignee.login} } | ConvertTo-Html > d:\ps1\temp.html
$body=Get-Content d:\ps1\temp.html

$username="me@domain.com"
$pwd = Get-Content d:\ps1\.mypwd | ConvertTo-SecureString
$mycred=New-Object System.Management.Automation.PSCredential -ArgumentList $username, $pwd
Send-MailMessage -From allendm@microsoft.com -To team@domain.com -Subject "Active Test Issues in Microsoft/vsts-tasks github repo" -SmtpServer "smtp.office365.com" -Credential $mycred -Port 587 -UseSsl -Body "$body" -BodyAsHtml


rm temp*.html
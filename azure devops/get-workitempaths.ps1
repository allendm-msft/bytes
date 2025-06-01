
$personalAccessToken = '<your pat>' 
$tfsUrl= "https://<your ado>.visualstudio.com/<your project>/_apis"
$tfsUrl2= "https://<your ado>.visualstudio.com/_apis"
$headers=@{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($personalAccessToken)"))
}
function Get-Children {
    param($ChildList)
   
    Foreach ($child in $childList) {
      # Write-Output $child

      if($child.hasChildren) {
        Get-Children $child.children
      }
      # Write-Output "hello"
      Write-Output $child.path
    }
}

# Get user stories for current sprint
# create a flat query "currentsprinttasksflat" for getting all tasks
$activeBugsQuery=Invoke-RestMethod -Uri $tfsUrl'/wit/classificationnodes?ids=<id>&$depth=10&api-version=4.1' -Method Get -Headers $headers

Get-Children -ChildList $activeBugsQuery.value.children

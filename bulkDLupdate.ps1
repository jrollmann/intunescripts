Connect-ExchangeOnline 
Import-CSV <FilePath> | foreach {  
 $UPN=$_.UPN 
 Write-Progress -Activity "Adding $UPN to group… " 
 Add-DistributionGroupMember –Identity <GroupUPN> -Member $UPN  
 If($?)  
 {  
 Write-Host $UPN Successfully added -ForegroundColor Green 
 }  
 Else  
 {  
 Write-Host $UPN - Error occurred –ForegroundColor Red  
 }  
} 

# ImportFromCsvToOutlook    
* Import contacts CSV file exported form active directory to outlook (2016 , 2019)     

### Export vcs From Active Directory
Command    
>Get-ADUser -Filter 'mail -like "*<domain>"' -SearchBase "<search base>"  -Properties * | Select -Property   DisplayName,UserPrincipalName | Export-CSV "<csv file location>" -NoTypeInformation -Encoding UTF8    

Example    
Get-ADUser -Filter 'mail -like "*test.com"' -SearchBase "OU=taipei,DC=test,DC=com"  -Properties * | Select -Property   DisplayName,UserPrincipalName | Export-CSV "c:\aaa.csv" -NoTypeInformation -Encoding UTF8    

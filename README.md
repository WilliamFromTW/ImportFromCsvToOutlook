# ImportFromCsvToOutlook    
* Import contacts to outlook (2016 , 2019)     

### Prerequisite    
* Active Directory 2008R2 or above
* visual studio 2019 community version
* outlook 2016 , 2019

### Export vcs From Active Directory by power shell
Command    
>Get-ADUser -Filter 'mail -like "*<domain>"' -SearchBase "<search base>"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "<csv file location>" -NoTypeInformation -Encoding UTF8    

Example    
>Get-ADUser -Filter 'mail -like "*test.com"' -SearchBase "OU=taipei,DC=test,DC=com"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "c:\aaa.csv" -NoTypeInformation -Encoding UTF8    

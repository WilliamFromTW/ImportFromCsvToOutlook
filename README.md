# ImportFromCsvToOutlook    
* Import contacts to outlook (2016 , 2019)     
  This APP will read csv file that export from active directory and import to contacts of outlook 2016 or 2019     

### Prerequisite    
* Develop    
  visual studio 2019 community version    
  - reference outlook module (lib)    
* server     
  Active Directory 2008R2 or above
* User Client    
  outlook 2016 , 2019

### Export vcs From Active Directory by power shell
Command    
>Get-ADUser -Filter 'mail -like "*\<domain\>"' -SearchBase "\<search base\>"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "\<csv file location\>" -NoTypeInformation -Encoding UTF8    

Example    
>Get-ADUser -Filter 'mail -like "*test.com"' -SearchBase "OU=taipei,DC=test,DC=com"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "c:\aaa.csv" -NoTypeInformation -Encoding UTF8    


### command mode    
Default installed path     
> C:\InMethod\ImportContactsToOutlook    

Example     
> C:\InMethod\ImportContactsToOutlook\ImportFromCsvToOutlook.exe   c:\xxx.csv 

# ImportFromCsvToOutlook    
* Import contacts to outlook 2013 or above     
  This APP will read csv file that export from active directory and import to contacts of outlook     

### Prerequisite    
* Develop    
  visual studio 2019 community version    
  - reference outlook module (lib)    
  - dot net framework 4.7.2 or above    
  
* server     
  Active Directory 2008R2 or above    
* User Client    
  outlook 2013 or above    

### Export vcs From Active Directory by power shell
Command    
>Get-ADUser -Filter 'mail -like "*\<domain\>"' -SearchBase "\<search base\>"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "\<csv file location\>" -NoTypeInformation -Encoding UTF8    

Example    
>Get-ADUser -Filter 'mail -like "*test.com"' -SearchBase "OU=taipei,DC=test,DC=com"  -Properties * | Select -Property   DisplayName,GivenName,Surname,mail,Title,Department,Office | Export-CSV "c:\aaa.csv" -NoTypeInformation -Encoding UTF8    


### Command mode    
Default installed path     
> C:\InMethod\ImportContactsToOutlook    

Example 1    
> C:\InMethod\ImportContactsToOutlook\ImportFromCsvToOutlook.exe   c:\xxx.csv 

Example 2    
> C:\InMethod\ImportContactsToOutlook\ImportFromCsvToOutlook.exe  https://website/xxx.csv 


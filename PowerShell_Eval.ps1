# 1. Import csv as a variable
$filepath = 'D:\Powershell_work\Take_Two'
$usercsv = Import-Csv $filepath\Users.csv


# 2. How many users?
Write-Host "There are $($usercsv.count) users in the csv file." -ForegroundColor Yellow

# 3. What is the total size of all mailboxes
$totalsize = $usercsv | Measure-Object -Property MailboxSizeGB -Sum
Write-Host "The total size of the mailboxes is $($totalsize.Sum)GB." -ForegroundColor Yellow

# 4. How many accounts exist with non identical EmailAddress/UserPrinicapName
# Case sensitive
$different_casesensitve = ($usercsv | Where-Object {$_.EmailAddress -cne $_.UserPrincipalName}).count
# Case insensitive
$different_caseinsensitve = ($usercsv | Where-Object {$_.EmailAddress -ne $_.UserPrincipalName}).count
Write-Host "There are $different_casesensitve different EmailAddress to UserPrincipalName, of which $($different_casesensitve-$different_caseinsensitve) are due to case sensitivity." -ForegroundColor Yellow

# 5. What is the total size of all mailboxes in NYC
$totalnycsize = $usercsv | Where-Object {$_.Site -eq "NYC"} | Measure-Object -Property MailboxSizeGB -Sum
Write-Host "The total size of the mailboxes in site NYC is $($totalnycsize.Sum)GB." -ForegroundColor Yellow

# 6. How many Employees have mailboxes larger than 10 GB
# Number of mailboxs greater or equal to 10 GB
$mbge10 = ($usercsv | Where-Object {$_.AccountType -eq 'Employee' -and $_.MailboxSizeGB -ge 10}).count
# Number of mailboxes greater than 10 GB
$mbgt10 = ($usercsv | Where-Object {$_.AccountType -eq 'Employee' -and $_.MailboxSizeGB -gt 10}).count
Write-Host "The number of mailboxes greater or equal to 10 GB is $mbge10, of which $($mbge10 - $mbgt10) are 10 GB in size." -ForegroundColor Yellow

# 7. Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending
# get the top 10
$domain2info = $usercsv | Where-Object {$_.EmailAddress -match "domain2.com"} | Sort-Object -Property MailboxSizeGB -Descending |  Select-Object -First 10
# now just the username
$usernames = ($domain2info.EmailAddress -replace("@domain2.com","")) -join " "
Write-Host $usernames -ForegroundColor Yellow

# 8. Create a new CSV file that summarizes Sites, using the following headers: Site, TotalUserCount, EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB
# get the unique sites
$usersites = $usercsv.Site | Select-Object -Unique
# set up a empty array
$siteinfo = @()
foreach ($site in $usersites){
    
$mbinfo = $usercsv | Where-Object {$_.Site -eq $site} | Measure-Object -Property MailboxSizeGB -Sum -Average
$empcount = ($usercsv | Where-Object {$_.Site -eq $site -and $_.Accounttype -eq 'Employee'}).count
$sinfo = [PSCustomObject]@{
    Site = $site
    TotalUserCount = $mbinfo.Count
    EmployeeCount = $empcount
    ContractorCount = ($mbinfo.Count - $empcount)
    TotalMailboxSizeGB = $mbinfo.Sum
    AverageMailboxSizeGB = [math]::Round($mbinfo.Average,1)
}
$siteinfo += $sinfo
$sinfo = $null
$mbinfo = $null
$empcount = $null

}
# write out the information
$siteinfo | export-csv $filepath\answer.csv -NoTypeInformation
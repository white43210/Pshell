[CmdletBinding()]
Param(
  [Parameter(Mandatory=$False,Position=1)]
   [string]$filePath = 'c:\temp\inactive_users.htm',	

   [Parameter(Mandatory=$False)]
   [string]$targetOU = 'OU=Domain Users,DC=BWWB,DC=pri'
)

$header = @"
<style>
body {font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;font-size: 10px;}

table{border-collapse: collapse;border: solid 1px;font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;color: black;margin-left: 20px;margin-bottom: 10px;width: 730px;}
table td{font-size: 10px;padding-left: 7px;text-align: left;border: solid 1px;}
table td:first-child{width: 150px;}
table td:nth-child(2){width: 325px;}
table td:nth-child(3){width: 100px;}
table td:nth-child(4){width: 55px;text-align: center;}
table td:nth-child(5){width: 100px;}
table td:nth-child(6){width: 180px;}
 
table tr th {font-size: 10px;font-weight: bold;padding-left: 0px;text-align: center;background: #BBBBBB;}
table tr {page-break-inside: auto;}
table tr:nth-child(odd){background: #CCCCCC;}
table tr:nth-child(even){background: #F2F2F2;}
</style>
"@

$cols =
 @{n="User ID";e={$_.samAccountName}},
 @{n="Department";e={$_.Department}},
 @{n="Last Logon";e={If($_.LastLogonTimestamp -eq $Null){"NA"}Else{[datetime]::FromFileTime($_.LastLogonTimestamp).ToString("d")}}},
 @{n="Locked";e={If($_.LockedOut){"Yes"}Else{"No"}}},
 @{n="Last Bad Logon";e={If($_.LastBadPasswordAttempt -eq $Null){"NA"}Else{$_.LastBadPasswordAttempt.ToShortdatestring()}}}

Search-ADAccount -AccountInactive -DateTime ((get-date).adddays(-90)) -UsersOnly -SearchBase $trgtOU |
 Get-ADUser -Properties SamAccountName,Department,LastLogonTimeStamp,LockedOut,LastBadPasswordAttempt | 
 sort -property LastLogonTimeStamp -Descending | 
 where {($_.Enabled -eq 'True') <#-and ($_.Description -like 'User Account - *')#>} | 
 select $cols |
 ConvertTo-Html -Head $header -Title "User Account Report" -PreContent "<H3>Inactive User Accounts (90) Days</H3>" -PostContent "Generated: $(Get-Date)" |
 Out-File $filePath -Encoding ascii

 Invoke-Expression $filePath

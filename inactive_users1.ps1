<#
.SYNOPSIS
    Returns all inactive users in the domain.
.DESCRIPTION
    This script will evaluate enabled users in the specified OU, and if they 
    have not logged in within defined period of time add them to an html report.
.PARAMETER filepath
    [optional] provides output path AND name of html file [default is 'c:\temp\inactive_users.htm']
.PARAMETER targetOU
    [optional] provides scope/filter [default is 'OU=Domain Users,DC=BWWB,DC=pri']
.EXAMPLE
    .\inactive-users.ps1
.EXAMPLE
    .\inactive-users.ps1 -filepath 'c:\temp\inactive_users.htm'
.EXAMPLE
    .\inactive-users.ps1 -targetOU 'OU=Information Technology,OU=Finance & Administration,OU=Domain Users,DC=BWWB,DC=pri'
.NOTES
    Written by: Chad White
#>
[CmdletBinding()]
Param(
  [Parameter(Mandatory=$False,Position=1)]
   [string]$filePath = 'c:\temp\inactive_users.htm',	

   [Parameter(Mandatory=$False)]
   [string]$targetOU = 'OU=Domain Users,DC=BWWB,DC=pri'
)

#region Data gathering
 
 $cols =
 @{n="User ID";e={$_.samAccountName}},
 @{n="Department";e={$_.Department}},
 @{n="Last Logon";e={If($_.LastLogonTimestamp -eq $Null){"NA"}Else{[datetime]::FromFileTime($_.LastLogonTimestamp).ToString("d")}}},
 @{n="Locked";e={If($_.LockedOut){"Yes"}Else{"No"}}},
 @{n="Last Bad Logon";e={If($_.LastBadPasswordAttempt -eq $Null){"NA"}Else{$_.LastBadPasswordAttempt.ToShortdatestring()}}}

 $Users = Search-ADAccount -AccountInactive -DateTime ((get-date).adddays(-90)) -UsersOnly -SearchBase $targetOU |
 Get-ADUser -Properties SamAccountName,Department,LastLogonTimeStamp,LockedOut,LastBadPasswordAttempt | 
 sort -property LastLogonTimeStamp -Descending | 
 where {($_.Enabled -eq 'True') <#-and ($_.Description -like 'User Account - *')#>} | 
 select $cols

 $Report = 'Inactive_Users.htm'

#endregion Data gathering

#region highlighting users with inactive accounts
 Add-Type -AssemblyName System.Xml.Linq 

 $head = @"
 <style>
 h6 {
 text-align:center;
 border-bottom:1px solid #666666;
 color:blue;
 }
 
 body {
 font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;
 font-size: 10px;
 font-smoothing: always;
 width: 100%
 text-align:left;
 }
 
 table {
 table-layout: fixed;
 border-collapse: collapse;
 border: solid 1px black;
 color: black;
 width: 100%;
 }
 
 * {
 margin:0;
 }
 
 .pageholder {
 margin:0px auto
 }
 
 td{
 vertical-align: Top;
 padding-left: 7px;
 border: solid 1px;
 word-wrap: break-word;
 }
 
 th{
 vertical-align: Top;
 padding-left: 7px;
 text-align: center;
 color:Black;
 background-color:LightSteelBlue;
 border: solid 1px black;
 }
 
 #table td:first-child{width: 150px;}
 #table td:nth-child(2){width: 325px;}
 #table td:nth-child(3){width: 100px;}
 #table td:nth-child(4){width: 55px;text-align: center;}
 #table td:nth-child(5){width: 100px;}
 #table td:nth-child(6){width: 180px;}
  
 .odd{ 
 background: #CCCCCC;
 }
 
 .even{
 background: #F2F2F2;
 }
 </style>
"@ 
 
 $pre = @"
 <H3>Inactive Users Report (90) Days</H3></br>Scope: *$targetOU"
"@

 $body = $Users | ConvertTo-Html -Fragment | Out-String
#region Linq parsing
 $xml = [System.Xml.Linq.XDocument]::Parse( $body)
 
 if($Namespace = $xml.Root.Attribute("xmlns").Value)  {
$Namespace = "{{{0}}}" -f $Namespace
 }
 
 $index = [Array]::IndexOf( $xml.Descendants("${Namespace}th").Value, "Locked")
 $i = 0
 foreach($row in $xml.Descendants("${Namespace}tr"))  {
 if ($i % 2) {
 Write-Verbose 'Set even' -Verbose
 $row.SetAttributeValue("class","even")
 } Else {
 Write-Verbose 'Set odd' -Verbose
 $row.SetAttributeValue("class","odd")
 }
 switch(@($row.Descendants("${Namespace}td"))[$Index])  {
 {'Yes' -eq $_.Value } {
 Write-Verbose 'Set red' -Verbose
 $_.SetAttributeValue("style","background: red;")
 continue
 }
 }
 $i++
 }
 $Body = $xml.Document.ToString()
 #endregion Linq parsing
$HTML = $pre, $body
 $post = "<br><i>Report Generated on: $((Get-Date).ToString())</i>"
 ConvertTo-Html -Head $head -PostContent $post -Body $HTML | Out-String | Out-File $filePath

 Invoke-Expression $filePath
 #endregion highlighting users with inactive accounts

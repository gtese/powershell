# powershell

set-executionpolicy unrestricted
Set-Location $home\Desktop
$Date=(get-date).ToString('yyyyMMdd')


##List User

Get-ADUser -SearchBase "OU=策源地产,OU=复星集团下属企业,DC=fosun,DC=com" -Filter {(Enabled -ne $false)} | select name,userPrincipalName,DistinguishedName |export-csv  report$date.csv -NoTypeInformation -Encoding UTF8

Get-ADUser -SearchBase "OU=上海云济信息科技有限公司,DC=fosun,DC=com" -Filter {(Enabled -ne $false)} -Properties * | select name,userPrincipalName,displayname,title,department |export-csv  report_ou_$date.csv -NoTypeInformation -Encoding UTF8 -Force

##list ou
Get-ADOrganizationalUnit -Filter *  | select name, DistinguishedName |export-csv List_OU_$date.csv -NoTypeInformation -Encoding UTF8

##Get Logon 

Get-LogonStatistics -server FXJTSEMB003  | ft username, logontime, lastaccesstime, servername, clientversion -autosize

##Add-Permission

Add-MailboxPermission -Identity 'CN=Brian Sun 孙州川,OU=财富管理集团,OU=复星集团总部,DC=fosun,DC=com' -User 'FOSUN\Administrator' -AccessRights 'FullAccess'

##Search Message

$Target=“noreply@fosun.com”
$targetmaillbox="linyiyao@fosun.com"
$recipient="linyiyao@fosun.com"
$TargetFolder="SearchResults"
$Keyword="456"
Search-Mailbox -Identity $Target -SearchQuery "$Keyword AND $recipient" -EstimateResultOnly -TargetMailbox $targetmaillbox -TargetFolder $TargetFolder -LogLevel Full -LogOnly 
Get-Mailbox -ResultSize unlimited | Search-Mailbox -SearchQuery '关于施瑜的任命通知' -TargetMailbox linyiyao@fosun.com -TargetFolder search$date -LogLevel Full -DeleteContent

## message trans limited

get-transportconfig | ft maxsendsize, maxreceivesize 
get-receiveconnector | ft name, maxmessagesize 
get-sendconnector | ft name, maxmessagesize 
get-mailbox Administrator |ft Name, Maxsendsize, maxreceivesize

##Add Send-as

Get-Mailbox  -resultsize unlimited | Add-Adpermission -User 'FOSUN\ecard2018' -ExtendedRights 'Send-as' 

##Get -7 days Created users

Get-Mailbox -ResultSize Unlimited | Where-Object {$_.WhenCreated –ge ((Get-Date).Adddays(-7))} | ft name,PrimarySmtpAddress,servername,database,WhenCreatedUTC -auto

##Msg log Track

$TransServer=Get-TransportServer

ForEach( $srv in $TransServer) 
{ 
Get-MessageTrackingLog -Server $srv.Name -ResultSize Unlimited -eventid DELIVER -start 2018/03/25 -End 2018/03/26 -sender noreply@fosun.com -Recipients linyiyao@fosun.com `
| select timestamp,sender,{$_.Recipients},messagesubject

## change mailbox size
Set-Mailbox -Identity chenwch@fosun.com -IssueWarningQuota 4Gb -ProhibitSendQuota 4.8Gb -ProhibitSendReceiveQuota 5.1Gb -UseDatabaseQuotaDefaults $false

## 
function Get-DistributionGroupMemberRecursive ($GroupIdentity) {
	$member_list = Get-DistributionGroupMember -Identity $GroupIdentity
	foreach ($member in $member_list) {
		if ($member.RecipientType -like '*Group*') {
			Get-DistributionGroupMemberRecursive -GroupIdentity $member.Identity
		} else {
			$member
		}
	}
}

$group = Get-DistributionGroup -Identity zhangxingbao
Get-DistributionGroupMemberRecursive -GroupIdentity $group.Identity | select name,PrimarySmtpAddress,DistinguishedName,WhenCreated |export-csv Groupmember_$date.csv -NoTypeInformation -Encoding UTF8

function 邮箱大小报表
{

$Results = @()
$MailboxUsers = get-mailbox -resultsize unlimited

foreach($user in $mailboxusers)
{
$UPN = $user.userprincipalname
$MbxStats = Get-MailboxStatistics $UPN

$Properties = @{
Name = $user.DisplayName
Email = $user.PrimarySmtpAddress
Dept=$user.OrganizationalUnit
MailboxSize =[math]::Round($MbxStats.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","")/1Gb,4)
TotalItem=$MbxStats.ItemCount
States=$MbxStats.StorageLimitStatus
MailboxQuote=$user.ProhibitSendReceiveQuota
Lastlogon=$MbxStats.LastLogonTime
Lastlogoff=$MbxStats.LastLogoffTime
Server = $MbxStats.servername
Database= $MbxStats.databasename
CreatedDate=$MailboxUsers.whencreated
}

$Results += New-Object psobject -Property $properties

}

$Results | select Name,email,dept,mailboxsize,MailboxQuote,States,Lastlogon,Lastlogoff,server,database,CreatedDate |export-csv report_mailbox_$date.csv -NoTypeInformation -Encoding UTF8

}

##list mailboxdatabase size

Get-MailboxDatabase -Status | select Name,DatabaseSize,AvailableNewMailboxSpace


## remove groupmember
$gname="fosungroup.list"
$list = Get-DistributionGroupMember -Identity $gname -ResultSize unlimited
foreach ($user in $list)
{ Write-Host $user.primarySmtpAddress
Remove-DistributionGroupMember -Identity $gname -Member $user.PrimarySmtpAddress -Confirm:$false }

## list client api
Get-MailboxServer |Sort-Object name | foreach {Get-LogonStatistics -Server $_.name} | Export-Csv client$date.csv -NoTypeInformation -Encoding UTF8

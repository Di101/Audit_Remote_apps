Import-Module ActiveDirectory
Import-Module ImportExcel
<#
1.Скрипт атоматический добавляет всех новых пользователей АД в группу блокировки телеграма.Выборка пользователей осуществляется из ЦО и только включенные УЗ.
    и формирует отчет. 

2.Если пользоватлей есть в файле "Remote_Apps.xlsx" и он в группе блокировки телеграмма, скрипт должен удалять его из группы.
    и формироват отчет.
3.Если пользователя нет в файле то скрипт добавляет его в группу блокировки телеграмма.
    и формирует отчет

#>


$Users_allowed_remote_apps= Import-Excel -Path '\\hostname\Remote_Apps.xlsx'  | ForEach-Object {$_.Учётная_запись} 
#Проверяем членоство в группе
$users_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny = foreach ($User_allowed_remote_apps in $Users_allowed_remote_apps) 
{
try {
Get-ADUser -Identity $User_allowed_remote_apps -Properties memberof | Where-Object {$_.memberof -like "*App_Remote_access_Deny*"} | ForEach {$_.SamAccountName}
}
catch {
$Cannot_find_users_error += $User_allowed_remote_apps + "," + "`n"
}
}

#Удаляем из группы
foreach ($user_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny in $users_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny)
{Remove-ADGroupMember -Identity App_Remote_access_Deny -Member $user_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny -Confirm:$false
}

#Необходимо добавить больше инфы о пользователях. логин+должность+департамент+ и тд.
$Report_delate_users_in_group_more_info = foreach ($user_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny in $users_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny)
{
 Get-ADUser -Identity $user_allowed_remote_apps_but_users_add_to_group_App_Remote_access_Deny -Properties CN,co,City , Department,Title, physicalDeliveryOfficeName,Enabled,CanonicalName | select SamAccountName,CN,co,City , Department,Title, physicalDeliveryOfficeName,Enabled,CanonicalName
 }

#Выгружаем всех пользователей центрального офиса 


$all_users_add_to_group_App_Remote_access_Deny = Get-ADUser -Filter * -SearchBase "OU=Office,OU=Moscow,OU=RU,OU=you_company_name,DC=you,DC=domain" -Properties Department, enabled |`
 where {$_.enabled -eq $True } | foreach {$_.SamAccountName} 



#Сравниваем пользователей центрально офиса  с пользователями файле "Remote_Apps.xlsx" тем самым выявляем пользователей которых необходимо добавить в группу App_Remote_access_Deny. 
$users_add_to_group =  Compare-Object -ReferenceObject $($all_users_add_to_group_App_Remote_access_Deny) -DifferenceObject $($Users_allowed_remote_apps) |where {$_.SideIndicator -eq '<='} | foreach {$_.InputObject}


#Выявляем уже добавленных пользователей в группу App_Remote_access_Deny и добавляем оставшихся в App_Remote_access_Deny

$Report_add_users_in_group = foreach ($user_add_to_group in $users_add_to_group) 
{
Get-ADUser -Identity $user_add_to_group -Properties CN,co,City, Department,Title, physicalDeliveryOfficeName,Enabled,CanonicalName, memberof| Where-Object {!($_.memberof -like "*App_Remote_access_Deny*")} | select SamAccountName,CN,co,City , Department,Title, physicalDeliveryOfficeName,Enabled,CanonicalName
}
$Report_add_users_in_group_for_add_users = $Report_add_users_in_group | foreach  SamAccountName
$add_users_ADGroupMember = Add-ADGroupMember -Identity App_Remote_access_Deny -Members $Report_add_users_in_group_for_add_users

#Формируем HTML таблицу
$a = "<style>"
$a = $a + "BODY{background-color:peachpuff;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:PaleGoldenrod}"
$a = $a + "</style>"


#Формируем отчет о удаленных польхователях из группы App_Telegram_Test

if (([string]::IsNullOrEmpty($Report_delate_users_in_group_more_info)))
{
$Report_delate_users_in_group_more_info = "" | Out-File C:\you_dir\Report_delate_users_in_group_more_info.html 
}
else {
$Report_delate_users_in_group_more_info | ConvertTo-Html -Head $a | Out-File C:\you_dir\Report_delate_users_in_group_more_info.html 
}

#Формируем отчет о добавленных пользователях в группы App_Telegram_Test

if (([string]::IsNullOrEmpty($Report_add_users_in_group)))
{
$Report_add_users_in_group = "" | Out-File C:\you_dir\Report_add_users_in_group.html 
}
else {
$Report_add_users_in_group | ConvertTo-Html -Head $a | Out-File C:\you_dir\Report_add_users_in_group.html 
}


#Обработка контента в письме.

$get_delate_users_in_group = Get-Content C:\you_dir\Report_delate_users_in_group_more_info.html 
 if (([string]::IsNullOrEmpty($get_delate_users_in_group)))
 {$get_delate_users_in_group = "Пользователей которым необходимо предоставлить доступ к средствам удаленного доступа не обнаружено."
  
 }
 else {
 $get_delate_users_in_group_else= Get-Content C:\you_dir\Report_delate_users_in_group_more_info.html  -raw 
 $get_delate_users_in_group = 'Пользователи были удалены из группы App_Remote_access_Deny тем самым им предоставлен доступ к средставам удаленного доступа:
 ' +  $get_delate_users_in_group_else
 }


 $get_add_users_in_group = Get-Content C:\you_dir\Report_add_users_in_group.html 
 if (([string]::IsNullOrEmpty($get_add_users_in_group)))
 {$get_add_users_in_group = "Пользователей которым необходимо заблокировать доступ к средствам удаленного доступа не обнаружено."
  
 }
 else {
 $get_add_users_in_group_else= Get-Content C:\you_dir\Report_add_users_in_group.html   -raw 
 $get_add_users_in_group = 'Пользователи добавлены в группу блокировки средств удаленного доступа:
 ' +  $get_add_users_in_group_else
 }

   if (([string]::IsNullOrEmpty($Cannot_find_users_error )))
 {$Cannot_find_users_error  = ""
 }
 else {
 $Cannot_find_users_error_msg  = 'Пользователей не удалось найти в AD:
 ' +  $Cannot_find_users_error
 }

 #Отправляем письмо
$smtpServer = "mail1.you_domain.com"
$MailTo = 'user_name@you_domain.com'
$MailTo_2 ='amvolkov@you_domain.com' 
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$gCont_1 =  $get_delate_users_in_group
$gCont_2 =  $get_add_users_in_group
$gCont_3 = $Cannot_find_users_error_msg


$msg.From = "Remote_apps_Block@no-reply.com"
$msg.To.Add($MailTo)
$msg.To.Add($MailTo_2)
$msg.Subject = "Блокировка приложений удаленного доступа"
$msg.IsBodyHTML = $True
$msg.Body =  $gCont_1  , '<br>' ,'<br>' , $gCont_2 , '<br>' ,'<br>' , $gCont_3 

$smtp.Send($msg) 
#

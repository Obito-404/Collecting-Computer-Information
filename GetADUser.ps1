# 指定导出文件的路径
$ExportPath = "D:"

# 指定要导出的用户属性
$Properties = "SamAccountName", "Name", "EmailAddress", "Department","OfficePhone"

# 使用Get-ADUser命令获取所有用户信息，并使用Where-Object筛选Department不为空并且SamAccountName开头字母是s和h的用户，最后将结果导出到CSV文件
Get-ADUser  -filter * -Properties $Properties | Where-Object { $_.Department -ne "" -and $_.SamAccountName -like "[sh]*" } | Select-Object $Properties | Export-CSV -Path "$ExportPath\ADUsers.csv" -NoTypeInformation

#Get-ADUser -Filter * -Properties * | Export-CSV -Path "$ExportPath\ADUsers_$Date.csv" -NoTypeInformation

# 显示导出完成的消息
Write-Host "AD用户信息已导出到$ExportPath\ADUsers.csv"

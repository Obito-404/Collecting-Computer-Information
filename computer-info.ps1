$computer = Get-WMIObject Win32_ComputerSystem | Select-Object Name   
$user = $(Get-WMIObject -class Win32_ComputerSystem | select username)   
$os = Get-WmiObject -class Win32_OperatingSystem | Select-Object Caption,Version   
$ip = Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $(Get-NetConnectionProfile | Select-Object -ExpandProperty InterfaceIndex) | Select-Object IPAddress   

$dns = (Get-DnsClientServerAddress -AddressFamily IPv4 | where {$_.ServerAddresses -ne $null}) | Select-Object ServerAddresses   
$mac = (Get-WmiObject Win32_NetworkAdapterConfiguration | where {$_.ipenabled -EQ $true}) | Select-Object Description,Macaddress   
$sn = Get-WmiObject -ComputerName localhost -Class Win32_BIOS | Select-Object SerialNumber   
$admin = Get-LocalGroupMember -Group "Administrators" | Select-Object Name   

$cpu = Get-WmiObject -Class Win32_Processor | Select-Object Name,Caption,SocketDesignation   
$memory = Get-WmiObject -Class Win32_PhysicalMemory | Select-Object @{Label="MemorySize";Expression={$_.Capacity / 1gb -as [int] }}
$disk = Get-WmiObject -Class Win32_DiskDrive | Select-Object Model,@{Label="DiskSize";Expression={$_.Size / 1gb -as [int] }}   
$videocontroller = Get-WmiObject -Class Win32_VideoController | Select-Object Caption   

#软件信息
$software = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {($_.DisplayName -ne $null)}    
$ms_software = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*  | Where-Object {($_.DisplayVersion -ne $null)}    


Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | Select-Object DisplayName| Where-Object {($_.DisplayName -Like "*Visual Studio 2022")}

$all_software = @($software + $ms_software)

#筛选软件名单
$softwareList = @(
    "*Soliwork*",
    "*NX*",
    "*Bartender*"
    "*ZWCAD*",
    "*PADS*",
    "*altium designer*",
    "*Suse*",
    "*buymanager*",
    "*SUPPLYON*",
    "*Veeam*",
    "*Keil*"
)

# 创建一个空的数组来保存查询结果
$results = @()

# 遍历变量并查询值
foreach ($software in $softwareList) {
    $x86Result = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Where-Object {($_.DisplayName -Like $software)}
    $x64Result = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Where-Object {($_.DisplayName -Like $software)}

    # 将查询结果添加到 $results 数组中
    if ($x86Result) { $results += $x86Result }
    if ($x64Result) { $results += $x64Result }
}

function main {
    [PSCustomObject]@{
        Computer = $computer.Name
        User = $user.Username
        OS = $os.Caption
        OSVersion = $os.Version
        IPAddress = $ip.IPAddress
        DNS = $dns.ServerAddresses
        MACAddress = $mac.MacAddress
        SerialNumber = $sn.SerialNumber
        Admin = $admin.Name
        CPU = $cpu.Name
        CPUCaption = $cpu.Caption
        CPUSocket = $cpu.SocketDesignation
        MemorySize = $memory.MemorySize
        Disk = $disk.Model
        DiskSize = $disk.DiskSize
        VideoController = $videocontroller.Caption
        Checksoftware=$results. DisplayName
        Software = $all_software.DisplayName
    } | ConvertTo-Json
}

main | Out-File -Encoding "UTF8" -FilePath #保存文件的路径 \\D:


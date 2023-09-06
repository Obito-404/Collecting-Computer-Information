# 🖥️  收集计算机信息

一套能从 AD 域中的计算机自动获取硬件和软件信息的工具。它使用 PowerShell 和 Python 构建。

- AD Group Policy Management 推送脚本，将每次登录时自动写入到指定文件夹路径
- 导出AD域中所有用户信息(这一步非必须!!!仅用来通过匹配用户，将hostanme与user匹配起来，可以追加一些AD中导出的字段，如Display name,E-mail,Telephone number...)
- python将收集的电脑信息和用户合并后添加或更新到数据库


## 📋 步骤

使用该工具的步骤如下：

1. 创建所有用户都可以访问的共享文件，并添加Everyone的读写权限。
2. 将`Computer Information.ps1`添加到Group Policy Management的OU。每当用户登录时，他们的计算机都会自动将其硬件和软件信息以 JSON 格式写入刚刚创建的指定路径。
<br>

```shell
 Out-File -Encoding "UTF8" -FilePath \\保存文件的路径\$env:COMPUTERNAME.json
```
**⚠️ 说明: 根据自身所需要收集的硬件或软件信息更改脚本里的变量。**



3. 定时运行 `Get ADuser.ps1`，导出域中所有用户到指定路径。
4. 创建数据库，将对应的字段匹配如下。定时运行 `Computer-User.py`，将更新从 `Get ADuser` 中获取的每台计算机的信息。
```shell
 query = "UPDATE ComputerInfo SET [User] = ?, OS = ?, OSVersion = ?, IPAddress = ?, DNS = ?, MACAddress = ?, SerialNumber = ?, Admin = ?, CPU = ?, CPUCaption = ?, CPUSocket = ?, MemorySize = ?, Disk = ?, DiskSize = ?, VideoController = ?, Checksoftware = ?, Software = ? , Updatetime =? WHERE Computer = ?"
```
**⚠️ 说明: 将所有任务设置成定时任务，即可每天自动更新电脑信息与用户信息到数据库。**

## 🛠️ Requirements

- Windows operating system
- PowerShell
- Python
- Pyodbc (Python library for connecting to SQL database)

## 📝 License

本项目采用 MIT 许可，详情请参见 [LICENSE](LICENSE) 文件。

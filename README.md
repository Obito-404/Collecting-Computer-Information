# 🖥️  Collecting computer information

<p align="center">
    【English | <a href="README/README.zh_CN.md">Chinese</a>】
</p>

A toolset designed to automatically gather hardware and software information from computers in an Active Directory (AD) domain. It is built using PowerShell and Python.

- AD Group Policy Management deploys scripts that automatically write data into a specified folder path each time a user logs in.
- Exporting all user information from the AD domain (this step is optional! It's only used for matching users to hostnames and can include additional AD-exported fields such as Display Name, Email, Telephone Number).
- Python combines the collected computer information and user data, adding or updating it in the database.


## 📋 Steps

Here are the steps for using this tool:

1. Create a shared folder accessible to all users and grant Everyone read and write permissions.
2. Add Computer Information.ps1 to the Group Policy Management Organizational Unit (OU). Whenever a user logs in, their computer will automatically write its hardware and software information in JSON format to the specified path that was created in step 1.
<br>

```shell
 Out-File -Encoding "UTF8" -FilePath \\保存文件的路径\$env:COMPUTERNAME.json
```
**⚠️ Note: Modify the variables in the script to collect hardware or software information based on your specific requirements.**



3. Schedule the execution of Get ADuser.ps1 to export all domain users to a specified path.
4. Create a database and match the corresponding fields as follows. Schedule the execution of 'Computer-User.py' to update the information for each computer obtained from 'Get ADuser'.
```shell
 query = "UPDATE ComputerInfo SET [User] = ?, OS = ?, OSVersion = ?, IPAddress = ?, DNS = ?, MACAddress = ?, SerialNumber = ?, Admin = ?, CPU = ?, CPUCaption = ?, CPUSocket = ?, MemorySize = ?, Disk = ?, DiskSize = ?, VideoController = ?, Checksoftware = ?, Software = ? , Updatetime =? WHERE Computer = ?"
```
**⚠️ Note:  Set up all tasks as scheduled tasks to automatically update computer information and user information to the database every day."**


## 🛠️ Requirements

- Windows operating system
- PowerShell
- Python
- Pyodbc (Python library for connecting to SQL database)

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

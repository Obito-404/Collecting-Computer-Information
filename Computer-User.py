import os
import json
import pyodbc
import csv
from datetime import datetime

# database connection information
server = 'xxx'
database = 'xxx'
username = 'xxx'
password = 'xxx'

# connect to database with UTF-8 character set
conn = pyodbc.connect(
    f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password};CHARSET=UTF8")

cursor = conn.cursor()

# set path to directory containing JSON files
#电脑信息收集共享文件夹路径
path = r'\\xxx'
#ad导出用户的文件夹路径
csv_file = r'D:\ADUsers.csv'

def add_computer():
    # iterate over files in directory
    for file in os.listdir(path):
        file_name_path = os.path.join(path, file)
        with open(file_name_path, 'r', encoding='utf-8') as f:
            file_contents = f.read()
            if file_contents.startswith(u'\ufeff'):
                file_contents = file_contents.encode('utf8')[3:].decode('utf8')
            if not file_contents:
                continue
            data = json.loads(file_contents)
            computer_name = str(data['Computer'])
            user_name = str(data['User'])
            if '\\' in user_name:
                user_name = user_name.split('\\')[1]
            os_name = data['OS'][:30]
            os_version = str(data['OSVersion'])[:30]
            ip_address = str(data['IPAddress'])[:15]
            dns_servers = str(data['DNS'])[:100]
            mac_addresses = str(data['MACAddress'])[:50]
            serial_number = str(data['SerialNumber'])[:30]
            admin = str(data['Admin'])
            if admin == "None":
                admin = ""
            else:
                admin_accounts = str(data['Admin'])
                admin_counts = admin.count(',') + 1
                i = 0
                admin_list = []
                for i in range(admin_counts):
                    admin_account = str(data['Admin'][i])
                    admin_account = admin_account.split('\\')[1]
                    admin_list += [admin_account]
                    admin_accounts = ','
                    admin = admin_accounts.join(admin_list)
            cpu_model = str(data['CPU'])[:50]
            cpu_caption = str(data['CPUCaption'])[:100]
            cpu_socket = str(data['CPUSocket'])[:30]
            memory_size = str(data['MemorySize'])[:30]
            disk_models = str(data['Disk'])[:100]
            disk_size = str(data['DiskSize'])[:30]
            video_controller = str(data['VideoController'])[:100]
            check_software = str(data['Checksoftware'])
            Software = str(data['Software'])
            Updatetime= datetime.now()
    #-----------------------------------
            # check if computer_name already exists in the database
            query = "SELECT COUNT(*) FROM ComputerInfo WHERE Computer = ?"
            cursor.execute(query, [computer_name])
            result = cursor.fetchone()

            if result[0] > 0:  # update existing record
                query = "UPDATE ComputerInfo SET [User] = ?, OS = ?, OSVersion = ?, IPAddress = ?, DNS = ?, MACAddress = ?, SerialNumber = ?, Admin = ?, CPU = ?, CPUCaption = ?, CPUSocket = ?, MemorySize = ?, Disk = ?, DiskSize = ?, VideoController = ?, Checksoftware = ?, Software = ? , Updatetime =? WHERE Computer = ?"
                values = (user_name, os_name, os_version, ip_address, dns_servers, mac_addresses, serial_number, admin, cpu_model, cpu_caption, cpu_socket, memory_size, disk_models, disk_size, video_controller, check_software,Software,Updatetime,computer_name)
                cursor.execute(query, values)
                print(computer_name+"已更新")
            else:  # insert new record
                query = "INSERT INTO ComputerInfo (Computer, [User], OS, OSVersion, IPAddress, DNS, MACAddress, SerialNumber, Admin, CPU, CPUCaption, CPUSocket, MemorySize, Disk, DiskSize, VideoController,Checksoftware,Software,Updatetime) VALUES (?, ?, ?, ?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?,?)"
                values = (computer_name, user_name, os_name, os_version, ip_address, dns_servers, mac_addresses, serial_number,
                                  admin, cpu_model, cpu_caption, cpu_socket, memory_size, disk_models, disk_size,
                                  video_controller,check_software,Software,Updatetime)
                cursor.execute(query, values)
                print(computer_name+"已添加")

def add_user():
    # 读取CSV文件并解析为Python字典
    with open(csv_file, "r", newline="") as f:
        reader = csv.DictReader(f)
        query_conditions = [row["SamAccountName"] for row in reader]

    # 根据CSV文件中的条件查询SQL Server数据库记录，并插入对应的 Name，EmailAddress 和 Department
    cursor = conn.cursor()
    for qc in query_conditions:
        cursor.execute("SELECT * FROM ComputerInfo WHERE [User] = ?", qc)
        row = cursor.fetchone()
        if row:
            # 获取CSV文件中对应行的Name，EmailAddress和Department
            with open(csv_file, "r", newline="") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    if row["SamAccountName"] == qc:
                        name = row["Name"]
                        email = row["EmailAddress"]
                        department = row["Department"]
                        OfficePhone = row["OfficePhone"]
                        # 插入 Name，EmailAddress 和 Department 到 SQL Server 表
                        cursor.execute(
                            "UPDATE ComputerInfo SET Name = ?, EmailAddress = ?, Department = ? , OfficePhone = ? WHERE [User] = ?",
                            name, email, department, OfficePhone, qc)
                        conn.commit()
                        break

def main():
    add_computer()
    add_user()


if __name__ == '__main__':
    main()






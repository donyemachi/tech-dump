from getpass import getpass
from pprint import pprint
#import datetime
import json
from napalm import get_network_driver
from netmiko import juniper
from netmiko import ConnectHandler,ssh_exception
from getpass import getpass
import openpyxl
from datetime import datetime
from time import gmtime, strftime
#import os
#import pprint
import argparse
import copy
from datetime import datetime

import pandas as pd
from netmiko import Netmiko
from getpass import getpass

vlan_id = [2,200,202,10,14,1901,1902,103,21,22,23,25,26,99,5,101,6,104,105,106,107,108,109,110,111,112,116,118,128,129,130,152,160,176,192,340,350,360,370,380,390,401,402,403,410,411,412,413,420,421,404,405,432,440,450,460,470,480,490,540]
vlan_name = ['ADLS','Security','QF-BMS','PABX/VoIP','NY Printers','RP4VM - Data (Dell - Internal VLANs)','RP4VM - Replication (Dell - Internal VLANs)','Load Balancer','Scientific Computing - Heart Beat','Scientific Computing - Public Mgmt','Scientific Computing - VPN','Scientific Computing - TSM Backup','Research - Backbone','Management WiFi','Management LAN','Cluster Management - Heatbeat','Management2','vSphere vMotion','vSan','Research Lab Patched Network','Racknet - SSID','Research-Lab-Network','Research Monitoring System','Printers','Open WiFi','Employees SSID','exam SSID','edls SSID','eduroam SSID','IoT','Extron VLAN for AV Team','Guest SSID','BYOD SSID - Students','BYOD SSID - Employees','Dual-Stack - IPv6','Private-Admin-1','Private-Admin-2','Private-Admin-4','Private-Student-1','Private-Student-2','Private-Admin-3','Trust (Inside)','DMZ - 1','DMZ - 2','Backbone 1','Backbone 2','Backbone 3','Backbone 4','Backbone 5',	'Backbone 6','DMZ - SC','DMZ-EXT','VPN','Admin 1','Admin 2','Admin 4','Student 1','Student 2','Admin 3','Audio Visual-1']
parser = argparse.ArgumentParser(prog= 'switch_inventory.py', usage='%(prog)s device username',description="login credentials")

vlan_detail = {2: 'ADLS', 200: 'Security', 202: 'QF-BMS', 10: 'PABX/VoIP', 14: 'NY Printers', 1901: 'RP4VM - Data (Dell - Internal VLANs)', 1902: 'RP4VM - Replication (Dell - Internal VLANs)', 103: 'Load Balancer', 21: 'Scientific Computing - Heart Beat', 22: 'Scientific Computing - Public Mgmt', 23: 'Scientific Computing - VPN', 25: 'Scientific Computing - TSM Backup', 26: 'Research - Backbone', 99: 'Management WiFi', 5: 'Management LAN', 101: 'Cluster Management - Heatbeat', 6: 'Management2', 104: 'vSphere vMotion', 105: 'vSan', 106: 'Research Lab Patched Network', 107: 'Racknet - SSID', 108: 'Research-Lab-Network', 109: 'Research Monitoring System', 110: 'Printers', 111: 'Open WiFi', 112: 'Employees SSID', 116: 'exam SSID', 118: 'edls SSID', 128: 'eduroam SSID', 129: 'IoT', 130: 'Extron VLAN for AV Team', 152: 'Guest SSID', 160: 'BYOD SSID - Students', 176: 'BYOD SSID - Employees', 192: 'Dual-Stack - IPv6', 340: 'Private-Admin-1', 350: 'Private-Admin-2', 360: 'Private-Admin-4', 370: 'Private-Student-1', 380: 'Private-Student-2', 390: 'Private-Admin-3', 401: 'Trust (Inside)', 402: 'DMZ - 1', 403: 'DMZ - 2', 410: 'Backbone 1', 411: 'Backbone 2', 412: 'Backbone 3', 413: 'Backbone 4', 420: 'Backbone 5', 421: 'Backbone 6', 404: 'DMZ - SC', 405: 'DMZ-EXT', 432: 'VPN', 440: 'Admin 1', 450: 'Admin 2', 460: 'Admin 4', 470: 'Student 1', 480: 'Student 2', 490: 'Admin 3', 540: 'Audio Visual-1', 888:' NYTVLAN'}
interface_status = {'True':'Interface_Up', 'False' : 'Interface_Down'}

parser.add_argument('--device', help="hostname or IP",type=str,default=None  )
parser.add_argument('--username', help="login name", type=str, default=None)
#parser.add_argument('password', help="password",type=str, default=None )

args = parser.parse_args()

host = args.device
user = args.username


junos_driver = get_network_driver("junos")

junos_password = getpass(" SSH key password: ")

with junos_driver(hostname=host, username=user, password=junos_password) as dev:
    dev.open()
    info = dev.get_interfaces()
    dev_info = {'interface': [entry1 for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'mac-address': [entry2['mac_address'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'Flap' : [entry2['last_flapped'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'DESC' : [entry2['description'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'status' : [entry2['is_up'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'Speed' : [entry2['speed'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            'MTU' : [entry2['mtu'] for entry1, entry2 in info.items()if entry1 != 'ae0.0'],
            }
            
    mac_table = dev.get_mac_address_table()
    mac_data = {'mac':  [entry['mac'] for entry in mac_table if entry['interface'] != 'ae0.0'],
            'interface': [entry['interface'] for entry in mac_table if entry['interface'] != 'ae0.0'],
            'vlan': [entry['vlan'] for entry in mac_table if entry['interface'] != 'ae0.0']
#            'Last_flap': [entry['last_move'] for entry in mac_table if entry['interface'] != 'ae0.0']
            }

#df['eq'] = df.apply(lambda row: row['coname1'] == row['coname2'], axis=1).astype(int)
df6 = pd.DataFrame(mac_data, columns=list(mac_data.keys()))
df = pd.DataFrame(dev_info, columns=list(dev_info.keys()))
df11 = copy.deepcopy(df)
#df11['Intf_status'] = df11.apply(lambda row: row['Intf_status'] == row['status'], axis=1).astype(int)

#df['status'].astype(str)
df['status'].astype(str).replace(to_replace=interface_status, inplace=True)

df11['port_status'] = df['status'].astype(str)
df11['port_status'].replace(to_replace=interface_status, inplace=True)
df12 = df11[df11.columns[~df11.columns.isin(['status', 'mac-address'])]]
df9 = df.set_index('status')
#df7 = df.groupby('status')[['interface', 'Flap']].rename({'Flap': 'Last_interface_Flap'},axis = 'columns')
df6['vlan name'] = df6['vlan']

df6['vlan name'].replace(to_replace=vlan_detail, inplace=True)
df2 = copy.deepcopy(df6)
#df3 = df2.groupby('vlan name')[['mac','interface']].count()
df4 = df2.groupby('vlan name')[['mac']].count().rename({'mac': 'users'},axis = 'columns')
#df13 = df.groupby('vlan name')[['mac','interface']].count()

current_datetime = datetime.now()
date_time = current_datetime.strftime("%m_%d_%Y - %H_%M_%S")
str_time = str(date_time)
filename = 'report' +' ' + str_time + '.xlsx'
file = str(filename)

with pd.ExcelWriter(file,  engine='xlsxwriter') as writer3:


    df.to_excel(writer3, sheet_name="Raw_Interf_Data", index=False)
    df6.to_excel(writer3, sheet_name="user_Details", index=False)
    df11.to_excel(writer3, sheet_name="port status", index=True)
    df12.to_excel(writer3, sheet_name="users_count", index=True)
    df9.to_excel(writer3, sheet_name="status_count3", index=True)


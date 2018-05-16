# int2excel

## What does this do:
int2excel uses Nekmiko (https://github.com/ktbyers/netmiko) which in turn utilises Paramiko SSH connections
to connect to Juniper devices, issue a command, analyse output and populate an excel file.<br>
This script will create an excel workbook `int2excel.xlsx` with details of each host defined in `device-list.txt`.
Each device will have it's own named sheet and will contain information on interface name, physical and logical status. 

## In short what does this do:
Collect device interface status and write to an excel file, one sheet per device.

## What use is this to me:
- Collate network wide interface details.
- Quick network-wide audit of all interfaces.

## How do I use int2excel:
- Enter your device hostnames or IP addresses into `device-list.txt`
- Run the script: `python int2excel.py`
- Enter your common username and password.
- Observe output in the excel file `int2excel.xlsx`

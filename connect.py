import uiautomator2 as u
import os
from devices import getDevicesAll


# 无线连接手机
def wifi_connect_phone():
    devices_list = getDevicesAll()
    adress_list = ['192.168.31.xxx:xxxx']
    for ip in adress_list:
        if ip not in devices_list:
            os.system('adb -s {} tcpip 5566'.format(ip))
            d = u.connect(ip)
            print(d.info)
            print('连接成功！')
        else:
            print('手机已连接，无需重新连接！')


wifi_connect_phone()

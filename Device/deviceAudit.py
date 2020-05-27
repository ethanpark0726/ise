import sys
import os
import subprocess
import openpyxl
import datetime

from ise import ERS  # noqa E402
from pprint import pprint  # noqa E402
from config import uri, endpoint, endpoint_group, user, identity_group, device, device_group, trustsec  # noqa E402

def getList(iseObj):

    pageNumber = 1
    deviceList = list()

    temp = iseObj.get_devices(page=pageNumber)
    deviceList.append(temp.get('response'))

    while temp.get('nextPage'):
        pageNumber += 1
        temp = iseObj.get_devices(page=pageNumber)
        deviceList.append(temp.get('response'))
    
    print("--- List gathering Complete!")
    return deviceList

def getDeviceIDList(deviceList):

    deviceIDList = list()
    for i in range(len(deviceList)):
        for j in range(len(deviceList[i])):
            deviceIDList.append(deviceList[i][j][0])
    print("--- Device ID gathering Complete!")
    return deviceIDList

def getIPList(iseObj, deviceIDList):
    
    ipList = list()

    for deviceID in deviceIDList:
        ipList.append(iseObj.get_device(deviceID).get('response').get('NetworkDeviceIPList')[0].get('ipaddress'))
    print("--- Device IP gathering Complete!")
    return ipList

def getPingResult(deviceIPList):

    pingList = dict()

    for deviceIP in deviceIPList:
        
        result = os.system('ping -n 1 -w 2 ' + deviceIP)

        if result == 0:
            pingList[deviceIP] = 'Okay'
        else:
            pingList[deviceIP] = 'Needs to check'
    print("--- Device Ping list gathering Complete!")
    return pingList

def createExcelFile():
    # Excel File Creation
    nowDate = 'Report Date: ' + str(datetime.datetime.now().strftime('%Y-%m-%d'))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'ISE Device Audit'
    ws['A2'] = nowDate
    ws['A4'] = 'Hostname'
    ws['B4'] = 'IP Address'
    ws['C4'] = 'Ping Status'

    wb.save('ISE_Device_Audit.xlsx')
    wb.close()

def saveExcelFile(deviceIDList, deviceIPList, pingResult):

    wb = openpyxl.load_workbook('ISE_Device_Audit.xlsx')
    ws = wb.active
    cellNumber = 5

    for i in range(len(deviceIDList)):
        ws['A' + str(cellNumber)] = deviceIDList[i]
        ws['B' + str(cellNumber)] = deviceIPList[i]
        ws['C' + str(cellNumber)] = pingResult.get(deviceIPList[i])
        cellNumber += 1

    wb.save('ISE_Device_Audit.xlsx')
    wb.close()

if __name__ == "__main__":

    ise = ERS(ise_node=uri['ise_node'], ers_user=uri['ers_user'], ers_pass=uri['ers_pass'], verify=False,
          disable_warnings=True, timeout=15)
    
    createExcelFile()

    deviceList = getList(ise)
    
    deviceIDList = getDeviceIDList(deviceList)
    
    deviceIPList = getIPList(ise, deviceIDList)
    pingResult = getPingResult(deviceIPList)

    saveExcelFile(deviceIDList, deviceIPList, pingResult)
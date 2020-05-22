import sys
import os
import subprocess
sys.path.append('./')

from ise import ERS  # noqa E402
from pprint import pprint  # noqa E402
from config import uri, endpoint, endpoint_group, user, identity_group, device, device_group, trustsec  # noqa E402

def getList(iseObj):

    pageNumber = 1
    deviceList = list()

    temp = iseObj.get_devices(page=pageNumber)

    while temp.get('nextPage'):
        deviceList.append(temp.get('response'))
        pageNumber += 1
        temp = iseObj.get_devices(page=pageNumber)
    
    return deviceList

def getDeviceIDList(deviceList):

    ipList = list()
    for i in range(len(deviceList)):
        for j in range(20):
            ipList.append(deviceList[i][j][0])
    return ipList

if __name__ == "__main__":

    ise = ERS(ise_node=uri['ise_node'], ers_user=uri['ers_user'], ers_pass=uri['ers_pass'], verify=False,
          disable_warnings=True, timeout=15)
    
    #deviceList = getList(ise)
    #deviceIDList = getDeviceIDList(deviceList)
    #print(deviceIDList)
    rep = os.system('ping -n 1 ' + '10.4.22.199')

    print(rep)
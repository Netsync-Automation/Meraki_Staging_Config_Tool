import requests
import json
import pandas as pd
import numpy as np
import os
import time
# import sys
import logging

time_str = time.strftime('%m%d%Y_%H%M%S')
print(time_str)
logging.basicConfig(filename=f"LOG_file_{time_str}.log",
					format='%(asctime)s %(message)s',
					filemode='w')
logger = logging.getLogger()
# logger.setLevel(logging.INFO)
# Uncomment line below if need debugging in Log_file
logger.setLevel(logging.DEBUG)

# old_stdout = sys.stdout
# log_file = open("message.log","w")
# sys.stdout = log_file

df = pd.read_excel("Config_File.xlsx", 'Step 0', skiprows=0)
df2 = df.replace(np.nan, '', regex=True)
data = []

for index, rows in df2.iterrows():
    data.append({rows[0]: rows[1]})

API_VALUE = data[0]['API Key']
ORG_ID = data[1]['Organization ID']

# Helping def. Used for various steps
def getAllNetworks():
    url = f"https://api.meraki.com/api/v1/organizations/{ORG_ID}/networks"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Cisco-Meraki-API-Key": API_VALUE
    }
    payload = None
    response = requests.request('GET', url, headers=headers, data=payload)
    parsed = json.loads(response.text)
    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
    if response.status_code != 200:
        print(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info("API RESPONSE\n")
        logger.info(apiPrintLogger)
        return []
    logger.info("API RESPONSE\n")
    logger.info(apiPrintLogger)
    # print(json.dumps(parsed, indent=4, sort_keys=True))
    return response.json()

# Helping def. Used for various steps. Return a single device.
def getDevice(DeviceSerial):
    url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Cisco-Meraki-API-Key": API_VALUE
    }
    payload = None
    response = requests.request('GET', url, headers=headers, data=payload)
    parsed = json.loads(response.text)
    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
    if response.status_code != 200:
        print(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info("API RESPONSE\n")
        logger.info(apiPrintLogger)
        return []
    logger.info("API RESPONSE\n")
    logger.info(apiPrintLogger)
    return response.json()

# Helping def. Used for various steps. Return a single port.
def getPort(DeviceSerial, PortUnique):
    url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}/switch/ports/{PortUnique}"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Cisco-Meraki-API-Key": API_VALUE
    }
    payload = None
    response = requests.request('GET', url, headers=headers, data=payload)
    parsed = json.loads(response.text)
    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
    if response.status_code != 200:
        print(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info(f'Error with status code {response.status_code}, text: {response.text}')
        return []
        logger.info("API RESPONSE\n")
        logger.info(apiPrintLogger)
    logger.info("API RESPONSE\n")
    logger.info(apiPrintLogger)
    return response.json()

# Helping def. Used for step7. Update a single Access Point port, particular variables.
def updateAPport(DeviceSerial, portToUpdate, changeName, changeTag, changeType, changeNativeVlan, changeAllowedVlans):
    payload_dict = {"name": changeName, "type": changeType, "allowedVlans": changeAllowedVlans, "vlan": changeNativeVlan,
              "tags": [changeTag]}
    if changeName is None:
        del payload_dict["name"]
    if changeTag is None:
        del payload_dict["tags"]
    if changeType is None:
        del payload_dict["type"]
    if changeNativeVlan is None:
        del payload_dict["vlan"]
    if changeAllowedVlans is None:
        del payload_dict["allowedVlans"]

    url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}/switch/ports/{portToUpdate}"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Cisco-Meraki-API-Key": API_VALUE
    }
    payload = json.dumps(payload_dict)
    response = requests.request('PUT', url, headers=headers, data=payload)
    parsed = json.loads(response.text)
    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
    logger.info("API RESPONSE\n")
    logger.info(apiPrintLogger)
    # print(f'    Printing API response####: \n')
    # print(json.dumps(parsed, indent=4, sort_keys=True))

# Helping def for Step8. Get all Client for Network ID. API responce specified to 1000 lines per page, it's max for Meraki API v1. If you have more Clients - this script needs to be enhansed.
def getCleints(networkID, targetOUI):
    url = str(f'https://api.meraki.com/api/v1/networks/{networkID}/clients?perPage=999&s&mac={targetOUI}')
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-Cisco-Meraki-API-Key": API_VALUE
    }
    payload = None
    response = requests.request('GET', url, headers=headers, data=payload)
    parsed = json.loads(response.text)
    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
    if response.status_code != 200:
        print(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info(f'Error with status code {response.status_code}, text: {response.text}')
        logger.info(apiPrintLogger)
        return []
    logger.info(apiPrintLogger)
    return response.json()
    # print(json.dumps(parsed, indent=4, sort_keys=True))

# Helping def for Step8. Updating switchport with particular OUI online and wired with target attributes.
# Trunk ports will be skipped by default to avoid accidental config change on inter-switch port.
def updateOUIport(switchName, switchSerial, switchPort, targetPortName, targetVlan, targetTag):
    portResponse = getPort(switchSerial, switchPort)
    portType = portResponse["type"]
    if portType != 'access':
        print(f'     Port type is not ACCESS. Skipping this port')
        logger.info(f'     Port type is not ACCESS. Skipping this port')
        return
    else:
        portName = portResponse["name"]
        portVlan = str(portResponse["vlan"])
        portTag = portResponse["tags"]
        if portName == targetPortName and portVlan == targetVlan and portTag == targetTag:
            print(f'    Switch: {switchSerial}, Port: {switchPort} already have correct configuration')
            logger.info(f'    Switch: {switchSerial}, Port: {switchPort} already have correct configuration')
            return
        else:
            changeName = None
            changeVlan = None
            changeTag = None
            if portName == targetPortName:
                print(f'    NAME. Name on port {switchPort} already {targetPortName}. No changes needed')
                logger.info(f'    NAME. Name on port {switchPort} already {targetPortName}. No changes needed')
            else:
                print(
                    f'    *NAME. Port {switchPort} does not have correct Name. Port name will be changed to {targetPortName}\n'
                    f'        ACTUAL NAME: {portName}\n'
                    f'        TARGET NAME: {targetPortName}')
                logger.info(
                    f'    *NAME. Port {switchPort} does not have correct Name. Port name will be changed to {targetPortName}\n'
                    f'        ACTUAL NAME: {portName}\n'
                    f'        TARGET NAME: {targetPortName}')
                changeName = targetPortName
            if portVlan == targetVlan:
                print(
                    f'    ACCESS vlan. Access vlan on port {switchPort} already {targetVlan}. No changes needed')
                logger.info(
                    f'    ACCESS vlan. Access vlan on port {switchPort} already {targetVlan}. No changes needed')
            else:
                print(
                    f'    *ACCESS vlan. Port {switchPort} does not have correct Access Vlan. Port Access vlan will be '
                    f'changed to {targetVlan}\n'
                    f'        ACTUAL VLAN: {portVlan}\n'
                    f'        TARGET VLAN: {targetVlan}')
                logger.info(
                    f'    *ACCESS vlan. Port {switchPort} does not have correct Access Vlan. Port Access vlan will be '
                    f'changed to {targetVlan}\n'
                    f'        ACTUAL VLAN: {portVlan}\n'
                    f'        TARGET VLAN: {targetVlan}')
                changeVlan = targetVlan
            if portTag == targetTag:
                print(f'    TAG. Tag on port {switchPort} already {targetTag}. No changes needed')
                logger.info(f'    TAG. Tag on port {switchPort} already {targetTag}. No changes needed')
            else:
                print(
                    f'    *TAG. Port {switchPort} does not have correct Tag. Port Tag will be changed to {targetTag}\n'
                    f'        ACTUAL TAG: {portTag}\n'
                    f'        TARGET TAG: {targetTag}')
                logger.info(
                    f'    *TAG. Port {switchPort} does not have correct Tag. Port Tag will be changed to {targetTag}\n'
                    f'        ACTUAL TAG: {portTag}\n'
                    f'        TARGET TAG: {targetTag}')
                # converting tag list to string for API
                str_targetTag = str(targetTag[0])
                changeTag = str_targetTag
            print(f'Making Following Changes on Port: {switchPort}, Switch: {switchName} ({switchSerial}): \n'
                  f'    NAME. {changeName}\n'
                  f'    ACCESS vlan. {changeVlan}\n'
                  f'    TAG. {changeTag}\n')
            logger.info(f'Making Following Changes on Port: {switchPort}, Switch: {switchName} ({switchSerial}): \n'
                  f'    NAME. {changeName}\n'
                  f'    ACCESS vlan. {changeVlan}\n'
                  f'    TAG. {changeTag}\n')
            payload_dict = {"name": changeName,
                            "vlan": changeVlan,
                            "tags": [changeTag]}
            if changeName is None:
                del payload_dict["name"]
            if changeVlan is None:
                del payload_dict["vlan"]
            if changeTag is None:
                del payload_dict["tags"]
            url = f"https://api.meraki.com/api/v1/devices/{switchSerial}/switch/ports/{switchPort}"
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-Cisco-Meraki-API-Key": API_VALUE
            }
            payload = json.dumps(payload_dict)
            response = requests.request('PUT', url, headers=headers, data=payload)
            parsed = json.loads(response.text)
            apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
            if "errors" in parsed:
                print(f'UNSUCCESSFUL. Printing API response: \n')
                logger.info(f'UNSUCCESSFUL. Printing API response: \n')
                print(json.dumps(parsed, indent=4, sort_keys=True))
                logger.info(apiPrintLogger)
                return
            else:
                print("PORT SUCCESSFULLY UPDATED\n")
                logger.info("PORT SUCCESSFULLY UPDATED\n")
                logger.info(apiPrintLogger)
                return

# Step 2. Assign Devices to Site
def addToSite():
    df = pd.read_excel("Config_File.xlsx", 'Step 2', skiprows=2)
    df2 = df.replace(np.nan, '', regex=True)
    data = []

    for index, rows in df2.iterrows():
        temp = [rows["Serial Number"],
                rows["Network ID"]]
        data.append(temp)
    for list in data:
        serials = str(list[0].strip())
        networkId = str(list[1].strip())
        url = f"https://api.meraki.com/api/v1/networks/{networkId}/devices/claim"
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "X-Cisco-Meraki-API-Key": API_VALUE
        }
        payload = json.dumps({"serials": [serials]})
        response = requests.request('POST', url, headers=headers, data=payload)
        #parsed = json.loads(response.text)
        #apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
        # print(json.dumps(parsed, indent=4, sort_keys=True))
        print(f"Assigning Device {serials} to: {networkId}")
        ##logger.info(f"Assigning Device {serials} to: {networkId}")
        #logger.info("API RESPONSE\n")
        #logger.info(apiPrintLogger)

# Step 3. See The List of Networks for ORG ID
def getNetworks():
    print(f"Network IDs for Organization {ORG_ID}: ")
    logger.info(f"Network IDs for Organization {ORG_ID}: ")
    networks = getAllNetworks()
    for net in networks:
        net_id = net['id']
        net_name = net['name']
        print(f'    Network ID: {net_id}, Name: {net_name}')
        logger.info(f'    Network ID: {net_id}, Name: {net_name}')

# Step 4. Update Network name
def updateNetworks():
    networks = getAllNetworks()
    for net in networks:
        net_id = net['id']
        net_name = net['name']
        df = pd.read_excel("Config_File.xlsx", 'Step 4', skiprows=2)
        df2 = df.replace(np.nan, '', regex=True)
        for index, rows in df2.iterrows():
            NetworkID = rows["Network ID"]
            CurrentName = rows["Current Name"]
            NewName = rows["New Name"]
            if net_name == NewName:
                print(
                    f'    The current Network Name: {net_name} already match the new Network Name: {NewName}.\n    No need to update.')
                logger.info(
                    f'    The current Network Name: {net_name} already match the new Network Name: {NewName}.\n    No need to update.')
                continue
            else:
                if net_id != NetworkID or net_name != CurrentName:
                    continue
                else:
                    print(
                        f'    Upgrading:\n    Network ID: {net_id}, Name: {net_name} \n    with new Name:\n    Network ID: {NetworkID}, Name: {NewName}')
                    logger.info(
                        f'    Upgrading:\n    Network ID: {net_id}, Name: {net_name} \n    with new Name:\n    Network ID: {NetworkID}, Name: {NewName}')
                    url = f"https://api.meraki.com/api/v1/networks/{NetworkID}"
                    headers = {
                        "Accept": "application/json",
                        "Content-Type": "application/json",
                        "X-Cisco-Meraki-API-Key": API_VALUE
                    }
                    payload = json.dumps({"name": NewName})
                    response = requests.request('PUT', url, headers=headers, data=payload)
                    parsed = json.loads(response.text)
                    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
                    print(f'    Printing API response: \n')
                    logger.info(f'    Printing API response: \n')
                    print(json.dumps(parsed, indent=4, sort_keys=True))
                    logger.info("API RESPONSE\n")
                    logger.info(apiPrintLogger)

# Step 5. Update Devices Atributes.
def updateDeviceAtributes():
    df = pd.read_excel("Config_File.xlsx", 'Step 5', skiprows=2)
    df2 = df.replace(np.nan, '', regex=True)
    for index, rows in df2.iterrows():
        NetworkID = rows["Network ID"]
        DeviceSerial = rows["Serial Number"]
        NewName = rows["New Name"]
        Address = rows["Address"]
        MGMTvlan = rows["MGMT vlan"]
        staticIP_YesNo = rows["Static IP (true or false)"]
        # Run def to get Network ID for Device Serial Number from dashboard
        DeviceInfo = getDevice(DeviceSerial)
        # print(DeviceInfo)
        logger.info(DeviceInfo)

        net_id = DeviceInfo['networkId']
        if NetworkID != net_id:
            print(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            logger.info(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            continue
        else:
            print(f'    Upgrading Device: {DeviceSerial}\n    New Name: {NewName}\n    Address: {Address}')
            logger.info(f'    Upgrading Device: {DeviceSerial}\n    New Name: {NewName}\n    Address: {Address}')
            url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}"
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-Cisco-Meraki-API-Key": API_VALUE
            }
            payload = json.dumps({"name": NewName, "address": Address})
            response = requests.request('PUT', url, headers=headers, data=payload)
            parsed = json.loads(response.text)
            apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
            logger.info("API RESPONSE\n")
            logger.info(apiPrintLogger)
            #print(json.dumps(parsed, indent=4, sort_keys=True))

            url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}/managementInterface"
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-Cisco-Meraki-API-Key": API_VALUE
            }
            payload = json.dumps({"wan1": {"usingStaticIp": staticIP_YesNo, "vlan": MGMTvlan}})
            response = requests.request('PUT', url, headers=headers, data=payload)
            parsed = json.loads(response.text)
            apiPrintLogger1 = (json.dumps(parsed, indent=4, sort_keys=True))
            print(f'    Updating Device Attributes and MGMT Interface: {DeviceSerial}')
            logger.info(f'    Updating Device Attributes: {DeviceSerial}')
            logger.info("API RESPONSE\n")
            logger.info(apiPrintLogger1)
            # print(json.dumps(parsed, indent=4, sort_keys=True))

# Step 6. Update switchports
def updateSwitchports():
    df = pd.read_excel("Config_File.xlsx", 'Step 6', skiprows=2)
    df2 = df.replace(np.nan, '', regex=True)
    for index, rows in df2.iterrows():
        NetworkID = rows["Network ID"]
        DeviceSerial = rows["Serial Number"]
        # Run def to get Network ID for Device Serial Number from dashboard
        DeviceInfo = getDevice(DeviceSerial)
        logger.info(DeviceInfo)
        #print(DeviceInfo)
        net_id = DeviceInfo['networkId']
        if NetworkID != net_id:
            print(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            logger.info(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            continue
        else:
            print(f'    Updating Ports on Device: {DeviceSerial}')
            logger.info(f'    Updating Ports on Device: {DeviceSerial}')
            for index, rows in df2.iterrows():
                Port = rows["Ports"]
                PortName = rows["Port Name"]
                State = rows["Enabled/Disabled (true/false)"]
                PortMode = rows["Access/Trunk"]
                AccessVlan = rows["Access Vlan (native if trunk)"]
                if PortMode == 'access':
                    VoiceVlan = int(rows["Voice Vlan"])
                    print(
                        f'Updating Port: {Port}\n    Port Name: {PortName}\n    Port Mode: {PortMode}\n    Access Vlan: {AccessVlan}\n    Voice Vlan: {VoiceVlan}\n    State: {State}')
                    logger.info(
                        f'Updating Port: {Port}\n    Port Name: {PortName}\n    Port Mode: {PortMode}\n    Access Vlan: {AccessVlan}\n    Voice Vlan: {VoiceVlan}\n    State: {State}')
                elif PortMode == 'trunk':
                    AllowedVlans = rows["Allowed vlans (1,3,5-10)"]
                    print(
                        f'Updating Port: {Port}\n    Port Name: {PortName}\n    Port Mode: {PortMode}\n    Allowed Vlans: {AllowedVlans}\n    Access Vlan: {AccessVlan}\n    State: {State}')
                    logger.info(
                        f'Updating Port: {Port}\n    Port Name: {PortName}\n    Port Mode: {PortMode}\n    Allowed Vlans: {AllowedVlans}\n    Access Vlan: {AccessVlan}\n    State: {State}')
                else:
                    print(f'    Wrong Port Mode. Check Spreadsheet')
                    logger.info(f'    Wrong Port Mode. Check Spreadsheet')
                url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}/switch/ports/{Port}"
                headers = {
                    "Accept": "application/json",
                    "Content-Type": "application/json",
                    "X-Cisco-Meraki-API-Key": API_VALUE
                }
                if PortMode == 'access':
                    payload = json.dumps(
                        {"name": PortName, "type": PortMode, "voiceVlan": VoiceVlan, "vlan": AccessVlan,
                         "enabled": State})
                    response = requests.request('PUT', url, headers=headers, data=payload)
                    parsed = json.loads(response.text)
                    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
                    logger.info("API RESPONSE\n")
                    logger.info(apiPrintLogger)
                    # print(json.dumps(parsed, indent=4, sort_keys=True))
                elif PortMode == 'trunk':
                    payload = json.dumps(
                        {"name": PortName, "type": PortMode, "allowedVlans": AllowedVlans, "vlan": AccessVlan,
                         "enabled": State})
                    response = requests.request('PUT', url, headers=headers, data=payload)
                    parsed = json.loads(response.text)
                    apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
                    logger.info("API RESPONSE\n")
                    logger.info(apiPrintLogger)
                    #print(json.dumps(parsed, indent=4, sort_keys=True))
                else:
                    print(f'    Wrong Port Mode. Check Spreadsheet')
                    logger.info(f'    Wrong Port Mode. Check Spreadsheet')

# Step 7. Update Port with Meraki Access Point (MR46) based on lldp info
def getCDPLLDPneighbours():
    # Reading Info from spreadsheet Step 7B.
    df = pd.read_excel("Config_File.xlsx", 'Step 7B', skiprows=0)
    df2 = df.replace(np.nan, '', regex=True)
    data = []
    for index, rows in df2.iterrows():
        data.append({rows[0]: rows[1]})
    targetLLDPname = str(data[0]['Search for lldp name (can be partial, but must be unique for target Device)'])
    targetPortName = str(data[1]['Target Port Name'])
    targetAllowedVlan = str(data[2]['Target Allowed Vlans List'])
    targetNativeVlan = str(data[3]['Target Native Vlan'])
    targetTag = ([data[4]['Target Tag']])

    # Actual script. Starts from reading info in Step 7A sheet.
    df = pd.read_excel("Config_File.xlsx", 'Step 7A', skiprows=2)
    df2 = df.replace(np.nan, '', regex=True)
    for index, rows in df2.iterrows():
        NetworkID = rows["Network ID"]
        DeviceSerial = rows["Serial Number"]
        # Run def to get Network ID for Device Serial Number from dashboard
        DeviceInfo = getDevice(DeviceSerial)
        #print(DeviceInfo)
        net_id = DeviceInfo['networkId']
        DeviceName = DeviceInfo["name"]
        print(f"Checking Device: {DeviceName}, {DeviceSerial}")
        logger.info(f"Checking Device: {DeviceName}, {DeviceSerial}")
        if NetworkID != net_id:
            print(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            logger.info(
                f'    Device: {DeviceSerial} is not part of Network: {NetworkID}\n    It is part of Network: {net_id}')
            continue
        else:
            print(f'    Getting LLDP neighbours for Device: {DeviceSerial}, {DeviceName}\n'
                  f'    Looking for Meraki Access Pints MR46 connected to the Device: {DeviceSerial}, {DeviceName}')
            logger.info(f'    Getting LLDP neighbours for Device: {DeviceSerial}, {DeviceName}\n'
                  f'    Looking for Meraki Access Pints MR46 connected to the Device: {DeviceSerial}, {DeviceName}')
            url = f"https://api.meraki.com/api/v1/devices/{DeviceSerial}/lldpCdp"
            headers = {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-Cisco-Meraki-API-Key": API_VALUE
            }
            payload = None
            response = requests.request('GET', url, headers=headers, data=payload)
            parsed = json.loads(response.text)
            apiPrintLogger = (json.dumps(parsed, indent=4, sort_keys=True))
            logger.info("Printing Full API RESPONSE\n")
            logger.info(apiPrintLogger)
            if response.status_code != 200:
                print(f'Error with status code {response.status_code}, text: {response.text}')
                logger.info(f'Error with status code {response.status_code}, text: {response.text}')
                return []
            portsList = []
            Meraki_MR46_List = []
            # Meraki_MR46 = str("Meraki")
            if len(parsed) == 0:
                print(f'    API call returned NO LLDP neighbours for the Device: {DeviceSerial}, {DeviceName}')
                logger.info(f'    API call returned NO LLDP neighbours for the Device: {DeviceSerial}, {DeviceName}')
                continue
            else:
                targetDict = parsed['ports']
                for dict in targetDict:
                    portsList.append(dict)
                    lldpInfo = str(parsed['ports'][dict])
                    if "systemName" not in lldpInfo:
                        print(f"    LLDP neighbour info does not have System Name field in API call return. Skipping "
                              f"Port: {dict}")
                        logger.info(f"    LLDP neighbour info does not have System Name field in API call return. Skipping "
                              f"Port: {dict}")
                        continue
                    else:
                        lldpName = parsed['ports'][dict]['lldp']['systemName']
                        lldpSourcePort = parsed['ports'][dict]['lldp']['sourcePort']
                        if dict != lldpSourcePort:
                            print(
                                f'    Skipping this port: {dict}. Port ID does not match LLDP Source Port in API '
                                f'response. Might be Uplink Fiber Port on Network Module')
                            logger.info(
                                f'    Skipping this port: {dict}. Port ID does not match LLDP Source Port in API '
                                f'response. Might be Uplink Fiber Port on Network Module')
                            continue
                        else:
                            if targetLLDPname in lldpName:
                                Meraki_MR46_List.append(dict)

            print(f'    Following ports have Meraki Access Points MR46 connected: {Meraki_MR46_List}\n'
                  f'    Checking configuration on those ports: NAME, TYPE, TAG, NATIVE vlan, ALLOWED vlans')
            logger.info(f'    Following ports have Meraki Access Points MR46 connected: {Meraki_MR46_List}\n'
                  f'    Checking configuration on those ports: NAME, TYPE, TAG, NATIVE vlan, ALLOWED vlans')
            for merakiAPport in Meraki_MR46_List:
                portResponse = getPort(DeviceSerial, merakiAPport)
                portName = portResponse["name"]
                portType = portResponse["type"]
                portVlan = str(portResponse["vlan"])
                portAllowedVlans = portResponse["allowedVlans"]
                portTag = portResponse["tags"]
                print(f'*******\nChecking Port: {merakiAPport}')
                logger.info(f'*******\nChecking Port: {merakiAPport}')
                if portName == targetPortName and portType == "trunk" and portVlan == targetNativeVlan and portAllowedVlans == targetAllowedVlan and portTag == targetTag:
                    print(f'    Port: {merakiAPport} already have correct configuration')
                    logger.info(f'    Port: {merakiAPport} already have correct configuration')
                    continue
                else:
                    changeName = None
                    changeType = None
                    changeTag = None
                    changeNativeVlan = None
                    changeAllowedVlans = None
                    if portName == targetPortName:
                        print(f'    NAME. Name on port {merakiAPport} already {targetPortName}. No changes needed')
                        logger.info(f'    NAME. Name on port {merakiAPport} already {targetPortName}. No changes needed')
                    else:
                        print(f'    *NAME. Port {merakiAPport} does not have correct Name. Port name will be changed to {targetPortName}\n'
                              f'        ACTUAL NAME: {portName}\n'
                              f'        TARGET NAME: {targetPortName}')
                        logger.info(f'    *NAME. Port {merakiAPport} does not have correct Name. Port name will be changed to {targetPortName}\n'
                              f'        ACTUAL NAME: {portName}\n'
                              f'        TARGET NAME: {targetPortName}')
                        changeName = targetPortName
                    if portType == "trunk":
                        print(f'    TYPE. Port {merakiAPport} already TRUNK. No changes needed')
                        logger.info(f'    TYPE. Port {merakiAPport} already TRUNK. No changes needed')
                    else:
                        print(f'    *TYPE. Port {merakiAPport} is not TRUNK. Port Type will be changed to TRUNK')
                        logger.info(f'    *TYPE. Port {merakiAPport} is not TRUNK. Port Type will be changed to TRUNK')
                        changeType = 'trunk'
                    if portTag == targetTag:
                        print(f'    TAG. Tag on port {merakiAPport} already {targetTag}. No changes needed')
                        logger.info(f'    TAG. Tag on port {merakiAPport} already {targetTag}. No changes needed')
                    else:
                        print(f'    *TAG. Port {merakiAPport} does not have correct Tag. Port Tag will be changed to {targetTag}\n'
                              f'        ACTUAL TAG: {portTag}\n'
                              f'        TARGET TAG: {targetTag}')
                        logger.info(f'    *TAG. Port {merakiAPport} does not have correct Tag. Port Tag will be changed to {targetTag}\n'
                              f'        ACTUAL TAG: {portTag}\n'
                              f'        TARGET TAG: {targetTag}')
                        # converting tag list to string for API
                        str_targetTag = str(targetTag[0])
                        changeTag = str_targetTag
                    if portVlan == targetNativeVlan:
                        print(f'    NATIVE vlan. Native vlan on port {merakiAPport} already {targetNativeVlan}. No changes needed')
                        logger.info(f'    NATIVE vlan. Native vlan on port {merakiAPport} already {targetNativeVlan}. No changes needed')
                    else:
                        print(f'    *NATIVE vlan. Port {merakiAPport} does not have correct Native Vlan. Port Native vlan will be changed to {targetNativeVlan}\n'
                              f'        ACTUAL NATIVE VLAN: {portVlan}\n'
                              f'        TARGET NATIVE VLAN: {targetNativeVlan}')
                        logger.info(f'    *NATIVE vlan. Port {merakiAPport} does not have correct Native Vlan. Port Native vlan will be changed to {targetNativeVlan}\n'
                              f'        ACTUAL NATIVE VLAN: {portVlan}\n'
                              f'        TARGET NATIVE VLAN: {targetNativeVlan}')
                        changeNativeVlan = targetNativeVlan
                    if portAllowedVlans == targetAllowedVlan:
                        print(f'    ALLOWED vlans. Tag on port {merakiAPport} already {targetAllowedVlan}. No changes needed')
                        logger.info(f'    ALLOWED vlans. Tag on port {merakiAPport} already {targetAllowedVlan}. No changes needed')
                    else:
                        print(f'    *ALLOWED vlans. Port {merakiAPport} does not have correct Allowed Vlans List. Port Tag will be changed to {targetAllowedVlan}\n'
                              f'        ACTUAL VLAN: {portAllowedVlans}\n'
                              f'        TARGET VLAN: {targetAllowedVlan}')
                        logger.info(f'    *ALLOWED vlans. Port {merakiAPport} does not have correct Allowed Vlans List. Port Tag will be changed to {targetAllowedVlan}\n'
                              f'        ACTUAL VLAN: {portAllowedVlans}\n'
                              f'        TARGET VLAN: {targetAllowedVlan}')
                        changeAllowedVlans = targetAllowedVlan

                    print(f'Making Following Changes for Port {merakiAPport}: \n'
                          f'    NAME. {changeName}\n'
                          f'    TYPE. {changeType}\n'
                          f'    TAG. {changeTag}\n'
                          f'    NATIVE vlan. {changeNativeVlan}\n'
                          f'    ALLOWED vlans. {changeAllowedVlans}')
                    logger.info(f'Making Following Changes for Port {merakiAPport}: \n'
                          f'    NAME. {changeName}\n'
                          f'    TYPE. {changeType}\n'
                          f'    TAG. {changeTag}\n'
                          f'    NATIVE vlan. {changeNativeVlan}\n'
                          f'    ALLOWED vlans. {changeAllowedVlans}')
                    updateAPport(DeviceSerial, merakiAPport, changeName, changeTag, changeType, changeNativeVlan, changeAllowedVlans)

# Step 8. Update Port Configuration based on OUI online.
def updatePortBasedOnOUI():
    df = pd.read_excel("Config_File.xlsx", 'Step 8', skiprows=0)
    df2 = df.replace(np.nan, '', regex=True)
    data = []

    for index, rows in df2.iterrows():
        data.append({rows[0]: rows[1]})
    networkID = data[0]['Network ID']
    targetOUI = data[1]['Serach for OUI']
    targetPortName = data[2]['Target Port Name']
    targetVlan = str(data[3]['Target Access Vlan'])
    targetTag = ([data[4]['Target Tag']])
    # should be only one tag. Otherwise, comparison with actual tag won't work
    clients = getCleints(networkID, targetOUI)
    for client in clients:
        status = client['status']
        recentConnection = client["recentDeviceConnection"]
        macAddress = client["mac"]
        switchName = client["recentDeviceName"]
        switchSerial = client["recentDeviceSerial"]
        switchPort = client["switchport"]
        print(f'CLIENT: {status}, {recentConnection}, {macAddress}, {switchName}, {switchSerial}, {switchPort}')
        logger.info(f'CLIENT: {status}, {recentConnection}, {macAddress}, {switchName}, {switchSerial}, {switchPort}')
        if targetOUI not in macAddress:
            print(f'Error with API call. MAC address returned: {macAddress} does not include target OUI: {targetOUI}.')
            logger.info(f'Error with API call. MAC address returned: {macAddress} does not include target OUI: {targetOUI}.')
        else:
            if status != "Online" or recentConnection != "Wired":
                continue
            else:
                print(f'#*#*#*#*#*#\nUpdating Port: {switchPort} on switch {switchName}:')
                logger.info(f'#*#*#*#*#*#\nUpdating Port: {switchPort} on switch {switchName}:')
                updateOUIport(switchName, switchSerial, switchPort, targetPortName, targetVlan, targetTag)

if __name__ == "__main__":

    main_menu = True
    while main_menu:
        os.system('clear')
        print("=====================================================")
        print("=    Meraki staging and Configuration tool          =")
        print("=====================================================\n\n")
        print("*[1] Claim Devices. Not available. In process")
        print("[2] Claim Devices and Assign them to Site")
        print("[3] See The List of Networks for ORG ID")
        print("[4] Update Network Name")
        print("[5] Update Device Atributes")
        print("[6] Update Switchports")
        print("[7] Configure Access Points ports")
        print("[8] Configure ports based on client's OUI")
        while True:
            step = input("Select a number [1-8]: ")
            logger.info(f'               SELECTED STEP: {step}')
            if step == "2":
                addToSite()
                input("\nPress Enter to continue:")
                break
            if step == "3":
                getNetworks()
                input("\nPress Enter to continue:")
                break
            if step == "4":
                updateNetworks()
                input("\nPress Enter to continue:")
                break
            if step == "5":
                updateDeviceAtributes()
                input("\nPress Enter to continue:")
                break
            if step == "6":
                updateSwitchports()
                input("\nPress Enter to continue:")
                break
            if step == "7":
                getCDPLLDPneighbours()
                input("\nPress Enter to continue:")
                break
            if step == "8":
                updatePortBasedOnOUI()
                input("\nPress Enter to continue:")
                break

#sys.stdout = old_stdout
#log_file.close()
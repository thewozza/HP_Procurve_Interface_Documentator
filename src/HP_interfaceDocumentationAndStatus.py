import csv
from netmiko import ConnectHandler
from netmiko.ssh_exception import NetMikoTimeoutException,NetMikoAuthenticationException
import pandas as pd

def getDocumentation(switch_ip,username,password):
    switch = {
        'device_type': 'hp_procurve',
        'ip': switch_ip,
        'username': username,
        'password': password,
        'secret': password,
        'port' : 22,          # optional, defaults to 22
        'verbose': False,       # optional, defaults to False
        'global_delay_factor': 4 # for remote systems when the network is slow
    }
    
    try:# if the switch is reponsive we do our thing, otherwise we hit the exeption below
        # this actually logs into the device
        net_connect = ConnectHandler(**switch)
        print switch_ip + ':we are in'
        net_connect.send_command('term len 1000')
        
        hostname = ''

        # I wanted to be sure of the hostname of the switch I'm on
        show_system = net_connect.send_command('show system').split("\n")
        for system_line in show_system:
            if len(system_line) <= 1: # we ignore empty lines
                continue
            elif "Name" in system_line.split()[1]:
                hostname = system_line.split()[3]

        print switch_ip + ':this is ' + hostname
                
        interfaceDict = {}
        
        # this gives a table of interfaces and their basic status        
        show_interface_brief = net_connect.send_command('show interface brief').split("\n")
        
        # walk through the table and start building the interface dictionary structure
        for interface in show_interface_brief:
            if len(interface) <= 1: # we ignore empty lines
                continue
            if interface.split()[0].isdigit():
                interfaceDict[interface.split()[0]] = {}
                
                # due to weirdness in exporting, we have to create
                # an entry for the port number, this goes away at the end
                interfaceDict[interface.split()[0]]['port'] = int(interface.split()[0])
                
                # if an SFP interface has no module, it messes up the columns
                # so we ignore the line if this happens
                if "|" in str(interface.split()[2]):
                    interfaceDict[interface.split()[0]]['port'] = int(interface.split()[0])
                    interfaceDict[interface.split()[0]]['status'] = interface.split()[5]
                    interfaceDict[interface.split()[0]]['mode'] = interface.split()[6]
                    # we want to track if the interface went up or down in the logs
                    # so this is just a helpful place to initialize the variable
                    # just in case there are no entries in the log, so at least we
                    # have a zero here
                    interfaceDict[interface.split()[0]]['onlineCount'] = 0
        
        print switch_ip + ':got the interfaces'
        
        vlanList = []
        
        # this gives a table of the known vlans on the switch
        show_vlans = net_connect.send_command('show vlans').split("\n")
        
        # create a usable list of vlans
        for vlans in show_vlans:    
            if len(vlans) <= 1: # we ignore empty lines
                continue
            if vlans.split()[0].isdigit():
                vlanList.append(vlans.split()[0])
        
        print switch_ip + ':got the vlans'
        
        # walk through the vlans and populate our interface dictionary with the memberships
        for vlan in vlanList:
            show_vlan_membership = net_connect.send_command('show vlan ' + vlan).split("\n")
            
            for vlan_membership in show_vlan_membership:    
                if len(vlan_membership) <= 1: # we ignore empty lines
                    continue
                if vlan_membership.split()[0].isdigit():
                    interfaceDict[vlan_membership.split()[0]]['vlan' + vlan] = vlan_membership.split()[1]

        print switch_ip + ':got the vlan memberships'

        # we gather the logs - filtering on the "on-line" keyword
        logData = net_connect.send_command('show logging on-line').split("\n")

        print switch_ip + ':got the log data'
        
        # we always sanely disconnect
        net_connect.disconnect()
        print switch_ip + ':sane disconnect'
        
        # now we parse the logs looking for interface up/down status entries
        for logLine in logData:
            if len(logLine) <= 1: # we ignore empty lines
                continue
            
            # sometimes the formatting is off, but if we see the keyword "port"
            # then we know that the interface number is next
            # so this just initializes the variable
            nextIsPort = 0 
            if logLine.split()[0] == "I": # the log entries we care about are preceeded with an "I"
                for word in logLine.split():
                    if nextIsPort: # this only happens if the previous word in the log line was "port"
                        if interfaceDict[word]['onlineCount'] >= 0:
                            interfaceDict[word]['onlineCount'] += 1
                        else:
                            interfaceDict[word]['onlineCount'] = 1
                        nextIsPort = 0
                    if word == "port":
                        nextIsPort = 1

        print switch_ip + ':processed the log data'

        # we convert the python dictionary to a pandas dataframe
        # I don't know what that means
        dataframe = pd.DataFrame(interfaceDict).T # transpose the tables so interfaces are in a column
        dataframe.sort_values(by=['port'], inplace=True) # sort the values by the "port" column we made
        dataframe = dataframe.reset_index(drop=True) # reset the index to match this
        
        # we want to re-order the columns, so we pull the names into a list
        dfColumns = dataframe.columns.tolist()
        # we change the order so that the "port" column header is first
        dfColumns.insert(0, dfColumns.pop(2))
        # then we re-insert that topology into the dataframe
        dataframe = dataframe[dfColumns]
        
        # finally we can export as an excel document
        dataframe.to_excel('/home/paul/' + hostname + '.xls',index=False)
        
        print switch_ip + ':exported to /home/paul/' + hostname + '.xls'
    except (NetMikoTimeoutException, NetMikoAuthenticationException):
        print switch_ip + ':no_response'

switches = csv.DictReader(open("switches.csv"))

try:
    for row in switches:
        getDocumentation(row['IP'],row['username'],row['password'])
except IndexError:
    pass
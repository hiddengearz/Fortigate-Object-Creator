import os
import sys
import time
import xlrd


def main():
    startTime = time.time() #Begin recording how long it takes to execute this script

    file = ("systems.xlsx") #default file name
    wb = xlrd.open_workbook(file) #open the excel file


    for x, sheet_name in enumerate(wb.sheet_names()): #for each sheet in the workbook
        print(f'######################################################################{sheet_name}######################################################################')
        sheet = wb.sheet_by_index(x)
        config_file = sheet_name + '.conf' #FGT config file is hardcoded to be sheet_name.conf so make sure you rename it
        notAdded = [] #Servers that already exist in the config (not perfect, can definitely be improved)
        if os.path.isfile(config_file): #Make sure the config file exists
            with open(config_file) as f: #Open the file
                content = f.read()
                for i in range(sheet.nrows): #For each row in the file
                    name = "Srv_" + sheet.cell_value(i, 0)
                    if sheet.cell_value(i, 0) != "": #If the cell isn't empty
                        if sheet.cell_value(i, 0) not in content: #if the string in the cell isn't in the config file
                            print("config firewall address")
                            print(f"edit \"{name}\"")
                            print(f'set associated-interface "Trust"')
                            try: #Column B has to be FQDN, if there is a FQDN use that
                                if(sheet.cell_value(i, 1)) != "":
                                    print(f'set type fqdn')
                                    print(f'set fqdn "{sheet.cell_value(i, 1)}"')
                            except: #Otherwise coulmn C needs to have an IP, use the IP instead
                                if (sheet.cell_value(i, 2)) != "":
                                    print(f'set subnet {sheet.cell_value(i, 2)} 255.255.255.255')
                            print('end')

                            #Add the server to the object group
                            print("config firewall addrgrp")
                            print(f'edit "Blocked 2008 Servers"')
                            print(f'append member "{name}"')
                            print(f'end')

                        else: #If the cells value is in the config you need to add it to the group manually
                            notAdded.append(sheet.cell_value(i, 0))

        #Print out all of the servers you need to add manually to the user
        print("############################ MANUALLY ADD ########################")
        for server in notAdded:
            print(server)



    print('The script took {0} second!'.format(time.time() - startTime))

if __name__ == "__main__":
    sys.exit(main())
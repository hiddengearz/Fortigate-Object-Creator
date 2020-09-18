import os
import sys
import time
import xlrd


def main():
    startTime = time.time() #Begin recording how long it takes to execute this script
    help = False;
    file = ""

    for i, arg in enumerate(sys.argv):
        if '-h' in arg or '-help' in arg:
            help = True;
        elif '-f' in arg or '-file' in arg:
            file = str(sys.argv[i+1])

    if help:
        print('Fortigate object creator!')
        print('Adding dozens or even hundreads of objects to a fortigate firewall is a tedious and time consuming '
              'task that can take hours. This script speeds up this process by automatically creating the commands needed for creating '
              'these objects for you. Once it\s done all you need to do is copy and paste the command into the Fortigate CLI.\n')
        print(
            'This script will avoid making duplicate objects by checking if an object with the same name is in the config\n')
        print('###REQUIREMENTS###')
        print(
            '1) An Excel workbook (.xlsx format) that contains the following: Column A: Object Name Column B: DNS '
            'Name Or Coulmn C: IP address. If Both Column B and C is provided, DNS will be prioritized. '
            'Please note this script supports multiple sheets, allowing you to generate objects for multiple firewalls at once')
        print(
            '2) The latest copy of you\'re firewall configuration. The script will compare new objects to old once to avoid creating duplicates.'
            'Please note the firewall config name must be the same as the worksheet in excel.'
            'Please see example.xlsx and the example .conf files')
        print('-f [file], -file [file] the location of the file containing a list of the systems')

    else:
        file = (file) #default file name
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
                        name = sheet.cell_value(i, 0)
                        if sheet.cell_value(i, 0) != "": #If the cell isn't empty
                            if sheet.cell_value(i, 0) not in content: #if the string in the cell isn't in the config file
                                print("config firewall address")
                                print(f"edit \"{name}\"")
                                print(f'set associated-interface "Trust"')
                                try: #Column B has to be FQDN, if there is a FQDN use that
                                    if(sheet.cell_value(i, 1)) != "":
                                        print(f'set type fqdn')
                                        print(f'set fqdn "{sheet.cell_value(i, 1)}"')
                                    else:
                                        print(f'set subnet {sheet.cell_value(i, 2)} 255.255.255.255')
                                except: #Otherwise coulmn C needs to have an IP, use the IP instead
                                    if (sheet.cell_value(i, 2)) != "":
                                        print(f'set subnet {sheet.cell_value(i, 2)} 255.255.255.255')
                                print('end')

                                #Add the server to the object group
                                print("config firewall addrgrp")
                                print(f'edit "Xerox Printers"')
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
#!/usr/bin/python

import sys
import os
import re
import csv
import time

from bs4 import BeautifulSoup
from optparse import OptionParser
from xlsxwriter import Workbook




def BuildExcelTable(text_file_list,base_name):

	newName = base_name + '.xlsx'
	workbook = Workbook(newName, {'strings_to_numbers': True})
	
	for i in range(len(text_file_list)):
		print(text_file_list[i])
		worksheet_name = text_file_list[i] [:12]
		worksheet = workbook.add_worksheet(worksheet_name)
		FileReader = csv.reader((open(text_file_list[i], 'r', encoding='latin1')),
                            delimiter='*', quotechar='"')
		for rowx, row in enumerate(FileReader):
			for colx, value in enumerate(row):
				worksheet.write(rowx, colx, value)

	workbook.close()


def ExtractPlugin(Nessus_file, output_file, plugin):
	print("Processing plugin information. Please wait...")
	sys.stdout=open(output_file,'w')
	infile = open(Nessus_file,"r")
	contents = infile.read()
	soup = BeautifulSoup(contents,'xml')
	hosts = soup.find_all('ReportHost')
	for host in hosts:
		items = host.find_all('ReportItem')
		# print("\n")
		# print('Plugin information:')
		i=0
		for item in items:
			# print(plugin)
			# print(item.get("pluginID"))
			if plugin == item.get("pluginID"):
				print("********************************")
				print('Plugin ID: ' + item.get("pluginID"))
				print("Machine name: " + host.get('name'))
				print("Port: " + item.get('port'))
				i=i+1
				print(i)
				hosttags = host.find_all('HostProperties')
				for hosttag in hosttags:
					innertaglisting = hosttag.find_all("tag")
					for message in innertaglisting:
						search_string = str(message.attrs)
						if search_string.find("host-ip") != -1:
							print("Host-ip:	" + message.text)
							# print("\n")
						if search_string.find("mac-address") != -1:
							print("MAC address:	" + message.text)
							# print("\n")
	infile.close()
	sys.stdout.flush()
	time.sleep(5)


def ExtractPluginToXLS(Nessus_file, output_file, plugin):
	sys.stdout=open(output_file,'w')
	infile = open(Nessus_file,"r")
	contents = infile.read()
	soup = BeautifulSoup(contents,'xml')
	hosts = soup.find_all('ReportHost')

	print("HostIP*name*hostname*os*OperatingSystem*SystemType*MACaddress*port*plugin")

	for host in hosts:
		items = host.find_all('ReportItem')
		i=0
		for item in items:
			if plugin == item.get("pluginID"):
				# print('Plugin ID: ' + item.get("pluginID"))
				value_plugin = item.get("pluginID")
				# print("Machine name: " + host.get('name'))
				value_machine_name = host.get('name')
				# print("Port: " + item.get('port'))
				value_port = item.get('port')
				value_host_ip = " "
				value_MAC_address = " "
				value_os = " "
				value_hostname = " "
				value_operating_system = " "
				value_system_type = " "
				
				i=i+1
				hosttags = host.find_all('HostProperties')
				for hosttag in hosttags:
					innertaglisting = hosttag.find_all("tag")
					for message in innertaglisting:
						search_string = str(message.attrs)
						# value_host_ip = " "
						# value_MAC_address = " "
						# value_os = " "
						# value_hostname = " "
						# value_operating_system = " "
						# value_system_type = " "
						if search_string.find("host-ip") != -1:
							value_host_ip = message.text
						if search_string.find("mac-address") != -1:
							value_MAC_address = message.text
						if search_string.find("os") != -1:
							value_os = message.text
						if search_string.find("hostname") != -1:
							value_hostname = message.text
						if search_string.find("operating-system") != -1:
							value_operating_system = message.text
						if search_string.find("system-type") != -1:
							value_system_type = message.text
				print(value_host_ip + "*" + value_machine_name + "*" + value_hostname + "*" + value_os + "*" + value_operating_system + "*" + value_system_type + "*" + value_MAC_address + "*" + value_port + "*" + value_plugin)
							
	infile.close()
	sys.stdout.flush()
	time.sleep(5)



def CombineHostnamePlugins(Nessus_file):
	print("Generating reports...")

	Basetime = time.strftime("%Y%m%d-%H%M%S")
	base_name = "Names_plugins" + Basetime

	text_file_list = []

	filename55472 = "Output_55472_" + Basetime + ".txt"
	ExtractPluginToXLS(Nessus_file, filename55472, "55472")
	text_file_list.append(filename55472)
	
	filename10150 = "Output_10150_" + Basetime + ".txt"
	ExtractPluginToXLS(Nessus_file, filename10150, "10150")
	text_file_list.append(filename10150)

	filename46180 = "Output_46180_" + Basetime + ".txt"
	ExtractPluginToXLS(Nessus_file, filename46180, "46180")
	text_file_list.append(filename46180)

	time.sleep(10)

	BuildExcelTable(text_file_list,base_name)


def ParseNessusFile(Nessus_file,output_file_name):
	base_name = os.path.splitext(output_file_name)[0]
	# text_file = base_name + '.txt'
	sys.stdout = open(output_file_name,'w')

	infile = open(Nessus_file,"r")
	contents = infile.read()
	soup = BeautifulSoup(contents,'xml')
	hosts = soup.find_all('ReportHost')

	print("Machine*IP*PluginNB*PluginName*Port*Protocol")

	for host in hosts:
		## print("**************\n")
		## print("Machine:IP:PluginNB:PluginName:Port:Protocol")
		## print("Machine name: " + host.get('name'))
		MachineName = host.get('name')
		HostIP = " "
		# print(host.get('name'))
		hosttags = host.find_all('HostProperties')
		for hosttag in hosttags:
			
			innertaglisting = hosttag.find_all("tag")
			# print(innertaglisting[-2])

			for message in innertaglisting:
				
				# print(message.attrs)
				# print(message.text)
				
				search_string = str(message.attrs)
				if search_string.find("host-ip") != -1:
					## print("\n")
					## print("Host-ip:	" + message.text)
					HostIP = message.text
					# print(message.text)


		items = host.find_all('ReportItem')

		## print("\n")
		## print('Plugin information:')
		for item in items:
			# print('pointer')
			print(MachineName + "*" + HostIP + "*" + item.get("pluginID") + "*" + item.get("pluginName") + "*" + item.get("port") + "*" + item.get("protocol"))
			## print('Plugin ID: ' + item.get("pluginID"))
			## print(item.get("pluginID") + ":")
			# print(item.get("pluginID"))
			## print('Plugin name: ' + item.get("pluginName"))
			# print(item.get("pluginName"))
			## print('Port: ' + item.get("port"))
			## print('Protocol: ' + item.get("protocol"))
			# print(item.find('plugin_output'))
			## print("\n")
	text_file_list = []
	text_file_list.append(output_file_name)


	BuildExcelTable(text_file_list,base_name)



def main():


	choice ='0'
	while choice == '0':
	    print("MENU")
	    print("Usage: python3 dparser report.xml output.txt|xls")
	    print("Choose 1 for simple text and xls output of hosts and plugins")
	    print("Choose 2 for text report for one plugin")
	    print("Choose 3 for .xls report for one plugin")
	    print("Choose 4 to find all hosts / scan should be DNS aware for comprehensive results")
	    print("Choose 5 to exit")
	    # print(sys.argv)
	    choice = input ("Please make a choice: ")


	if choice == "5":
		print("Exiting: ...")

	elif choice == "1":
		Nessus_file = input ("Please indicate the .nessus file: ")
		while True:
			try:
				infile = open(Nessus_file,"r")
			except FileNotFoundError:
				print("Wrong file or file path")
				break
			else:
				base_extension = os.path.splitext(Nessus_file)[1]
				if base_extension != ".nessus":
					print("This file does not have the correct extension! We're quitting.")
					break
				else:
					Basetime = time.strftime("%Y%m%d-%H%M%S")
					output_file = "ParserOutput" + "_" + Basetime + ".txt"
					ParseNessusFile(Nessus_file, output_file)
					break

	elif choice == "2":
		Nessus_file = input ("Please indicate the .nessus file: ")
		while True:
			try:
				infile = open(Nessus_file,"r")
			except FileNotFoundError:
				print("Wrong file or file path")
				break
			else:
				base_extension = os.path.splitext(Nessus_file)[1]
				if base_extension != ".nessus":
					print("This file does not have the correct extension! We're quitting.")
					break
				else:
					plugin = input ("Plugin number: ")
					Basetime = time.strftime("%Y%m%d-%H%M%S")
					output_file = plugin + "_" + Basetime + ".txt"
					ExtractPlugin(Nessus_file, output_file, plugin)	
					break

	elif choice == "3":
		Nessus_file = input ("Please indicate the .nessus file: ")
		while True:
			try:
				infile = open(Nessus_file,"r")
			except FileNotFoundError:
				print("Wrong file or file path")
				break
			else:
				base_extension = os.path.splitext(Nessus_file)[1]
				if base_extension != ".nessus":
					print("This file does not have the correct extension! We're quitting.")
					break
				else:
					plugin = input ("Plugin number: ")
					Basetime = time.strftime("%Y%m%d-%H%M%S")
					output_file = plugin + "_" + Basetime + ".txt"
					ExtractPluginToXLS(Nessus_file, output_file, plugin)	
					break					

	elif choice == "4":
		Nessus_file = input ("Please indicate the .nessus file:")
		while True:
			try:
				infile = open(Nessus_file,"r")
			except FileNotFoundError:
				print("Wrong file or file path")
				break
			else:
				base_extension = os.path.splitext(Nessus_file)[1]
				if base_extension != ".nessus":
					print("This file does not have the correct extension! We're quitting.")
					break
				else:
					CombineHostnamePlugins(Nessus_file)
					break
	else:
		print("I don't understand your choice. Please run the program again :)")


if __name__ == '__main__':
    main()


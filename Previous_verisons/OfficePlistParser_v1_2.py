#
# Recent Documents for MS Office are stored in the com.microsoft.office.plist file
# This program parses out the access dates, file paths and file names. MS Office 2010
# also supplies the long name which is also parsed out.
#
# Required: biplist from http://github.com/wooster/biplist
# use easy install : sudo easy_install biplist
#
# OfficePlistParser.py = Python Script to parse an MS Office 2008, 2012 plist file
#
# Copyright (C) 2013 Mari DeGrazia (arizona4n6@gmail.com)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You can view the GNU General Public License at <http://www.gnu.org/licenses/>
#
# Version History:
# v1.2 2013-8-4
#	
# Special thanks to Adrian Leong for aiding in the research of the
# com.microsoft.office.plist file
#

__author__ = 'arizona4n6@gmail.com (Mari DeGrazia)'
__version__ = '1.2'
__copyright__ = 'Copyright (C) 2013 Mari DeGrazia'
__license__ = 'GNU'

import sys
import datetime
from optparse import OptionParser
import string
from biplist import *


###############################  Functions  ################################################

#Convert HEX to yyyy-mm-dd hh:mm:ss
def convert_hex_to_Hfs(hex_date):

	hfs32 = []
	 
	 #the hex date looks like 0000 B951 20CE, we need to switch it to Big Endian, 0xCE20 51B9.
	hfs32.extend([hex_date[10],hex_date[11],hex_date[8],hex_date[9],hex_date[6],hex_date[7],hex_date[4],hex_date[5]])
	hfs32_big_Endian = "".join(hfs32)
	
	 #HFS+ time is the number of seconds since 1/1/1904. We will need to subract this from Epoch, which is number of seconds since 1/1/1970		 
	hfs_timestamp = str(datetime.datetime.fromtimestamp(int(hfs32_big_Endian,16)-2082844800))
	 
	return(hfs_timestamp)
	

#Takes the file alias, and returns the file path for Office 2008, and file path and long name for Office 2010
#MRUType is either '2008' or '2010'

def get_path(this_file_aliases,MRUType):

	#encode to hex so we can search for the end of the path designated by Hex 000e00
	#the full file name is the next chunk after the path for 2010	
	path,temp1,full_file_name = this_file_aliases.encode('hex').partition('000e00')
		
	#we know that 000200 exisits right before the path starts, but sometimes there are false postives before the path starts. Since 000200 will not occur in the file path
	#reverse look through the path for the first occurance of the string.

	index = path.rfind('000200',0,len(path))		
		
	#now slice that puppy out, remember to strip out the leading 000200 and an additional obnoxious byte
	path=path[index+8:len(path)]
			
	if MRUType == "2008":				
		return(path.decode('hex'))
	
	if MRUType == "2010":

		full_file_name,temp1,temp2 = full_file_name.partition('000f00')
					
		#remove those pesky leading bytes
							
		full_file_name = full_file_name[6:]	
		
		return(path.decode('hex'), full_file_name.decode('hex'))

def remove_ascii_non_printable(str):
	 return ''.join([ch for ch in str if ord(ch) > 31 and ord(ch) < 126 or ord(ch) ==9])


###############################  MAIN  ################################################

#help menu, etc
usage = "\n%prog [-h|help] [-f file] [-o output]\nExample: OfficePlistParser.py -f com.microsoft.office.plist -o recentdocs.tsv"

parser = OptionParser(usage=usage)

parser.add_option("-f", "--f", dest = "infile", help = "binary plist file", metavar = "input.plist")
parser.add_option("-o", "--o", dest = "outfile", help = "output to a tsv file", metavar = "output.tsv")

(options,args)=parser.parse_args()

if options.infile == None or options.outfile == None:
	parser.error("Filename not given")
	parser.print_help()
	

#try to open the files, if not there bail out
try:
	f = open(options.infile, "rb")
except IOError as e:
	print 'File Not Found :' + options.infile
	exit(0)

#get the output file ready
output = open(options.outfile, 'w')

#make sure the file is a binary plist file
file_header = f.read(6)
if str(file_header) != 'bplist':
 	print "Sorry, " + options.infile + 'is not a Binary Plist file'
	exit(0)

#be polite and move back to the beginning of the file
f.seek(0)

plist = readPlist(options.infile)

MS2008Entries = {} 	#array to hold Office 2008 Entires: MRUID,Path and Date
MS14 = {} 			#array to hole Office 14 (2010) Entires: Key
userinfo = {} 		#array to hold userinfo: Keyname and Value
count = 0 			#counter for records processed
j = 0				#counter for the MS14 incrementor

for key,value in plist.iteritems():
	
	#for giggles, let's see if there is any user information in this file
	if "User" in key and value != "":
		userinfo[key]=value
		
	#find the 2008 File Aliases, create the array and add the parsed path	
	if "2008\\File Aliases" in key:
			
		keyname = key.split('\\')
		MRUID = keyname[2]	
		
		if MS2008Entries.has_key(MRUID) == True:
			MS2008Entries[MRUID][0]=get_path(value,"2008")
			
		else:	
			MS2008Entries[MRUID]=[get_path(value,"2008"),"No Date"]
		
	#find the 2008 Access Dated, create the array and add the converted timestamp	
	if "2008\\MRU Access Date" in key:	
		
		keyname = key.split('\\')
		MRUID = keyname[2]
		
		#convert the date
		hex_access_date = value.encode('hex')
		date=(convert_hex_to_Hfs(hex_access_date))
			
		#add the values to the array		
		if MS2008Entries.has_key(MRUID) == True:		
			MS2008Entries[MRUID][1]=date	
		
		else:
			MS2008Entries[MRUID]=["No File Alias",date] 
			

 	#check for Office 2010 MRU Format		
	if key == "14\File MRU\XCEL" or key == "14\File MRU\PPT3" or key == "14\File MRU\MSWD" :
		#loop through each item under the MRU Key		
		i=0		
		for item in value:		
			
			this_access_date = value[i]["Access Date"]
			this_file_alias = value[i]["File Alias"]
	  		i = i+1
			file_path = get_path(this_file_alias, "2010")	
			
			#convert the access date to hex so it can be converted to a timestamp.	 
			hex_access_date = this_access_date.encode('hex')
		 	
			#add the values to the array
			MS14[j]=[key,convert_hex_to_Hfs(hex_access_date),file_path[0],file_path[1]]
			j = j+1
			
			

#write out the information
output.write("MRUID\tAccess Date(UTC)\tFile Alias\tFile Name (Office2010 Only)\t\n")

#print out the dictionary holding the 2008 MRU filepath and date
for key,value in MS2008Entries.iteritems():
	output.write(key + "\t" + remove_ascii_non_printable(value[1]) + "\t" + remove_ascii_non_printable(value[0]) + "\n")
	
#print out the dictionary holding the Office 14 (2010) Entries
for key,value in MS14.iteritems():
	output.write(value[0] + "\t" + remove_ascii_non_printable(value[1]) + "\t" + remove_ascii_non_printable(value[2]) + "\t" + remove_ascii_non_printable(value[3])+"\n")

print "\nNumber of Keys processed: ", len(MS2008Entries)+len(MS14)

#if there was userinformtion, print it
if len(userinfo):
	print "\nUser Information:"
	for key,value in userinfo.iteritems():
		print("Key: " + key + "\t\tValue: " + value)
else:
	print "\nUser Information not located"

output.close


	 

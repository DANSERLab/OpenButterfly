# Script to search differnt data files, grab relevant time frames, then consolidate them into one excel work book where each sheet is a differnt trial.
# requires this file to be at the top of the directory with all the other data folders
# the directory path should be of "DATA/Session#'s/Userfiles"



import os
import os.path
import copy
import csv
#import xlsxwriter module --> may need to download pip then run the command "pip install xlsxwriter"
import xlsxwriter 


#==============================================================
# Variable Block

subject="Mike"						# Subject's name as formated in File name
session="Session1"					# Session number to check
clock_start_times = []				# Array of global clock times that were recorded
hr_clock_start_time = []			# Time the HR monitor began recording
hr_offset_min = []
hr_offset_sec = []
heart_rate_start_times = []			# Array of all the beginning times to clip
gsr_start_times = []				# Array of all the beginning times to clip
date = []							# Date of session
# number_of_recordings = []			# The number of Muse/Neulog recordings
recording_transitions = []			# Array for each "part".  Shows transitions from recording 1 to recording 2 to recording 3, etc
muse_start_time = []				# Array of all the begining times to clip

#==============================================================
# FUNCTION:  make excel workbook with correct name and sheet names
# I'm working on this section right now

def make_excel_book(name,session):

	if session in ["Session2","Session3","Session4","Session5","Session6"]:
		# Workbook() takes one, non-optional, argument  
		# which is the filename that we want to create. 
		workbook = xlsxwriter.Workbook('clipped_' + name + '_' + session + '.xlsx') 
		  
		# By default worksheet names in the spreadsheet will be  
		# Sheet1, Sheet2 etc., but we can also specify a name. 
		worksheet = workbook.add_worksheet("baseline") 
		worksheet = workbook.add_worksheet("FAR1") 
		worksheet = workbook.add_worksheet("SAR1") 
		worksheet = workbook.add_worksheet("SR1") 
		worksheet = workbook.add_worksheet("FAR2") 
		worksheet = workbook.add_worksheet("SAR2") 
		worksheet = workbook.add_worksheet("SAR2") 
		worksheet = workbook.add_worksheet("FAR3") 
		worksheet = workbook.add_worksheet("SAR3") 
		worksheet = workbook.add_worksheet("SR3") 
		  
		# Finally, close the Excel file 
		# via the close() method. 
		workbook.close() 

	elif session in ["Session7","Session8","Session9","Session10","Session11","Session12"]:
		# Workbook() takes one, non-optional, argument  
		# which is the filename that we want to create. 
		workbook = xlsxwriter.Workbook('clipped_' + name + '_' + session + '.xlsx') 
		  
		# By default worksheet names in the spreadsheet will be  
		# Sheet1, Sheet2 etc., but we can also specify a name. 
		worksheet = workbook.add_worksheet("baseline") 
		worksheet = workbook.add_worksheet("ExR1") 
		worksheet = workbook.add_worksheet("AbdR1") 
		worksheet = workbook.add_worksheet("MxdPr1") 
		worksheet = workbook.add_worksheet("MxdCr1") 
		worksheet = workbook.add_worksheet("ExR2") 
		worksheet = workbook.add_worksheet("AbdR2") 
		worksheet = workbook.add_worksheet("MxdPr2") 
		worksheet = workbook.add_worksheet("MxdPr2") 
		  
		# Finally, close the Excel file 
		# via the close() method. 
		workbook.close() 

	else:
			print("This is not a valid session name. Should be Session1, Session2,...,Session12.  No Session0 or Session1")

#==============================================================
# FUNCTION:  Start Times ./start_time/Session1/time_Mike_2019-04-23.csv
	#make array of start times
	# csv file with column 0 of HR,Muse/GSR,Baseline,Trial1, Trial 2... 9, Survey, ROM
	#				column 1 of Clock time of the form HOUR:MIN
	#				column 2 of GSR time
def make_start_time_array(name,session):

	# his function will be passed a name (Mike) and the session number (Session1)
	# fill the array with the clock times and fill the array with neulog times
	# so there should be 2 arrays at the end, of different length


	

	with open(file_time, 'rb') as csvFile:
    		reader = csv.reader(csvFile)							# create reader wrapped around an object.  These means use one time and done, so can't call twice.
    		originalFile_time = list(reader)								# make a list of spread sheet, need a list to index to a specfic cell and overwrite old spread sheet
	
		
    	# loop for making a column an array
		#for row in originalFile:
		#	gsr_start_times.append(row[2])
		global number_of_recordings
		date.append(originalFile_time[2][1])
		number_of_recordings = int(originalFile_time[3][1])

		if session in ["Session1","Session2","Session3","Session4","Session5","Session6"]:
			i = 7
			while i <= 17:
				gsr_start_times.append(originalFile_time[i][2])
				recording_transitions.append(originalFile_time[i][3]) 
				i += 1

			j = 5
			while j <= 18:
				clock_start_times.append(originalFile_time[j][1]) 
				j += 1


 		elif session in ["Session7","Session8","Session9","Session10","Session11","Session12"]:
			# make speciic cells of a column into an array
			i = 8
			while i <= 22:
				gsr_start_times.append(originalFile_time[i][2])
				recording_transitions.append(originalFile_time[i][3]) 
				i += 1

			j = 5
			while j <= 23:
				clock_start_times.append(originalFile_time[j][1]) 
				j += 1

		


		# Get heart 	rate start times and muse start times.  Figure out differences 



	 	csvFile.close()
		
#==============================================================
# FUNCTION: Heart Rate ./heart_rate/Session1/HR_Mike_Session1_2019-04-23.csv
def heart_rate_clock_start_time(name,session):

	# First get the global start time for the Heart Rate Recording and break into hours,min, and sec

	with open(file_HR, 'rb') as csvFile:
    		reader = csv.reader(csvFile)							# create reader wrapped around an object.  These means use one time and done, so can't call twice.
    		originalFile_HR = list(reader)								# make a list of spread sheet, need a list to index to a specfic cell and overwrite old spread sheet
	
	hr_clock_start_time.append(originalFile_HR[1][3])

	global hr_hour, hr_min, hr_sec
	hr_hour, hr_min, hr_sec = hr_clock_start_time[0].split(":")
	hr_hour = int(hr_hour)
	hr_min = int(hr_min)
	hr_sec = int(hr_sec)
	
	csvFile.close()

	
#==============================================================
# FUNCTION: find the offset from beginning of HR file to beginning of particular Muse file
def hr_find_offsets(input_file):
	# Next get start minute of Muse which should be later than the Heart rate.  
	# Need to take care of cases where there are multiple muse recordings
	with open(input_file, 'rb') as csvFile:
    		reader = csv.reader(csvFile)							# create reader wrapped around an object.  These means use one time and done, so can't call twice.
    		originalFile_muse = list(reader)	

    	muse_start_time.append(originalFile_muse[1][0])

    	muse_hour, muse_min, muse_sec = muse_start_time[0].split(":")
    	muse_date, muse_hour = muse_hour.split(" ")
    	muse_sec, muse_decisec = muse_sec.split(".")
    	muse_hour = int(muse_hour)
    	muse_min = int(muse_min)
    	muse_sec = int(muse_sec)
    	muse_decisec = int(muse_decisec)



    	csvFile.close()

    	# If both muse and HR hours match then just subtract HR min and sec from Muse min and sec to get offse
    	# If muse and HR hours differ subtract HR min and sec from 60 then add to Muse min and sec

    	if hr_hour == muse_hour:
    		diff_min = muse_min - (hr_min + 1)
    		diff_sec = muse_sec + (60-hr_sec)
    		if diff_sec > 59:
    			diff_min = diff_min + 1
    			diff_sec = diff_sec - 60

    		hr_offset_min.append(diff_min)
    		hr_offset_sec.append(diff_sec)


    	if hr_hour!=muse_hour:
    		diff_min = (60 - (hr_min + 1)) + muse_min
    		diff_sec = (60 - hr_sec) + muse_sec
    		if diff_sec > 59:
    			diff_min = diff_min + 1
    			diff_sec = diff_sec - 60

    		hr_offset_min.append(diff_min)
    		hr_offset_sec.append(diff_sec)



#==============================================================
# FUNCTION: Muse ./muse/Session1/muse_Mike_Session1_2019-04-23.csv
# This function should be take a singular start time (not an array), copy the appropiate 
# 60 seconds of data, and append it to the sheet of the excel book.  
# The time to beging will be a string you can search for in the Muse file.

#def clip_muse(time_to_begin_at, sheet_to_paste_to)

	
#==============================================================
# FUNCTION: GSR ./gsr/Session1/gsr_Mike_Session1







################################################################
################################################################
	           		# MAIN CODE BLOCK #
################################################################
################################################################


# Current working directory
fileDir = os.path.dirname(os.path.realpath('__file__'))
#For accessing the file in a folder contained in the current folder
file_time = os.path.join(fileDir, 'start_time/', session, 'time_'+ subject +'.csv')
file_gsr = os.path.join(fileDir, 'gsr/', session, 'gsr_'+ subject +'.csv')
file_HR = os.path.join(fileDir, 'heart_rate/', session, 'HR_'+ subject +'.csv')
file_muse = os.path.join(fileDir, 'muse/', session, 'muse_'+ subject +'.csv')



# take in subject and session 
#subject = raw_input("Enter subject's name to consolidate: ")
#session = input("Enter subject's session to consolidate (ie Sesion2, Session5): ")
#print ("You entered " + subject + "and " + session) 

make_start_time_array(subject,session)
heart_rate_clock_start_time(subject,session)

for i in range(len(recording_transitions)): # should be length ofrecording transitions not number
	if i == 1:
		hr_find_offsets(file_muse)
	if i > 1 and recording_transitions[i] > recording_transitions[i-1]:
		file_muse_updated = os.path.join(fileDir, 'muse/', session, 'muse_'+ subject + '_part' + recording_transitions[i] + '.csv')
		#hr_find_offsets(file_muse_updated). # Need to make sure there are matching parts for each recording

print(hr_offset_min)
print(hr_offset_sec)


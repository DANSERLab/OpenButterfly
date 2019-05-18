

# Script to parse cruzid.txt for grade and write to canvas spreadsheet
# requires grade to be in file "cruzid.txt"
# the directory path should be "  "  with this file in " ", cruzid.txt within " ", and canvas_spreadsheet_to_import.csv within ""
# format of last line of cruzid.txt  "Grade: 5/10 (50.0%)"



import csv
import os
import os.path
import copy

global cruzid_row
global assignmentName_column
cruzid_list = []
finalPoints_list = []
row_list = []
column_list = []
row_list_counter = 0
#==============================================================
# Get assignment from grader
lab_dir = raw_input("What assignment are you grading? (name must be contained in csv file)\n ")

#==============================================================
# creating 2 lists, students' cruzids and grades
for root, dirs, files in os.walk("./"+lab_dir+"/results", topdown=False):
	for name in files:
		# Make a list of all the cruzids
		file_name = name.split(".")
		cruzid = file_name[0]
		cruzid_list.append(cruzid)
		# Open the cruzid.txt file and read all lines of file.
		f=open('./'+lab_dir+'/results/'+name,"r")
		flines = f.readlines()
		f.close()
		#Read last line and extract percentage value
		finalLine = flines[-1]		
		#print cruzid, finalLine
		lineComponents = finalLine.split()
		#print lineComponents
		if not lineComponents:
			finalPoints = 0
		elif len(lineComponents) < 2:
			finalPoints = 0
		elif "/" in lineComponents[1]:
			pointComponents = lineComponents[1].split("/")
			finalPoints = pointComponents[0]
		else:
			finalPoints = 0
		finalPoints_list.append(finalPoints)
		#print cruzid, finalPoints
#print len(cruzid_list)
row_list = [0] * (len(cruzid_list))
#===========================================================

# scan through the csv file to get row number of cruzid 
#    and get column number of assignment (numbers each row and col)
with open("canvas_spreadsheet_to_import.csv","r") as canvasFile:
	canvasFileReader = csv.reader(canvasFile)				# create reader wrapped around an object.  These means use one time and done, so can't call twice.
	origCanvasFile = list(canvasFileReader)					# make a list of spread sheet, need a list to index to a specfic cell and overwrite old spread sheet
	updatedCanvasFile = copy.copy(origCanvasFile)			# need a copy for searching througn rows/cols and another to hang on to 
	for h in range(len(cruzid_list)):
		for i, row in enumerate(origCanvasFile):			#original values.  This copy allow the avoidance of another open.
			for field in row:  
				if cruzid_list[h] in field:
					cruzid_row = i
					row_list[h] = i
					print cruzid_list[h], row_list[h], finalPoints_list[h]
					#row_list_counter = row_list_counter + 1
			for j, column in enumerate(row):
				if lab_dir in column:
					lab_dir_column = j
					column_list.append(j)
					#print "column", j 
print len(row_list), len(column_list)
canvasFile.close()
# write grade to desired cell of list
for k in range(len(cruzid_list)):
	#print k, cruzid_list[k]
	if row_list[k] != 0:
		updatedCanvasFile[row_list[k]][lab_dir_column] = finalPoints_list[k]
		print cruzid_list[k], row_list[k], finalPoints_list[k]

#==============================================================
# write list to canvas spread sheet
# there doesn't seem to be any easy function within csv library to support writing to a specific cell.  
# There are other libraries but they look like 3rd party libraries, wasn't sure if that was a good idea.
with open("canvas_spreadsheet_to_import.csv","w") as canvasFile:
	canvasFileWriter = csv.writer(canvasFile)
	canvasFileWriter.writerows(updatedCanvasFile)
canvasFile.close()



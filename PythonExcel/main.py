import csv

'''*****************************************
STEP: 1
Reading all row into memory to process later
CSV File from: https://simplemaps.com/data/us-cities
*****************************************'''
#create list to store row object
city_list =[]
#will change to False and skip first row with title informaiton
skip_first_row = True

#Read CSV file into memory
with open('uscities.csv') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')
    for row in readCSV:
        #skip first row
        if skip_first_row:
            skip_first_row = False
            continue
        #store each row in a dictionary
        cityInfo ={
            'city' : row[0],
            'state_id': row[2],
            'state_name': row[3],
            'city_population':row[10],
            'military_base': row[13],
            'time_zone' : row[15]
        }

        #print(cityInfo)
        city_list.append(cityInfo)

#print(len(city_list))
#print(city_list[0]['city'])

#get list of all city id
all_state_id = []
for id in city_list:
    #check in city id is in list. If not then add to list
    if id['state_id'] not in all_state_id:
        all_state_id.append(id['state_id'])
#print(len(all_state_id))



'''**********************************************************
STEP: 2
check if current directory has a folder called "stateFiles"
If not then create one.
**********************************************************'''
import os
excel_folder_name = os.getcwd()+'\stateFiles'

if os.path.exists(excel_folder_name) == False or os.path.isdir(excel_folder_name) == False:

    try:
        os.mkdir(excel_folder_name)
    except OSError:
        print("Creation of the directory %s failed" % excel_folder_name)
    else:
        print("Successfully created the directory %s " % excel_folder_name)



'''**********************************************************
STEP: 3
    1. Use Openpyxl to create workbook for each state_id
    2. Create table in each workbook containing all information for that state
    3. Highlight rows for cities with Military Base
**********************************************************'''


from NewFile_FUNC import newFile

for i in all_state_id:
    newFile(i,city_list,excel_folder_name)
    #ONLY CREATE ONE FILE WHEN TESTING
    #break



__name__ = "__main__"


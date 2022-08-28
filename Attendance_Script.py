import pandas as pd
import os
import time
from os import path, system
import datetime

#console color
os.system('color 4F') # Colors in the CLI  
export_folder = ''
inputpath = ''
(Date,Time) = ('','')
#Root dirrectory 
'''
root = os.getcwd()
root1 = root.split('\\')
root1.pop()
rootpath = root1[0]+'/'+root1[1]
'''


#Function for path choice
def user(ch):
    global inputpath
    if ch == 'n':
        inputpath = input ('\nEnter the Absolute path of this folder:')
        os.chdir(inputpath)
    
    
    elif ch =='y':
        inputpath = rootpath+'/Attendance_Sheets'    #Default Input path for pickup!
        os.chdir(inputpath)
    else:
        print('Not a valid choice!')
        time.sleep(2)
        quit()
#time lag function
def Time(t):
    print('...')
    time.sleep(t)    

def export(choice1):
    global export_folder
    if choice1 == 'y':
        export_folder = rootpath+'/Output/Attendance_merge'         #Default Export Path
    elif choice1 == 'n':
        export_folder = input('\nEnter the absolute path of export folder')
    else:
        print('Not a valid choice!')
        Time(2)
        quit()
           
#initializing dataframe and path
Attendance = pd.DataFrame()
month = input('\nEnter the Abbreviation of the month for which u need the merged records:\neg:JAN,FEB,MAR...,NOV,DEC,etc:')
choice = input('\nIs the input directory, default one?\n[Attendance_Sheets] Folder (y/n):')
#os.system('cd ..')
rootpath = input('Enter the rootpath: ')
user(choice)

for file in os.listdir(os.getcwd()):
     if file.__contains__(month):
         df = pd.read_excel(file,dtype = 'str')
         Attendance = pd.concat([Attendance,df],axis = 1)
         
Attendance = Attendance.loc[:,~Attendance.columns.duplicated()]
Attendance = Attendance.replace(['a','A','p','P'],['Absent','Absent','Present','Present'])

choice1 = input('\nDo you want to export this file to the default location?\n[Output] Folder (y/n):')
export(choice1)
export_name = input('Enter name for this merge file')

Attendance.to_excel(export_folder+'/'+export_name + ".xlsx", index=False)
print('Your merged file has been generated! \n location: '+export_folder+' name: '+export_name)

choice2 = input('Do you want to merge this with the master data?\n[Master_Sheets] Folder (y/n):')

if choice2 == 'y':
        Master = pd.read_excel(rootpath+'/Master_Sheets/Master_Attendance.xlsx', dtype = 'str') #Master Attendance file!
        Master = Master.append(Attendance)
        Master.to_excel(rootpath+'/Master_Sheets/Master_Attendance.xlsx', index=False)
        Time(0.5)
        Time(0.5)
        Time(0.5)
        print('\n Your Master Data has been updated!')
        Time(0.5)
        Time(0.5)
        print('Writing entry into log file!')
        #log file - time date
        right_now = datetime.datetime.now()
        (Date,Time) = str(right_now).split()

        Folder_path = rootpath+'/Scripts'
        os.chdir(Folder_path)
        #making a log file to store job details
        time.sleep(1)
        file = open('logs.txt','a+')
        file.write("\nJob for Attendance_merge completed on date: "+Date+" at time: "+Time+"\n"+"Input Folder->"+inputpath+" Output Folder->"+export_folder+"\nAlso appended to Master_Attendace"+"\n________________________________________________________________________________________________")
        file.close()
         
else:
    print('Ok Bye..')
    Time(1.5)
    quit()
        

        
            



    
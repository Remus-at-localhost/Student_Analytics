import datetime
import pandas as pd
import os
from os import path
import time
#initialising the Root Dir
'''root = os.getcwd()
Root = root.split('\\')
Root_Path = Root[0]+'/'+Root[1]
'''
#initialising the dfs and variables
os.system('color 4F') # Colors in the CLI
ExamSheet = pd.DataFrame()
inputpath = ''
export_folder=''
export_name = ''
Attendance = pd.DataFrame()

# Function for time lag
def Time(t):
    print('...')
    time.sleep(t)

#Outputpath choice
def export(choice1):
    global export_folder
    if choice1 == 'y':    
        export_folder = Root_Path+'/Output/Exam_merge'         #Default Output Path
    elif choice1 == 'n':
        export_folder = input('\nEnter the absolute path of export folder')
    else:
        print('Not a valid choice!')
        Time(2)
        quit()    

#Inputpath choice
def user(ch):
    global inputpath
    if ch == 'n':
        inputpath = input ('\nEnter the Absolute path of this folder:')
        os.chdir(inputpath)
    elif ch =='y':
        inputpath = Root_Path+'/Exam_Sheets'                   #Default Input Path for pickup!
        os.chdir(inputpath)
    else:
        print('Not a valid choice!')
        Time(2)
        quit()
          
print('Result Sheets Merge')
Root_Path=input('Enter the root path:')
choice = input('\nIs the input directory for result records, default one?\n[Exam_Sheets] Folder (y/n):')
user(choice)
columns = ['Class','Subject','Exam']
choice0 = input('\nExamination Type: (Unit-1/Semester-1/Unit-2/Semester-2)')

#loading the required Data into the Dataframe from the selected path.
for file in os.listdir(os.getcwd()):
    if file.__contains__(choice0): 
        cols = file.split(' ')
        df = pd.read_excel(file,dtype='str')
        #cleaning [removing unessesary columns]
        df.drop(df.iloc[:,5:-1], inplace = True, axis = 1)
        df.pop(df.columns[-1])
        #making new rows for Class subject and exam type
        df[columns[0]] = cols[0]
        df[columns[1]] = cols[2]
        df[columns[2]] = cols[3]
        ExamSheet = ExamSheet.append(df)

#output of this dataframe
choice1 = input('\nDo you want to export this file to the default location?\n[Output] Folder (y/n):')
export(choice1)
export_name = input('Enter name for this merge file')
Attendance.to_excel(export_folder+'/'+export_name + ".xlsx", index=False)
print('Your merged file has been generated! \n location: '+export_folder+' name: '+export_name)        
#Merge with master data / a custom sheet
choice2 = input('\nDo you want to merge this to the default Master file? (y/n):')
if choice2 == 'y':
    #Export Path
    Master_folder = Root_Path+'/Master_Sheets'
    master = pd.read_excel(Master_folder+'/Master_Exam.xlsx',dtype = 'str') #Master file for Exam Result data
    master = master.append(ExamSheet)
    master.to_excel(Master_folder+'/Master_Exam.xlsx',index = False)
                    
elif choice2 == 'n':
    Master_folder = input('\nEnter the absolute path of export folder')
    name = input('\n Give a name to this file:')
    ExamSheet.to_excel(Master_folder+'/'+name+'.xlsx', index = False)
    
else:
    print('Not a valid choice!')
    Time(2)
    quit()

print('\n...')
Time(0.5)
print('\n...')
Time(0.5)
print('Your result merge is generated..')
print('\n...')
Time(0.5)
print('Writing entry into log file!')
Time(1)
#log file - time date
right_now = datetime.datetime.now()
(Date,Time) = str(right_now).split()

Folder_path = Root_Path+'/Scripts'
os.chdir(Folder_path)
#making a log file to store job details
file = open('logs.txt','a+')
file.write("\nJob for Exam_merge completed on date: "+Date+" at time: "+Time+"\n"+"Input Folder->"+inputpath+" Output Folder->"+export_folder+"\nAlso appended to Master_Exam"+"\n________________________________________________________________________________________________")
file.close()        

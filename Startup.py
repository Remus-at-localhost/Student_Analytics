#Making the proper file system and importing the required libraries..
import os
import time
import subprocess
import shutil as sh
import xlsxwriter

os.system('color 4F')
Extracted_at = os.getcwd() 
#importing the required libraries.
libraries = ['pandas','openpyxl','xlsxwriter']

for library in libraries:
    print('Installing ',library)
    subprocess.check_call(['pip','install',library])
    time.sleep(1.5)
print('\nAll the required Libraries are imported to the System..')
PATH = input('Enter the dirrectory path where you want to implement the suggested file structure: ')
Base_Folder = input('Enter the name of the main folder: ')
PATH2 = PATH+'/'+Base_Folder
#list of suggested folders
Folders = ['Attendance_Sheets','Exam_Sheets','Master_Sheets','Dashboard','Output','Scripts']
output_list = ['Attendance_Merge','Exam_Merge']
Master_sheets = ['Master_Attendance.xlsx','Master_Exam.xlsx']
#Scripts = ['Exam_Merge_Script.py','Attendance_Script.py']
#Creating these folders..
os.chdir(PATH)
os.mkdir(PATH2)
os.chdir(PATH2)

#Creator Function
def Creator(LIST):
    for x in LIST:
        print(x)
        os.mkdir(x)

Creator(Folders)
os.chdir(PATH2 +'/'+'Output')
Creator(output_list)
os.chdir(PATH2 +'/'+'Master_Sheets')
for sheets in Master_sheets:
    workbook = xlsxwriter.Workbook(sheets)
    workbook.close() 

    
#os.chdir(PATH2 +'/'+'Scripts')
for file in os.listdir(Extracted_at):
    Source = Extracted_at + '/' + file
    Destination = PATH2 +'/'+'Scripts'+'/'+file
    sh.move(Source,Destination)
    


# for folder in Folders:
#     print(folder)
#     os.mkdir(PATH2+'/'+folder)

# os.chdir(PATH2 +'/'+'Output')
# for output in output_list:
#     os.mkdir(output)

# os.chdir(PATH2 +'/'+'Output')
# for output in output_list:
#     os.mkdir(output)

time.sleep(1)    
print('The required File system is ready..\nLocation->'+str(PATH2))
exit = input('Press any key to exit..')



    
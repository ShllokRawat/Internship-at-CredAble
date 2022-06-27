import zipfile
import os
import re
import shutil
#Code to find the path of a file in a folder
#stores the first path in a variable called 'paths'
#use regex to extract only the name of the zip file(ensure that .zip is not included) and store it as folder_name

#THIS IS THE CODE TO UNZIP THE FOLDER INTO ANOTHER FOLDER
folder = folder_name
with zipfile.ZipFile(paths, 'r') as zip_ref:
    zip_ref.extractall(folder)
#THIS IS THE CODE TO UNZIP THE FOLDER INTO ANOTHER FOLDER
    
#use regex to remove the .zip from the 'paths' variable, and store it as folder_path

#THIS IS THE CODE TO MOVE THE UNZIPPED FOLDER INTO A PROCESSING FOLDER
os.chdir('C:\\')
os.system('mkdir Processing')
shutil.move(folder_path , 'C:\\Users\\shllo\\Desktop\\Internship\\Data\\Processing')
#THIS IS THE CODE TO MOVE THE UNZIPPED FOLDER INTO A PROCESSING FOLDER

#use regex to remove everything before the '_' in the folder_name variable, and store it as file_name

path1 = str('C:\\Users\\shllo\\Desktop\\Internship\\Data\\Processing' + folder_name + file_name)
path2 = str('C:\\Users\\shllo\\Desktop\\Internship\\Data\\Processing' + folder_name)
#EXTRACTION CODE GOES HERE
os.chdir('C:\\')
os.system('mkdir Processed')
shutil.move( path2, 'C:\\Users\\shllo\\Desktop\\Internship\\Data\\Processed')


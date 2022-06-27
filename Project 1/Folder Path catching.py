import os
import zipfile
import re
import shutil
FOLDER_PATH = r'C:\Users\shllo\Desktop\Internship\Data\BSA Reports - Unzipped'

def listDir(dir):
    fileNames = os.listdir(dir)
    for fileName in fileNames:
            folder_name = str(os.path.abspath(os.path.join(dir, fileName)))
            folder_path = folder_name
            #print(fileName)
            firstPass = re.findall("_.*$",fileName)[0]
            excelsheet_name = re.sub("_","",firstPass)
            print(excelsheet_name)
            #print(fileName)
            #path1 = str(folder_path + '\\' + excelsheet_name +'.xlsx')
            #print(path1)
            #os.chdir('C:\\')
            #os.system('mkdir Processed')
            #shutil.move( path2, 'C:\\Users\\shllo\\Desktop\\Internship\\Data\\Processed')   

if __name__ == '__main__':
    listDir(FOLDER_PATH)


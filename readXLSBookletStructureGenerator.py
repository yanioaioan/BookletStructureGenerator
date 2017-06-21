#! /usr/bin/bash
#title           :readXLSBookletStructureGenerator
#description     :This script is a booklet folder structure generator, used for undergraduate and postgraduate degree showreel booklets
#author		 :Ioannis Ioannidis
#date            :21/06/2017
#version         :1.0
#usage           :./readXLSBookletStructureGenerator
#                : you need a .xlsl file to read which is provided from the user
#date modified   :--/--/----
#==============================================================================


#How to use the `readXLSBookletStructureGenerator script`:
#
#1)open bash profile: geany ~/.bashrc
#2)append this line to the end: export PATH=$PATH:/public/bin/yanScripts
#3)Save and Close bash
#4)Source bash: source ~/.bashrc
#5)make sure there's an excel file to read from (usually this is retrieved from Google Drive after students have submitted their information)
#6)run the following command:  `./readXLSBookletStructureGenerator`


import pandas,os
import shutil


def isNaN(num):
    return num != num

def assure_path_exists(path):
        dir = os.path.dirname(path)
        if not os.path.exists(dir):#if not there - create it
                os.makedirs(dir)
                return True
        if os.path.exists(path):#if there - delete first and then create it
            print "PATH already exists: %s "%(path)
            shutil.rmtree(path)
            print "PATH deleted so as to be recreated: %s "%(dir)
            #os.makedirs(path)

            print path
            print dir



excelFile='excel/DegreeShow2017.xlsx'
print "Please Give the absolute path to the excel file: such as '/home/yioannidis/Downloads/BookletStructureGenerator/excel/DegreeShow2017.xlsx'"
read excelFile
df = pandas.read_excel(open(excelFile,'rb'), sheetname='Sheet1')
#print the column names
#print df.columns
cols=df.columns

sumbissions=[]
timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software = "","","","","","","","","","",""

'''
for j in df.index:#get the values for a given
    #print "ROW",j
    for i in cols:#get the values for a given column
        #print "COLUMN:",i,df[i][j]
        sumbissions.append(df[i][j])
'''

#get rows number
#print df.count

'''
row=df.loc[0][:]#1st row all elements EXCLUDING numbering on the left hand side
print len(row)
print row
print row[0]
print row[10]
'''

#add all row into the list submissions
for i in df.index:
  row=df.loc[i][:]
  sumbissions.append(row)
  #print i

#create outter folder named "studentsFolders"
cwd = os.getcwd()
outterStudentsFolder=cwd+"/studentsFolder"
assure_path_exists(outterStudentsFolder)

for row in sumbissions:#each row
    #for rowElement in row:#individual elements of each row
    timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software = row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10]
    print timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software
    #name,inumber,email,phone,url,affiliation,projectname,description,skills,software = "","","","","","","","","",""

    #now for this row create a folder in the following format ("name - projectname")
    studentFoldername="/%r-%r/"%(name,projectname)
    studentFoldername=str(studentFoldername)
    print studentFoldername
    #cwd = os.getcwd()
    studentFoldername=outterStudentsFolder+studentFoldername
    print studentFoldername
    studnetFolderCreatedSucessfully=assure_path_exists(studentFoldername)


    #then for this foldername create a description txt file of the following format ("name - inumber")
    if studnetFolderCreatedSucessfully:

      os.chdir(studentFoldername)
      cwd = os.getcwd()
      print cwd

      studentFileDescriptionName="%s-%s.txt"%(name,inumber)

      file = open(studentFileDescriptionName,"w")


      if not isNaN(name) and not isNaN(inumber):
        file.write("%r\t%r\n\n"%(name,inumber))
      if not isNaN(email):
        file.write("%r\n"%(email))
      if not isNaN(phone):
        file.write("%r\n"%(phone))
      if not isNaN(url):
        file.write("%r\n\n"%(url))
      if not isNaN(affiliation):
        file.write("%r\n"%(affiliation))
      if not isNaN(projectname):
        file.write("%r\n\n"%(projectname))
      if not isNaN(description):
        file.write("Synopsis:\n%r\n\n"%(description))
      if not isNaN(skills):
        file.write("%r\n"%(skills))
      if not isNaN(software):
        file.write("%r\n"%(software))
      file.close()












#for s in sumbissions:#list holding all submission
#  print s
#  break



#values = df['Arm_id'].values
#get a data frame with selected columns
#FORMAT = ['Arm_id', 'DSPName', 'Pincode']
#df_selected = df[FORMAT]

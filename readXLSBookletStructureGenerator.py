import pandas,os
import shutil

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



def isNaN(num):
    return num != num

def assure_path_exists(path):
        dir = os.path.dirname(path)
        print dir
        if not os.path.exists(path):#if not there - create it
                os.makedirs(path)
                return True
        if os.path.exists(path):#if there - delete first and then create it
            print "PATH already exists: %s "%(path)
            shutil.rmtree(path)
            print "PATH deleted so as to be recreated: %s "%(dir)
            #os.makedirs(path)

            print path
            print dir



excelFile='excel/DegreeShow2017.xlsx'

excelFile=raw_input("Please Give the absolute path to the excel file: such as '/home/yioannidis/Downloads/BookletStructureGenerator/excel/DegreeShow2017.xlsx'")
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

groupsStudentsFolder=outterStudentsFolder+"/_____GROUPS_____"
assure_path_exists(groupsStudentsFolder)
print groupsStudentsFolder

individualsStudentsFolder=outterStudentsFolder+"/_____INDIVIDUALS_____"
assure_path_exists(individualsStudentsFolder)
print individualsStudentsFolder


import unicodedata

for row in sumbissions:#each row
    #for rowElement in row:#individual elements of each row
    timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software = row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10]


    #encoding
    timestamp = ((timestamp)).encode('utf-8')
    if not isNaN(name):
      name = ((name)).encode('utf-8')
    if not isNaN(inumber):
        inumber = ((inumber)).encode('utf-8')
    if not isNaN(email):
        email = ((email)).encode('utf-8')
    if not isNaN(phone):
        phone = (str(phone)).encode('utf-8')
    if not isNaN(url):
        url = (str(url)).encode('utf-8')
    if not isNaN(affiliation):
        affiliation = ((affiliation)).encode('utf-8')
    if not isNaN(projectname):
        projectname = ((projectname)).encode('utf-8')
    if not isNaN(description):
        #description = (str(description)).encode('utf-8')

        print "DESCRIPTION=",description

        #It's important to notice that using the ignore option is dangerous because it silently drops any unicode(and internationalization) support from the code that uses it, as seen here:
        description=unicodedata.normalize('NFKD', description).encode('ascii','ignore')

    if not isNaN(skills):
        skills = (str(skills)).encode('utf-8')
    if not isNaN(software):
        software = ((software)).encode('utf-8')

    print timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software
    #name,inumber,email,phone,url,affiliation,projectname,description,skills,software = "","","","","","","","","",""

    #now for this row create a folder in the following format ("name - projectname") if it's an individual project otherwise,
    #create a folder in the following format ("projectname") and under this one create a folder in the following format ("name") for each of the members of the group
    studentFoldername=""
    #cwd = os.getcwd()

    #dive into either Groups or Individual Folder depending on the specified affiliation
    affiliationFolderChosen=""
    if not isNaN(affiliation):
        affiliation = ((affiliation)).encode('utf-8')

        if affiliation == "group":
          os.chdir(groupsStudentsFolder)
          affiliationFolderChosen = os.getcwd()

          #list individual folders each group AND in there.. for each student
          groupProjectFoldername="/%s/"%(projectname)
          studentFoldername="%s/"%(name)
          studentFoldername=str(groupProjectFoldername+studentFoldername)

          print studentFoldername


        elif affiliation == "individual":

          #list individual folders for each student
          studentFoldername="/%s-%s/"%(name,projectname)
          studentFoldername=str(studentFoldername)
          print studentFoldername

          os.chdir(individualsStudentsFolder)
          affiliationFolderChosen = os.getcwd()

    print affiliationFolderChosen+studentFoldername

    #update studentFoldername
    studentFoldername=affiliationFolderChosen+studentFoldername#outterStudentsFolder+studentFoldername
    print studentFoldername

    studnetFolderCreatedSucessfully=assure_path_exists(studentFoldername)#studentFoldername

    #then for this foldername create a description txt file of the following format ("name - inumber")
    if studnetFolderCreatedSucessfully:


      #dive into studentFoldername
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
        file.write("%r\n\n"%(phone))
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

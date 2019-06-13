try:
	import pandas,os,re
except:
	print 'The following python libraries need to be installed, so if you have admin rights under linux \n \
	open a terminal and execute the following commands:'
	print 'sudo easy_install pip'
	print 'sudo pip install pandas'	
	print 'sudo pip install xlrd'
	exit(0)

import shutil

#! /usr/bin/bash
#title           :readXLSBookletStructureGenerator
#description     :This script is a booklet folder structure generator, used for undergraduate and postgraduate degree showreel booklets
#author		 :Ioannis Ioannidis
#date            :21/06/2017
#version         :1.0
#usage           :python readXLSBookletStructureGenerator
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




excelFile=raw_input("Please Give the absolute path to the excel file: such as '/public/bin/yanScripts/excel/DegreeShow2017.xlsx'")
photosPath=raw_input("Please Give the absolute path to the photos: such as '/public/bin/yanScripts/photos'")

#excelFile="/public/bin/yanScripts/excel/DegreeShow2017.xlsx"
#photosPath="/public/bin/yanScripts/photos"

#excelFile="/home/yioannidis/Downloads/BookletStructureGenerator/2018/readXLSBookletStructureGenerator2018Test/excel/Booklet2018-Downloaded.xlsx"
#photosPath="/home/yioannidis/Downloads/BookletStructureGenerator/2018/readXLSBookletStructureGenerator2018Test/photos/Work_Collector"

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
outterStudentsFolder=os.path.join(cwd,"studentsFolder")
assure_path_exists(outterStudentsFolder)

groupsStudentsFolder=os.path.join(outterStudentsFolder,"_____GROUPS_____")
assure_path_exists(groupsStudentsFolder)
print groupsStudentsFolder


def recursive_glob(treeroot, pattern):
    results = []
    for base, dirs, files in os.walk(treeroot):
        goodfiles = fnmatch.filter(files, pattern)
        results.extend(os.path.join(base, f) for f in goodfiles)
    return results

individualsStudentsFolder=os.path.join(outterStudentsFolder,"_____INDIVIDUALS_____")
assure_path_exists(individualsStudentsFolder)
print individualsStudentsFolder


import unicodedata

#Collect inumbers to check whether there are inumbers in the project photos folder submitted,
#that don't correspond to the submission inumbers on the booklet information excel sheet
#all inumbers submtitted in the excel
excellInumbers=[]
for row in sumbissions:
    timestamp,name,inumber,email,phone,url,affiliation,projectname,description,skills,software = row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10]
    excellInumbers.append(((inumber)).encode('utf-8'))
print  'excellInumbers',excellInumbers

t=[f for f in os.listdir(photosPath) if re.match(r'.*'+str('i')+'.*_.*', f)]
inumbersDetectedInPhotosPath = []
for i in t:
    print i
    inumbersDetectedInPhotosPath.append(i.split('_')[0])

print '\ninumbersDetectedInPhotosPath\n',inumbersDetectedInPhotosPath

suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo=[]
#detect the inumbers that have been correctly submitted as part of the name of the project photos BUT NOT correctly submitted as part of the booklet info submission
for inum in inumbersDetectedInPhotosPath:
    if (inum not in excellInumbers) and (inum not in suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo):#i number not detected in the booklet sumbission excel sheet AND not already in suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo list, then add it for further investigation
        suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo.append(inum)

groupMembersThaHaveNotSubmittedAnyProjectFilesThemselves=[]
#detect the inumbers that have NOT been correctly submitted as part of the name of the project photos BUT correctly submitted as part of the booklet info submission
for inum in excellInumbers:
    if (inum not in inumbersDetectedInPhotosPath) and (inum not in groupMembersThaHaveNotSubmittedAnyProjectFilesThemselves):#i number not detected in the booklet sumbission excel sheet AND not already in suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo list, then add it for further investigation
        print inum,(inum not in inumbersDetectedInPhotosPath)
        groupMembersThaHaveNotSubmittedAnyProjectFilesThemselves.append(inum)


#if not empty
if suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo:
  print '\nAttention possible WRONG inumbers sumbitted\n as part of the excel booklet info that don\'t match the project photos inumbers.\nPlease, investigate further inumbers sumbitted'
  print 'Please investigate submissions of the following inumbers\n\n'

  print '----------------------------------------------------------------------------------------------'
  print '----------------------------------------------------------------------------------------------'
  print '     INUMBER ERRORS IN BOOKLET SUBMISSION - INVESTIGATE THE FOLLOWING \'%d\' INUMBER SUBMISSIONS'%(len(suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo))
  print '----------------------------------------------------------------------------------------------'
  print '----------------------------------------------------------------------------------------------'
  counter = 1
  for i in suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo:
      print counter,')  !!!!!-Investigate booklet submission of -->',i,'-!!!!!\n'
      counter+=1
  print '\nAlso check groupMembersThaHaveNotSubmittedAnyProjectFilesThemselves:\n'
  print 'double check booklet submissions of the follwoing inumbers too are indeed members of a groups that haven\'t submitted project photos because another group member did!'
  print groupMembersThaHaveNotSubmittedAnyProjectFilesThemselves


  file = open(outterStudentsFolder+str('/ATTENTION.txt'),"w")
  for i in suspiciousInumberMisMatchBetweenProjectPhotosAndBookletInfo:
    file.write('!!!!!-Investigate booklet submission of -->%s-!!!!!, inumber possibly mistyped\n'%(i))
  file.close()


'''
for inumber in excellInumbers:
  print 'Testing inumber-->..',str(inumber)
  t=[f for f in os.listdir(photosPath) if re.match(r'.*'+str(inumber)+'.*', f)]
  print 'matched',t
exit()
'''

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

    print timestamp
    print name 
    print inumber
    print email
    print phone
    print url
    print affiliation
    print projectname
    print description
    print skills
    print software
    #name,inumber,email,phone,url,affiliation,projectname,description,skills,software = "","","","","","","","","",""

    #now for this row create a folder in the following format ("name - projectname") if it's an individual project otherwise,
    #create a folder in the following format ("projectname") and under this one create a folder in the following format ("name") for each of the members of the group
    studentFoldername=""
    #cwd = os.getcwd()

    #dive into either Groups or Individual Folder depending on the specified affiliation
    affiliationFolderChosen=""
    if affiliation:
        affiliation = ((affiliation)).encode('utf-8')

        if affiliation == "As part of a group":
          os.chdir(groupsStudentsFolder)
          print 'group affiliation'
          affiliationFolderChosen = os.getcwd()

          #list individual folders each group AND in there.. for each student
          groupProjectFoldername="%s"%(projectname)
          studentFoldername="%s"%(name)
          studentFoldername=os.path.join(groupProjectFoldername,studentFoldername)

          print studentFoldername


        elif affiliation == "As an individual":

          #list individual folders for each student
          print 'individual affiliation'
          studentFoldername="%s-%s"%(name,projectname)
          studentFoldername=str(studentFoldername)
          print studentFoldername

          os.chdir(individualsStudentsFolder)
          affiliationFolderChosen = os.getcwd()

    print os.path.join(affiliationFolderChosen,studentFoldername)

    #update studentFoldername
    studentFoldername=os.path.join(affiliationFolderChosen,studentFoldername)#outterStudentsFolder+studentFoldername
    print "studentFoldername %r"%(studentFoldername)

    studnetFolderCreatedSucessfully=assure_path_exists(studentFoldername)#studentFoldername

    #then for this foldername create a description txt file of the following format ("name - inumber")
    if studnetFolderCreatedSucessfully:



      #dive into studentFoldername
      os.chdir(studentFoldername)
      cwd = os.getcwd()
      print cwd

      #create local images under each person
      localImages=os.path.join(studentFoldername,"images")
      assure_path_exists(localImages)
      #print localImages

      #search, find & copy inumber-related image to pre-created local folder named 'images'
      filesMatched = [f for f in os.listdir(photosPath) if re.match(r'.*'+str(inumber)+'.*', f)]
      print filesMatched

      for file in filesMatched:

            #print localImages
            imagepath=dir = os.path.join(os.path.abspath(photosPath),file)

            print imagepath
            shutil.copy2(imagepath, localImages)


      studentFileDescriptionName="%s-%s.txt"%(name,inumber)

      file = open(studentFileDescriptionName,"w")


      if not isNaN(name) and not isNaN(inumber):
        file.write("name: %r\t%r\n\n"%(name,inumber))
      if not isNaN(email):
        file.write("email: %r\n"%(email))
      if not isNaN(phone):
        file.write("phone: %r\n\n"%(phone))
      if not isNaN(url):
        file.write("url: %r\n\n"%(url))
      if not isNaN(affiliation):
        file.write("affiliation: %r\n"%(affiliation))
      if not isNaN(projectname):
        file.write("projectname: %r\n\n"%(projectname))
      if not isNaN(description):
        file.write("description: \n%r\n\n"%(description))
      if not isNaN(skills):
        file.write("skills :%r\n"%(skills))
      if not isNaN(software):
        file.write("software: %r\n"%(software))
      file.close()


#Check folders not having images
import os,glob

empty_dirs = []
for root, dirs, files in os.walk(outterStudentsFolder):
   newpath = ''
   noImagesAtAllInProject = 1

   if not len(dirs) and not len(files):
      print '\nroot',root
      groupProjectPath=root.split('/')
      newpath = ''
      for i in range(len(groupProjectPath)-2):
          newpath += str(groupProjectPath[i])+'/'
      print '\nnewpath',newpath


      noImagesAtAllInProject=1
      import fnmatch
      matches = []

      matchesjpg = recursive_glob(newpath,'*.jpg')
      matchespng = recursive_glob(newpath,'*.png')
      matchesJPG = recursive_glob(newpath,'*.JPG')
      matchesPNG = recursive_glob(newpath,'*.PNG')
      matchesjpeg = recursive_glob(newpath,'*.jpeg')
      matchesJPEG = recursive_glob(newpath,'*.JPEG')

      if matchesjpg:
          noImagesAtAllInProject=0
          print i

      elif matchespng:
          noImagesAtAllInProject=0
          print i

      elif matchesJPG:
          noImagesAtAllInProject=0
          print i

      elif matchesPNG:
          noImagesAtAllInProject=0
          print i
      elif matchesjpeg:
          noImagesAtAllInProject=0
          print i

      elif matchesJPEG:
          noImagesAtAllInProject=0
          print i

      else:
          print '\nNo images in this project at all -->%s\n'%(newpath)
          file = open(outterStudentsFolder+str('/ATTENTION.txt'),"a")
          file.write('\nInvestigate Project %s, as there are no images at all there (ex. project title mistyping or anything really / Email the corresponding students \n'%(newpath))
          file.close()


#for s in sumbissions:#list holding all submission
#  print s
#  break



#values = df['Arm_id'].values
#get a data frame with selected columns
#FORMAT = ['Arm_id', 'DSPName', 'Pincode']
#df_selected = df[FORMAT]

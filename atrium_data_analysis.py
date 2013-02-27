##Atrium Photo analysis
import os
#from datetime import datetime
from xlwt.Workbook import *

#gathers info from photos across a full day
class Day:
    def __init__(self, day):
        self.day = day
        self.photoList = []
        
    def addPhoto(self, photo):
        self.photoList.append(photo)

#stores info of a single photo      
class Photo:
    def __init__(self, exactDate, dictList):
        month = ["", "Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]
        self.exactDate = str(exactDate[0]) + " " + month[int(exactDate[1])] + " " + str(exactDate[2]) + " " + str(exactDate[3]) + ":" + str(exactDate[4])
        self.dictList = dictList
        self.numPeople = 0
        self.numGroups = 0
        self.dayDate = self.exactDate[:3]
        self.personList = []
        self.groupList = []
    
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople += 1
    
    def addGroup(self, group):
        self.groupList.append(group)
        self.numGroups += 1

    def __str__(self):
         return self.exactDate.toString()       

#knows number of people in it, and which people     
class Group:
    def __init__(self, dicty):
        self.minX = dicty["x"]
        self.minY = dicty["y"]
        self.maxX = dicty["x"] + dicty["width"]
        self.maxY = dicty["y"] + dicty["height"]
        self.numPeople = 0
        self.personList = []
        
    def addPerson(self, person):
        self.personList.append(person)
        self.numPeople += 1
        
    def isPersoninGroup(self, person):
        if person.midX >= self.minX and person.midX <= maxX and person.midY >= self.minY and person.midY <= self.maxY:
            return true
        else:
            return false

#knows center point of person
class Person:
    def __init__(self, dicty):
        self.midX = float(dicty["x"]) + float(dicty["width"]) / 2.0
        self.midY = float(dicty["y"]) + float(dicty["height"]) / 2.0
        
    

# input: directory path(string)
# goes through all the files in a directory and calls parseFile on them
# joins the outputs of all the parseFile calls and returns it
def readFiles(path):
    allDictList = []
    for file in os.listdir(path):   
        print "reading" + str(file)
        for dicty in parseFile(path + "\\" + file):
            allDictList.append(dicty)
    
    masterDict = {}
    for dicty in allDictList:
        date = (int(dicty["year"]), int(dicty["month"]), int(dicty["day"]), int(dicty["hour"]), int(dicty["minute"]))
        if date in masterDict:
            masterDict[date].append(dicty)
        else:
            masterDict[date] = [dicty]
    
    keyList = []
    for key in masterDict:
        keyList.append(key)
    
    keyList.sort()
    
    allPhotoList = []
    i = 0
    while i < len(keyList):
        photo = Photo(keyList[i], masterDict[keyList[i]])
        allPhotoList.append(photo)
        analyzePhoto(photo)     
        print "analyzing photo from " + str(photo.exactDate)
        i += 1
            
    peoplePerPhoto(allPhotoList)
    groupsPerPhoto(allPhotoList)

    
# input: a filename (string)
# opens the file, goes through all the lines and extracts the fields
# returns an unique list of dictionaries (one for each line)    
def parseFile(filePath):
    f = open(filePath) 
    linesList = f.readlines() 
    f.close()
    
    noDuplicates = list(set(linesList))
    
    fileObjectDicts = []
    for line in noDuplicates:
        fileObjectDicts.append(tokenize(line))      
    
   # dict = fileObjectDicts[0]
   # exactDate = (dict["year"], dict["month"], dict["day"], dict["hour"], dict["minute"])
   # photo = Photo(fileObjectDicts, exactDate)
        
    return fileObjectDicts

    
separators = [',', '_', '.']
# Converts string into a list of tokens (strings)
# save tokens in a dictionary, and returns it
def tokenize(string):
    tokensList = []  
    token = ''    #keeps track of a single element
    for x in string:
        if x in separators:
            if token != '':  #add nonempty token when a separator is reached, then reset token
                tokensList.append(token) 
                token = ''
        else:   
            token += x
    
    objectDict = {'ID':tokensList[0], 'x':tokensList[1], 'y':tokensList[2], 'width':tokensList[3], \
        'height':tokensList[4], 'year':tokensList[5], 'month':tokensList[6], 'day':tokensList[7], \
        'hour':tokensList[8],   'minute':tokensList[9], 'second':tokensList[10]}
                    
    return objectDict


def analyzePhoto(photo):
    for dicty in photo.dictList:
        if dicty["ID"] == '1':
            photo.addPerson(Person(dicty))
        elif dicty["ID"] == "2":
            photo.addGroup(Group(dicty))
                        

def peoplePerPhoto(photoList):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numPeople))
    
    workbook = Workbook()
    sheet = workbook.add_sheet("peoplePerPhoto")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of people")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\peoplePerPhoto.xls")

def groupsPerPhoto(photoList):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numGroups))
    
    workbook = Workbook()
    sheet = workbook.add_sheet("groupsPerPhoto")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of groups")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\groupsPerPhoto.xls")
    
# this is a generic function to help with calculating statistics on the dataset
# specific functions could be written instead
# input: data: list of dictionaries, fun: a lambda function that takes in a dictionary and outputs a number
# returns: a list of numbers that were obtained by applying fun on each element in data
def calculateStat(data, fun):
    
    pass

    
# generates an Excel spreadsheet
# input: a list of lists or dictionaries to put in the spreadsheet
# each first-level list is a sheet in Excel
# example usage: outputToExcel(calculateStat(allData, *function that calculates average of X*), calculateStat(allData, *function that calculates average of Y*))
def outputToExel(data):
    pass
    


readFiles("C:\Users\Nicole\Documents\UROP 2013\data")

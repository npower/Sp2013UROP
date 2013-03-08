import os
#from datetime import datetime
from xlwt.Workbook import *
from atrium_objects import *       
    

# input: directory path(string)
# goes through all the files in a directory and calls parseFile on them
# joins the outputs of all the parseFile calls and returns it
def readFiles(path):
    allDictList = []
    for file in os.listdir(path):   
        #print "reading" + str(file)
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
        #print "analyzing photo from " + str(photo.exactDate)
        i += 1
            
    peoplePerPhoto(allPhotoList)
    groupsPerPhoto(allPhotoList)
    peoplePerDay(allPhotoList)
    groupsPerDay(allPhotoList)
    averagePeoplePerGroupPerPhoto(allPhotoList)
    peopleUsingSmallChairsPerPhoto(allPhotoList)
    peopleUsingSofasPerPhoto(allPhotoList)

    
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
        
    return fileObjectDicts

    
separators = [',', '_', '.']
# Converts string into a list of tokens (strings)
# save tokens in a dictionary, and returns it
def tokenize(string):
    tokensList = []  
    token = ''    #keeps track of a single elt
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
    #determine what objects are in the photo
    for dicty in photo.dictList:
        if dicty["ID"] == '1':
            photo.addPerson(Person(dicty))
        elif dicty["ID"] == "2":
            photo.addGroup(Group(dicty))
        elif dicty["ID"] == "3":
            photo.addChair(Chair(dicty))
        elif dicty["ID"] == "4":
            photo.addSofa(Sofa(dicty))
        elif dicty["ID"] == "5":
            photo.addSmallTable(Small_Table(dicty))
        elif dicty["ID"] == "6":
            photo.addLargeTable(Large_Table(dicty))
##        elif dicty["ID"] == "7":
##            photo.addPingPongTable(Ping_Pong_Table(dicty))
    #determine which group a person is in, and what furniture theyre using
    for person in photo.personList:
        for group in photo.groupList:
            if group.isInGroup(person):
                group.addPerson(person)
        for chair in photo.chairList:
            if chair.usingChair(person):
                chair.addPerson(person)
        for sofa in photo.sofaList:
            if sofa.usingSofa(person):
                sofa.addPerson(person)
        for smallTable in photo.smallTableList:
            if smallTable.usingSmallTable(person):
                smallTable.addPerson(person)
        for largeTable in photo.largeTableList:
            if largeTable.usingLargeTable(person):
                largeTable.addPerson(person)
    #determine which furniture group is using
    for group in photo.groupList:
        for chair in photo.chairList:
            if chair.usingChair(group):
                chair.addGroup(group)
        for sofa in photo.sofaList:
            if sofa.usingSofa(group):
                sofa.addGroup(group)
        for smallTable in photo.smallTableList:
            if smallTable.usingSmallTable(group):
                smallTable.addGroup(group)
    for largeTable in photo.largeTableList:
        for group in photo.groupList:
            if group.isInGroup(largeTable):
                largeTable.addGroup(group)
            
                               
                        

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

def peoplePerDay(photoList):
    excelList = []
    currentDate = photoList[0].dayDateString
    currentNumPeople = 0
    excelList.append((currentDate, currentNumPeople))
    i = 0
    for photo in photoList:
        if photo.dayDateString == currentDate:
            currentNumPeople += photo.numPeople 
            excelList[i] = (currentDate, currentNumPeople)
        else:
            currentDate = photo.dayDateString
            currentNumPeople = photo.numPeople
            excelList.append((currentDate, currentNumPeople))
            i += 1
            
    workbook = Workbook()
    sheet = workbook.add_sheet("peoplePerPDay")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of people")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\peoplePerDay.xls")

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

def groupsPerDay(photoList):
    excelList = []
    currentDate = photoList[0].dayDateString
    currentNumGroups = 0
    excelList.append((currentDate, currentNumGroups))
    i = 0
    for photo in photoList:
        if photo.dayDateString == currentDate:
            currentNumGroups += photo.numGroups 
            excelList[i] = (currentDate, currentNumGroups)
        else:
            currentDate = photo.dayDateString
            currentNumGroups = photo.numGroups
            excelList.append((currentDate, currentNumGroups))
            i += 1
            
    workbook = Workbook()
    sheet = workbook.add_sheet("groupsPerPDay")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of groups")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\groupsPerDay.xls")

def averagePeoplePerGroupPerPhoto(photoList):
    excelList = []
    for photo in photoList:
        totalPeopleinGroups = 0
        for group in photo.groupList:
            totalPeopleinGroups += group.numPeople
        averagePeoplePerGroup = 0
        if photo.numGroups != 0:
            averagePeoplePerGroup = float(totalPeopleinGroups)/photo.numGroups
        excelList.append((photo.exactDate, averagePeoplePerGroup))
    
    workbook = Workbook()
    sheet = workbook.add_sheet("peoplePerGroup")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "average people per group")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\AveragePeoplePerGroupPerPhoto.xls")

def peopleUsingSmallChairsPerPhoto(photoList):
    excelList = []
    for photo in photoList:
        totalPeopleinChairs = 0
        for chair in photo.chairList:
            totalPeopleinChairs += chair.numPeople
        excelList.append((photo.exactDate, totalPeopleinChairs))
    
    workbook = Workbook()
    sheet = workbook.add_sheet("peopleUsingChairs")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of people using chairs")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\peopleUsingChairsPerPhoto.xls")

def peopleUsingSofasPerPhoto(photoList):
    excelList = []
    for photo in photoList:
        totalPeopleinSofas = 0
        for sofa in photo.sofaList:
            totalPeopleinSofas += sofa.numPeople
        excelList.append((photo.exactDate, totalPeopleinSofas))
    
    workbook = Workbook()
    sheet = workbook.add_sheet("peopleUsingChairs")
    sheet.write(0, 0, "date")
    sheet.write(0, 1, "number of people using sofaa")
    
    r = 1
    for item in excelList:
        sheet.write(r, 0, item[0])
        sheet.write(r, 1, item[1])
        r += 1
    
    workbook.save("C:\Users\Nicole\Documents\UROP 2013\peopleUsingSofasPerPhoto.xls")


readFiles("C:\Users\Nicole\Documents\UROP 2013\data")

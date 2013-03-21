import os
from datetime import *
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
        date = datetime(int(dicty["year"]), int(dicty["month"]), int(dicty["day"]), int(dicty["hour"]), int(dicty["minute"]))
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
    emptyDict = {}
    while i < len(keyList):
        if keyList[i] in masterDict:
            photo = Photo(keyList[i], masterDict[keyList[i]])
        else:
            photo = Photo(keyList[i], emptyDict)
        allPhotoList.append(photo)
        analyzePhoto(photo)     
        #print "analyzing photo from " + str(photo.exactDate)
        i += 1
    
    workbook = Workbook()
            
    peoplePerPhoto(allPhotoList, workbook)
    groupsPerPhoto(allPhotoList, workbook)
    peoplePerDay(allPhotoList, workbook)
    groupsPerDay(allPhotoList, workbook)
    averagePeoplePerGroupPerPhoto(allPhotoList, workbook)
    peopleUsingSmallChairsPerPhoto(allPhotoList, workbook)
    peopleUsingSofasPerPhoto(allPhotoList, workbook)
    peopleUsingSmallTablesPerPhoto(allPhotoList, workbook)
    peopleUsingLargeTablesPerPhoto(allPhotoList, workbook)
    groupsUsingSofasPerPhoto(allPhotoList, workbook)
    groupsUsingSmallChairsPerPhoto(allPhotoList, workbook)
    groupsUsingSmallTablesPerPhoto(allPhotoList, workbook)
    groupsUsingLargeTablesPerPhoto(allPhotoList, workbook)
    
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
##        for largeTable in photo.largeTableList:
##            if largeTable.usingLargeTable(person):
##                largeTable.addPerson(person)
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
            if group.isInGroup(largeTable):
                largeTable.addGroup(group)
                for person in group.personList:
                    largeTable.addPerson(person)

def writeToExcel(workbook, sheetName, column1name, column2name, excelList):
    ##excelList is a lit of (photo.exactDate, attribute) tuples
    sheet = workbook.add_sheet(sheetName)
    sheet.write(0, 0, column1name)
    sheet.write(0, 1, column2name)
    
    r = 1
    prev = None
    for item in excelList:
        current = item[0]        
        if prev != None and  current - prev >= timedelta(minutes=30):
            difference = current - prev
            numExtras = difference.total_seconds()/(60*25) - 1
            i = 0
            while i < numExtras:
                sheet.write(r, 0, str(prev + timedelta(minutes = 15)))
                sheet.write(r, 1, 0)
                prev = prev + timedelta(minutes = 15)
                r += 1
                i += 1
        sheet.write(r, 0, str(item[0]))
        sheet.write(r, 1, item[1])
        r += 1
        prev = current

    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")           

                               
def peoplePerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numPeople))

    writeToExcel(workbook, "peoplePerPhoto", "Date", "number of people per photo", excelList)
    
    
def peoplePerDay(photoList, workbook):
    excelList = []
    currentDate = photoList[0].dayDate
    currentNumPeople = 0
    excelList.append((currentDate, currentNumPeople))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            currentNumPeople += photo.numPeople 
            excelList[i] = (currentDate, currentNumPeople)
        else:
            currentDate = photo.dayDate
            currentNumPeople = photo.numPeople
            excelList.append((currentDate, currentNumPeople))
            i += 1

    writeToExcel(workbook, "peoplePerDay", "Date", "number of people per day", excelList)
    
    
def groupsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numGroups))

    writeToExcel(workbook, "groupsPerPhoto", "Date", "number of groups per photo", excelList)
    
    
def groupsPerDay(photoList, workbook):
    excelList = []
    currentDate = photoList[0].dayDate
    currentNumGroups = 0
    excelList.append((currentDate, currentNumGroups))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            currentNumGroups += photo.numGroups 
            excelList[i] = (currentDate, currentNumGroups)
        else:
            currentDate = photo.dayDate
            currentNumGroups = photo.numGroups
            excelList.append((currentDate, currentNumGroups))
            i += 1

    writeToExcel(workbook, "groupsPerDay", "Date", "number of groups per day", excelList)
    

def averagePeoplePerGroupPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleinGroups = 0
        for group in photo.groupList:
            totalPeopleinGroups += group.numPeople
        averagePeoplePerGroup = 0
        if photo.numGroups != 0:
            averagePeoplePerGroup = float(totalPeopleinGroups)/photo.numGroups
        excelList.append((photo.exactDate, averagePeoplePerGroup))

    writeToExcel(workbook, "peoplePerGroup", "Date", "average people per group per photo", excelList)
    
    
def peopleUsingSmallChairsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleinChairs = 0
        for chair in photo.chairList:
            totalPeopleinChairs += chair.numPeople
        excelList.append((photo.exactDate, totalPeopleinChairs))

    writeToExcel(workbook, "peopleUsingChairs", "Date", "number of people using chairs per photo", excelList)
    

def peopleUsingSofasPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleinSofas = 0
        for sofa in photo.sofaList:
            totalPeopleinSofas += sofa.numPeople
        excelList.append((photo.exactDate, totalPeopleinSofas))

    writeToExcel(workbook, "peopleUsingSofas", "Date", "number of people using sofas per photo", excelList)
    
   
def peopleUsingSmallTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleUsingSmallTables = 0
        for smallTable in photo.smallTableList:
            totalPeopleUsingSmallTables += smallTable.numPeople
        excelList.append((photo.exactDate, totalPeopleUsingSmallTables))
        
    writeToExcel(workbook, "peopleUsingSmallTables", "Date", "number of groups using small tables per photo", excelList)

    
def peopleUsingLargeTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleUsingLargeTables = 0
        for largeTable in photo.largeTableList:
            totalPeopleUsingLargeTables += largeTable.numPeople
        excelList.append((photo.exactDate, totalPeopleUsingLargeTables))
    
    writeToExcel(workbook, "peopleUsingLargeTables", "Date", "number of people using large tables per photo", excelList) 

def groupsUsingSmallChairsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingChairs = 0
        for chair in photo.chairList:
            totalGroupsUsingChairs += chair.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingChairs))

    writeToExcel(workbook, "groupsUsingChairs", "Date", "number of groups using chairs per photo", excelList)  
    

def groupsUsingSofasPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingSofas = 0
        for sofa in photo.sofaList:
            totalGroupsUsingSofas += sofa.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingSofas))

    writeToExcel(workbook, "groupsUsingSofas", "Date", "number of groups using sofas per photo", excelList)   


def groupsUsingSmallTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingSmallTables = 0
        for smallTable in photo.smallTableList:
            totalGroupsUsingSmallTables += smallTable.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingSmallTables))

    writeToExcel(workbook, "groupsUsingSmallTables", "Date", "number of groups using small tables per photo", excelList)

    
def groupsUsingLargeTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingLargeTables = 0
        for largeTable in photo.largeTableList:
            totalGroupsUsingLargeTables += largeTable.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingLargeTables))

    writeToExcel(workbook, "groupsUsingLargeTables", "Date", "number of groups using large tables per photo", excelList)
    

readFiles("C:\Users\Nicole\Documents\UROP2013\data")

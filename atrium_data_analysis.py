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
        i += 1
    
    workbook = Workbook()
            
    peoplePerPhoto(allPhotoList, workbook)
    peoplePerDay(allPhotoList, workbook)
    
    groupsPerPhoto(allPhotoList, workbook)    
    groupsPerDay(allPhotoList, workbook)
    
    averagePeoplePerGroupPerPhoto(allPhotoList, workbook)
    averagePeoplePerGroupPerDay(allPhotoList, workbook)
    
    peopleUsingSmallChairsPerPhoto(allPhotoList, workbook)
    peopleUsingSmallChairsPerDay(allPhotoList, workbook)
    
    peopleUsingSofasPerPhoto(allPhotoList, workbook)
    peopleUsingSofasPerDay(allPhotoList, workbook)
    
    peopleUsingSmallTablesPerPhoto(allPhotoList, workbook)
    peopleUsingSmallTablesPerDay(allPhotoList, workbook)
    
    peopleUsingLargeTablesPerPhoto(allPhotoList, workbook)
    peopleUsingLargeTablesPerDay(allPhotoList, workbook)

    groupsUsingSmallChairsPerPhoto(allPhotoList, workbook)
    groupsUsingSmallChairsPerDay(allPhotoList, workbook)
    
    groupsUsingSofasPerPhoto(allPhotoList, workbook)
    groupsUsingSofasPerDay(allPhotoList, workbook)
    
    groupsUsingSmallTablesPerPhoto(allPhotoList, workbook)
    groupsUsingSmallTablesPerDay(allPhotoList, workbook)
    
    groupsUsingLargeTablesPerPhoto(allPhotoList, workbook)
    groupsUsingLargeTablesPerDay(allPhotoList, workbook)
    
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

def writeToExcelbyPhoto(workbook, sheetName, column1name, column2name, excelList):
    ##excelList is a lit of (photo.exactDate, attribute) tuples
    sheet = workbook.add_sheet(sheetName)
    sheet.write(0, 0, column1name)
    sheet.write(0, 1, column2name)
    sheet.write(0, 3, column1name)
    sheet.write(0, 4, column2name + " (daytime only)")
    
    r = 1
    row = 1
    prev = None
    for item in excelList:
        current = item[0]        
        if prev != None and  current - prev >= timedelta(minutes=30):
            difference = current - prev
            numExtras = difference.total_seconds()/(60*25) - 1
            i = 0
            while i < numExtras:
                toadd = prev + timedelta(minutes = 15)
                sheet.write(r, 0, str(toadd))
                sheet.write(r, 1, 0)
                if toadd.hour >= 8 and toadd.hour < 20:
                    sheet.write(row, 3, str(toadd))
                    sheet.write(row, 4, 0)
                    row += 1
                prev = prev + timedelta(minutes = 15)
                r += 1
                i += 1
        sheet.write(r, 0, str(item[0]))
        sheet.write(r, 1, item[1])
        if item[0].hour >= 8 and item[0].hour < 20:
            sheet.write(row, 3, str(item[0]))
            sheet.write(row, 4, item[1])
            row += 1
        r += 1
        prev = current

    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")

def writeToExcelbyDay(workbook, sheetName, column1name, column2name, excelList, daytimeList):
    ##excelList is a lit of (photo.exactDate, attribute) tuples
    sheet = workbook.add_sheet(sheetName)
    sheet.write(0, 0, column1name)
    sheet.write(0, 1, column2name)
    sheet.write(0, 3, column1name)
    sheet.write(0, 4, column2name + " during day")
    
    r = 1
    prev = None
    for item in excelList:
        current = item[0]        
        if prev != None and  current - prev >= timedelta(days = 2):
            difference = current - prev
            numExtras = difference.total_seconds()/(24*60*60) - 1
            i = 0
            while i < numExtras:
                sheet.write(r, 0, str(prev + timedelta(days = 1)))
                sheet.write(r, 1, 0)
                prev = prev + timedelta(days = 1)
                r += 1
                i += 1
        sheet.write(r, 0, str(item[0]))
        sheet.write(r, 1, item[1])
        r += 1
        prev = current

    row = 1
    previ = None
    for item in daytimeList:
        current = item[0]        
        if previ != None and  current - previ >= timedelta(days = 2):
            difference = current - previ
            numExtras = difference.total_seconds()/(24*60*60) - 1
            i = 0
            while i < numExtras:
                sheet.write(row, 3, str(previ + timedelta(days = 1)))
                sheet.write(row, 4, 0)
                previ = previ + timedelta(days = 1)
                row += 1
                i += 1
        sheet.write(row, 3, str(item[0]))
        sheet.write(row, 4, item[1])
        row += 1
        previ = current

    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")    

                               
def peoplePerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numPeople))

    writeToExcelbyPhoto(workbook, "peoplePerPhoto", "Date", "number of people per photo", excelList)
    
    
def peoplePerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentDate = photoList[0].dayDate
    currentNumPeople = 0
    daytimeNumPeople = 0
    daytimeList.append((currentDate, daytimeNumPeople))
    excelList.append((currentDate, currentNumPeople))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                daytimeNumPeople += photo.numPeople
                daytimeList[i] = ((currentDate, daytimeNumPeople))
            currentNumPeople += photo.numPeople 
            excelList[i] = (currentDate, currentNumPeople)
        else:
            currentDate = photo.dayDate
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                daytimeNumPeople = photo.numPeople
            else:
                daytimeNumPeople = 0
            currentNumPeople = photo.numPeople
            excelList.append((currentDate, currentNumPeople))
            daytimeList.append((currentDate, daytimeNumPeople))
            i += 1

    writeToExcelbyDay(workbook, "peoplePerDay", "Date", "number of people per day", excelList, daytimeList)
    
    
def groupsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        excelList.append((photo.exactDate, photo.numGroups))

    writeToExcelbyPhoto(workbook, "groupsPerPhoto", "Date", "number of groups per photo", excelList)
    
    
def groupsPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentDate = photoList[0].dayDate
    currentNumGroups = 0
    daytimeNumGroups = 0
    excelList.append((currentDate, currentNumGroups))
    daytimeList.append((currentDate, daytimeNumGroups))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            currentNumGroups += photo.numGroups 
            excelList[i] = (currentDate, currentNumGroups)
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                daytimeNumGroups += photo.numGroups
                daytimeList[i] = (currentDate, daytimeNumGroups)
        else:
            currentDate = photo.dayDate
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                daytimeNumGroups = photo.numGroups
            else:
                daytimeNumGroups = 0
            currentNumGroups = photo.numGroups
            excelList.append((currentDate, currentNumGroups))
            daytimeList.append((currentDate, daytimeNumGroups))
            i += 1

    writeToExcelbyDay(workbook, "groupsPerDay", "Date", "number of groups per day", excelList, daytimeList)
    

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

    writeToExcelbyPhoto(workbook, "peoplePerGroupPerPhoto", "Date", "average people per group per photo", excelList)
    
def averagePeoplePerGroupPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentDate = photoList[0].dayDate
    currentPeopleinGroups = 0
    daytimePeopleinGroups = 0
    currentNumGroups = 0
    daytimeNumGroups = 0
    currentAveragePeoplePerGroup = 0
    daytimeAveragePeoplePerGroup = 0
    excelList.append((currentDate, currentAveragePeoplePerGroup))
    daytimeList.append((currentDate, daytimeAveragePeoplePerGroup))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:            
            currentNumGroups += photo.numGroups
            for group in photo.groupList:
                currentPeopleinGroups += group.numPeople
            if currentNumGroups != 0:
                currentAveragePeoplePerGroup = float(currentPeopleinGroups)/currentNumGroups
                
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                daytimeNumGroups += photo.numGroups
                for groupB in photo.groupList:
                    daytimePeopleinGroups += groupB.numPeople                
                if daytimeNumGroups != 0:
                    daytimeAveragePeoplePerGroup = float(daytimePeopleinGroups)/daytimeNumGroups
                    
            excelList[i] = (currentDate, currentAveragePeoplePerGroup)
            daytimeList[i] = (currentDate, daytimeAveragePeoplePerGroup)
            
        else:
            currentDate = photo.dayDate
            currentNumGroups = photo.numGroups
            currentPeopleinGroups = 0
            for group in photo.groupList:
                currentPeopleinGroups += group.numPeople
            currentAveragePeoplePerGroup = 0
            if currentNumGroups != 0:
                currentAveragePeoplePerGroup = float(currentPeopleinGroups)/currentNumGroups

            daytimeNumGroups = 0
            daytimePeopleinGroups = 0
            daytimeAveragePeoplePerGroup = 0
            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:                
                daytimeNumGroups = photo.numGroups
                for groupB in photo.groupList:
                    daytimePeopleinGroups += groupB.numPeople
                if daytimeNumGroups != 0:
                    daytimeAveragePeoplePerGroup = float(daytimePeopleinGroups)/daytimeNumGroups
                
            excelList.append((currentDate, currentAveragePeoplePerGroup))
            daytimeList.append((currentDate, daytimeAveragePeoplePerGroup))
            i += 1

    writeToExcelbyDay(workbook, "peoplePerGroupPerDay", "Date", "average people per group per day", excelList, daytimeList)
    

def peopleUsingSmallChairsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleinChairs = 0
        for chair in photo.chairList:
            totalPeopleinChairs += chair.numPeople
        excelList.append((photo.exactDate, totalPeopleinChairs))

    writeToExcelbyPhoto(workbook, "peopleUsingChairsPerPhoto", "Date", "number of people using chairs per photo", excelList)

def peopleUsingSmallChairsPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentPeopleinChairs = 0
    daytimePeopleinChairs = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentPeopleinChairs))
    daytimeList.append((currentDate, daytimePeopleinChairs))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:            
            for chair in photo.chairList:
                currentPeopleinChairs += chair.numPeople
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleinChairs += chair.numPeople
            excelList[i] = (currentDate, currentPeopleinChairs)
            daytimeList[i] = (currentDate, daytimePeopleinChairs)
        else:
            currentDate = photo.dayDate
            currentPeopleinChairs = 0
            daytimePeopleinChairs = 0
            for chair in photo.chairList:
                currentPeopleinChairs += chair.numPeople
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleinChairs += chair.numPeople
            excelList.append((currentDate, currentPeopleinChairs))
            daytimeList.append((currentDate, daytimePeopleinChairs))
            i += 1

    writeToExcelbyDay(workbook, "peopleUsingChairsPerDay", "Date", "number of people using chairs per day", excelList, daytimeList)
    

def peopleUsingSofasPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleinSofas = 0
        for sofa in photo.sofaList:
            totalPeopleinSofas += sofa.numPeople
        excelList.append((photo.exactDate, totalPeopleinSofas))

    writeToExcelbyPhoto(workbook, "peopleUsingSofasPerPhoto", "Date", "number of people using sofas per photo", excelList)

def peopleUsingSofasPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentPeopleinSofas = 0
    daytimePeopleinSofas = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentPeopleinSofas))
    daytimeList.append((currentDate, daytimePeopleinSofas))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for sofa in photo.sofaList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleinSofas += sofa.numPeople
                currentPeopleinSofas += sofa.numPeople
            excelList[i] = (currentDate, currentPeopleinSofas)
            daytimeList[i] = (currentDate, daytimePeopleinSofas)
        else:
            currentDate = photo.dayDate
            currentPeopleinSofas = 0
            daytimePeopleinSofas = 0
            for sofa in photo.sofaList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleinSofas += sofa.numPeople
                currentPeopleinSofas += sofa.numPeople
            excelList.append((currentDate, currentPeopleinSofas))
            daytimeList.append((currentDate, daytimePeopleinSofas))
            i += 1

    writeToExcelbyDay(workbook, "peopleUsingSofasPerDay", "Date", "number of people using sofas per day", excelList, daytimeList)   

   
def peopleUsingSmallTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleUsingSmallTables = 0
        for smallTable in photo.smallTableList:
            totalPeopleUsingSmallTables += smallTable.numPeople
        excelList.append((photo.exactDate, totalPeopleUsingSmallTables))
        
    writeToExcelbyPhoto(workbook, "peopleUsingSmallTablesPerPhoto", "Date", "number of groups using small tables per photo", excelList)

def peopleUsingSmallTablesPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentPeopleUsingSmallTables = 0
    daytimePeopleUsingSmallTables = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentPeopleUsingSmallTables))
    daytimeList.append((currentDate, daytimePeopleUsingSmallTables))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for table in photo.smallTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleUsingSmallTables += table.numPeople
                currentPeopleUsingSmallTables += table.numPeople
            excelList[i] = (currentDate, currentPeopleUsingSmallTables)
            daytimeList[i] = (currentDate, daytimePeopleUsingSmallTables)
        else:
            currentDate = photo.dayDate
            currentPeopleUsingSmallTables = 0
            daytimePeopleUsingSmallTables = 0
            for table in photo.smallTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleUsingSmallTables += table.numPeople
                currentPeopleUsingSmallTables+= table.numPeople
            excelList.append((currentDate, currentPeopleUsingSmallTables))
            daytimeList.append((currentDate, daytimePeopleUsingSmallTables))
            i += 1

    writeToExcelbyDay(workbook, "peopleUsingSmallTablesPerDay", "Date", "number of people using small tables per day", excelList, daytimeList)
    
    
def peopleUsingLargeTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalPeopleUsingLargeTables = 0
        for largeTable in photo.largeTableList:
            totalPeopleUsingLargeTables += largeTable.numPeople
        excelList.append((photo.exactDate, totalPeopleUsingLargeTables))
    
    writeToExcelbyPhoto(workbook, "peopleUsingLargeTablesPerPhoto", "Date", "number of people using large tables per photo", excelList) 

def peopleUsingLargeTablesPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentPeopleUsingLargeTables = 0
    daytimePeopleUsingLargeTables = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentPeopleUsingLargeTables))
    daytimeList.append((currentDate, daytimePeopleUsingLargeTables))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for table in photo.largeTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleUsingLargeTables += table.numPeople
                currentPeopleUsingLargeTables += table.numPeople
            excelList[i] = (currentDate, currentPeopleUsingLargeTables)
            daytimeList[i] = (currentDate, daytimePeopleUsingLargeTables)
        else:
            currentDate = photo.dayDate
            currentPeopleUsingLargeTables = 0
            daytimePeopleUsingLargeTables = 0
            for table in photo.largeTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimePeopleUsingLargeTables += table.numPeople
                currentPeopleUsingLargeTables+= table.numPeople
            excelList.append((currentDate, currentPeopleUsingLargeTables))
            daytimeList.append((currentDate, daytimePeopleUsingLargeTables))
            i += 1

    writeToExcelbyDay(workbook, "peopleUsingLargeTablesPerDay", "Date", "number of people using large tables per day", excelList, daytimeList)
    
def groupsUsingSmallChairsPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingChairs = 0
        for chair in photo.chairList:
            totalGroupsUsingChairs += chair.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingChairs))

    writeToExcelbyPhoto(workbook, "groupsUsingChairsPerPhoto", "Date", "number of groups using chairs per photo", excelList)  
    
def groupsUsingSmallChairsPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentGroupsinChairs = 0
    daytimeGroupsinChairs = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentGroupsinChairs))
    daytimeList.append((currentDate, daytimeGroupsinChairs))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for chair in photo.chairList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsinChairs += chair.numGroups
                currentGroupsinChairs += chair.numGroups
            excelList[i] = (currentDate, currentGroupsinChairs)
            daytimeList[i] = (currentDate, daytimeGroupsinChairs)
        else:
            currentDate = photo.dayDate
            currentGroupsinChairs = 0
            daytimeGroupsinChairs = 0
            for chair in photo.chairList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsinChairs += chair.numGroups
                currentGroupsinChairs += chair.numGroups
            excelList.append((currentDate, currentGroupsinChairs))
            daytimeList.append((currentDate, daytimeGroupsinChairs))
            i += 1

    writeToExcelbyDay(workbook, "groupsUsingChairsPerDay", "Date", "number of groups using chairs per day", excelList, daytimeList)

def groupsUsingSofasPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingSofas = 0
        for sofa in photo.sofaList:
            totalGroupsUsingSofas += sofa.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingSofas))

    writeToExcelbyPhoto(workbook, "groupsUsingSofasPerPhoto", "Date", "number of groups using sofas per photo", excelList)

def groupsUsingSofasPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentGroupsinSofas = 0
    daytimeGroupsinSofas = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentGroupsinSofas))
    daytimeList.append((currentDate, daytimeGroupsinSofas))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for sofa in photo.sofaList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsinSofas += sofa.numGroups
                currentGroupsinSofas += sofa.numGroups
            excelList[i] = (currentDate, currentGroupsinSofas)
            daytimeList[i] = (currentDate, daytimeGroupsinSofas)
        else:
            currentDate = photo.dayDate
            currentGroupsinSofas = 0
            daytimeGroupsinSofas = 0
            for chair in photo.sofaList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsinSofas += sofa.numGroups
                currentGroupsinSofas += sofa.numGroups
            excelList.append((currentDate, currentGroupsinSofas))
            daytimeList.append((currentDate, daytimeGroupsinSofas))
            i += 1

    writeToExcelbyDay(workbook, "groupsUsingSofasPerDay", "Date", "number of groups using sofas per day", excelList, daytimeList)


def groupsUsingSmallTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingSmallTables = 0
        for smallTable in photo.smallTableList:
            totalGroupsUsingSmallTables += smallTable.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingSmallTables))

    writeToExcelbyPhoto(workbook, "groupsUsingSmallTablesPerPhoto", "Date", "number of groups using small tables per photo", excelList)

def groupsUsingSmallTablesPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentGroupsUsingSmallTables = 0
    daytimeGroupsUsingSmallTables = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentGroupsUsingSmallTables))
    daytimeList.append((currentDate ,daytimeGroupsUsingSmallTables))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for table in photo.smallTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsUsingSmallTables += table.numGroups
                currentGroupsUsingSmallTables += table.numGroups
            excelList[i] = (currentDate, currentGroupsUsingSmallTables)
            daytimeList[i] = (currentDate, daytimeGroupsUsingSmallTables)
        else:
            currentDate = photo.dayDate
            currentGroupsUsingSmallTables = 0
            daytimeGroupsUsingSmallTables = 0
            for table in photo.smallTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsUsingSmallTables += table.numGroups
                currentGroupsUsingSmallTables += table.numGroups
            excelList.append((currentDate, currentGroupsUsingSmallTables))
            daytimeList.append((currentDate, daytimeGroupsUsingSmallTables))
            i += 1

    writeToExcelbyDay(workbook, "groupsUsingSmallTablesPerDay", "Date", "number of groups using small tables per day", excelList, daytimeList)

    
def groupsUsingLargeTablesPerPhoto(photoList, workbook):
    excelList = []
    for photo in photoList:
        totalGroupsUsingLargeTables = 0
        for largeTable in photo.largeTableList:
            totalGroupsUsingLargeTables += largeTable.numGroups
        excelList.append((photo.exactDate, totalGroupsUsingLargeTables))

    writeToExcelbyPhoto(workbook, "groupsUsingLargeTablesPerPhoto", "Date", "number of groups using large tables per photo", excelList)

def groupsUsingLargeTablesPerDay(photoList, workbook):
    excelList = []
    daytimeList = []
    currentGroupsUsingLargeTables = 0
    daytimeGroupsUsingLargeTables = 0
    currentDate = photoList[0].dayDate
    excelList.append((currentDate, currentGroupsUsingLargeTables))
    daytimeList.append((currentDate, daytimeGroupsUsingLargeTables))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            for table in photo.largeTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsUsingLargeTables += table.numGroups
                currentGroupsUsingLargeTables += table.numGroups
            excelList[i] = (currentDate, currentGroupsUsingLargeTables)
            daytimeList[i] = (currentDate, daytimeGroupsUsingLargeTables)
        else:
            currentDate = photo.dayDate
            currentGroupsUsingLargeTables = 0
            daytimeGroupsUsingLargeTables = 0
            for table in photo.largeTableList:
                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
                    daytimeGroupsUsingLargeTables += table.numGroups
                currentGroupsUsingLargeTables += table.numGroups
            excelList.append((currentDate, currentGroupsUsingLargeTables))
            daytimeList.append((currentDate, daytimeGroupsUsingLargeTables))
            i += 1

    writeToExcelbyDay(workbook, "groupsUsingLargeTablesPerDay", "Date", "number of groups using large tables per day", excelList, daytimeList)


    

readFiles("C:\Users\Nicole\Documents\UROP2013\data")

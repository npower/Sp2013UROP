import os
from datetime import *
from xlwt.Workbook import *
from atrium_objects import *       
    

# input: directory path(string)
# goes through all the files in a directory and calls parseFile on them
# joins the outputs of all the parseFile calls and returns it
def readFiles(path):
    allDictList = []
    coveredDatesDict = {}
    for file in os.listdir(path):
        fileDatesList = []
        for dicty, date in parseFile(path + "\\" + file):
            if date not in coveredDatesDict:
                allDictList.append(dicty)
                fileDatesList.append(date)
        for date in fileDatesList:
            coveredDatesDict[date] = 1
    
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

    noOutliersPhotoList = []
    for photo in allPhotoList:
        if not isOutlier(photo.exactDate):
            noOutliersPhotoList.append(photo)

    daytimePhotoList = []
    daytimeNoOutliersPhotoList = []
    nighttimeNoOutliersPhotoList = []
    for photo in allPhotoList:
        if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
            daytimePhotoList.append(photo)
            if not isOutlier(photo.exactDate):
                daytimeNoOutliersPhotoList.append(photo)
        else:
            if not isOutlier(photo.exactDate):
                nighttimeNoOutliersPhotoList.append(photo)
        
    workbook = Workbook()
        
##    peoplePerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    peoplePerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    groupsPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)    
##    groupsPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    averagePeoplePerGroupPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    averagePeoplePerGroupPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    peopleUsingSmallChairsPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    peopleUsingSmallChairsPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    peopleUsingSofasPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    peopleUsingSofasPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    peopleUsingSmallTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    peopleUsingSmallTablesPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    groupsUsingSmallChairsPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    groupsUsingSmallChairsPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    groupsUsingSofasPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    groupsUsingSofasPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
    groupsUsingSmallTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    groupsUsingSmallTablesPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    smallChairUtilizationPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    smallChairUtilizationPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    sofaUtilizationPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    sofaUtilizationPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    smallTableUtilizationPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    smallTableUtilizationPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    portionOfPeopleUsingFurniturePerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfPeopleUsingFurniturePerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    portionOfPeopleUsingSmallChairsPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfPeopleUsingSmallChairsPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    portionOfPeopleUsingSofasPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfPeopleUsingSofasPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    portionOfPeopleUsingSmallTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfPeopleUsingSmallTablesPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##    
##    portionOfGroupsUsingSmallChairsPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfGroupsUsingSmallChairsPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    portionOfGroupsUsingSofasPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfGroupsUsingSofasPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
##
##    portionOfGroupsUsingSmallTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    portionOfGroupsUsingSmallTablesPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)
    
##    peopleUsingLargeTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    peopleUsingLargeTablesPerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook)    
##    groupsUsingLargeTablesPerPhoto(allPhotoList, daytimePhotoList, nighttimeNoOutliersPhotoList, workbook)
##    groupsUsingLargeTablesPerDay(allPhotoList, workbook)
##    largeTableUtilizationPerPhoto(allPhotoList, workbook)
##    largeTableUtilizationPerDay(allPhotoList, workbook)
##    portionOfPeopleUsingLargeTablesPerPhoto(allPhotoList, workbook)
##    portionOfPeopleUsingLargeTablesPerDay(allPhotoList, workbook)
##    portionOfGroupsUsingLargeTablesPerPhoto(allPhotoList, workbook)
##    portionOfGroupsUsingLargeTablesPerDay(allPhotoList, workbook)

    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")
    
# input: a filename (string)
# opens the file, goes through all the lines and extracts the fields
# returns an unique list of dictionaries (one for each line)    
def parseFile(filePath):
    f = open(filePath) 
    linesList = f.readlines()
    f.close()
    
    noDuplicates = list(set(linesList))
    
    fileObjectDictsDatesList = []
    dateList = []
    for line in noDuplicates:        
        fileObjectDictsDatesList.append(tokenize(line))
        
    return fileObjectDictsDatesList

    
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

    date = datetime(int(objectDict['year']), int(objectDict['month']), int(objectDict['day']),
                    int(objectDict['hour']), int(objectDict['minute']))
                    
    return objectDict, date


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
        
def isOutlier(x):
        if x >= datetime(2013, 2, 1, 9, 30) and x <= datetime(2013, 2, 2, 15, 45):
            return True
        elif x >= datetime(2013, 2, 21, 13, 0) and x <= datetime(2013, 2, 21, 14, 45):
            return True
        elif x >= datetime(2013, 2, 22, 12, 00) and x <= datetime(2013, 2, 25, 10, 0):
            return True
        elif x >= datetime(2013, 2, 28, 10, 0) and x <= datetime(2013, 2, 28, 2, 40):
            return True
        elif x >= datetime(2013, 2, 28, 8, 0) and x <= datetime(2013, 2, 28, 11, 0):
            return True
        elif x >= datetime(2013, 3, 6, 9, 15) and x <= datetime(2013, 3, 7, 14, 0):
            return True
        elif x >= datetime(2013, 3, 8, 16, 0) and x <= datetime(2013, 3, 8, 17, 0):
            return True
        elif x >= datetime(2013, 3, 9, 18, 0) and x <= datetime(2013, 3, 10, 13, 45):
            return True
        elif x >= datetime(2013, 3, 19, 16, 0) and x <= datetime(2013, 3, 20, 14, 0):
            return True
        elif x >= datetime(2013, 3, 28, 10, 0) and x <= datetime(2013, 3, 29, 12, 0):
            return True
        elif x >= datetime(2013, 4, 1, 12, 0) and x <= datetime(2013, 4, 1, 13, 0):
            return True
        else:
            return False

def writeToExcelbyPhoto(workbook, sheetName, column1name, column2name, excelList, daytimeList, nighttimeNoOutliersList):
    ##excelList is a lit of (photo.exactDate, attribute) tuples
    sheet = workbook.add_sheet(sheetName)
    sheet.write(0, 0, column1name)
    sheet.write(0, 1, column2name)
    sheet.write(0, 3, column1name)
    sheet.write(0, 4, column2name + " (daytime only)")

    overallNoOutliers = [(a.date(), b) for (a, b) in excelList if not isOutlier(a)]
    daytimeNoOutliers = [(a.date(), b) for (a, b) in daytimeList if not isOutlier(a)]
    nighttimeNoOutliers = [(a.date(), b) for (a, b) in nighttimeNoOutliersList]

    periodList = [(date.min, date.max, "jan 8 to april 1"),
        (date(2013, 1, 8), date(2013, 1, 22), "jan 8 to jan 22"),
        (date(2013, 1, 23), date(2013, 2, 4), "jan 23 to feb 4"),
        (date(2013, 2, 5), date(2013, 2, 17), "feb 5 to feb 17"),
        (date(2013, 2, 18), date(2013, 3, 3), "feb 18 to march 3"),
        (date(2013, 3, 4), date(2013, 3, 17), "march 4 to march 17"),
        (date(2013, 3, 18), date(2013, 4, 1), "march 18 to april 1")]

    y = 1
    for (a, b, c) in periodList:
        stats(a, b, sheet, overallNoOutliers, daytimeNoOutliers, nighttimeNoOutliers, c, y)
        y += 7
    
    r = 1
    row = 1
    prev = None
    for item in excelList:
        current = item[0]        
        if prev != None and  current - prev >= timedelta(minutes=28):
            difference = current - prev
            numExtras = difference.total_seconds()/(60*15) - 1
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

##    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")
    
def average(l):
    return (1.0*sum(l))/len(l)                


def stats(startDate, endDate, sheet, overallNoOutliers, daytimeNoOutliers, nighttimeNoOutliers, period, i):
    overall = [b for (a, b) in overallNoOutliers if a >= startDate and a <= endDate]
    daytime = [b for (a, b) in daytimeNoOutliers if a >= startDate and a <= endDate]
    nighttime = [b for (a, b) in nighttimeNoOutliers if a >= startDate and a <= endDate]

    sheet.write(i, 6, period + " overall max")
    sheet.write(i, 7, max(overall))
    sheet.write(i+1, 6, period + " overall average")
    sheet.write(i+1, 7, average(overall)) 
    sheet.write(i+2, 6, period + " daytime max")
    sheet.write(i+2, 7, max(daytime))
    sheet.write(i+3, 6, period + " daytime average")
    sheet.write(i+3, 7, average(daytime))
    sheet.write(i+4, 6, period + " nighttime max")
    sheet.write(i+4, 7, max(nighttime))
    sheet.write(i+5, 6, period + " nighttime average")
    sheet.write(i+5, 7, average(nighttime)) 


def writeToExcelbyDay(workbook, sheetName, column1name, column2name, excelList, daytimeList, noOutliersList, daytimeNoOutliersList, nighttimeNoOutliersList):
    ##excelList is a lit of (photo.dayDate, attribute) tuples
    sheet = workbook.add_sheet(sheetName)
    sheet.write(0, 0, column1name)
    sheet.write(0, 1, column2name)
    sheet.write(0, 3, column1name)
    sheet.write(0, 4, column2name + " during day")
    sheet.write(0, 6, column2name)

    periodList = [(date.min, date.max, "jan 8 to april 1"), 
        (date(2013, 1, 8), date(2013, 1, 22), "jan 8 to jan 22"),
        (date(2013, 1, 23), date(2013, 2, 4), "jan 23 to feb 4"),
        (date(2013, 2, 5), date(2013, 2, 17), "feb 5 to feb 17"),
        (date(2013, 2, 18), date(2013, 3, 3), "feb 18 to march 3"),
        (date(2013, 3, 4), date(2013, 3, 17), "march 4 to march 17"),
        (date(2013, 3, 18), date(2013, 4, 1), "march 18 to april 1")]

    y = 1
    for (a, b, c) in periodList:
        stats(a, b, sheet, noOutliersList, daytimeNoOutliersList, nighttimeNoOutliersList, c, y)
        y += 7
    
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

##    workbook.save("C:\Users\Nicole\Documents\UROP2013\AtriumData.xls")    

                               
def peoplePerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overallList = [(photo.exactDate, photo.numPeople) for photo in photoList]
    daytimeList = [(photo.exactDate, photo.numPeople) for photo in daytimePhotoList]
    nighttimeList = [(photo.exactDate, photo.numPeople) for photo in nighttimePhotoList]

    writeToExcelbyPhoto(workbook, "peoplePerPhoto", "Date", "number of people per photo", overallList, daytimeList, nighttimeList)
    

def A_per_day_helper(photoList, getNumA):
    myList = []
    currentDate = photoList[0].dayDate
    currentNumA = 0
    myList.append((currentDate, currentNumA))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            currentNumA += getNumA(photo) 
            myList[i] = (currentDate, currentNumA)
        else:
            currentDate = photo.dayDate
            currentNumA = getNumA(photo)
            myList.append((currentDate, currentNumA))
            i += 1
    return myList

    
def peoplePerDay(allPhotoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overallList = A_per_day_helper(allPhotoList, lambda x: x.numPeople)
    daytimeList = A_per_day_helper(daytimePhotoList, lambda x: x.numPeople)
    noOutliersList = A_per_day_helper(noOutliersPhotoList, lambda x: x.numPeople)
    daytimeNoOutliersList = A_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.numPeople)
    nighttimeNoOutliersList = A_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.numPeople)

    writeToExcelbyDay(workbook, "peoplePerDay", "Date", "number of people per day", overallList, daytimeList, noOutliersList, daytimeNoOutliersList, nighttimeNoOutliersList)
    
    
def groupsPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overallList = [(photo.exactDate, photo.numGroups) for photo in photoList]
    daytimeList = [(photo.exactDate, photo.numGroups) for photo in daytimePhotoList]
    nighttimeList = [(photo.exactDate, photo.numGroups) for photo in nighttimePhotoList]

    writeToExcelbyPhoto(workbook, "groupsPerPhoto", "Date", "number of groups per photo", overallList, daytimeList, nighttimeList)
    
    
def groupsPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overallList = A_per_day_helper(photoList, lambda x: x.numGroups)
    daytimeList = A_per_day_helper(daytimePhotoList, lambda x: x.numGroups)
    noOutliersList = A_per_day_helper(noOutliersPhotoList, lambda x: x.numGroups)
    daytimeNoOutliersList = A_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.numGroups)
    nighttimeList = A_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.numGroups)

    writeToExcelbyDay(workbook, "groupsPerDay", "Date", "number of groups per day", overallList, daytimeList, noOutliersList, daytimeNoOutliersList, nighttimeList)

def averagePeoplePerGroupPerPhotoHelper(photoList):
    myList = []
    for photo in photoList:
        totalPeopleinGroups = 0
        for group in photo.groupList:
            totalPeopleinGroups += group.numPeople
        averagePeoplePerGroup = 0
        if photo.numGroups != 0:
            averagePeoplePerGroup = float(totalPeopleinGroups)/photo.numGroups
        myList.append((photo.exactDate, averagePeoplePerGroup))
    return myList

def averagePeoplePerGroupPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = averagePeoplePerGroupPerPhotoHelper(photoList)
    daytime = averagePeoplePerGroupPerPhotoHelper(daytimePhotoList)
    nighttime = averagePeoplePerGroupPerPhotoHelper(nighttimePhotoList)

    writeToExcelbyPhoto(workbook, "peoplePerGroupPerPhoto", "Date", "average people per group per photo", overall, daytime, nighttime)

def averagePeoplePerGroupPerDayHelper(photoList):
    myList = []
    currentDate = photoList[0].dayDate
    currentPeopleinGroups = 0
    currentNumGroups = 0
    currentAveragePeoplePerGroup = 0
    myList.append((currentDate, currentAveragePeoplePerGroup))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:            
            currentNumGroups += photo.numGroups
            for group in photo.groupList:
                currentPeopleinGroups += group.numPeople
            if currentNumGroups != 0:
                currentAveragePeoplePerGroup = float(currentPeopleinGroups)/currentNumGroups
            myList[i] = (currentDate, currentAveragePeoplePerGroup)            
        else:
            currentDate = photo.dayDate
            currentNumGroups = photo.numGroups
            currentPeopleinGroups = 0
            for group in photo.groupList:
                currentPeopleinGroups += group.numPeople
            currentAveragePeoplePerGroup = 0
            if currentNumGroups != 0:
                currentAveragePeoplePerGroup = float(currentPeopleinGroups)/currentNumGroups                
            myList.append((currentDate, currentAveragePeoplePerGroup))
            i += 1
    return myList
    
def averagePeoplePerGroupPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overallList = averagePeoplePerGroupPerDayHelper(photoList)
    daytimeList = averagePeoplePerGroupPerDayHelper(daytimePhotoList)
    noOutliers = averagePeoplePerGroupPerDayHelper(noOutliersPhotoList)
    daytimeNoOutliers = averagePeoplePerGroupPerDayHelper(daytimeNoOutliersPhotoList)
    nighttimeNoOutliers = averagePeoplePerGroupPerDayHelper(nighttimeNoOutliersPhotoList)

    writeToExcelbyDay(workbook, "peoplePerGroupPerDay", "Date", "average people per group per day", overallList, daytimeList, noOutliers, daytimeNoOutliers, nighttimeNoOutliers)
    

def A_using_B_per_Photo_Helper(photoList, getBList, getNumA):
    myList = []
    for photo in photoList:
        totalAinB = 0
        for b in getBList(photo):
            totalAinB += getNumA(b)
        myList.append((photo.exactDate, totalAinB))    
    return myList
            
def peopleUsingSmallChairsPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.chairList, lambda y: y.numPeople)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.numPeople)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.chairList, lambda y: y.numPeople)

    writeToExcelbyPhoto(workbook, "peopleUsingChairsPerPhoto", "Date", "number of people using chairs per photo", overall, daytime, nighttime)

def A_using_B_per_day_helper(photoList, getBList, getNumA):
    myList = []
    currentAinB = 0
    currentDate = photoList[0].dayDate
    myList.append((currentDate, currentAinB))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:            
            for b in getBList(photo):
                currentAinB += getNumA(b)                
            myList[i] = (currentDate, currentAinB)
        else:
            currentDate = photo.dayDate
            currentAinB = 0
            for b in getBList(photo):
                currentAinB += getNumA(b)
            myList.append((currentDate, currentAinB))
            i += 1
    return myList

def peopleUsingSmallChairsPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.chairList, lambda y: y.numPeople)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.numPeople)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.chairList, lambda y: y.numPeople)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.numPeople)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.chairList, lambda y: y.numPeople)    

    writeToExcelbyDay(workbook, "peopleUsingChairsPerDay", "Date", "number of people using chairs per day", overall, daytime, noOut, daytimeNoOut, nighttime)
    

def peopleUsingSofasPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.sofaList, lambda y: y.numPeople)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.numPeople)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.numPeople)

    writeToExcelbyPhoto(workbook, "peopleUsingSofasPerPhoto", "Date", "number of people using sofas per photo", overall, daytime, nighttime)

def peopleUsingSofasPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.sofaList, lambda y: y.numPeople)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.numPeople)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.sofaList, lambda y: y.numPeople)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.numPeople)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.numPeople)

    writeToExcelbyDay(workbook, "peopleUsingSofasPerDay", "Date", "number of people using sofas per day", overall, daytime, noOut, daytimeNoOut, nighttime)   

   
def peopleUsingSmallTablesPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)
        
    writeToExcelbyPhoto(workbook, "peopleUsingSmallTablesPerPhoto", "Date", "number of groups using small tables per photo", overall, daytime, nighttime)

def peopleUsingSmallTablesPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.numPeople)

    writeToExcelbyDay(workbook, "peopleUsingSmallTablesPerDay", "Date", "number of people using small tables per day", overall, daytime, noOut, daytimeNoOut, nighttime)
    
    
##def peopleUsingLargeTablesPerPhoto(photoList, workbook):
##    excelList = []
##    for photo in photoList:
##        totalPeopleUsingLargeTables = 0
##        for largeTable in photo.largeTableList:
##            totalPeopleUsingLargeTables += largeTable.numPeople
##        excelList.append((photo.exactDate, totalPeopleUsingLargeTables))
##    
##    writeToExcelbyPhoto(workbook, "peopleUsingLargeTablesPerPhoto", "Date", "number of people using large tables per photo", excelList) 
##
##def peopleUsingLargeTablesPerDay(photoList, workbook):
##    excelList = []
##    daytimeList = []
##    currentPeopleUsingLargeTables = 0
##    daytimePeopleUsingLargeTables = 0
##    currentDate = photoList[0].dayDate
##    excelList.append((currentDate, currentPeopleUsingLargeTables))
##    daytimeList.append((currentDate, daytimePeopleUsingLargeTables))
##    i = 0
##    for photo in photoList:
##        if photo.dayDate == currentDate:
##            for table in photo.largeTableList:
##                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                    daytimePeopleUsingLargeTables += table.numPeople
##                currentPeopleUsingLargeTables += table.numPeople
##            excelList[i] = (currentDate, currentPeopleUsingLargeTables)
##            daytimeList[i] = (currentDate, daytimePeopleUsingLargeTables)
##        else:
##            currentDate = photo.dayDate
##            currentPeopleUsingLargeTables = 0
##            daytimePeopleUsingLargeTables = 0
##            for table in photo.largeTableList:
##                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                    daytimePeopleUsingLargeTables += table.numPeople
##                currentPeopleUsingLargeTables+= table.numPeople
##            excelList.append((currentDate, currentPeopleUsingLargeTables))
##            daytimeList.append((currentDate, daytimePeopleUsingLargeTables))
##            i += 1
##
##    writeToExcelbyDay(workbook, "peopleUsingLargeTablesPerDay", "Date", "number of people using large tables per day", excelList, daytimeList)
    
def groupsUsingSmallChairsPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.chairList, lambda y: y.numGroups)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.numGroups)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.chairList, lambda y: y.numGroups)

    writeToExcelbyPhoto(workbook, "groupsUsingChairsPerPhoto", "Date", "number of groups using chairs per photo", overall, daytime, nighttime)  
    
def groupsUsingSmallChairsPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.chairList, lambda y: y.numGroups)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.numGroups)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.chairList, lambda y: y.numGroups)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.numGroups)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.chairList, lambda y: y.numGroups)  

    writeToExcelbyDay(workbook, "groupsUsingChairsPerDay", "Date", "number of groups using chairs per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def groupsUsingSofasPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.sofaList, lambda y: y.numGroups)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.numGroups)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.numGroups)

    writeToExcelbyPhoto(workbook, "groupsUsingSofasPerPhoto", "Date", "number of groups using sofas per photo", overall, daytime, nighttime)

def groupsUsingSofasPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.sofaList, lambda y: y.numGroups)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.numGroups)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.sofaList, lambda y: y.numGroups)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.numGroups)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.numGroups)  

    writeToExcelbyDay(workbook, "groupsUsingSofasPerDay", "Date", "number of groups using sofas per day", overall, daytime, noOut, daytimeNoOut, nighttime)


def groupsUsingSmallTablesPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_Photo_Helper(photoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    daytime = A_using_B_per_Photo_Helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    nighttime = A_using_B_per_Photo_Helper(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)

    writeToExcelbyPhoto(workbook, "groupsUsingSmallTablesPerPhoto", "Date", "number of groups using small tables per photo", overall, daytime, nighttime)

def groupsUsingSmallTablesPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = A_using_B_per_day_helper(photoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    daytime = A_using_B_per_day_helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    noOut = A_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    daytimeNoOut = A_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    nighttime = A_using_B_per_day_helper(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.numGroups)
    
    writeToExcelbyDay(workbook, "groupsUsingSmallTablesPerDay", "Date", "number of groups using small tables per day", overall, daytime, noOut, daytimeNoOut, nighttime)

    
##def groupsUsingLargeTablesPerPhoto(photoList, workbook):
##    excelList = []
##    for photo in photoList:
##        totalGroupsUsingLargeTables = 0
##        for largeTable in photo.largeTableList:
##            totalGroupsUsingLargeTables += largeTable.numGroups
##        excelList.append((photo.exactDate, totalGroupsUsingLargeTables))
##
##    writeToExcelbyPhoto(workbook, "groupsUsingLargeTablesPerPhoto", "Date", "number of groups using large tables per photo", excelList)
##
##def groupsUsingLargeTablesPerDay(photoList, workbook):
##    excelList = []
##    daytimeList = []
##    currentGroupsUsingLargeTables = 0
##    daytimeGroupsUsingLargeTables = 0
##    currentDate = photoList[0].dayDate
##    excelList.append((currentDate, currentGroupsUsingLargeTables))
##    daytimeList.append((currentDate, daytimeGroupsUsingLargeTables))
##    i = 0
##    for photo in photoList:
##        if photo.dayDate == currentDate:
##            for table in photo.largeTableList:
##                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                    daytimeGroupsUsingLargeTables += table.numGroups
##                currentGroupsUsingLargeTables += table.numGroups
##            excelList[i] = (currentDate, currentGroupsUsingLargeTables)
##            daytimeList[i] = (currentDate, daytimeGroupsUsingLargeTables)
##        else:
##            currentDate = photo.dayDate
##            currentGroupsUsingLargeTables = 0
##            daytimeGroupsUsingLargeTables = 0
##            for table in photo.largeTableList:
##                if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                    daytimeGroupsUsingLargeTables += table.numGroups
##                currentGroupsUsingLargeTables += table.numGroups
##            excelList.append((currentDate, currentGroupsUsingLargeTables))
##            daytimeList.append((currentDate, daytimeGroupsUsingLargeTables))
##            i += 1
##
##    writeToExcelbyDay(workbook, "groupsUsingLargeTablesPerDay", "Date", "number of groups using large tables per day", excelList, daytimeList)

def furnUtilHelper(photoList, getFurnList):
    myList = []
    for photo in photoList:
        totalFurninUse = 0
        for a in getFurnList(photo):
            if a.numPeople != 0 or a.numGroups != 0:
                totalFurninUse += 1
        percentFurninUse = 0
        if len(getFurnList(photo)) != 0:
            percentFurninUse = (100.0 * totalFurninUse) / len(getFurnList(photo))
        myList.append((photo.exactDate, percentFurninUse))
    return myList

def smallChairUtilizationPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    ##chairs in use / total chairs
    overall = furnUtilHelper(photoList, lambda x: x.chairList)
    daytime = furnUtilHelper(photoList, lambda x: x.chairList)
    nighttime = furnUtilHelper(photoList, lambda x: x.chairList)

    writeToExcelbyPhoto(workbook, "smallChairUtilizationPerPhoto", "Date", "percent of chairs being used per photo", overall, daytime, nighttime)

def furnUtilHelperPerDay(photoList, getFurnList):
    myList = []
    currentDate = photoList[0].dayDate
    currentFurnUsed = 0
    currentTotalFurn = 0
    currentPercentFurnUsed = 0
    myList.append((currentDate, currentPercentFurnUsed))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:            
            currentTotalFurn += len(getFurnList(photo))
            for furn in getFurnList(photo):
                if furn.numPeople != 0 or furn.numGroups != 0:
                    currentFurnUsed += 1
            if currentTotalFurn != 0:
                currentPercentFurnUsed = (100.0 * currentFurnUsed) / currentTotalFurn
            myList[i] = (currentDate, currentPercentFurnUsed)
            
        else:
            currentDate = photo.dayDate
            currentTotalFurn = len(getFurnList(photo))
            currentFurnUsed = 0
            for furn in getFurnList(photo):
                if furn.numPeople != 0 or furn.numGroups != 0:
                    currentFurnUsed += 1
            currentPercentFurnUsed = 0
            if currentTotalFurn != 0:
                currentPercentFurn = (100.0 * currentFurnUsed) / currentTotalFurn

            myList.append((currentDate, currentPercentFurnUsed))
            i += 1
    return myList

def smallChairUtilizationPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = furnUtilHelperPerDay(photoList, lambda x: x.chairList)
    daytime = furnUtilHelperPerDay(daytimePhotoList, lambda x: x.chairList)
    noOut = furnUtilHelperPerDay(noOutliersPhotoList, lambda x: x.chairList)
    daytimeNoOut = furnUtilHelperPerDay(daytimeNoOutliersPhotoList, lambda x: x.chairList)
    nighttime = furnUtilHelperPerDay(nighttimePhotoList, lambda x: x.chairList)

    writeToExcelbyDay(workbook, "smallChairUtilPerDay", "Date", "small chair utlization per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def sofaUtilizationPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = furnUtilHelper(photoList, lambda x: x.sofaList)
    daytime = furnUtilHelper(photoList, lambda x: x.sofaList)
    nighttime = furnUtilHelper(photoList, lambda x: x.sofaList)

    writeToExcelbyPhoto(workbook, "sofaUtilizationPerPhoto", "Date", "percent of sofas being used per photo", overall, daytime, nighttime)

def sofaUtilizationPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = furnUtilHelperPerDay(photoList, lambda x: x.sofaList)
    daytime = furnUtilHelperPerDay(daytimePhotoList, lambda x: x.sofaList)
    noOut = furnUtilHelperPerDay(noOutliersPhotoList, lambda x: x.sofaList)
    daytimeNoOut = furnUtilHelperPerDay(daytimeNoOutliersPhotoList, lambda x: x.sofaList)
    nighttime = furnUtilHelperPerDay(nighttimePhotoList, lambda x: x.sofaList)
    
    writeToExcelbyDay(workbook, "sofaUtilPerDay", "Date", "sofa utlization per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def smallTableUtilizationPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = furnUtilHelper(photoList, lambda x: x.smallTableList)
    daytime = furnUtilHelper(photoList, lambda x: x.smallTableList)
    nighttime = furnUtilHelper(photoList, lambda x: x.smallTableList)

    writeToExcelbyPhoto(workbook, "smallTableUtilizationPerPhoto", "Date", "percent of small tables being used per photo", overall, daytime, nighttime)

def smallTableUtilizationPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimePhotoList, workbook):
    overall = furnUtilHelperPerDay(photoList, lambda x: x.smallTableList)
    daytime = furnUtilHelperPerDay(daytimePhotoList, lambda x: x.smallTableList)
    noOut = furnUtilHelperPerDay(noOutliersPhotoList, lambda x: x.smallTableList)
    daytimeNoOut = furnUtilHelperPerDay(daytimeNoOutliersPhotoList, lambda x: x.smallTableList)
    nighttime = furnUtilHelperPerDay(nighttimePhotoList, lambda x: x.smallTableList)

    writeToExcelbyDay(workbook, "smallTableUtilPerDay", "Date", "small table utlization per day", overall, daytime, noOut, daytimeNoOut, nighttime)

##def largeTableUtilizationPerPhoto(photoList, workbook):
##    excelList = []
##    for photo in photoList:
##        totalLargeTablesinUse = 0
##        for table in photo.largeTableList:
##            if table.numPeople != 0 or table.numGroups != 0:
##                totalLargeTablesinUse += 1
##        percentLargeTablesinUse = 0
##        if len(photo.largeTableList) != 0:
##            percentLargeTablesinUse = (100.0 * totalLargeTablesinUse) / len(photo.largeTableList)
##        excelList.append((photo.exactDate, percentLargeTablesinUse))
##
##    writeToExcelbyPhoto(workbook, "largeTableUtilizationPerPhoto", "Date", "percent of large tables being used per photo", excelList)
##
##def largeTableUtilizationPerDay(photoList, workbook):
##    excelList = []
##    daytimeList = []
##    currentDate = photoList[0].dayDate
##    currentTablesUsed = 0
##    daytimeTablesUsed = 0
##    currentTotalTables = 0
##    daytimeTotalTables = 0
##    currentPercentTablesUsed = 0
##    daytimePercentTablesUsed = 0
##    excelList.append((currentDate, currentPercentTablesUsed))
##    daytimeList.append((currentDate, daytimePercentTablesUsed))
##    i = 0
##    for photo in photoList:
##        if photo.dayDate == currentDate:            
##            currentTotalTables += len(photo.largeTableList)
##            for table in photo.largeTableList:
##                if table.numPeople != 0 or table.numGroups != 0:
##                    currentTablesUsed += 1
##            if currentTotalTables != 0:
##                currentPercentTablesUsed = (100.0 * currentTablesUsed) / currentTotalTables
##                
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                daytimeTotalTables += len(photo.largeTableList)
##                for table in photo.largeTableList:
##                    if table.numPeople != 0 or table.numGroups != 0:
##                        daytimeTablesUsed += 1
##                if daytimeTotalTables != 0:
##                    daytimePercentTablesUsed = (100.0 * daytimeTablesUsed) / daytimeTotalTables
##                    
##            excelList[i] = (currentDate, currentPercentTablesUsed)
##            daytimeList[i] = (currentDate, daytimePercentTablesUsed)
##            
##        else:
##            currentDate = photo.dayDate
##            currentTotalTables = len(photo.largeTableList)
##            currentTablesUsed = 0
##            for table in photo.largeTableList:
##                if table.numPeople != 0 or table.numGroups != 0:
##                    currentTablesUsed += 1
##            currentPercentTablesUsed = 0
##            if currentTotalTables != 0:
##                currentPercentTables = (100.0 * currentTablesUsed) / currentTotalTables
##
##            daytimeTablesUsed = 0
##            daytimeTotalTables = 0
##            daytimePercentTablesUsed = 0
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:                
##                daytimeTotalTables = len(photo.largeTableList)
##                for table in photo.largeTableList:
##                    if table.numPeople != 0 or table.numGroups != 0:
##                        daytimeTablesUsed += 1
##                if daytimeTotalTables != 0:
##                    daytimePercentTablesUsed = (100.0 * daytimeTablesUsed) / daytimeTotalTables
##                
##            excelList.append((currentDate, currentPercentTablesUsed))
##            daytimeList.append((currentDate, daytimePercentTablesUsed))
##            i += 1
##
##    writeToExcelbyDay(workbook, "largeTableUtilPerDay", "Date", "large table utlization per day", excelList, daytimeList)


def portion_of_A_using_B_per_photo(photoList, getBList, getAList, getNumA):
    myList = []
    for photo in photoList:
        AUsingBList = []
        for b in getBList(photo):
            for a in getAList(b):
                AUsingBList.append(a)
        percentAUsingB = 0
        if getNumA(photo) != 0:
            percentAUsingB = (100.0 * len(set(AUsingBList))) / getNumA(photo)
        myList.append((photo.exactDate, percentAUsingB))
    return myList
                                   

def portionOfPeopleUsingFurniturePerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)

    writeToExcelbyPhoto(workbook, "percentPeopleinFurniturePhoto", "Date", "percent of people using furniture per photo", overall, daytime, nighttime)

def portion_ofA_using_B_per_day_helper(photoList, getBList, getAList, getNumA):
    myList = []
    currentDate = photoList[0].dayDate
    currentAUsingBList = []
    currentPercentAUsingB = 0
    currentTotalA = 0
    myList.append((currentDate, currentPercentAUsingB))
    i = 0
    for photo in photoList:
        if photo.dayDate == currentDate:
            currentTotalA += getNumA(photo)
            for b in getBList(photo):
                for a in getAList(b):
                    currentAUsingBList.append(a)
            if currentTotalA != 0:
                currentPercentAUsingB = (100.0 * len(set(currentAUsingBList))) / currentTotalA
            myList[i] = (currentDate, currentPercentAUsingB)
        else:
            currentDate = photo.dayDate
            currentAUsingBList = []
            currentTotalA = getNumA(photo)
            currentPercentAUsingB = 0
            for b in getBList(photo):
                for a in getAList(b):
                    currentAUsingBList.append(a)
            if currentTotalA != 0:
                currentPercentAUsingB = (100.0 * len(set(currentAUsingBList))) / currentTotalA                
            myList.append((currentDate, currentPercentAUsingB))
            i += 1
    return myList

    
def portionOfPeopleUsingFurniturePerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.chairList + x.sofaList + x.smallTableList + x.largeTableList, lambda y: y.personList, lambda z: z.numPeople)
    
    writeToExcelbyDay(workbook, "percentPeopleinFurnitureDay", "Date", "percent of people using furniture per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def portionOfPeopleUsingSmallChairsPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)

    writeToExcelbyPhoto(workbook, "percentPeopleinChairsPhoto", "Date", "percent of people using small chairs per photo", overall, daytime, nighttime)

def portionOfPeopleUsingSmallChairsPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.personList, lambda z: z.numPeople)

    writeToExcelbyDay(workbook, "percentPeopleinChairsDay", "Date", "percent of people using small chairs per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def portionOfPeopleUsingSofasPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)

    writeToExcelbyPhoto(workbook, "percentPeopleinSofasPhoto", "Date", "percent of people using sofas per photo", overall, daytime, nighttime)

def portionOfPeopleUsingSofasPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.personList, lambda z: z.numPeople) 

    writeToExcelbyDay(workbook, "percentPeopleinSofasPerDay", "Date", "percent of people using sofas per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def portionOfPeopleUsingSmallTablesPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)

    writeToExcelbyPhoto(workbook, "percentPeopleinSmallTablesPhoto", "Date", "percent of people using small tables per photo", overall, daytime, nighttime)

def portionOfPeopleUsingSmallTablesPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.personList, lambda z: z.numPeople)
    
    writeToExcelbyDay(workbook, "percentPeopleinSmallTablesDay", "Date", "percent of people using small tables per day", overall, daytime, noOut, daytimeNoOut, nighttime)

##def portionOfPeopleUsingLargeTablesPerPhoto(photoList, workbook):
##    excelList = []
##    for photo in photoList:
##        peopleinLargeTablesList = []
##        for table in photo.largeTableList:
##            for person in table.personList:
##                peopleinLargeTablesList.append(person)
##        percentPeopleinLargeTables = 0
##        if photo.numPeople != 0:
##            percentPeopleinLargeTables = (100.0 * len(set(peopleinLargeTablesList))) / photo.numPeople
##        excelList.append((photo.exactDate, percentPeopleinLargeTables))
##
##    writeToExcelbyPhoto(workbook, "percentPeopleinLargeTablesPhoto", "Date", "percent of people using large tables per photo", excelList)
##
##def portionOfPeopleUsingLargeTablesPerDay(photoList, workbook):
##    excelList = []
##    daytimeList = []
##    currentDate = photoList[0].dayDate
##    currentPeopleinTablesList = []
##    daytimePeopleinTablesList = []
##    currentPercentPeopleinTables = 0
##    daytimePercentPeopleinTables = 0
##    currentTotalPeople = 0
##    daytimeTotalPeople = 0
##    daytimeList.append((currentDate, daytimePercentPeopleinTables))
##    excelList.append((currentDate, currentPercentPeopleinTables))
##    i = 0
##    for photo in photoList:
##        if photo.dayDate == currentDate:
##            currentTotalPeople += photo.numPeople
##            for table in photo.largeTableList:
##                for person in table.personList:
##                    currentPeopleinTablesList.append(person)
##            if currentTotalPeople != 0:
##                currentPercentPeopleinTables = (100.0 * len(set(currentPeopleinTablesList))) / currentTotalPeople
##
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                daytimeTotalPeople += photo.numPeople
##                for table in photo.largeTableList:
##                    for person in table.personList:
##                        daytimePeopleinTablesList.append(person)
##                if daytimeTotalPeople != 0:
##                    daytimePercentPeopleinTables = (100.0 * len(set(daytimePeopleinTablesList))) / daytimeTotalPeople
##
##            excelList[i] = (currentDate, currentPercentPeopleinTables)
##            daytimeList[i] = (currentDate, daytimePercentPeopleinTables)
##        else:
##            currentDate = photo.dayDate
##            currentPeopleinTablesList = []
##            currentTotalPeople = photo.numPeople
##            currentPercentPeopleinTables = 0
##            for table in photo.largeTableList:
##                for person in table.personList:
##                    currentPeopleinTablesList.append(person)
##            if currentTotalPeople != 0:
##                currentPercentPeopleinTables = (100.0 * len(set(currentPeopleinTablesList))) / currentTotalPeople
##
##            daytimePeopleinTablesList = []
##            daytimeTotalPeople = 0
##            daytimePercentPeopleinTables = 0
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:                
##                daytimeTotalPeople = len(photo.personList)
##                for table in photo.largeTableList:
##                    for person in table.personList:
##                        daytimePeopleinTablesList.append(person)
##                if daytimeTotalPeople != 0:
##                    daytimePercentPeopleinTables = (100.0 * len(set(daytimePeopleinTablesList))) / daytimeTotalPeople
##                
##            excelList.append((currentDate, currentPercentPeopleinTables))
##            daytimeList.append((currentDate, daytimePercentPeopleinTables))
##            i += 1
##
##    writeToExcelbyDay(workbook, "percentPeopleinLargeTablesDay", "Date", "percent of people using large tables per day", excelList, daytimeList)

def portionOfGroupsUsingSmallChairsPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)

    writeToExcelbyPhoto(workbook, "percentGroupsinSmallChairsPhoto", "Date", "percent of groups using small chairs per photo", overall, daytime, nighttime)


def portionOfGroupsUsingSmallChairsPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):    
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.chairList, lambda y: y.groupList, lambda z: z.numGroups)

    writeToExcelbyDay(workbook, "percentGroupsinChairsDay", "Date", "percent of groups using small chairs per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def portionOfGroupsUsingSofasPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)

    writeToExcelbyPhoto(workbook, "percentGroupsinSofasPhoto", "Date", "percent of groups using sofas per photo", overall, daytime, nighttime)

def portionOfGroupsUsingSofasPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.sofaList, lambda y: y.groupList, lambda z: z.numGroups)
    
    writeToExcelbyDay(workbook, "percentGroupsinSofasDay", "Date", "percent of groups using sofas per day", overall, daytime, noOut, daytimeNoOut, nighttime)

def portionOfGroupsUsingSmallTablesPerPhoto(photoList, daytimePhotoList, nighttimePhotoList, workbook):
    overall = portion_of_A_using_B_per_photo(photoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_of_A_using_B_per_photo(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_of_A_using_B_per_photo(nighttimePhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)

    writeToExcelbyPhoto(workbook, "percentGroupsinSmallTablesPhoto", "Date", "percent of groups using small tables per photo", overall, daytime, nighttime)

def portionOfGroupsUsingSmallTablesPerDay(photoList, daytimePhotoList, noOutliersPhotoList, daytimeNoOutliersPhotoList, nighttimeNoOutliersPhotoList, workbook):
    overall = portion_ofA_using_B_per_day_helper(photoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    daytime = portion_ofA_using_B_per_day_helper(daytimePhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    noOut = portion_ofA_using_B_per_day_helper(noOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    daytimeNoOut = portion_ofA_using_B_per_day_helper(daytimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    nighttime = portion_ofA_using_B_per_day_helper(nighttimeNoOutliersPhotoList, lambda x: x.smallTableList, lambda y: y.groupList, lambda z: z.numGroups)
    
    writeToExcelbyDay(workbook, "percentGroupsinSmallTablesDay", "Date", "percent of groups using small tables per day", overall, daytime, noOut, daytimeNoOut, nighttime)

##def portionOfGroupsUsingLargeTablesPerPhoto(photoList, workbook):
##    excelList = []
##    for photo in photoList:
##        groupsinLargeTablesList = []
##        for table in photo.largeTableList:
##            for group in table.groupList:
##                groupsinLargeTablesList.append(group)
##        percentGroupsinLargeTables = 0
##        if photo.numGroups != 0:
##            percentGroupsinLargeTables = (100.0 * len(set(groupsinLargeTablesList))) / photo.numGroups
##        excelList.append((photo.exactDate, percentGroupsinLargeTables))
##
##    writeToExcelbyPhoto(workbook, "percentGroupsinLargeTablesPhoto", "Date", "percent of groups using large tables per photo", excelList)
##
##def portionOfGroupsUsingLargeTablesPerDay(photoList, workbook):
##    excelList = []
##    daytimeList = []
##    currentDate = photoList[0].dayDate
##    currentGroupsinTablesList = []
##    daytimeGroupsinTablesList = []
##    currentPercentGroupsinTables = 0
##    daytimePercentGroupsinTables = 0
##    currentTotalGroups = 0
##    daytimeTotalGroups = 0
##    daytimeList.append((currentDate, daytimePercentGroupsinTables))
##    excelList.append((currentDate, currentPercentGroupsinTables))
##    i = 0
##    for photo in photoList:
##        if photo.dayDate == currentDate:
##            currentTotalGroups += photo.numGroups
##            for table in photo.largeTableList:
##                for group in table.groupList:
##                    currentGroupsinTablesList.append(group)
##            if currentTotalGroups != 0:
##                currentPercentGroupsinTables = (100.0 * len(set(currentGroupsinTablesList))) / currentTotalGroups
##
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:
##                daytimeTotalGroups += photo.numGroups
##                for table in photo.largeTableList:
##                    for person in table.groupList:
##                        daytimeGroupsinTablesList.append(person)
##                if daytimeTotalGroups != 0:
##                    daytimePercentGroupsinTables = (100.0 * len(set(daytimeGroupsinTablesList))) / daytimeTotalGroups
##
##            excelList[i] = (currentDate, currentPercentGroupsinTables)
##            daytimeList[i] = (currentDate, daytimePercentGroupsinTables)
##        else:
##            currentDate = photo.dayDate
##            currentGroupsinTablesList = []
##            currentTotalGroups = photo.numGroups
##            currentPercentGroupsinTables = 0
##            for table in photo.largeTableList:
##                for group in table.groupList:
##                    currentGroupsinTablesList.append(group)
##            if currentTotalGroups != 0:
##                currentPercentGroupsinTables = (100.0 * len(set(currentGroupsinTablesList))) / currentTotalGroups
##
##            daytimeGroupsinTablesList = []
##            daytimeTotalGroups = 0
##            daytimePercentGroupsinTables = 0
##            if photo.exactDate.hour >= 8 and photo.exactDate.hour < 20:                
##                daytimeTotalGroups = len(photo.groupList)
##                for table in photo.largeTableList:
##                    for group in table.groupList:
##                        daytimeGroupsinTablesList.append(group)
##                if daytimeTotalGroups != 0:
##                    daytimePercentGroupsinTables = (100.0 * len(set(daytimeGroupsinTablesList))) / daytimeTotalGroups
##                
##            excelList.append((currentDate, currentPercentGroupsinTables))
##            daytimeList.append((currentDate, daytimePercentGroupsinTables))
##            i += 1
##
##    writeToExcelbyDay(workbook, "percentGroupsinLargeTablesDay", "Date", "percent of groups using large tables per day", excelList, daytimeList)


readFiles("C:\Users\Nicole\Documents\UROP2013\data")
